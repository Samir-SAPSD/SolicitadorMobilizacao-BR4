---
sidebar_position: 1
title: routes.py
---

# `backend/api/routes.py`

## Responsabilidade

> Blueprint Flask com todos os endpoints HTTP da aplicação. Orquestra o fluxo completo de validação, submissão ao SharePoint e download de relatórios.

Este arquivo é a **camada de apresentação** do sistema: recebe requisições HTTP, delega ao domínio (módulos `backend/`) e retorna respostas estruturadas. Contém a lógica de orquestração de mais alto nível, incluindo o processamento em streaming por grupos.

## Localização

```
backend/api/routes.py
```

## Dependências

### Internas
| Módulo | Símbolos Usados |
|---|---|
| `backend.config` | `SCRIPTS_DIR` |
| `backend.state` | `start_job`, `end_job` |
| `backend.services` | Todos os demais símbolos (via Facade) |

### Externas
| Biblioteca | Propósito |
|---|---|
| `flask` | `Blueprint`, `Response`, `jsonify`, `render_template`, `request`, `stream_with_context`, `send_from_directory` |
| `werkzeug.utils.secure_filename` | Sanitização de nomes de arquivo enviados pelo usuário |
| `subprocess` | Execução de scripts PowerShell |
| `threading` | Thread para shutdown assíncrono |

---

## Blueprint

```python
api_bp = Blueprint('api', __name__)
```

Registrado em `app.py` sem prefixo de URL — todas as rotas ficam na raiz.

---

## Endpoints

### `GET /` — Index

```python
@api_bp.route('/')
def index():
    return render_template('index.html')
```

Serve o frontend HTML via Jinja2. O template é resolvido pelo Flask usando `FRONTEND_TEMPLATES_DIR` configurado em `app.py`.

---

### `GET /download-template` — Download do Template Excel

```python
@api_bp.route('/download-template')
def download_template():
    return send_from_directory(current_app.config['UPLOAD_FOLDER'], TEMPLATE_FILENAME, as_attachment=True)
```

Permite ao analista baixar o template `ModeloSolicitacaoMob.xlsx` para preencher antes do upload.

---

### `GET /template-update-status` — Status de Atualização do Template

Retorna `{"last_updated": "2024-11-15T14:30:00"}` ou `{"last_updated": null}`. Usado pelo frontend para exibir quando o template foi atualizado pela última vez.

---

### `POST /update-template` — Atualizar Opções do Template

**Fluxo**:
1. Incrementa job counter (`start_job()`)
2. Verifica se o template e o script existem
3. Executa `Update-ExcelTemplateChoices.ps1` via `subprocess.Popen` (síncrono — aguarda conclusão)
4. Decodifica output com `decode_powershell_output()`
5. Se sucesso: persiste timestamp com `save_template_status()`, retorna JSON com log
6. Se falha: retorna erro com log do PowerShell

**Por que síncrono (não streaming)?** A atualização do template é uma operação administrativa infrequente. Aguardar a conclusão e retornar o resultado em um único JSON é mais simples e suficiente.

---

### `POST /validate` — Validação do Excel por Grupos

**Propósito**: Fase 1 do processo. Valida a estrutura e os dados do Excel *antes* de qualquer escrita no SharePoint.

**Fluxo detalhado**:

```mermaid
flowchart TD
    A[Receber arquivo via multipart/form-data] --> B{Arquivo válido e extensão xlsx?}
    B -->|Não| C[HTTP 400]
    B -->|Sim| D{Tamanho > 0?}
    D -->|Não| E[HTTP 400 - arquivo vazio]
    D -->|Sim| F[secure_filename + salvar em UPLOAD_FOLDER]
    F --> G[_read_grouped_excel: validar estrutura e agrupar]
    G -->|errors| H[HTTP 200 com status=failed e erros]
    G -->|ok| I{Script Validate-ExcelData.ps1 existe?}
    I -->|Não| J[HTTP 500]
    I -->|Sim| K[Para cada GRUPO]
    K --> L[_create_group_workbook: criar temp .xlsx do grupo]
    L --> M[Para cada aba PESSOAS e EQUIPAMENTOS]
    M --> N[build_powershell_command + subprocess.Popen]
    N --> O[decode_powershell_output]
    O --> P{Marcadores JSON presentes?}
    P -->|Não| Q[Adicionar erro de resultado estruturado ausente]
    P -->|Sim| R[json.loads do bloco JSON]
    R --> S{Tem erros no resultado?}
    S -->|Sim| T[Acumular erros com prefixo GRUPO|ABA]
    S -->|Não| U[Continuar]
    U --> K
    T --> K
    K -->|Todos os grupos processados| V[Retornar JSON com status, erros, group_summary]
```

**Por que verificar tamanho do arquivo duas vezes** (antes e depois de salvar)? O tamanho do stream pode ser 0 antes de salvar por limitação do Flask, mas o arquivo salvo ter 0 bytes indica outro problema (ex: disco cheio, permissão). A segunda verificação é mais confiável.

**Limpeza de arquivos temporários por grupo**:
```python
finally:
    try:
        if os.path.exists(temp_group_path):
            os.remove(temp_group_path)
    except Exception:
        pass
```

O `finally` garante que o workbook temporário de cada grupo é removido após a validação, independentemente de sucesso ou falha. A exceção silenciada previne que um erro de limpeza mascare um erro de validação.

---

### `POST /run-script` — Submissão ao SharePoint (Streaming por Grupos)

**Propósito**: Fase 2. Submete os dados ao SharePoint grupo a grupo, enviando progresso em tempo real ao browser via HTTP streaming.

**Por que streaming?** A operação pode levar vários minutos para grupos grandes. Sem streaming, o browser ficaria bloqueado com "carregando" até o fim. Com streaming, o analista vê progresso em tempo real linha a linha.

**Protocolo de streaming**:
O servidor emite texto plano (`text/plain; charset=utf-8`) com linhas especiais:

| Marcador | Propósito |
|---|---|
| `---GROUP_PROGRESS:1/3:NomeGrupo---` | Notifica início do grupo N de Total |
| `---REPORT_FILE:Relatorio_Mob_xxx.xlsx---` | Nome do arquivo de relatório gerado |
| `---GROUP_RESULT:{json}---` | Resultado do grupo (status, qtd, id_mob) |
| `---GROUP_SUMMARY_JSON_START---` ... `---GROUP_SUMMARY_JSON_END---` | JSON consolidado de todos os grupos |

**Lógica de parada em falha**:
```python
if not group_ok:
    stop_processing = True
    yield f'[GRUPO {group_key}] Falhou. Encerrando processamento dos próximos grupos.\n'
```

Se um grupo falha, os grupos subsequentes **não são processados**. Isso é intencional: evita criar registros SharePoint parciais quando um problema sistêmico (ex: credencial expirada, lista bloqueada) afetaria todos os grupos igualmente.

**Detecção de erros via marcadores em stdout**:
```python
error_markers = [
    'FALHA CRÍTICA', 'UPLOAD CANCELADO', '--- RESULT: ERROR ---',
    'Write-Error', 'Erro ao adicionar item',
    'Não foi possível gerar um ID_Mobilizacao único',
]
if any(marker in decoded for marker in error_markers):
    detected_error_in_output = True
```

O `process.returncode` não é suficiente — scripts PowerShell frequentemente retornam código 0 mesmo em falha parcial. A busca por marcadores textuais no stdout é a verificação primária de falha.

---

### `POST /shutdown` — Encerrar Servidor

```python
def _shutdown():
    os._exit(0)
t = threading.Timer(0.5, _shutdown)
t.daemon = True
t.start()
return jsonify({'message': 'Servidor encerrado.'}), 200
```

Encerra o processo Python após 0.5 segundos. O delay permite que o Flask complete o envio da resposta HTTP antes de encerrar. `os._exit(0)` é usado em vez de `sys.exit()` porque `sys.exit()` lança `SystemExit`, que pode ser capturado por handlers, enquanto `os._exit()` encerra imediatamente sem cleanup.

---

### `GET /list-reports` — Listar Relatórios

Lista todos os `.xlsx` em `REPORTS_FOLDER`, ordenados do mais recente ao mais antigo. Retorna nome, tamanho e data de modificação formatada.

---

### `GET /download-report/<filename>` — Download de Relatório

**Validação de path traversal**:
```python
safe_name = secure_filename(filename)
reports_abs = os.path.abspath(current_app.config['REPORTS_FOLDER'])
file_abs = os.path.abspath(os.path.join(reports_abs, safe_name))
if not file_abs.startswith(reports_abs + os.sep):
    return jsonify({'error': 'Acesso negado.'}), 403
```

Esta verificação previne **path traversal attack** (ex: `../../etc/passwd`). Mesmo com `secure_filename`, a verificação de que o caminho resolvido começa com o diretório de reports é a defesa definitiva.

---

### `POST /suggest-corrections` — Sugestões de Correção de Choices

**Request body**:
```json
{
    "filename": "ModeloSolicitacaoMob.xlsx",
    "errors": [
        {"campo": "FuncaoPessoa", "valor": "Eletrisista", "line": 5, "sheet": "PESSOAS", "group": "GRUPO-A"}
    ]
}
```

**Response**:
```json
{
    "suggestions": [
        {
            "campo": "FuncaoPessoa",
            "valor": "Eletrisista",
            "options": ["Eletricista", "Mecânico", "Técnico de Segurança"],
            "suggested": "Eletricista",
            "column_name": "Função de Pessoa",
            "internal_name": "FuncaoPessoa"
        }
    ]
}
```

**Fluxo detalhado de matching**:

```mermaid
flowchart TD
        A[Para cada entry em errors] --> B[normalize_key campo]
        B --> C{Match exato em column_name ou internal_name?}
        C -->|Sim| D[Usar coluna encontrada]
        C -->|Não| E{Match parcial: campo contém coluna ou vice-versa?}
        E -->|Sim| D
        E -->|Não| F[Sem coluna correspondente → options=[]]
        D --> G[Para cada option da coluna]
        G --> H{normalize_key option == normalize_key valor?}
        H -->|Sim| I[suggested = option — match exato normalizado]
        H -->|Não| J{option contém valor ou valor contém option?}
        J -->|Sim| K[Calcular score = diferença de comprimento]
        K --> L{Score < best_score?}
        L -->|Sim| M[suggested = option — melhor match parcial]
        L -->|Não| G
        J -->|Não| G
        I --> N[Adicionar à lista de suggestions]
        M --> N
        F --> N
```

**Algoritmo de sugestão automática**:

```python
for opt in options:
        opt_norm = normalize_key(opt)
        if opt_norm == valor_norm:
                suggested = opt
                break                     # Match exato — para imediatamente
        if valor_norm in opt_norm or opt_norm in valor_norm:
                score = abs(len(opt_norm) - len(valor_norm))
                if score < best_score:
                        best_score = score
                        suggested = opt       # Candidato parcial — continua buscando melhor
```

**Por que a diferença de comprimento como score?** `"Eletrisista"` (11 chars) tem distância 1 de `"Eletricista"` (11 chars) e distância 5 de `"Técnico Eletricista"` (18 chars). O menor score favorece a opção mais similar em tamanho — uma heurística eficiente sem Levenshtein.

**Por que duas tentativas de matching de coluna (exato + parcial)?** Os nomes dos campos nos erros retornados pelo SharePoint via PowerShell podem ter capitalização diferente, abreviações ou o nome interno em vez do nome de display. O match parcial cobre casos como `"FuncaoPessoa"` encontrando `"Função de Pessoa [FuncaoPessoa]"`.

---

### `POST /apply-corrections` — Aplicar Correções no Excel

**Propósito**: Modifica diretamente o arquivo Excel salvo no servidor, substituindo valores inválidos pelos valores corretos escolhidos pelo analista no modal de correção guiada. Após aplicar, o analista pode clicar "Iniciar Importação" sem fazer novo upload.

**Request body**:
```json
{
    "filename": "ModeloSolicitacaoMob.xlsx",
    "corrections": [
        {
            "sheet": "PESSOAS",
            "excel_row": 5,
            "internal_name": "FuncaoPessoa",
            "new_value": "Eletricista"
        }
    ]
}
```

**Fluxo de execução**:

```mermaid
flowchart TD
        A[Receber JSON com filename e corrections[]] --> B{filename presente?}
        B -->|Não| C[HTTP 400]
        B -->|Sim| D{Arquivo existe em UPLOAD_FOLDER?}
        D -->|Não| E[HTTP 404]
        D -->|Sim| F[openpyxl.load_workbook arquivo]
        F --> G[Para cada correction]
        G --> H{sheet_name existe no workbook?}
        H -->|Não| I[Pular correção]
        H -->|Sim| J[Para cada coluna da linha 1 cabeçalho]
        J --> K{Cabeçalho tem formato 'Nome [InternalName]'?}
        K -->|Sim| L{normalize_key internal ou display == normalize_key internal_name?}
        K -->|Não| M{normalize_key header == normalize_key internal_name?}
        L -->|Sim| N[target_col encontrado]
        M -->|Sim| N
        L -->|Não| J
        M -->|Não| J
        N --> O[ws.cell row=excel_row col=target_col .value = new_value]
        O --> P[applied++]
        P --> G
        G -->|Todas as correções processadas| Q[wb.save file_path]
        Q --> R[HTTP 200 com applied count]
```

**Lógica de localização da coluna de destino**:

```python
m = re.match(r'^(?P<disp>.+?)\s*\[(?P<internal>[^\[\]]+)\]\s*$', header)
if m:
        if (normalize_key(m.group('internal')) == normalize_key(internal_name) or
                        normalize_key(m.group('disp')) == normalize_key(internal_name)):
                target_col = col
else:
        if normalize_key(header) == normalize_key(internal_name):
                target_col = col
```

A mesma regex de `correction_helper.py` é usada aqui para localizar a coluna. Aceita tanto o nome interno (`FuncaoPessoa`) quanto o nome de display (`Função de Pessoa`) como `internal_name` na requisição. Isso é importante porque o frontend pode enviar qualquer um dos dois, dependendo de qual foi extraído do erro do SharePoint.

**Por que modificar o arquivo no servidor em vez de retornar o arquivo corrigido?** O arquivo já está salvo com o nome que o backend usa nas fases seguintes. Modificar in-place evita que o analista precise fazer novo upload — mantém a continuidade do fluxo: validação → correção → submissão, tudo com o mesmo `filename`.

**Risco de segurança mitigado**: O `secure_filename(filename)` sanitiza o nome antes de construir o caminho, e `os.path.join(UPLOAD_FOLDER, ...)` garante que o arquivo só pode estar dentro de `UPLOAD_FOLDER`. Não há path traversal possível.

**Nota**: Não há validação de que `new_value` seja uma das opções válidas da lista. Um usuário malicioso poderia enviar qualquer string como `new_value`. Como o sistema é interno e a validação definitiva acontece no PowerShell (que consulta o SharePoint), esta ausência de validação local é aceitável.
