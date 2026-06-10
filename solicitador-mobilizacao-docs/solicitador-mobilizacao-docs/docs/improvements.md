---
sidebar_position: 3
title: Melhorias Identificadas
---

# Melhorias e Dívidas Técnicas

## Problemas Críticos — Alta Prioridade

### 1. Duplicação de Controle de Jobs (`config.py` vs `state.py`)

**Arquivo**: [config.py](./files/config.md), [state.py](./files/state.md)

`config.py` define `ACTIVE_JOBS`, `JOBS_LOCK`, `start_job()`, `end_job()`, `get_active_jobs()`. `state.py` redefine as mesmas variáveis e funções (sem `get_active_jobs`). O `routes.py` importa de `state.py`, mas `config.py` ainda tem código morto de controle de jobs.

**Risco**: Duas instâncias separadas do contador. Se alguém importar de `config.py` em vez de `state.py`, os contadores ficam dessincronizados.

**Solução**: Remover `ACTIVE_JOBS`, `JOBS_LOCK`, `start_job`, `end_job`, `get_active_jobs` de `config.py`. Manter apenas em `state.py`. Adicionar `get_active_jobs()` em `state.py`.

---

### 2. Nomenclatura Contraintuitiva de Pastas

**Arquivo**: [config.py](./files/config.md)

```python
UPLOAD_FOLDER  = "backend/data/templates"  # Armazena o TEMPLATE (não uploads)
REPORTS_FOLDER = "backend/data/uploads"    # Armazena uploads E relatórios
```

Os nomes são inversos ao que fazem: `UPLOAD_FOLDER` guarda o template; `REPORTS_FOLDER` guarda uploads temporários e relatórios gerados.

**Risco**: Confusão para manutenção. Um desenvolvedor pode salvar o arquivo errado no diretório errado.

**Solução**: Renomear para `TEMPLATES_FOLDER` e `WORKDIR_FOLDER` (ou similar), atualizando todas as referências.

---

### 3. Ausência de Autenticação nos Endpoints

**Arquivo**: [api/routes.py](./files/routes.md)

Nenhum endpoint tem autenticação ou autorização. Qualquer pessoa com acesso à rede que descobrir a URL pode:
- Baixar o template (`GET /download-template`)
- Submeter dados ao SharePoint (`POST /run-script`)
- Listar e baixar todos os relatórios (`GET /list-reports`, `GET /download-report/...`)
- Encerrar o servidor (`POST /shutdown`)

**Risco**: Alto em ambientes de rede corporativa com múltiplos usuários.

**Solução**: Implementar autenticação básica (OAuth2 com Azure AD / Entra ID, que está disponível na Vestas) ou ao menos restringir por IP de origem.

---

### 4. `POST /shutdown` sem Proteção

**Arquivo**: [api/routes.py](./files/routes.md)

O endpoint `POST /shutdown` encerra o processo Python completamente. Qualquer usuário na rede pode encerrar o servidor de outros analistas.

**Solução Imediata**: Restringir o endpoint para aceitar apenas requisições de `127.0.0.1` (localhost):

```python
@api_bp.route('/shutdown', methods=['POST'])
def shutdown():
    if request.remote_addr != '127.0.0.1':
        return jsonify({'error': 'Forbidden'}), 403
    ...
```

---

### 5. Arquivo de Upload não Removido Após Validação

**Arquivo**: [api/routes.py](./files/routes.md) — endpoint `/validate`

O arquivo `.xlsx` enviado pelo usuário é salvo em `UPLOAD_FOLDER` durante a validação, mas **nunca é removido** após o endpoint `/validate` retornar. Somente o endpoint `/run-script` reutiliza o arquivo.

Se o usuário validar mas não submeter, o arquivo fica permanentemente no disco.

**Solução**: Implementar limpeza periódica de arquivos antigos (ex: arquivos com mais de 24h) ou remover após validação bem-sucedida quando não há submissão pendente.

---

## Dívidas Técnicas — Média Prioridade

### 6. Ausência de Testes Automatizados

Nenhum arquivo de teste foi encontrado no projeto. Sem testes:
- Refatorações são arriscadas
- Bugs em `encoding_repair.py` e `excel_processor.py` (lógica complexa) não são detectados automaticamente

**Sugestão**: Priorizar testes para:
- `encoding_repair.py` — lógica de pontuação e reparo (pura, fácil de testar)
- `excel_processor.py` — casos edge: linha vazia, duplicata, GRUPO vazio
- `correction_helper.py` — normalização e matching de campos

---

### 7. Dependências não Utilizadas em `requirements.txt`

```
customtkinter   ← GUI desktop, não usado no servidor web
packaging       ← Utilitário de versões, não usado
```

**Impacto**: Instalação mais lenta, ambiguidade sobre o propósito do projeto.

**Solução**: Remover ou mover para `requirements-desktop.txt` se houver uma interface desktop legada.

---

### 8. `score_decoded_text` penaliza `â` legítimo

**Arquivo**: [encoding_repair.py](./files/encoding-repair.md)

```python
mojibake_penalty = sum(text.count(marker) for marker in ('Ã', 'Â', 'â')) * 6
```

O caractere `â` (a com circunflexo) é **legítimo** em português (ex: "câmara", "âmbito"). Penalizá-lo pode fazer o algoritmo preferir uma decodificação pior para textos com essas palavras.

**Solução**: Remover `'â'` dos marcadores de mojibake, pois `'Ã'` e `'Â'` já são suficientes para detectar o padrão de erro UTF-8→cp1252.

---

### 9. Sem Logging Estruturado

O sistema usa `print()` e strings formatadas via `yield` para logging. Sem logging estruturado:
- Impossível filtrar logs por nível (INFO/WARNING/ERROR)
- Impossível integrar com sistemas de monitoramento

**Solução**: Substituir `print()` por `logging.getLogger(__name__)` com handlers configurados no startup.

---

### 10. Sem Limite de Tamanho de Upload

**Arquivo**: [api/routes.py](./files/routes.md)

Não há `MAX_CONTENT_LENGTH` configurado no Flask. Um usuário pode enviar um arquivo Excel de centenas de megabytes, causando consumo excessivo de memória.

**Solução**:
```python
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16 MB
```

---

## Sugestões de Melhoria — Baixa Prioridade

### 11. Endpoint de Health Check

Adicionar `GET /health` que retorna status do servidor, jobs ativos e última atualização do template:

```python
@api_bp.route('/health')
def health():
    return jsonify({
        'status': 'ok',
        'active_jobs': get_active_jobs(),
        'template_last_updated': get_template_status()['last_updated']
    })
```

### 12. Persistência de Relatórios

Relatórios são salvos em disco local. Se o servidor for reiniciado em outra máquina ou o disco for limpo, todos os relatórios históricos se perdem.

**Sugestão**: Salvar relatórios no SharePoint (mesma integração já existente) ou em Azure Blob Storage.

### 13. Tratamento de `report_data` sem `items`

**Arquivo**: [report_generator.py](./files/report-generator.md)

```python
items = report_data.get('items', [])
```

Se `items` for `None` (não apenas ausente), `get()` retorna `None`, causando erro na iteração. Usar `report_data.get('items') or []` seria mais robusto.

### 14. Largura de Colunas Baseada Apenas no Cabeçalho

**Arquivo**: [report_generator.py](./files/report-generator.md)

A largura estimada `max(14, min(len(col_name) + 4, 40))` considera apenas o tamanho do nome da coluna. Valores longos nos dados (ex: nomes completos) serão truncados visualmente.

**Sugestão**: Calcular o máximo entre o comprimento do cabeçalho e o comprimento máximo dos valores daquela coluna.

---

## Problemas no Frontend — Média/Alta Prioridade

### 15. CDN sem Verificação de Integridade (SRI)

**Arquivo**: [index.html](./files/index-html.md)

```html
<!-- Atual — sem verificação -->
<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">

<!-- Correto — com Subresource Integrity -->
<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css"
      rel="stylesheet"
      integrity="sha384-..."
      crossorigin="anonymous">
```

Um atacante que comprometesse o CDN poderia injetar JavaScript malicioso. SRI garante que o browser recuse o arquivo se o hash não bater.

---

### 16. Endpoint `/apply-corrections` não documentado no Backend
### 16. ~~Endpoint `/apply-corrections` não documentado~~ — Resolvido ✓

**Arquivo**: [routes.py](./files/routes.md)

O endpoint `POST /apply-corrections` **está implementado** em `api/routes.py` (linha 514). A documentação foi adicionada em [routes.md](./files/routes.md#post-apply-corrections--aplicar-correções-no-excel). O item foi erroneamente classificado como ausente na análise inicial.

---

### 17. Espaço no Nome do Asset de Logo

**Arquivo**: [index.html](./files/index-html.md)

```
frontend/static/Vestas Primary Logo RGB.png   ← Espaços no nome
```

Espaços em nomes de arquivo podem causar problemas com ferramentas de build, proxies ou servidores que não encodam paths automaticamente.

**Solução**: Renomear para `vestas-logo.png` e atualizar a referência no HTML.

---

### 18. `dataset.suggestions` — Serialização no DOM

**Arquivo**: [index.html](./files/index-html.md)

```javascript
modal.dataset.suggestions = JSON.stringify(suggestions);
```

Armazenar JSON no `dataset` do DOM funciona, mas tem limite de tamanho e pode causar problemas com caracteres especiais (aspas, etc.). Para volumes maiores de sugestões, um módulo de estado em memória seria mais robusto.

---

## Boas Práticas Ausentes

| Prática | Impacto | Módulo |
|---|---|---|
| Type hints em funções de `routes.py` | Legibilidade, IDE support | `api/routes.py` |
| Docstrings em funções de `routes.py` | Manutenção | `api/routes.py` |
| Context manager para workbook openpyxl | Resource leak prevention | `excel_processor.py` |
| Validação de schema do JSON do PS1 | Falha mais clara em mudanças de contrato | `api/routes.py` |
| Timeout em `subprocess.Popen` | Previne hang indefinido | `api/routes.py` |
| SRI em tags de CDN (Bootstrap) | Segurança contra CDN comprometida | `index.html` |
| Separar CSS/JS do HTML | Manutenibilidade do frontend | `index.html` |

---

## Próximos Passos — Priorizados por Impacto

1. **[CRÍTICO]** Adicionar restrição de IP ou autenticação no `/shutdown`
2. **[ALTO]** Remover duplicação `config.py` vs `state.py`
3. **[ALTO]** Configurar `MAX_CONTENT_LENGTH` no Flask
4. **[ALTO]** Adicionar `integrity` SRI nas tags de CDN do Bootstrap
5. **[MÉDIO]** Corrigir penalização de `â` em `score_decoded_text`
6. **[MÉDIO]** Adicionar testes para `encoding_repair.py` e `excel_processor.py`
7. **[MÉDIO]** Implementar limpeza de arquivos de upload antigos
8. **[BAIXO]** Renomear `Vestas Primary Logo RGB.png` para `vestas-logo.png`
9. **[BAIXO]** Adicionar endpoint `/health`
10. **[BAIXO]** Remover dependências não utilizadas do `requirements.txt`
