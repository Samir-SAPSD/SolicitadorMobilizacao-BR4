# Preenchimento Automático do Analista Responsável

## Visão Geral

Foi implementado um sistema automático de preenchimento da coluna **"Analista Responsável"** baseado em 4 regras de negócio. Este sistema é aplicado automaticamente durante o envio de dados para o SharePoint.

## Como Funciona

O processo ocorre em dois arquivos principais:

### 1. **Apply-AnalistRules.ps1** (Módulo de Regras)
Contém todas as funções que implementam a lógica de atribuição automática do Analista.

### 2. **Populate-SharePointList.ps1** (Script Principal)
Chama o módulo `Apply-AnalistRules.ps1` automaticamente após ler os dados do Excel e antes de enviá-los para o SharePoint.

## As 4 Regras de Negócio

As regras são aplicadas em **ordem de prioridade**. Uma vez que um analista é atribuído, as próximas regras não o modificam.

### **Regra 1: FORNECEDOR (Prioridade Alta)**
```
Se: FORNECEDOR em (FLEXWIND, ARTHWIND, TETRACE, REVTECH, TECH SERVICES, AERIS, PRIME WIND, BELA VISTA, DRONE BASE)
    E Analista Responsável está vazio
Então: Atribui "MAFDO"
```

### **Regra 2: TIPO DE MOBILIZAÇÃO**
```
Se: TIPO DE MOBILIZAÇÃO = "Máquinas e Equipamentos"
    E Analista Responsável está vazio
Então: Atribui "SAIOI / LPHDS"
```

### **Regra 3: TIPO DE ATIVIDADE**
```
Se: TIPO DE ATIVIDADE inicia com "ST - "
    E Analista Responsável está vazio
Então: Atribui "SAIOI / LPHDS"
```

### **Regra 4: PARQUE (Fallback)**
```
Se: Nenhuma das regras anteriores foi aplicada
    E Analista Responsável está vazio
    E Existe mapeamento Parque -> Analista no SharePoint
Então: Busca o Parque do item e atribui o analista correspondente
```

## Campos Reconhecidos

O módulo reconhece variações dos nomes de colunas (case-insensitive) para maior flexibilidade:

| Regra | Variações de Nome Aceitas |
|-------|--------------------------|
| **FORNECEDOR** | Fornecedor, FORNECEDOR, Supplier |
| **TIPO DE MOBILIZAÇÃO** | Tipo, TipoMobilizacao, Mobilizacao, TipoMobiliza, Mobilization Type |
| **TIPO DE ATIVIDADE** | Atividade, TIPODEATIVIDADE, TipoAtividade, Activity, Type |
| **PARQUE** | Parque, Park, WindPark, LocationPark |
| **ANALISTA** | Analista, Analista Responsável, AnalistaResponsavel, Analyst, Responsible Analyst |

## Exemplo de Uso

Quando você envia um arquivo Excel para o **Solicitador de Mobilizações**:

1. O arquivo é carregado e validado normalmente
2. Antes de enviar para o SharePoint, o módulo `Apply-AnalistRules` é acionado
3. Cada linha é processada pelas 4 regras em ordem
4. Os analistas são atribuídos automaticamente conforme aplicável
5. Um resumo mostra quantos itens foram afetados por cada regra
6. Os dados são then enviados para o SharePoint com o Analista já preenchido

### Exemplo de Saída no Console:
```
Resumo de Preenchimento do Analista Responsável:
  Regra 1 (Fornecedor)          : 5
  Regra 2 (Tipo de Mobilização): 3
  Regra 3 (Tipo de Atividade)  : 2
  Regra 4 (Parque)             : 8
  Ainda Vazios                 : 4
  Total Processado             : 22
```

## Campos do SharePoint Usados

- **Fornecedor**: FORNECEDOR
- **Tipo de Mobilização**: TipoMobiliza_x00e7__x00e3_o
- **Tipo de Atividade**: TIPODEATIVIDADE
- **Parque**: Parque (com lookup para o Analista)
- **Analista Responsável**: AnalistaRespons_x00e1_vel

## Tratamento de Acentos e Case-Insensitivity

Todas as comparações (exceto para valores como nomes de analistas) são feitas com normalização:
- Removem acentos (ex: "Máquinas" → "Maquinas")
- Convertem para maiúsculas
- Ignoram espaços extras

Isso garante que pequenas variações de digitação não quebrem a lógica.

## Regra 4 - Mapeamento Parque → Analista

A Regra 4 busca dinamicamente o mapeamento na lista de Parques do SharePoint.

**Lista de Parques**: ID = `678f10f9-8d46-404b-a451-70dfe938a1ee`

O campo **"Analista Responsável"** em cada registro de Parque é usado como valor padrão quando aplica a Regra 4.

## Troubleshooting

### Nenhum analista foi atribuído
- Verifique se os nomes dos campos estão corretos no Excel
- Verifique se há dados nas colunas obrigatórias (Fornecedor, Tipo de Mobilização, etc)
- Verifique se o mapeamento de Parques no SharePoint está preenchido (Regra 4)

### Alguns itens têm "Ainda Vazios"
- Esses itens não se encaixam em nenhuma das 4 regras
- Você pode preenchê-los manualmente ou ajustar as regras no arquivo `Apply-AnalistRules.ps1`

### Erro "Módulo Apply-AnalistRules não disponível"
- Verifique se o arquivo `Apply-AnalistRules.ps1` está no mesmo diretório que `Populate-SharePointList.ps1`
- Verifique se a conexão ao SharePoint foi estabelecida corretamente (necessária para Regra 4)

## Customização

Para modificar as regras, edite o arquivo `Apply-AnalistRules.ps1`:

```powershell
# Linha ~70: Alterar lista de fornecedores
$FornecedoresRegra1 = @("FLEXWIND", "ARTHWIND", "..." )

# Linha ~72: Alterar valor atribuído
$ValorRegra1 = "MAFDO"

# Linha ~73: Alterar tipo de mobilização
$TipoMobilizacaoRegra2 = "Maquinas e Equipamentos"
```

## Integração com Outras Ferramentas

O mesmo módulo pode ser usado por outros scripts PowerShell. Basta importar e chamar:

```powershell
. .\Apply-AnalistRules.ps1
$items = Apply-AnalistRules -Items $items -Verbose
```

## Histórico de Implementação

- **Data**: Abril 2026
- **Base**: Script `Set-AnalistaPorTipoAtividade.ps1` (modelo de referência)
- **Adaptação**: Implementado para processar dados antes do envio ao SharePoint (vs. após estarem já no SP)
