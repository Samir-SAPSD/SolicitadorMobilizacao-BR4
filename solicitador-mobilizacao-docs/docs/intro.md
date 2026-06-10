---
sidebar_position: 1
title: Introdução
---

# Solicitador de Mobilização — Documentação Técnica

## O que é este sistema?

O **Solicitador de Mobilização** é uma aplicação web interna da Vestas que automatiza o processo de solicitação e registro de mobilização de pessoas e equipamentos. A aplicação permite que analistas façam upload de planilhas Excel estruturadas, valide os dados contra regras de negócio definidas via SharePoint, e submeta os registros diretamente para listas do SharePoint Online, gerando relatórios Excel com estilização Vestas ao final.

## Problema que resolve

Antes desta ferramenta, o processo de mobilização era manual, propenso a erros de digitação, inconsistências com os valores aceitos pelo SharePoint e dependente de operações repetitivas. O sistema automatiza:

1. Validação estrutural do Excel (abas obrigatórias, colunas, grupos)
2. Validação de dados contra as opções válidas do SharePoint (via scripts PowerShell)
3. Criação automática de registros no SharePoint, agrupados por `GRUPO`
4. Geração de relatório Excel formatado com status de cada item submetido

## Stack Tecnológico

| Componente | Tecnologia | Versão / Notas |
|---|---|---|
| Servidor Web | Python / Flask | Com `threaded=True` |
| Processamento Excel | openpyxl | Leitura e escrita de `.xlsx` |
| Integração SharePoint | PowerShell 5.1 + PnP | Scripts externos invocados via `subprocess` |
| Frontend | HTML + Jinja2 | Servido como template Flask |
| Execução de scripts | `subprocess.Popen` | Com janela oculta no Windows |

## Dependências (`requirements.txt`)

```
flask
openpyxl
customtkinter
packaging
```

> **Nota**: `customtkinter` e `packaging` aparecem no `requirements.txt` mas não são utilizados pelo backend web — indicam uso anterior em uma interface desktop (GUI).

## Estrutura de Diretórios

```
SolicitadorMobilizacao/
├── api/
│   ├── __init__.py
│   └── routes.py           ← Blueprint Flask com todos os endpoints HTTP
├── backend/
│   ├── __init__.py         ← Docstring de pacote
│   ├── server.py           ← Entrypoint da aplicação
│   ├── app.py              ← Fábrica da instância Flask
│   ├── config.py           ← Constantes globais e controle de jobs (threading)
│   ├── services.py         ← Camada de compatibilidade (re-exports)
│   ├── state.py            ← Contador de jobs ativos (threading)
│   ├── excel_processor.py  ← Leitura, agrupamento e manipulação de Excel
│   ├── encoding_repair.py  ← Reparo de mojibake em saída PowerShell
│   ├── powershell_executor.py ← Construção de comandos PowerShell
│   ├── correction_helper.py   ← Sugestões de correção e validação de arquivo
│   ├── report_generator.py    ← Geração de relatório Excel estilizado
│   └── data/
│       ├── templates/      ← Template Excel e status de atualização
│       └── uploads/        ← Arquivos de upload e relatórios gerados
├── frontend/
│   ├── templates/index.html  ← SPA completa (HTML + CSS + JS inline)
│   └── static/
│       └── Vestas Primary Logo RGB.png
└── docs/                   ← Esta documentação
```

## Como navegar nesta documentação

- **[Arquitetura](./architecture.md)**: visão macro, diagrama de componentes, fluxo ponta a ponta
- **[Files](./files/server.md)**: documentação detalhada de cada arquivo Python
- **[Módulos](./modules/overview.md)**: agrupamento funcional e integrações externas
- **[Melhorias](./improvements.md)**: problemas identificados, dívidas técnicas e sugestões
- **[Frontend](./files/index-html.md)**: documentação do `index.html` — design system, fluxo de streaming, correção guiada
