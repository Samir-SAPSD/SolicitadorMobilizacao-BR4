---
sidebar_position: 2
title: __init__.py (api)
---

# `api/__init__.py`

## Responsabilidade

> Marcador de pacote Python para o módulo `api`. Contém apenas a docstring do pacote.

## Localização

```
api/__init__.py
```

## Conteúdo

```python
"""HTTP API layer for the Mobilization request app."""
```

Este arquivo declara o diretório `api/` como um pacote Python importável. A docstring `"""HTTP API layer..."""` serve como descrição do pacote — visível em ferramentas como `help(api)` e geradores de documentação automática.

**Por que é necessário?** Sem este arquivo, `from api.routes import api_bp` em `backend/app.py` lançaria `ModuleNotFoundError`. O Python trata qualquer diretório com `__init__.py` como um pacote importável.

**Por que está vazio (apenas docstring)?** O pacote `api` não precisa exportar nada no nível do pacote — o único símbolo utilizado externamente é `api_bp`, importado diretamente de `api.routes`. Inicializações em `__init__.py` seriam executadas a cada importação do pacote, o que seria desnecessário aqui.

## Interações com Outros Módulos

```mermaid
graph LR
    app[backend/app.py] -->|from api.routes import api_bp| init[api/__init__.py\npacote]
    init --> routes[api/routes.py]
```
