# Inventário e Classificação de Arquivos

## Algoritmo de Inventário

Ao receber um caminho (`docs/` ou outro), execute:

```bash
# Listar todos os .md recursivamente
find <docs-path> -name "*.md" | sort
```

No Windows PowerShell:
```powershell
Get-ChildItem -Path <docs-path> -Recurse -Filter "*.md" | Select-Object FullName
```

---

## Extração de Metadados de Frontmatter

Para cada `.md`, extrair o bloco YAML entre os marcadores `---`:

```
---
sidebar_position: 2
title: Arquitetura
---
```

Se não houver frontmatter:
- `title`: usar o primeiro `# Heading` do arquivo
- `sidebar_position`: usar ordem alfabética do arquivo

---

## Mapa de Classificação de Categorias

| Caminho do arquivo | Categoria Docusaurus | Label no sidebar |
|---|---|---|
| `docs/intro.md` | raiz | — (item direto) |
| `docs/architecture.md` | raiz | — (item direto) |
| `docs/improvements.md` | raiz | — (item direto) |
| `docs/files/*.md` | `files` | "Arquivos de Código" |
| `docs/modules/*.md` | `modules` | "Módulos" |
| `docs/scripts/*.md` | `scripts` | "Scripts" |
| `docs/integrations/*.md` | `integrations` | "Integrações" |
| `docs/classes/*.md` | `classes` | "Classes" |
| `docs/functions/*.md` | `functions` | "Funções" |
| Qualquer outro subdiretório | nome do diretório | Capitalizar nome |

---

## Detecção do Nome do Projeto

1. Abrir `docs/intro.md`
2. Encontrar o primeiro `# Heading`
3. Extrair o texto após `#`
4. Usar como `title` no `docusaurus.config.js`
5. Sanitizar para `project-name` (lowercase, hífens): usar como nome da pasta do projeto

**Exemplo**:
```markdown
# Solicitador de Mobilização — Documentação Técnica
```
→ `title`: `"Solicitador de Mobilização"`
→ `project-name`: `solicitador-mobilizacao-docs`

---

## Regras de Ordenação Final

```
1. intro.md                  (sidebar_position: 1)
2. architecture.md           (sidebar_position: 2)
3. improvements.md           (sidebar_position: 3)
4. [Categoria: files/]       (sidebar_position: 4, se existir)
5. [Categoria: modules/]     (sidebar_position: 5, se existir)
6. [demais categorias]       (ordem alfabética)
```

Dentro de cada categoria, respeitar `sidebar_position` dos arquivos individuais.
