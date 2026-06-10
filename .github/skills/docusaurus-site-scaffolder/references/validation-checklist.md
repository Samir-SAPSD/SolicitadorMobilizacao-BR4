# Checklist de Validação

## Pré-Build

Execute antes de `npm run build`:

- [ ] Todos os arquivos `.md` têm frontmatter com `title`
  - Verificar: arquivos sem `title` usarão o primeiro `#` heading (Docusaurus faz isso automaticamente, mas é melhor ter frontmatter explícito)
- [ ] Nenhum `sidebar_position` duplicado dentro do mesmo diretório
- [ ] `sidebars.js` referencia apenas arquivos que existem em `docs/`
- [ ] `docusaurus.config.js` tem `title`, `url` e `baseUrl` preenchidos (sem placeholder `{{...}}`)
- [ ] `src/pages/index.js` tem links que apontam para rotas válidas
- [ ] Node.js versão >= 18: `node --version`

---

## Validação de Links no `npm run build`

O Docusaurus detecta links quebrados automaticamente durante o build. Configuração recomendada em `docusaurus.config.js`:

```js
onBrokenLinks: 'warn',          // 'throw' em produção; 'warn' durante desenvolvimento
onBrokenMarkdownLinks: 'warn',
```

**Interpretar a saída do build**:

```
[WARNING] Docs markdown link "./files/server.md" in "docs/modules/overview.md" resolves to the file "docs/files/server.md" which does not exist.
```

→ O link `./files/server.md` está incorreto. Em Docusaurus, links entre docs devem ser sem extensão ou usar o caminho relativo correto.

**Formato correto de links internos**:
```markdown
<!-- Errado -->
[server.py](./files/server.md)

<!-- Correto (sem extensão) -->
[server.py](./files/server)

<!-- Também correto (caminho absoluto do site) -->
[server.py](/files/server)
```

---

## Execução de Build de Validação

```bash
cd <project-name>
npm install
npm run build 2>&1 | tee build-output.txt
```

**Interpretar resultado**:

| Saída | Significado | Ação |
|---|---|---|
| `Generated static files in "build".` | ✅ Build bem-sucedido | Continuar |
| `[WARNING] Docs markdown link...` | ⚠️ Link quebrado | Corrigir `.md` referenciado |
| `[ERROR] ...` | ❌ Erro crítico | Verificar `docusaurus.config.js` e `sidebars.js` |
| `Cannot find module...` | ❌ Dependência faltando | `npm install` novamente |

---

## Verificação de Arquivos `.md` sem Frontmatter

Executar para detectar arquivos sem bloco `---`:

```powershell
# PowerShell
Get-ChildItem -Path docs -Recurse -Filter "*.md" | ForEach-Object {
    $content = Get-Content $_.FullName -Raw
    if (-not $content.StartsWith("---")) {
        Write-Warning "Sem frontmatter: $($_.FullName)"
    }
}
```

---

## Teste de Navegação Local

Após build bem-sucedido:

```bash
npm run start
```

Verificar manualmente:
- [ ] Homepage carrega sem erro
- [ ] Sidebar exibe todas as categorias
- [ ] Navegar até pelo menos 3 páginas de categorias diferentes
- [ ] Diagrama Mermaid renderiza (requer plugin `@docusaurus/theme-mermaid`)
- [ ] Blocos de código têm syntax highlighting

---

## Suporte a Mermaid (Diagramas)

Se os `.md` contêm blocos ` ```mermaid `, instalar o plugin:

```bash
npm install @docusaurus/theme-mermaid
```

Adicionar ao `docusaurus.config.js`:

```js
markdown: {
  mermaid: true,
},
themes: ['@docusaurus/theme-mermaid'],
```

---

## Relatório Final de Validação

Ao concluir, exibir este resumo:

```
╔══════════════════════════════════════════════╗
║   DOCUSAURUS SITE SCAFFOLDER — RELATÓRIO     ║
╠══════════════════════════════════════════════╣
║ Projeto criado em : ./<project-name>/        ║
║ Arquivos .md      : N                        ║
║ Categorias        : X                        ║
║ Links verificados : OK (ou N quebrados)       ║
║ Build status      : ✅ PASSOU / ❌ FALHOU     ║
╠══════════════════════════════════════════════╣
║ Para iniciar:                                ║
║   cd <project-name>                          ║
║   npm run start                              ║
╚══════════════════════════════════════════════╝
```
