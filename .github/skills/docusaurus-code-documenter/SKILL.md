---
name: docusaurus-code-documenter
description: >-
  Especialista em análise profunda de código-fonte e geração de documentação técnica completa no padrão Docusaurus.
  Use quando precisar: documentar um repositório inteiro; gerar docs para Docusaurus; explicar arquitetura de sistemas;
  mapear dependências entre módulos; documentar APIs, funções, classes e integrações; entender decisões de engenharia;
  gerar diagramas de fluxo (Mermaid); identificar anti-padrões e sugerir melhorias. Produz arquivos .md prontos para
  publicação com intro, architecture, módulos por arquivo, funções, classes e integrações.
argument-hint: 'Caminho do repositório ou escopo (ex: backend/, módulo específico, repositório completo)'
---

# Docusaurus Code Documenter

## Visão Geral

Esta Skill analisa código-fonte em profundidade e gera documentação técnica completa, estruturada no padrão do Docusaurus. O foco não é apenas descrever *o que* o código faz, mas explicar *por que* cada decisão foi tomada, como os componentes se relacionam e qual é o fluxo de execução ponta a ponta.

## Quando Usar

- Documentar um repositório inteiro do zero
- Gerar ou atualizar docs para publicação em Docusaurus
- Explicar arquitetura e decisões de engenharia para novos membros
- Mapear dependências e integrações entre módulos
- Identificar anti-padrões, dívidas técnicas e pontos de melhoria
- Criar material de referência técnica de nível profissional

---

## Fluxo de Execução

Siga rigorosamente estas etapas em sequência:

### Fase 1 — Ingestão e Mapeamento Estrutural

1. **Listar todos os arquivos** do escopo definido (recursivamente), incluindo:
   - Arquivos de código-fonte (`.py`, `.ts`, `.js`, `.cs`, etc.)
   - Arquivos de configuração (`requirements.txt`, `package.json`, `.env.example`, `config.*`)
   - Scripts de automação e utilitários
   - Templates e assets relevantes

2. **Construir mapa de estrutura de diretórios** com anotações do papel de cada pasta.

3. **Detectar o stack tecnológico**: linguagens, frameworks, bibliotecas principais, banco de dados, serviços externos.

4. **Identificar o ponto de entrada** do sistema (ex: `main.py`, `server.py`, `app.py`, `index.ts`).

Consulte o guia detalhado: [análise estrutural](./references/structural-analysis.md)

---

### Fase 2 — Análise Profunda por Arquivo

Para **cada arquivo de código**, execute:

1. **Leitura completa** — nunca resumir ou truncar.
2. **Extração de entidades**:
   - Funções e métodos (assinatura, parâmetros, retorno)
   - Classes e seus atributos
   - Constantes e configurações globais
   - Decorators, middlewares e hooks
3. **Análise de lógica interna**:
   - Qual problema este arquivo resolve?
   - Quais são os fluxos principais de execução?
   - Existem ramificações condicionais importantes? Por quê?
   - Há tratamento de erros? É adequado?
4. **Identificar importações e dependências** (internas e externas).
5. **Detectar padrões arquiteturais** usados (Repository, Service Layer, Factory, Singleton, etc.).
6. **Registrar anti-padrões** encontrados (God Class, código duplicado, acoplamento excessivo, etc.).

---

### Fase 3 — Mapeamento de Dependências e Integrações

1. Construir **grafo de dependências** entre módulos internos.
2. Identificar todas as **integrações externas**:
   - APIs REST/GraphQL chamadas
   - Bancos de dados e ORMs
   - Sistemas de mensageria (queues, webhooks)
   - Serviços de cloud (SharePoint, Azure, AWS, etc.)
   - Ferramentas de automação (PowerShell, bash, etc.)
3. Mapear **fluxo de dados**: onde entram, como são transformados, onde saem.
4. Identificar **pontos de falha** e dependências críticas.

---

### Fase 4 — Geração da Documentação

Gere os arquivos seguindo o [padrão de saída](./references/output-format.md).

**Regras obrigatórias de documentação:**

- Cada arquivo de código → uma página `.md` própria
- Cada função/método → seção dedicada com: propósito, parâmetros, retorno, lógica interna e *por que* foi implementada assim
- Cada classe → seção com: responsabilidade, atributos, métodos, padrão utilizado
- Nenhum bloco lógico relevante pode ser omitido
- Sempre incluir blocos de código com syntax highlighting
- Usar diagramas Mermaid para fluxos complexos, arquitetura e dependências
- Conectar explicitamente cada componente aos outros que interagem com ele

---

### Fase 5 — Revisão e Sugestões

Após gerar a documentação, produzir uma seção final `improvements.md` com:

1. **Problemas identificados**: bugs potenciais, race conditions, falta de validação
2. **Dívidas técnicas**: código legado, dependências desatualizadas, falta de testes
3. **Sugestões de refatoração**: com justificativa técnica
4. **Boas práticas ausentes**: tratamento de erros, logging, segurança, performance
5. **Próximos passos recomendados**: priorizados por impacto

---

## Comportamento da IA

### Como Analisar Código

- Leia o código **como um engenheiro sênior revisando uma PR**: questione cada decisão.
- Não aceite código "funcional" como suficiente — avalie qualidade, manutenibilidade e segurança.
- Ao encontrar lógica complexa, **decomponha passo a passo** antes de documentar.
- Ao encontrar um nome ambíguo (função, variável), **infira o propósito pelo contexto de uso**.

### Como Deduzir Regras de Negócio

- Analise nomes de funções + parâmetros + corpo para inferir regras implícitas.
- Observe validações, condicionais e mensagens de erro — eles revelam invariantes do negócio.
- Conecte regras de negócio a casos de uso reais do sistema.

### Como Identificar Padrões Arquiteturais

| Sinal no Código | Padrão Provável |
|---|---|
| Classe com apenas um método estático | Utility / Helper |
| Classe que orquestra outras sem lógica própria | Facade / Service |
| Herança + método abstrato | Template Method |
| Injeção de dependência via construtor | Dependency Injection |
| Arquivo de config centralizado | Singleton Configuration |
| Separação clara entre rotas e lógica | MVC / Layered Architecture |

### Como Explicar Decisões de Engenharia

Sempre responda as três perguntas:
1. **O quê?** — O que este código faz concretamente.
2. **Como?** — O mecanismo técnico utilizado.
3. **Por quê?** — A motivação da decisão (performance, legibilidade, restrição externa, histórico).

---

## Referências

- [Análise Estrutural e Extração de Entidades](./references/structural-analysis.md)
- [Formato de Saída Docusaurus](./references/output-format.md)
