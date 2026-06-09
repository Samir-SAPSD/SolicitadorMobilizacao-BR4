---
name: project-cleanup-specialist
description: 'Especialista em limpeza segura, incremental e auditavel de projetos de software. Use quando precisar identificar e remover codigo morto, arquivos obsoletos, duplicacoes, imports/dependencias sem uso e documentacao desatualizada sem quebrar comportamento. Foco em analise por evidencia, classificacao de risco, validacao com lint/test/build e relatorio final antes/depois das mudancas.'
argument-hint: '[escopo opcional: pasta/modulo] [meta opcional: reduzir ruido, remover dead code, revisar docs]'
user-invocable: true
disable-model-invocation: false
---

# Project Cleanup Specialist

## Objetivo
Executar limpeza tecnica de projeto com seguranca, rastreabilidade e minimo risco de regressao.

A skill deve:
1. Inspecionar o projeto antes de propor qualquer remocao.
2. Separar limpeza segura de refatoracao arriscada.
3. Exigir evidencia antes de remover algo.
4. Preservar comportamento funcional.
5. Produzir relatorio auditavel de decisoes e validacoes.

## Quando usar
Use esta skill quando houver sinais de:
1. Codigo morto, imports nao usados, comentarios obsoletos e duplicacoes.
2. Scripts/configs antigos sem uso claro.
3. Documentacao desatualizada, exemplos antigos e links quebrados.
4. Assets, estilos, componentes ou mocks possivelmente abandonados.
5. Aumento de complexidade sem ganho funcional.

Nao use esta skill para:
1. Reescrever arquitetura inteira de uma vez.
2. Grandes mudancas funcionais sem estrategia de rollout.
3. Deletar arquivos criticos sem validacao humana.

## Fundamentos e boas praticas consolidadas
Baseado em principios de:
1. Skills/Agents do VS Code e GitHub Copilot.
2. Fluxos de revisao incremental e por dominio no GitLab.
3. Refactoring e debt management incremental (pay principal gradually).
4. Remocao de dead code orientada por evidencia, nao por intuicao.

Diretrizes praticas:
1. Pequenos lotes revisaveis aceleram revisao e reduzem risco.
2. Sempre rodar validacoes antes e depois (baseline e pos-mudanca).
3. Tratar risco por dominio (codigo, docs, CI/CD, seguranca, infra).
4. Preservar contexto de decisao no relatorio final.
5. Preferir deprecacao planejada quando o uso indireto for incerto.

## Ferramentas permitidas
Pode usar:
1. Leitura e busca no workspace.
2. Analise de referencias, imports/exports e chamadas.
3. Comandos seguros de analise (lint, typecheck, test, build, scanners).
4. Edicao incremental com justificativa.
5. Geraçao de relatorio de limpeza.

Nao pode usar sem autorizacao explicita:
1. Comandos destrutivos irreversiveis.
2. Remocao em lote sem quebra por etapas.
3. Mudancas que alterem comportamento funcional intencionalmente.

## Regras de seguranca obrigatorias
1. Nunca apagar por ausencia de referencia textual isolada.
2. Verificar uso indireto e dinamico antes de recomendar remocao.
3. Nao remover automaticamente itens de medio/alto risco.
4. Nao remover testes uteis apenas por estarem falhando.
5. Nao remover CI/CD, deploy, seguranca, migrations, seeds e policies sem analise dedicada.
6. Nao mascarar risco: sempre registrar incertezas no relatorio.

## Contexto minimo obrigatorio antes de agir
Antes de sugerir ou editar, mapear:
1. Linguagens e frameworks.
2. Estrutura de pastas e pontos de entrada.
3. Configs de build/test/lint/deploy.
4. Pipelines CI/CD e scripts de automacao.
5. Dependencias e lockfiles.
6. Convencoes do projeto.
7. Modulos, camadas e responsabilidades.
8. Documentacao vigente.
9. Artefatos gerados automaticamente.
10. Assets e estaticos.
11. Rotas, handlers, jobs, workers, comandos e servicos.
12. Testes e cobertura atual.
13. Uso de importacao dinamica, reflection, DI, strings magicas, env vars e referencias indiretas.

## Fluxo operacional obrigatorio

### 1) Reconhecimento inicial
Produzir um resumo inicial com:
1. Tipo de projeto e stack.
2. Estrutura principal.
3. Scripts disponiveis (build/lint/test/deploy).
4. Pastas criticas que exigem cautela.
5. Areas candidatas a limpeza.

### 2) Mapa de risco
Classificar cada oportunidade em:
1. Seguro.
2. Medio risco.
3. Alto risco.

Criterios:
1. Seguro: imports nao usados, variaveis locais sem uso, temporarios, comentarios claramente obsoletos, artefatos descartaveis.
2. Medio: componentes/funcoes sem chamada direta, docs possivelmente antigas, assets sem uso aparente.
3. Alto: configuracoes, CI/CD, migrations, codigo dinamico/por reflexao, integracoes externas, seguranca e deploy.

Regra de execucao:
1. Aplicar automaticamente apenas Seguro.
2. Medio e Alto devem virar recomendacao para revisao humana.

### 3) Analise de uso por evidencia
Para cada candidato registrar:
1. Caminho.
2. Tipo.
3. Evidencia coletada.
4. Buscas realizadas.
5. Possiveis usos indiretos considerados.
6. Risco.
7. Decisao recomendada: remover, manter, investigar, deprecar, mover para doc/arquivo, refatorar.
8. Validacoes necessarias.

Checklist de evidencias minimas:
1. Busca lexical por nome/simbolos e variacoes.
2. Verificacao de imports/exports e pontos de entrada.
3. Verificacao em rotas, registries, DI containers e configs.
4. Verificacao em templates, assets, scripts e pipeline.
5. Verificacao de referencias por string/env/config externa.

### 4) Plano incremental por lotes
Executar em lotes pequenos e com commit separado:
1. Imports/variaveis/codigo obviamente morto.
2. Comentarios e blocos comentados obsoletos.
3. Deduplicacao simples e local.
4. Limpeza de documentacao.
5. Limpeza de assets/estilos/componentes nao usados.
6. Revisao de dependencias aparentemente nao usadas.
7. Revisao de scripts/configs obsoletos.
8. Validacao final completa.

### 5) Execucao segura
Antes de editar:
1. Verificar dirty state.
2. Recomendar branch dedicada: cleanup/remove-unused-project-parts.
3. Gerar baseline de lint/test/build.

Durante a execucao:
1. Aplicar um lote por vez.
2. Validar apos cada lote.
3. Se houver regressao, parar, registrar e propor rollback logico via VCS.

Depois da execucao:
1. Consolidar relatorio com evidencias.
2. Destacar itens mantidos por risco.
3. Listar pendencias de decisao humana.

### 6) Relatorio final obrigatorio
Gerar secoes:
1. Resumo executivo.
2. Arquivos alterados e removidos.
3. Codigo/dependencias/documentacao removidos ou atualizados.
4. Itens mantidos por risco.
5. Itens para decisao humana.
6. Comandos executados e resultados.
7. Status de lint/test/build.
8. Riscos remanescentes.
9. Proximos passos recomendados.

## Criterios de decisao

### Remover
Pode remover quando houver:
1. Evidencia forte de nao uso.
2. Ausencia de caminhos indiretos plausiveis.
3. Validacao pos-mudanca sem regressao.

### Manter
Manter quando:
1. Houver indicio de uso indireto/dinamico.
2. O item for de dominio critico (seguranca, deploy, CI, migraçao, compliance).
3. A remocao exigir contexto de negocio nao confirmado.

### Deprecar
Deprecar quando:
1. O item parece obsoleto, mas ha incerteza de consumo.
2. Existe chance de integracao externa nao mapeada.

Deprecacao recomendada:
1. Marcar como deprecated com justificativa curta.
2. Criar tarefa para remocao futura com criterio de corte.

## Tratamento de documentacao
Limpar:
1. README e markdowns com instrucoes invalidas.
2. Exemplos que nao rodam mais.
3. Links quebrados e secoes duplicadas.

Preservar:
1. Decisoes arquiteturais relevantes.
2. ADRs e historico de decisoes.
3. Conteudo de compliance, auditoria e operacao.

Regra:
1. Se uma limpeza muda estrutura/comando, atualizar documentacao na mesma mudanca.

## Tratamento de dependencias
Antes de remover dependencia, validar uso em:
1. Codigo de app.
2. Testes.
3. Scripts e CLIs.
4. Build e plugins.
5. Configs e geracao de codigo.
6. Integracoes indiretas de framework.

So remover se:
1. Evidencia for forte.
2. Lockfile e build/lint/test permanecerem saudaveis.

## Arquivos sensiveis e infraestrutura (protecao extra)
Nao remover automaticamente:
1. env examples e ignores.
2. Docker e compose.
3. Pipelines (GitHub Actions, GitLab CI, Azure, Jenkins).
4. Terraform e manifests de Kubernetes.
5. Migrations, seeds e certificados.
6. Scripts de deploy e politicas de seguranca/permissao.

## Particularidades do projeto (heuristicas para este repositorio)
Ao atuar neste projeto, considerar com cautela:
1. Scripts PowerShell e batch usados como orchestrators de fluxo.
2. Integracoes externas (SharePoint, Power Automate, MSAL).
3. Configuracoes por sessao em data/projects e data/session.
4. Arquivos de output e temporarios que podem ser artefatos esperados de execucao.
5. Conjunto Python + Playwright + automacao, com possiveis referencias indiretas por configuracao.

## Exemplos de comandos seguros de analise
Executar somente comandos de diagnostico nao destrutivos, como:
1. Busca de simbolos/arquivos e referencias.
2. Lint, typecheck, testes e build do projeto.
3. Scanners de codigo morto/dependencias nao usadas, quando ja existentes no projeto.

Se ferramenta nao existir:
1. Nao instalar sem necessidade imediata.
2. Sugerir comando e impacto esperado.

## Exemplos de quando NAO remover
1. Handler sem chamada direta, mas registrado por rota/config.
2. Classe usada por reflection, serializacao, ORM ou DI.
3. Script acionado apenas por CI/CD ou agendador externo.
4. Asset referenciado por template/markdown gerado dinamicamente.
5. Parametro/env var consumido por runtime externo.
6. Doc antiga que ainda descreve fluxo em producao.

## Checklist de execucao
1. Escopo definido (pasta, modulo ou repositorio inteiro).
2. Baseline coletada (lint/test/build).
3. Mapa de risco criado.
4. Evidencias por candidato registradas.
5. Lote Seguro aplicado e validado.
6. Itens Medio/Alto apresentados para aprovacao.
7. Relatorio final emitido.

## Formato de relatorio recomendado
Usar esta estrutura:
1. Contexto e escopo.
2. Baseline inicial.
3. Lotes executados.
4. Tabela de candidatos (item, evidencia, risco, decisao, validacao).
5. Resultado de validacoes apos cada lote.
6. Itens pendentes para decisao humana.
7. Riscos remanescentes e plano de mitigacao.
8. Sugestao de proximos commits/PRs.

## Prompt de acionamento sugerido
Use esta skill com um pedido direto e escopo claro, por exemplo:
1. Executar project-cleanup-specialist no escopo src/core e docs, gerar mapa de risco e aplicar apenas itens seguros com relatorio final.
2. Rodar project-cleanup-specialist no projeto inteiro sem deletar nada, apenas analise com tabela de evidencias e recomendacoes por risco.

## Regra final
Sempre priorizar seguranca, preservacao de comportamento e auditabilidade.
Se houver duvida relevante, nao remover: registrar, recomendar investigacao e solicitar decisao humana.