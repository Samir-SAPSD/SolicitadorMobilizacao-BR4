---
name: code-refactoring-specialist
description: 'Especialista em refatoracao segura, incremental e orientada a evidencias. Use quando precisar reduzir duplicacao de codigo, simplificar condicionais complexas, melhorar nomes, separar responsabilidades, extrair abstracao reutilizavel, reduzir complexidade ciclomatica ou melhorar a manutenibilidade sem alterar comportamento funcional. Cobre Python, PowerShell, batch e qualquer stack do projeto.'
argument-hint: '[escopo opcional: pasta/modulo/arquivo] [objetivo opcional: simplificar condicionais, extrair funcoes, renomear, separar responsabilidades]'
user-invocable: true
disable-model-invocation: false
---

# Code Refactoring Specialist

## Principio central

> Refatoracao e a pratica de melhorar a estrutura interna do codigo sem alterar seu comportamento externo.
> (Martin Fowler, Refactoring: Improving the Design of Existing Code, 2018)

A skill nao adiciona funcionalidades, nao remove comportamentos e nao altera contratos funcionais.
Cada mudanca deve ser comportamentalmente equivalente a versao anterior.

## Quando usar

Use esta skill quando houver sinais de:
1. Funcoes ou metodos longos demais com multiplas responsabilidades.
2. Duplicacao de logica entre arquivos, funcoes ou modulos.
3. Condicionais aninhadas dificeis de entender.
4. Nomes de variaveis, funcoes ou classes pouco descritivos.
5. Parametros em excesso ou listas de argumentos repetidos em multiplos lugares.
6. Acoplamento alto entre modulos.
7. Coesao baixa dentro de modulos.
8. Codigo dificil de testar por mistura de responsabilidades.
9. Constantes magicas espalhadas pelo codigo.
10. Padrao de implementacao inconsistente entre partes do projeto.

Nao use esta skill para:
1. Reescrever o projeto inteiro de uma vez.
2. Trocar tecnologia, framework ou biblioteca.
3. Adicionar novas funcionalidades.
4. Alterar regras de negocio.
5. Mudar APIs publicas ou contratos sem analise de impacto.

## Catalogo de tecnicas aplicaveis

### Composicao de metodos
1. Extract Method / Extract Function: extrair bloco logico em funcao nomeada.
2. Inline Method / Inline Function: remover funcao trivial que apenas delega.
3. Extract Variable: nomear expressao complexa.
4. Inline Temp: remover variavel temporaria desnecessaria.
5. Replace Temp with Query: substituir variavel temporaria por chamada de funcao.
6. Split Temporary Variable: separar variavel reutilizada para fins diferentes.
7. Replace Method with Method Object: transformar metodo longo em classe propria.
8. Substitute Algorithm: substituir algoritmo por versao mais clara e equivalente.

### Mover funcionalidade
1. Move Method / Move Function: mover para o modulo com maior afinidade.
2. Extract Class: extrair responsabilidade em nova classe ou modulo.
3. Inline Class: eliminar classe desnecessaria que apenas repassa.
4. Remove Middle Man: eliminar excesso de delegacao.

### Organizar dados
1. Replace Magic Number with Named Constant: nomear literais numericos.
2. Encapsulate Collection: encapsular acesso direto a colecoes mutaveis.
3. Introduce Parameter Object: agrupar parametros relacionados em objeto/dataclass.

### Simplificar condicionais
1. Replace Nested Conditional with Guard Clauses: retornar cedo para eliminar aninhamento.
2. Decompose Conditional: extrair expressoes complexas em funcoes nomeadas.
3. Consolidate Conditional Expression: unir condicionais com mesmo resultado.
4. Consolidate Duplicate Conditional Fragments: mover codigo repetido para fora do condicional.
5. Simplify Boolean Expression: simplificar expressoes booleanas redundantes.
6. Remove Control Flag: substituir flag de controle por break/return/throw.

### Simplificar chamadas
1. Rename Method / Rename Variable: melhorar nome para revelar intencao.
2. Separate Query from Modifier: separar funcao que consulta de funcao que altera estado.
3. Parameterize Method: unificar funcoes similares com parametro.
4. Remove Parameter: eliminar parametro nunca usado.
5. Replace Parameter with Method Call: calcular internamente em vez de receber como argumento.
6. Replace Error Code with Exception: substituir codigo de erro por excecao.
7. Replace Exception with Test: substituir try/catch evitavel por verificacao previa.

### Generalizar
1. Pull Up Method: mover metodo duplicado para ancestral comum.
2. Push Down Method: mover para subclasse quando relevante apenas para ela.
3. Extract Superclass / Extract Interface: consolidar comportamento compartilhado.
4. Replace Inheritance with Delegation: preferir composicao quando heranca e abuso.

## Contexto obrigatorio antes de agir

Antes de propor ou executar qualquer refatoracao, mapear:
1. Linguagens e versoes.
2. Frameworks e bibliotecas principais.
3. Estrutura de pastas e convencoes de nomenclatura.
4. Padroes arquiteturais: camadas, modulos, responsabilidades.
5. Pontos de entrada da aplicacao.
6. Scripts disponiveis: build, lint, typecheck, test, format.
7. Ferramentas de CI/CD.
8. Padrao de testes existentes.
9. Estilo de tratamento de erros e logs.
10. Uso de tipagem estatica ou dinamica.
11. Uso de imports dinamicos, reflection, decorators, DI ou configuracao em runtime.
12. Dependencias criticas e suas versoes.
13. Arquivos gerados automaticamente.
14. Areas sensiveis: seguranca, autenticacao, autorizacao, criptografia, migrations, deploy, infra.

## Fluxo obrigatorio

### 1. Reconhecimento inicial

Inspecionar o projeto sem fazer alteracoes. Produzir:
1. Tipo do projeto e stack.
2. Linguagens, frameworks e versoes detectados.
3. Estrutura principal e convencoes detectadas.
4. Padroes arquiteturais percebidos.
5. Scripts disponiveis.
6. Ferramentas de validacao: test, lint, typecheck, build.
7. Areas candidatas a refatoracao com justificativa.
8. Areas criticas que exigem cautela extra.

Nenhuma alteracao deve ser feita nesta etapa.

### 2. Baseline de validacao

Antes de qualquer mudanca, identificar e executar os comandos de validacao disponiveis.

Exemplos:
```
pytest
pytest -q
npm test
npm run lint
npm run typecheck
python -m mypy src/
flake8 src/
```

Registrar o baseline:
1. Testes passando ou falhando.
2. Erros de lint e tipo.
3. Cobertura, se disponivel.

Se nenhuma ferramenta estiver disponivel, registrar ausencia e propor adicao de validacao antes de refatorar.

### 3. Catalogo de oportunidades de refatoracao

Para cada oportunidade identificada, registrar:
1. Arquivo e linha aproximada.
2. Code smell detectado com categoria.
3. Tecnica de refatoracao recomendada.
4. Justificativa clara e objetiva.
5. Risco de regressao: Baixo, Medio ou Alto.
6. Dependencias e referencias identificadas.
7. Decisao recomendada: aplicar agora, adiar, investigar ou rejeitar.

Classificacao de risco:
1. Baixo: rename, extrair variavel local, substituir magic number, simplificar expressao booleana sem side effects.
2. Medio: extrair funcao/classe, mover funcao, simplificar condicional, inline class, introduce parameter object.
3. Alto: alterar assinatura publica, alterar modulo compartilhado por multiplos contextos, alterar fluxo de controle principal, mover arquivos entre modulos.

### 4. Plano incremental por lotes

Agrupar refatoracoes em lotes pequenos e logicamente coerentes:
1. Renomeacoes e constantes: rename, magic numbers, extract variable.
2. Simplificacao de condicionais: guard clauses, decompose conditional.
3. Extracao de funcoes: extract method, extract variable, replace temp with query.
4. Mover responsabilidades: move method, extract class, remove middle man.
5. Limpeza de interfaces: separate query from modifier, parameterize method, introduce parameter object.
6. Validacao final.

Cada lote deve:
1. Ter commit semantico separado.
2. Passar baseline de validacao antes e depois.
3. Ser pausado se houver regressao.

### 5. Execucao segura

Antes de editar:
1. Verificar se ha alteracoes pendentes no repositorio.
2. Recomendar branch dedicada: refactor/nome-do-escopo.
3. Registrar baseline de testes e lint.

Durante a execucao:
1. Aplicar um lote por vez.
2. Validar apos cada lote.
3. Se houver regressao: parar, registrar e propor rollback via VCS.
4. Preservar comportamento externo em cada passo.
5. Nao misturar refatoracao com correcao de bug ou nova funcionalidade.

Depois da execucao:
1. Consolidar relatorio com evidencias.
2. Destacar itens mantidos por risco.
3. Listar pendencias para decisao humana.

### 6. Relatorio final obrigatorio

Estrutura do relatorio:
1. Resumo executivo.
2. Baseline antes e depois (testes, lint, typecheck, build).
3. Lotes executados com lista de mudancas por lote.
4. Tabela de candidatos: arquivo, code smell, tecnica, risco, decisao, validacao.
5. Itens mantidos por risco.
6. Itens para decisao humana.
7. Riscos remanescentes.
8. Proximos commits/PRs recomendados.

## Criterios de decisao por risco

### Aplicar (Baixo risco)
1. Mudanca estritamente local e sem efeito externo.
2. Evidencia clara de equivalencia comportamental.
3. Nenhuma interface publica afetada.
4. Testes passam apos mudanca.

### Propor para aprovacao humana (Medio risco)
1. Mudanca afeta mais de um arquivo.
2. Interface publica pode ser afetada indiretamente.
3. Logica de dominio envolvida.
4. Teste de caracterizacao necessario antes de proceder.

### Nao aplicar automaticamente (Alto risco)
1. Alterar assinatura de funcao publica sem mapeamento de todos os chamadores.
2. Mover modulo que e consumido dinamicamente.
3. Alterar comportamento de tratamento de erro em camada de integracao.
4. Qualquer mudanca em autenticacao, autorizacao, criptografia ou seguranca.
5. Qualquer mudanca em scripts de deploy, migrations ou pipelines.

## Regras de seguranca obrigatorias

1. Nunca refatorar e adicionar funcionalidade no mesmo commit.
2. Nunca alterar comportamento externo sem autorizacao explicita.
3. Nunca alterar contratos de API publica sem analise de impacto.
4. Nunca alterar arquivos gerados automaticamente.
5. Nunca alterar pipelines, migrations ou infra sem analise especifica.
6. Nunca usar comandos destrutivos irreversiveis.
7. Se houver duvida sobre equivalencia comportamental, nao refatorar: registrar e solicitar decisao humana.

## Particularidades do projeto (heuristicas para este repositorio)

Ao atuar neste projeto, considerar com cautela:
1. Scripts PowerShell e batch sao orchestrators de fluxo com multiplos chamadores e caminhos dinamicos.
2. Integracoes externas com SharePoint, Power Automate e MSAL usam configuracao via env vars.
3. Modulos Python com imports condicionais podem ser usados em contextos diferentes.
4. Configuracoes por sessao em data/projects e data/session sao consumidas dinamicamente.
5. Artefatos de output sao esperados por integradores externos.
6. Caminho de diretorio de trabalho em scripts batch e critico para execucao correta.

## Exemplos de prompts de acionamento

1. `code-refactoring-specialist` no escopo `src/core`: reconhecimento inicial, catalogo de oportunidades e aplicar apenas Baixo risco com relatorio.
2. `code-refactoring-specialist` no arquivo `scripts/scrape-automate.ps1`: identificar funcoes longas, condicionais complexas e propor plano de extracao.
3. `code-refactoring-specialist` sem escopo: analise completa do projeto, apenas catalogo de oportunidades por risco, sem aplicar mudancas.

## Regra final

Sempre priorizar seguranca, equivalencia comportamental e incrementalidade.
Se houver duvida relevante sobre o impacto de uma mudanca, nao refatorar: registrar, justificar e solicitar decisao humana.
O melhor refactoring e o que torna o codigo mais claro para a proxima pessoa que vai ler, sem surpresas em runtime.
