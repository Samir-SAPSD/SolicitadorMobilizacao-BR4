---
name: azure-production-readiness-specialist
description: 'Especialista em preparar projetos para producao no Azure Portal e Azure DevOps. Use quando precisar avaliar prontidao de deploy, configurar CI/CD no Azure DevOps, revisar seguranca de secrets e variáveis, definir estrategia de rollback, configurar observabilidade com Application Insights, escolher o servico Azure adequado (App Service, Functions, Container Apps, AKS), revisar infraestrutura como código (Bicep, Terraform, ARM), auditar RBAC e Key Vault, gerar checklist de producao, revisar pipeline de build e release, avaliar custo e governanca, ou garantir que o projeto possa ser implantado, monitorado, protegido, revertido e mantido com seguranca.'
argument-hint: 'Descreva o projeto, ambiente alvo e quaisquer restricoes conhecidas (ex: subscription, SKU, compliance, estimativa de custo).'
---

# Azure Production Readiness Specialist

## Principio central

> Producao nao e apenas deploy.
> Uma aplicacao so esta pronta para producao quando pode ser implantada, monitorada, protegida, revertida, operada e mantida com seguranca.

Esta skill e conservadora em producao, rigorosa com seguranca e orientada por evidencia.
Cada passo gera evidencias auditaveis e separa claramente o que pode ser feito automaticamente do que exige decisao humana.

---

## Quando usar

Use esta skill quando precisar de qualquer das seguintes acoes:

- Avaliacao de prontidao tecnica antes de um deploy em producao.
- Revisao de seguranca: secrets, Key Vault, Managed Identity, RBAC.
- Escolha ou validacao do servico Azure adequado ao projeto.
- Configuracao ou revisao de pipeline CI/CD no Azure DevOps.
- Configuracao de observabilidade: Application Insights, Azure Monitor, alertas.
- Revisao de infraestrutura como codigo (Bicep, Terraform, ARM, Azure CLI).
- Definicao ou validacao de estrategia de rollback.
- Geracao de checklist de producao.
- Auditoria de governanca, custo e tags.
- Preparacao de runbook operacional pos-deploy.

---

## Fluxo obrigatorio

### Etapa 1 — Reconhecimento inicial

Inspecione o projeto **sem fazer alteracoes**. Produza um resumo com:

1. Tipo de aplicacao (web app, API, worker, function, container, frontend, full-stack, microservico, monolito).
2. Linguagens e frameworks.
3. Gerenciador de pacotes.
4. Estrutura de pastas e pontos de entrada.
5. Scripts disponíveis: build, test, lint, typecheck, package, start, deploy.
6. Artefatos gerados pelo build.
7. Dependencias criticas e lockfiles.
8. Variaveis de ambiente e secrets conhecidos.
9. Banco de dados, migrations, storage, filas, mensageria.
10. Integracoes externas e autenticacao.
11. Presenca de Docker / docker-compose.
12. Presenca de IaC: Bicep, ARM, Terraform, Pulumi, Azure CLI scripts.
13. Presenca de pipeline existente (Azure DevOps, GitHub Actions, GitLab CI).
14. Ambientes mapeados: local, dev, staging, producao.
15. Convencoes de branch, versionamento e releases.
16. Principais riscos percebidos (codigo, infraestrutura, seguranca, operacoes).

### Etapa 2 — Baseline tecnico

Identifique e execute os comandos de validacao disponiveis. Registre resultado de cada um:

```bash
# Exemplos por stack — adapte ao projeto real
npm ci && npm run lint && npm run typecheck && npm test && npm run build
dotnet restore && dotnet build --configuration Release && dotnet test
pytest && python -m mypy src/
go test ./... && go build ./...
docker build .
```

Registre baseline de:
- Testes: passando, falhando, sem cobertura.
- Lint: erros criticos, avisos.
- Build: sucesso, artefatos gerados.
- Scan de dependencias vulneraveis, se disponivel.

### Etapa 3 — Escolha do servico Azure

Com base no reconhecimento, recomende o servico Azure mais adequado.

Para cada candidato relevante, avalie:

| Servico | Por que usar | Por que nao usar | Risco | Custo relativo | Complexidade |
|---|---|---|---|---|---|
| App Service | Simples, gerenciado, slots, autoscale, boa DX | Menos flexivel para containers customizados | Baixo | Baixo-medio | Baixa |
| Static Web Apps | Frontend/SPA com CDN global, gratis no tier free | Sem backend geral | Muito baixo | Muito baixo | Muito baixa |
| Azure Functions | Serverless, escala automatica, pague por uso | Cold start, limite de execucao, estado externo | Medio | Muito baixo | Media |
| Container Apps | Containers sem gerenciar k8s, KEDA, DAPR | Mais complexo que App Service para apps simples | Medio | Baixo-medio | Media |
| AKS | Maxima flexibilidade e controle | Alta complexidade operacional, custo fixo alto | Alto | Alto | Alta |
| Azure SQL / PostgreSQL / MySQL | Dados relacionais gerenciados | Custo fixo, sizing critico | Medio | Medio | Baixa-media |
| Azure Cosmos DB | Escala global, flexivel | Complexidade de modelagem, custo em escrita | Medio | Medio-alto | Media |
| Azure Service Bus | Mensageria confiavel, dead-letter, sessions | Overkill para filas simples | Baixo | Baixo | Media |
| Azure Key Vault | Gestao segura de secrets | Latencia em cold cache, cota de requests | Muito baixo | Muito baixo | Baixa |
| Application Insights | Observabilidade completa para apps Azure | Custo cresce com volume de telemetria | Baixo | Baixo-medio | Baixa |

**Regra**: nao escolha AKS ou arquitetura complexa se App Service, Static Web Apps, Functions ou Container Apps resolverem com menor complexidade e custo.

### Etapa 4 — Avaliacao por pilares

Avalie cada pilar. Classifique cada item como:
- ✅ Pronto
- ⚠️ Risco parcial / acao recomendada
- ❌ Bloqueador — nao fazer deploy sem resolver

#### Pilar 1 — Seguranca

Verificar:

- [ ] Secrets hardcoded no codigo ou em arquivos versionados.
- [ ] Connection strings no codigo.
- [ ] Arquivos `.env` no repositorio.
- [ ] Azure Key Vault configurado para secrets de producao.
- [ ] Managed Identity para acesso a Key Vault, Storage, DB, Service Bus.
- [ ] RBAC com menor privilegio: pipelines, servicos, desenvolvedor.
- [ ] Service connections do Azure DevOps com escopo restrito.
- [ ] Aprovacoes obrigatorias no pipeline antes de producao.
- [ ] HTTPS obrigatorio, TLS >= 1.2.
- [ ] CORS configurado corretamente.
- [ ] IP restrictions ou Private Endpoint, quando aplicavel.
- [ ] Storage e containers sem acesso publico indevido.
- [ ] Logs nao expondo dados sensiveis (PII, tokens, senhas).
- [ ] Dependencias sem vulnerabilidades criticas (CVE).
- [ ] Imagens de container sem CVE critico, quando aplicavel.
- [ ] Contas administrativas nao persistentes.
- [ ] Auditoria de acessos habilitada (Azure Activity Log, Defender for Cloud).

**Bloquear deploy se:**
- Secret real ou credencial de producao encontrados no codigo.
- Falta de autenticacao em API sensivel.
- Permissoes excessivas criticas no pipeline ou servico.
- Banco de producao acessivel publicamente sem protecao.
- Storage publico com dados sensiveis.
- Pipeline com service connection sem controle de aprovacao.

#### Pilar 2 — Confiabilidade

Verificar:

- [ ] Health check endpoint implementado (`/health`, `/ready`).
- [ ] Liveness e readiness probes configurados, quando aplicavel.
- [ ] Retry com backoff exponencial para chamadas externas.
- [ ] Timeout configurado para todas as chamadas externas.
- [ ] Circuit breaker, quando aplicavel.
- [ ] Graceful shutdown implementado.
- [ ] Slots de deployment configurados (App Service).
- [ ] Estrategia de rollback documentada e testada.
- [ ] Versionamento de artefatos e imagens.
- [ ] Migrations seguras (nao destrutivas, reversiveis).
- [ ] Compatibilidade backward entre versoes.
- [ ] Backup configurado para banco de dados.
- [ ] Restore testado.
- [ ] Plano de DR ou descricao de RTO/RPO.
- [ ] Autoscale configurado.
- [ ] Redundancia de zonas, quando exigido por SLA.

#### Pilar 3 — Excelencia operacional

Verificar:

- [ ] Pipeline CI/CD versionado no repositorio.
- [ ] Build reprodutivel (lockfile, versao fixa de runtime).
- [ ] Infraestrutura como codigo para o ambiente de producao.
- [ ] Separacao de ambientes (dev, staging, producao).
- [ ] Aprovacoes e gates antes de producao no Azure DevOps.
- [ ] Logs centralizados (Application Insights, Log Analytics).
- [ ] Metricas customizadas instrumentadas.
- [ ] Alertas configurados para erros, latencia e disponibilidade.
- [ ] Dashboard operacional.
- [ ] Runbook de deploy.
- [ ] Runbook de rollback.
- [ ] Runbook de incidente.
- [ ] Responsaveis por aprovacao de producao mapeados.
- [ ] Registro de mudancas (change log, release notes).
- [ ] Estrategia de hotfix documentada.

#### Pilar 4 — Performance

Verificar:

- [ ] Build configurado em modo producao.
- [ ] Minificacao e bundling para frontend, quando aplicavel.
- [ ] Cache HTTP configurado.
- [ ] Compressao (gzip/brotli) habilitada.
- [ ] CDN configurado, quando aplicavel.
- [ ] Tamanho de imagem otimizado, quando aplicavel.
- [ ] Startup time medido e aceitavel.
- [ ] Pooling de conexoes de banco configurado.
- [ ] Queries sem scan completo em tabelas grandes.
- [ ] Limites de CPU e memoria configurados.
- [ ] Autoscale configurado.
- [ ] Testes de carga realizados para sistemas criticos.

#### Pilar 5 — Custo

Verificar:

- [ ] SKU e tamanho de plano adequados ao workload.
- [ ] Autoscale com limite maximo configurado.
- [ ] Retencao de logs dimensionada (evitar acumulo desnecessario).
- [ ] Recursos orfaos identificados (ambientes de dev nao apagados, IPs nao associados).
- [ ] Storage com lifecycle policy, quando aplicavel.
- [ ] Alertas de orcamento no Azure Cost Management.
- [ ] Tags de custo aplicadas (projeto, ambiente, centro de custo, time).
- [ ] Estimativa de custo mensal calculada no Azure Pricing Calculator.

#### Pilar 6 — Governanca

Verificar:

- [ ] Naming convention consistente com padrao da organizacao.
- [ ] Tags obrigatorias em todos os recursos.
- [ ] Resource groups por ambiente (dev, staging, producao).
- [ ] Separacao de subscriptions, quando exigido.
- [ ] RBAC aplicado com menor privilegio em todos os recursos.
- [ ] Azure Policy aplicada para regioes, naming, tags, TLS.
- [ ] Controle de acesso a secrets documentado.
- [ ] Controle de acesso ao pipeline documentado.
- [ ] Registro de decisoes arquiteturais.

### Etapa 5 — Checklist de prontidao para producao

Gere um checklist consolidado com status de cada item dos pilares.
Separe em tres secoes:
1. **Bloqueadores** — nao fazer deploy enquanto pendentes.
2. **Riscos pendentes** — acao recomendada antes de producao, mas nao bloqueador absoluto.
3. **Melhorias futuras** — itens para sprints seguintes.

### Etapa 6 — Plano de acao

Para cada item de bloqueador ou risco pendente, gere:

```
Item: <nome do item>
Status: Bloqueador / Risco
Evidencia: <o que foi encontrado>
Acao recomendada: <o que fazer>
Responsavel sugerido: Desenvolvedor / DevOps / Arquiteto / Lider tecnico
Prazo sugerido: Antes do deploy / Proximo sprint / Roadmap
```

### Etapa 7 — Configuracao da infraestrutura e pipeline

Propor ou revisar, conforme o caso:

1. **Bicep ou Terraform** para provisionar os recursos Azure necessarios.
2. **azure-pipelines.yml** com:
   - Stages: `build`, `test`, `deploy-staging`, `approval`, `deploy-production`.
   - Environments com aprovacao obrigatoria para producao.
   - Gates de qualidade (lint, test, cobertura minima, scan de seguranca).
   - Artefatos versionados.
   - Rollback automatizado por slot swap ou re-deploy de versao anterior.
3. **App Settings** separados por ambiente (sem secrets hardcoded).
4. **Key Vault references** no App Service ou Container Apps, quando aplicavel.

### Etapa 8 — Observabilidade pos-deploy

Propor ou revisar:

1. Application Insights configurado e instrumentado no codigo.
2. Metricas customizadas para KPIs do negocio.
3. Alertas para:
   - Taxa de erros HTTP 5xx acima de threshold.
   - Latencia de resposta acima de threshold.
   - Disponibilidade abaixo de threshold.
   - CPU / memoria acima de threshold.
4. Dashboard no Azure Monitor com metricas principais.
5. Log Analytics workspace com queries salvas para troubleshooting.
6. Runbook de resposta a incidente.

### Etapa 9 — Relatorio final

Producao apenas com relatorio auditavel contendo:

1. Resumo executivo.
2. Tipo de aplicacao e servico Azure escolhido.
3. Baseline tecnico.
4. Status de cada pilar com evidencias.
5. Checklist consolidado.
6. Plano de acao com responsaveis e prazos.
7. Pipeline proposto ou revisado.
8. Infraestrutura proposta ou revisada.
9. Estrategia de rollback.
10. Runbooks.
11. Riscos remanescentes e mitigacao.
12. Proximos passos recomendados.

---

## Regras de seguranca obrigatorias

1. **Nunca commitar secrets, tokens ou credenciais reais no repositorio.**
2. **Nunca fazer deploy direto em producao sem aprovacao humana registrada.**
3. **Nunca usar service connection sem escopo restrito no Azure DevOps.**
4. **Nunca expor banco de dados de producao publicamente sem WAF ou Private Endpoint.**
5. **Nunca remover mecanismo de rollback sem substituto equivalente.**
6. **Se houver duvida sobre impacto em producao: bloquear, registrar e solicitar decisao.**
7. **Nao misturar infra de producao com dev/staging sem segregacao explicita.**

---

## Referencias e recursos

Consulte quando necessario:

- [Azure Well-Architected Framework](https://learn.microsoft.com/pt-br/azure/well-architected/)
- [Azure DevOps Pipelines — YAML schema](https://learn.microsoft.com/pt-br/azure/devops/pipelines/yaml-schema/)
- [Azure DevOps — Environments, approvals e gates](https://learn.microsoft.com/pt-br/azure/devops/pipelines/process/environments)
- [Azure App Service — Deploy slots](https://learn.microsoft.com/pt-br/azure/app-service/deploy-staging-slots)
- [Azure Key Vault — Managed Identity](https://learn.microsoft.com/pt-br/azure/key-vault/general/managed-identity)
- [Azure RBAC — Principio do menor privilegio](https://learn.microsoft.com/pt-br/azure/role-based-access-control/best-practices)
- [Application Insights — Visao geral](https://learn.microsoft.com/pt-br/azure/azure-monitor/app/app-insights-overview)
- [Azure Monitor — Alertas](https://learn.microsoft.com/pt-br/azure/azure-monitor/alerts/alerts-overview)
- [Microsoft Cloud Adoption Framework](https://learn.microsoft.com/pt-br/azure/cloud-adoption-framework/)
- [Azure Pricing Calculator](https://azure.microsoft.com/pt-br/pricing/calculator/)
- [Container Apps vs App Service — Escolha](https://learn.microsoft.com/pt-br/azure/container-apps/compare-options)
- [Bicep — Documentacao](https://learn.microsoft.com/pt-br/azure/azure-resource-manager/bicep/)
- [Defender for Cloud — Recomendacoes](https://learn.microsoft.com/pt-br/azure/defender-for-cloud/recommendations-reference)
