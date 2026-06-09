# Justificativa de Acessos para PnP Online com Certificado

## Objetivo do Documento

Este documento apresenta os casos de negócio em que o programa de solicitações do setor precisa se conectar ao SharePoint via **PnP Online com certificado** (identidade de aplicação), além de justificar a necessidade das permissões:

- `Sites.FullControl.All`
- `User.ReadWrite.All`

O foco é suportar decisão de aprovação por TI/Security com base em continuidade operacional, governança e escalabilidade.

## Contexto Atual e Evolução Esperada

O projeto já executa operações críticas no SharePoint para validar dados, atualizar templates e registrar solicitações. Atualmente, parte do fluxo depende de login interativo.

A evolução prevista é transformar o projeto em um **programa setorial de solicitações gerais**, com execução não assistida e capacidade de administrar acessos e perfis de usuários em escala.

Sem autenticação por certificado e sem as permissões apropriadas, o programa fica limitado a execução manual, com risco de indisponibilidade operacional e baixa governança de acesso.

## Por Que PnP Online com Certificado

A autenticação por certificado é necessária para:

- Permitir execução não assistida (agendamentos, jobs noturnos, rotinas automáticas).
- Evitar dependência de credenciais pessoais e sessões interativas.
- Garantir trilha de auditoria por identidade de aplicação.
- Facilitar rotação e revogação controlada de credencial.
- Sustentar escala para múltiplas frentes do setor sem risco de bloqueio por MFA interativo.

## Casos de Negócio Onde o Modelo é Necessário

### 1. Operação Noturna de Processamento de Solicitações

**Cenário:** processar lotes de solicitações fora do horário comercial para reduzir fila operacional.

**Necessidade de certificado:** não há operador disponível para autenticação interativa.

**Necessidade de permissões:** leitura e atualização de listas/campos para validar, inserir e corrigir registros.

**Resultado esperado:** menor tempo de ciclo e previsibilidade de processamento diário.

### 2. Programa Setorial de Solicitações Gerais

**Cenário:** expansão do sistema para atender múltiplos tipos de solicitação do setor em um modelo centralizado.

**Necessidade de certificado:** uma identidade de aplicação estável para operar de forma padronizada entre equipes e ambientes.

**Necessidade de permissões:** controle completo de estruturas de listas e itens para evolução contínua do serviço.

**Resultado esperado:** escala operacional com governança única e redução de retrabalho manual.

### 3. Sincronização de Templates e Metadados

**Cenário:** manter templates (como Excel) sincronizados com escolhas e regras do SharePoint para evitar erro de preenchimento.

**Necessidade de certificado:** execução automática e recorrente sem intervenção humana.

**Necessidade de permissões:** consulta e manutenção de campos, escolhas e estruturas que suportam a qualidade dos dados.

**Resultado esperado:** redução de inconsistência e menor taxa de rejeição na validação.

### 4. Gestão de Usuários e Acessos para Governança

**Cenário:** controlar quem pode solicitar, aprovar, validar e administrar fluxo no SharePoint conforme crescimento do programa.

**Necessidade de certificado:** ações de governança executadas por identidade institucional, sem depender de contas individuais.

**Necessidade de permissões:** leitura/escrita de dados de usuário para automações de atribuição, manutenção de perfil e ajustes de acesso.

**Resultado esperado:** melhor segregação de função, rastreabilidade e compliance interno.

## Mapa de Permissões Solicitadas

| Permissão | Justificativa de Negócio | Capacidades habilitadas no programa |
|---|---|---|
| `Sites.FullControl.All` | Garantir operação ponta a ponta do programa setorial com autonomia para manter listas, campos, itens e regras sem interrupções manuais. | Criação/atualização de itens, leitura de metadados, evolução de estrutura SharePoint, automações de sincronização e manutenção contínua. |
| `User.ReadWrite.All` | Permitir governança de identidades e acessos em escala, com automações de atribuição e manutenção de perfis. | Enriquecimento de dados de usuário, atualização de informações de perfil necessárias ao fluxo, apoio a regras de distribuição e controle de acesso. |

## Risco de Não Concessão

Sem as permissões solicitadas e sem modelo por certificado:

- A operação permanece dependente de execução manual.
- O programa não escala com segurança para múltiplas demandas do setor.
- A governança de acesso fica fragmentada e reativa.
- A rastreabilidade de mudanças e ações administrativas é enfraquecida.
- O risco de atraso e falhas operacionais aumenta em períodos críticos.

## Controles de Segurança e Mitigações Propostos

Para aderência a boas práticas de segurança, recomenda-se:

- Concessão faseada por ambiente e escopo de uso.
- Monitoramento de logs e auditoria das ações da aplicação.
- Revisão periódica de necessidade de permissões.
- Rotação regular do certificado e política de revogação.
- Segregação entre conta de desenvolvimento e conta de produção.
- Processo formal de change management para alterações de escopo.

## Solicitação Formal para TI/Security

Solicita-se aprovação para uso de autenticação PnP Online por certificado e concessão das permissões `Sites.FullControl.All` e `User.ReadWrite.All`, com o objetivo de:

1. Viabilizar execução não assistida e contínua do programa setorial.
2. Garantir governança centralizada de dados, usuários e acessos SharePoint.
3. Sustentar expansão do sistema para solicitações gerais com controle e auditoria.
4. Reduzir risco operacional associado a processos manuais e autenticação interativa.

## Critérios de Sucesso Após Aprovação

- Rotinas automatizadas executando sem intervenção humana.
- Fluxos de solicitação operando com menor tempo de ciclo.
- Governança de acesso implementada com rastreabilidade.
- Redução de erros de dados por sincronização automática de regras e metadados.
- Base técnica pronta para expansão do programa para novas frentes do setor.
