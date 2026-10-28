# Etapa 3 · Automações (Power Automate)

Quatro fluxos, todos com **conectores padrão** (SharePoint, Office 365 Outlook): sem licença extra.

| Fluxo | Quando roda | O que faz |
|---|---|---|
| **F0 · Sincronizar acessos** | Quando alguém muda a lista **Acessos** | Coloca ou tira a pessoa do grupo certo do SharePoint. A lista Acessos vira o **único lugar** onde se define acesso |
| **F1 · Pós-lançamento** | A cada envio de aferição pelo app | Atualiza a situação dos equipamentos e manda **alerta imediato** de não conformidade para a instrumentação |
| **F2 · Rotina diária** | Todo dia às 07:00 | Recalcula a situação de todos os equipamentos e envia o **relatório diário**: agenda de ontem (feito x pendente), não conformidades e calibrações. **Parte E:** avisa quais calibrações venceram ontem, para você trocar a imagem por NÃO UTILIZAR |
| **F3 · Calibração registrada** | A cada calibração registrada pelo app | Avisa você ("foi calibrada a balança tal") com a **linha pronta** para o gerador de imagens de consulta |

> **As imagens de consulta (o que a câmera abre) só mudam em dois momentos:** quando uma calibração é **aprovada** (F3) e quando uma calibração **vence** (F2, parte E). Aferição não conforme **não** muda a imagem: ela bloqueia o equipamento só no app, para a equipe.

> **Crie os quatro com a sua conta** (Proprietária do site). Eles gravam em listas que os laboratoristas não podem editar.
>
> **Renomeie as ações com os nomes indicados** (três pontinhos da ação > Renomear). As expressões usam esses nomes. Nomes sem acento e sem espaço evitam erro.
>
> Onde aparece `SITE`, escolha o site das listas.

---

## F0 · Sincronizar acessos

**+ Criar > Fluxo da nuvem automatizado** · Gatilho **SharePoint · Quando um item é criado ou modificado** · Lista `Acessos`.

**Antes:** descubra o **ID** de cada grupo. Configurações do site > Pessoas e grupos > clique no grupo > o endereço termina com `MembershipGroupId=NN`.

| # | Ação | Nome | Configuração |
|---|---|---|---|
| 1 | Compor | `Login` | `concat('i:0#.f|membership|', toLower(triggerOutputs()?['body/Pessoa/Email']))` |
| 2 | Compor | `GrupoLab` | ID do grupo Metrologia - Laboratoristas (ex.: `12`) |
| 3 | Compor | `GrupoIns` | ID do grupo Metrologia - Instrumentação (ex.: `13`) |
| 4 | Escopo | `Remover` | Dentro dele, as ações 4.1 e 4.2 |
| 4.1 | SharePoint · Enviar uma solicitação HTTP para o SharePoint | `RemoverLab` | Método **POST** · URI `_api/web/sitegroups/getbyid(@{outputs('GrupoLab')})/users/removeByLoginName(@v)?@v='@{encodeUriComponent(outputs('Login'))}'` · Cabeçalho `Accept: application/json;odata=nometadata` |
| 4.2 | (igual) | `RemoverIns` | Igual à 4.1, com `GrupoIns`. **Configurar execução após**: marque *é bem-sucedido* e *falhou* (a pessoa pode não estar no grupo) |
| 5 | Condição | `Ativo` | `triggerOutputs()?['body/Ativo']` **é igual a** `true` · **Configurar execução após** de 5: *é bem-sucedido* e *falhou* do escopo `Remover` |
| 5.1 | Se sim: Enviar uma solicitação HTTP | `Adicionar` | Método **POST** · URI `_api/web/sitegroups/getbyid(@{if(equals(triggerOutputs()?['body/Perfil/Value'],'Laboratorista'), outputs('GrupoLab'), outputs('GrupoIns'))})/users` · Cabeçalhos `Accept` e `Content-Type`: `application/json;odata=nometadata` · Corpo: `{ "LoginName": "@{outputs('Login')}" }` |

**Regra para a equipe:** para tirar o acesso, **desmarque Ativo**. Não apague a linha, ou o fluxo não fica sabendo.

---

## F1 · Pós-lançamento

**Gatilho:** SharePoint · **Quando um item é criado** · Lista `Lancamentos`.
**Condição do gatilho** (Configurações do gatilho > Condições do gatilho > + Adicionar):
```
@equals(triggerOutputs()?['body/Origem'], 'App')
```
Assim os 80 lançamentos importados da planilha não disparam e-mails.

| # | Ação | Nome | Configuração |
|---|---|---|---|
| 1 | SharePoint · Obter itens | `ObterAfericoes` | Lista `Afericoes` · Consulta de filtro: `Title eq '@{triggerOutputs()?['body/Title']}'` · Contagem superior: `500` |
| 2 | Selecionar | `SelecionarTAGs` | De: `body('ObterAfericoes')?['value']` · Mapear (modo texto, ícone ⇄): `item()?['TAG']` |
| 3 | Compor | `TAGsUnicas` | `union(body('SelecionarTAGs'), body('SelecionarTAGs'))` |
| 4 | Aplicar a cada | `ParaCadaTAG` | Saída: `outputs('TAGsUnicas')` |
| 4.1 | Condição | `TemTAG` | `items('ParaCadaTAG')` **não é igual a** (vazio) |
| 4.2 | Se sim: Filtrar matriz | `NaoConformesDaTAG` | De: `body('ObterAfericoes')?['value']` · Modo avançado: `@and(equals(item()?['TAG'], items('ParaCadaTAG')), equals(item()?['Resultado'], 'Não conforme'))` |
| 4.3 | SharePoint · Obter itens | `ObterEquipamento` | Lista `Equipamentos` · Filtro: `Title eq '@{items('ParaCadaTAG')}'` · Contagem: `1` |
| 4.4 | Condição | `Existe` | `length(body('ObterEquipamento')?['value'])` **é maior que** `0` |
| 4.5 | Se sim: SharePoint · Atualizar item | `AtualizarEquipamento` | Lista `Equipamentos` · Id: `first(body('ObterEquipamento')?['value'])?['ID']` · Título: `items('ParaCadaTAG')` · **UltimaAfericao**: `triggerOutputs()?['body/DataEnsaio']` · **UltimaAfericaoResultado**: `if(greater(length(body('NaoConformesDaTAG')), 0), 'Não conforme', 'Conforme')` · **UltimaAfericaoLancamento**: `triggerOutputs()?['body/Title']` · **Status**: `if(greater(length(body('NaoConformesDaTAG')), 0), 'Bloqueado · Aferição não conforme', first(body('ObterEquipamento')?['value'])?['Status'])` |
| 5 | Condição (fora do loop) | `TemNaoConforme` | `triggerOutputs()?['body/QtdNaoConformes']` **é maior que** `0` |
| 5.1 | Se sim: Obter itens | `ObterInstrumentacao` | Lista `Acessos` · Filtro: `Ativo eq 1 and Perfil eq 'Instrumentação'` |
| 5.2 | Selecionar | `EmailsInstrumentacao` | De: `body('ObterInstrumentacao')?['value']` · Mapear (modo texto): `item()?['Pessoa']?['Email']` |
| 5.3 | Obter itens | `ObterEmailsExtras` | Lista `Configuracoes` · Filtro: `Title eq 'EmailsAlertaNaoConforme'` |
| 5.4 | Obter itens | `ObterLinkApp` | Lista `Configuracoes` · Filtro: `Title eq 'LinkApp'` |
| 5.5 | Filtrar matriz | `NaoConformes` | De: `body('ObterAfericoes')?['value']` · `@equals(item()?['Resultado'], 'Não conforme')` |
| 5.6 | Selecionar | `TabelaNC` | De: `body('NaoConformes')` · Mapear: TAG = `item()?['TAG']` · Parâmetro = `item()?['Parametro']` · Valor = `item()?['Valor']` · Referência = `item()?['Referencia']` · Diferença = `item()?['Diferenca']` · Itens = `item()?['ItensNaoOK']` |
| 5.7 | Criar tabela HTML | `HTMLNC` | De: `body('TabelaNC')` · Colunas: Automático |
| 5.8 | Office 365 Outlook · Enviar um email (V2) | `EmailAlerta` | Veja abaixo |

**E-mail do alerta (5.8):**
- **Para:** `@{join(body('EmailsInstrumentacao'), ';')};@{first(body('ObterEmailsExtras')?['value'])?['Valor']}`
- **Assunto:** `[Metrologia] Não conforme · @{triggerOutputs()?['body/Ensaio']} · @{triggerOutputs()?['body/TAGsNaoConformes']}`
- **Importância:** Alta
- **Corpo** (clique em `</>` para o modo código e cole o arquivo `email-alerta-nc.html` desta pasta)

---

## F2 · Rotina diária (situação + relatório)

**Gatilho:** Agendado · **Recorrência**: a cada 1 dia, às **07:00**, fuso **(UTC-03:00) Brasília**.

### Parte 0 · Datas e configurações

| # | Ação | Nome | Configuração |
|---|---|---|---|
| 1 | Inicializar variável | `varHoje` | Cadeia · `convertFromUtc(utcNow(), 'E. South America Standard Time', 'yyyy-MM-dd')` |
| 2 | Inicializar variável | `varOntem` | Cadeia · `convertFromUtc(addDays(utcNow(), -1), 'E. South America Standard Time', 'yyyy-MM-dd')` |
| 3 | Inicializar variável | `varPrefixo` | Cadeia · `createArray('Dom','Seg','Ter','Qua','Qui','Sex','Sab')[dayOfWeek(variables('varOntem'))]` |
| 4 | Obter itens | `ObterConfig` | Lista `Configuracoes` (sem filtro) |
| 5 | Filtrar matriz | `CfgDias` | De: `body('ObterConfig')?['value']` · `@equals(item()?['Title'], 'DiasAvisoCalibracao')` |
| 6 | Inicializar variável | `varLimite` | Cadeia · `formatDateTime(addDays(variables('varHoje'), int(coalesce(first(body('CfgDias'))?['Valor'], '30'))), 'yyyy-MM-dd')` |
| 7 | Filtrar matriz | `CfgEmails` | De: `body('ObterConfig')?['value']` · `@equals(item()?['Title'], 'EmailsRelatorioDiario')` |
| 8 | Inicializar variável | `varAlertasEquip` | Matriz · `[]` |
| 9 | Filtrar matriz | `CfgAviso` | De: `body('ObterConfig')?['value']` · `@equals(item()?['Title'], 'EmailsAvisoImagem')` |
| 10 | Filtrar matriz | `CfgPasta` | De: `body('ObterConfig')?['value']` · `@equals(item()?['Title'], 'PastaImagensQR')` |

### Parte A · Situação de todos os equipamentos

| # | Ação | Nome | Configuração |
|---|---|---|---|
| A1 | Obter itens | `ObterEquipamentos` | Lista `Equipamentos` · Filtro `Ativo eq 1` · Contagem `5000` |
| A2 | Aplicar a cada | `ParaCadaEquip` | `body('ObterEquipamentos')?['value']` |
| A2.1 | Compor | `StatusNovo` | expressão abaixo |
| A2.2 | Condição | `Mudou` | `outputs('StatusNovo')` **não é igual a** `items('ParaCadaEquip')?['Status']` |
| A2.3 | Se sim: Atualizar item | `AtualizarStatus` | Id: `items('ParaCadaEquip')?['ID']` · Título: `items('ParaCadaEquip')?['Title']` · **Status**: `outputs('StatusNovo')` |
| A2.4 | Condição | `Alertar` | Modo avançado: `@or(startsWith(outputs('StatusNovo'), 'Bloqueado'), equals(outputs('StatusNovo'), 'Liberado · Calibração vence em breve'))` |
| A2.5 | Se sim: Acrescentar à variável de matriz | | `varAlertasEquip` · valor: `{"TAG": "@{items('ParaCadaEquip')?['Title']}", "Equipamento": "@{items('ParaCadaEquip')?['Descricao']}", "Situação": "@{outputs('StatusNovo')}", "Próxima calibração": "@{if(empty(items('ParaCadaEquip')?['ProximaCalibracao']), '-', formatDateTime(items('ParaCadaEquip')?['ProximaCalibracao'], 'dd/MM/yyyy'))}"}` |

**Expressão de `StatusNovo`** (mesma regra do app e do Power BI; cole numa linha só se o editor reclamar):
```
if(and(equals(items('ParaCadaEquip')?['ControlaCalibracao'], true), equals(items('ParaCadaEquip')?['UltimoResultado'], 'Recusada')), 'Bloqueado · Calibração recusada',
if(and(equals(items('ParaCadaEquip')?['ControlaCalibracao'], true), equals(items('ParaCadaEquip')?['ManutencaoSolicitada'], true)), 'Bloqueado · Em manutenção',
if(and(equals(items('ParaCadaEquip')?['ControlaCalibracao'], true), not(empty(items('ParaCadaEquip')?['ProximaCalibracao'])), less(formatDateTime(coalesce(items('ParaCadaEquip')?['ProximaCalibracao'], '2999-12-31'), 'yyyy-MM-dd'), variables('varHoje'))), 'Bloqueado · Calibração vencida',
if(and(equals(items('ParaCadaEquip')?['UltimaAfericaoResultado'], 'Não conforme'), not(and(equals(items('ParaCadaEquip')?['UltimoResultado'], 'Aprovada'), greaterOrEquals(formatDateTime(coalesce(items('ParaCadaEquip')?['UltimaCalibracao'], '1900-01-01'), 'yyyy-MM-dd'), formatDateTime(coalesce(items('ParaCadaEquip')?['UltimaAfericao'], '2999-12-31'), 'yyyy-MM-dd'))))), 'Bloqueado · Aferição não conforme',
if(and(equals(items('ParaCadaEquip')?['ControlaCalibracao'], true), empty(items('ParaCadaEquip')?['ProximaCalibracao'])), 'Liberado · Sem data de calibração',
if(and(equals(items('ParaCadaEquip')?['ControlaCalibracao'], true), lessOrEquals(formatDateTime(coalesce(items('ParaCadaEquip')?['ProximaCalibracao'], '2999-12-31'), 'yyyy-MM-dd'), variables('varLimite'))), 'Liberado · Calibração vence em breve',
'Liberado'))))))
```

### Parte B · Agenda de ontem (feito x pendente)

| # | Ação | Nome | Configuração |
|---|---|---|---|
| B1 | Obter itens | `ObterEnsaios` | Lista `Ensaios` · Filtro `Ativo eq 1` |
| B2 | Obter itens | `ObterLancOntem` | Lista `Lancamentos` · Filtro: `DataEnsaio ge '@{variables('varOntem')}T00:00:00Z' and DataEnsaio lt '@{variables('varHoje')}T00:00:00Z' and Situacao eq 'Ativo'` |
| B3 | Selecionar | `ChavesEnviadas` | De `body('ObterLancOntem')?['value']` · modo texto: `concat(item()?['Ensaio'], '|', item()?['Turno'])` |
| B4 | Compor | `Enviados` | `join(body('ChavesEnviadas'), ';')` |
| B5 | Selecionar | `Agenda` | De `body('ObterEnsaios')?['value']` · Mapear: **Ensaio** = `item()?['Nome']` · **07x19** = expressão D · **19x07** = expressão N (abaixo) |
| B6 | Filtrar matriz | `AgendaRelevante` | De `body('Agenda')` · `@not(and(equals(item()?['07x19'], '—'), equals(item()?['19x07'], '—')))` |
| B7 | Filtrar matriz | `Pendentes` | De `body('AgendaRelevante')` · `@or(startsWith(item()?['07x19'], '✖'), startsWith(item()?['19x07'], '✖'))` |
| B8 | Criar tabela HTML | `HTMLAgenda` | De `body('AgendaRelevante')` |

**Expressão D** (troque `D` e `07x19` por `N` e `19x07` na outra coluna):
```
if(equals(item()?[concat(variables('varPrefixo'), 'D')], true),
   if(contains(outputs('Enviados'), concat(item()?['NomeBD'], '|07x19')), '✔ Enviado', '✖ Pendente'),
   if(contains(outputs('Enviados'), concat(item()?['NomeBD'], '|07x19')), '✔ Enviado (fora da agenda)', '—'))
```

### Parte C · Não conformidades de ontem

| # | Ação | Nome | Configuração |
|---|---|---|---|
| C1 | Obter itens | `ObterNCOntem` | Lista `Afericoes` · Filtro: `DataEnsaio ge '@{variables('varOntem')}T00:00:00Z' and DataEnsaio lt '@{variables('varHoje')}T00:00:00Z' and Resultado eq 'Não conforme'` |
| C2 | Selecionar | `TabelaNCOntem` | TAG = `item()?['TAG']` · Ensaio = `item()?['Ensaio']` · Parâmetro = `item()?['Parametro']` · Valor = `item()?['Valor']` · Responsável = `item()?['Responsavel']` · Observação = `item()?['Observacao']` |
| C3 | Criar tabela HTML | `HTMLNCOntem` | De `body('TabelaNCOntem')` |
| C4 | Criar tabela HTML | `HTMLEquip` | De `variables('varAlertasEquip')` |

### Parte D · E-mail

| # | Ação | Configuração |
|---|---|---|
| D1 | Enviar um email (V2) | **Para:** `first(body('CfgEmails'))?['Valor']` · **Assunto:** `[Metrologia] Relatório diário · @{formatDateTime(variables('varOntem'), 'dd/MM/yyyy')} · @{length(body('Pendentes'))} pendente(s) · @{length(body('ObterNCOntem')?['value'])} não conforme(s)` · **Corpo:** modo código, cole `email-relatorio-diario.html` |

### Parte E · Calibrações que venceram ontem (imagem de consulta)

Uma calibração "válida até 27/10" passa a vencida no dia 28. A rotina das 07:00 do dia 28 encontra esse equipamento e manda as linhas prontas para o gerador de imagens.

| # | Ação | Nome | Configuração |
|---|---|---|---|
| E1 | Filtrar matriz | `VenceramOntem` | De: `body('ObterEquipamentos')?['value']` · Modo avançado: `@and(equals(item()?['ControlaCalibracao'], true), not(empty(item()?['ProximaCalibracao'])), equals(formatDateTime(coalesce(item()?['ProximaCalibracao'], '1900-01-01'), 'yyyy-MM-dd'), variables('varOntem')))` |
| E2 | Selecionar | `LinhasImagem` | De: `body('VenceramOntem')` · Mapear (modo texto, ícone ⇄): expressão abaixo |
| E3 | Condição | `TemVencida` | `length(body('VenceramOntem'))` **é maior que** `0` |
| E4 | Se sim: Enviar um email (V2) | `EmailAvisoImagem` | **Para:** `first(body('CfgAviso'))?['Valor']` · **Assunto:** `[Metrologia] Atualize a imagem: @{length(body('VenceramOntem'))} calibração(ões) vencida(s)` · **Importância:** Alta · **Corpo:** modo código, cole `email-aviso-imagem.html` |

**Expressão de E2** (uma linha no formato do gerador: TAG; Descrição; Local; Situação; Última; Próxima; Certificado; Periodicidade):
```
concat(item()?['Title'], '; ', coalesce(item()?['Descricao'], ''), '; ', coalesce(item()?['Local'], ''), '; NÃO UTILIZAR; ', if(empty(item()?['UltimaCalibracao']), '', formatDateTime(item()?['UltimaCalibracao'], 'dd/MM/yyyy')), '; ', formatDateTime(item()?['ProximaCalibracao'], 'dd/MM/yyyy'), '; ', coalesce(item()?['UltimoCertificado'], ''), '; ', coalesce(string(item()?['PeriodicidadeMeses']), ''))
```

> Prefere um e-mail só? Em vez de E4, cole o bloco de `email-aviso-imagem.html` dentro do relatório diário (D1). O separado é melhor porque chega com importância alta e você não perde no meio do relatório.

---

## F3 · Calibração registrada (aviso para atualizar a imagem)

**Gatilho:** SharePoint · **Quando um item é criado** · Lista `Calibracoes`.

| # | Ação | Nome | Configuração |
|---|---|---|---|
| 1 | Obter itens | `ObterEquipamento` | Lista `Equipamentos` · Filtro: `Title eq '@{triggerOutputs()?['body/Title']}'` · Contagem: `1` |
| 2 | Compor | `Equip` | `first(body('ObterEquipamento')?['value'])` |
| 3 | Obter itens | `ObterConfig` | Lista `Configuracoes` (sem filtro) |
| 4 | Filtrar matriz | `CfgAviso` | De: `body('ObterConfig')?['value']` · `@equals(item()?['Title'], 'EmailsAvisoImagem')` |
| 5 | Filtrar matriz | `CfgPasta` | Igual, com `'PastaImagensQR'` |
| 6 | Filtrar matriz | `CfgLink` | Igual, com `'LinkApp'` |
| 7 | Compor | `Aprovada` | `equals(triggerOutputs()?['body/Resultado'], 'Aprovada')` |
| 8 | Compor | `Linha` | expressão abaixo |
| 9 | Enviar um email (V2) | `EmailCalibracao` | **Para:** `first(body('CfgAviso'))?['Valor']` · **Assunto:** `[Metrologia] Calibração registrada · @{triggerOutputs()?['body/Title']} · @{triggerOutputs()?['body/Resultado']}` · **Corpo:** modo código, cole `email-calibracao-registrada.html` |

**Expressão de `Linha`** (a situação sai da data: NÃO UTILIZAR só se a nova validade já tiver passado, o que acontece quando se registra uma calibração antiga):
```
concat(triggerOutputs()?['body/Title'], '; ', coalesce(outputs('Equip')?['Descricao'], ''), '; ', coalesce(triggerOutputs()?['body/Local'], outputs('Equip')?['Local'], ''), '; ', if(less(formatDateTime(triggerOutputs()?['body/ProximaCalibracao'], 'yyyy-MM-dd'), convertFromUtc(utcNow(), 'E. South America Standard Time', 'yyyy-MM-dd')), 'NÃO UTILIZAR', 'CALIBRADO'), '; ', formatDateTime(triggerOutputs()?['body/DataCalibracao'], 'dd/MM/yyyy'), '; ', formatDateTime(triggerOutputs()?['body/ProximaCalibracao'], 'dd/MM/yyyy'), '; ', coalesce(triggerOutputs()?['body/Certificado'], ''), '; ', coalesce(string(outputs('Equip')?['PeriodicidadeMeses']), ''))
```

**Calibração recusada:** o e-mail chega em vermelho e **sem** a linha do gerador. A imagem de consulta não muda (regra: só calibração aprovada ou vencimento mudam a imagem). Para a equipe, o app já mostra NÃO UTILIZAR.

---

## Relatório visual do Power BI (complemento, sem programar)

No Power BI online, abra o relatório > **Assinar** > **Criar assinatura**: escolha a página do painel, frequência **Diária**, 07:30, destinatários. Todo dia chega a imagem do painel por e-mail com o link. Precisa de licença Pro (ou capacidade Premium).

## Testar

- **F0:** mude o Perfil de uma pessoa de teste e confira o grupo em Pessoas e grupos.
- **F1:** lance pelo app um Blaine de 1985 (não conforme). Em 1 a 2 minutos: e-mail para a instrumentação e equipamento bloqueado na lista.
- **F2:** abra o fluxo > **Testar** > **Manualmente**. Confira o e-mail e a coluna Status da lista Equipamentos.
- **F2, parte E:** num equipamento de teste, coloque `ProximaCalibracao` = ontem e rode o F2 manualmente. Chega o aviso "Atualize a imagem" com a linha NÃO UTILIZAR.
- **F3:** registre pelo app uma calibração aprovada. Em 1 a 2 minutos chega o e-mail com a linha; cole no gerador de imagens e confira a imagem.
