# Etapa 3: automações (Power Automate)

Tempo estimado: **30 minutos**. Usa só conectores padrão (SharePoint e Outlook), incluídos no Microsoft 365.

> O app **já atualiza a planilha central sozinho** ao salvar uma calibração. Os fluxos abaixo cuidam dos **avisos**:
> 1. alertar a manutenção quando houver recusa ou pedido de manutenção;
> 2. atualizar o status todo dia e mandar um resumo do que está vencendo.

Acesse **https://make.powerautomate.com** > **+ Criar**.

---

## Fluxo 1: "Calibração - Alerta de recusa e manutenção"

**Tipo:** Fluxo da nuvem automatizado.

| # | Ação | Configuração |
|---|---|---|
| 1 | **Gatilho:** SharePoint > *Quando um item é criado* | Endereço do site: seu site. Nome da lista: `Calibracoes` |
| 2 | **Condição** | Clique em *Editar no modo avançado* e cole: `@or(equals(triggerOutputs()?['body/Resultado'], 'Recusada'), equals(triggerOutputs()?['body/ManutencaoSolicitada'], true))` |
| 3 | **Se sim:** Office 365 Outlook > *Enviar um email (V2)* | Veja abaixo |
| 4 | **Se não:** (vazio) | |

**E-mail (passo 3):**
- **Para:** e-mail da manutenção e do responsável do laboratório (pode ser uma caixa compartilhada).
- **Assunto** (expressão):
  ```
  concat('[Calibração] ', triggerOutputs()?['body/Title'], ' - ', if(equals(triggerOutputs()?['body/ManutencaoSolicitada'], true), 'MANUTENÇÃO SOLICITADA', 'CALIBRAÇÃO RECUSADA'))
  ```
- **Corpo:** use o conteúdo dinâmico do gatilho:
  ```
  Equipamento: [Título]
  Data da calibração: [DataCalibracao]
  Responsável: [Responsavel]
  Local: [Local]
  Resultado: [Resultado]
  Motivo da recusa: [MotivoRecusa]
  Manutenção solicitada: [ManutencaoSolicitada]
  Observações: [Observacoes]

  Abrir no app: <LINK DA WEB DO APP>&equip=[Título]
  ```
- **Importância:** Alta.

**Extras opcionais** (no mesmo "Se sim"):
- **Microsoft Teams > Postar mensagem em um chat ou canal**, no canal da manutenção.
- **Planner > Criar uma tarefa**, para acompanhar a manutenção até o fim.

---

## Fluxo 2: "Calibração - Status diário e aviso de vencimentos"

**Tipo:** Fluxo da nuvem agendado. **Repetir a cada:** 1 dia, às **06:00**, fuso **(UTC-03:00) Brasília**.

| # | Ação | Nome sugerido | Configuração |
|---|---|---|---|
| 1 | **Gatilho:** Recorrência | | 1 dia, 06:00, fuso de Brasília |
| 2 | *Inicializar variável* | `varHoje` | Tipo **Cadeia de caracteres**. Valor (expressão): `convertFromUtc(utcNow(), 'E. South America Standard Time', 'yyyy-MM-dd')` |
| 3 | *Inicializar variável* | `varLimite` | Tipo **Cadeia de caracteres**. Valor: `convertFromUtc(addDays(utcNow(), 30), 'E. South America Standard Time', 'yyyy-MM-dd')` |
| 4 | *Inicializar variável* | `varAlertas` | Tipo **Matriz**. Valor: `[]` |
| 5 | SharePoint > *Obter itens* | | Lista `Equipamentos`. Consulta de filtro: `Ativo eq 1`. Contagem superior: `5000` |
| 6 | *Aplicar a cada* | | Saída: `value` do "Obter itens" |
| 6.1 | dentro do loop: *Compor* | **renomeie para `StatusNovo`** | Expressão abaixo |
| 6.2 | dentro do loop: *Condição* | | `outputs('StatusNovo')` **não é igual a** `Status` (conteúdo dinâmico) |
| 6.2.a | Se sim: SharePoint > *Atualizar item* | | Lista `Equipamentos`. Id: `ID`. Título: `Título`. **Status:** `outputs('StatusNovo')` |
| 6.3 | dentro do loop: *Condição* | | Modo avançado: `@contains(createArray('Vencido','Próximo','Recusada','Manutenção solicitada'), outputs('StatusNovo'))` |
| 6.3.a | Se sim: *Acrescentar à variável de matriz* | | `varAlertas`, valor abaixo |
| 7 | *Condição* (fora do loop) | | `length(variables('varAlertas'))` **é maior que** `0` |
| 7.a | Se sim: *Criar tabela HTML* | | De: `variables('varAlertas')`. Colunas: Automático |
| 7.b | Se sim: Outlook > *Enviar um email (V2)* | | Para: laboratório + instrumentação. Assunto: `Calibrações - vencidas e a vencer (30 dias)`. Corpo: saída da tabela HTML |

**Expressão do passo 6.1 (`StatusNovo`):**
```
if(equals(item()?['ManutencaoSolicitada'], true), 'Manutenção solicitada',
 if(equals(item()?['UltimoResultado'], 'Recusada'), 'Recusada',
  if(empty(item()?['ProximaCalibracao']), 'Sem registro',
   if(less(formatDateTime(coalesce(item()?['ProximaCalibracao'], '1900-01-01'), 'yyyy-MM-dd'), variables('varHoje')), 'Vencido',
    if(lessOrEquals(formatDateTime(coalesce(item()?['ProximaCalibracao'], '1900-01-01'), 'yyyy-MM-dd'), variables('varLimite')), 'Próximo',
     'Válido')))))
```
(Cole tudo numa linha só se o editor reclamar das quebras.)

**Valor do passo 6.3.a:**
```json
{
  "TAG": "@{item()?['Title']}",
  "Equipamento": "@{item()?['Descricao']}",
  "Local": "@{item()?['Local']}",
  "Próxima calibração": "@{if(empty(item()?['ProximaCalibracao']), '-', formatDateTime(item()?['ProximaCalibracao'], 'dd/MM/yyyy'))}",
  "Status": "@{outputs('StatusNovo')}"
}
```

**Dicas:**
- Quer o resumo só às segundas? No gatilho, escolha **Frequência: Semana**, **Nestes dias: Segunda-feira**. Ou crie um terceiro fluxo semanal só para o e-mail.
- Para deixar o e-mail mais bonito, coloque antes da tabela um `<style>table{border-collapse:collapse}td,th{border:1px solid #ccc;padding:4px 8px}th{background:#123b63;color:#fff}</style>`.

---

## Por que o status também é calculado no fluxo?

O app e a formatação da coluna calculam o status **na hora**, a partir das datas. O fluxo diário grava esse status na coluna `Status`. Assim também funcionam **filtros, exportação para Excel e um futuro Power BI**, que leem o valor gravado.
