# 13 — v15/v16: painel gerencial, Equipamentos & Paradas e Aferições

## v19.1: acerto de dimensionamento do mapa

Não mexe em medidas nem em fórmulas.

- **Mapa:**
  - o texto não quebra mais de linha. Antes, "qua", "dom" e "hoje" quebravam em duas linhas, desalinhavam as alturas das linhas e escondiam os nomes dos ensaios;
  - a coluna dos ensaios tem largura fixa, e o cabeçalho dela passou a se chamar "Ensaio".
- **Legenda:**
  - agora é uma caixa de texto só, com os 6 itens na mesma linha. Na v19 eram 12 caixinhas, e o amarelo e o azul-claro não apareciam;
  - ganhou uma linha que explica o amarelo e o ✖.

## v19: mapa de aferições pela regra do dia certo

Aprovado a partir do mockup `entrega/Mockup-Mapa-Dia-Certo.png`. Só a
página **Aferições** mudou; as demais só trocaram o número da versão.

**Mapa**

- Cada ensaio tem um **dia certo**, e a marca fica na coluna dele:
  - ✔ azul: feito no dia certo;
  - **amarelo**: feito com atraso;
  - ✖ laranja: não feito;
  - ○: programado;
  - ½: Blaine incompleto.
- O dia em que a atrasada foi feita aparece em **✔ azul-claro**.
- O mapa usa a largura toda e mostra **todos os dias do mês**. A primeira
  linha traz o dia da semana e destaca **hoje**.
- Embaixo do mapa ficam a **legenda** e a tabela **"Regras de cada ensaio
  · o dia certo"**. Clicar num ensaio do mapa filtra essa tabela.
- Regras e exemplos: `docs/14-metas-e-origem-dos-numeros.md`, seção 3.

**Números de rotina**

- A aderência passou a ser **feitas no dia certo ÷ programadas até
  hoje**.
- Cartão: "Rotina · feitas no dia certo".
- Situação por ensaio: colunas "No dia certo", "Com atraso" e "Não
  feitos". A coluna de frequência saiu e foi para a tabela de regras.

**Modelo**

- **DimEnsaio** ganhou a coluna **Regra**, que é o texto da tabela de
  regras.
- Nova tabela **tbl_MapaLinhas**: são as linhas do mapa, o dia da semana
  mais os ensaios. Ela não tem relacionamento com as outras tabelas.
- Novas medidas: `_Mapa Nível Data`, `_Mapa Data`, `Rotina · No Dia Certo`,
  `Rotina · Com Atraso`, `Rotina · Não Feitas`, `Rotina · Programadas`,
  `Regra · Dias Certos`, `Regra · No Dia Certo`, `Regra Aparece`.
- A página Aferições ficou mais alta (1920×1900) e continua ajustada à
  largura.

## v18: ajustes finais

**Laboratórios**

- Sai o **LDP**: da dimensão de laboratórios e da calibração. O filtro no
  modelo garante que ele não volte mesmo que apareça de novo na planilha.

**Todas as páginas**

- Abrem **ajustadas à largura**. Antes umas ajustavam à altura e pareciam
  menores no serviço.
- Cabeçalho azul-marinho com o título em branco.
- Saiu o subtítulo "Gestão de Equipamentos" sob o logo e saiu a frase sob
  o título.

**Visão Geral**

- Saiu a faixa com a frase do mês, e as frentes subiram.
- Os pontos de parada em aberto mostram o nome do equipamento; quando a
  coluna automática vem vazia, mostram a TAG.

**Calibrações**

- A tabela "Última calibração, vencimento e situação" ganhou a coluna
  **TAG**.

**Aferições**

- Saiu a tabela de não conformidades por equipamento, e o **mapa** ocupa a
  metade de baixo.
- O **✖ do mapa só aparece quando o prazo venceu**: ensaio feito dentro do
  prazo em outro dia não gera ✖. As regras estão em
  `docs/14-metas-e-origem-dos-numeros.md`.

**Aferições · Detalhe**

- O gráfico "diferença ÷ tolerância" deu lugar à tabela **Resultados por
  equipamento**: última aferição, último valor, média do período, faixa
  aceita, último resultado e nº de registros.

## v16: correções a partir dos prints da v15 e o mockup na íntegra

| Na v15 | Causa | Na v16 |
|---|---|---|
| Botões do menu apareciam como retângulos azuis vazios | O botão de navegação ficava por cima da pílula, com o preenchimento padrão | A própria pílula (imagem) navega; não há botão por cima |
| Pilares sem situação, meta e barra; "Período" e "Última atualização" cortados | Caixas baixas demais para o tamanho do texto: o Power BI esconde o que não cabe | Caixas redimensionadas; cabeçalho com "Período: …" e "Atualizado em …" em texto |
| "Pontos que pedem decisão" e "Resumo do mês" vazios | Tabelas só com medidas, sem coluna de linha | Coluna Tema (pontos) e # (resumo) |
| Linha em branco em "Saúde por grupo" e linhas "Total" | Instrumentos sem grupo; totais ligados | Grupo vazio filtrado; totais desligados |
| Disponibilidade por semana negativa (−482%) | Horas paradas lançadas num dia maior que a capacidade da semana | Disponibilidade limitada a 0% |
| Faltavam itens do mockup | — | Mapa ensaio × dia, carta de controle abrindo no Blaine, placar por laboratório, barras de meta |

As 4 páginas reconstruídas (Visão Geral, Equipamentos & Paradas, Aferições,
Aferições · Detalhe) passaram a ter **1920 × 1440** e abrem **ajustadas à
largura**. Com isso a página rola na vertical e o texto fica maior. As
demais continuam em 1920 × 1080.

### Mapa de aferições (ensaio × dia)

| Símbolo | Significado |
|---|---|
| ✔ azul | conforme |
| ✖ laranja | houve não conformidade |
| ○ ocre | estava na rotina e não foi feito (conta até ontem) |
| em branco | sem rotina naquele dia |

### Carta de controle

Sem nada escolhido na lateral, a carta abre no **Blaine Manual (66PB08)**.
Para trocar, escolha na lateral o ensaio, o equipamento e o parâmetro.


Construído a partir do mockup aprovado (rodada 2) e dos ajustes pedidos.

## O que mudou

**Menu.** O menu tem 8 botões:

1. Visão Geral
2. Equip. & Paradas
3. Calibrações
4. **Aferições**, com 2 sub-botões:
   - Resumo do mês
   - Detalhe por ensaio
5. Intervenções
6. Inspeções
7. Notas / OMs
8. Sobressalentes

Os sub-botões só aparecem quando você está numa página de Aferições, como um
menu que se abre. Saíram as páginas **RELATÓRIOS** e **PARADAS**. Paradas
agora faz parte de Equipamentos.

**Cabeçalho.** Todas as páginas usam o cabeçalho branco do mockup, com a
faixa dourada e o rodapé marcado "v15".

### Visão Geral

É a página gerencial, com quatro blocos:

- **Frase do mês** (gerada pelo painel). Exemplo: "Setembro 2026: 1 de 4
  frentes na meta. Críticas: Calibração, Inspeções…"
- **4 frentes:** Disponibilidade, Calibração em dia, Aferições conformes e
  Inspeções de amostradores. Cada uma traz:
  - o número grande;
  - a situação (● na meta · ▲ atenção · ◆ crítico), com cor e símbolo;
  - a meta;
  - uma barra de progresso;
  - dois fatos de apoio;
  - um mini gráfico: por semana; vencimentos em 30/60/90 dias; ou realizado
    × esperado.
- **Saúde por grupo operacional.** Cada grupo mostra só o que se aplica a
  ele:
  - grupos de amostragem mostram só disponibilidade;
  - grupos de laboratório (tambores, prensas, analisadores, Alpine) mostram
    só calibração e aferição.
  - Não aparece mais "não se aplica". Grupo sem nenhum indicador não entra
    na tabela.
- **Pontos que pedem decisão.** A lista é montada pelo próprio painel, com o
  quê, o porquê e a ação sugerida. Cada linha só aparece quando existe:
  - calibrações vencidas (a ação sugere começar pelas que também reprovaram
    na aferição);
  - os 2 ensaios com falhas repetidas no mês;
  - inspeções abaixo do ritmo;
  - paradas em aberto;
  - disponibilidade abaixo da meta.

### Equipamentos & Paradas

- 6 indicadores.
- Disponibilidade por grupo × meta, em um gráfico só (acaba a repetição por
  grupo e por área).
- Horas paradas por equipamento. Laranja indica parada ainda em aberto.
- Horas paradas por dia do mês, separadas por situação.
- Tabela de equipamentos com a última parada e a situação.

### Aferições · Resumo do mês

Sem o mapa diário, como você pediu.

- 4 indicadores com meta:
  - conformidade;
  - aderência à rotina;
  - não conformidades;
  - equipamentos reincidentes.
- Situação por ensaio: situação, % conforme, rotina, frequência, número de
  NC e último registro.
- Resumo do mês em texto.
- Não conformidades por equipamento: o que falhou, quantas vezes, em que
  datas, a situação da calibração externa e a gravidade.

### Aferições · Detalhe por ensaio

- **Carta de controle.** Escolha na lateral o ensaio, o equipamento e o
  parâmetro. A carta mostra:
  - valor × limites, nos ensaios de calibração;
  - diferença × ± tolerância, nos comparativos.
- **Comparativos** em % da tolerância. Acima de 100% está fora.
- **Registros do período.**

## As 3 fontes no SharePoint

| Fonte | Alimenta | Situação |
|---|---|---|
| `Base_PowerBI_Gestao_Equipamentos.xlsx` | equipamentos, calibração, paradas, intervenções, notas, sobressalentes, configuração | já funcionando |
| `Inspeção de Amostradores.xlsx` (Microsoft Forms) | inspeções, junto com a aba da Base | já funcionando |
| **Pasta das planilhas mensais `CALIBRAÇÃO MM - MÊS.xlsm`** | aferições e comparativos | **já configurada** |

### Aferições: como a pasta é lida

O painel lê a pasta "Calibração Integrada mensal" e entra em cada pasta
"Calibração mensal AAAA" de 2026 em diante:

    automacao / laboratorios / Laboratório de Controle da Produção - Ubu /
    Resultados rotina / Laboratorio Fisico Ubu / Calibração Integrada  -LCP Ubu /
    Calibração Integrada mensal / Calibração mensal 2026 / CALIBRAÇÃO 09 - SETEMBRO (1).xlsm

- Cada planilha `.xlsm` dessas pastas é lida.
- O ano novo (pasta "Calibração mensal 2027") entra sozinho.
- Só entram as planilhas que têm a aba oculta `BD_Afericoes`. A dos
  operadores, antes da troca, é ignorada sem erro.
- O mesmo lançamento em duas planilhas (por exemplo, "SETEMBRO" e
  "SETEMBRO (1)") entra uma vez só.
- O endereço é fixo, como o serviço do Power BI exige para agendar a
  atualização.

**Primeira atualização.** O Power BI pede credencial para
`https://smineracao.sharepoint.com/sites/automacao`. Escolha **Conta
organizacional**, entre com a conta Samarco e defina a privacidade
**Organizacional**.

**Se a página Aferições disser "Pasta não encontrada".** Isso significa que
"laboratorios" é um subsite, e não uma biblioteca do site "automacao". Daqui
não dá para saber qual dos dois é. A correção é em uma linha, nas duas
consultas `tbl_Afericoes` e `tbl_Afericoes_Status` (Transformar dados >
Editor Avançado):

1. Troque o Site para `"https://smineracao.sharepoint.com/sites/automacao/laboratorios"`.
2. Tire `"laboratorios",` do início do Caminho.

**Para os dados aparecerem**, a planilha da pasta precisa ser a versão com
as abas ocultas (`afericoes/CALIBRAÇÃO 09 - SETEMBRO.xlsm` deste
repositório). Enquanto for a versão antiga:

- a página Aferições mostra "Pasta lida: N planilhas .xlsm, 0 com a aba
  oculta BD_Afericoes";
- o restante do painel funciona normalmente.

**Setembro.** A planilha que você mandou é uma cópia de 23/09. Os
operadores continuam lançando na original.

- **Opção A.** Substitua a original agora pela versão com abas ocultas e
  relance o que entrou depois de 23/09.
- **Opção B.** Use a versão com abas ocultas como base de outubro: "Salvar
  como" e limpar os valores digitados. O painel passa a ler a partir de
  outubro.

## Metas

As metas são lidas da aba **Configuração** da Base. Quando o parâmetro não
existe lá, o painel usa o valor-padrão da tabela. Para mudar uma meta, basta
acrescentar a linha na aba.

| Parâmetro (coluna Parâmetro) | Padrão |
|---|---|
| `Meta Disponibilidade (%)` | já existe: 95 |
| `Meta Inspeções/Mês` | já existe: 100 |
| `Meta Calibração em Dia (%)` | 90 |
| `Meta Aferições Conformes (%)` | 95 |
| `Meta Aderência Rotina (%)` | 90 |

**Regra da situação:**

- **Na meta:** valor ≥ meta.
- **Atenção:** até 10% abaixo da meta. Em aferições, até 5 pontos abaixo.
- **Crítico:** abaixo disso.
- **Inspeções:** comparadas com o esperado até hoje, proporcional ao dia do
  mês.

## Período

- **Visão Geral e Equipamentos & Paradas** mostram sempre o **mês
  corrente**. A calibração é uma foto do momento.
- **Aferições:** sem mês escolhido na lateral, mostram o mês corrente.
  Escolhendo um mês na lateral, mostram aquele mês (últimos 13 meses).

## Rotina das aferições

Desde a v19 a rotina segue a **regra do dia certo**. A tabela de rotinas,
a regra de atraso e como ajustar estão em
`docs/14-metas-e-origem-dos-numeros.md`, seção 3.

## O que foi verificado aqui e o que só o Power BI confirma

**Verificado aqui:**

- Os 425 JSON do relatório são válidos. Os visuais novos foram conferidos
  contra o esquema público do PBIR (microsoft/json-schemas).
- Todas as colunas e medidas que as páginas usam existem no modelo.
- As 109 medidas novas têm referências e parênteses conferidos por script.
- O M das 19 partições está íntegro. As vírgulas entre passos continuam
  checadas: foi a causa do erro da v12/v13.
- A navegação chega às 9 páginas. Os recursos estão registrados.
- O maior caminho de arquivo tem 133 caracteres (limite 259).

**Só o Power BI confirma:**

- A sintaxe DAX em si. Não há motor DAX aqui. Se alguma medida der erro, o
  visual mostra o aviso e o resto funciona.
- O visual exato das novas composições: cartões clássicos com cor por medida
  e barra 100% empilhada usada como barra de progresso.
- O desempenho de `SharePoint.Files`, que lista o site inteiro antes de
  filtrar a pasta.
