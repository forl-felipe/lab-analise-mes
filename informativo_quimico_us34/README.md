# Relatório de Turno do Laboratório Químico: Usinas 3 e 4

Planilha Excel com macros (`Relatorio_Quimico_US3_US4.xlsm`) para o Laboratório Químico. É a mesma ideia do
Informativo do Laboratório Físico (`../informativo_us34`):

- turnos de 12 horas (Dia 07h–19h, Noite 19h–07h) e letras A, B, C e D, com **um técnico por letra**;
- **modelo diário**: um arquivo por dia, com os dois turnos;
- **estrutura simples, 4 abas**: quase tudo funciona por fórmula. O VBA só busca o MES, copia a imagem e esconde as
  linhas vazias do resumo. Não há Painel nem "Finalizar turno".

A ocorrência segue a planilha que o laboratório usa hoje (`01-09-26.xlsm`). As análises e as tags seguem o
**Plano Amostral do Laboratório Físico/Químico de Ubu (rev. 06)**.

## As abas

| Aba | Para que serve |
|---|---|
| **Preenchimento** | A única aba em que o técnico escreve. No alto fica a **data do dia**; abaixo, o **Turno Dia** e, mais embaixo, o **Turno Noite**, com os mesmos campos: letra, técnico, produto US3/US4, tarefas realizadas, solicitações, equipamentos, tarefas a realizar, **Dragas**, **Mineroduto 03 (batch e teores)**, **Filtragem**, **comentários por usina**, controle do laboratório, cadinhos de platina e observações |
| **Resumo Dia** / **Resumo Noite** | Relatório técnico do turno (imagem para o e-mail), em formato paisagem e **no mesmo formato do Preenchimento**: identificação, ocorrências linha a linha (numeradas, com o número junto do texto), comentários por usina, controle do laboratório, cadinhos, observações e resultados do MES por **horário de amostra: Dia 07:30 a 19:30, Noite 19:30 a 07:30** (o resultado das 19:30 aparece nos dois turnos). Tudo vem por fórmula. Botões: **Atualizar dados do MES** e **Copiar imagem** |
| **Resultados gerais** | Resultados químicos de qualquer período de até 24 h, nas **mesmas janelas do turno (amostras 07:30, 09:30…)**; o início é ajustado para o começo da janela de 2 h em que está: **Início** e **Fim** no alto, **Atualizar dados do MES** e **Copiar imagem** |
| Limites (oculta) | **Farol por produto**: limites químicos de cada produto (SMIN-POP-GEA-001 rev. 12). Pode ser editada e ampliada à mão (botão direito numa guia → Reexibir → Limites) |
| Configurações (oculta) | Fonte dos dados, servidor, **tags do MES**, LIE/LSE e faixa válida. Para abrir: botão direito numa guia → Reexibir |

**Texto longo:** no Preenchimento e no Resumo o texto quebra a linha, e a altura da linha se ajusta sozinha (macro).
**O que não foi preenchido não vai para o Resumo:** linha vazia, campo vazio e assunto inteiro vazio (ex.: Cadinhos de
platina, Equipamentos, Comentário US4, Controle do laboratório) ficam ocultos, inclusive o título. Os campos do
Controle do laboratório não vêm mais pré-preenchidos.

Navegação por hiperlinks (sem macro): "Ir para: Turno Noite / Resumo Dia / Resumo Noite / Resultados gerais" e "Voltar ao Preenchimento".

## Rotina do técnico

1. Todo dia: copie o arquivo modelo para a pasta do mês e renomeie com a data. Ao abrir, a **data do dia** é preenchida pelo relógio, se estiver vazia.
2. Preencha o seu turno na aba **Preenchimento**.
3. Abra o **Resumo** do seu turno: as informações já estão lá. Clique em **Atualizar dados do MES** e depois em **Copiar imagem**, e cole no e-mail (Ctrl+V). Também dá para tirar o print com Windows + Shift + S.

Os resultados de cada Resumo ficam gravados na própria aba. O turno Noite atualiza o seu resumo sem mexer no do Dia.

## Farol (verde / vermelho) por produto

No **Preenchimento**, cada turno escolhe o **Produto US3** e o **Produto US4** numa lista (PDR/MX, PDR/STD, PBF/MB45, PBF/STD, PBF/HB, PBF/SF, PBF/SA). Os resultados dos Resumos são comparados com os limites desse produto:

- **VERDE**: dentro do limite, ou seja, **igual ou acima do mínimo** e **igual ou abaixo do máximo**. A comparação usa o valor arredondado, como aparece na tela;
- **VERMELHO**: fora do limite;
- sem cor: análise sem limite para o produto, ou produto não selecionado.

O farol vale para as janelas de 2 h e para a média do turno. Nos Resultados gerais, os produtos vêm do Turno Dia e podem ser trocados ali.

Limites usados, conforme o padrão **SMIN-POP-GEA-001 rev. 12**:

- **Pelota Queimada US3/US4:** colunas Máx./Mín. do **Processo (GPU), dados horários** (itens 10.4.1 e 10.4.2). Fe mín., SiO2 máx., P máx., CaO mín. (PDR) e B2 (PBF). Nas revisões com seta, vale o valor novo (ex.: Fe PDR/STD 67,10 → **67,19**).
- **Pellet Feed (Filtragem, linha US3/4):** limites do **concentrado** do produto da US3 (ou da US4, se a US3 estiver sem produto). Concentrados: CLS, CNS, CHS ou CSP.
  - SiO2 máx. bi-horário do concentrador III (item 10.2);
  - P e PPC máx. diários do batch (item 10.1).

| Produto | Concentrado | Fe mín. | SiO2 máx. | P máx. | CaO mín. | B2 |
|---|---|---|---|---|---|---|
| PDR/MX | CLS | 67,39 | 1,54 | 0,050 | 0,70 | – |
| PDR/STD | CNS | 67,19 | 2,05 | 0,074 | 0,65 | – |
| PBF/MB45 | CHS | 65,30 | 3,45 | 0,074 | – | máx. 0,55 |
| PBF/STD | CHS | 65,10 | 3,20 | 0,074 | – | mín. 0,75 |
| PBF/HB | CNS | 65,30 | 2,80 | 0,074 | – | mín. 0,95 |
| PBF/SF | CNS | 65,00 | 3,00 | 0,072 | – | mín. 1,10 |
| PBF/SA | CSP | 63,50 | 5,30 | 0,100 | – | mín. 0,40 |

| Concentrado (Pellet Feed) | SiO2 máx. | P máx. | PPC máx. |
|---|---|---|---|
| CLS | 1,36 | 0,050 | 4,30 |
| CNS | 1,99 | 0,075 | 4,30 |
| CHS | 2,50 | 0,075 | 4,30 |
| CSP | 5,30 | 0,110 | 4,30 |

O P da pelota não está no Plano Amostral (Pelota Queimada), por isso não aparece no relatório.

## Análises e tags do MES

Tags montadas a partir do Plano Amostral (rev. 06): `M` + ponto + `-` + código da análise (4 dígitos) + `-` + fase + área.
Exemplo: `M650030010-0004-HHLQU` (fase HH = bi-horário; LQU = Laboratório Químico).

Só entram os **resultados químicos bi-horários (HH)** da Filtragem e da Pelota Queimada. Insumos e demais pontos
ficam fora.

| Amostra | US3 | US4 |
|---|---|---|
| Filtragem, Pellet Feed | `M650030010` (U03-02TP004), **uma linha só "US3/4"** (mesmo material nas duas usinas) | – |
| Pelota Queimada | `M710050020` (U03-07TP002) | `M4710050020` (U04-07TP015) |

| Amostra | Análises (código) |
|---|---|
| Filtragem | SiO2 (0004), P (0008), MgO (0007), CaO (0006), PPC (0013) |
| Pelota Queimada | Fe (0002), SiO2 (0004), Al2O3 (0005), CaO (0006), MgO (0007), B2 (0018), Mn (0115). Sem PPC (não é feito na pelota) |

**Pontos de atenção (conferir no MES):**

- **Fe e Al2O3 da Filtragem** são **diários (DD)** no plano e ficaram fora do relatório de turno.
  - US3: `M650030010-0002-DDLQU` e `-0005-DDLQU`;
  - US4: `M4650030010-…` (no plano, o ponto da US4 é `4650030010` para Fe/Al2O3 e `4650060010` para as demais análises).
- **Pelota Queimada:** no plano, os códigos de **CaO (0115), MgO (0006), B2 (0122), Mn (0018) e PPC (0007)** não batem com os
  códigos de todos os outros pontos do próprio plano (embarque, estocagem, filtragem) nem com as tags que já funcionam
  nas planilhas atuais (CaO 0006, B2 0018). Por isso, foram usados os códigos consistentes:
  - CaO 0006, MgO 0007, B2 0018, Mn 0115, PPC 0013.

  Se o MES confirmar os códigos do plano, basta trocar as tags em **Configurações**.

Toda tag pode ser alterada na aba **Configurações**. Análise sem tag fica oculta e fora da consulta.

## Consulta ao MES

**Velocidade:**
- A consulta é calculada uma única vez por clique (antes o Excel a recalculava mais de uma vez).
- Depois de ler os resultados, a fórmula do Aspen vira texto e as matrizes são apagadas, para o MES não ser consultado de novo a cada edição da planilha.
- O início da fórmula (como o Excel a mostra) fica guardado em `Dados_MES!E3`, e o tempo da última consulta em `Dados_MES!E4`.

**Imagem em alta resolução:** com zoom em 100%, a área é copiada como figura vetorial (aparência de impressão), ampliada num gráfico temporário até cerca de 3600 px de largura e exportada como PNG; esse PNG vai para a área de transferência. Se algo falhar, é usada a cópia comum.

Uma única consulta `GetCalculationValues` com todas as análises configuradas: tipo de cálculo `"1"`, janelas de `"2h"` e listas em texto literal, no mesmo formato validado no Físico. A consulta inclui só as análises com as duas tags preenchidas.

- **O botão pode ser usado a qualquer hora do turno**, para acompanhar os resultados:
  - consulta só até o fim da janela de 2 h atual e nunca pede horários futuros ao MES;
  - a janela em andamento aparece com `*` (média parcial);
  - as janelas seguintes ficam em branco;
  - no início do turno, sem análises ainda, a consulta termina em cerca de 15 s com o aviso "ainda não há resultados", sem erro;
  - um segundo clique durante a consulta é ignorado com um aviso.
- Cada botão do turno consulta 7 janelas de 2 h (14 h). Os Resultados gerais fazem mais de uma consulta para períodos maiores.
- **Tag recusada pelo MES** ("Tag Name … is invalid"): uma tag inválida derrubava a consulta inteira. Agora a planilha
  lê o erro, tira as tags inválidas da consulta, consulta de novo com as demais e avisa quais tags foram recusadas.
  A lista também fica em `Dados_MES!E5`. Basta corrigir as tags na aba Configurações.
- Antes de cada consulta, as matrizes de resultado do Aspen são apagadas, para evitar o erro "Não é possível alterar parte de uma matriz".
- A macro espera os valores estabilizarem antes de gravar.
- Os valores de 2 em 2 h são gravados direto na tabela da aba. Média, mínimo e máximo são fórmulas, e o laranja vem de formatação condicional (LIE/LSE em Configurações).

## Primeiro uso

1. Desbloqueie o arquivo baixado: Propriedades → **Desbloquear**. Depois clique em **Habilitar conteúdo**. Sem macros, o preenchimento e os resumos continuam funcionando; só os botões do MES e da imagem ficam inativos.
2. No **arquivo modelo**, confira as tags e preencha LIE/LSE em Configurações.
3. Para treinar sem o MES: **Fonte dos dados** = `SIMULAÇÃO`.
4. O suplemento **Aspen Process Data** precisa estar ativo, como no Informativo do Físico.

Guia de uma página: [`guia/Guia_rapido_Relatorio_Quimico.png`](guia/Guia_rapido_Relatorio_Quimico.png).

## Como gerar a planilha

```
cd build
python3 build_quimico.py           # gera ../Relatorio_Quimico_US3_US4.xlsm
python3 build_quimico.py --teste   # versão com o módulo de teste automático (LibreOffice)
```

O código VBA fica em `vba/`. O `ModLayout.bas` é gerado pelo build.

Validação no LibreOffice (modo simulação):

- fórmula de consulta montada pelo VBA idêntica à do Python;
- migração por fórmula e ocultação das linhas vazias;
- Resumo Dia e Resumo Noite independentes;
- análises sem tag ocultas;
- Resultados gerais de 24 h e de 10 h (colunas que sobram ficam ocultas);
- leitura das tags recusadas pelo MES a partir do texto de erro do Aspen;
- limpeza e recriação das matrizes do MES.
