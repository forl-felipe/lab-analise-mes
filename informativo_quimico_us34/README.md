# Relatório de Turno do Laboratório Químico: Usinas 3 e 4

Planilha Excel com macros (`Relatorio_Quimico_US3_US4.xlsm`) para o Laboratório Químico. É a mesma ideia do
Informativo do Laboratório Físico (`../informativo_us34`):

- turnos de 12 horas (Dia 07h–19h, Noite 19h–07h) e letras A, B, C e D, com **um técnico por letra**;
- **modelo diário**: um arquivo por dia, com os dois turnos;
- **estrutura simples, 4 abas**: quase tudo funciona por fórmula. O VBA só busca o MES, copia a imagem e esconde as
  linhas vazias do resumo. Não há Painel nem "Finalizar turno".

A ocorrência segue a planilha que o laboratório usa hoje (`01-09-26.xlsm`). As análises seguem a planilha
`CONTROLE DE PRODUÇÃO USINAS 03 E 04` (aba Resultados).

## As 4 abas

| Aba | Para que serve |
|---|---|
| **Preenchimento** | A única aba em que o técnico escreve. No alto fica a **data do dia**; abaixo, o **Turno Dia** e, mais embaixo, o **Turno Noite**, com os mesmos campos: letra e técnico, tarefas realizadas, solicitações, equipamentos, tarefas a realizar, controle do laboratório, pessoal, cadinhos de platina e observações |
| **Resumo Dia** / **Resumo Noite** | Página pronta para o e-mail. As ocorrências vêm **sozinhas, por fórmula**, do Preenchimento. Tem **2 botões**: **Atualizar dados do MES** (busca os resultados químicos do período do turno na data do Preenchimento) e **Copiar imagem**. Linhas vazias ficam ocultas ao abrir a aba e antes de copiar |
| **Resultados gerais** | Resultados químicos de qualquer período de até 24 h: **Início** e **Fim** no alto, **Atualizar dados do MES** e **Copiar imagem** |
| Limites (oculta) | **Farol por produto**: limites químicos de cada produto (SMIN-POP-GEA-001 rev. 12). Pode ser editada e ampliada à mão (botão direito numa guia → Reexibir → Limites) |
| Configurações (oculta) | Fonte dos dados, servidor, **tags do MES**, LIE/LSE e faixa válida. Para abrir: botão direito numa guia → Reexibir |

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

- **Pelota (Linha de Mistura US3/US4):** colunas Máx./Mín. do **Processo (GPU), dados horários** (itens 10.4.1 e 10.4.2). Fe mín., SiO2 máx., P máx., CaO mín. (PDR) e B2 (PBF). Nas revisões com seta, vale o valor novo (ex.: Fe PDR/STD 67,10 → **67,19**).
- **Pellet Feed MD03:** limites do **concentrado** da pelota da US3 (CLS, CNS, CHS ou CSP):
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

Fe e P do farol passam a aparecer quando as tags dessas análises forem configuradas.

## Análises e tags do MES

Pontos de amostragem (os mesmos do Informativo do Físico e da planilha padrão de 2014):

| Amostra | Ponto no MES |
|---|---|
| Pellet Feed, Mineroduto 03 (amostra única para US3 e US4, "MD03") | `M650030010` |
| Linha de Mistura / Pelota US3 | `M710050020` |
| Linha de Mistura / Pelota US4 | `M4710050020` |

Códigos de análise (sufixo `-HHLQU` = Laboratório Químico) **já confirmados** nas planilhas existentes:

| Código | Análise | Pellet Feed | Mistura/Pelota US3 e US4 |
|---|---|---|---|
| 0004 | SiO2 | ✔ | ✔ |
| 0006 | CaO | ✔ | ✔ |
| 0013 | PPC | ✔ | |
| 0018 | B2 | | ✔ |
| 0084 | Carvão (kg/t) | | ✔ |
| 0493 | Carbono fixo | | ✔ |

**Ainda sem tag** (não aparecem em nenhuma das planilhas recebidas): FeT, Al2O3, MgO, P, Mn e TiO2 (Pellet Feed e Mistura/Pelota) e pH do Mineroduto.

Essas análises já estão no relatório, mas ficam **ocultas** e **fora da consulta** até a tag ser preenchida na aba Configurações (célula em laranja). Não foi colocada nenhuma tag "chutada". Quando o código for confirmado no MES, basta digitar a tag (ex.: `M710050020-00xx-HHLQU`) e atualizar. A linha aparece sozinha no relatório.

## Consulta ao MES

Uma única consulta `GetCalculationValues` com todas as análises configuradas: tipo de cálculo `"1"`, janelas de `"2h"` e listas em texto literal, no mesmo formato validado no Físico. A consulta inclui só as análises com as duas tags preenchidas.

- **O botão pode ser usado a qualquer hora do turno**, para acompanhar os resultados:
  - consulta só até o fim da janela de 2 h atual e nunca pede horários futuros ao MES;
  - a janela em andamento aparece com `*` (média parcial);
  - as janelas seguintes ficam em branco;
  - no início do turno, sem análises ainda, a consulta termina em cerca de 15 s com o aviso "ainda não há resultados", sem erro;
  - um segundo clique durante a consulta é ignorado com um aviso.
- Cada botão consulta 12 h (6 janelas). Os Resultados gerais fazem mais de uma consulta para períodos maiores.
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
- limpeza e recriação das matrizes do MES.
