# Relatório de Turno do Laboratório Químico: Usinas 3 e 4

Planilha Excel com macros (`Relatorio_Quimico_US3_US4.xlsm`) para o Laboratório Químico. É a mesma ideia do
Informativo do Laboratório Físico (`../informativo_us34`):

- turnos de 12 horas (Dia 07h–19h, Noite 19h–07h) e letras A, B, C e D, com **um técnico por letra**;
- **modelo diário**: um arquivo por dia, com os dois turnos;
- **resultados químicos puxados do MES** (Aspen IP.21) de 2 em 2 horas, com a mesma consulta do Físico;
- **relatório do turno em uma página**: ocorrências em cima e informativo de qualidade químico embaixo, pronto
  para copiar como imagem e colar no e-mail.

A ocorrência segue a planilha que o laboratório usa hoje (`01-09-26.xlsm`). As análises seguem a planilha
`CONTROLE DE PRODUÇÃO USINAS 03 E 04` (aba Resultados).

## Abas

| Aba | Para que serve |
|---|---|
| **Painel** | Data, turno, letra e técnico; 4 botões; situação do turno; indicadores (SiO2 PF, SiO2 e B2 da mistura/pelota, carbono fixo) |
| **Relatório Dia** / **Relatório Noite** | Uma página por turno: ocorrências + resultados de 2 em 2 h com média, mín. e máx. (fora da especificação em laranja). Botão **Copiar relatório** (imagem em alta resolução) ou print com Windows+Shift+S. Os valores são gravados como uma "foto" do turno |
| **Ocorrência** | Entrada de dados: tarefas realizadas, solicitações, equipamentos, tarefas a realizar, controle do laboratório (programas em uso, padrões, sistema de ar), pessoal (ausência, troca, hora extra, letra que recebe), cadinhos de platina e observações |
| **Informativo** | Resultados do MES de qualquer período de até 24 h (Início/Fim no alto da aba), com **Copiar como imagem** |
| **Configurações** (oculta) | Fonte dos dados, servidor, **tags do MES**, LIE/LSE e faixa válida. Para abrir: botão direito numa guia → Reexibir |
| Dados_MES (oculta) | Fórmula do suplemento Aspen (`GetCalculationValues` / `ShowCalculationValues`) |

## Os 4 botões do Painel

1. **Atualizar dados do MES**: busca os resultados químicos do turno e atualiza o relatório do turno.
2. **Ocorrência do turno**: abre a aba de preenchimento.
3. **Relatório do turno**: gera (ou refaz) a página do turno selecionado e a abre.
4. **Finalizar turno**:
   - busca o MES de novo, se preciso;
   - fixa o relatório do turno e salva o arquivo;
   - oferece **preparar a ocorrência do próximo turno**: as "Tarefas a realizar" passam para "Tarefas realizadas" e os demais campos voltam ao padrão.

**Editar turno Dia** / **Editar turno Noite** trocam o turno sem mudar a data. Um turno que ainda não começou pode ser preenchido; os resultados ficam em branco.

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

Uma única consulta `GetCalculationValues` com todas as análises configuradas: tipo de cálculo `"1"`, janelas de `"2h"`, listas em texto literal, no mesmo formato validado no Físico. A consulta inclui só as análises com as duas tags preenchidas. O resto é igual ao Físico:

- limpeza das matrizes de resultado antes de cada consulta, para evitar o erro "Não é possível alterar parte de uma matriz";
- espera até os valores estabilizarem;
- período livre de até 24 h, feito com duas consultas de 12 h.

## Primeiro uso

1. Desbloqueie o arquivo baixado: Propriedades → **Desbloquear**. Depois clique em **Habilitar conteúdo**.
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
- fluxo: MES → ocorrência → finalizar Dia → preparar a Noite → finalizar Noite, com o Relatório Dia intacto;
- análises sem tag ocultas;
- período de 24 h;
- limpeza e recriação das matrizes do MES.
