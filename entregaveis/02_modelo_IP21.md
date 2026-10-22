# Modelo de dados no PIMS (Aspen InfoPlus.21): prensa automática de compressão

Documento da Fase 1. Substitui o entregável "02_modelo_AF.md" previsto no planejamento inicial, pois o PIMS da Unidade UBU utiliza a plataforma AspenTech (Aspen InfoPlus.21 e Aspen Process Graphics V14.2), e não o PI System. Ver hipótese H01 em `hipoteses.md`.

Os nomes de campos citados seguem o padrão das definições do IP.21 e devem ser conferidos pelo administrador do PIMS no ambiente da Samarco, com o Aspen InfoPlus.21 Administrator.

## 1. Caminho do dado

```
Prensa automática  >  PLC 307-PLC-001  >  DeltaV01 (servidor OPC)  >  Aspen Cim-IO  >  IP.21 (fonte UBU)  >  Aspen Process Graphics
```

| Etapa | Responsável provável | O que conferir |
|---|---|---|
| PLC e DeltaV | Automação | Lógica dos discretos 307-QI-001 a 003, tipo do parâmetro PV_D, momento em que os valores são atualizados |
| Cim-IO e IP.21 | Administração do PIMS (atendimento do RITM0273793) | Nome dos registros, definição, unidade, arquivamento histórico |
| Tela | LCE, com apoio do PIMS | Montagem no Process Graphics Editor e publicação |

## 2. Registros

Os cinco registros já foram criados pelo chamado RITM0273793 (última leitura em 30/09/2026, 11:28, qualidade Good). Este documento não propõe recriá-los, apenas completar a configuração que a tela utiliza.

### 2.1 Analógicos: 307-GI-001 (carga média) e 307-GI-002 (desvio padrão)

| Campo IP.21 | 307-GI-001 | 307-GI-002 | Por que importa para a tela |
|---|---|---|---|
| IP_DESCRIPTION | Carga média | Desvio padrão | Exibida na tabela de tags e nas dicas de tela |
| IP_ENG_UNITS | kgf/pel | kgf/pel | Unidade padrão do LCE, igual à tela atual do Laboratório Químico |
| IP_GRAPH_MINIMUM / IP_GRAPH_MAXIMUM | 0 / 400 (confirmada) | 0 / 100 (proposta) | Escala padrão das tendências, evitando ajuste manual em cada tela |
| IP_VALUE_FORMAT | 1 casa decimal | 1 casa decimal | Exibição com vírgula decimal, conforme a configuração regional da estação |
| IP_INPUT_QUALITY | Leitura | Leitura | Base para exibir "Sem dado confiável" quando diferente de Good |
| IP_INPUT_TIME | Leitura | Leitura | Base para a indicação de defasagem e de perda de comunicação |
| IP_ARCHIVING | ON | ON | Necessário para a tendência de 8 horas e para o histórico |
| IP_DC_SIGNIFICANCE | 0 (proposta) | 0 (proposta) | Garante que todo resultado de ensaio seja gravado, mesmo repetido. Como a atualização é por ensaio, o volume de dados é pequeno |
| IP_LOW_LIMIT e demais limites | Não preencher por enquanto | Não preencher | Não há meta de especificação definida (H03). Os limites serão revistos após a aprovação da severidade na Fase 2 |

### 2.2 Discretos: 307-QI-001, 307-QI-002 e 307-QI-003

| Campo IP.21 | Valor proposto | Por que importa |
|---|---|---|
| IP_DESCRIPTION | Menor que 100, Menor que 150, Menor que 200 | Texto da tabela de tags |
| Conjunto de estados | 0 = Inativo; 1 = Ativo | A tela mostra o estado por texto, e não só por cor (ISA-101) |
| IP_ARCHIVING | ON | Permite consultar quando cada faixa foi acionada |
| IP_INPUT_QUALITY / IP_INPUT_TIME | Leitura | Mesmo tratamento de qualidade e defasagem dos analógicos |

## 3. Cálculos opcionais (Aspen Calc ou registro calculado)

Nenhum cálculo é obrigatório para a primeira versão da tela. Os itens abaixo são propostas para avaliação, cada um com sua justificativa.

| Cálculo | Fórmula | Justificativa | Prioridade |
|---|---|---|---|
| 307-GI-001_SUP e 307-GI-001_INF | Carga média mais e menos o desvio padrão | A tendência do Process Graphics plota tags, não expressões. Sem esses dois registros, a banda de mais ou menos um desvio padrão não pode ser desenhada | Recomendado, se a banda for mantida no projeto |
| 307-QI_FAIXA | 0 a 3, indicando a faixa mais severa ativa | Um único campo multiestado na tela, em vez de três sinais independentes. Útil também para relatórios | Opcional |
| 307-GI-001_IDADE | Minutos desde o último IP_INPUT_TIME | Permite indicar "sem ensaio recente" e "possível perda de comunicação" por cor e texto na tela | Recomendado, caso o Process Graphics não consiga calcular a idade do dado diretamente (avaliar na Fase 3) |

Esses registros precisam ser criados pelo administrador do PIMS. Sem eles, a tela funciona com as cinco tags, mas sem banda de desvio e sem indicação automática de defasagem.

## 4. Reaproveitamento para outras prensas

O IP.21 não possui um equivalente direto aos templates do PI AF. A replicação será feita assim:

1. **Padrão de nomes:** os registros seguem o padrão `<área>-<função>-<número>`, como 307-GI-001. Uma nova prensa terá o mesmo conjunto, com outro prefixo de área.
2. **Tela modelo:** a tela será montada uma única vez, com todas as tags agrupadas e identificadas. Para outra prensa, a tela é copiada e as tags são substituídas pela função de substituição de tags do Process Graphics Editor. O procedimento será detalhado e conferido na Fase 3.
3. **Dicionário:** cada nova prensa ganha cinco linhas no `01_dicionario_tags.csv`, com o mesmo formato.

## 5. Pendências deste documento

- Confirmar nome dos registros, fonte de dados e definição (H04).
- Confirmar o tipo do parâmetro PV_D no DeltaV para as tags analógicas (H09).
- Decidir sobre os cálculos opcionais da seção 3 com o administrador do PIMS.
