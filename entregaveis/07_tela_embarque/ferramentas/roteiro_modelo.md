# Passo a passo: tela Qualidade de Embarque no Aspen (versão 1)

**Tela:** Qualidade de Embarque, padrão visual Samarco: ritmo de carregamento, total carregado por berço e por porão, resultados físicos e químicos do embarque e tendências

**Resultado esperado:** igual a `tela_exemplo_normal.png` (dados simulados)

**Tempo estimado:** 20 minutos

![Resultado esperado](tela_exemplo_normal.png)

## O que vem pronto no arquivo

O arquivo `Qualidade_de_Embarque.atgraphic` foi gerado pelo mesmo processo da tela Compressão Online Usina 4, que já abriu no seu PIMS. Tudo é objeto nativo do Aspen e não perde nitidez ao maximizar. A tela não tem limites de especificação: mostra apenas os resultados.

| Conteúdo | Quantidade |
|---|---|
| Painéis, linhas e logotipo oficial em polígonos | {{N_FORMAS}} |
| Textos fixos | {{N_TEXTOS}} |
| Campos de dados ligados às tags | {{N_CAMPOS}} |
| Tendências com penas e escala definidas | {{N_TEND}} |
| Lâmpadas animadas (berço com carga) | 2 |

## Parte A. Abrir

**A1.** Copie `Qualidade_de_Embarque.atgraphic` para a pasta de rede das telas do laboratório.

**A2.** Abra no **Aspen Process Graphics Editor**, compare com `tela_exemplo_normal.png` e maximize.

**Se der erro ao abrir:** tente `teste_minimo.atgraphic` e me envie a captura da mensagem, dizendo se o arquivo de teste abriu.

## Parte B. Ajustes nos campos (no editor)

**B1. Atributo de data e hora.** Os campos "Último resultado" do Físico (objeto 23) e do Químico (objeto 30) vêm com o atributo `VAL` e mostram um número no lugar da hora. Dê duplo clique em cada um e troque `VAL` pelo atributo de data e hora da lista. Se preferir, apague os dois campos e o texto "Último resultado".

**B2. Mapa das tags de qualidade.** Os 15 resultados físicos e químicos usam o mapa `IP_MESVALOR`, o mesmo da tag de compressão do laboratório que já funciona na outra tela. Se algum campo mostrar `***`, abra o campo e confira o mapa na lista.

**B3. Carregamento.** Ritmo, totais e porões usam o mapa `IP_ANALOGMAP`, como na tela de embarque atual. Nada a fazer.

## Parte C. Tendências

| Objeto | Penas | Escala |
|---|---|---|
| 40 (ritmo de carregamento) | `008-WI-010` (balança principal, azul claro) | 0 a 2500 t/h |
| 41 (compressão) | `M0620010000-0031-TNLFU` (16,0 mm, azul claro) e `M0620010000-0035-TNLFU` (12,5 mm, branca) | 150 a 450 nas duas |
| 42 (Fe e SiO2) | `M0620010000-0002-TNLQU` (Fe, azul claro) e `M0620010000-0004-TNLQU` (SiO2, branca) | Fe 62 a 68; SiO2 0 a 6 |

**C1.** Na barra de tempo de cada tendência, use o período que preferir. Sugestão: **`24:00:00`** no ritmo e **`72:00:00`** na qualidade (os resultados do embarque são menos frequentes).

**C2.** Os números das escalas são parte fixa da tela. Se mudar a escala de uma pena, os números deixam de coincidir com a linha. Se os valores reais ficarem fora das escalas, me avise que eu ajusto.

## Parte D. Lâmpadas (já prontas)

| Lâmpada | Tag | Acende quando |
|---|---|---|
| Berço Oeste | `309-WIT-004.TOT_NAV_OE` | total carregado maior que zero, em verde |
| Berço Leste | `309-WIT-004.TOT_NAV_LE` | total carregado maior que zero, em verde |

A lâmpada indica o berço com carga registrada. Se o totalizador não for zerado ao fim de cada navio, as duas podem ficar acesas; nesse caso, me passe a tag que indica o berço em operação.

## Parte E. Salvar

**E1.** **Arquivo > Salvar**. Abra no Process Explorer e maximize para conferir.

## Parte F. Problemas comuns

| Problema | Solução |
|---|---|
| Erro ao abrir | Tente `teste_minimo.atgraphic` e envie a captura (Parte A) |
| Hora mostra um número | Parte B1 |
| Campo com `***` | Tag ou mapa não encontrado: confira na lista do campo (Parte B2) |
| Linha fora da escala da tendência | Me avise para ajustar a escala (Parte C2) |

## Anexo. Fichas dos objetos dinâmicos

Referência caso algum objeto precise ser refeito à mão. Posições em unidades da tela de 50 x 35.

![Mapa de posicionamento](mapa_posicionamento.png)

{{FICHAS}}
