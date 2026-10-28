# Passo a passo: tela de compressão de pelotas no Aspen (versão 7)

**Tela:** Monitoramento do ensaio de compressão de pelotas, padrão visual Samarco, comparação composto com composto (2 h)

**Resultado esperado:** igual a `tela_exemplo_normal.png`

**Tempo estimado:** 30 minutos

![Resultado esperado](tela_exemplo_normal.png)

## O que vem pronto no arquivo

O arquivo `LCE_Compressao_Pelotas_Prensa01.atgraphic` segue o mesmo formato da versão que já abriu no seu PIMS. Tudo é objeto nativo do Aspen e não perde nitidez ao maximizar:

| Conteúdo | Quantidade |
|---|---|
| Painéis, linhas, limites tracejados e logotipo oficial em polígonos | {{N_FORMAS}} |
| Textos fixos | {{N_TEXTOS}} |
| Campos de dados ligados às tags | {{N_CAMPOS}} |
| Tendências com penas e escala definidas | {{N_TEND}} |
| Lâmpadas animadas (cortadores e faixas de compressão) | {{N_LAMPADAS}} |

## Parte A. Abrir

**A1.** Copie `LCE_Compressao_Pelotas_Prensa01.atgraphic` para a pasta de rede das telas do laboratório. Se a versão anterior estiver aberta, feche-a antes de substituir.

**A2.** Abra no **Aspen Process Graphics Editor**, compare com `tela_exemplo_normal.png` e maximize.

**Se der erro ao abrir:** tente `teste_minimo.atgraphic` e me envie a captura da mensagem, dizendo se o arquivo de teste abriu.

## Parte B. Ajustes nos campos (no editor)

**B1. Atributo de data e hora.** Os campos abaixo vêm com o atributo `VAL` e mostram um número no lugar da hora. Dê duplo clique em cada um e troque `VAL` pelo atributo de data e hora da lista:

| Objeto | Campo |
|---|---|
| 1 | Última atualização (cabeçalho) |
| 4 | Média prensa automática: atualizado às |
| 8 | Desvio padrão: atualizado às |
| 12 | Composto 2 h: fim da janela |
| 15 a 18 | Cortes do composto: hora de cada corte |

Dica: depois de achar o atributo certo no objeto 1, me diga o nome. Na próxima versão ele já vem configurado em todos.

**B2. Atributo de qualidade** nos objetos 5 e 9 (qualidade da média e do desvio padrão), com a regra de cor: `Good` em branco e qualquer outro texto em laranja (RGB 243, 112, 33).

**B3. Mapa dos discretos.** Os objetos 23, 24 e 25 (faixas de compressão) e as três lâmpadas ao lado deles usam o mapa `IP_DISCRETEMAP`, o mesmo da versão anterior, que funcionou. Nada a fazer, a menos que apareça erro.

## Parte C. Tendências

| Objeto | Penas | Escala |
|---|---|---|
| 3 (minitendência da média) | `307-GI-001` | 0 a 500 |
| 7 (minitendência do desvio padrão) | `307-GI-002` | 0 a 150 |
| 14 (tendência principal) | `307-GI-001` (cada ensaio, linha azul escura), `M4710050020-0031-HHLFU` (laboratório, branca) e `307-GI-001-C2H` (composto da prensa, azul claro) | 50 a 450 nas três |

**C1.** Na barra de tempo, embaixo da janela, use **`8:00:00`**.

**C2.** Os números de 100 a 400 e as linhas de 200, 150 e 100 são parte fixa da tela. **Não altere a escala das penas.**

**C3. Teste do composto sem tag calculada:** nas propriedades da pena `307-GI-001` da tendência 14, veja se existe opção de **média por período** (agregado). Se existir, adicione mais uma pena `307-GI-001` com **média de 2 h**, cor azul claro (RGB 89, 198, 242) e escala 50 a 450. Ela desenha o composto da prensa enquanto a tag calculada não existe.

## Parte D. Lâmpadas (já prontas)

| Lâmpadas | Tag | Acende quando |
|---|---|---|
| 07CR001 e 07CR002, Automático | `U04-07CR001`, `U04-07CR002` (mapa `IP_PIDMAP03`) | valor 32, em verde |
| 07CR001 e 07CR002, Manual | mesmas tags | valor 16, em laranja |
| Faixas abaixo de 200, 150 e 100 kgf/pelota | `307-QI-003`, `307-QI-002`, `307-QI-001` | valor maior que zero, na cor da faixa |

Os valores 32 e 16 foram tirados dos indicadores Manual/Automático da tela PQ US04. Confira com o cortador em manual: a lâmpada laranja deve acender.

## Parte E. Salvar

**E1.** **Arquivo > Salvar**. Abra no Process Explorer e maximize para conferir.

## Parte F. Tags calculadas (pedir ao administrador do PIMS)

Os campos do composto da prensa (10), da diferença (13), da quantidade de cortes (26) e dos cortes (15 a 22) e a linha azul clara da tendência já estão ligados às tags abaixo. **Assim que o administrador criar as tags com estes nomes, a tela passa a mostrar os valores sozinha**, sem nenhuma edição. Até lá, o Aspen mostra `***` nesses campos, que é a indicação de tag inexistente.

| Tag | Cálculo |
|---|---|
| `307-GI-001-C2H` | Média de `307-GI-001` **dos cortes com resultado** na janela de 2 h do composto do laboratório (1 a 4 cortes), gravada no fim da janela |
| `307-GI-001-N2H` | Quantidade de cortes com resultado na janela (0 a 4) |
| `307-GI-001-DIF2H` | `307-GI-001-C2H` menos `M4710050020-0031-HHLFU` da mesma janela |
| `307-GI-001-CORTE1` a `307-GI-001-CORTE4` | Resultado de `307-GI-001` de cada um dos 4 cortes da janela atual, gravado no horário do corte; sem valor quando o corte não teve resultado |

Detalhes (janela pelo horário do corte, defasagem de 20 min até a entrada no PIMS, janelas incompletas e cortadores em manual) em `../05_proposta_visual/composto_2h_especificacao.md`. Se o administrador preferir outros nomes, me informe que eu regero o arquivo.

## Parte G. Problemas comuns

| Problema | Solução |
|---|---|
| Erro ao abrir | Tente `teste_minimo.atgraphic` e envie a captura (Parte A) |
| Hora mostra um número | Parte B1 |
| Campos do composto com `***` | Normal até as tags calculadas existirem (Parte F) |
| Lâmpada de modo não muda | Confira se o valor de Manual é 16 e de Automático é 32 na tela PQ US04 |
| Escala da tendência diferente | Volte para 50 a 450 (Parte C) |

## Anexo. Fichas dos objetos dinâmicos

Referência caso algum objeto precise ser refeito à mão. Posições em unidades da tela de 50 x 35.

![Mapa de posicionamento](mapa_posicionamento.png)

{{FICHAS}}
