# Passo a passo: tela de compressão de pelotas no Aspen (versão 7)

**Tela:** Monitoramento do ensaio de compressão de pelotas, padrão visual Samarco, comparação composto com composto (2 h)

**Resultado esperado:** igual a `tela_exemplo_normal.png`

**Tempo estimado:** 30 minutos

![Resultado esperado](tela_exemplo_normal.png)

## O que vem pronto no arquivo

O arquivo `LCE_Compressao_Pelotas_Prensa01.atgraphic` segue o mesmo formato da versão que já abriu no seu PIMS. Tudo é objeto nativo do Aspen e não perde nitidez ao maximizar:

| Conteúdo | Quantidade |
|---|---|
| Painéis, linhas, limites tracejados e logotipo oficial em polígonos | 178 |
| Textos fixos | 55 |
| Campos de dados ligados às tags | 23 |
| Tendências com penas e escala definidas | 3 |
| Lâmpadas animadas (cortadores e faixas de compressão) | 7 |

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

### Objeto 1: Última atualização (cabeçalho)

| Propriedade | Valor |
|---|---|
| Ferramenta | Campo de dados |
| Posição em unidades (X ; Y) | 42,75 ; 1,03 |
| Tamanho em unidades (largura ; altura) | 6,25 ; 0,81 |
| Tag | `307-GI-001` |
| Mapa | `IP_ANALOGMAP` |
| Atributo | `VAL` no arquivo: **trocar pelo atributo de data e hora** (Parte B) |
| Fonte | Arial negrito, já definida no arquivo |
| Alinhamento | Direita |
| Cor do texto | RGB 255, 255, 255 |
| Cor de fundo | RGB 0, 51, 90 (ou transparente), sem borda |
| Formato | dd/mm/aaaa hh:mm:ss |
| Regra de cor | Nenhuma |

### Objeto 2: Média prensa automática: valor

| Propriedade | Valor |
|---|---|
| Ferramenta | Campo de dados |
| Posição em unidades (X ; Y) | 1,78 ; 5,47 |
| Tamanho em unidades (largura ; altura) | 6,41 ; 2,63 |
| Tag | `307-GI-001` |
| Mapa | `IP_ANALOGMAP` |
| Atributo | `VAL` |
| Fonte | Arial negrito, já definida no arquivo |
| Alinhamento | Direita |
| Cor do texto | RGB 255, 255, 255 |
| Cor de fundo | RGB 0, 51, 90 (ou transparente), sem borda |
| Formato | 1 casa decimal (0,0) |
| Regra de cor | Nenhuma |

### Objeto 3: Média prensa automática: minitendência

| Propriedade | Valor |
|---|---|
| Ferramenta | Tendência (plot) |
| Posição em unidades (X ; Y) | 11,66 ; 5,97 |
| Tamanho em unidades (largura ; altura) | 4,38 ; 2,81 |
| Observação | Já vem pronta no arquivo, escala 0 a 500. Ver Parte C. |

### Objeto 4: Média prensa automática: atualizado às

| Propriedade | Valor |
|---|---|
| Ferramenta | Campo de dados |
| Posição em unidades (X ; Y) | 4,66 ; 9,69 |
| Tamanho em unidades (largura ; altura) | 5,16 ; 0,69 |
| Tag | `307-GI-001` |
| Mapa | `IP_ANALOGMAP` |
| Atributo | `VAL` no arquivo: **trocar pelo atributo de data e hora** (Parte B) |
| Fonte | Arial, já definida no arquivo |
| Alinhamento | Esquerda |
| Cor do texto | RGB 255, 255, 255 |
| Cor de fundo | RGB 0, 51, 90 (ou transparente), sem borda |
| Formato | dd/mm/aaaa hh:mm:ss |
| Regra de cor | Nenhuma |

### Objeto 5: Média prensa automática: qualidade

| Propriedade | Valor |
|---|---|
| Ferramenta | Campo de dados |
| Posição em unidades (X ; Y) | 12,75 ; 9,69 |
| Tamanho em unidades (largura ; altura) | 2,50 ; 0,69 |
| Tag | `307-GI-001` |
| Mapa | `IP_ANALOGMAP` |
| Atributo | `VAL` no arquivo: **trocar pelo atributo de qualidade** (Parte B) |
| Fonte | Arial negrito, já definida no arquivo |
| Alinhamento | Esquerda |
| Cor do texto | RGB 255, 255, 255 |
| Cor de fundo | RGB 0, 51, 90 (ou transparente), sem borda |
| Formato | Texto |
| Regra de cor | Texto igual a `Good`: RGB 255, 255, 255. Qualquer outro texto: RGB 243, 112, 33 |

### Objeto 6: Desvio padrão: valor

| Propriedade | Valor |
|---|---|
| Ferramenta | Campo de dados |
| Posição em unidades (X ; Y) | 18,22 ; 5,47 |
| Tamanho em unidades (largura ; altura) | 6,41 ; 2,63 |
| Tag | `307-GI-002` |
| Mapa | `IP_ANALOGMAP` |
| Atributo | `VAL` |
| Fonte | Arial negrito, já definida no arquivo |
| Alinhamento | Direita |
| Cor do texto | RGB 255, 255, 255 |
| Cor de fundo | RGB 0, 51, 90 (ou transparente), sem borda |
| Formato | 1 casa decimal (0,0) |
| Regra de cor | Nenhuma |

### Objeto 7: Desvio padrão: minitendência

| Propriedade | Valor |
|---|---|
| Ferramenta | Tendência (plot) |
| Posição em unidades (X ; Y) | 28,09 ; 5,97 |
| Tamanho em unidades (largura ; altura) | 4,38 ; 2,81 |
| Observação | Já vem pronta no arquivo, escala 0 a 150. Ver Parte C. |

### Objeto 8: Desvio padrão: atualizado às

| Propriedade | Valor |
|---|---|
| Ferramenta | Campo de dados |
| Posição em unidades (X ; Y) | 21,10 ; 9,69 |
| Tamanho em unidades (largura ; altura) | 5,16 ; 0,69 |
| Tag | `307-GI-002` |
| Mapa | `IP_ANALOGMAP` |
| Atributo | `VAL` no arquivo: **trocar pelo atributo de data e hora** (Parte B) |
| Fonte | Arial, já definida no arquivo |
| Alinhamento | Esquerda |
| Cor do texto | RGB 255, 255, 255 |
| Cor de fundo | RGB 0, 51, 90 (ou transparente), sem borda |
| Formato | dd/mm/aaaa hh:mm:ss |
| Regra de cor | Nenhuma |

### Objeto 9: Desvio padrão: qualidade

| Propriedade | Valor |
|---|---|
| Ferramenta | Campo de dados |
| Posição em unidades (X ; Y) | 29,18 ; 9,69 |
| Tamanho em unidades (largura ; altura) | 2,50 ; 0,69 |
| Tag | `307-GI-002` |
| Mapa | `IP_ANALOGMAP` |
| Atributo | `VAL` no arquivo: **trocar pelo atributo de qualidade** (Parte B) |
| Fonte | Arial negrito, já definida no arquivo |
| Alinhamento | Esquerda |
| Cor do texto | RGB 255, 255, 255 |
| Cor de fundo | RGB 0, 51, 90 (ou transparente), sem borda |
| Formato | Texto |
| Regra de cor | Texto igual a `Good`: RGB 255, 255, 255. Qualquer outro texto: RGB 243, 112, 33 |

### Objeto 10: Composto 2 h: prensa automática

| Propriedade | Valor |
|---|---|
| Ferramenta | Campo de dados |
| Posição em unidades (X ; Y) | 34,66 ; 5,97 |
| Tamanho em unidades (largura ; altura) | 5,63 ; 2,19 |
| Tag | `307-GI-001-C2H` |
| Mapa | `IP_ANALOGMAP` |
| Atributo | `VAL` |
| Fonte | Arial negrito, já definida no arquivo |
| Alinhamento | Esquerda |
| Cor do texto | RGB 89, 198, 242 |
| Cor de fundo | RGB 0, 51, 90 (ou transparente), sem borda |
| Formato | 1 casa decimal (0,0) |
| Regra de cor | Nenhuma |
| Observação | Média dos cortes com resultado na janela (1 a 4). Tag calculada a criar pelo administrador do PIMS (Parte F); até lá o campo fica sem valor. |

### Objeto 11: Composto 2 h: laboratório (MES)

| Propriedade | Valor |
|---|---|
| Ferramenta | Campo de dados |
| Posição em unidades (X ; Y) | 43,71 ; 5,97 |
| Tamanho em unidades (largura ; altura) | 4,06 ; 2,19 |
| Tag | `M4710050020-0031-HHLFU` |
| Mapa | `IP_MESVALOR` |
| Atributo | `VAL` |
| Fonte | Arial negrito, já definida no arquivo |
| Alinhamento | Esquerda |
| Cor do texto | RGB 255, 255, 255 |
| Cor de fundo | RGB 0, 51, 90 (ou transparente), sem borda |
| Formato | Sem casas decimais (0) |
| Regra de cor | Nenhuma |
| Observação | Mesmo vínculo do campo Compressão da tela PQ US04. |

### Objeto 12: Composto 2 h: fim da janela

| Propriedade | Valor |
|---|---|
| Ferramenta | Campo de dados |
| Posição em unidades (X ; Y) | 37,53 ; 9,06 |
| Tamanho em unidades (largura ; altura) | 3,94 ; 0,69 |
| Tag | `307-GI-001-C2H` |
| Mapa | `IP_ANALOGMAP` |
| Atributo | `VAL` no arquivo: **trocar pelo atributo de data e hora** (Parte B) |
| Fonte | Arial, já definida no arquivo |
| Alinhamento | Esquerda |
| Cor do texto | RGB 255, 255, 255 |
| Cor de fundo | RGB 0, 51, 90 (ou transparente), sem borda |
| Formato | dd/mm/aaaa hh:mm |
| Regra de cor | Nenhuma |
| Observação | Tag calculada a criar pelo administrador do PIMS (Parte F); até lá o campo fica sem valor. |

### Objeto 13: Composto 2 h: diferença prensa menos laboratório

| Propriedade | Valor |
|---|---|
| Ferramenta | Campo de dados |
| Posição em unidades (X ; Y) | 44,03 ; 9,06 |
| Tamanho em unidades (largura ; altura) | 1,75 ; 0,69 |
| Tag | `307-GI-001-DIF2H` |
| Mapa | `IP_ANALOGMAP` |
| Atributo | `VAL` |
| Fonte | Arial negrito, já definida no arquivo |
| Alinhamento | Direita |
| Cor do texto | RGB 255, 255, 255 |
| Cor de fundo | RGB 0, 51, 90 (ou transparente), sem borda |
| Formato | 1 casa decimal (0,0) |
| Regra de cor | Nenhuma |
| Observação | Tag calculada a criar pelo administrador do PIMS (Parte F); até lá o campo fica sem valor. |

### Objeto 14: Tendência: composto da prensa x composto do laboratório

| Propriedade | Valor |
|---|---|
| Ferramenta | Tendência (plot) |
| Posição em unidades (X ; Y) | 3,16 ; 15,06 |
| Tamanho em unidades (largura ; altura) | 30,94 ; 17,50 |
| Observação | Já vem pronta no arquivo, com as 3 penas e escala 50 a 450. Ver Parte C. |

### Objeto 15: Cortes do composto: hora do corte 1

| Propriedade | Valor |
|---|---|
| Ferramenta | Campo de dados |
| Posição em unidades (X ; Y) | 37,72 ; 15,80 |
| Tamanho em unidades (largura ; altura) | 2,50 ; 0,75 |
| Tag | `307-GI-001-CORTE1` |
| Mapa | `IP_ANALOGMAP` |
| Atributo | `VAL` no arquivo: **trocar pelo atributo de data e hora** (Parte B) |
| Fonte | Arial negrito, já definida no arquivo |
| Alinhamento | Esquerda |
| Cor do texto | RGB 255, 255, 255 |
| Cor de fundo | RGB 0, 51, 90 (ou transparente), sem borda |
| Formato | hh:mm |
| Regra de cor | Nenhuma |
| Observação | Tag calculada a criar pelo administrador do PIMS (Parte F); até lá o campo fica sem valor. |

### Objeto 16: Cortes do composto: hora do corte 2

| Propriedade | Valor |
|---|---|
| Ferramenta | Campo de dados |
| Posição em unidades (X ; Y) | 37,72 ; 17,05 |
| Tamanho em unidades (largura ; altura) | 2,50 ; 0,75 |
| Tag | `307-GI-001-CORTE2` |
| Mapa | `IP_ANALOGMAP` |
| Atributo | `VAL` no arquivo: **trocar pelo atributo de data e hora** (Parte B) |
| Fonte | Arial negrito, já definida no arquivo |
| Alinhamento | Esquerda |
| Cor do texto | RGB 255, 255, 255 |
| Cor de fundo | RGB 0, 51, 90 (ou transparente), sem borda |
| Formato | hh:mm |
| Regra de cor | Nenhuma |
| Observação | Tag calculada a criar pelo administrador do PIMS (Parte F); até lá o campo fica sem valor. |

### Objeto 17: Cortes do composto: hora do corte 3

| Propriedade | Valor |
|---|---|
| Ferramenta | Campo de dados |
| Posição em unidades (X ; Y) | 37,72 ; 18,30 |
| Tamanho em unidades (largura ; altura) | 2,50 ; 0,75 |
| Tag | `307-GI-001-CORTE3` |
| Mapa | `IP_ANALOGMAP` |
| Atributo | `VAL` no arquivo: **trocar pelo atributo de data e hora** (Parte B) |
| Fonte | Arial negrito, já definida no arquivo |
| Alinhamento | Esquerda |
| Cor do texto | RGB 255, 255, 255 |
| Cor de fundo | RGB 0, 51, 90 (ou transparente), sem borda |
| Formato | hh:mm |
| Regra de cor | Nenhuma |
| Observação | Tag calculada a criar pelo administrador do PIMS (Parte F); até lá o campo fica sem valor. |

### Objeto 18: Cortes do composto: hora do corte 4

| Propriedade | Valor |
|---|---|
| Ferramenta | Campo de dados |
| Posição em unidades (X ; Y) | 37,72 ; 19,55 |
| Tamanho em unidades (largura ; altura) | 2,50 ; 0,75 |
| Tag | `307-GI-001-CORTE4` |
| Mapa | `IP_ANALOGMAP` |
| Atributo | `VAL` no arquivo: **trocar pelo atributo de data e hora** (Parte B) |
| Fonte | Arial negrito, já definida no arquivo |
| Alinhamento | Esquerda |
| Cor do texto | RGB 255, 255, 255 |
| Cor de fundo | RGB 0, 51, 90 (ou transparente), sem borda |
| Formato | hh:mm |
| Regra de cor | Nenhuma |
| Observação | Tag calculada a criar pelo administrador do PIMS (Parte F); até lá o campo fica sem valor. |

### Objeto 19: Cortes do composto: resultado do corte 1

| Propriedade | Valor |
|---|---|
| Ferramenta | Campo de dados |
| Posição em unidades (X ; Y) | 42,09 ; 15,77 |
| Tamanho em unidades (largura ; altura) | 3,75 ; 0,81 |
| Tag | `307-GI-001-CORTE1` |
| Mapa | `IP_ANALOGMAP` |
| Atributo | `VAL` |
| Fonte | Arial negrito, já definida no arquivo |
| Alinhamento | Direita |
| Cor do texto | RGB 255, 255, 255 |
| Cor de fundo | RGB 0, 51, 90 (ou transparente), sem borda |
| Formato | 1 casa decimal (0,0) |
| Regra de cor | Nenhuma |
| Observação | Tag calculada a criar pelo administrador do PIMS (Parte F); até lá o campo fica sem valor. |

### Objeto 20: Cortes do composto: resultado do corte 2

| Propriedade | Valor |
|---|---|
| Ferramenta | Campo de dados |
| Posição em unidades (X ; Y) | 42,09 ; 17,02 |
| Tamanho em unidades (largura ; altura) | 3,75 ; 0,81 |
| Tag | `307-GI-001-CORTE2` |
| Mapa | `IP_ANALOGMAP` |
| Atributo | `VAL` |
| Fonte | Arial negrito, já definida no arquivo |
| Alinhamento | Direita |
| Cor do texto | RGB 255, 255, 255 |
| Cor de fundo | RGB 0, 51, 90 (ou transparente), sem borda |
| Formato | 1 casa decimal (0,0) |
| Regra de cor | Nenhuma |
| Observação | Tag calculada a criar pelo administrador do PIMS (Parte F); até lá o campo fica sem valor. |

### Objeto 21: Cortes do composto: resultado do corte 3

| Propriedade | Valor |
|---|---|
| Ferramenta | Campo de dados |
| Posição em unidades (X ; Y) | 42,09 ; 18,27 |
| Tamanho em unidades (largura ; altura) | 3,75 ; 0,81 |
| Tag | `307-GI-001-CORTE3` |
| Mapa | `IP_ANALOGMAP` |
| Atributo | `VAL` |
| Fonte | Arial negrito, já definida no arquivo |
| Alinhamento | Direita |
| Cor do texto | RGB 255, 255, 255 |
| Cor de fundo | RGB 0, 51, 90 (ou transparente), sem borda |
| Formato | 1 casa decimal (0,0) |
| Regra de cor | Nenhuma |
| Observação | Tag calculada a criar pelo administrador do PIMS (Parte F); até lá o campo fica sem valor. |

### Objeto 22: Cortes do composto: resultado do corte 4

| Propriedade | Valor |
|---|---|
| Ferramenta | Campo de dados |
| Posição em unidades (X ; Y) | 42,09 ; 19,52 |
| Tamanho em unidades (largura ; altura) | 3,75 ; 0,81 |
| Tag | `307-GI-001-CORTE4` |
| Mapa | `IP_ANALOGMAP` |
| Atributo | `VAL` |
| Fonte | Arial negrito, já definida no arquivo |
| Alinhamento | Direita |
| Cor do texto | RGB 255, 255, 255 |
| Cor de fundo | RGB 0, 51, 90 (ou transparente), sem borda |
| Formato | 1 casa decimal (0,0) |
| Regra de cor | Nenhuma |
| Observação | Tag calculada a criar pelo administrador do PIMS (Parte F); até lá o campo fica sem valor. |

### Objeto 23: Faixa abaixo de 200 kgf/pelota

| Propriedade | Valor |
|---|---|
| Ferramenta | Campo de dados |
| Posição em unidades (X ; Y) | 46,03 ; 24,52 |
| Tamanho em unidades (largura ; altura) | 2,00 ; 0,94 |
| Tag | `307-QI-003` |
| Mapa | `IP_DISCRETEMAP` no arquivo: **conferir** (Parte B) |
| Atributo | `VAL` |
| Fonte | Arial negrito, já definida no arquivo |
| Alinhamento | Direita |
| Cor do texto | RGB 255, 255, 255 |
| Cor de fundo | RGB 0, 51, 90 (ou transparente), sem borda |
| Formato | Sem casas decimais (0) |
| Regra de cor | Nenhuma |
| Observação | A lâmpada ao lado acende na cor da faixa quando o valor é maior que zero (já vem pronta no arquivo). |

### Objeto 24: Faixa abaixo de 150 kgf/pelota

| Propriedade | Valor |
|---|---|
| Ferramenta | Campo de dados |
| Posição em unidades (X ; Y) | 46,03 ; 25,89 |
| Tamanho em unidades (largura ; altura) | 2,00 ; 0,94 |
| Tag | `307-QI-002` |
| Mapa | `IP_DISCRETEMAP` no arquivo: **conferir** (Parte B) |
| Atributo | `VAL` |
| Fonte | Arial negrito, já definida no arquivo |
| Alinhamento | Direita |
| Cor do texto | RGB 255, 255, 255 |
| Cor de fundo | RGB 0, 51, 90 (ou transparente), sem borda |
| Formato | Sem casas decimais (0) |
| Regra de cor | Nenhuma |
| Observação | A lâmpada ao lado acende na cor da faixa quando o valor é maior que zero (já vem pronta no arquivo). |

### Objeto 25: Faixa abaixo de 100 kgf/pelota

| Propriedade | Valor |
|---|---|
| Ferramenta | Campo de dados |
| Posição em unidades (X ; Y) | 46,03 ; 27,27 |
| Tamanho em unidades (largura ; altura) | 2,00 ; 0,94 |
| Tag | `307-QI-001` |
| Mapa | `IP_DISCRETEMAP` no arquivo: **conferir** (Parte B) |
| Atributo | `VAL` |
| Fonte | Arial negrito, já definida no arquivo |
| Alinhamento | Direita |
| Cor do texto | RGB 255, 255, 255 |
| Cor de fundo | RGB 0, 51, 90 (ou transparente), sem borda |
| Formato | Sem casas decimais (0) |
| Regra de cor | Nenhuma |
| Observação | A lâmpada ao lado acende na cor da faixa quando o valor é maior que zero (já vem pronta no arquivo). |

### Objeto 26: Composto 2 h: cortes com resultado na janela

| Propriedade | Valor |
|---|---|
| Ferramenta | Campo de dados |
| Posição em unidades (X ; Y) | 40,10 ; 10,00 |
| Tamanho em unidades (largura ; altura) | 0,75 ; 0,69 |
| Tag | `307-GI-001-N2H` |
| Mapa | `IP_ANALOGMAP` |
| Atributo | `VAL` |
| Fonte | Arial negrito, já definida no arquivo |
| Alinhamento | Direita |
| Cor do texto | RGB 255, 255, 255 |
| Cor de fundo | RGB 0, 51, 90 (ou transparente), sem borda |
| Formato | Sem casas decimais (0) |
| Regra de cor | Nenhuma |
| Observação | Quantidade de cortes (1 a 4) usados na média. Tag calculada a criar pelo administrador do PIMS (Parte F); até lá o campo fica sem valor. |

