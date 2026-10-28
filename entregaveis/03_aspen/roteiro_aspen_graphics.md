# Passo a passo: tela de compressão de pelotas no Aspen

**Tela:** Monitoramento do ensaio de compressão de pelotas (versão 4: arquivo `.atgraphic` pronto, sem imagem de fundo)

**Resultado esperado:** igual a `tela_exemplo_normal.png`, nítida também com a janela maximizada

**Tempo estimado:** 30 a 60 minutos

![Resultado esperado](tela_exemplo_normal.png)

## O que mudou nesta versão

A versão anterior usava uma imagem de fundo com os textos. Quando a janela é maximizada, o Aspen estica a imagem e as letras perdem a nitidez. A tela PQ US04 não tem esse problema porque todos os textos e quadros são **objetos do próprio Aspen**, redesenhados em qualquer tamanho.

Agora a tela segue o mesmo princípio e já vem **pronta, em um arquivo `.atgraphic`**, gerado no mesmo formato da tela PQ US04:

| Conteúdo do arquivo | Quantidade |
|---|---|
| Painéis, linhas, grade da tendência e linhas de referência 100, 150 e 200 (polígonos) | 30 |
| Textos fixos (títulos, rótulos, unidades) | 51 |
| Campos de dados já ligados às tags | 14 |
| Tendências já ligadas às tags, com escala fixa | 3 |

Tamanho da tela: **50 x 35 unidades**, o padrão do editor e o mesmo da tela PQ US04. Não é preciso digitar posição nenhuma.

### Arquivos

| Arquivo | Para que serve |
|---|---|
| `LCE_Compressao_Pelotas_Prensa01.atgraphic` | **A tela.** É o arquivo que você vai abrir |
| `teste_minimo.atgraphic` | Arquivo de diagnóstico com 4 objetos (um painel, um texto, um campo e uma tendência). Só use se a tela não abrir |
| `tela_exemplo_*.png` | Como a tela deve ficar: normal, faixa ativa, qualidade Bad e cortador em manual |
| `mapa_posicionamento.png` | Número de cada objeto dinâmico, usado nas fichas do Anexo |
| `posicoes_objetos.xlsx` | Posição e tamanho de cada objeto dinâmico, em unidades |

> **Importante:** o arquivo foi gerado por engenharia reversa do formato da tela PQ US04 e conferido por releitura completa, mas **não pôde ser aberto no Aspen** durante a preparação, porque o ambiente de desenvolvimento não tem acesso ao PIMS. A Parte A diz o que fazer se ele não abrir.

## Parte A. Abrir a tela

**A1.** Copie `LCE_Compressao_Pelotas_Prensa01.atgraphic` para a pasta de rede onde ficam as telas do laboratório.

**A2.** Abra o arquivo no **Aspen Process Graphics Editor** (para editar) ou no **Aspen Process Explorer** (para usar).

**A3.** Compare com `tela_exemplo_normal.png` e **maximize a janela**. Os textos devem continuar nítidos, como na tela PQ US04.

**Se der erro ao abrir:**

1. Tente abrir `teste_minimo.atgraphic`.
2. Me envie uma captura da mensagem de erro e diga se o arquivo de teste abriu ou não. Com isso eu localizo o ponto do formato a corrigir.

## Parte B. Ajustar os campos de dados (no editor)

Os 14 campos já vêm com tag, mapa, fonte, cor e posição. Faltam só os itens que dependem de nomes que só existem no seu PIMS e as regras de cor. Para abrir as propriedades de um campo, dê duplo clique nele.

**B1. Conheça o diálogo pela tela PQ US04.** Abra a tela PQ US04 no editor, dê duplo clique no valor de **Compressão** e veja onde ficam: tag, mapa, atributo, formato e cores condicionais. Abra a lista de **atributos** (onde aparece `VAL`) e anote:

| O que | Nome exato na lista (anote aqui) |
|---|---|
| Data e hora do valor | ______________________ |
| Qualidade do valor | ______________________ |

Feche a tela PQ US04 **sem salvar**.

**B2. Troque o atributo** nos campos abaixo. No arquivo eles vêm com `VAL`, que mostra o valor e não a hora ou a qualidade:

| Objeto | Campo | Trocar `VAL` por |
|---|---|---|
| 2 | Última atualização (cabeçalho) | atributo de data e hora |
| 5 | Carga média: atualizado em | atributo de data e hora |
| 9 | Desvio padrão: atualizado em | atributo de data e hora |
| 12 | Compressão MES: horário do resultado | atributo de data e hora |
| 6 | Carga média: qualidade | atributo de qualidade |
| 10 | Desvio padrão: qualidade | atributo de qualidade |

**B3. Confira o mapa dos discretos** (objetos 16, 17 e 18, faixas de classificação). No arquivo eles vêm com o mapa `IP_DISCRETEMAP`. Se o campo mostrar erro ou ficar vazio, abra a lista de mapas e escolha o mapa de discreto disponível.

**B4. Configure as regras de cor** (aba de cores condicionais):

| Objetos | Condição | Cor do texto |
|---|---|---|
| 6 e 10 (qualidade) | igual a `Good` | RGB 230, 234, 240 |
| | qualquer outro valor | RGB 242, 153, 74 (laranja) |
| 13 e 14 (modo dos cortadores) | `Automático` | RGB 230, 234, 240 |
| | `Manual` | RGB 242, 153, 74 (laranja) |
| 16 (menor que 200) | `1` ou `Ativo` | RGB 242, 201, 76 (amarelo) |
| 17 (menor que 150) | `1` ou `Ativo` | RGB 242, 153, 74 (laranja) |
| 18 (menor que 100) | `1` ou `Ativo` | RGB 235, 87, 87 (vermelho) |
| 16, 17 e 18 | `0` ou `Inativo` | RGB 139, 150, 165 (cinza) |

**B5. Confira as casas decimais:** carga média (3) e desvio padrão (7) com 1 casa; compressão MES (11) sem casas. Se o editor usar o formato da própria tag, deixe como estiver.

**B6. Alinhamento:** os valores grandes (3, 7 e 11) ficam melhor **alinhados à direita**, encostados na unidade "daN/pelota". Se o editor tiver a opção, use alinhamento à direita também em 2, 12, 13, 14, 16, 17 e 18.

## Parte C. Conferir as tendências

As três tendências já vêm prontas:

| Objeto | Penas | Escala |
|---|---|---|
| 4 (minitendência da carga média) | `307-GI-001` (azul) | 0 a 500 |
| 8 (minitendência do desvio padrão) | `307-GI-002` (azul) | 0 a 150 |
| 15 (tendência principal) | `307-GI-001` (azul) e `M4710050020-0031-HHLFU`, mapa `IP_MESVALOR` (cinza) | 0 a 500 nas duas |

**C1.** Na barra de tempo, embaixo da janela, troque o período de `2:00:00` para **`8:00:00`**.

**C2.** A escala de 0 a 500, a grade e as linhas de 100, 150 e 200 da tendência principal são **parte fixa da tela**, desenhadas por cima da tendência. Por isso **não altere a escala das penas**: se ela mudar, os números da lateral deixam de corresponder às linhas.

**C3.** Se o fundo da tendência aparecer preto em vez de cinza-azulado, ajuste a cor de fundo nas propriedades da tendência para RGB 31, 38, 48.

## Parte D. Logotipo (objeto 1)

Com a ferramenta **Imagem**, insira o arquivo **oficial** do logotipo fornecido pela Comunicação no canto superior esquerdo (objeto 1 do mapa), mantendo a proporção. Não redesenhe a marca. O logotipo é a única imagem da tela; se ele perder um pouco de nitidez ao maximizar, use um arquivo de maior resolução.

## Parte E. Salvar e testar

**E1.** No editor, **Arquivo > Salvar como**, na pasta de rede, com o nome `LCE_Compressao_Pelotas_Prensa01`.

**E2.** Abra no Process Explorer, maximize e confira a Parte F.

> A janela mostra "(Somente leitura)" quando a tela é aberta para uso. Para editar, abra pelo Graphics Editor.

## Parte F. Conferência final

| Nº | Verificar | Esperado |
|---|---|---|
| 1 | Maximizar a janela | Textos e quadros nítidos, sem serrilhado |
| 2 | Carga média e desvio padrão | 1 casa decimal com vírgula, por exemplo `290,1` |
| 3 | Horários | Data e hora do último ensaio, não o valor |
| 4 | Qualidade | `Good` em branco; outro texto em laranja |
| 5 | Faixas | `Inativo` (ou 0) em cinza; `Ativo` (ou 1) em amarelo, laranja ou vermelho |
| 6 | Compressão MES e modo dos cortadores | Mesmos valores da tela PQ US04 no mesmo horário |
| 7 | Tendência | Período de 8 h, linha azul do módulo e linha cinza do laboratório entre 0 e 500 |

## Parte G. Pedido ao administrador do PIMS (opcional)

Texto pronto para enviar:

> Prezados, para a tela de monitoramento do ensaio de compressão de pelotas (LCE), solicito um **descritor de estados** para os discretos `307-QI-001`, `307-QI-002` e `307-QI-003`: 0 = Inativo, 1 = Ativo, para que a tela mostre o texto em vez do número.
>
> Se possível, também duas **tags calculadas** para a banda da tendência: `307-GI-001-SUP` = `307-GI-001` + `307-GI-002` e `307-GI-001-INF` = `307-GI-001` menos `307-GI-002`. Os nomes podem seguir o padrão da área.
>
> Obrigado.

Quando as tags calculadas existirem, adicione-as como penas na tendência 15, com escala de 0 a 500 e cor RGB 92, 200, 240.

## Parte H. Problemas comuns

| Problema | Causa provável | Solução |
|---|---|---|
| Erro ao abrir o arquivo | Diferença de formato não prevista | Tente `teste_minimo.atgraphic` e envie a captura do erro (Parte A) |
| Campo mostra erro ou "?" | Mapa ou atributo inexistente no PIMS | Partes B2 e B3 |
| Horário mostra um número | Atributo ainda em `VAL` | Parte B2 |
| Tendência com escala diferente | Escala da pena alterada | Volte para 0 a 500 (ou 0 a 150 no objeto 8) |
| Número com ponto em vez de vírgula | Configuração regional da estação | Verifique em Ferramentas > Opções; se não houver ajuste, peça à TI |
| Faixa mostra 0 ou 1 | Descritor de estados não configurado | Normal até o pedido da Parte G ser atendido |
| Texto cortado ou sobrando | Fonte da estação diferente da prevista | Aumente um pouco a largura do objeto de texto no editor |

## Limitações

- O arquivo não pôde ser testado no Aspen durante a preparação. Se algo aparecer diferente, envie uma captura para ajuste.
- A tendência do Aspen desenha só as linhas. Os horários do eixo de tempo aparecem na barra de tempo, embaixo da janela, como na tela PQ US04.
- A severidade das faixas (Atenção, Alerta e Crítico) está **proposta** e aguarda aprovação do LCE.

## Anexo. Fichas dos objetos dinâmicos

Referência caso algum objeto precise ser refeito à mão. As posições estão em unidades da tela de 50 x 35.

![Mapa de posicionamento](mapa_posicionamento.png)

### Objeto 1: Logotipo oficial

| Propriedade | Valor |
|---|---|
| Ferramenta | Imagem |
| Posição em unidades (X ; Y) | 1,00 ; 0,48 |
| Tamanho em unidades (largura ; altura) | 6,25 ; 1,75 |
| Observação | Arquivo oficial fornecido pela Comunicação. Manter a proporção dentro do retângulo. Ver Parte D. |

### Objeto 2: Última atualização (cabeçalho)

| Propriedade | Valor |
|---|---|
| Ferramenta | Campo de dados |
| Posição em unidades (X ; Y) | 42,75 ; 1,31 |
| Tamanho em unidades (largura ; altura) | 6,25 ; 0,81 |
| Tag | `307-GI-001` |
| Mapa | `IP_ANALOGMAP` |
| Atributo | `VAL` no arquivo: **trocar pelo atributo de data e hora** (Parte B) |
| Fonte | Arial negrito, já definida no arquivo |
| Alinhamento | Direita |
| Cor do texto | RGB 230, 234, 240 |
| Cor de fundo | RGB 17, 22, 29 (ou transparente), sem borda |
| Formato | dd/mm/aaaa hh:mm:ss |
| Regra de cor | Nenhuma |
| Observação | Fica no cabeçalho, que tem fundo mais escuro que os painéis. |

### Objeto 3: Carga média: valor

| Propriedade | Valor |
|---|---|
| Ferramenta | Campo de dados |
| Posição em unidades (X ; Y) | 1,78 ; 5,25 |
| Tamanho em unidades (largura ; altura) | 6,41 ; 2,63 |
| Tag | `307-GI-001` |
| Mapa | `IP_ANALOGMAP` |
| Atributo | `VAL` |
| Fonte | Arial negrito, já definida no arquivo |
| Alinhamento | Direita |
| Cor do texto | RGB 230, 234, 240 |
| Cor de fundo | RGB 31, 38, 48 (ou transparente), sem borda |
| Formato | 1 casa decimal (0,0) |
| Regra de cor | Nenhuma |
| Observação | Objeto de calibração da fonte (ver início da Parte E). A unidade daN/pelota já está no fundo. |

### Objeto 4: Carga média: minitendência

| Propriedade | Valor |
|---|---|
| Ferramenta | Tendência (plot) |
| Posição em unidades (X ; Y) | 11,78 ; 5,44 |
| Tamanho em unidades (largura ; altura) | 4,69 ; 3,13 |
| Observação | Já vem pronta no arquivo. Ver Parte C. |

### Objeto 5: Carga média: atualizado em

| Propriedade | Valor |
|---|---|
| Ferramenta | Campo de dados |
| Posição em unidades (X ; Y) | 4,80 ; 9,53 |
| Tamanho em unidades (largura ; altura) | 5,16 ; 0,69 |
| Tag | `307-GI-001` |
| Mapa | `IP_ANALOGMAP` |
| Atributo | `VAL` no arquivo: **trocar pelo atributo de data e hora** (Parte B) |
| Fonte | Arial, já definida no arquivo |
| Alinhamento | Esquerda |
| Cor do texto | RGB 230, 234, 240 |
| Cor de fundo | RGB 31, 38, 48 (ou transparente), sem borda |
| Formato | dd/mm/aaaa hh:mm:ss |
| Regra de cor | Nenhuma |

### Objeto 6: Carga média: qualidade

| Propriedade | Valor |
|---|---|
| Ferramenta | Campo de dados |
| Posição em unidades (X ; Y) | 12,83 ; 9,53 |
| Tamanho em unidades (largura ; altura) | 2,81 ; 0,69 |
| Tag | `307-GI-001` |
| Mapa | `IP_ANALOGMAP` |
| Atributo | `VAL` no arquivo: **trocar pelo atributo de qualidade** (Parte B) |
| Fonte | Arial negrito, já definida no arquivo |
| Alinhamento | Esquerda |
| Cor do texto | RGB 230, 234, 240 |
| Cor de fundo | RGB 31, 38, 48 (ou transparente), sem borda |
| Formato | Texto |
| Regra de cor | Texto igual a `Good`: RGB 230, 234, 240. Qualquer outro texto: RGB 242, 153, 74 |

### Objeto 7: Desvio padrão: valor

| Propriedade | Valor |
|---|---|
| Ferramenta | Campo de dados |
| Posição em unidades (X ; Y) | 18,66 ; 5,25 |
| Tamanho em unidades (largura ; altura) | 6,41 ; 2,63 |
| Tag | `307-GI-002` |
| Mapa | `IP_ANALOGMAP` |
| Atributo | `VAL` |
| Fonte | Arial negrito, já definida no arquivo |
| Alinhamento | Direita |
| Cor do texto | RGB 230, 234, 240 |
| Cor de fundo | RGB 31, 38, 48 (ou transparente), sem borda |
| Formato | 1 casa decimal (0,0) |
| Regra de cor | Nenhuma |
| Observação | Copie o objeto 3 e troque a tag e a posição. |

### Objeto 8: Desvio padrão: minitendência

| Propriedade | Valor |
|---|---|
| Ferramenta | Tendência (plot) |
| Posição em unidades (X ; Y) | 28,66 ; 5,44 |
| Tamanho em unidades (largura ; altura) | 4,69 ; 3,13 |
| Observação | Já vem pronta no arquivo. Ver Parte C. |

### Objeto 9: Desvio padrão: atualizado em

| Propriedade | Valor |
|---|---|
| Ferramenta | Campo de dados |
| Posição em unidades (X ; Y) | 21,68 ; 9,53 |
| Tamanho em unidades (largura ; altura) | 5,16 ; 0,69 |
| Tag | `307-GI-002` |
| Mapa | `IP_ANALOGMAP` |
| Atributo | `VAL` no arquivo: **trocar pelo atributo de data e hora** (Parte B) |
| Fonte | Arial, já definida no arquivo |
| Alinhamento | Esquerda |
| Cor do texto | RGB 230, 234, 240 |
| Cor de fundo | RGB 31, 38, 48 (ou transparente), sem borda |
| Formato | dd/mm/aaaa hh:mm:ss |
| Regra de cor | Nenhuma |
| Observação | Copie o objeto 5 e troque a tag e a posição. |

### Objeto 10: Desvio padrão: qualidade

| Propriedade | Valor |
|---|---|
| Ferramenta | Campo de dados |
| Posição em unidades (X ; Y) | 29,70 ; 9,53 |
| Tamanho em unidades (largura ; altura) | 2,81 ; 0,69 |
| Tag | `307-GI-002` |
| Mapa | `IP_ANALOGMAP` |
| Atributo | `VAL` no arquivo: **trocar pelo atributo de qualidade** (Parte B) |
| Fonte | Arial negrito, já definida no arquivo |
| Alinhamento | Esquerda |
| Cor do texto | RGB 230, 234, 240 |
| Cor de fundo | RGB 31, 38, 48 (ou transparente), sem borda |
| Formato | Texto |
| Regra de cor | Texto igual a `Good`: RGB 230, 234, 240. Qualquer outro texto: RGB 242, 153, 74 |
| Observação | Copie o objeto 6 e troque a tag e a posição. |

### Objeto 11: Compressão MES: último resultado

| Propriedade | Valor |
|---|---|
| Ferramenta | Campo de dados |
| Posição em unidades (X ; Y) | 35,53 ; 5,81 |
| Tamanho em unidades (largura ; altura) | 5,31 ; 2,63 |
| Tag | `M4710050020-0031-HHLFU` |
| Mapa | `IP_MESVALOR` |
| Atributo | `VAL` |
| Fonte | Arial negrito, já definida no arquivo |
| Alinhamento | Direita |
| Cor do texto | RGB 230, 234, 240 |
| Cor de fundo | RGB 31, 38, 48 (ou transparente), sem borda |
| Formato | Sem casas decimais (0) |
| Regra de cor | Nenhuma |
| Observação | Mesmo vínculo do campo Compressão da tela PQ US04. Mesma fonte do objeto 3: copie o 3 e troque tag, mapa e formato. |

### Objeto 12: Compressão MES: horário do resultado

| Propriedade | Valor |
|---|---|
| Ferramenta | Campo de dados |
| Posição em unidades (X ; Y) | 43,53 ; 8,69 |
| Tamanho em unidades (largura ; altura) | 4,69 ; 0,69 |
| Tag | `M4710050020-0031-HHLFU` |
| Mapa | `IP_MESVALOR` |
| Atributo | `VAL` no arquivo: **trocar pelo atributo de data e hora** (Parte B) |
| Fonte | Arial, já definida no arquivo |
| Alinhamento | Esquerda |
| Cor do texto | RGB 230, 234, 240 |
| Cor de fundo | RGB 31, 38, 48 (ou transparente), sem borda |
| Formato | dd/mm/aaaa hh:mm |
| Regra de cor | Nenhuma |
| Observação | Copie o objeto 5 e troque tag, mapa e formato. |

### Objeto 13: Modo do cortador 07CR001

| Propriedade | Valor |
|---|---|
| Ferramenta | Campo de dados |
| Posição em unidades (X ; Y) | 44,78 ; 11,22 |
| Tamanho em unidades (largura ; altura) | 3,44 ; 0,81 |
| Tag | `U04-07CR001` |
| Mapa | `IP_PIDMAP03` |
| Atributo | `VAL` |
| Fonte | Arial negrito, já definida no arquivo |
| Alinhamento | Direita |
| Cor do texto | RGB 230, 234, 240 |
| Cor de fundo | RGB 31, 38, 48 (ou transparente), sem borda |
| Formato | Texto |
| Regra de cor | Texto `Automático`: RGB 230, 234, 240. Texto `Manual`: RGB 242, 153, 74 |
| Observação | Mesmo vínculo do indicador Manual/Automático da tela PQ US04. Se o campo mostrar número, use a regra de cor pelo número equivalente. |

### Objeto 14: Modo do cortador 07CR002

| Propriedade | Valor |
|---|---|
| Ferramenta | Campo de dados |
| Posição em unidades (X ; Y) | 44,78 ; 12,47 |
| Tamanho em unidades (largura ; altura) | 3,44 ; 0,81 |
| Tag | `U04-07CR002` |
| Mapa | `IP_PIDMAP03` |
| Atributo | `VAL` |
| Fonte | Arial negrito, já definida no arquivo |
| Alinhamento | Direita |
| Cor do texto | RGB 230, 234, 240 |
| Cor de fundo | RGB 31, 38, 48 (ou transparente), sem borda |
| Formato | Texto |
| Regra de cor | Texto `Automático`: RGB 230, 234, 240. Texto `Manual`: RGB 242, 153, 74 |
| Observação | Copie o objeto 13 e troque a tag. |

### Objeto 15: Tendência principal da carga média

| Propriedade | Valor |
|---|---|
| Ferramenta | Tendência (plot) |
| Posição em unidades (X ; Y) | 3,03 ; 14,63 |
| Tamanho em unidades (largura ; altura) | 29,63 ; 18,00 |
| Observação | Já vem pronta no arquivo, com as penas e a escala. Ver Parte C. |

### Objeto 16: Faixa menor que 200 (Atenção)

| Propriedade | Valor |
|---|---|
| Ferramenta | Campo de dados |
| Posição em unidades (X ; Y) | 44,31 ; 19,91 |
| Tamanho em unidades (largura ; altura) | 3,25 ; 1,00 |
| Tag | `307-QI-003` |
| Mapa | `IP_DISCRETEMAP` no arquivo: **conferir** (Parte B) |
| Atributo | `VAL` |
| Fonte | Arial negrito, já definida no arquivo |
| Alinhamento | Direita |
| Cor do texto | RGB 230, 234, 240 |
| Cor de fundo | RGB 31, 38, 48 (ou transparente), sem borda |
| Formato | Texto (ou número inteiro) |
| Regra de cor | Valor `1` ou `Ativo`: RGB 242, 201, 76. Valor `0` ou `Inativo`: RGB 139, 150, 165 |

### Objeto 17: Faixa menor que 150 (Alerta)

| Propriedade | Valor |
|---|---|
| Ferramenta | Campo de dados |
| Posição em unidades (X ; Y) | 44,31 ; 24,53 |
| Tamanho em unidades (largura ; altura) | 3,25 ; 1,00 |
| Tag | `307-QI-002` |
| Mapa | `IP_DISCRETEMAP` no arquivo: **conferir** (Parte B) |
| Atributo | `VAL` |
| Fonte | Arial negrito, já definida no arquivo |
| Alinhamento | Direita |
| Cor do texto | RGB 230, 234, 240 |
| Cor de fundo | RGB 31, 38, 48 (ou transparente), sem borda |
| Formato | Texto (ou número inteiro) |
| Regra de cor | Valor `1` ou `Ativo`: RGB 242, 153, 74. Valor `0` ou `Inativo`: RGB 139, 150, 165 |
| Observação | Copie o objeto 16. |

### Objeto 18: Faixa menor que 100 (Crítico)

| Propriedade | Valor |
|---|---|
| Ferramenta | Campo de dados |
| Posição em unidades (X ; Y) | 44,31 ; 29,16 |
| Tamanho em unidades (largura ; altura) | 3,25 ; 1,00 |
| Tag | `307-QI-001` |
| Mapa | `IP_DISCRETEMAP` no arquivo: **conferir** (Parte B) |
| Atributo | `VAL` |
| Fonte | Arial negrito, já definida no arquivo |
| Alinhamento | Direita |
| Cor do texto | RGB 230, 234, 240 |
| Cor de fundo | RGB 31, 38, 48 (ou transparente), sem borda |
| Formato | Texto (ou número inteiro) |
| Regra de cor | Valor `1` ou `Ativo`: RGB 235, 87, 87. Valor `0` ou `Inativo`: RGB 139, 150, 165 |
| Observação | Copie o objeto 16. |

