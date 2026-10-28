# Roteiro de montagem da tela no Aspen Process Graphics Editor V14.2

Tela: Monitoramento do ensaio de compressão de pelotas
Tags: 307-GI-001, 307-GI-002, 307-QI-001, 307-QI-002, 307-QI-003 (coluna "Tag" do chamado RITM0273793)

> **Observação importante.** Este roteiro foi escrito sem acesso ao seu PIMS. A sequência de trabalho e os valores (posições, cores, tags e campos) estão prontos para uso. Os nomes exatos de menus e abas do editor podem mudar um pouco na V14.2. Quando um nome não bater, procure a opção equivalente e me envie uma captura da janela de propriedades para eu ajustar o texto.

## 1. Como a tela foi pensada

As ferramentas nativas do editor (texto, retângulo, campo de dados, barra e tendência) desenham uma tela funcional, mas com aparência antiga. Para chegar a um visual atual, a tela é montada em duas camadas:

1. **Camada de fundo (imagem):** tudo o que não muda: cabeçalho, painéis com cantos arredondados, títulos, rótulos, nomes das tags e a coluna de endereços OPC. Arquivo pronto: `fundo_1920x1080.png`.
2. **Camada dinâmica (objetos do Aspen):** somente o que vem do histórico: valores, horários, qualidades, estados das faixas e tendências. São 44 objetos (37 do módulo e 7 do painel complementar do laboratório), cada um com posição e tamanho definidos em `posicoes_objetos.csv` e numerados na imagem `mapa_posicionamento.png`.

Vantagem: a parte bonita fica pronta numa única imagem, e você só configura os campos de dados. Para outra prensa, basta trocar as tags.

### Arquivos desta pasta

| Arquivo | Uso |
|---|---|
| `fundo_1920x1080.png` | Imagem de fundo definitiva |
| `mapa_posicionamento.png` | Mesmo fundo com retângulos numerados. Use temporariamente como guia de posicionamento |
| `posicoes_objetos.csv` | Número, tipo, tag, campo, posição X e Y, largura e altura de cada objeto, em pixels |
| `tela_exemplo_normal.png` e demais | Como a tela deve ficar em cada situação (dados simulados) |

## 2. Antes de começar (5 minutos)

1. Abra o Aspen Process Explorer (ou o Tag Browser) e pesquise `307-GI-001`. Confirme que a tag existe e anote o tipo de registro (definição), que deve ser algo como `IP_AnalogDef`. Repita para `307-QI-001`, que deve ser `IP_DiscreteDef`.
2. Ainda no Tag Browser, abra os campos da tag e confirme a existência de `IP_INPUT_VALUE` (valor), `IP_INPUT_TIME` (data e hora) e `IP_INPUT_QUALITY` (qualidade). Se os nomes forem outros, anote e me informe.
3. **Padrão de vínculo usado na UBU.** Na tela PIMS PQ US04, os campos de dados apontam para **tag + mapa + atributo `VAL`** (exemplo: `406-FX-001`, mapa `IP_ANALOGMAP`, atributo `VAL`). Siga o mesmo padrão: para `307-GI-001` e `307-GI-002`, use provavelmente o mapa `IP_ANALOGMAP`; para os discretos, o mapa de discretos disponível na lista. Para data e hora e para qualidade, escolha na lista de atributos do mapa os equivalentes a `IP_INPUT_TIME` e `IP_INPUT_QUALITY`. Neste roteiro e no CSV, os nomes `IP_INPUT_*` indicam o conteúdo esperado.
4. Nos discretos, veja como o valor aparece: `0` e `1`, ou textos como `OFF` e `ON`. Isso define o passo 7.
5. Copie os arquivos desta pasta para uma pasta de rede acessível pela estação onde a tela vai rodar.

## 3. Criar o gráfico e definir o tamanho

1. No editor, **Arquivo > Novo**.
2. Abra as propriedades do gráfico (duplo clique numa área vazia da tela cinza, ou menu **Editar > Propriedades** / **Visualização > Propriedades do gráfico**).
3. Defina o tamanho em **1920 x 1080 pixels**.
4. Defina a cor de fundo em **RGB 22, 27, 34** (hexadecimal `#161B22`).
5. Salve com um nome descritivo, por exemplo `LCE_Compressao_Pelotas_Prensa01`.

## 4. Inserir a imagem de fundo

1. Selecione a ferramenta de **imagem/bitmap** na paleta da esquerda (ícone de figura, terceira coluna da segunda linha da paleta).
2. Clique na tela e escolha `mapa_posicionamento.png`. Começar pelo mapa facilita o posicionamento; no fim ele será trocado pelo fundo definitivo.
3. Nas propriedades da imagem, defina **X = 0, Y = 0, largura = 1920, altura = 1080**. Não estique a imagem por arraste, porque isso distorce os textos.
4. Envie a imagem para trás (**Editar > Enviar para trás**, ou botão de ordem na paleta) e, se houver a opção, bloqueie o objeto para não movê-lo por engano.

## 5. Padrão visual dos objetos dinâmicos

Use estes valores em todos os campos de dados, para que eles "desapareçam" dentro dos painéis da imagem:

| Propriedade | Valor |
|---|---|
| Fonte | Arial (a mesma barra de fonte que aparece no editor) |
| Cor do texto normal | RGB 230, 234, 240 (`#E6EAF0`) |
| Cor de fundo do campo | RGB 31, 38, 48 (`#1F2630`), igual ao painel. Se existir "fundo transparente", prefira |
| Borda do campo | Nenhuma |
| Cor para dado não confiável | RGB 107, 116, 128 (`#6B7480`), cinza |
| Atenção | RGB 242, 201, 76 (`#F2C94C`), amarelo |
| Alerta e problema de qualidade | RGB 242, 153, 74 (`#F2994A`), laranja |
| Crítico | RGB 235, 87, 87 (`#EB5757`), vermelho |
| Linha de tendência | RGB 92, 200, 240 (`#5CC8F0`), azul claro |

Regra de ouro (ISA-101): cor só aparece em situação anormal, e sempre acompanhada de texto ou símbolo. Em operação normal, a tela fica em tons neutros.

## 6. Campos de dados (valores, horários e qualidades)

Ferramenta: **campo de dados** (ícone de lápis com campo, na paleta da esquerda).

Para cada linha do `posicoes_objetos.csv` com objeto "Campo de dados":

1. Selecione a ferramenta e desenhe um retângulo aproximado sobre o número correspondente do mapa.
2. Na janela de propriedades que abre, preencha:
   - **Tag:** coluna `tag` do CSV (exemplo: `307-GI-001`).
   - **Campo/atributo:** coluna `atributo` do CSV (`IP_INPUT_VALUE`, `IP_INPUT_TIME` ou `IP_INPUT_QUALITY`).
   - **Posição e tamanho:** colunas `x`, `y`, `largura` e `altura`.
3. Ajuste a fonte conforme a tabela abaixo.

| Objetos (nº do mapa) | Conteúdo | Fonte | Formato |
|---|---|---|---|
| 3 e 8 | Carga média e desvio padrão | Arial negrito 60 a 72 pt, alinhado à direita | 1 casa decimal |
| 2 | Última atualização (cabeçalho) | Arial negrito 14 pt, alinhado à direita | Data e hora `dd/mm/aaaa hh:mm:ss` |
| 6 e 10 | Atualizado em | Arial 11 pt | Data e hora |
| 7, 11, 13 a 17 | Qualidade | Arial 11 pt | Texto |
| 20 a 22 | Estado das faixas | Arial negrito 14 pt, alinhado à direita | Ver passo 7 |
| 23 a 37 | Tabela: valor, timestamp e qualidade | Arial 11 pt | Analógicos com 1 casa decimal |

Dicas:

- **Copiar e colar acelera muito.** Configure o objeto 23 (valor de 307-GI-001 na tabela), copie, cole, mude a tag para 307-GI-002 e o Y para 895. Repita descendo 35 pixels por linha.
- **Separador decimal:** se o campo mostrar ponto em vez de vírgula, o padrão vem da configuração regional da estação ou do servidor. Verifique em **Ferramentas > Opções** do editor; se não houver ajuste, registre como pendência para a TI.
- **Pontos (pt) e pixels:** se o tamanho da fonte do editor for em pontos, 72 pt corresponde a cerca de 96 px em telas comuns. Ajuste visualmente até o número preencher o retângulo do mapa.

## 7. Cores condicionais (qualidade e faixas)

Nas propriedades do campo de dados, procure a aba de **cores dinâmicas**, **animação** ou **condições** (no Aspen ela costuma permitir trocar a cor do texto conforme o valor ou conforme limites).

### 7.1 Campos de qualidade (objetos 7, 11, 13 a 17, 25, 28, 31, 34, 37)

| Condição | Cor do texto |
|---|---|
| Qualidade igual a Good | Normal (`#E6EAF0`) |
| Qualquer outra (Bad, Suspect, sem dado) | Laranja (`#F2994A`) |

### 7.2 Estados das faixas (objetos 20, 21 e 22)

Tag e campo: `307-QI-003`, `307-QI-002` e `307-QI-001`, campo `IP_INPUT_VALUE`.

| Valor | Texto exibido | Cor do texto |
|---|---|---|
| 0 | Inativo | Cinza (`#8B96A5`) |
| 1, objeto 20 (menor que 200) | Ativo | Amarelo (`#F2C94C`) |
| 1, objeto 21 (menor que 150) | Ativo | Laranja (`#F2994A`) |
| 1, objeto 22 (menor que 100) | Ativo | Vermelho (`#EB5757`) |

Se o campo mostrar `0` e `1` em vez de texto, há duas saídas:

- **Melhor:** pedir ao administrador do InfoPlus.21 que associe as três tags a um conjunto de estados (select set) com `0 = Inativo` e `1 = Ativo`. O campo passa a exibir o texto sozinho.
- **Imediata:** deixar o campo mostrando `0` ou `1`, com a regra de cor acima, e escrever ao lado, na própria tela, a legenda "0 = Inativo, 1 = Ativo" com a ferramenta de texto.

### 7.3 Aviso na carga média (objeto 4)

Opcional. Um campo de texto ligado a `307-QI-003` que só aparece quando o valor for 1, com o texto "▲ Atenção: carga média abaixo de 200". Se o editor não oferecer visibilidade condicional, deixe esse objeto de fora; as faixas no painel da direita já cumprem a função.

## 8. Tendências

Ferramenta: **tendência/plot** (ícone de gráfico, última posição da quarta linha da paleta).

### 8.1 Tendência principal (objeto 19)

1. Desenhe o objeto e defina **X = 57, Y = 466, largura = 1290, altura = 266**.
2. Adicione a tag `307-GI-001`.
3. **Escala do eixo vertical:** fixa, de 0 a 500. Evite autoescala, que esconde a referência visual das faixas.
4. **Janela de tempo:** 8 horas. Na captura que você enviou, a barra inferior do editor mostra `2:00:00`; troque para `8:00:00`.
5. **Linha:** cor `#5CC8F0`, espessura 2 ou 3, com marcadores nos pontos (cada ponto é um ensaio).
6. **Fundo do gráfico:** `#1F2630`. **Grade:** `#2E3744`. **Textos dos eixos:** `#8B96A5`.
7. Inclua a compressão do laboratório `M4710050020-0031-HHLFU` (mapa `IP_MESVALOR`) na **mesma escala de 0 a 500**, só com marcadores quadrados cinza (`#AAB3BF`) e sem linha, ou com linha em degrau fina. É informação complementar: deve aparecer menos que a linha do módulo.
8. Opcional: `307-GI-002` em eixo secundário, com linha fina cinza, para acompanhar a dispersão.

### 8.2 Referências em 100, 150 e 200 e banda de mais ou menos um desvio padrão

O plot do Aspen não desenha faixas sombreadas nem linhas fixas sem uma tag por trás. Opções, em ordem de preferência:

1. **Limites da própria tag:** se a tendência tiver a opção de mostrar limites (por exemplo, limite baixo e baixo-baixo), peça ao administrador do PIMS para preencher os limites de `307-GI-001` com 200, 150 e 100.
2. **Tags calculadas no Aspen Calc** (nomes provisórios, a definir com o administrador):
   - `307-GI-001-SUP` = `307-GI-001 + 307-GI-002` (limite superior da banda)
   - `307-GI-001-INF` = `307-GI-001 - 307-GI-002` (limite inferior da banda)
   - `307-REF-100`, `307-REF-150`, `307-REF-200` com valores fixos
   Plote as bandas em linha fina tracejada `#5CC8F0` e as referências em tracejado amarelo, laranja e vermelho.
3. **Sem cálculo:** deixe só a carga média na tendência nesta primeira versão. É a opção recomendada para colocar a tela no ar rápido.

### 8.3 Minitendências (objetos 5 e 9)

Mesmo procedimento, com tamanho 210 x 100, escala fixa (0 a 500 para a carga média e 0 a 150 para o desvio padrão), janela de 8 horas, sem legenda, sem eixos visíveis (se houver a opção) e linha `#5CC8F0`.

## 9. Painel complementar "Referência do laboratório" (objetos 38 a 44)

Fica no canto inferior direito, com letras menores que as do módulo, para não competir com o destaque da tela.

| Objeto | Conteúdo | Tag e mapa | Observação |
|---|---|---|---|
| 38 | Compressão do último resultado MES | `M4710050020-0031-HHLFU`, `IP_MESVALOR`, `VAL` | Arial negrito 18 pt, sem casas decimais |
| 39 | Horário do resultado MES | `M4710050020-0031-HHLFU`, atributo de data e hora | Arial 11 pt |
| 40 | Média do módulo na mesma janela | `307-CMP-MOD-JAN` (tag calculada, opcional) | Deixe de fora até a tag existir |
| 41 | Diferença módulo menos MES | `307-CMP-DIF` (tag calculada, opcional) | Deixe de fora até a tag existir |
| 42 | Modo do 07CR001 | `U04-07CR001`, `IP_PIDMAP03`, `VAL` | Mesmo vínculo da tela PQ US04. Manual em laranja com símbolo ◆ |
| 43 | Modo do 07CR002 | `U04-07CR002`, `IP_PIDMAP03`, `VAL` | Idem |
| 44 | Situação da comparação | `307-CMP-SITUACAO` (tag calculada, opcional) | Sem a tag, a leitura é feita pelos modos 42 e 43 |

Dica: abra a tela PQ US04 no editor, copie os indicadores de modo do 07CR001 e do 07CR002 e cole nesta tela. Assim o vínculo com `IP_PIDMAP03` vem pronto; depois ajuste só a aparência.

O método completo da comparação está em `../04_comparacao_modulo_mes/analise_e_proposta.md`.

## 10. Indicadores que dependem de cálculo (opcionais)

| Objeto | Indicador | Tag calculada sugerida (Aspen Calc) | Sem a tag calculada |
|---|---|---|---|
| 12 | Quantidade de tags com qualidade Good | `307-LCE-QTD-GOOD`: soma de 1 para cada tag com qualidade Good | Não escreva um número fixo. Deixe o objeto de fora e mantenha só as cinco qualidades individuais (13 a 17) |
| 18 | Tempo desde o último ensaio | `307-LCE-DEFASAGEM-H`: horas entre agora e o `IP_INPUT_TIME` de 307-GI-001; cor laranja acima de 4 | Use o campo de data e hora (objeto 2) e oriente a equipe a conferir visualmente |

## 11. Logotipo (objeto 1)

Com a ferramenta de imagem, insira o arquivo oficial do logotipo em **X = 32, Y = 16**, mantendo a proporção dentro de 200 x 56 pixels. Não redesenhe a marca.

## 12. Trocar o mapa pelo fundo definitivo

1. Selecione a imagem de fundo e, nas propriedades, troque o arquivo para `fundo_1920x1080.png` (ou apague e insira de novo com X = 0, Y = 0, 1920 x 1080).
2. Envie para trás novamente.
3. Salve.

## 13. Teste rápido de aceitação

| Verificação | Esperado |
|---|---|
| Cinco tags aparecem com valor, horário e qualidade | Valores iguais aos do Tag Browser |
| Carga média com 1 casa decimal | Exemplo: `286,4` |
| Qualidade diferente de Good (peça à automação para simular ou consulte um período com falha) | Texto da qualidade em laranja |
| Faixa ativa | Texto "Ativo" (ou 1) na cor da severidade |
| Tendência | Janela de 8 h, escala 0 a 500 |
| Leitura a distância | Valores principais legíveis a 3 metros do monitor |

## 14. Limitações conhecidas desta versão

- O tratamento "Sem dado confiável" (valor em cinza no lugar do número quando a qualidade não é Good) só é completo com cor condicional baseada na qualidade, se o editor permitir, ou com tag calculada. Sem isso, a qualidade fica visível ao lado de cada valor.
- A severidade das faixas está **proposta** e precisa da aprovação do LCE.
- Com as tags ainda em 0 (leitura de 30/09/2026), a faixa "menor que 100" pode parecer ativa se o PLC passar a calcular a partir de 0. Validar com o primeiro ensaio real.

## 15. O que me enviar para a próxima rodada

1. Captura da janela de propriedades de um **campo de dados** e de uma **tendência** no seu editor.
2. Resultado do passo 2 (tipo de registro e nomes dos campos).
3. Arquivo oficial do logotipo.
4. Aprovação ou ajuste da severidade das faixas e do limite de 4 h.
