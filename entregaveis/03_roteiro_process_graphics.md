# Roteiro de construção no Aspen Process Graphics Editor V14.2

Tela: monitoramento do ensaio de compressão de pelotas (prensa automática, PLC 307-PLC-001).

Este roteiro é construído por partes. A Parte 1 pode ser executada agora. As demais partes serão detalhadas na Fase 3, após a aprovação do protótipo (`04_prototipo/index.html`).

Os nomes de menus seguem a interface em português da captura de 22/10/2026 (Arquivo, Editar, Visualização, Extrair, Ferramentas, Janela, Ajuda). Se algum comando estiver em outro lugar, procure pelo mesmo nome no menu do botão direito sobre o objeto. Os pontos não testados no editor estão marcados com **(conferir)**.

## Parte 1: preparação da tela base

Objetivo: partir do arquivo `PIMS laboratório atualizado`, manter apenas o campo Comp. Fx Maior e apagar o restante, **sem perder a tela original do Laboratório Químico**.

### 1.1 Salvar uma cópia antes de apagar qualquer coisa

1. Abra `PIMS laboratório atualizado` no editor.
2. Clique em **Arquivo > Salvar como**.
3. Salve com o nome `PIMS compressao prensa 307`, na mesma pasta.
4. Confira se o nome novo aparece na barra de título, no alto da janela: "Aspen Process Graphics Editor V14.2 - aspenONE - [PIMS compressao prensa 307]".

Só continue depois de confirmar o nome na barra de título. Assim, tudo o que for apagado a seguir afeta apenas a cópia, e a tela do Laboratório Químico continua disponível para quem a utiliza.

### 1.2 Identificar o que será mantido

O campo Comp. Fx Maior é formado por três objetos:

| Objeto | Na tela atual | Como confirmar |
|---|---|---|
| Rótulo | Caixa amarela com o texto "Comp. Fx Maior" | Clique uma vez: só a caixa deve ficar selecionada |
| Campo de dados | Valor numérico (ex.: 280) | Clique duas vezes: as propriedades mostram a tag, o mapa `IP_MESVALOR` e a fonte `UBU` |
| Unidade | Texto "Kgf/pel" | Clique uma vez |

Na tela existem dois campos Comp. Fx Maior:

| Lado da tela | Tag | Valor na captura |
|---|---|---|
| Parâmetros USINA 03 | `M710050020-0031-HHLFU` | 280 Kgf/pel |
| Parâmetros USINA 04 | `M4710050020-0031-HHLFU` | 263 Kgf/pel |

Mantenha o da usina definida na hipótese H19 (ou os dois, se for o caso).

Se, ao clicar no campo, vários objetos forem selecionados ao mesmo tempo, eles estão agrupados. Clique com o botão direito e procure **Desagrupar** **(conferir)**. Repita até conseguir selecionar só o rótulo, o valor e a unidade.

### 1.3 Apagar os demais objetos

Forma recomendada, por partes, para reduzir o risco de apagar o que deve ficar:

1. Com a ferramenta de seleção (a seta, no alto da barra de ferramentas da esquerda), clique numa área vazia e arraste um retângulo sobre um bloco de objetos, por exemplo "Nível TQ prep. 01" inteiro.
2. Pressione **Delete**.
3. Repita bloco por bloco: níveis dos tanques, Carregamento, Campanha, Parâmetros das usinas, as quatro tendências LM 01 e LM 02, Set Point de Insumos, Tamboramento e Comp. Fx Menor.
4. Se apagar algo por engano, use **Editar > Desfazer** (Ctrl+Z) imediatamente.

Forma alternativa, mais rápida: **Editar > Selecionar tudo** (Ctrl+A), depois Ctrl+clique ou Shift+clique nos três objetos que devem ficar, para tirá-los da seleção **(conferir qual tecla o editor usa)**, e então **Delete**.

### 1.4 Conferir e salvar

1. Na tela devem restar só o rótulo, o valor e a unidade do Comp. Fx Maior.
2. Use **Editar > Selecionar tudo** (Ctrl+A) para conferir: se aparecerem marcas de seleção em áreas que parecem vazias, ainda há objetos invisíveis ou sobrepostos. O arquivo original tem mais de 3.000 textos "SP Calcário" repetidos, provavelmente ocultos ou empilhados. Apague o que restar.
3. Arraste os três objetos mantidos para o canto inferior esquerdo. Eles serão posicionados no lugar definitivo na Parte 2.
4. Clique em **Arquivo > Salvar**.
5. Não publique a tela no servidor ainda.

Resultado esperado: tela cinza vazia, apenas com o Comp. Fx Maior, e arquivo bem menor que o original (o original tem cerca de 400 KB).

## Correspondência entre o protótipo e os objetos do editor

Visão antecipada da Parte 2 (detalhamento na Fase 3):

| Área do protótipo | Objeto do editor | Tags |
|---|---|---|
| Valores de carga média e desvio padrão | Campo de dados, mapa IP_AnalogMap, fonte UBU | 307-GI-001, 307-GI-002 |
| Minigráficos das últimas 8 h | Tendência pequena, sem eixos | 307-GI-001, 307-GI-002 |
| Comp. Fx Maior (laboratório) | Campo de dados mantido na Parte 1 | M710050020-0031-HHLFU ou M4710050020-0031-HHLFU |
| Tendência da carga média | Tendência com até quatro penas e escala fixa de 0 a 400 | 307-GI-001, banda (cálculos opcionais `_SUP` e `_INF`), Comp. Fx Maior |
| Linhas de referência 100, 150 e 200 | Linhas tracejadas desenhadas sobre a tendência, ou limites de escala **(conferir)** | Não se aplica |
| Faixas de classificação | Campo de dados discreto (IP_DiscreteMap) com cor e texto por estado | 307-QI-001, 307-QI-002, 307-QI-003 |
| Tabela de tags | Campos de dados com os atributos de valor, horário e qualidade **(conferir)** | As cinco tags |
| Janela de 2 h, 8 h e 24 h | Barra de tempo nativa, na parte inferior do editor | Não se aplica |

## Parte 2: montagem da tela

A detalhar na Fase 3.

## Parte 3: botão de acesso e publicação

A detalhar na Fase 3.
