# 07 — Ícones e menu lateral

## O que mudou nesta entrega

A faixa de navegação horizontal que ficava abaixo do cabeçalho foi removida.
No lugar dela, cada uma das 9 páginas ganhou um **menu lateral próprio**, com
9 botões empilhados dentro da barra azul-marinho da esquerda:

| # | Botão | Ícone | Destino |
|---|-------|-------|---------|
| 1 | VISÃO GERAL | casa | VISÃO GERAL |
| 2 | EQUIPAMENTOS | cubo | EQUIPAMENTOS |
| 3 | CALIBRAÇÕES | selo de aprovação | CALIBRAÇÕES |
| 4 | INTERVENÇÕES | chave de boca | INTERVENÇÕES |
| 5 | PARADAS | pausa | PARADAS |
| 6 | SOBRESSALENTES | caixa | SOBRESSALENTES |
| 7 | NOTAS / OMs | folha | NOTAS / OMs |
| 8 | INSPEÇÕES | lupa | INSPEÇÕES |
| 9 | RELATÓRIOS | barras | RELATÓRIOS |

O botão da página em que você está fica **azul claro (#1F6FB2)**; os demais
ficam no azul-marinho do fundo (#002643). É assim que o painel mostra onde
você está sem precisar de nenhuma medida DAX.

Cada botão é um *botão de navegação* de verdade (ação "Navegação de página"),
o mesmo recurso que você usou no exemplo que me mandou. O ícone ao lado é uma
imagem PNG separada, posicionada por cima do botão.

## Ícones nos cartões de indicador

Todos os 42 cartões de KPI das 9 páginas receberam um ícone circular colorido
no canto direito. A cor segue o significado do número:

| Cor | Significado | Exemplos |
|-----|-------------|----------|
| azul `#1F6FB2` | contagem neutra | Total Equipamentos, Inspeções Realizadas |
| verde `#0E7C5A` | situação boa | Em Dia, Aprovados, Concluídas, Disponibilidade |
| âmbar `#EDA100` | atenção / prazo | A Vencer, Em Andamento, Registros com data parcial |
| vermelho `#D03B3B` | problema | Vencidas, Reprovados, Equipamentos com Desvio |
| roxo `#6B4E9B` | tempo | Lead Time Médio, Total Intervenções |
| petróleo `#0F766E` | percentual / base | % Conformidade, Cobertura Média, % Base Confiável |

## Arquivos de imagem

As 32 imagens ficam em
`Gestão equioamentos power pbip.Report/StaticResources/RegisteredResources/`
e estão declaradas em `definition/report.json`, na seção `resourcePackages`.
São PNG com fundo transparente, 128×128 px — se um dia você quiser trocar
algum ícone, basta substituir o arquivo mantendo **o mesmo nome**.

## Depois de instalar

1. Abra o `.pbip` no Power BI Desktop.
2. Clique em **Página Inicial → Atualizar**.

O passo 2 não é opcional: a correção que elimina os `--` dos segmentadores
está no Power Query (a de-para de inspeções passou a descartar as 3 linhas
sem `TagKey`). Sem atualizar, os dados antigos continuam em memória e os
`--` continuam aparecendo.

---

## Correção da 1ª versão (menu branco)

Na primeira versão os 9 botões apareceram como **pastilhas brancas vazias**
na barra lateral. A causa estava no meu próprio arquivo de tema: em
`visualStyles` eu tinha a regra coringa

```json
"*": { "*": { "background": [{ "show": true, "color": "#FFFFFF" }],
              "border":     [{ "show": true, "radius": 8 }],
              "dropShadow": [{ "show": true }] } }
```

Essa regra é o que dá aos cartões de KPI o visual de cartão branco com
sombra — e ela vale para **todo** visual, botões e imagens inclusive. Em
cima da barra azul-marinho isso virou uma pastilha branca por botão, e o
rótulo branco e o ícone branco ficaram invisíveis dentro dela.

Duas correções:

1. **No tema**, `actionButton`, `image`, `shape` e `textbox` passaram a ter
   `background`, `border` e `dropShadow` desligados. Só `cardVisual` e os
   gráficos continuam com o visual de cartão.
2. **No relatório**, cada item do menu virou uma imagem PNG única de
   204×36 px que já traz o fundo, o ícone e o texto desenhados. Por cima
   dela fica um botão transparente que só carrega a ação de navegação.

A segunda mudança é o que garante o resultado: o desenho do menu não
depende mais de nenhuma propriedade de formatação do Power BI — só de a
imagem carregar (o que os ícones dos cartões já provaram que funciona) e
de a navegação do botão funcionar (o que o seu botão de exemplo provou).

São 18 imagens: cada um dos 9 itens em dois estados, `-on` (fundo azul
`#1F6FB2`, página atual) e `-off` (fundo transparente, deixa passar o
azul-marinho da barra).

## Indicadores de OM

`Total OMs`, `OMs Abertas`, `OMs Em Andamento` e `OMs Concluídas` passaram
a terminar em `+ 0`, para mostrarem **0** em vez de `--`.

Atenção ao que esse 0 significa hoje: a aba **"Notas Manutenção" da
planilha está vazia** — só o cabeçalho, nenhuma linha. Então o 0 é
literalmente "nenhum registro carregado", não "nenhuma OM aberta". A
página NOTAS / OMs traz esse aviso escrito na tela. Assim que a aba for
preenchida, os quatro indicadores passam a contar de verdade sem nenhuma
mudança no modelo.

---

## Versão 4 — o menu branco e a VISÃO GERAL do mockup

### O que estava acontecendo de verdade

Medindo os pixels dos prints (e não só olhando), ficou claro que:

- o **ícone** do menu aparecia, no tamanho e na posição certos;
- o **texto** ao lado não aparecia — nem um pixel claro, nem com contraste
  forçado ao máximo;
- a cor da pastilha era **uniforme nas 9 linhas** da mesma página, e mudava
  de página para página (9 azuis numa, 7 brancas + 2 azuis noutra, 9 brancas
  na terceira).

Ícone e texto estão **no mesmo arquivo PNG**. Se o ícone aparece, o texto
tem de aparecer. Como não aparecia, o que estava sendo desenhado não era a
pílula nova — era o ícone solto de 20×20 px da versão anterior, com os
botões novos (sem preenchimento definido, portanto com a cor padrão do
tema) desenhados por cima das pílulas.

Ou seja: **a pasta instalada tinha os visuais da versão antiga e os da nova
ao mesmo tempo.** Quando se copia a pasta nova por cima sem apagar a
antiga, o Windows substitui os arquivos de mesmo nome mas **mantém** as
pastas de visuais que só existiam na versão antiga. O Power BI então
desenha as duas gerações empilhadas.

### As três defesas desta versão

1. **O botão de navegação desceu para baixo da pílula** (z 6900 contra
   7000). Ele não tem mais como encobrir o desenho, tenha ou não
   preenchimento. Os dois — pílula e botão — carregam a ação de navegação.
2. **A moldura subiu no eixo Z** (menu 7000, filtros 7500, barra lateral
   6000). Mesmo que sobrem visuais de uma versão antiga na pasta, eles
   ficam por baixo e o menu continua legível.
3. **Carimbo de versão no rodapé.** Todas as páginas terminam com
   `· v4`. Se o rodapé não mostrar `v4`, a pasta instalada não é esta.

### VISÃO GERAL reconstruída

Agora segue a grade do mockup:

- **8 indicadores numa faixa só**, na ordem do mockup: Total Equipamentos,
  Equipamentos Críticos, Calibrações Em Dia, Calibrações A Vencer,
  Calibrações Vencidas, OMs Abertas, Horas Paradas, Total Intervenções.
- **9 painéis numa grade 3×3**, que é o que estava faltando:

  | | coluna 1 | coluna 2 | coluna 3 |
  |---|---|---|---|
  | **linha 1** | Disponibilidade por Grupo Operacional | Status das Calibrações | Próximas Calibrações a Vencer |
  | **linha 2** | Status das OMs | Intervenções por Disciplina | Horas Paradas por Equipamento |
  | **linha 3** | Notas Abertas por Área | Inspeções · Conformidade | Alertas de Sobressalentes |

- **4 filtros** na barra lateral, como no mockup: Laboratório, Grupo
  Operacional, Área e **Status Operacional** (este é novo).

Duas diferenças conscientes em relação ao mockup: "Inspeções ·
Conformidade" é um anel em vez de um velocímetro, e "Notas Abertas por
Área" e "Horas Paradas por Equipamento" são tabelas montadas a partir de
tabelas que já funcionam no arquivo. Preferi reaproveitar visuais já
provados a introduzir um tipo de visual novo que eu não teria como testar
aqui.
