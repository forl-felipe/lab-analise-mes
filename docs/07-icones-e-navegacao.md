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
