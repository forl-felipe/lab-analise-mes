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
