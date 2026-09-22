# 08 — Identidade visual Samarco

## Paleta institucional (o arquivo que você me mandou)

`tema/samarco-cores.css` é o arquivo original, guardado aqui sem alteração:

| Variável | Hex | Onde aparece no painel |
|---|---|---|
| `--samarco-azul-profundo` | `#002643` | fundo da barra lateral |
| `--samarco-azul-samarco` | `#00335A` | cabeçalho de todas as páginas |
| `--samarco-azul-titulo` | `#003B6C` | títulos dos painéis |
| `--samarco-azul-medio` | `#004070` | — |
| `--samarco-amarelo` | `#FFC000` | régua abaixo do cabeçalho e trilho do item ativo do menu |
| `--samarco-azul-claro` | `#59C6F2` | base da cor de percentuais |
| `--samarco-laranja` | `#F37021` | base da 2ª cor categórica |
| `--samarco-cinza-logo` | `#B5BEC4` | subtítulo da barra lateral |
| `--samarco-azul-acinzentado` | `#61889B` | texto terciário |

O logo é o SVG oficial do seu kit visual, rasterizado em PNG a 3× e
reduzido — por isso fica nítido. Era a palavra "SAMARCO" escrita à mão
que aparecia antes.

## Por que as cores dos gráficos não são os hex puros do manual

Os tons institucionais foram feitos para cabeçalho, logo e fundo. Quando
viram barra, fatia de rosca ou ponto de gráfico, três deles reprovam:

- `#00335A` é **escuro demais** — some contra o texto e contra os outros;
- `#B5BEC4` e `#61889B` têm **saturação baixa demais** — lidos lado a
  lado com um cinza de grade, parecem a mesma coisa;
- `#FFC000` e `#59C6F2` ficam **abaixo de 3:1 de contraste** sobre fundo
  branco — a barra some no cartão.

Rodei a paleta no validador de contraste e daltonismo (OKLab, separação
mínima ΔE 8 para deuteranopia/protanopia/tritanopia). O resultado com os
hex puros: **reprovado em 3 dos 5 testes**.

Então mantive as **matizes** da Samarco e ajustei só a luminosidade e a
saturação até passar em tudo:

| Papel | Cor | Origem |
|---|---|---|
| categórica 1 | `#1B6FB0` | azul Samarco clareado |
| categórica 2 | `#E0620F` | laranja Samarco escurecido |
| categórica 3 | `#36A9C9` | azul claro Samarco escurecido |
| categórica 4 | `#C99400` | amarelo Samarco escurecido |
| categórica 5 | `#8A4B9E` | complemento (a paleta não tem 5ª matiz) |

Resultado do validador: **aprovado nos 5 testes**, com aviso de contraste
em duas cores — aviso que fica coberto porque todo gráfico do painel tem
legenda e rótulo de dados.

## Cores de status

Status é um papel reservado: não entra no rodízio das categóricas.

| Estado | Cor |
|---|---|
| em dia | `#1B8F5E` |
| a vencer | `#C99400` (amarelo Samarco) |
| vencido | `#C0361B` |
| sem data | `#8E96A0` (neutro, de propósito) |

Verde e vermelho não existem no manual da Samarco, e foi uma escolha
consciente mantê-los: "vencido" em laranja institucional fica a ΔE 5,1 do
amarelo de "a vencer" — abaixo do piso de 15 até para quem enxerga todas
as cores. Duas situações opostas ficariam indistinguíveis. Na identidade
institucional eles não aparecem; só no semáforo dos indicadores.

**Um ponto honesto:** o Power BI pinta as fatias da rosca na ordem da
paleta, não pelo nome da categoria. Então a rosca "Status das Calibrações"
usa as cores categóricas, não o semáforo acima. Fixar cor por categoria é
uma propriedade de formatação que eu não consegui verificar num arquivo
real, e eu já quebrei o menu duas vezes chutando propriedade. Se quiser,
dá para fazer em 30 segundos na interface: clique na rosca → Formatar →
Cores dos dados → escolha a cor de cada status.
