# Gestão de Equipamentos e Calibrações — Laboratórios

Código-fonte do modelo e do tema do painel Power BI. Gerado a partir da análise
do `.pbix` e da base `Base_PowerBI_Gestao_Equipamentos.xlsx`.

## Estrutura

```
docs/        diagnóstico, decisões de arquitetura e plano de implementação
powerquery/  consultas M — uma por arquivo, colar no Editor Avançado
dax/         tabelas calculadas e o conjunto completo de medidas
tema/        tema JSON (Exibir ▸ Temas ▸ Procurar temas)
```

**Comece por** [`docs/02-decisoes-arquitetura.md`](docs/02-decisoes-arquitetura.md)
— as 10 decisões e a evidência de cada uma.
Depois [`docs/03-plano-implementacao.md`](docs/03-plano-implementacao.md).

## Os três achados que mudam o projeto

**1 · São três domínios, não um.** Equipamentos de amostragem (34 tags `U03-*`),
instrumentos de laboratório (208 tags `66 XX NN`) e inspeções de campo (5 tags
`AM-XX-NN`). Interseção entre eles: **zero**. O relacionamento
`tbl_Calibracao → tbl_Equipamentos` existe no modelo e não casa **nenhuma** das
214 linhas. A única dimensão realmente compartilhada é `Laboratório`.

**2 · 48 instrumentos reprovados estão invisíveis.** A base traz o resultado
metrológico (`OK` 137 · `N/OK` 47 · `E/C` 23) e o painel usa só o prazo. Um
instrumento com certificado válido e resultado reprovado está em uso medindo
errado — e é o pior caso, não o vencido.

**3 · As regras de negócio já estão na planilha e ninguém as usa.** `tbl_Config`
tem Meta de Disponibilidade, Horas do Mês e os limiares de alerta (que estão
duplicados, fixos, dentro do código M). `tbl_Grupos` tem a regra de
série/redundância de cada grupo. Nenhuma das duas aparece em qualquer medida.

## Grade do layout — 1920×1080

O mockup está em proporção 3:2; o canvas é 16:9. Transposição com folga:

```
Barra lateral       x=0      y=0       224 × 1080
Cabeçalho           x=224    y=0      1696 × 80
Conteúdo            x=248 → 1896   (útil 1648, gutter 12)

Faixa 1 · KPIs      y=96     altura 150
  8 cards de 195px  x = 248, 455, 662, 869, 1076, 1283, 1490, 1697
Faixa 2 · gráficos  y=258    altura 300    3 colunas de 541px
Faixa 3             y=570    altura 226    x = 248, 801, 1354
Faixa 4             y=808    altura 226
Rodapé              y=1044   altura 24
```

## Cores de status

Verde e vermelho medem **ΔE 4,1** sob deuteranopia (o piso seguro é 8) — quem
tem deuteranomalia, cerca de 8% dos homens, não distingue os dois. Azul e
vermelho medem **18,2**. Por isso a codificação de status é:

| Estado | Cor | |
|---|---|---|
| Em Dia · Conforme · Aprovado | `#1F6FB2` | azul |
| A Vencer · Atenção | `#EDA100` | âmbar |
| Vencido · Não Conforme · Reprovado | `#D03B3B` | vermelho |
| Sem Data · Não informado | `#898781` | cinza |

`#002643` fica para a identidade — barra lateral, cabeçalho, títulos — e não
para codificar dado. Toda cor de status vem acompanhada de ícone e rótulo;
nunca sozinha.
