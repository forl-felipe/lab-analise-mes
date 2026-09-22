# Decisões de arquitetura — Painel Gestão de Equipamentos

> Registro das decisões tomadas a partir da análise do `.pbix` e da base
> `Base_PowerBI_Gestao_Equipamentos.xlsx`. Cada decisão traz a evidência que a sustenta.

---

## D1 — O modelo tem TRÊS domínios independentes, não um

A investigação das chaves mostrou três universos sem nenhuma interseção:

| Domínio | Tabela | Linhas | Padrão de TAG | Exemplo |
|---|---|---|---|---|
| A · Equipamentos de amostragem | `tbl_Equipamentos` | 34 | `U03-*` / `U04-*` | `U03-02CR001` |
| B · Instrumentos de laboratório | `tbl_Calibracao` | 208 | `66 XX NN` / numérico | `66 AN 13`, `449212` |
| C · Inspeções de campo | `tbl_Inspecoes` | 5 | `AM-XX-NN` | `AM-CT-01` |

Interseção entre A e B: **0 tags**. Entre A e C: **0 tags**.
Também são disjuntos os nomes de equipamento e os grupos:

- Grupos em A: `PF · Usina 03`, `PF · Usina 04`, `PQ · Usina 03`, `PQ · Usina 04`, `Embarque · Torre 3`
- Grupos em C: `Amostragem Primaria`, `Amostragem Secundaria`, `Amostragem de Polpa`, `Divisao de Amostra`

**Decisão:** modelar três estrelas ligadas apenas pelas dimensões que realmente
compartilham. Não criar relacionamentos que não se sustentam nos dados.

---

## D2 — A ponte entre A e B é `Laboratório`, não o equipamento

Única dimensão com valores em comum:

| | LCP | LCE | LDP |
|---|---|---|---|
| `tbl_Equipamentos` | 26 | 8 | — |
| `tbl_Calibracao` | 119 | 57 | 32 |

A coluna `Local` **não** serve de ponte: em A são locais de planta
(`USINA 03`, `USINA 04`, `Torre 3`); em B são salas de laboratório
(`LTF`, `INSUMOS`, `PREPARAÇÃO`, `QUÍMICO`, `METALÚRGICO`, `SÓLIDO LÍQUIDO`,
`POT GRATE`, `MICROSCOPIA`). Único encontro é `Torre 3` / `TORRE 3`, que é
uma localização física compartilhada, não uma chave.

**Decisão:** criar `DimLaboratorio` (LCP, LCE, LDP, N/A) como dimensão
compartilhada. Os filtros de laboratório passam a atravessar os dois domínios.
Cruzamentos no grão de equipamento continuam impossíveis — e isso deve ficar
explícito no painel, não escondido.

---

## D3 — O domínio C resolve-se com um de-para de 5 linhas

São apenas **5 tags distintas** em inspeções. Um de-para manual conecta o
domínio C ao mestre de equipamentos a custo quase zero.

**Decisão:** criar `DePara_Inspecao_Equipamento` na planilha de origem, com
os pares preenchidos por quem conhece o campo. Enquanto não for preenchido,
as inspeções permanecem isoladas (comportamento honesto, e igual ao de hoje).

**Ação de fundo:** o formulário de inspeção precisa ter a lista de
equipamentos vinculada ao mestre — senão o de-para cresce a cada equipamento novo.

---

## D4 — Separar STATUS METROLÓGICO de STATUS DE PRAZO

Esta é a descoberta de maior impacto. A base tem duas informações distintas
que o painel atual trata como uma só — na verdade, ignora uma delas por completo.

**`Status Calibração` (vem da origem) — resultado metrológico:**

| Valor | Linhas | Significado |
|---|---|---|
| `OK` | 137 | Aprovado |
| `N/OK` | 47 | **Reprovado** |
| `E/C` | 23 | Em calibração |
| `NOK` | 1 | erro de digitação de `N/OK` |

**`StatusAtual` (calculado no Power Query) — situação do prazo:**

| Valor | Linhas |
|---|---|
| Em Dia | 122 |
| Vencido | 65 |
| Sem Data | 22 |
| A Vencer | 5 |

O relatório atual usa **somente** o prazo. Resultado: **48 instrumentos
reprovados na calibração estão invisíveis no painel** — inclusive os que estão
"Em Dia" de prazo. Um instrumento com certificado válido e resultado N/OK é
pior que um vencido: ele está em uso e medindo errado.

**Decisão:** promover o status metrológico a KPI de primeira classe e criar um
indicador combinado `Instrumentos em Risco` = reprovados **ou** vencidos.

---

## D5 — Sobressalentes: carregar a tabela, não construir o painel do mockup

`tbl_Sobressalentes` existe na planilha (27 itens) e nunca foi carregada. Mas as
três colunas que o mockup usa estão **100% vazias**:

| Coluna | Preenchimento |
|---|---|
| `Cobertura (%)` | 0 de 27 |
| `Situação` | 0 de 27 |
| `Compatível com` | 0 de 27 |
| `Qtd. Atual` | 27 de 27 — todos `1` |
| `Estoque Mín.` | 27 de 27 — todos `1` |

Além disso, `Categoria` tem `'Consumivel '` (com espaço à direita) em 26 linhas e
`'sobressalente'` em 1. O conteúdo é majoritariamente consumível de laboratório
(etiquetas, papel de filtro, potes, pincéis), não sobressalente de equipamento.

**Decisão:** carregar a tabela já tratada, e no lugar do painel "Alertas de
Sobressalentes" exibir o que existe de fato (itens por Local e Categoria) com um
aviso de cobertura pendente. O painel de alertas entra quando `Qtd. Atual`,
`Estoque Mín.` e `Compatível com` estiverem preenchidos — aí `Cobertura (%)`
passa a ser calculada, não digitada.

---

## D6 — Status de OMs: usar o domínio real, não o do mockup

`tbl_Listas` define o domínio válido de `Status Nota`:

`Em andamento` · `Concluído` · `Aguardando peças` · `Cancelado`

O mockup propõe `Concluídas` / `Em Andamento` / `Abertas`. **`Abertas` não existe**
no domínio. `tbl_Notas` está vazia (0 linhas), então nada quebra hoje — mas a
rosca precisa nascer com as categorias certas.

**Decisão:** usar os quatro valores do domínio. "Aberta" passa a ser a medida
derivada `OMs Abertas = Em andamento + Aguardando peças`, com o nome explicando
a regra.

---

## D7 — Parser de datas explícito no Power Query

`Última Calibração` tem **seis formatos concorrentes** na mesma coluna:

| Formato | Linhas | Exemplo |
|---|---|---|
| texto `dd/mm/aaaa` | 100 | `01/06/2025` |
| texto `mm/aa` — **sem dia** | 61 | `05/26` |
| texto livre | 29 | `NOVA`, `nova`, `SEM `, `05/2026` |
| texto `dd/mm/aa` | 6 | `01/07/26` |
| vazio | 6 | |
| serial do Excel | 6 | `46174` |

Apenas 48% está em formato de data completo e não ambíguo.

`Próximo Vencimento` está bem melhor: **191 de 208 (92%) são seriais do Excel**
válidos. Os 17 restantes são `NOVA`/`nova` (8), vazio (6) e dois casos em que
o texto *"SEM CALIBRAÇÃO"* transbordou da coluna anterior, deixando `CALIBRAÇÃO`
e `CALIBR` dentro de uma coluna de data.

**Decisão:** `Próximo Vencimento` é a fonte primária dos indicadores de prazo.
`Última Calibração` passa por um parser explícito de 6 formatos e ganha uma
coluna `Qualidade Data` que marca cada linha como `OK`, `Mês sem dia` ou
`Não interpretável` — para dirigir a limpeza na origem em vez de escondê-la.

---

## D8 — Fuso horário no cálculo de vencimento

O código atual usa `DateTime.LocalNow()`. No Power BI Service isso é avaliado
em **UTC**, não no horário de Brasília — o que desloca o corte de dia.

**Decisão:** `DateTimeZone.SwitchZone(DateTimeZone.UtcNow(), -3)` em todo
cálculo de "hoje". Os limiares (30 e 15 dias) passam a ser lidos de
`tbl_Config`, onde já existem, em vez de ficarem fixos no código.

---

## D9 — Status em azul/âmbar/vermelho, não verde/âmbar/vermelho

Medição com o validador de paleta (ΔE em OKLab ×100, piso seguro = 8):

| Par | ΔE deuteranopia | Veredito |
|---|---|---|
| verde `#0CA30C` ↔ vermelho `#D03B3B` | **4,1** | reprovado |
| azul `#1F6FB2` ↔ vermelho `#D03B3B` | **18,2** | aprovado |

Verde e vermelho são praticamente indistinguíveis para quem tem deuteranomalia
— cerca de 8% dos homens. Numa rosca de status, onde a cor é o único código,
isso é uma falha de leitura, não uma questão de gosto.

**Decisão:** a codificação de status usa **azul (em dia) · âmbar (a vencer) ·
vermelho (vencido) · cinza (sem data)**. Verde fica reservado para variação
positiva em texto. Nos cards de KPI, onde sempre há ícone + rótulo, a regra é
mais folgada — mas mantemos a mesma paleta por consistência.

Se você preferir a convenção verde/vermelho, o hex é `#0CA30C`: nesse caso
todo gráfico de status precisa de rótulo direto visível, nunca cor sozinha.

---

## D10 — Não renomear as tabelas existentes agora

Renomear `Tbl_Paradas` → `tbl_Paradas` ou adotar prefixos `Dim`/`f` quebraria a
ligação de 60+ visuais já construídos.

**Decisão:** manter os nomes atuais nesta fase. Tabelas **novas** nascem com a
convenção correta (`DimLaboratorio`, `DimCalendario`). A padronização completa
entra na Fase 3, quando as páginas forem reconstruídas de qualquer forma.
