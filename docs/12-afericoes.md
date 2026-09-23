# 12 — Aferições e comparativos: da planilha dos operadores para o painel

## Resposta curta

**É possível, e está feito do lado da planilha.** O arquivo
`afericoes/CALIBRAÇÃO 09 - SETEMBRO.xlsm` é a planilha de setembro que você
mandou, com **duas abas ocultas a mais**. Nenhuma célula, fórmula, macro,
gráfico ou modo de exibição das abas dos operadores foi alterado. Eles
continuam preenchendo tudo nos mesmos lugares.

| Aba oculta | Tabela | Para que serve |
|---|---|---|
| `BD_Afericoes` | `tbl_Afericoes` | Consolida as 9 abas de ensaio em uma tabela única, no formato que o Power BI lê melhor: **1 linha por equipamento, ensaio e data**. São 796 linhas de fórmula (setembro ocupa 294). |
| `BD_Limites` | `tbl_Limites` | Nominal, limites e tolerância de cada ensaio, **num lugar só**, com a fonte de cada número. |

## Colunas de `tbl_Afericoes`

| Coluna | Conteúdo |
|---|---|
| Ensaio | Tambor de Abrasão, Gran. Fina Alpine, Blaine, Comparativo Fisher, Compressão, Granulometria, Umidade, Tamb 5 kg x 15 kg, Verificação Peneiradores |
| Tipo | Calibração (valor dentro de faixa), Comparativo (diferença entre dois métodos) ou Verificação (checklist) |
| Data, Responsável, Letra | Copiados da linha do ensaio |
| Equipamento | TAG (66TA05, 66AG09, 66PS04, PN0423…) ou o par comparado (`66AN11 x Estufa`) |
| Parâmetro, Unidade | O que foi medido |
| Valor, Referência, Diferença | Valor medido, valor de referência (nominal ou o outro método) e a diferença entre os dois |
| Lim. Inferior, Lim. Superior, Tolerância | Lidos de `tbl_Limites` |
| Resultado | **Conforme / Não conforme / Informativo**, recalculado aqui com a mesma regra para todos os ensaios |
| Observação | Texto da coluna de observação, "PARADA" da compressão, itens NÃO OK e "Sem TAG" dos peneiradores |
| Origem | Aba e célula de onde veio o valor, para auditoria |

**Regras de resultado:**

- **Calibração:** Conforme se o valor está entre o limite inferior e o superior.
- **Comparativo:** Conforme se |diferença| ≤ tolerância.
- **Peneiradores:** Não conforme se qualquer item estiver "NÃO OK".
  - Valor = itens OK.
  - Referência = itens verificados.

## Por que assim

- **O operador não muda nada.** As fórmulas apontam para as células de
  sempre. A consolidação é invisível para quem preenche.
- **Formato longo (tidy).** Um único visual atende todos os ensaios: basta
  filtrar por Ensaio ou Equipamento. Um ensaio novo vira linhas novas, não
  colunas novas nem consultas novas.
- **Limites num lugar só.** Hoje cada limite está escrito dentro de uma
  fórmula diferente, às vezes com número diferente para o mesmo ensaio. Na
  `tbl_Limites` você muda o limite uma vez.
- **Não herda fórmulas quebradas.**
  - Granulometria e Tamb 5×15 são recalculados a partir das massas em
    gramas.
  - Por isso o bloco da direita da Granulometria passa a funcionar, embora
    a coluna K das abas esteja com `=J12-CR12`.
- **Só funções do Excel 2007** (IF, SUM, COUNTIF, INDEX/MATCH…). Funciona
  no Excel desktop e no Excel Online do SharePoint.

## Como foi verificado

Não é possível abrir o Excel aqui, então cada uma das 7.164 fórmulas foi
calculada num motor de fórmulas à parte, com os dados reais de setembro.
Os valores batem com os que a própria planilha mostra.

| Ensaio | Painel (BD) | Aba do operador |
|---|---|---|
| Tambor 66TA05, 09/09 | 24,79 rpm | E6 = 24,79 |
| Granulometria −6,3 mm, 14/09 | 1,604 × 1,555 | Q18 / R18 |
| Granulometria RG, 14/09 | 0,829 × 0,796 | Q21 / R21 |
| Tamb 5×15, 14/09 | 93,80 × 93,53 | S10 / T10 |
| Umidade 66AN11, 06/09 | 0,97 → Não conforme | J6 = CONFIRMAR |

Outras verificações:

- As partes novas do arquivo foram validadas contra o esquema oficial do
  formato Excel (ISO 29500).
- O projeto VBA está byte a byte idêntico ao original.
- As macros do menu só mostram e ocultam abas pelo nome. Nenhuma percorre
  todas as abas, então as abas BD_ não aparecem.
- O arquivo está marcado para recalcular tudo ao abrir.

**Primeiro teste seu:**

1. Abra o arquivo no Excel e digite um valor de teste numa aba qualquer.
2. Clique com o botão direito numa guia → **Reexibir** → `BD_Afericoes`.
3. Confira se a linha correspondente mudou.
4. Oculte a aba de novo.

## O que o operador precisa saber (quase nada)

1. **Não inserir nem excluir linhas ou colunas** nas abas de ensaio. As
   fórmulas apontam para posições fixas. Digitar, apagar e corrigir valores
   pode à vontade.
2. **Mês novo:** "Salvar como" a partir do mês anterior e limpar os valores
   digitados, como já fazem. As abas ocultas vão junto. Não recriar a partir
   de um modelo antigo, porque ele não tem as abas BD_.

## Pontos que precisam da sua decisão

| # | Ponto | O que eu fiz |
|---|---|---|
| 1 | **Compressão, valor nominal.** O título da aba diz 355 kgf/pel (FX −16,0 +12,5). As fórmulas usam 345 na linha 6 e 360 nas linhas 7–8, e o cabeçalho diz 360. | Usei **355**, do título. Com 355, o 66PS04 de 14/09 (378) dá diferença 23 → Não conforme; com 360 daria 18 → Conforme. Qual é o certo? Troque em `BD_Limites`, linha COM16. |
| 2 | **Umidade.** O texto diz diferença aceitável 0,05; a fórmula da aba aceita < 0,06. | Usei 0,05. Uma diferença de 0,055 fica Não conforme aqui e VERDADEIRO na aba. |
| 3 | **Granulometria, −6,3 mm.** A aba só marca "Confirmar" quando LTF − LCE > 0,30, positivo. | Usei a diferença em módulo: LCE maior que LTF em 0,4 também é Não conforme. |
| 4 | **Tambor, bloco "200 voltas"** (linhas 25–34). O 66TA06 calcula com 188 voltas (E25) e o 66TA07 com 200. | Usei o RPM que a aba calcula. Se o 66TA06 também roda 200 voltas, o RPM dele está subestimado em 6 %. |
| 5 | **Blaine Automático** (coluna L). Usa a mesma faixa do manual, 1986–2074. | Mantive. A coluna está vazia em setembro. |

## Problemas que encontrei nas abas dos operadores (não corrigi)

Não corrigi porque você pediu para não mexer onde eles trabalham. A BD não
depende de nenhum deles.

- **Granulometria, bloco da direita:**
  - A coluna K (%) calcula `=J12-CR12`, em vez de `=(J12/J22)*100`.
  - Por isso W17:W22, AE17:AE22 e AI17:AI22 e a média Q26/Q27 dão `#VALUE!`.
- **Resultados 1:**
  - 66PS 05 com `#DIV/0!`.
  - 66PS 10 aponta para a coluna de velocidade.
  - Os limites 280–340 não conversam com os 355/310 da aba Compressão.
- **Compressão:**
  - A coluna T (66PS10) usa 345 enquanto as vizinhas usam 360.
  - "PARADA" em L20 gera `#VALUE!` em T20.
- **Calibração Blaine:** "---" em J (Star) faz a validação K mostrar
  "Confirmar" em dia sem ensaio.

## Power BI: o que vem na v15

**Leitura.** A consulta está pronta em `powerquery/24-tbl_Afericoes.pq`.

- Ela lê a **pasta** do SharePoint, não um arquivo: cada mês novo entra
  sozinho e o histórico se acumula.
- Arquivos sem a tabela são ignorados.
- Lançamentos repetidos (mês copiado sem limpar) entram uma vez só.
- Falta só o endereço da pasta.

**Navegação:**

- **Equipamentos + Paradas numa página só.** A disponibilidade aparece uma
  vez. Embaixo ficam o gráfico de paradas por equipamento e a lista de
  paradas.
- **Sai RELATÓRIOS.**
- Menu com 4 botões: **Visão Geral · Equipamentos · Calibração · Aferições**.

**Página AFERIÇÕES:**

1. **Faixa de indicadores:**
   - % conforme no período.
   - Nº de aferições.
   - Nº de não conformidades.
   - **Aderência à rotina**: aferições feitas ÷ programadas. O Tambor é toda
     segunda, a Alpine seg/qua/sáb, o Blaine a cada turno, os Peneiradores
     ter/qui/sáb, a Granulometria e o Tamb 5×15 semanais.
2. **Mapa de calor ensaio × dia.** Mostra num relance o que foi feito, o
   que falhou e o que não foi feito.
3. **Carta de controle**, com seletor de ensaio e equipamento:
   - Valor ao longo do tempo com a faixa Lim. Inferior–Superior sombreada e
     a linha do nominal.
   - A **detecção de anomalias** nativa do Power BI destaca pontos fora do
     padrão antes de estourarem o limite.
4. **Comparativos:** barras de diferença ÷ tolerância. Acima de 100 %
   passam a vermelho.
5. **Tabela de não conformidades,** com observação e responsável. Um clique
   leva ao detalhe.
6. No serviço do Power BI:
   - **Alerta de dados** no cartão de não conformidades, que envia um aviso
     quando passa de 0.
   - Opcional: **Power Automate** manda o e-mail ao técnico do turno.

**Para montar a v15 preciso de:**

1. O **endereço da pasta** no SharePoint onde ficam os arquivos mensais.
   Basta o link da pasta, copiado do navegador.
2. As respostas aos pontos 1 a 5 acima. Pode ser só "mantém como fez".
