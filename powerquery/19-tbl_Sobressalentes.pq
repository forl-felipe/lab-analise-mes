// ─────────────────────────────────────────────────────────────────────────────
// CONSULTA: tbl_Sobressalentes                          (27 itens · NOVA)
// ─────────────────────────────────────────────────────────────────────────────
// A tabela existe na planilha desde sempre e nunca foi carregada no modelo.
//
// ⚠ ESTADO DA BASE — o painel "Alertas de Sobressalentes" do mockup ainda NÃO
//   pode ser construído. As três colunas de que ele depende estão vazias:
//       Cobertura (%)   0 de 27 preenchidos
//       Situação        0 de 27
//       Compatível com  0 de 27
//   E os dois insumos do cálculo são constantes:
//       Qtd. Atual      = 1 em todas as 27
//       Estoque Mín.    = 1 em todas as 27
//
//   Cobertura NÃO deve ser digitada — é Qtd. Atual / Estoque Mín. Esta consulta
//   já a calcula. Quando as quantidades reais forem lançadas, o indicador passa
//   a funcionar sozinho, sem retrabalho.

let
    Origem = fx_Tabela( "tbl_Sobressalentes" ),

    Tipos = Table.TransformColumnTypes(
        Origem,
        {
            { "Código",         type text },
            { "Descrição",      type text },
            { "Categoria",      type text },
            { "Qtd. Atual",     Int64.Type },
            { "Unidade",        type text },
            { "Estoque Mín.",   Int64.Type },
            { "Local",          type text },
            { "Compatível com", type text }
        },
        "pt-BR"
    ),

    // 'Consumivel ' (com espaço à direita, 26x) e 'sobressalente' (minúsculo, 1x)
    CategoriaLimpa = Table.TransformColumns(
        Tipos,
        {{ "Categoria", each
            let C = if _ = null then "" else Text.Proper( Text.Trim( Text.From( _ ) ) )
            in  if C = "" then "Não classificado" else C,
           type text }}
    ),

    DescricaoLimpa = Table.TransformColumns(
        CategoriaLimpa,
        {{ "Descrição", each if _ = null then null else Text.Trim( Text.Replace( Text.From( _ ), "#(lf)", " " ) ), type nullable text }}
    ),

    // cobertura CALCULADA, substituindo a coluna vazia da origem
    SemCobertura = Table.RemoveColumns( DescricaoLimpa, { "Cobertura (%)", "Situação" }, MissingField.Ignore ),

    Cobertura = Table.AddColumn(
        SemCobertura, "Cobertura",
        each if Record.Field( _, "Estoque Mín." ) = null or Record.Field( _, "Estoque Mín." ) = 0
             then null
             else Record.Field( _, "Qtd. Atual" ) / Record.Field( _, "Estoque Mín." ),
        type nullable number
    ),

    Situacao = Table.AddColumn(
        Cobertura, "Situação",
        each if [Cobertura] = null      then "Sem parâmetro"
             else if [Cobertura] < 0.5  then "Crítico"
             else if [Cobertura] < 1    then "Atenção"
             else "Adequado",
        type text
    ),

    // enquanto "Compatível com" estiver vazio não há ligação com equipamento
    TemVinculo = Table.AddColumn(
        Situacao, "Vinculado a Equipamento",
        each [Compatível com] <> null and Text.Trim( Text.From( [Compatível com] ) ) <> "",
        type logical
    )
in
    TemVinculo
