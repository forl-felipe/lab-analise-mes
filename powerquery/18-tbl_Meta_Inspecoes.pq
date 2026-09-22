// ─────────────────────────────────────────────────────────────────────────────
// CONSULTA: tbl_Meta_Inspecoes                       (5 responsáveis · meta 10)
// ─────────────────────────────────────────────────────────────────────────────
// Observação de nomenclatura: esta consulta lê o item "tbl_Mediçõesinspeções"
// do Excel — o nome da aba não tem relação com o conteúdo (são METAS, não
// medições). Vale renomear a tabela na planilha para tbl_Meta_Inspecoes.
//
// A versão atual carregava 9 linhas, 4 delas em branco. fx_Tabela já filtra.

let
    Origem = fx_Tabela( "tbl_Mediçõesinspeções" ),

    Tipos = Table.TransformColumnTypes(
        Origem,
        { { "Responsavel", type text }, { "Meta_Semanal", Int64.Type } },
        "pt-BR"
    ),

    Limpo = Table.SelectRows(
        Tipos,
        each [Responsavel] <> null and Text.Trim( [Responsavel] ) <> ""
    ),

    NomeLimpo = Table.TransformColumns(
        Limpo,
        {{ "Responsavel", each Text.Trim( Text.From( _ ) ), type text }}
    )
in
    NomeLimpo
