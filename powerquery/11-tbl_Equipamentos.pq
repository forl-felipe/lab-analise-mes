// ─────────────────────────────────────────────────────────────────────────────
// CONSULTA: tbl_Equipamentos        (grão: 1 linha por equipamento · 34 linhas)
// ─────────────────────────────────────────────────────────────────────────────
// ATENÇÃO — qualidade da base, não do código:
//   Crítico            = "Sim" nas 34 linhas  -> o card "Equipamentos Críticos"
//                        mostra sempre o mesmo valor que "Total Equipamentos"
//   Status Operacional = "Operante" nas 34    -> coluna sem poder de filtro
//   Data Cadastro      = vazia nas 34
// Enquanto essas colunas não forem preenchidas de verdade, os indicadores que
// dependem delas não informam nada. Ver docs/02-decisoes-arquitetura.md.

let
    Origem = fx_Tabela( "tbl_Equipamentos" ),

    Tipos = Table.TransformColumnTypes(
        Origem,
        {
            { "Tag",                type text },
            { "Equipamento",        type text },
            { "Categoria",          type text },
            { "Área",               type text },
            { "Sistema",            type text },
            { "Tipo",               type text },
            { "Grupo Operacional",  type text },
            { "Laboratório",        type text },
            { "Local",              type text },
            { "Crítico",            type text },
            { "Status Operacional", type text },
            { "Data Cadastro",      type date },
            { "Observações",        type text }
        },
        "pt-BR"
    ),

    ComTagKey = Table.AddColumn( Tipos, "TagKey", each fx_NormalizarTag( [Tag] ), type nullable text ),

    // 1/0 para permitir soma; a medida DAX usa CALCULATE sobre [Crítico] = "Sim"
    Criticidade = Table.AddColumn(
        ComTagKey, "Criticidade",
        each if [Crítico] = "Sim" then 1 else 0,
        Int64.Type
    ),

    // alinha o nome da coluna com tbl_Calibracao (lá é "Laboratorio", sem acento)
    LabNormalizado = Table.AddColumn(
        Criticidade, "Laboratorio",
        each if [Laboratório] = null or Text.Trim( Text.From( [Laboratório] ) ) = "" then "N/A"
             else Text.Upper( Text.Trim( Text.From( [Laboratório] ) ) ),
        type text
    )
in
    LabNormalizado
