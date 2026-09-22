// ─────────────────────────────────────────────────────────────────────────────
// CONSULTA: tbl_Intervencoes                         (3 linhas com conteúdo)
// ─────────────────────────────────────────────────────────────────────────────
// Correção: a coluna DataHora estava sempre NULA no modelo. O código antigo
// fazia  DateTime.From([Data]) + Duration.From([Hora])  — mas [Hora] é do tipo
// time, e Duration.From sobre time não produz o deslocamento esperado quando a
// hora vem do Excel com base 1899-12-30. Aqui a composição é explícita.

let
    Origem = fx_Tabela( "tbl_Intervencoes" ),

    Tipos = Table.TransformColumnTypes(
        Origem,
        {
            { "Data",                    type date },
            { "Hora",                    type time },
            { "Turno",                   type text },
            { "Equip. (Tag)",            type text },
            { "Disciplina",              type text },
            { "Descrição da Ocorrência", type text },
            { "Ação Executada",          type text },
            { "Técnico",                 type text },
            { "Status",                  type text },
            { "Equipamento (auto)",      type text },
            { "Grupo (auto)",            type text }
        },
        "pt-BR"
    ),

    ComTagKey = Table.AddColumn( Tipos, "TagKey", each fx_NormalizarTag( Record.Field( _, "Equip. (Tag)" ) ), type nullable text ),

    DataHora = Table.AddColumn(
        ComTagKey, "DataHora",
        each if [Data] = null then null
             else if [Hora] = null then DateTime.From( [Data] )
             else #datetime( Date.Year([Data]), Date.Month([Data]), Date.Day([Data]),
                             Time.Hour([Hora]), Time.Minute([Hora]), 0 ),
        type nullable datetime
    ),

    ComID = Table.AddIndexColumn( DataHora, "ID_Intervencao", 1, 1, Int64.Type )
in
    ComID
