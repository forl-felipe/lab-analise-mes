// ─────────────────────────────────────────────────────────────────────────────
// CONSULTA: tbl_Medicoes_Operacionais                (3 linhas com conteúdo)
// ─────────────────────────────────────────────────────────────────────────────
// Também estava sem tipagem: Velocidade_Calculada chegava como texto "0,471".
// O gráfico da página INSPEÇÕES usava Contagem de Velocidade_Calculada — ou
// seja, contava ocorrências de texto em vez de medir velocidade.

let
    Origem = fx_Tabela( "tbl_Medicoes_Operacionais" ),

    Texto = Table.TransformColumnTypes(
        Origem,
        {
            { "ID_Inspecao",         type text },
            { "Equip_Tag",           type text },
            { "Equipamento",         type text },
            { "Velocidade_Uniforme", type text },
            { "Resultado_Medicao",   type text }
        },
        "pt-BR"
    ),

    ComData = Table.AddColumn( Texto, "Data Medicao", each fx_ParseData( [Data], false ), type nullable date ),
    TiraData = Table.RemoveColumns( ComData, { "Data" }, MissingField.Ignore ),
    RenData  = Table.RenameColumns( TiraData, {{ "Data Medicao", "Data" }} ),

    // decimais com vírgula -> número (cultura pt-BR)
    ParaNumero = List.Accumulate(
        { "Comprimento_Cortador", "Distancia_Percurso",
          "Tempo_1", "Tempo_2", "Tempo_3", "Tempo_4", "Tempo_5",
          "Tempo_Medio", "Velocidade_Calculada" },
        RenData,
        ( tabela, coluna ) =>
            if List.Contains( Table.ColumnNames( tabela ), coluna )
            then Table.TransformColumns(
                    tabela,
                    {{ coluna, each try Number.FromText( Text.Replace( Text.Trim( Text.From( _ ) ), ",", "." ) ) otherwise null, type nullable number }}
                 )
            else tabela
    ),

    ComTagKey = Table.AddColumn( ParaNumero, "TagKey Inspecao", each fx_NormalizarTag( [Equip_Tag] ), type nullable text )
in
    ComTagKey
