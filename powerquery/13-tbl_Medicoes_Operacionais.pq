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
                    // Number.FromText usa a cultura do modelo (pt-BR), onde "." e separador
                    // de MILHAR. Trocar virgula por ponto sem dizer a cultura fazia 0,471
                    // virar 471 e 0,47100000000000003 virar 4,71E+16.
                    {{ coluna,
                      each try ( if _ is number
                                 then Number.From( _ )
                                 else Number.FromText( Text.Replace( Text.Trim( Text.From( _ ) ), ",", "." ), "en-US" ) )
                           otherwise null,
                      type nullable number }}
                 )
            else tabela
    ),

    ComTagKey = Table.AddColumn( ParaNumero, "TagKey Inspecao", each fx_NormalizarTag( [Equip_Tag] ), type nullable text )
in
    ComTagKey
