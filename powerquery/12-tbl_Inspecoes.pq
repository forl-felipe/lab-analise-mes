// ─────────────────────────────────────────────────────────────────────────────
// CONSULTA: tbl_Inspecoes                            (5 linhas com conteúdo)
// ─────────────────────────────────────────────────────────────────────────────
// A versão atual é só  Fonte{[Item="tbl_Inspeções"]}[Data]  — sem filtro e sem
// tipagem. Consequências que isso causava:
//   · 195 linhas em branco entravam no modelo e apareciam na tabela do relatório
//   · as 47 colunas viravam TEXTO, incluindo Qtd_NC (somada por SUM) e Data
//
// BUG NA ORIGEM: ID_Inspecao traz o literal "yyyy" —  INS-yyyy0921-AM-CT-01-001
// contra  INS-20260921-AM-CT-01-001  em tbl_Medicoes_Operacionais. A fórmula da
// planilha não substitui o ano, e as duas tabelas nunca cruzam. Corrigir no Excel.

let
    Origem = fx_Tabela( "tbl_Inspeções" ),

    Tipos = Table.TransformColumnTypes(
        Origem,
        {
            { "ID_Inspecao",        type text },
            { "Carimbo_DataHora",   type text },
            { "Semana",             type text },
            { "Mes",                type text },
            { "Ano",                type text },
            { "Turno",              type text },
            { "Responsavel",        type text },
            { "Equip_Tag",          type text },
            { "Equipamento",        type text },
            { "Grupo",              type text },
            { "Tipo_Equipamento",   type text },
            { "Resultado_Inspecao", type text },
            { "Requer_Acao",        type text },
            { "Criticidade",        type text },
            { "Observacoes",        type text }
        },
        "pt-BR"
    ),

    // Data chega como texto "21/09/2026"
    ComData = Table.AddColumn( Tipos, "Data Inspecao", each fx_ParseData( [Data], false ), type nullable date ),
    SemDataAntiga = Table.RemoveColumns( ComData, { "Data" }, MissingField.Ignore ),
    RenomeiaData  = Table.RenameColumns( SemDataAntiga, {{ "Data Inspecao", "Data" }} ),

    // Qtd_NC vem como texto ('0','2','4') e era somada por SUM()
    QtdNC = Table.TransformColumns(
        RenomeiaData,
        {{ "Qtd_NC", each try Number.FromText( Text.Trim( Text.From( _ ) ), "pt-BR" ) otherwise null, Int64.Type }}
    ),

    ComTagKey = Table.AddColumn( QtdNC, "TagKey Inspecao", each fx_NormalizarTag( [Equip_Tag] ), type nullable text ),

    // sinalizador do bug de origem, para ficar visível em vez de silencioso
    IdValido = Table.AddColumn(
        ComTagKey, "ID Válido",
        each [ID_Inspecao] <> null and not Text.Contains( Text.Lower( [ID_Inspecao] ), "yyyy" ),
        type logical
    )
in
    IdValido
