// ─────────────────────────────────────────────────────────────────────────────
// CONSULTA: Tbl_Paradas                              (1 linha com conteúdo)
// ─────────────────────────────────────────────────────────────────────────────
// Nome mantido com T maiúsculo para não quebrar os visuais existentes.
// Padronizar para tbl_Paradas na Fase 3 (ver D10).
//
// As colunas "(auto)" — Equipamento (auto), Grupo (auto), Mês (auto) — são
// cópias desnormalizadas do que deveria vir de tbl_Equipamentos. Mantidas por
// compatibilidade com os visuais atuais; devem sair quando as páginas forem
// reconstruídas usando a dimensão.

let
    Origem = fx_Tabela( "Tbl_Paradas" ),

    Tipos = Table.TransformColumnTypes(
        Origem,
        {
            { "Data",               type date },
            { "Equip. (Tag)",       type text },
            { "Horas Parado",       type number },
            { "Ocorrência",         type text },
            { "Responsável",        type text },
            { "Situação",           type text },
            { "Equipamento (auto)", type text },
            { "Grupo (auto)",       type text }
        },
        "pt-BR"
    ),

    ComTagKey = Table.AddColumn( Tipos, "TagKey", each fx_NormalizarTag( Record.Field( _, "Equip. (Tag)" ) ), type nullable text ),
    ComID     = Table.AddIndexColumn( ComTagKey, "ID_Parada", 1, 1, Int64.Type ),
    ComMesAno = Table.AddColumn( ComID, "MesAno", each if [Data] = null then null else Date.ToText( [Data], "yyyy-MM" ), type nullable text )
in
    ComMesAno
