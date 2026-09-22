// ─────────────────────────────────────────────────────────────────────────────
// CONSULTA: tbl_Grupos                                          (10 linhas)
// ─────────────────────────────────────────────────────────────────────────────
// Tabela subaproveitada: carrega a regra de disponibilidade de cada grupo
// (Série · todos necessários  vs  Redundância parcial + Mín. em Operação) e
// hoje não aparece em nenhuma medida.
//
// Os 5 primeiros grupos são de planta e têm equipamento em tbl_Equipamentos.
// Os 5 últimos (Tambores, Peneiradores, Prensas, Analisadores, Espectrometria)
// são de laboratório e NÃO têm equipamento — correspondem aos instrumentos de
// tbl_Calibracao. A coluna "Domínio" torna essa separação explícita.

let
    Origem = fx_Tabela( "tbl_Grupos" ),

    SemColuna1 = Table.RemoveColumns( Origem, { "Coluna1" }, MissingField.Ignore ),

    Tipos = Table.TransformColumnTypes(
        SemColuna1,
        {
            { "Grupo Operacional",        type text },
            { "Área",                     type text },
            { "Categoria",                type text },
            { "Nº Equipamentos",          Int64.Type },
            { "Mín. em Operação",         Int64.Type },
            { "Regra",                    type text },
            { "Meta Disponibilidade (%)", Int64.Type }
        },
        "pt-BR"
    ),

    Dominio = Table.AddColumn(
        Tipos, "Domínio",
        each if [Regra] = "Série · todos necessários" then "Equipamentos de amostragem" else "Instrumentos de laboratório",
        type text
    ),

    // meta como fração, para formatar direto como percentual no visual
    MetaFracao = Table.AddColumn(
        Dominio, "Meta Disponibilidade",
        each if Record.Field( _, "Meta Disponibilidade (%)" ) = null then null
             else Record.Field( _, "Meta Disponibilidade (%)" ) / 100,
        type nullable number
    )
in
    MetaFracao
