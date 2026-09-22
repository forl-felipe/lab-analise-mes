// ─────────────────────────────────────────────────────────────────────────────
// CONSULTA: tbl_Notas                                        (0 linhas hoje)
// ─────────────────────────────────────────────────────────────────────────────
// A aba "Notas Manutenção" está VAZIA na origem. Dois cards do painel
// (OMs Abertas e Lead Time Médio) dependem dela.
// A medida antiga usava  COUNTROWS(tbl_Notas) + 0  — o "+0" transformava
// "não tenho dado" em "zero OM aberta", que é uma afirmação diferente.
// A solução correta é o estado vazio no visual, não um zero inventado.
//
// Domínio válido de Status (vem de tbl_Listas):
//   Em andamento · Concluído · Aguardando peças · Cancelado
// Note que "Aberta" NÃO existe no domínio — ver D6 em docs/02-decisoes.

let
    Origem = fx_Tabela( "tbl_Notas" ),

    Tipos = Table.TransformColumnTypes(
        Origem,
        {
            { "OM",          type text },
            { "Disciplina",  type text },
            { "Atividade",   type text },
            { "Tag",         type text },
            { "Área",        type text },
            { "Status",      type text },
            { "Semana",      type text },
            { "Abertura",    type date },
            { "Atendimento", type date },
            { "Empresa",     type text },
            { "Site",        type text },
            { "Observação",  type text }
        },
        "pt-BR"
    ),

    ComTagKey = Table.AddColumn( Tipos, "TagKey", each fx_NormalizarTag( [Tag] ), type nullable text ),

    ComID = Table.AddIndexColumn( ComTagKey, "ID_Nota", 1, 1, Int64.Type ),

    LeadTime = Table.AddColumn(
        ComID, "LeadTimeDias",
        each if [Atendimento] = null or [Abertura] = null
             then null
             else Duration.Days( [Atendimento] - [Abertura] ),
        Int64.Type
    ),

    // agrupamento derivado: "aberta" = ainda demanda ação
    Agrupado = Table.AddColumn(
        LeadTime, "Situação",
        each
            let S = if [Status] = null then "" else Text.Lower( Text.Trim( [Status] ) )
            in  if S = "concluído" or S = "concluido" then "Concluída"
                else if S = "cancelado"               then "Cancelada"
                else if S = "em andamento"            then "Em andamento"
                else if S = "aguardando peças" or S = "aguardando pecas" then "Aguardando peças"
                else "Não informado",
        type text
    )
in
    Agrupado
