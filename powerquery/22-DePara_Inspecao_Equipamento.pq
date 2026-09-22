// ─────────────────────────────────────────────────────────────────────────────
// CONSULTA: DePara_Inspecao_Equipamento                     (5 linhas · NOVA)
// ─────────────────────────────────────────────────────────────────────────────
// Liga o domínio das INSPEÇÕES (tags AM-XX-NN) ao mestre de EQUIPAMENTOS
// (tags U03-*/U04-*/U00-*). Hoje a interseção entre os dois é ZERO, e por isso
// nenhuma inspeção consegue ser analisada por grupo, área ou laboratório.
//
// São só 5 tags — por isso um de-para manual resolve, e é mais barato que
// qualquer outra alternativa.
//
// ⚠ AÇÃO NECESSÁRIA: preencher a coluna TagEquipamento.
//    · Os 2 cortadores têm candidatos no mestre, mas o número não fecha:
//      a inspeção diz "PQ 01" e o cadastro só tem PQ 03 e PQ 04.
//    · Os 3 últimos NÃO EXISTEM no cadastro de equipamentos. Cross Belt,
//      Bomba de Polpa e Divisor Rotativo precisam ser cadastrados primeiro —
//      o mestre de 34 só contém cortadores e correias.
//
// AÇÃO DE FUNDO: a lista suspensa de equipamento no formulário de inspeção
// precisa ser alimentada pelo cadastro mestre. Sem isso, este de-para cresce
// a cada equipamento novo e volta a dessincronizar.

let
    Origem = #table(
        type table [
            TagInspecao      = text,
            DescricaoOrigem  = text,
            TagEquipamento   = text,   // ← PREENCHER
            Confianca        = text
        ],
        {
            { "AM-CT-01", "Cortador Primario PQ 01",      "",  "Candidatos: U03-07CR002 (PQ 03) ou U04-07CR001 (PQ 04) — CONFIRMAR" },
            { "AM-CT-03", "Cortador Secundario PQ 01",    "",  "Candidatos: U03-07CR005 (PQ 03) ou U04-07TP018 (PQ 04) — CONFIRMAR" },
            { "AM-CB-01", "Amostrador Cross Belt 01",     "",  "AUSENTE no cadastro mestre — cadastrar equipamento" },
            { "AM-BP-01", "Bomba de Polpa Amostragem 01", "",  "AUSENTE no cadastro mestre — cadastrar equipamento" },
            { "AM-DR-01", "Divisor Rotativo 01",          "",  "AUSENTE no cadastro mestre — cadastrar equipamento" }
        }
    ),

    ChaveInspecao    = Table.AddColumn( Origem,        "TagKey Inspecao", each fx_NormalizarTag( [TagInspecao] ),    type nullable text ),
    ChaveEquipamento = Table.AddColumn( ChaveInspecao, "TagKey",          each fx_NormalizarTag( [TagEquipamento] ), type nullable text ),

    Resolvido = Table.AddColumn(
        ChaveEquipamento, "Resolvido",
        each [TagKey] <> null,
        type logical
    )
in
    Resolvido
