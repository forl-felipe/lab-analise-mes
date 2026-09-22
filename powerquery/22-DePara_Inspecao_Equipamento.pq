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
// STATUS: respondido pelo responsável em 22/09/2026.
//   · Os dois cortadores são PQ 03 e PQ 04 — mapeados abaixo.
//   · Cross Belt, Bomba de Polpa e Divisor Rotativo ENTRAM no cadastro mestre.
//     Enquanto não forem cadastrados em tbl_Equipamentos, o de-para os mantém
//     com TagEquipamento vazio e a coluna [Resolvido] = FALSE — o painel mostra
//     quantos faltam em vez de fingir que estão ligados.
//
// ⚠ CONFIRMAR uma leitura: "os cortadores são PQ3 e PQ4" foi lido como
//    AM-CT-01 (primário)  -> PQ 03
//    AM-CT-03 (secundário) -> PQ 04
//   Se a intenção era os dois na MESMA usina, trocar U04-07TP018 por
//   U03-07CR005 (Cortador Secundário PQ 03). É uma célula.

let
    Origem = #table(
        type table [
            TagInspecao      = text,
            DescricaoOrigem  = text,
            TagEquipamento   = text,   // ← PREENCHER
            Confianca        = text
        ],
        {
            { "AM-CT-01", "Cortador Primario PQ 01",      "U03-07CR002", "Cortador Primário PQ 03 — confirmar leitura" },
            { "AM-CT-03", "Cortador Secundario PQ 01",    "U04-07TP018", "Cortador Secundário PQ 04 — confirmar leitura" },
            { "AM-CB-01", "Amostrador Cross Belt 01",     "",            "A cadastrar em tbl_Equipamentos" },
            { "AM-BP-01", "Bomba de Polpa Amostragem 01", "",            "A cadastrar em tbl_Equipamentos" },
            { "AM-DR-01", "Divisor Rotativo 01",          "",            "A cadastrar em tbl_Equipamentos" }
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
