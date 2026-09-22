// ─────────────────────────────────────────────────────────────────────────────
// CONSULTA: DimLaboratorio                                   (4 linhas · NOVA)
// ─────────────────────────────────────────────────────────────────────────────
// A ÚNICA dimensão compartilhada entre os dois domínios (D2).
//
//              LCP   LCE   LDP
//   Equipamentos  26     8     —
//   Calibração   119    57    32
//
// Com ela, o slicer de laboratório passa a filtrar os dois domínios ao mesmo
// tempo — algo que hoje é impossível, porque o slicer está preso a
// tbl_Calibracao[Laboratorio] e a relação com equipamentos não casa nenhuma linha.
//
// ⚠ AJUSTAR: os nomes por extenso abaixo são marcadores. Substituir pelos
//    nomes oficiais dos laboratórios.

let
    Origem = #table(
        type table [
            Laboratorio = text,
            #"Nome do Laboratório" = text,
            Ordem = Int64.Type
        ],
        {
            { "LCP", "LCP — ajustar nome oficial", 1 },
            { "LCE", "LCE — ajustar nome oficial", 2 },
            { "LDP", "LDP — ajustar nome oficial", 3 },
            { "N/A", "Não informado",              9 }
        }
    )
in
    Origem
