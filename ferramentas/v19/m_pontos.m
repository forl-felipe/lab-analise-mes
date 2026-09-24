let
    // Linhas fixas da lista "Pontos que pedem decisão" (Visão Geral).
    // Cada linha só aparece quando a medida [Ponto Ativo] = 1.
    Origem = #table(
        type table [ Chave = text, Ordem = Int64.Type, Tema = text ],
        {
            { "CAL", 1, "Calibração" }, { "AF1", 2, "Aferição · 1º ensaio" }, { "AF2", 3, "Aferição · 2º ensaio" },
            { "INS", 4, "Inspeções" }, { "PAR", 5, "Paradas" }, { "DISP", 6, "Disponibilidade" }
        }
    )
in
    Origem
