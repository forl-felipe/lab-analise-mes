let
    // Ensaios da planilha de calibração e a rotina de cada um (tirada do cabeçalho
    // de cada aba). Dias Semana: 1 = segunda ... 7 = domingo.
    // Por Dia = sessões por dia programado (Blaine: uma por turno).
    Origem = #table(
        type table [ Ensaio = text, #"Ensaio Curto" = text, Ordem = Int64.Type, #"Frequência" = text,
                     #"Tem Rotina" = logical, Semanal = logical, #"Dias Semana" = text, #"Por Dia" = Int64.Type ],
        {
            { "Tambor de Abrasão",        "Tambor de abrasão", 1, "toda segunda",          true,  false, "1",     1 },
            { "Verificação Peneiradores", "Peneiradores",      2, "ter, qui e sáb",        true,  false, "2,4,6", 1 },
            { "Umidade",                  "Umidade",           3, "sem rotina definida",   false, false, "",      0 },
            { "Compressão",               "Compressão",        4, "toda segunda",          true,  false, "1",     1 },
            { "Blaine",                   "Blaine",            5, "início de cada turno",  true,  false, "1,2,3,4,5,6,7", 2 },
            { "Gran. Fina Alpine",        "Alpine",            6, "seg, qua e sáb",        true,  false, "1,3,6", 1 },
            { "Granulometria",            "Granulometria",     7, "1 por semana",          true,  true,  "",      1 },
            { "Tamb 5 kg x 15 kg",        "Tamb 5 × 15 kg",    8, "toda segunda",          true,  false, "1",     1 },
            { "Comparativo Fisher",       "Fisher",            9, "sem rotina definida",   false, false, "",      0 }
        }
    )
in
    Origem
