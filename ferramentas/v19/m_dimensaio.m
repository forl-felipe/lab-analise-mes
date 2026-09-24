let
    // Ensaios da planilha de calibração e a rotina de cada um (tirada do cabeçalho
    // de cada aba). Dias Semana: 1 = segunda ... 7 = domingo.
    // Por Dia = sessões por dia programado (Blaine: uma por turno).
    // Regra = texto da "cola" embaixo do mapa de aferições (pode ser editado aqui).
    // O DIA CERTO de cada ensaio sai de Dias Semana: mudar a rotina = mudar Dias Semana.
    Origem = #table(
        type table [ Ensaio = text, #"Ensaio Curto" = text, Ordem = Int64.Type, #"Frequência" = text,
                     #"Tem Rotina" = logical, Semanal = logical, #"Dias Semana" = text, #"Por Dia" = Int64.Type,
                     Regra = text ],
        {
            { "Tambor de Abrasão",        "Tambor de abrasão", 1, "toda segunda",          true,  false, "1",     1,
              "Toda segunda-feira, no turno da noite" },
            { "Verificação Peneiradores", "Peneiradores",      2, "ter, qui e sáb",        true,  false, "2,4,6", 1,
              "Terça, quinta e sábado" },
            { "Umidade",                  "Umidade",           3, "sem rotina definida",   false, false, "",      0,
              "Sem rotina definida: o mapa mostra só os dias em que foi feito" },
            { "Compressão",               "Compressão",        4, "toda segunda",          true,  false, "1",     1,
              "Toda segunda-feira" },
            { "Blaine",                   "Blaine",            5, "início de cada turno",  true,  false, "1,2,3,4,5,6,7", 2,
              "Todos os dias, no início de cada turno (2 por dia)" },
            { "Gran. Fina Alpine",        "Alpine",            6, "seg, qua e sáb",        true,  false, "1,3,6", 1,
              "Segunda, quarta e sábado" },
            { "Granulometria",            "Granulometria",     7, "1 por semana",          true,  true,  "1",     1,
              "1 por semana; o dia certo é a segunda-feira" },
            { "Tamb 5 kg x 15 kg",        "Tamb 5 × 15 kg",    8, "toda segunda",          true,  false, "1",     1,
              "Toda segunda-feira" },
            { "Comparativo Fisher",       "Fisher",            9, "sem rotina definida",   false, false, "",      0,
              "Sem rotina definida: o mapa mostra só os dias em que foi feito" }
        }
    )
in
    Origem
