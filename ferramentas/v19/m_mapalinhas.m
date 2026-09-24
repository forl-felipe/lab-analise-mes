let
    // Linhas do mapa de aferições: a 1ª linha mostra o dia da semana (seg, ter, ... e "hoje");
    // as demais são os ensaios de DimEnsaio (mesma ordem e mesma regra).
    // Tabela separada de DimEnsaio: clicar num ensaio do mapa filtra só a "cola" de regras.
    Ensaios = Table.RenameColumns(
                  Table.SelectColumns( DimEnsaio, { "Ensaio", "Ensaio Curto", "Ordem", "Regra" } ),
                  { { "Ensaio Curto", "Linha" } } ),
    DiaSemana = #table(
                    type table [ Ensaio = text, Linha = text, Ordem = Int64.Type, Regra = text ],
                    { { "", "", 0, "" } } ),
    Linhas = Table.Combine( { DiaSemana, Ensaios } ),
    Tipado = Table.TransformColumnTypes( Linhas, { { "Ensaio", type text }, { "Linha", type text },
                                                   { "Ordem", Int64.Type }, { "Regra", type text } } )
in
    Tipado
