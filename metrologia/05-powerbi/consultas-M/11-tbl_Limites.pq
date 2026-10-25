// Substitui a consulta tbl_Limites do relatório atual.
let
    Lista = fnLista("Limites"),
    Colunas = Table.SelectColumns(Lista, {"Title", "Ensaio", "Parametro", "Unidade", "Nominal", "LimInferior", "LimSuperior", "Tolerancia", "Fonte"}, MissingField.UseNull),
    Tipos = Table.TransformColumnTypes(Colunas, {{"Nominal", type number}, {"LimInferior", type number}, {"LimSuperior", type number}, {"Tolerancia", type number}}),
    Renomeada = Table.RenameColumns(Tipos, {
        {"Title", "Chave"}, {"Parametro", "Parâmetro"}, {"LimInferior", "Lim. Inferior"},
        {"LimSuperior", "Lim. Superior"}, {"Tolerancia", "Tolerância"}
    })
in
    Renomeada
