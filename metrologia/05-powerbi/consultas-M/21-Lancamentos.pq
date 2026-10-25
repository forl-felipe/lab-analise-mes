// Nova consulta: Lancamentos (um registro por envio)
let
    Lista = fnLista("Lancamentos"),
    Colunas = Table.SelectColumns(Lista, {"Title", "Ensaio", "DataEnsaio", "Turno", "Letra", "Responsavel", "QtdResultados", "QtdNaoConformes", "TAGsNaoConformes", "Causa", "Observacao", "RegistradoEm", "ForaDoTurno", "Situacao", "Origem"}, MissingField.UseNull),
    Datas = Table.TransformColumns(Colunas, {{"DataEnsaio", fnData, type date}}),
    Ativos = Table.SelectRows(Datas, each [Situacao] <> "Desfeito"),
    Tipos = Table.TransformColumnTypes(Ativos, {{"QtdResultados", Int64.Type}, {"QtdNaoConformes", Int64.Type}}),
    Renomeada = Table.RenameColumns(Tipos, {{"Title", "Lançamento"}, {"DataEnsaio", "Data"}}),
    ComChave = Table.AddColumn(Renomeada, "Chave", each [Ensaio] & "|" & Date.ToText([Data], "yyyy-MM-dd") & "|" & [Turno], type text)
in
    ComChave
