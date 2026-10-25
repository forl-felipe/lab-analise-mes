// Nova consulta: Equipamentos (situação de calibração e aferição)
let
    Lista = fnLista("Equipamentos"),
    Colunas = Table.SelectColumns(Lista, {
        "Title", "Descricao", "TipoEquipamento", "Local", "Grupo", "EnsaioAfericao", "ControlaCalibracao", "PeriodicidadeMeses",
        "UltimaCalibracao", "ProximaCalibracao", "UltimoResultado", "UltimoResponsavel", "ManutencaoSolicitada",
        "UltimaAfericao", "UltimaAfericaoResultado", "Status", "Ativo"
    }, MissingField.UseNull),
    Datas = Table.TransformColumns(Colunas, {{"UltimaCalibracao", fnData, type date}, {"ProximaCalibracao", fnData, type date}, {"UltimaAfericao", fnData, type date}}),
    Ativos = Table.SelectRows(Datas, each [Ativo] = true),
    Renomeada = Table.RenameColumns(Ativos, {{"Title", "TAG"}}),
    // Situação resumida para gráficos: Liberado / Bloqueado
    ComSituacao = Table.AddColumn(Renomeada, "Situação", each if Text.StartsWith([Status] ?? "", "Bloqueado") then "Bloqueado" else "Liberado", type text),
    ComMotivo = Table.AddColumn(ComSituacao, "Motivo", each let p = Text.Split([Status] ?? "", "·") in if List.Count(p) > 1 then Text.Trim(p{1}) else "Em dia", type text)
in
    ComMotivo
