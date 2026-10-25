// Função auxiliar: lê uma lista do SharePoint e devolve as colunas pelo nome interno.
// Nome da consulta: fnLista
(nomeLista as text) as table =>
let
    Fonte = SharePoint.Tables(SiteUrl, [Implementation = "2.0", ViewMode = "All"]),
    Itens = Fonte{[Title = nomeLista]}[Items],
    // Em sites em português a coluna Título pode chegar como "Título"
    Padronizada = Table.RenameColumns(Itens, {{"Título", "Title"}}, MissingField.Ignore)
in
    Padronizada
