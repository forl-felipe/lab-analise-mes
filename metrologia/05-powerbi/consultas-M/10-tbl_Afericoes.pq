// Substitui a consulta tbl_Afericoes do relatório atual (Editor Avançado > colar).
// Mantém o NOME e as COLUNAS da tabela da planilha, então gráficos e medidas continuam funcionando.
let
    Lista = fnLista("Afericoes"),
    Colunas = Table.SelectColumns(Lista, {
        "Ensaio", "TipoRegistro", "DataEnsaio", "Equipamento", "TAG", "Parametro", "Unidade", "Valor", "Referencia",
        "Diferenca", "LimInferior", "LimSuperior", "Tolerancia", "Resultado", "Responsavel", "ResponsavelEmail", "Letra",
        "Observacao", "Causa", "ItensNaoOK", "Title", "Turno", "RegistradoEm", "TurnoRegistro", "ForaDoTurno", "Origem", "ModoLeitura"
    }, MissingField.UseNull),
    Datas = Table.TransformColumns(Colunas, {{"DataEnsaio", fnData, type date}}),
    Tipos = Table.TransformColumnTypes(Datas, {
        {"Valor", type number}, {"Referencia", type number}, {"Diferenca", type number},
        {"LimInferior", type number}, {"LimSuperior", type number}, {"Tolerancia", type number},
        {"RegistradoEm", type datetimezone}
    }),
    Local = Table.TransformColumns(Tipos, {{"RegistradoEm", each if _ = null then null else DateTimeZone.RemoveZone(DateTimeZone.SwitchZone(_, -3)), type datetime}}),
    Renomeada = Table.RenameColumns(Local, {
        {"TipoRegistro", "Tipo"}, {"DataEnsaio", "Data"}, {"Parametro", "Parâmetro"}, {"Referencia", "Referência"},
        {"Diferenca", "Diferença"}, {"LimInferior", "Lim. Inferior"}, {"LimSuperior", "Lim. Superior"},
        {"Tolerancia", "Tolerância"}, {"Responsavel", "Responsável"}, {"Observacao", "Observação"},
        {"Title", "Lançamento"}, {"RegistradoEm", "Registrado em"}, {"TurnoRegistro", "Turno do registro"},
        {"ForaDoTurno", "Fora do turno"}, {"Origem", "Fonte"}, {"ModoLeitura", "Modo de leitura"}
    }),
    // Coluna "Origem" no formato da planilha (código do lançamento | turno)
    ComOrigem = Table.AddColumn(Renomeada, "Origem", each [Lançamento] & " | " & ([Turno] ?? ""), type text)
in
    ComOrigem
