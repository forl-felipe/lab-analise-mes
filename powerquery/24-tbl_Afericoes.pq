// ─────────────────────────────────────────────────────────────────────────────
// tbl_Afericoes — aferições, calibrações e comparativos do laboratório
// ─────────────────────────────────────────────────────────────────────────────
// Fonte: as planilhas mensais "CALIBRAÇÃO MM - MÊS.xlsm" que os operadores já
// preenchem. Cada uma tem a aba OCULTA "BD_Afericoes" com a tabela
// tbl_Afericoes (formato longo: 1 linha = equipamento x ensaio x data), montada
// só com fórmulas que apontam para as células de sempre. Nada muda para o
// operador.
//
// A consulta lê a PASTA (não um arquivo): todo mês novo que entrar na pasta
// aparece sozinho no painel, com histórico acumulado.
//
// ⚠ PREENCHER as duas linhas marcadas com ►.

let
    // ► endereço do SITE (até /sites/<nome>), sem a pasta
    Site  = "https://smineracao.sharepoint.com/sites/COLOCAR_O_SITE",
    // ► um trecho do caminho da pasta onde ficam os arquivos mensais
    Pasta = "/COLOCAR_A_PASTA/",

    Arquivos  = SharePoint.Files( Site, [ ApiVersion = 15 ] ),
    DaPasta   = Table.SelectRows( Arquivos, each
                    Text.Contains( [Folder Path], Pasta, Comparer.OrdinalIgnoreCase )
                    and Text.EndsWith( Text.Lower( [Name] ), ".xlsm" )
                    and not Text.StartsWith( [Name], "~$" ) ),

    // arquivo sem a tabela (mês antigo, cópia de trabalho) é ignorado, não quebra
    ComDados  = Table.AddColumn( DaPasta, "Dados", each
                    let
                        wb = try Excel.Workbook( [Content], null, true ) otherwise null,
                        t  = if wb = null then null
                             else Table.SelectRows( wb, each [Kind] = "Table" and [Name] = "tbl_Afericoes" )
                    in
                        if t = null or Table.IsEmpty( t ) then null else t{0}[Data] ),
    Validos   = Table.SelectRows( ComDados, each [Dados] <> null ),
    Enxuto    = Table.SelectColumns( Validos, { "Name", "Dados" } ),
    Renomeado = Table.RenameColumns( Enxuto, { { "Name", "Arquivo" } } ),

    Colunas   = { "Ensaio", "Tipo", "Data", "Equipamento", "Parâmetro", "Unidade", "Valor",
                  "Referência", "Diferença", "Lim. Inferior", "Lim. Superior", "Tolerância",
                  "Resultado", "Responsável", "Letra", "Observação", "Origem" },
    Expandido = Table.ExpandTableColumn( Renomeado, "Dados", Colunas ),

    // fórmula que devolve "" chega como texto vazio: vira null antes de tipar
    Numericas = { "Data", "Valor", "Referência", "Diferença", "Lim. Inferior", "Lim. Superior", "Tolerância" },
    SemVazio  = Table.ReplaceValue( Expandido, "", null, Replacer.ReplaceValue, Numericas ),
    Tipado    = Table.TransformColumnTypes( SemVazio, {
                    { "Data", type date }, { "Valor", type number }, { "Referência", type number },
                    { "Diferença", type number }, { "Lim. Inferior", type number },
                    { "Lim. Superior", type number }, { "Tolerância", type number },
                    { "Ensaio", type text }, { "Tipo", type text }, { "Equipamento", type text },
                    { "Parâmetro", type text }, { "Unidade", type text }, { "Resultado", type text },
                    { "Responsável", type text }, { "Letra", type text }, { "Observação", type text },
                    { "Origem", type text }, { "Arquivo", type text } } ),
    SemErro   = Table.ReplaceErrorValues( Tipado, List.Transform( Numericas, each { _, null } ) ),

    // só o que foi de fato registrado: tem data e tem resultado
    Registros = Table.SelectRows( SemErro, each [Data] <> null and [Resultado] <> null and [Resultado] <> "" ),

    // o mesmo lançamento não entra duas vezes (arquivo copiado para o mês seguinte sem limpar):
    // Origem é a célula da aba de preenchimento, única dentro do arquivo
    Chave     = Table.AddColumn( Registros, "Chave", each
                    [Origem] & "|" & [Parâmetro] & "|" & Date.ToText( [Data], "yyyy-MM-dd" ), type text ),
    Unicos    = Table.Distinct( Chave, { "Chave" } ),

    Conforme  = Table.AddColumn( Unicos, "Conforme", each if [Resultado] = "Não conforme" then 0 else if [Resultado] = "Conforme" then 1 else null, Int64.Type )
in
    Conforme
