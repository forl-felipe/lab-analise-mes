// ─────────────────────────────────────────────────────────────────────────────
// CONSULTA: tbl_Config                                          (10 parâmetros)
// ─────────────────────────────────────────────────────────────────────────────
// Ocultar do painel de campos: serve de insumo para fx_Config e para as
// medidas de cabeçalho (Empresa / Unidade / Área), não para o usuário filtrar.

let
    Origem = fx_Tabela( "tbl_Config" ),
    Tipos  = Table.TransformColumnTypes( Origem, {{ "Parâmetro", type text }}, "pt-BR" ),
    Texto  = Table.TransformColumns( Tipos, {{ "Valor", each if _ = null then null else Text.From( _ ), type nullable text }} )
in
    Texto
