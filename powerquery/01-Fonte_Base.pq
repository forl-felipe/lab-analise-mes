// ─────────────────────────────────────────────────────────────────────────────
// CONSULTA: Fonte_Base          ⚠ DESMARCAR "Habilitar carga"
// ─────────────────────────────────────────────────────────────────────────────
// Abre o arquivo UMA vez. Todas as demais consultas referenciam esta.
// Antes: 12 consultas abriam Excel.Workbook(Web.Contents(...)) separadamente,
// baixando o mesmo arquivo 12 vezes por atualização.

let
    Fonte = Excel.Workbook(
        Web.Contents( Parametro_CaminhoBase ),
        null,
        true
    )
in
    Fonte
