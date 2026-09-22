// ─────────────────────────────────────────────────────────────────────────────
// PARÂMETRO: Parametro_CaminhoBase      (Tipo: Texto)
// ─────────────────────────────────────────────────────────────────────────────
// Criar via  Página Inicial ▸ Gerenciar Parâmetros ▸ Novo Parâmetro
//   Nome:            Parametro_CaminhoBase
//   Tipo:            Texto
//   Valor Atual:     a URL do .xlsx no SharePoint
//
// Por que parametrizar: hoje a URL está repetida em 12 consultas. Quando a base
// sair da pasta de teste (ZTESTE) para produção, é 1 alteração em vez de 12.
//
// Valor atual (pasta de TESTE — trocar ao publicar):
//   https://<tenant>.sharepoint.com/sites/automacao/laboratorios/ZTESTE/Base_PowerBI_Gestao_Equipamentos.xlsx
