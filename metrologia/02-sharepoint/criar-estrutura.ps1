<#
  SISTEMA METROLOGIA · criação da estrutura no SharePoint (gerado por ferramentas/gerar_pnp.py)

  O que faz:
    1. Cria as listas que faltam e acrescenta as colunas novas nas listas da versão 1
       (Equipamentos e Calibracoes), sem apagar nada.
    2. Cria índices (necessários para a lista passar de 5.000 itens sem travar o app).
    3. Cria os grupos 'Metrologia - Laboratoristas' e 'Metrologia - Instrumentação' e
       aplica as permissões de cada lista.
    4. (Opcional, -ImportarDados) importa a carga inicial dos CSVs gerados por gerar_carga.py.

  Pré-requisitos: PowerShell 7+, módulo PnP.PowerShell e um App Registration do Entra ID
  (Register-PnPEntraIDAppForInteractiveLogin). Peça à TI se não tiver.

  Uso:
    ./criar-estrutura.ps1 -SiteUrl https://EMPRESA.sharepoint.com/sites/SITE -ClientId <guid>
    ./criar-estrutura.ps1 -SiteUrl ... -ClientId ... -ImportarDados -PastaCsv ..\dados\csv
#>
param(
    [Parameter(Mandatory = $true)] [string] $SiteUrl,
    [Parameter(Mandatory = $true)] [string] $ClientId,
    [switch] $ImportarDados,
    [string] $PastaCsv = (Join-Path $PSScriptRoot 'csv')
)
$ErrorActionPreference = 'Stop'
Import-Module PnP.PowerShell
Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId
$inv = [System.Globalization.CultureInfo]::InvariantCulture

function Garantir-Lista([string] $Nome, [string] $Descricao) {
    if (-not (Get-PnPList -Identity $Nome -ErrorAction SilentlyContinue)) {
        New-PnPList -Title $Nome -Template GenericList -OnQuickLaunch | Out-Null
        Write-Host "Lista criada: $Nome" -ForegroundColor Green
    } else { Write-Host "Lista já existe: $Nome" -ForegroundColor Yellow }
    Set-PnPList -Identity $Nome -Description $Descricao -EnableVersioning $true | Out-Null
}
function Garantir-Coluna([string] $Lista, [string] $Nome, [string] $Xml) {
    if (Get-PnPField -List $Lista -Identity $Nome -ErrorAction SilentlyContinue) { return }
    $x = $Xml.Replace("Name='$Nome'", "ID='{" + [guid]::NewGuid().ToString() + "}' Name='$Nome'")
    Add-PnPFieldFromXml -List $Lista -FieldXml $x | Out-Null
    Write-Host "   + $Lista.$Nome"
}
function Indexar([string] $Lista, [string] $Nome) {
    Set-PnPField -List $Lista -Identity $Nome -Values @{ Indexed = $true } | Out-Null
}

# ---------------------------------------------------------------- Acessos
Garantir-Lista 'Acessos' 'Define quem é Laboratorista e quem é Instrumentação. Quem não está aqui (ou está inativo) entra como Colaborador, somente consulta.'
Garantir-Coluna 'Acessos' 'Pessoa' '<Field Type=''User'' UserSelectionMode=''PeopleOnly'' UserSelectionScope=''0'' ShowField=''ImnName'' Name=''Pessoa'' StaticName=''Pessoa'' DisplayName=''Pessoa'' Required=''TRUE'' />'
Garantir-Coluna 'Acessos' 'Perfil' '<Field Type=''Choice'' Format=''Dropdown'' FillInChoice=''FALSE'' Name=''Perfil'' StaticName=''Perfil'' DisplayName=''Perfil'' Required=''TRUE''><CHOICES><CHOICE>Laboratorista</CHOICE><CHOICE>Instrumentação</CHOICE></CHOICES></Field>'
Garantir-Coluna 'Acessos' 'Letra' '<Field Type=''Choice'' Format=''Dropdown'' FillInChoice=''FALSE'' Name=''Letra'' StaticName=''Letra'' DisplayName=''Letra''><CHOICES><CHOICE>A</CHOICE><CHOICE>B</CHOICE><CHOICE>C</CHOICE><CHOICE>D</CHOICE><CHOICE>ADM</CHOICE></CHOICES></Field>'
Garantir-Coluna 'Acessos' 'Ativo' '<Field Type=''Boolean'' Name=''Ativo'' StaticName=''Ativo'' DisplayName=''Ativo''><Default>1</Default></Field>'
Garantir-Coluna 'Acessos' 'Observacao' '<Field Type=''Text'' Name=''Observacao'' StaticName=''Observacao'' DisplayName=''Observacao'' MaxLength=''255''></Field>'
$v = Get-PnPView -List 'Acessos' | Where-Object { $_.DefaultView }
Set-PnPView -List 'Acessos' -Identity $v.Id -Fields @('LinkTitle', 'Pessoa', 'Perfil', 'Letra', 'Ativo', 'Observacao') | Out-Null

# ---------------------------------------------------------------- Equipamentos
Garantir-Lista 'Equipamentos' 'Cadastro único de equipamentos: calibração (instrumentação) e situação da última aferição (laboratório).'
Garantir-Coluna 'Equipamentos' 'Descricao' '<Field Type=''Text'' Name=''Descricao'' StaticName=''Descricao'' DisplayName=''Descricao'' MaxLength=''255''></Field>'
Garantir-Coluna 'Equipamentos' 'TipoEquipamento' '<Field Type=''Text'' Name=''TipoEquipamento'' StaticName=''TipoEquipamento'' DisplayName=''TipoEquipamento'' MaxLength=''255''></Field>'
Garantir-Coluna 'Equipamentos' 'FabricanteModelo' '<Field Type=''Text'' Name=''FabricanteModelo'' StaticName=''FabricanteModelo'' DisplayName=''FabricanteModelo'' MaxLength=''255''></Field>'
Garantir-Coluna 'Equipamentos' 'NumeroSerie' '<Field Type=''Text'' Name=''NumeroSerie'' StaticName=''NumeroSerie'' DisplayName=''NumeroSerie'' MaxLength=''255''></Field>'
Garantir-Coluna 'Equipamentos' 'Local' '<Field Type=''Text'' Name=''Local'' StaticName=''Local'' DisplayName=''Local'' MaxLength=''255''></Field>'
Garantir-Coluna 'Equipamentos' 'Grupo' '<Field Type=''Text'' Name=''Grupo'' StaticName=''Grupo'' DisplayName=''Grupo'' MaxLength=''255''></Field>'
Garantir-Coluna 'Equipamentos' 'Conjunto' '<Field Type=''Text'' Name=''Conjunto'' StaticName=''Conjunto'' DisplayName=''Conjunto'' MaxLength=''255''></Field>'
Garantir-Coluna 'Equipamentos' 'EnsaioAfericao' '<Field Type=''Text'' Name=''EnsaioAfericao'' StaticName=''EnsaioAfericao'' DisplayName=''EnsaioAfericao'' MaxLength=''255''></Field>'
Garantir-Coluna 'Equipamentos' 'ControlaCalibracao' '<Field Type=''Boolean'' Name=''ControlaCalibracao'' StaticName=''ControlaCalibracao'' DisplayName=''ControlaCalibracao''><Default>1</Default></Field>'
Garantir-Coluna 'Equipamentos' 'PeriodicidadeMeses' '<Field Type=''Number'' Name=''PeriodicidadeMeses'' StaticName=''PeriodicidadeMeses'' DisplayName=''PeriodicidadeMeses''><Default>12</Default></Field>'
Garantir-Coluna 'Equipamentos' 'UltimaCalibracao' '<Field Type=''DateTime'' Format=''DateOnly'' Name=''UltimaCalibracao'' StaticName=''UltimaCalibracao'' DisplayName=''UltimaCalibracao'' />'
Garantir-Coluna 'Equipamentos' 'ProximaCalibracao' '<Field Type=''DateTime'' Format=''DateOnly'' Name=''ProximaCalibracao'' StaticName=''ProximaCalibracao'' DisplayName=''ProximaCalibracao'' />'
Garantir-Coluna 'Equipamentos' 'UltimoResultado' '<Field Type=''Text'' Name=''UltimoResultado'' StaticName=''UltimoResultado'' DisplayName=''UltimoResultado'' MaxLength=''255''></Field>'
Garantir-Coluna 'Equipamentos' 'UltimoResponsavel' '<Field Type=''Text'' Name=''UltimoResponsavel'' StaticName=''UltimoResponsavel'' DisplayName=''UltimoResponsavel'' MaxLength=''255''></Field>'
Garantir-Coluna 'Equipamentos' 'MotivoRecusa' '<Field Type=''Note'' Name=''MotivoRecusa'' StaticName=''MotivoRecusa'' DisplayName=''MotivoRecusa'' NumLines=''6'' RichText=''FALSE'' />'
Garantir-Coluna 'Equipamentos' 'ManutencaoSolicitada' '<Field Type=''Boolean'' Name=''ManutencaoSolicitada'' StaticName=''ManutencaoSolicitada'' DisplayName=''ManutencaoSolicitada''><Default>0</Default></Field>'
Garantir-Coluna 'Equipamentos' 'UltimaAfericao' '<Field Type=''DateTime'' Format=''DateOnly'' Name=''UltimaAfericao'' StaticName=''UltimaAfericao'' DisplayName=''UltimaAfericao'' />'
Garantir-Coluna 'Equipamentos' 'UltimaAfericaoResultado' '<Field Type=''Text'' Name=''UltimaAfericaoResultado'' StaticName=''UltimaAfericaoResultado'' DisplayName=''UltimaAfericaoResultado'' MaxLength=''255''></Field>'
Garantir-Coluna 'Equipamentos' 'UltimaAfericaoLancamento' '<Field Type=''Text'' Name=''UltimaAfericaoLancamento'' StaticName=''UltimaAfericaoLancamento'' DisplayName=''UltimaAfericaoLancamento'' MaxLength=''255''></Field>'
Garantir-Coluna 'Equipamentos' 'Status' '<Field Type=''Text'' Name=''Status'' StaticName=''Status'' DisplayName=''Status'' MaxLength=''255''></Field>'
Garantir-Coluna 'Equipamentos' 'ObservacaoCadastro' '<Field Type=''Note'' Name=''ObservacaoCadastro'' StaticName=''ObservacaoCadastro'' DisplayName=''ObservacaoCadastro'' NumLines=''6'' RichText=''FALSE'' />'
Garantir-Coluna 'Equipamentos' 'Ativo' '<Field Type=''Boolean'' Name=''Ativo'' StaticName=''Ativo'' DisplayName=''Ativo''><Default>1</Default></Field>'
Set-PnPField -List 'Equipamentos' -Identity 'Title' -Values @{ Indexed = $true; EnforceUniqueValues = $true } | Out-Null
Indexar 'Equipamentos' 'Conjunto'
$v = Get-PnPView -List 'Equipamentos' | Where-Object { $_.DefaultView }
Set-PnPView -List 'Equipamentos' -Identity $v.Id -Fields @('LinkTitle', 'Descricao', 'TipoEquipamento', 'FabricanteModelo', 'NumeroSerie', 'Local', 'Grupo', 'Conjunto', 'EnsaioAfericao', 'ControlaCalibracao', 'PeriodicidadeMeses', 'UltimaCalibracao', 'ProximaCalibracao') | Out-Null

# ---------------------------------------------------------------- Ensaios
Garantir-Lista 'Ensaios' 'Catálogo dos ensaios e agenda semanal (mesma grade da aba Config da planilha).'
Garantir-Coluna 'Ensaios' 'Nome' '<Field Type=''Text'' Name=''Nome'' StaticName=''Nome'' DisplayName=''Nome'' MaxLength=''255'' Required=''TRUE''></Field>'
Garantir-Coluna 'Ensaios' 'NomeBD' '<Field Type=''Text'' Name=''NomeBD'' StaticName=''NomeBD'' DisplayName=''NomeBD'' MaxLength=''255'' Required=''TRUE''></Field>'
Garantir-Coluna 'Ensaios' 'Formulario' '<Field Type=''Choice'' Format=''Dropdown'' FillInChoice=''FALSE'' Name=''Formulario'' StaticName=''Formulario'' DisplayName=''Formulario''><CHOICES><CHOICE>Medição</CHOICE><CHOICE>Checklist</CHOICE></CHOICES><Default>Medição</Default></Field>'
Garantir-Coluna 'Ensaios' 'Frequencia' '<Field Type=''Text'' Name=''Frequencia'' StaticName=''Frequencia'' DisplayName=''Frequencia'' MaxLength=''255''></Field>'
Garantir-Coluna 'Ensaios' 'Ordem' '<Field Type=''Number'' Name=''Ordem'' StaticName=''Ordem'' DisplayName=''Ordem''></Field>'
Garantir-Coluna 'Ensaios' 'Instrucoes' '<Field Type=''Note'' Name=''Instrucoes'' StaticName=''Instrucoes'' DisplayName=''Instrucoes'' NumLines=''6'' RichText=''FALSE'' />'
Garantir-Coluna 'Ensaios' 'Ativo' '<Field Type=''Boolean'' Name=''Ativo'' StaticName=''Ativo'' DisplayName=''Ativo''><Default>1</Default></Field>'
Garantir-Coluna 'Ensaios' 'SegD' '<Field Type=''Boolean'' Name=''SegD'' StaticName=''SegD'' DisplayName=''SegD''><Default>0</Default></Field>'
Garantir-Coluna 'Ensaios' 'SegN' '<Field Type=''Boolean'' Name=''SegN'' StaticName=''SegN'' DisplayName=''SegN''><Default>0</Default></Field>'
Garantir-Coluna 'Ensaios' 'TerD' '<Field Type=''Boolean'' Name=''TerD'' StaticName=''TerD'' DisplayName=''TerD''><Default>0</Default></Field>'
Garantir-Coluna 'Ensaios' 'TerN' '<Field Type=''Boolean'' Name=''TerN'' StaticName=''TerN'' DisplayName=''TerN''><Default>0</Default></Field>'
Garantir-Coluna 'Ensaios' 'QuaD' '<Field Type=''Boolean'' Name=''QuaD'' StaticName=''QuaD'' DisplayName=''QuaD''><Default>0</Default></Field>'
Garantir-Coluna 'Ensaios' 'QuaN' '<Field Type=''Boolean'' Name=''QuaN'' StaticName=''QuaN'' DisplayName=''QuaN''><Default>0</Default></Field>'
Garantir-Coluna 'Ensaios' 'QuiD' '<Field Type=''Boolean'' Name=''QuiD'' StaticName=''QuiD'' DisplayName=''QuiD''><Default>0</Default></Field>'
Garantir-Coluna 'Ensaios' 'QuiN' '<Field Type=''Boolean'' Name=''QuiN'' StaticName=''QuiN'' DisplayName=''QuiN''><Default>0</Default></Field>'
Garantir-Coluna 'Ensaios' 'SexD' '<Field Type=''Boolean'' Name=''SexD'' StaticName=''SexD'' DisplayName=''SexD''><Default>0</Default></Field>'
Garantir-Coluna 'Ensaios' 'SexN' '<Field Type=''Boolean'' Name=''SexN'' StaticName=''SexN'' DisplayName=''SexN''><Default>0</Default></Field>'
Garantir-Coluna 'Ensaios' 'SabD' '<Field Type=''Boolean'' Name=''SabD'' StaticName=''SabD'' DisplayName=''SabD''><Default>0</Default></Field>'
Garantir-Coluna 'Ensaios' 'SabN' '<Field Type=''Boolean'' Name=''SabN'' StaticName=''SabN'' DisplayName=''SabN''><Default>0</Default></Field>'
Garantir-Coluna 'Ensaios' 'DomD' '<Field Type=''Boolean'' Name=''DomD'' StaticName=''DomD'' DisplayName=''DomD''><Default>0</Default></Field>'
Garantir-Coluna 'Ensaios' 'DomN' '<Field Type=''Boolean'' Name=''DomN'' StaticName=''DomN'' DisplayName=''DomN''><Default>0</Default></Field>'
Set-PnPField -List 'Ensaios' -Identity 'Title' -Values @{ Indexed = $true; EnforceUniqueValues = $true } | Out-Null
$v = Get-PnPView -List 'Ensaios' | Where-Object { $_.DefaultView }
Set-PnPView -List 'Ensaios' -Identity $v.Id -Fields @('LinkTitle', 'Nome', 'NomeBD', 'Formulario', 'Frequencia', 'Ordem', 'Instrucoes', 'Ativo', 'SegD', 'SegN', 'TerD', 'TerN', 'QuaD') | Out-Null

# ---------------------------------------------------------------- PontosAfericao
Garantir-Lista 'PontosAfericao' 'Cada linha é um campo do formulário de um ensaio: qual equipamento, qual parâmetro, que tipo de entrada e qual critério usar. É o equivalente às linhas da aba Lançamento. Nos peneiradores, cada linha é uma POSIÇÃO (peneirador + malha) e a coluna TAG diz qual peneira está nela.'
Garantir-Coluna 'PontosAfericao' 'Ensaio' '<Field Type=''Text'' Name=''Ensaio'' StaticName=''Ensaio'' DisplayName=''Ensaio'' MaxLength=''255'' Required=''TRUE''></Field>'
Garantir-Coluna 'PontosAfericao' 'Ordem' '<Field Type=''Number'' Name=''Ordem'' StaticName=''Ordem'' DisplayName=''Ordem''></Field>'
Garantir-Coluna 'PontosAfericao' 'Secao' '<Field Type=''Text'' Name=''Secao'' StaticName=''Secao'' DisplayName=''Secao'' MaxLength=''255''></Field>'
Garantir-Coluna 'PontosAfericao' 'Conjunto' '<Field Type=''Text'' Name=''Conjunto'' StaticName=''Conjunto'' DisplayName=''Conjunto'' MaxLength=''255''></Field>'
Garantir-Coluna 'PontosAfericao' 'TAG' '<Field Type=''Text'' Name=''TAG'' StaticName=''TAG'' DisplayName=''TAG'' MaxLength=''255''></Field>'
Garantir-Coluna 'PontosAfericao' 'EquipamentoBD' '<Field Type=''Text'' Name=''EquipamentoBD'' StaticName=''EquipamentoBD'' DisplayName=''EquipamentoBD'' MaxLength=''255''></Field>'
Garantir-Coluna 'PontosAfericao' 'Parametro' '<Field Type=''Text'' Name=''Parametro'' StaticName=''Parametro'' DisplayName=''Parametro'' MaxLength=''255'' Required=''TRUE''></Field>'
Garantir-Coluna 'PontosAfericao' 'Unidade' '<Field Type=''Text'' Name=''Unidade'' StaticName=''Unidade'' DisplayName=''Unidade'' MaxLength=''255''></Field>'
Garantir-Coluna 'PontosAfericao' 'Entrada' '<Field Type=''Choice'' Format=''Dropdown'' FillInChoice=''FALSE'' Name=''Entrada'' StaticName=''Entrada'' DisplayName=''Entrada'' Required=''TRUE''><CHOICES><CHOICE>Valor</CHOICE><CHOICE>Tempo</CHOICE><CHOICE>Comparativo</CHOICE><CHOICE>ComparativoReferencia</CHOICE><CHOICE>Massa</CHOICE><CHOICE>Checklist</CHOICE></CHOICES></Field>'
Garantir-Coluna 'PontosAfericao' 'ChaveLimite' '<Field Type=''Text'' Name=''ChaveLimite'' StaticName=''ChaveLimite'' DisplayName=''ChaveLimite'' MaxLength=''255''></Field>'
Garantir-Coluna 'PontosAfericao' 'Fator' '<Field Type=''Number'' Name=''Fator'' StaticName=''Fator'' DisplayName=''Fator''></Field>'
Garantir-Coluna 'PontosAfericao' 'RotuloA' '<Field Type=''Text'' Name=''RotuloA'' StaticName=''RotuloA'' DisplayName=''RotuloA'' MaxLength=''255''></Field>'
Garantir-Coluna 'PontosAfericao' 'RotuloB' '<Field Type=''Text'' Name=''RotuloB'' StaticName=''RotuloB'' DisplayName=''RotuloB'' MaxLength=''255''></Field>'
Garantir-Coluna 'PontosAfericao' 'TipoRegistro' '<Field Type=''Choice'' Format=''Dropdown'' FillInChoice=''FALSE'' Name=''TipoRegistro'' StaticName=''TipoRegistro'' DisplayName=''TipoRegistro''><CHOICES><CHOICE>Calibração</CHOICE><CHOICE>Verificação</CHOICE><CHOICE>Comparativo</CHOICE></CHOICES></Field>'
Garantir-Coluna 'PontosAfericao' 'Ativo' '<Field Type=''Boolean'' Name=''Ativo'' StaticName=''Ativo'' DisplayName=''Ativo''><Default>1</Default></Field>'
Set-PnPField -List 'PontosAfericao' -Identity 'Title' -Values @{ Indexed = $true; EnforceUniqueValues = $true } | Out-Null
Indexar 'PontosAfericao' 'Ensaio'
$v = Get-PnPView -List 'PontosAfericao' | Where-Object { $_.DefaultView }
Set-PnPView -List 'PontosAfericao' -Identity $v.Id -Fields @('LinkTitle', 'Ensaio', 'Ordem', 'Secao', 'Conjunto', 'TAG', 'EquipamentoBD', 'Parametro', 'Unidade', 'Entrada', 'ChaveLimite', 'Fator', 'RotuloA') | Out-Null

# ---------------------------------------------------------------- Limites
Garantir-Lista 'Limites' 'Critérios de aceitação. É AQUI que se define nominal, limites e tolerância de cada ensaio. O app usa o valor novo imediatamente.'
Garantir-Coluna 'Limites' 'Ensaio' '<Field Type=''Text'' Name=''Ensaio'' StaticName=''Ensaio'' DisplayName=''Ensaio'' MaxLength=''255''></Field>'
Garantir-Coluna 'Limites' 'Parametro' '<Field Type=''Text'' Name=''Parametro'' StaticName=''Parametro'' DisplayName=''Parametro'' MaxLength=''255''></Field>'
Garantir-Coluna 'Limites' 'Unidade' '<Field Type=''Text'' Name=''Unidade'' StaticName=''Unidade'' DisplayName=''Unidade'' MaxLength=''255''></Field>'
Garantir-Coluna 'Limites' 'Nominal' '<Field Type=''Number'' Name=''Nominal'' StaticName=''Nominal'' DisplayName=''Nominal''></Field>'
Garantir-Coluna 'Limites' 'LimInferior' '<Field Type=''Number'' Name=''LimInferior'' StaticName=''LimInferior'' DisplayName=''LimInferior''></Field>'
Garantir-Coluna 'Limites' 'LimSuperior' '<Field Type=''Number'' Name=''LimSuperior'' StaticName=''LimSuperior'' DisplayName=''LimSuperior''></Field>'
Garantir-Coluna 'Limites' 'Tolerancia' '<Field Type=''Number'' Name=''Tolerancia'' StaticName=''Tolerancia'' DisplayName=''Tolerancia''></Field>'
Garantir-Coluna 'Limites' 'Fonte' '<Field Type=''Note'' Name=''Fonte'' StaticName=''Fonte'' DisplayName=''Fonte'' NumLines=''6'' RichText=''FALSE'' />'
Set-PnPField -List 'Limites' -Identity 'Title' -Values @{ Indexed = $true; EnforceUniqueValues = $true } | Out-Null
$v = Get-PnPView -List 'Limites' | Where-Object { $_.DefaultView }
Set-PnPView -List 'Limites' -Identity $v.Id -Fields @('LinkTitle', 'Ensaio', 'Parametro', 'Unidade', 'Nominal', 'LimInferior', 'LimSuperior', 'Tolerancia', 'Fonte') | Out-Null

# ---------------------------------------------------------------- Afericoes
Garantir-Lista 'Afericoes' 'Um resultado por linha. Mesmo conteúdo da aba BD_Afericoes (o Power BI lê esta lista).'
Garantir-Coluna 'Afericoes' 'Ensaio' '<Field Type=''Text'' Name=''Ensaio'' StaticName=''Ensaio'' DisplayName=''Ensaio'' MaxLength=''255''></Field>'
Garantir-Coluna 'Afericoes' 'TipoRegistro' '<Field Type=''Text'' Name=''TipoRegistro'' StaticName=''TipoRegistro'' DisplayName=''TipoRegistro'' MaxLength=''255''></Field>'
Garantir-Coluna 'Afericoes' 'DataEnsaio' '<Field Type=''DateTime'' Format=''DateOnly'' Name=''DataEnsaio'' StaticName=''DataEnsaio'' DisplayName=''DataEnsaio'' />'
Garantir-Coluna 'Afericoes' 'Equipamento' '<Field Type=''Text'' Name=''Equipamento'' StaticName=''Equipamento'' DisplayName=''Equipamento'' MaxLength=''255''></Field>'
Garantir-Coluna 'Afericoes' 'TAG' '<Field Type=''Text'' Name=''TAG'' StaticName=''TAG'' DisplayName=''TAG'' MaxLength=''255''></Field>'
Garantir-Coluna 'Afericoes' 'Parametro' '<Field Type=''Text'' Name=''Parametro'' StaticName=''Parametro'' DisplayName=''Parametro'' MaxLength=''255''></Field>'
Garantir-Coluna 'Afericoes' 'Unidade' '<Field Type=''Text'' Name=''Unidade'' StaticName=''Unidade'' DisplayName=''Unidade'' MaxLength=''255''></Field>'
Garantir-Coluna 'Afericoes' 'Valor' '<Field Type=''Number'' Name=''Valor'' StaticName=''Valor'' DisplayName=''Valor''></Field>'
Garantir-Coluna 'Afericoes' 'Referencia' '<Field Type=''Number'' Name=''Referencia'' StaticName=''Referencia'' DisplayName=''Referencia''></Field>'
Garantir-Coluna 'Afericoes' 'Diferenca' '<Field Type=''Number'' Name=''Diferenca'' StaticName=''Diferenca'' DisplayName=''Diferenca''></Field>'
Garantir-Coluna 'Afericoes' 'LimInferior' '<Field Type=''Number'' Name=''LimInferior'' StaticName=''LimInferior'' DisplayName=''LimInferior''></Field>'
Garantir-Coluna 'Afericoes' 'LimSuperior' '<Field Type=''Number'' Name=''LimSuperior'' StaticName=''LimSuperior'' DisplayName=''LimSuperior''></Field>'
Garantir-Coluna 'Afericoes' 'Tolerancia' '<Field Type=''Number'' Name=''Tolerancia'' StaticName=''Tolerancia'' DisplayName=''Tolerancia''></Field>'
Garantir-Coluna 'Afericoes' 'Resultado' '<Field Type=''Text'' Name=''Resultado'' StaticName=''Resultado'' DisplayName=''Resultado'' MaxLength=''255''></Field>'
Garantir-Coluna 'Afericoes' 'Responsavel' '<Field Type=''Text'' Name=''Responsavel'' StaticName=''Responsavel'' DisplayName=''Responsavel'' MaxLength=''255''></Field>'
Garantir-Coluna 'Afericoes' 'ResponsavelEmail' '<Field Type=''Text'' Name=''ResponsavelEmail'' StaticName=''ResponsavelEmail'' DisplayName=''ResponsavelEmail'' MaxLength=''255''></Field>'
Garantir-Coluna 'Afericoes' 'Letra' '<Field Type=''Text'' Name=''Letra'' StaticName=''Letra'' DisplayName=''Letra'' MaxLength=''255''></Field>'
Garantir-Coluna 'Afericoes' 'Turno' '<Field Type=''Text'' Name=''Turno'' StaticName=''Turno'' DisplayName=''Turno'' MaxLength=''255''></Field>'
Garantir-Coluna 'Afericoes' 'Observacao' '<Field Type=''Note'' Name=''Observacao'' StaticName=''Observacao'' DisplayName=''Observacao'' NumLines=''6'' RichText=''FALSE'' />'
Garantir-Coluna 'Afericoes' 'Causa' '<Field Type=''Text'' Name=''Causa'' StaticName=''Causa'' DisplayName=''Causa'' MaxLength=''255''></Field>'
Garantir-Coluna 'Afericoes' 'ItensNaoOK' '<Field Type=''Text'' Name=''ItensNaoOK'' StaticName=''ItensNaoOK'' DisplayName=''ItensNaoOK'' MaxLength=''255''></Field>'
Garantir-Coluna 'Afericoes' 'RegistradoEm' '<Field Type=''DateTime'' Format=''DateTime'' Name=''RegistradoEm'' StaticName=''RegistradoEm'' DisplayName=''RegistradoEm'' />'
Garantir-Coluna 'Afericoes' 'TurnoRegistro' '<Field Type=''Text'' Name=''TurnoRegistro'' StaticName=''TurnoRegistro'' DisplayName=''TurnoRegistro'' MaxLength=''255''></Field>'
Garantir-Coluna 'Afericoes' 'ForaDoTurno' '<Field Type=''Text'' Name=''ForaDoTurno'' StaticName=''ForaDoTurno'' DisplayName=''ForaDoTurno'' MaxLength=''255''></Field>'
Garantir-Coluna 'Afericoes' 'Origem' '<Field Type=''Text'' Name=''Origem'' StaticName=''Origem'' DisplayName=''Origem'' MaxLength=''255''></Field>'
Indexar 'Afericoes' 'Title'
Indexar 'Afericoes' 'Ensaio'
Indexar 'Afericoes' 'DataEnsaio'
Indexar 'Afericoes' 'TAG'
$v = Get-PnPView -List 'Afericoes' | Where-Object { $_.DefaultView }
Set-PnPView -List 'Afericoes' -Identity $v.Id -Fields @('LinkTitle', 'Ensaio', 'TipoRegistro', 'DataEnsaio', 'Equipamento', 'TAG', 'Parametro', 'Unidade', 'Valor', 'Referencia', 'Diferenca', 'LimInferior', 'LimSuperior') | Out-Null

# ---------------------------------------------------------------- Lancamentos
Garantir-Lista 'Lancamentos' 'Um resumo por envio (substitui a aba Historico). Dispara o fluxo F1.'
Garantir-Coluna 'Lancamentos' 'Ensaio' '<Field Type=''Text'' Name=''Ensaio'' StaticName=''Ensaio'' DisplayName=''Ensaio'' MaxLength=''255''></Field>'
Garantir-Coluna 'Lancamentos' 'DataEnsaio' '<Field Type=''DateTime'' Format=''DateOnly'' Name=''DataEnsaio'' StaticName=''DataEnsaio'' DisplayName=''DataEnsaio'' />'
Garantir-Coluna 'Lancamentos' 'Turno' '<Field Type=''Text'' Name=''Turno'' StaticName=''Turno'' DisplayName=''Turno'' MaxLength=''255''></Field>'
Garantir-Coluna 'Lancamentos' 'Letra' '<Field Type=''Text'' Name=''Letra'' StaticName=''Letra'' DisplayName=''Letra'' MaxLength=''255''></Field>'
Garantir-Coluna 'Lancamentos' 'Responsavel' '<Field Type=''Text'' Name=''Responsavel'' StaticName=''Responsavel'' DisplayName=''Responsavel'' MaxLength=''255''></Field>'
Garantir-Coluna 'Lancamentos' 'ResponsavelEmail' '<Field Type=''Text'' Name=''ResponsavelEmail'' StaticName=''ResponsavelEmail'' DisplayName=''ResponsavelEmail'' MaxLength=''255''></Field>'
Garantir-Coluna 'Lancamentos' 'QtdResultados' '<Field Type=''Number'' Name=''QtdResultados'' StaticName=''QtdResultados'' DisplayName=''QtdResultados''></Field>'
Garantir-Coluna 'Lancamentos' 'QtdNaoConformes' '<Field Type=''Number'' Name=''QtdNaoConformes'' StaticName=''QtdNaoConformes'' DisplayName=''QtdNaoConformes''></Field>'
Garantir-Coluna 'Lancamentos' 'TAGsNaoConformes' '<Field Type=''Note'' Name=''TAGsNaoConformes'' StaticName=''TAGsNaoConformes'' DisplayName=''TAGsNaoConformes'' NumLines=''6'' RichText=''FALSE'' />'
Garantir-Coluna 'Lancamentos' 'Causa' '<Field Type=''Text'' Name=''Causa'' StaticName=''Causa'' DisplayName=''Causa'' MaxLength=''255''></Field>'
Garantir-Coluna 'Lancamentos' 'Observacao' '<Field Type=''Note'' Name=''Observacao'' StaticName=''Observacao'' DisplayName=''Observacao'' NumLines=''6'' RichText=''FALSE'' />'
Garantir-Coluna 'Lancamentos' 'RegistradoEm' '<Field Type=''DateTime'' Format=''DateTime'' Name=''RegistradoEm'' StaticName=''RegistradoEm'' DisplayName=''RegistradoEm'' />'
Garantir-Coluna 'Lancamentos' 'ForaDoTurno' '<Field Type=''Text'' Name=''ForaDoTurno'' StaticName=''ForaDoTurno'' DisplayName=''ForaDoTurno'' MaxLength=''255''></Field>'
Garantir-Coluna 'Lancamentos' 'Situacao' '<Field Type=''Text'' Name=''Situacao'' StaticName=''Situacao'' DisplayName=''Situacao'' MaxLength=''255''><Default>Ativo</Default></Field>'
Garantir-Coluna 'Lancamentos' 'Origem' '<Field Type=''Text'' Name=''Origem'' StaticName=''Origem'' DisplayName=''Origem'' MaxLength=''255''></Field>'
Set-PnPField -List 'Lancamentos' -Identity 'Title' -Values @{ Indexed = $true; EnforceUniqueValues = $true } | Out-Null
Indexar 'Lancamentos' 'Ensaio'
Indexar 'Lancamentos' 'DataEnsaio'
$v = Get-PnPView -List 'Lancamentos' | Where-Object { $_.DefaultView }
Set-PnPView -List 'Lancamentos' -Identity $v.Id -Fields @('LinkTitle', 'Ensaio', 'DataEnsaio', 'Turno', 'Letra', 'Responsavel', 'ResponsavelEmail', 'QtdResultados', 'QtdNaoConformes', 'TAGsNaoConformes', 'Causa', 'Observacao', 'RegistradoEm') | Out-Null

# ---------------------------------------------------------------- Calibracoes
Garantir-Lista 'Calibracoes' 'Histórico de calibrações (já existe da versão 1).'
Garantir-Coluna 'Calibracoes' 'EquipamentoID' '<Field Type=''Number'' Name=''EquipamentoID'' StaticName=''EquipamentoID'' DisplayName=''EquipamentoID''></Field>'
Garantir-Coluna 'Calibracoes' 'DataCalibracao' '<Field Type=''DateTime'' Format=''DateOnly'' Name=''DataCalibracao'' StaticName=''DataCalibracao'' DisplayName=''DataCalibracao'' />'
Garantir-Coluna 'Calibracoes' 'Responsavel' '<Field Type=''Text'' Name=''Responsavel'' StaticName=''Responsavel'' DisplayName=''Responsavel'' MaxLength=''255''></Field>'
Garantir-Coluna 'Calibracoes' 'Local' '<Field Type=''Text'' Name=''Local'' StaticName=''Local'' DisplayName=''Local'' MaxLength=''255''></Field>'
Garantir-Coluna 'Calibracoes' 'Resultado' '<Field Type=''Text'' Name=''Resultado'' StaticName=''Resultado'' DisplayName=''Resultado'' MaxLength=''255''></Field>'
Garantir-Coluna 'Calibracoes' 'MotivoRecusa' '<Field Type=''Note'' Name=''MotivoRecusa'' StaticName=''MotivoRecusa'' DisplayName=''MotivoRecusa'' NumLines=''6'' RichText=''FALSE'' />'
Garantir-Coluna 'Calibracoes' 'ManutencaoSolicitada' '<Field Type=''Boolean'' Name=''ManutencaoSolicitada'' StaticName=''ManutencaoSolicitada'' DisplayName=''ManutencaoSolicitada''><Default>0</Default></Field>'
Garantir-Coluna 'Calibracoes' 'ProximaCalibracao' '<Field Type=''DateTime'' Format=''DateOnly'' Name=''ProximaCalibracao'' StaticName=''ProximaCalibracao'' DisplayName=''ProximaCalibracao'' />'
Garantir-Coluna 'Calibracoes' 'Certificado' '<Field Type=''Text'' Name=''Certificado'' StaticName=''Certificado'' DisplayName=''Certificado'' MaxLength=''255''></Field>'
Garantir-Coluna 'Calibracoes' 'Observacoes' '<Field Type=''Note'' Name=''Observacoes'' StaticName=''Observacoes'' DisplayName=''Observacoes'' NumLines=''6'' RichText=''FALSE'' />'
Indexar 'Calibracoes' 'EquipamentoID'
$v = Get-PnPView -List 'Calibracoes' | Where-Object { $_.DefaultView }
Set-PnPView -List 'Calibracoes' -Identity $v.Id -Fields @('LinkTitle', 'EquipamentoID', 'DataCalibracao', 'Responsavel', 'Local', 'Resultado', 'MotivoRecusa', 'ManutencaoSolicitada', 'ProximaCalibracao', 'Certificado', 'Observacoes') | Out-Null

# ---------------------------------------------------------------- Configuracoes
Garantir-Lista 'Configuracoes' 'Parâmetros gerais: escala das letras, horário dos turnos, massas de ensaio, e-mails dos relatórios.'
Garantir-Coluna 'Configuracoes' 'Valor' '<Field Type=''Text'' Name=''Valor'' StaticName=''Valor'' DisplayName=''Valor'' MaxLength=''255'' Required=''TRUE''></Field>'
Garantir-Coluna 'Configuracoes' 'Descricao' '<Field Type=''Note'' Name=''Descricao'' StaticName=''Descricao'' DisplayName=''Descricao'' NumLines=''6'' RichText=''FALSE'' />'
Set-PnPField -List 'Configuracoes' -Identity 'Title' -Values @{ Indexed = $true; EnforceUniqueValues = $true } | Out-Null
$v = Get-PnPView -List 'Configuracoes' | Where-Object { $_.DefaultView }
Set-PnPView -List 'Configuracoes' -Identity $v.Id -Fields @('LinkTitle', 'Valor', 'Descricao') | Out-Null

# ---------------------------------------------------------------- Formatação de cores (Status / Resultado)
$fmt = Join-Path $PSScriptRoot 'formatacao'
if (Test-Path (Join-Path $fmt 'equipamentos-status.json')) {
    Set-PnPField -List 'Equipamentos' -Identity 'Status' -Values @{ CustomFormatter = (Get-Content -Raw -Encoding UTF8 (Join-Path $fmt 'equipamentos-status.json')) } | Out-Null
    Set-PnPField -List 'Afericoes' -Identity 'Resultado' -Values @{ CustomFormatter = (Get-Content -Raw -Encoding UTF8 (Join-Path $fmt 'afericoes-resultado.json')) } | Out-Null
}

# ---------------------------------------------------------------- Grupos e permissões
$colaborar = (Get-PnPRoleDefinition | Where-Object { $_.RoleTypeKind -eq 'Contributor' } | Select-Object -First 1).Name
$leitura   = (Get-PnPRoleDefinition | Where-Object { $_.RoleTypeKind -eq 'Reader' } | Select-Object -First 1).Name
if (-not (Get-PnPGroup -Identity 'Metrologia - Laboratoristas' -ErrorAction SilentlyContinue)) { New-PnPSiteGroup -Name 'Metrologia - Laboratoristas' -PermissionLevels $leitura | Out-Null; Write-Host 'Grupo criado: Metrologia - Laboratoristas' -ForegroundColor Green }
if (-not (Get-PnPGroup -Identity 'Metrologia - Instrumentação' -ErrorAction SilentlyContinue)) { New-PnPSiteGroup -Name 'Metrologia - Instrumentação' -PermissionLevels $leitura | Out-Null; Write-Host 'Grupo criado: Metrologia - Instrumentação' -ForegroundColor Green }
# As listas que recebem gravação deixam de herdar permissões; as demais continuam herdando do site (Leitura para visitantes).
Set-PnPList -Identity 'Afericoes' -BreakRoleInheritance -CopyRoleAssignments | Out-Null
Set-PnPList -Identity 'Calibracoes' -BreakRoleInheritance -CopyRoleAssignments | Out-Null
Set-PnPList -Identity 'Equipamentos' -BreakRoleInheritance -CopyRoleAssignments | Out-Null
Set-PnPList -Identity 'Lancamentos' -BreakRoleInheritance -CopyRoleAssignments | Out-Null
Set-PnPList -Identity 'PontosAfericao' -BreakRoleInheritance -CopyRoleAssignments | Out-Null
Set-PnPListPermission -Identity 'Afericoes' -Group 'Metrologia - Laboratoristas' -AddRole $colaborar | Out-Null
Set-PnPListPermission -Identity 'Lancamentos' -Group 'Metrologia - Laboratoristas' -AddRole $colaborar | Out-Null
Set-PnPListPermission -Identity 'Calibracoes' -Group 'Metrologia - Instrumentação' -AddRole $colaborar | Out-Null
Set-PnPListPermission -Identity 'Equipamentos' -Group 'Metrologia - Instrumentação' -AddRole $colaborar | Out-Null
Set-PnPListPermission -Identity 'PontosAfericao' -Group 'Metrologia - Instrumentação' -AddRole $colaborar | Out-Null
Write-Host 'Permissões aplicadas. Confira se o grupo de Visitantes do site contém "Todos, exceto usuários externos".' -ForegroundColor Cyan

# ---------------------------------------------------------------- Importação da carga inicial (opcional)
function Converter($lista, $coluna, $valor) {
    if ([string]::IsNullOrWhiteSpace($valor)) { return $null }
    $campo = Get-PnPField -List $lista -Identity $coluna
    switch ($campo.TypeAsString) {
        'Boolean'  { return ($valor -eq '1' -or $valor -eq 'True') }
        'Number'   { return [double]::Parse($valor, $inv) }
        'DateTime' { return [datetime]::Parse($valor, $inv) }
        default    { return $valor }
    }
}
function Importar([string] $Lista, [string] $Chave) {
    $arq = Join-Path $PastaCsv "$Lista.csv"
    if (-not (Test-Path $arq)) { Write-Host "Sem CSV para $Lista" -ForegroundColor Yellow; return }
    $linhas = Import-Csv -Path $arq -Encoding UTF8
    $existentes = @{}
    if ($Chave) { Get-PnPListItem -List $Lista -PageSize 2000 -Fields 'Title' | ForEach-Object { $existentes[$_['Title']] = $_.Id } }
    $lote = New-PnPBatch
    $n = 0
    foreach ($l in $linhas) {
        $valores = @{}
        foreach ($p in $l.PSObject.Properties) {
            if ($p.Name -eq 'Pessoa') { continue }
            $v = Converter $Lista $p.Name $p.Value
            if ($null -ne $v) { $valores[$p.Name] = $v }
        }
        if ($Chave -and $existentes.ContainsKey($l.Title)) {
            Set-PnPListItem -List $Lista -Identity $existentes[$l.Title] -Values $valores -Batch $lote | Out-Null
        } else {
            Add-PnPListItem -List $Lista -Values $valores -Batch $lote | Out-Null
        }
        $n++
    }
    Invoke-PnPBatch -Batch $lote
    Write-Host "${Lista}: $n linhas importadas" -ForegroundColor Green
}
if ($ImportarDados) {
    Importar 'Configuracoes' 'Title'
    Importar 'Limites' 'Title'
    Importar 'Ensaios' 'Title'
    Importar 'PontosAfericao' 'Title'
    Importar 'Equipamentos' 'Title'
    Importar 'Lancamentos' 'Title'
    Importar 'Afericoes' $null
    Write-Host 'Acessos não é importado: cadastre cada pessoa pelo seletor (coluna Pessoa).' -ForegroundColor Cyan
}
Write-Host 'Pronto.' -ForegroundColor Cyan
