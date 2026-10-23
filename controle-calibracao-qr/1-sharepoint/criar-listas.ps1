<#
  Controle de Calibração por QR Code - criação automática das listas no SharePoint

  OPCIONAL: se você não tiver o PnP.PowerShell (ou a TI não liberar), crie as listas
  manualmente seguindo o arquivo "estrutura-das-listas.md". O resultado é o mesmo.

  Pré-requisitos (normalmente feito pela TI):
    1. PowerShell 7+   ->  winget install Microsoft.PowerShell
    2. Install-Module PnP.PowerShell -Scope CurrentUser
    3. Um "App Registration" no Entra ID para o PnP (exigido pela Microsoft desde set/2024):
         Register-PnPEntraIDAppForInteractiveLogin -ApplicationName "PnP Calibracao" -Tenant SUAEMPRESA.onmicrosoft.com -Interactive
       Guarde o "ClientId" que ele mostrar.

  Uso:
    ./criar-listas.ps1 -SiteUrl "https://SUAEMPRESA.sharepoint.com/sites/Laboratorio" -ClientId "xxxxxxxx-..." -DadosExemplo
#>
param(
    [Parameter(Mandatory = $true)] [string] $SiteUrl,
    [Parameter(Mandatory = $true)] [string] $ClientId,
    [switch] $DadosExemplo
)

$ErrorActionPreference = 'Stop'
Import-Module PnP.PowerShell
Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId

function New-Lista([string] $Nome, [string] $Descricao) {
    if (-not (Get-PnPList -Identity $Nome -ErrorAction SilentlyContinue)) {
        New-PnPList -Title $Nome -Template GenericList -OnQuickLaunch | Out-Null
        Set-PnPList -Identity $Nome -Description $Descricao | Out-Null
        Write-Host "Lista criada: $Nome" -ForegroundColor Green
    } else {
        Write-Host "Lista já existe: $Nome" -ForegroundColor Yellow
    }
}

function Add-Coluna([string] $Lista, [string] $Nome, [string] $Tipo, [string] $Padrao = '') {
    if (Get-PnPField -List $Lista -Identity $Nome -ErrorAction SilentlyContinue) { return }
    $id  = [guid]::NewGuid().ToString()
    $base = "Name='$Nome' StaticName='$Nome' DisplayName='$Nome' ID='{$id}'"
    $xml = switch ($Tipo) {
        'Texto'  { "<Field Type='Text' $base MaxLength='255' />" }
        'Nota'   { "<Field Type='Note' $base NumLines='6' RichText='FALSE' />" }
        'Data'   { "<Field Type='DateTime' Format='DateOnly' $base />" }
        'Numero' { "<Field Type='Number' Decimals='0' $base>" + $(if ($Padrao) { "<Default>$Padrao</Default>" }) + "</Field>" }
        'SimNao' { "<Field Type='Boolean' $base><Default>$Padrao</Default></Field>" }
    }
    Add-PnPFieldFromXml -List $Lista -FieldXml $xml | Out-Null
    Write-Host "   + $Lista.$Nome ($Tipo)"
}

function Set-VisaoPadrao([string] $Lista, [string[]] $Campos) {
    $visao = Get-PnPView -List $Lista | Where-Object { $_.DefaultView }
    Set-PnPView -List $Lista -Identity $visao.Id -Fields $Campos | Out-Null
}

# ---------------------------------------------------------------- Equipamentos
New-Lista 'Equipamentos' 'Cadastro e situação de calibração dos equipamentos do laboratório (Title = TAG do equipamento)'
Add-Coluna 'Equipamentos' 'Descricao'           'Texto'
Add-Coluna 'Equipamentos' 'Tipo'                'Texto'
Add-Coluna 'Equipamentos' 'FabricanteModelo'    'Texto'
Add-Coluna 'Equipamentos' 'NumeroSerie'         'Texto'
Add-Coluna 'Equipamentos' 'Local'               'Texto'
Add-Coluna 'Equipamentos' 'PeriodicidadeMeses'  'Numero' '12'
Add-Coluna 'Equipamentos' 'UltimaCalibracao'    'Data'
Add-Coluna 'Equipamentos' 'ProximaCalibracao'   'Data'
Add-Coluna 'Equipamentos' 'UltimoResultado'     'Texto'
Add-Coluna 'Equipamentos' 'UltimoResponsavel'   'Texto'
Add-Coluna 'Equipamentos' 'MotivoRecusa'        'Nota'
Add-Coluna 'Equipamentos' 'ManutencaoSolicitada' 'SimNao' '0'
Add-Coluna 'Equipamentos' 'Status'              'Texto'
Add-Coluna 'Equipamentos' 'Ativo'               'SimNao' '1'
# TAG única e indexada (o QR Code depende disso)
Set-PnPField -List 'Equipamentos' -Identity 'Title' -Values @{ Indexed = $true; EnforceUniqueValues = $true } | Out-Null
Set-VisaoPadrao 'Equipamentos' @('LinkTitle','Descricao','Local','UltimaCalibracao','ProximaCalibracao','Status','MotivoRecusa','UltimoResponsavel','PeriodicidadeMeses')

# ---------------------------------------------------------------- Calibracoes
New-Lista 'Calibracoes' 'Histórico de todas as calibrações registradas pelo app (Title = TAG do equipamento)'
Add-Coluna 'Calibracoes' 'EquipamentoID'        'Numero'
Add-Coluna 'Calibracoes' 'DataCalibracao'       'Data'
Add-Coluna 'Calibracoes' 'Responsavel'          'Texto'
Add-Coluna 'Calibracoes' 'Local'                'Texto'
Add-Coluna 'Calibracoes' 'Resultado'            'Texto'
Add-Coluna 'Calibracoes' 'MotivoRecusa'         'Nota'
Add-Coluna 'Calibracoes' 'ManutencaoSolicitada' 'SimNao' '0'
Add-Coluna 'Calibracoes' 'ProximaCalibracao'    'Data'
Add-Coluna 'Calibracoes' 'Certificado'          'Texto'
Add-Coluna 'Calibracoes' 'Observacoes'          'Nota'
Set-PnPField -List 'Calibracoes' -Identity 'EquipamentoID' -Values @{ Indexed = $true } | Out-Null
Set-VisaoPadrao 'Calibracoes' @('LinkTitle','DataCalibracao','Resultado','Responsavel','Local','MotivoRecusa','ManutencaoSolicitada','ProximaCalibracao','Certificado','Created')

# ---------------------------------------------------------------- Calibradores
New-Lista 'Calibradores' 'Pessoas que podem registrar calibrações no app (Title = e-mail corporativo)'
Add-Coluna 'Calibradores' 'Nome' 'Texto'
Set-VisaoPadrao 'Calibradores' @('LinkTitle','Nome')

# ---------------------------------------------------------------- Formatação de cores da coluna Status
$json = Get-Content -Raw -Encoding UTF8 (Join-Path $PSScriptRoot 'formatacao-coluna-status.json')
Set-PnPField -List 'Equipamentos' -Identity 'Status' -Values @{ CustomFormatter = $json } | Out-Null

# ---------------------------------------------------------------- Dados de exemplo (opcional)
if ($DadosExemplo) {
    $exemplos = @(
        @{ Title='BAL-01'; Descricao='Balança Prix';       Tipo='Balança';     Local='Laboratório - Bancada 1'; PeriodicidadeMeses=6;  UltimaCalibracao='2026-10-10'; ProximaCalibracao='2027-04-10'; UltimoResultado='Aprovada' },
        @{ Title='REN-02'; Descricao='Reneira';            Tipo='Reneira';     Local='Laboratório - Peneiras';  PeriodicidadeMeses=6;  UltimaCalibracao='2026-05-05'; ProximaCalibracao='2026-11-05'; UltimoResultado='Aprovada' },
        @{ Title='BAL-03'; Descricao='Balança analítica';  Tipo='Balança';     Local='Laboratório - Bancada 2'; PeriodicidadeMeses=12; UltimaCalibracao='2025-08-01'; ProximaCalibracao='2026-08-01'; UltimoResultado='Aprovada' },
        @{ Title='EST-01'; Descricao='Estufa de secagem';  Tipo='Estufa';      Local='Laboratório - Estufas';   PeriodicidadeMeses=12; UltimaCalibracao='2026-06-15'; ProximaCalibracao='2027-06-15'; UltimoResultado='Aprovada' },
        @{ Title='TER-01'; Descricao='Termômetro digital'; Tipo='Termômetro';  Local='Laboratório - Bancada 1'; PeriodicidadeMeses=6;  UltimaCalibracao='2026-05-20'; ProximaCalibracao='2026-11-20'; UltimoResultado='Aprovada' },
        @{ Title='BAL-04'; Descricao='Balança de piso';    Tipo='Balança';     Local='Recebimento';             PeriodicidadeMeses=6;  UltimaCalibracao='2026-03-12'; ProximaCalibracao='2026-09-12'; UltimoResultado='Recusada'; MotivoRecusa='Desvio na linearidade' },
        @{ Title='REN-01'; Descricao='Reneira';            Tipo='Reneira';     Local='Laboratório - Peneiras';  PeriodicidadeMeses=6;  UltimaCalibracao='2026-07-18'; ProximaCalibracao='2027-01-18'; UltimoResultado='Recusada'; MotivoRecusa='Falha mecânica / motor'; ManutencaoSolicitada=$true },
        @{ Title='BAL-05'; Descricao='Balança Prix';       Tipo='Balança';     Local='Laboratório - Bancada 3'; PeriodicidadeMeses=6;  UltimaCalibracao='2026-08-22'; ProximaCalibracao='2027-02-22'; UltimoResultado='Aprovada' }
    )
    foreach ($e in $exemplos) {
        if (-not (Get-PnPListItem -List 'Equipamentos' -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>$($e.Title)</Value></Eq></Where></Query></View>")) {
            Add-PnPListItem -List 'Equipamentos' -Values $e | Out-Null
            Write-Host "   + equipamento de exemplo $($e.Title)"
        }
    }
}

Write-Host ''
Write-Host 'Pronto! Próximos passos:' -ForegroundColor Cyan
Write-Host ' 1. Adicione os e-mails da instrumentação na lista Calibradores.'
Write-Host ' 2. Ajuste as permissões (ver estrutura-das-listas.md, seção Permissões).'
Write-Host ' 3. Monte o app no Power Apps (pasta 2-powerapps).'
