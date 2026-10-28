"""
Gera 02-sharepoint/criar-estrutura.ps1 (PnP.PowerShell) a partir de modelo.py.

Uso:  python gerar_pnp.py
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from modelo import BIBLIOTECA_IMAGENS, GRUPOS, LISTAS  # noqa: E402

RAIZ = Path(__file__).resolve().parents[1]


def xml_campo(c) -> str:
    base = f"Name='{c.nome}' StaticName='{c.nome}' DisplayName='{c.nome}'"
    req = " Required='TRUE'" if c.obrigatoria else ""
    padrao = f"<Default>{c.padrao}</Default>" if c.padrao is not None else ""
    if c.tipo == "Text":
        return f"<Field Type='Text' {base} MaxLength='255'{req}>{padrao}</Field>"
    if c.tipo == "Note":
        return f"<Field Type='Note' {base} NumLines='6' RichText='FALSE'{req} />"
    if c.tipo == "Number":
        return f"<Field Type='Number' {base}{req}>{padrao}</Field>"
    if c.tipo == "Date":
        return f"<Field Type='DateTime' Format='DateOnly' {base}{req} />"
    if c.tipo == "DateTime":
        return f"<Field Type='DateTime' Format='DateTime' {base}{req} />"
    if c.tipo == "Boolean":
        return f"<Field Type='Boolean' {base}>{padrao or '<Default>0</Default>'}</Field>"
    if c.tipo == "Choice":
        ch = "".join(f"<CHOICE>{x}</CHOICE>" for x in c.choices)
        return f"<Field Type='Choice' Format='Dropdown' FillInChoice='FALSE' {base}{req}><CHOICES>{ch}</CHOICES>{padrao}</Field>"
    if c.tipo == "User":
        return f"<Field Type='User' UserSelectionMode='PeopleOnly' UserSelectionScope='0' ShowField='ImnName' {base}{req} />"
    raise ValueError(c.tipo)


def ps_str(s: str) -> str:
    return "'" + s.replace("'", "''") + "'"


def main():
    out = [
        "<#",
        "  SISTEMA METROLOGIA · criação da estrutura no SharePoint (gerado por ferramentas/gerar_pnp.py)",
        "",
        "  O que faz:",
        "    1. Cria as listas que faltam e acrescenta as colunas novas nas listas da versão 1",
        "       (Equipamentos e Calibracoes), sem apagar nada.",
        "    2. Cria índices (necessários para a lista passar de 5.000 itens sem travar o app).",
        "    3. Cria os grupos 'Metrologia - Laboratoristas' e 'Metrologia - Instrumentação' e",
        "       aplica as permissões de cada lista.",
        "    4. Cria a biblioteca ConsultaQR (uma imagem por equipamento, aberta pela câmera do celular)",
        "       e a exibição 'Imagens de consulta' na lista Equipamentos.",
        "    5. (Opcional, -ImportarDados) importa a carga inicial dos CSVs gerados por gerar_carga.py.",
        "",
        "  Pré-requisitos: PowerShell 7+, módulo PnP.PowerShell e um App Registration do Entra ID",
        "  (Register-PnPEntraIDAppForInteractiveLogin). Peça à TI se não tiver.",
        "",
        "  Uso:",
        "    ./criar-estrutura.ps1 -SiteUrl https://EMPRESA.sharepoint.com/sites/SITE -ClientId <guid>",
        "    ./criar-estrutura.ps1 -SiteUrl ... -ClientId ... -ImportarDados -PastaCsv ..\\dados\\csv",
        "#>",
        "param(",
        "    [Parameter(Mandatory = $true)] [string] $SiteUrl,",
        "    [Parameter(Mandatory = $true)] [string] $ClientId,",
        "    [switch] $ImportarDados,",
        "    [string] $PastaCsv = (Join-Path $PSScriptRoot 'csv')",
        ")",
        "$ErrorActionPreference = 'Stop'",
        "Import-Module PnP.PowerShell",
        "Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId",
        "$inv = [System.Globalization.CultureInfo]::InvariantCulture",
        "",
        "function Garantir-Lista([string] $Nome, [string] $Descricao) {",
        "    if (-not (Get-PnPList -Identity $Nome -ErrorAction SilentlyContinue)) {",
        "        New-PnPList -Title $Nome -Template GenericList -OnQuickLaunch | Out-Null",
        "        Write-Host \"Lista criada: $Nome\" -ForegroundColor Green",
        "    } else { Write-Host \"Lista já existe: $Nome\" -ForegroundColor Yellow }",
        "    Set-PnPList -Identity $Nome -Description $Descricao -EnableVersioning $true | Out-Null",
        "}",
        "function Garantir-Coluna([string] $Lista, [string] $Nome, [string] $Xml) {",
        "    if (Get-PnPField -List $Lista -Identity $Nome -ErrorAction SilentlyContinue) { return }",
        "    $x = $Xml.Replace(\"Name='$Nome'\", \"ID='{\" + [guid]::NewGuid().ToString() + \"}' Name='$Nome'\")",
        "    Add-PnPFieldFromXml -List $Lista -FieldXml $x | Out-Null",
        "    Write-Host \"   + $Lista.$Nome\"",
        "}",
        "function Indexar([string] $Lista, [string] $Nome) {",
        "    Set-PnPField -List $Lista -Identity $Nome -Values @{ Indexed = $true } | Out-Null",
        "}",
        "",
    ]
    for L in LISTAS:
        out.append(f"# ---------------------------------------------------------------- {L.nome}")
        out.append(f"Garantir-Lista {ps_str(L.nome)} {ps_str(L.descricao)}")
        for c in L.colunas:
            out.append(f"Garantir-Coluna {ps_str(L.nome)} {ps_str(c.nome)} {ps_str(xml_campo(c))}")
        if L.titulo_unico:
            out.append(f"Set-PnPField -List {ps_str(L.nome)} -Identity 'Title' -Values @{{ Indexed = $true; EnforceUniqueValues = $true }} | Out-Null")
        elif L.titulo_indexado:
            out.append(f"Indexar {ps_str(L.nome)} 'Title'")
        for c in L.colunas:
            if c.indexada:
                out.append(f"Indexar {ps_str(L.nome)} {ps_str(c.nome)}")
        campos = ["LinkTitle"] + [c.nome for c in L.colunas][:12]
        out.append(f"$v = Get-PnPView -List {ps_str(L.nome)} | Where-Object {{ $_.DefaultView }}")
        out.append(f"Set-PnPView -List {ps_str(L.nome)} -Identity $v.Id -Fields @({', '.join(ps_str(x) for x in campos)}) | Out-Null")
        out.append("")

    bib = BIBLIOTECA_IMAGENS
    out += [
        f"# ---------------------------------------------------------------- Biblioteca {bib['nome']} (imagens do QR Code)",
        f"if (-not (Get-PnPList -Identity {ps_str(bib['nome'])} -ErrorAction SilentlyContinue)) {{",
        f"    New-PnPList -Title {ps_str(bib['nome'])} -Template DocumentLibrary -OnQuickLaunch | Out-Null",
        f"    Write-Host 'Biblioteca criada: {bib['nome']}' -ForegroundColor Green",
        "}",
        f"Set-PnPList -Identity {ps_str(bib['nome'])} -Description {ps_str(bib['descricao'])} -EnableVersioning $true -MajorVersions 50 | Out-Null",
        "# Herda as permissões do site: todos leem; só os proprietários substituem as imagens.",
        f"if (-not (Get-PnPView -List 'Equipamentos' -Identity {ps_str(bib['exibicao'])} -ErrorAction SilentlyContinue)) {{",
        f"    Add-PnPView -List 'Equipamentos' -Title {ps_str(bib['exibicao'])} -Fields @({', '.join(ps_str(x) for x in bib['campos_exibicao'])}) -Query \"<OrderBy><FieldRef Name='Title' /></OrderBy><Where><Eq><FieldRef Name='Ativo' /><Value Type='Boolean'>1</Value></Eq></Where>\" -RowLimit 500 | Out-Null",
        f"    Write-Host 'Exibição criada: Equipamentos > {bib['exibicao']}' -ForegroundColor Green",
        "}",
        "",
        "# ---------------------------------------------------------------- Formatação de cores (Status / Resultado)",
        "$fmt = Join-Path $PSScriptRoot 'formatacao'",
        "if (Test-Path (Join-Path $fmt 'equipamentos-status.json')) {",
        "    Set-PnPField -List 'Equipamentos' -Identity 'Status' -Values @{ CustomFormatter = (Get-Content -Raw -Encoding UTF8 (Join-Path $fmt 'equipamentos-status.json')) } | Out-Null",
        "    Set-PnPField -List 'Afericoes' -Identity 'Resultado' -Values @{ CustomFormatter = (Get-Content -Raw -Encoding UTF8 (Join-Path $fmt 'afericoes-resultado.json')) } | Out-Null",
        "}",
        "",
        "# ---------------------------------------------------------------- Grupos e permissões",
        "$colaborar = (Get-PnPRoleDefinition | Where-Object { $_.RoleTypeKind -eq 'Contributor' } | Select-Object -First 1).Name",
        "$leitura   = (Get-PnPRoleDefinition | Where-Object { $_.RoleTypeKind -eq 'Reader' } | Select-Object -First 1).Name",
    ]
    for g in GRUPOS:
        out.append(f"if (-not (Get-PnPGroup -Identity {ps_str(g)} -ErrorAction SilentlyContinue)) {{ New-PnPSiteGroup -Name {ps_str(g)} -PermissionLevels $leitura | Out-Null; Write-Host 'Grupo criado: {g}' -ForegroundColor Green }}")
    listas_com_permissao = sorted({l for regras in GRUPOS.values() for l in regras})
    out.append("# As listas que recebem gravação deixam de herdar permissões; as demais continuam herdando do site (Leitura para visitantes).")
    for nome in listas_com_permissao:
        out.append(f"Set-PnPList -Identity {ps_str(nome)} -BreakRoleInheritance -CopyRoleAssignments | Out-Null")
    for g, regras in GRUPOS.items():
        for nome in regras:
            out.append(f"Set-PnPListPermission -Identity {ps_str(nome)} -Group {ps_str(g)} -AddRole $colaborar | Out-Null")
    out += [
        "Write-Host 'Permissões aplicadas. Confira se o grupo de Visitantes do site contém \"Todos, exceto usuários externos\".' -ForegroundColor Cyan",
        "",
        "# ---------------------------------------------------------------- Importação da carga inicial (opcional)",
        "function Converter($lista, $coluna, $valor) {",
        "    if ([string]::IsNullOrWhiteSpace($valor)) { return $null }",
        "    $campo = Get-PnPField -List $lista -Identity $coluna",
        "    switch ($campo.TypeAsString) {",
        "        'Boolean'  { return ($valor -eq '1' -or $valor -eq 'True') }",
        "        'Number'   { return [double]::Parse($valor, $inv) }",
        "        'DateTime' { return [datetime]::Parse($valor, $inv) }",
        "        default    { return $valor }",
        "    }",
        "}",
        "function Importar([string] $Lista, [string] $Chave) {",
        "    $arq = Join-Path $PastaCsv \"$Lista.csv\"",
        "    if (-not (Test-Path $arq)) { Write-Host \"Sem CSV para $Lista\" -ForegroundColor Yellow; return }",
        "    $linhas = Import-Csv -Path $arq -Encoding UTF8",
        "    $existentes = @{}",
        "    if ($Chave) { Get-PnPListItem -List $Lista -PageSize 2000 -Fields 'Title' | ForEach-Object { $existentes[$_['Title']] = $_.Id } }",
        "    $lote = New-PnPBatch",
        "    $n = 0",
        "    foreach ($l in $linhas) {",
        "        $valores = @{}",
        "        foreach ($p in $l.PSObject.Properties) {",
        "            if ($p.Name -eq 'Pessoa') { continue }",
        "            $v = Converter $Lista $p.Name $p.Value",
        "            if ($null -ne $v) { $valores[$p.Name] = $v }",
        "        }",
        "        if ($Chave -and $existentes.ContainsKey($l.Title)) {",
        "            Set-PnPListItem -List $Lista -Identity $existentes[$l.Title] -Values $valores -Batch $lote | Out-Null",
        "        } else {",
        "            Add-PnPListItem -List $Lista -Values $valores -Batch $lote | Out-Null",
        "        }",
        "        $n++",
        "    }",
        "    Invoke-PnPBatch -Batch $lote",
        "    Write-Host \"${Lista}: $n linhas importadas\" -ForegroundColor Green",
        "}",
        "if ($ImportarDados) {",
        "    Importar 'Configuracoes' 'Title'",
        "    Importar 'Limites' 'Title'",
        "    Importar 'Ensaios' 'Title'",
        "    Importar 'PontosAfericao' 'Title'",
        "    Importar 'Equipamentos' 'Title'",
        "    Importar 'Lancamentos' 'Title'",
        "    Importar 'Afericoes' $null",
        "    Write-Host 'Acessos não é importado: cadastre cada pessoa pelo seletor (coluna Pessoa).' -ForegroundColor Cyan",
        "}",
        "Write-Host 'Pronto.' -ForegroundColor Cyan",
    ]
    destino = RAIZ / "02-sharepoint" / "criar-estrutura.ps1"
    destino.write_text("\n".join(out) + "\n", encoding="utf-8-sig")
    print(f"gerado: {destino} ({len(out)} linhas)")


if __name__ == "__main__":
    main()
