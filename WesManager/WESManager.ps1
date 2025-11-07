# ===== IIS – Editar NumeroProviders (pool.config) e useCOMFree (web.config) via Out-GridView =====
# Sem backup .bak | Escolha alterar NumeroProviders, useCOMFree ou ambos
# Requer: PowerShell como Admin, módulo WebAdministration

Import-Module WebAdministration

# ---------- Leitura ----------
function Get-NumeroProviders {
    param([string]$physicalPath)
    try {
        if ([string]::IsNullOrWhiteSpace($physicalPath)) { return $null }
        $base = [Environment]::ExpandEnvironmentVariables($physicalPath)
        $cfg  = Join-Path -Path $base -ChildPath 'pool.config'
        if (-not (Test-Path -LiteralPath $cfg)) { return $null }

        $content = Get-Content -LiteralPath $cfg -Raw -ErrorAction Stop
        $rx = New-Object System.Text.RegularExpressions.Regex(
            '<\s*maxload\s*>\s*([^<]+)\s*<\s*/\s*maxload\s*>',
            [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
        )
        $m = $rx.Match($content)
        if ($m.Success) { return ($m.Groups[1].Value.Trim()) }
        return $null
    } catch { return $null }
}

function Get-UseCOMFree {
    param([string]$physicalPath)
    try {
        if ([string]::IsNullOrWhiteSpace($physicalPath)) { return $null }
        $base = [Environment]::ExpandEnvironmentVariables($physicalPath)
        $cfg  = Join-Path -Path $base -ChildPath 'web.config'
        if (-not (Test-Path -LiteralPath $cfg)) { return $null }

        [xml]$xml = Get-Content -LiteralPath $cfg -Raw -ErrorAction Stop
        $node = $xml.SelectSingleNode('//configuration/appSettings/add[@key="useCOMFree"]')
        if ($null -ne $node) { return $node.value } else { return $null }
    } catch { return $null }
}

# ---------- Escrita ----------
function Set-NumeroProviders {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$physicalPath,
        [Parameter(Mandatory)][int]$NewValue
    )
    $result = [ordered]@{
        CaminhoFisico   = $physicalPath
        Arquivo         = $null
        ValorAntigo     = $null
        ValorNovo       = $NewValue
        Status          = 'erro'
        Mensagem        = ''
    }
    try {
        $base = [Environment]::ExpandEnvironmentVariables($physicalPath)
        $cfg  = Join-Path -Path $base -ChildPath 'pool.config'
        $result.Arquivo = $cfg

        if (-not (Test-Path -LiteralPath $cfg)) {
            $result.Status = 'ausente'
            $result.Mensagem = 'pool.config não encontrado'
            return [pscustomobject]$result
        }

        $content = Get-Content -LiteralPath $cfg -Raw -ErrorAction Stop
        $rx = New-Object System.Text.RegularExpressions.Regex(
            '<\s*maxload\s*>\s*([^<]+)\s*<\s*/\s*maxload\s*>',
            [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
        )
        $m = $rx.Match($content)
        if (-not $m.Success) {
            $result.Status = 'sem_tag'
            $result.Mensagem = 'Tag <maxload> não encontrada (nenhuma alteração).'
            return [pscustomobject]$result
        }

        $result.ValorAntigo = $m.Groups[1].Value.Trim()

        # Substitui o valor da tag (sem backup)
        $novo = $rx.Replace($content, "<maxload>$NewValue</maxload>")
        Set-Content -LiteralPath $cfg -Value $novo -Encoding UTF8 -ErrorAction Stop

        $result.Status = 'ok'
        $result.Mensagem = 'Atualizado com sucesso'
    } catch {
        $result.Status = 'erro'
        $result.Mensagem = $_.Exception.Message
    }
    return [pscustomobject]$result
}

function Set-UseCOMFree {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$physicalPath,
        [Parameter(Mandatory)][ValidateSet('True','False')][string]$NewValue
    )
    $result = [ordered]@{
        CaminhoFisico   = $physicalPath
        Arquivo         = $null
        ValorAntigo     = $null
        ValorNovo       = $NewValue
        Status          = 'erro'
        Mensagem        = ''
    }
    try {
        $base = [Environment]::ExpandEnvironmentVariables($physicalPath)
        $cfg  = Join-Path -Path $base -ChildPath 'web.config'
        $result.Arquivo = $cfg

        if (-not (Test-Path -LiteralPath $cfg)) {
            $result.Status = 'ausente'
            $result.Mensagem = 'web.config não encontrado'
            return [pscustomobject]$result
        }

        [xml]$xml = Get-Content -LiteralPath $cfg -Raw -ErrorAction Stop
        $node = $xml.SelectSingleNode('//configuration/appSettings/add[@key="useCOMFree"]')
        if ($null -eq $node) {
            $result.Status = 'sem_tag'
            $result.Mensagem = 'Tag appSettings/add[@key="useCOMFree"] não encontrada.'
            return [pscustomobject]$result
        }

        $result.ValorAntigo = $node.value
        $node.value = $NewValue
        $xml.Save($cfg)

        $result.Status = 'ok'
        $result.Mensagem = 'Atualizado com sucesso'
    } catch {
        $result.Status = 'erro'
        $result.Mensagem = $_.Exception.Message
    }
    return [pscustomobject]$result
}

# ---------- Helpers ----------
function Format-Binding([object]$binding) {
    $proto = $binding.protocol
    $info  = $binding.bindingInformation
    if ($proto -in @('http','https')) {
        $parts = $info -split ':', 3
        if ($parts.Count -eq 3) { return "{0}://{1}:{2} {3}" -f $proto,$parts[0],$parts[1],$parts[2] }
    }
    return "{0}://{1}" -f $proto,$info
}

function Get-PortsFromBindings {
    param($bindings)
    $ports = @()
    foreach ($b in $bindings) {
        $proto = $b.protocol
        $info  = $b.bindingInformation
        if ($proto -in @('http','https')) {
            $p = ($info -split ':',3)[1]
            if ($p) { $ports += $p }
        } elseif ($proto -eq 'net.tcp') {
            $p = ($info -split ':',2)[0]
            if ($p) { $ports += $p }
        }
    }
    return ($ports | Where-Object { $_ -and $_ -ne '*' } | Sort-Object -Unique) -join ', '
}

# ---------- Coleta (1 linha por Site + Applications) ----------
$linhas = foreach ($site in Get-Website) {
    $siteItem = Get-Item "IIS:\Sites\$($site.Name)"
    $bindings = $site.Bindings.Collection
    $portsTxt = Get-PortsFromBindings $bindings
    $btxt = if ($bindings.Count) {
        ($bindings | ForEach-Object { Format-Binding $_ }) -join "`n"
    } else { '' }

    # Site
    [pscustomobject]@{
        Site             = $site.Name
        Tipo             = 'Site'
        Caminho          = '/'
        AppPool          = $site.applicationPool
        Estado           = $site.State
        Portas           = $portsTxt
        CaminhoFisico    = $siteItem.physicalPath
        NumeroProviders  = ''
        UseCOMFree       = ''
        Bindings         = $btxt
    }

    # Applications
    foreach ($app in Get-WebApplication -Site $site.Name) {
        $appItem = Get-Item ("IIS:\Sites\{0}{1}" -f $site.Name,$app.Path)
        [pscustomobject]@{
            Site             = $site.Name
            Tipo             = 'Application'
            Caminho          = $app.Path
            AppPool          = $app.ApplicationPool
            Estado           = (Get-WebAppPoolState $app.ApplicationPool).Value
            Portas           = $portsTxt
            CaminhoFisico    = $appItem.physicalPath
            NumeroProviders  = Get-NumeroProviders -physicalPath $appItem.physicalPath
            UseCOMFree       = Get-UseCOMFree -physicalPath $appItem.physicalPath
            Bindings         = ''
        }
    }
}

# ---------- Visualização + seleção ----------
$selecionados = $linhas |
    Where-Object { $_.Tipo -eq 'Application' } |
    Sort-Object Site, Caminho |
    Select-Object Site, Caminho, AppPool, Estado, Portas, CaminhoFisico, NumeroProviders, UseCOMFree |
    Out-GridView -PassThru -Title 'Selecione os Applications a editar e clique OK'

if (-not $selecionados) { Write-Host 'Nenhum Application selecionado. Cancelado.'; return }

# ---------- Escolha o que alterar ----------
$acao = Read-Host "O que deseja alterar? [1] NumeroProviders | [2] useCOMFree | [3] Ambos | [0] Cancelar"
if ($acao -match '^[0]$') { Write-Host 'Cancelado.'; return }
$doNP = $acao -match '^(1|3)$'
$doCF = $acao -match '^(2|3)$'

# Monta nomes para exibir no prompt
$maxShow = 12
$names = $selecionados | ForEach-Object { "{0} {1}" -f $_.Site, $_.Caminho }
if ($names.Count -le $maxShow) { $showNames = $names -join [Environment]::NewLine }
else {
    $resto = $names.Count - $maxShow
    $showNames = @($names[0..($maxShow-1)] + "... (+" + $resto + " restante(s))") -join [Environment]::NewLine
}

# Captura dos novos valores
if ($doNP) {
    $promptNP = ("Novo NumeroProviders (maxload) para {0} item(ns):{1}{2}" -f $selecionados.Count, [Environment]::NewLine, $showNames)
    $newNP = Read-Host $promptNP
    if ([string]::IsNullOrWhiteSpace($newNP)) { Write-Host 'Valor vazio. Cancelado.'; return }
    try { [int]$newNP | Out-Null } catch { Write-Host 'NumeroProviders deve ser inteiro.'; return }
}

if ($doCF) {
    $promptCF = ("Novo valor de useCOMFree (True/False) para {0} item(ns) [ENTER=False]:{1}{2}" -f $selecionados.Count, [Environment]::NewLine, $showNames)
    $newCF = Read-Host $promptCF
    if ([string]::IsNullOrWhiteSpace($newCF)) { $newCF = 'False' }
    $newCF = $newCF.Trim()
    if ($newCF -notin @('True','False','true','false')) { Write-Host 'useCOMFree deve ser True ou False.'; return }
    # normaliza capitalização
    if ($newCF -match '^(?i:true)$') { $newCF = 'True' } else { $newCF = 'False' }
}

# ---------- Aplicar alterações ----------
$resultados = @()
foreach ($row in $selecionados) {
    if ($doNP) { $resultados += (Set-NumeroProviders -physicalPath $row.CaminhoFisico -NewValue $newNP | Add-Member -PassThru NoteProperty TipoAlteracao 'NumeroProviders') }
    if ($doCF) { $resultados += (Set-UseCOMFree   -physicalPath $row.CaminhoFisico -NewValue $newCF | Add-Member -PassThru NoteProperty TipoAlteracao 'useCOMFree') }
}

# ---------- Relatório ----------
$null = New-Item -ItemType Directory -Path "C:\Temp" -ErrorAction SilentlyContinue
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$csv = "C:\Temp\edicoes_iis_$timestamp.csv"
$resultados | Select-Object TipoAlteracao, CaminhoFisico, Arquivo, ValorAntigo, ValorNovo, Status, Mensagem |
    Export-Csv $csv -NoTypeInformation -Encoding UTF8
Write-Host "Relatório salvo em: $csv"

# ---------- Reciclar AppPools (opcional) ----------
$apppools = ($selecionados.AppPool | Sort-Object -Unique)
$resp = Read-Host "Reciclar os AppPools afetados? (`"$($apppools -join ', ')`") [S/N]"
if ($resp -match '^[sS]') {
    foreach ($p in $apppools) {
        try {
            Restart-WebAppPool -Name $p -ErrorAction Stop
            Write-Host ("AppPool {0} reciclado." -f $p) -ForegroundColor Green
        }
        catch {
            $msg = $_.Exception.Message
            Write-Host ("Falha ao reciclar {0}: {1}" -f $p, $msg) -ForegroundColor Yellow
        }
    }
}

# ---------- Conferência ----------
$verifica = foreach ($row in $selecionados) {
    [pscustomobject]@{
        Site            = $row.Site
        Caminho         = $row.Caminho
        AppPool         = $row.AppPool
        Portas          = $row.Portas
        CaminhoFisico   = $row.CaminhoFisico
        NP_Antes        = $row.NumeroProviders
        NP_Depois       = (Get-NumeroProviders -physicalPath $row.CaminhoFisico)
        COMFree_Antes   = $row.UseCOMFree
        COMFree_Depois  = (Get-UseCOMFree -physicalPath $row.CaminhoFisico)
        PoolConfig      = (Join-Path ([Environment]::ExpandEnvironmentVariables($row.CaminhoFisico)) 'pool.config')
        WebConfig       = (Join-Path ([Environment]::ExpandEnvironmentVariables($row.CaminhoFisico)) 'web.config')
    }
}
$verifica | Out-GridView -Title 'Conferência: NumeroProviders / useCOMFree (Antes x Depois)'
