<#
Synopsis: Automate IIS setup for multiple clients (App Pools, Site, WebApps) with idempotency, validations, logs, dry-run, and removal.
Requires: PowerShell 5+, Administrator privileges, Windows Server with IIS capability.
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$ConfigPath = 'E:\Benner\WES\config.json',
    [switch]$CreateFolders,
    [switch]$Remove,
    [string]$Cliente,
    [string]$AppPoolUser = 'bennercloud\service.wes.rh',
    [string]$AppPoolPasswordPlain = '7Zv87P9kxPVQdp'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Show-MainMenu {
    param([string]$DefaultConfigPath)
    Write-Host "================ IIS Automação ================" -ForegroundColor Cyan
    Write-Host "1) Provisionar/Atualizar a partir do JSON"
    Write-Host "2) Remover configurações de um cliente"
    Write-Host "3) Sair"
    $opt = Read-Host 'Escolha uma opção (1-3)'
    switch ($opt) {
        '1' {
            $cfg = Read-Host ("Caminho do config.json [Enter p/ padrão: {0}]" -f $DefaultConfigPath)
            if ([string]::IsNullOrWhiteSpace($cfg)) { $cfg = $DefaultConfigPath }
            return [pscustomobject]@{ Action = 'provision'; ConfigPath = $cfg }
        }
        '2' {
            $cli = Read-Host 'Nome do cliente para remoção'
            if ([string]::IsNullOrWhiteSpace($cli)) { return $null }
            return [pscustomobject]@{ Action = 'remove'; Cliente = $cli }
        }
        '3' { return [pscustomobject]@{ Action = 'exit' } }
        default { return $null }
    }
}

function Write-Info {
    param([string]$Message)
    Write-Host "[INFO] $Message"
}

function Write-Warn {
    param([string]$Message)
    Write-Warning $Message
}

function Ensure-Transcript {
    $logDir = 'E:\Benner\WES'
    $logPath = Join-Path $logDir 'setup_iis.log'
    try {
        if (-not (Test-Path -LiteralPath $logDir)) {
            if ($PSCmdlet.ShouldProcess($logDir, 'Create transcript directory')) {
                New-Item -ItemType Directory -Force -Path $logDir | Out-Null
            }
        }
        Start-Transcript -Path $logPath -Append -ErrorAction SilentlyContinue | Out-Null
    } catch {
        Write-Warn "Falha ao iniciar transcript em ${logPath}: $($_.Exception.Message)"
    }
}

 

function Import-IISModule {
    if (-not (Get-Module -ListAvailable -Name WebAdministration)) {
        throw [System.InvalidOperationException]::new('Módulo WebAdministration não disponível. Verifique a instalação do IIS.')
    }
    Import-Module WebAdministration -ErrorAction Stop
}

function Get-CredentialFromFile {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        throw [System.IO.FileNotFoundException]::new("Arquivo de credencial não encontrado: $Path")
    }
    $cred = Import-Clixml -Path $Path
    if (-not ($cred -is [System.Management.Automation.PSCredential])) {
        throw [System.InvalidCastException]::new('O arquivo de credencial não contém um PSCredential válido.')
    }
    return $cred
}

function Get-ServerManager {
    $loaded = $false
    $candidatePaths = @(
        (Join-Path $env:windir 'System32\inetsrv\Microsoft.Web.Administration.dll'),
        (Join-Path $env:windir 'SysWOW64\inetsrv\Microsoft.Web.Administration.dll')
    )
    try {
        Add-Type -AssemblyName 'Microsoft.Web.Administration' -ErrorAction Stop | Out-Null
        $loaded = $true
    } catch {}
    if (-not $loaded) {
        foreach ($p in $candidatePaths) {
            if (Test-Path -LiteralPath $p) {
                try {
                    Add-Type -Path $p -ErrorAction Stop | Out-Null
                    $loaded = $true
                    break
                } catch {}
            }
        }
    }
    if (-not $loaded) {
        $pathsTried = $candidatePaths -join ', '
        throw [System.InvalidOperationException]::new("Não foi possível carregar 'Microsoft.Web.Administration'. Verifique se o IIS e as ferramentas de gerenciamento estão instalados. Caminhos verificados: $pathsTried")
    }
    return [Microsoft.Web.Administration.ServerManager]::new()
}

function Get-ClientConfigsFromIIS {
    $sm = Get-ServerManager
    $allPools = @($sm.ApplicationPools)
    $allSites = @($sm.Sites)
    $prodPools = $allPools | Where-Object { -not ($_.Name.ToUpperInvariant().EndsWith('_HML')) }
    $clients = @()
    foreach ($pp in $prodPools) {
        $clientName = $pp.Name
        $hmlName = "${clientName}_HML"
        $clients += [pscustomobject]@{
            ClientName = $clientName
            ProdPool = $clientName
            HmlPool = $hmlName
            ProdPoolExists = $true
            HmlPoolExists = ($allPools | Where-Object { $_.Name -eq $hmlName }) -ne $null
            SiteExists = ($allSites | Where-Object { $_.Name -eq $clientName }) -ne $null
        }
    }
    # Include orphan HML-only pools as separate entries if no prod exists
    $hmlOnly = $allPools | Where-Object { $_.Name.ToUpperInvariant().EndsWith('_HML') } |
        Where-Object { $prodPools.Name -notcontains ($_.Name -replace '_HML$', '') }
    foreach ($hp in $hmlOnly) {
        $base = $hp.Name -replace '_HML$', ''
        $clients += [pscustomobject]@{
            ClientName = $base
            ProdPool = $base
            HmlPool = $hp.Name
            ProdPoolExists = $false
            HmlPoolExists = $true
            SiteExists = ($allSites | Where-Object { $_.Name -eq $base }) -ne $null
        }
    }
    return $clients | Sort-Object -Property ClientName -Unique
}

function Show-ClientRemovalMenu {
    $items = Get-ClientConfigsFromIIS
    if (-not $items -or $items.Count -eq 0) {
        Write-Host 'Nenhuma configuração de cliente encontrada no IIS.' -ForegroundColor Yellow
        return @()
    }
    Write-Host 'Selecione as configurações (clientes) para remover:' -ForegroundColor Cyan
    for ($i = 0; $i -lt $items.Count; $i++) {
        $it = $items[$i]
        $details = @()
        if ($it.ProdPoolExists) { $details += "Pool=$($it.ProdPool)" }
        if ($it.HmlPoolExists) { $details += "Pool=$($it.HmlPool)" }
        if ($it.SiteExists) { $details += "Site=$($it.ClientName)" }
        $summary = if ($details.Count -gt 0) { $details -join ', ' } else { 'Sem objetos IIS encontrados' }
        Write-Host ("[{0}] {1}  ->  {2}" -f ($i+1), $it.ClientName, $summary)
    }
    while ($true) {
        $inputStr = Read-Host 'Digite os números separados por vírgula (ex.: 1,2,3)'
        if ([string]::IsNullOrWhiteSpace($inputStr)) { continue }
        $parts = $inputStr -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }
        $nums = @()
        $valid = $true
        foreach ($p in $parts) {
            $n = 0
            if (-not [int]::TryParse($p, [ref]$n)) { $valid = $false; break }
            if ($n -lt 1 -or $n -gt $items.Count) { $valid = $false; break }
            $nums += $n
        }
        if (-not $valid) {
            Write-Host 'Entrada inválida. Use somente números dentro do intervalo, separados por vírgula.' -ForegroundColor Yellow
            continue
        }
        $selectedClients = $nums | Sort-Object -Unique | ForEach-Object { $items[$_-1].ClientName }
        return $selectedClients
    }
}

function Assert-UniquePort {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][int]$Port,
        [Parameter(Mandatory)][string]$SiteName
    )
    $sm = Get-ServerManager
    foreach ($site in $sm.Sites) {
        foreach ($binding in $site.Bindings) {
            if ($binding.Protocol -eq 'http') {
                $parts = $binding.BindingInformation -split ':'
                if ($parts.Length -ge 2) {
                    $portStr = $parts[$parts.Length - 2]
                    $parsedPort = 0
                    if ([int]::TryParse($portStr, [ref]$parsedPort)) {
                        if ($parsedPort -eq $Port -and $site.Name -ne $SiteName) {
                            throw [System.InvalidOperationException]::new("Porta $Port já está em uso pelo site '$($site.Name)'.")
                        }
                    }
                }
            }
        }
    }
}

function Assert-UniqueNamesInConfig {
    param([object[]]$Clients)
    $nameGroups = $Clients | Group-Object -Property nome
    foreach ($g in $nameGroups) {
        if ($g.Count -gt 1) { throw [System.InvalidOperationException]::new("Nome de cliente duplicado no config: '$($g.Name)'.") }
    }
    $portGroups = $Clients | Group-Object -Property porta
    foreach ($g in $portGroups) {
        if ($g.Count -gt 1) { throw [System.InvalidOperationException]::new("Porta duplicada no config: $($g.Name).") }
    }
}

function Assert-UniquePaths {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object[]]$PlannedApps
        # Each item: @{ SiteName = 'NAME'; AppName = 'NAME'; PhysicalPath = 'PATH' }
    )
    $sm = Get-ServerManager

    $existingPathMap = @{}
    foreach ($site in $sm.Sites) {
        foreach ($app in $site.Applications) {
            if ($app.Path -ne '/') {
                $vd = $app.VirtualDirectories | Where-Object { $_.Path -eq '/' } | Select-Object -First 1
                if ($vd) {
                    $p = [IO.Path]::GetFullPath($vd.PhysicalPath)
                    $key = ($p.ToLowerInvariant())
                    if (-not $existingPathMap.ContainsKey($key)) {
                        $existingPathMap[$key] = @()
                    }
                    $existingPathMap[$key] += @{ SiteName = $site.Name; AppName = ($app.Path.TrimStart('/')) }
                }
            }
        }
    }

    # Within planned list
    $plannedGroups = $PlannedApps | Group-Object -Property PhysicalPath
    foreach ($grp in $plannedGroups) {
        if ($grp.Count -gt 1) {
            throw [System.InvalidOperationException]::new("Caminho físico duplicado no plano: $($grp.Name)")
        }
    }

    foreach ($plan in $PlannedApps) {
        $pp = [IO.Path]::GetFullPath($plan.PhysicalPath)
        $k = $pp.ToLowerInvariant()
        if ($existingPathMap.ContainsKey($k)) {
            $references = $existingPathMap[$k]
            $conflict = $references | Where-Object { $_.SiteName -ne $plan.SiteName -or $_.AppName -ne $plan.AppName }
            if ($conflict) {
                $desc = ($conflict | ForEach-Object { "Site='$($_.SiteName)' App='$($_.AppName)'" }) -join '; '
                throw [System.InvalidOperationException]::new("Caminho físico '$pp' já utilizado por: $desc")
            }
        }
    }
}

function Ensure-FolderIfNeeded {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        if ($PSCmdlet.ShouldProcess($Path, 'Criar diretório')) {
            New-Item -ItemType Directory -Force -Path $Path | Out-Null
            Write-Info "Diretório criado: $Path"
        }
    } else {
        Write-Verbose "Diretório já existe: $Path"
    }
}

function Get-AppPreload {
    param([Microsoft.Web.Administration.Application]$Application)
    try {
        return [bool]$Application.PreloadEnabled
    } catch {
        try {
            $val = $Application.GetAttributeValue('preloadEnabled')
            if ($null -ne $val) { return [bool]$val } else { return $false }
        } catch { return $false }
    }
}

function Ensure-AppPreload {
    param(
        [Microsoft.Web.Administration.Application]$Application,
        [bool]$Desired
    )
    $current = Get-AppPreload -Application $Application
    if ($current -ne $Desired) {
        if ($PSCmdlet.ShouldProcess($Application.Path, 'Atualizar preloadEnabled')) {
            try { $Application.PreloadEnabled = $Desired } catch { $Application.SetAttributeValue('preloadEnabled', $Desired) }
            return $true
        }
    }
    return $false
}

function Configure-AppPoolInternal {
    param(
        [Microsoft.Web.Administration.ApplicationPool]$Pool,
        [System.Management.Automation.PSCredential]$Credential,
        [bool]$IsProduction,
        [Nullable[TimeSpan]]$RecycleSchedule
    )
    $changed = $false

    if (-not $Pool) { throw [System.ArgumentNullException]::new('Pool') }

    if ($Pool.Enable32BitAppOnWin64 -ne $true) { $Pool.Enable32BitAppOnWin64 = $true; $changed = $true }

    # Identity
    $desiredUser = $Credential.UserName
    if ($Pool.ProcessModel.IdentityType -ne [Microsoft.Web.Administration.ProcessModelIdentityType]::SpecificUser) {
        $Pool.ProcessModel.IdentityType = [Microsoft.Web.Administration.ProcessModelIdentityType]::SpecificUser; $changed = $true
    }
    $mustSetPassword = $false
    if ($Pool.ProcessModel.UserName -ne $desiredUser) { $Pool.ProcessModel.UserName = $desiredUser; $changed = $true; $mustSetPassword = $true }
    if ($changed -and -not $mustSetPassword -and $Pool.ProcessModel.IdentityType -eq [Microsoft.Web.Administration.ProcessModelIdentityType]::SpecificUser) {
        # IdentityType change implies we should reapply password once
        $mustSetPassword = $true
    }
    if ($mustSetPassword) {
        $Pool.ProcessModel.Password = $Credential.GetNetworkCredential().Password
    }

    if ($IsProduction) {
        if ($Pool.StartMode -ne [Microsoft.Web.Administration.StartMode]::AlwaysRunning) { $Pool.StartMode = [Microsoft.Web.Administration.StartMode]::AlwaysRunning; $changed = $true }
        if ($Pool.ProcessModel.IdleTimeout -ne [TimeSpan]::Zero) { $Pool.ProcessModel.IdleTimeout = [TimeSpan]::Zero; $changed = $true }
        if ($Pool.ProcessModel.IdleTimeoutAction -ne [Microsoft.Web.Administration.IdleTimeoutAction]::Suspend) { $Pool.ProcessModel.IdleTimeoutAction = [Microsoft.Web.Administration.IdleTimeoutAction]::Suspend; $changed = $true }
        if ($Pool.ProcessModel.LoadUserProfile -ne $true) { $Pool.ProcessModel.LoadUserProfile = $true; $changed = $true }
        if ($Pool.ProcessModel.PingResponseTime -ne [TimeSpan]::FromSeconds(300)) { $Pool.ProcessModel.PingResponseTime = [TimeSpan]::FromSeconds(300); $changed = $true }
        if ($Pool.ProcessModel.PingInterval -ne [TimeSpan]::FromSeconds(180)) { $Pool.ProcessModel.PingInterval = [TimeSpan]::FromSeconds(180); $changed = $true }
        if ($Pool.ProcessModel.ShutdownTimeLimit -ne [TimeSpan]::FromSeconds(180)) { $Pool.ProcessModel.ShutdownTimeLimit = [TimeSpan]::FromSeconds(180); $changed = $true }
        if ($Pool.ProcessModel.StartupTimeLimit -ne [TimeSpan]::FromSeconds(300)) { $Pool.ProcessModel.StartupTimeLimit = [TimeSpan]::FromSeconds(300); $changed = $true }
        if ($Pool.Recycling.PeriodicRestart.Time -ne [TimeSpan]::Zero) { $Pool.Recycling.PeriodicRestart.Time = [TimeSpan]::Zero; $changed = $true }
        # Staggered schedule per production pool if provided
        if ($null -ne $RecycleSchedule) {
            $currentSchedule = @($Pool.Recycling.PeriodicRestart.Schedule | ForEach-Object { $_.Time })
            if (@($currentSchedule).Count -ne 1 -or $currentSchedule[0] -ne $RecycleSchedule) {
                $Pool.Recycling.PeriodicRestart.Schedule.Clear()
                [void]$Pool.Recycling.PeriodicRestart.Schedule.Add($RecycleSchedule)
                $changed = $true
            }
        }
        # Disable recycle event logging (0 = None)
        if ([int]$Pool.Recycling.LogEventOnRecycle -ne 0) {
            $Pool.Recycling.LogEventOnRecycle = 0
            $changed = $true
        }
    } else {
        # Homologacao defaults
        if ($Pool.ProcessModel.IdleTimeout -ne [TimeSpan]::FromMinutes(10)) { $Pool.ProcessModel.IdleTimeout = [TimeSpan]::FromMinutes(10); $changed = $true }
        if ($Pool.ProcessModel.LoadUserProfile -ne $true) { $Pool.ProcessModel.LoadUserProfile = $true; $changed = $true }
    }

    return $changed
}

function Ensure-AppPool {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][System.Management.Automation.PSCredential]$Credential,
        [Parameter(Mandatory)][bool]$IsProduction,
        [Nullable[TimeSpan]]$RecycleSchedule
    )
    $sm = Get-ServerManager
    $pool = $sm.ApplicationPools[$Name]
    $created = $false
    if (-not $pool) {
        if ($PSCmdlet.ShouldProcess($Name, 'Criar Application Pool')) {
            $pool = $sm.ApplicationPools.Add($Name)
            $created = $true
        }
    }
    $changed = $false
    if ($pool) {
        $changed = Configure-AppPoolInternal -Pool $pool -Credential $Credential -IsProduction:$IsProduction -RecycleSchedule $RecycleSchedule
    }
    if ($created -or $changed) {
        if ($PSCmdlet.ShouldProcess($Name, 'Aplicar alterações no Application Pool')) {
            $sm.CommitChanges()
        }
    }
    return @{ Name = $Name; Created = $created; Updated = ($changed -and -not $created); NoChange = (-not $created -and -not $changed) }
}

function Ensure-Site {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][int]$Port,
        [Parameter(Mandatory)][string]$PhysicalPath,
        [Parameter(Mandatory)][string]$AppPoolName,
        [Parameter()][bool]$EnsureRootPreload = $true
    )
    Assert-UniquePort -Port $Port -SiteName $Name
    Ensure-FolderIfNeeded -Path $PhysicalPath

    $sm = Get-ServerManager
    $site = $sm.Sites[$Name]
    $created = $false
    if (-not $site) {
        if ($PSCmdlet.ShouldProcess($Name, 'Criar Site')) {
            $bindingInfo = "*:$($Port):"
            $site = $sm.Sites.Add($Name, 'http', $bindingInfo, $PhysicalPath)
            $created = $true
        }
    }

    $changed = $false
    if ($site) {
        # Para sites existentes: preservar bindings atuais; apenas garantir que exista um binding na porta desejada
        $hasDesiredPort = $false
        foreach ($b in ($site.Bindings | Where-Object { $_.Protocol -eq 'http' })) {
            $parts = $b.BindingInformation -split ':'
            if ($parts.Length -ge 2) {
                $portStr = $parts[$parts.Length - 2]
                if ($portStr -eq "$Port") { $hasDesiredPort = $true; break }
            }
        }
        if (-not $hasDesiredPort) {
            if ($PSCmdlet.ShouldProcess($Name, 'Adicionar binding HTTP para a porta configurada')) {
                [void]$site.Bindings.Add("*:$($Port):", 'http')
                $changed = $true
            }
        }

        # Ensure physical path
        $rootApp = $site.Applications['/']
        $rootVDir = $rootApp.VirtualDirectories['/']
        if ($rootVDir.PhysicalPath -ne $PhysicalPath) {
            if ($PSCmdlet.ShouldProcess($Name, "Atualizar physicalPath do site para '$PhysicalPath'")) {
                $rootVDir.PhysicalPath = $PhysicalPath
                $changed = $true
            }
        }

        # Ensure app pool
        if ($rootApp.ApplicationPoolName -ne $AppPoolName) {
            if ($PSCmdlet.ShouldProcess($Name, "Vincular site ao App Pool '$AppPoolName'")) {
                $rootApp.ApplicationPoolName = $AppPoolName
                $changed = $true
            }
        }

        # Ensure preload on root application
        if ($EnsureRootPreload) {
            if (Ensure-AppPreload -Application $rootApp -Desired $true) { $changed = $true }
        }
    }

    if ($created -or $changed) {
        if ($PSCmdlet.ShouldProcess($Name, 'Aplicar alterações no Site')) { $sm.CommitChanges() }
    }
    return @{ Name = $Name; Created = $created; Updated = ($changed -and -not $created); NoChange = (-not $created -and -not $changed) }
}

function Ensure-App {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SiteName,
        [Parameter(Mandatory)][string]$AppName,
        [Parameter(Mandatory)][string]$PhysicalPath,
        [Parameter(Mandatory)][string]$AppPoolName,
        [Parameter(Mandatory)][bool]$PreloadEnabled
    )
    Ensure-FolderIfNeeded -Path $PhysicalPath
    $sm = Get-ServerManager
    $site = $sm.Sites[$SiteName]
    if (-not $site) { throw [System.InvalidOperationException]::new("Site '$SiteName' não encontrado para criar app '$AppName'.") }
    $appPath = '/' + $AppName
    $app = $site.Applications[$appPath]
    $created = $false
    if (-not $app) {
        if ($PSCmdlet.ShouldProcess($AppName, "Criar WebApp em '$SiteName'")) {
            $app = $site.Applications.Add($appPath, $PhysicalPath)
            $app.ApplicationPoolName = $AppPoolName
            try { $app.PreloadEnabled = $PreloadEnabled } catch { $app.SetAttributeValue('preloadEnabled', $PreloadEnabled) }
            $created = $true
        }
    }
    $changed = $false
    if ($app) {
        $vdir = $app.VirtualDirectories['/']
        if ($vdir.PhysicalPath -ne $PhysicalPath) {
            if ($PSCmdlet.ShouldProcess($AppName, 'Atualizar PhysicalPath')) {
                $vdir.PhysicalPath = $PhysicalPath
                $changed = $true
            }
        }
        if ($app.ApplicationPoolName -ne $AppPoolName) {
            if ($PSCmdlet.ShouldProcess($AppName, 'Atualizar App Pool')) {
                $app.ApplicationPoolName = $AppPoolName
                $changed = $true
            }
        }
        if (Ensure-AppPreload -Application $app -Desired $PreloadEnabled) { $changed = $true }
    }
    if ($created -or $changed) {
        if ($PSCmdlet.ShouldProcess($AppName, 'Aplicar alterações no WebApp')) { $sm.CommitChanges() }
    }
    return @{ Name = $AppName; Created = $created; Updated = ($changed -and -not $created); NoChange = (-not $created -and -not $changed) }
}

function Remove-Client {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$ClientName)
    $sm = Get-ServerManager
    $site = $sm.Sites[$ClientName]
    $prodPoolName = $ClientName
    $hmlPoolName = "${ClientName}_HML"
    $removed = @{ Site = $false; ProdPool = $false; HmlPool = $false }

    if ($site) {
        if ($PSCmdlet.ShouldProcess($ClientName, 'Remover Site (inclui webapps e bindings)')) {
            $sm.Sites.Remove($site)
            $removed.Site = $true
        }
    } else {
        Write-Verbose "Site '$ClientName' não existe."
    }

    $prodPool = $sm.ApplicationPools[$prodPoolName]
    if ($prodPool) {
        if ($PSCmdlet.ShouldProcess($prodPoolName, 'Remover App Pool de Produção')) {
            try { $prodPool.Stop() } catch {}
            $sm.ApplicationPools.Remove($prodPool)
            $removed.ProdPool = $true
        }
    } else {
        Write-Verbose "App Pool '$prodPoolName' não existe."
    }

    $hmlPool = $sm.ApplicationPools[$hmlPoolName]
    if ($hmlPool) {
        if ($PSCmdlet.ShouldProcess($hmlPoolName, 'Remover App Pool de Homologação')) {
            try { $hmlPool.Stop() } catch {}
            $sm.ApplicationPools.Remove($hmlPool)
            $removed.HmlPool = $true
        }
    } else {
        Write-Verbose "App Pool '$hmlPoolName' não existe."
    }

    if ($removed.Site -or $removed.ProdPool -or $removed.HmlPool) {
        if ($PSCmdlet.ShouldProcess($ClientName, 'Aplicar remoções')) { $sm.CommitChanges() }
    }
    return $removed
}

function Get-DefaultBasePath {
    param([object]$Config)
    if ($Config -and $Config.basePath) { return $Config.basePath }
    return 'E:\Benner\WES'
}

function Normalize-AppPath {
    param([string]$BasePath, [string]$ClientFolder, [string]$NomeProduto, [string]$Env)
    $productFolder = $NomeProduto
    if ($Env -eq 'homologacao') {
        if (-not $productFolder.ToUpperInvariant().EndsWith('_HML')) { $productFolder = "$productFolder`_HML" }
    }
    return (Join-Path (Join-Path $BasePath $ClientFolder) $productFolder)
}

Ensure-Transcript

try {
    Import-IISModule

    # Interactive menu when no explicit action parameters are provided
    $selected = $null
    $invokedByParams = $Remove.IsPresent -or $PSBoundParameters.ContainsKey('ConfigPath')
    if (-not $invokedByParams) {
        $selected = Show-MainMenu -DefaultConfigPath $ConfigPath
        if (-not $selected) { Write-Host 'Opção inválida.' -ForegroundColor Yellow; return }
        if ($selected.Action -eq 'exit') { return }
        if ($selected.Action -eq 'remove') {
            $Remove = $true
            # Interactive selection of clients from IIS if user did not type a name
            if ([string]::IsNullOrWhiteSpace($selected.Cliente)) {
                $toRemove = Show-ClientRemovalMenu
                if (-not $toRemove -or $toRemove.Count -eq 0) { return }
                foreach ($cn in $toRemove) {
                    $res = Remove-Client -ClientName $cn
                    Write-Info "Remoção concluída para '$cn' (Site=$($res.Site), ProdPool=$($res.ProdPool), HmlPool=$($res.HmlPool))"
                }
                try { Read-Host 'Pressione Enter para finalizar...' | Out-Null } catch {}
                return
            } else {
                $Cliente = $selected.Cliente
            }
        }
        if ($selected.Action -eq 'provision') { $ConfigPath = $selected.ConfigPath }        
    }

    if ($Remove) {
        if (-not $Cliente) { throw [System.ArgumentException]::new('Ao usar -Remove, forneça também -Cliente <NOME>.') }
        $res = Remove-Client -ClientName $Cliente
        Write-Info "Remoção concluída para '$Cliente' (Site=$($res.Site), ProdPool=$($res.ProdPool), HmlPool=$($res.HmlPool))"
        return
    }

    if (-not (Test-Path -LiteralPath $ConfigPath)) {
        throw [System.IO.FileNotFoundException]::new("Arquivo de configuração não encontrado: $ConfigPath")
    }

    $configRaw = Get-Content -LiteralPath $ConfigPath -Raw -ErrorAction Stop
    $config = $configRaw | ConvertFrom-Json

    if (-not $config) { throw [System.InvalidOperationException]::new('Config JSON inválido.') }
    if (-not $config.clientes -or $config.clientes.Count -eq 0) { throw [System.InvalidOperationException]::new('Nenhum cliente definido em config.clientes.') }

    $basePath = Get-DefaultBasePath -Config $config
    $secure = ConvertTo-SecureString -String $AppPoolPasswordPlain -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential($AppPoolUser, $secure)

    Assert-UniqueNamesInConfig -Clients $config.clientes

    $overallSummary = @()

    foreach ($cli in $config.clientes) {
        $clientName = [string]$cli.nome
        if ([string]::IsNullOrWhiteSpace($clientName)) { throw [System.ArgumentException]::new('Cliente.nome inválido no JSON.') }
        $port = [int]$cli.porta
        $rootPathName = if ($cli.rootPathName) { [string]$cli.rootPathName } else { $clientName }
        $sitePhysicalPath = Join-Path $basePath $rootPathName

        $prodPoolName = $clientName
        $hmlPoolName = "${clientName}_HML"

        # Plan apps and validate env
        $plannedApps = @()
        foreach ($a in $cli.apps) {
            $env = [string]$a.env
            if (@('producao', 'homologacao') -notcontains $env) {
                throw [System.ArgumentException]::new("Cliente '$clientName': env inválido em app '$($a.nomeProduto)': '$env'. Use 'producao' ou 'homologacao'.")
            }
            $appName = [string]$a.nomeProduto
            $physicalPath = if ($a.pathOverride) { [string]$a.pathOverride } else { Normalize-AppPath -BasePath $basePath -ClientFolder $rootPathName -NomeProduto $a.nomeProduto -Env $env }
            $preload = if ($null -ne $a.preloadOverride) { [bool]$a.preloadOverride } else { if ($env -eq 'producao') { $true } else { $false } }
            $targetPool = if ($env -eq 'producao') { $prodPoolName } else { $hmlPoolName }
            $plannedApps += [pscustomobject]@{ SiteName = $clientName; AppName = $appName; PhysicalPath = $physicalPath; AppPool = $targetPool; Preload = $preload }
        }

        # Validate paths uniqueness (planned vs existing)
        Assert-UniquePaths -PlannedApps $plannedApps

        # Ensure site root folder when needed
        Ensure-FolderIfNeeded -Path $sitePhysicalPath

        # App Pools
        # Calculate staggered recycle schedule for production pool: start 01:00, add 15 minutes per client index
        $clientIndex = [array]::IndexOf($config.clientes, $cli)
        if ($clientIndex -lt 0) { $clientIndex = 0 }
        $staggerMinutes = 15 * $clientIndex
        $prodRecycleSchedule = [TimeSpan]::FromHours(1).Add([TimeSpan]::FromMinutes($staggerMinutes))

        $prodRes = Ensure-AppPool -Name $prodPoolName -Credential $credential -IsProduction:$true -RecycleSchedule $prodRecycleSchedule
        $hmlRes  = Ensure-AppPool -Name $hmlPoolName  -Credential $credential -IsProduction:$false -RecycleSchedule $null

        # Site
        $siteRes = Ensure-Site -Name $clientName -Port $port -PhysicalPath $sitePhysicalPath -AppPoolName $prodPoolName -EnsureRootPreload:$true

        # Apps
        $appSummaries = @()
        foreach ($p in $plannedApps) {
            $appRes = Ensure-App -SiteName $p.SiteName -AppName $p.AppName -PhysicalPath $p.PhysicalPath -AppPoolName $p.AppPool -PreloadEnabled $p.Preload
            $appSummaries += $appRes
        }

        $overallSummary += [pscustomobject]@{
            Cliente = $clientName
            Porta = $port
            SitePath = $sitePhysicalPath
            Pools = @($prodRes, $hmlRes)
            Site = $siteRes
            Apps = $appSummaries
        }
    }

    # Output final summary per client
    foreach ($s in $overallSummary) {
        Write-Host "==== Cliente: $($s.Cliente) ====" -ForegroundColor Cyan
        Write-Host ("Porta: {0}  | SitePath: {1}" -f $s.Porta, $s.SitePath)
        foreach ($p in $s.Pools) {
            $status = if ($p.Created) { 'Created' } elseif ($p.Updated) { 'Updated' } else { 'No change' }
            Write-Host ("AppPool {0}: {1}" -f $p.Name, $status)
        }
        $siteStatus = if ($s.Site.Created) { 'Created' } elseif ($s.Site.Updated) { 'Updated' } else { 'No change' }
        Write-Host ("Site {0}: {1}" -f $s.Cliente, $siteStatus)
        foreach ($a in $s.Apps) {
            $astatus = if ($a.Created) { 'Created' } elseif ($a.Updated) { 'Updated' } else { 'No change' }
            Write-Host ("App {0}: {1}" -f $a.Name, $astatus)
        }
        Write-Host "" # spacer
    }
    # Pause on success to keep console open
    try { Read-Host 'Pressione Enter para finalizar...' | Out-Null } catch {}

} catch {
    $msg = $_.Exception.Message
    Write-Host ("Erro: {0}" -f $msg) -ForegroundColor Red
    try { Stop-Transcript | Out-Null } catch {}
    try { Read-Host 'Pressione Enter para finalizar...' | Out-Null } catch {}
    return
} finally {
    try { Stop-Transcript | Out-Null } catch {}
}


