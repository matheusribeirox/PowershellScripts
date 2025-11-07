# Verifica a versão do Windows Server
$osInfo = Get-CimInstance Win32_OperatingSystem
$version = [System.Version]$osInfo.Version

Write-Host "Versão do SO detectada: $($osInfo.Caption) ($($osInfo.Version))" -ForegroundColor Cyan

if ($version.Major -lt 10) {
    Write-Error "Este script suporta apenas Windows Server 2016 ou superior."
    exit 1
}

# Lista de Features para instalar
$features = @(
    "NET-Framework-Features",      # .NET Framework 3.5 (inclui 2.0 e 3.0)
    "NET-Framework-45-Features",   # .NET Framework 4.7/4.8
    "NET-WCF-HTTP-Activation45",   # WCF HTTP Activation
    "NET-WCF-TCP-Activation45",    # WCF TCP Activation
    "NET-WCF-Pipe-Activation45",   # Named Pipes Activation
    "NET-WCF-TCP-PortSharing45",   # TCP Port Sharing
    "Web-Server",                  # IIS Base
    "Web-WebServer",
    "Web-Common-Http",
    "Web-Default-Doc",
    "Web-Dir-Browsing",
    "Web-Http-Errors",
    "Web-Static-Content",
    "Web-Health",
    "Web-Http-Logging",
    "Web-Log-Libraries",
    "Web-Request-Monitor",
    "Web-Http-Tracing",
    "Web-Performance",
    "Web-Stat-Compression",
    "Web-Dyn-Compression",
    "Web-App-Dev",
    "Web-Asp-Net",
    "Web-Asp-Net45",
    "Web-Net-Ext",
    "Web-Net-Ext45",
    "Web-ISAPI-Ext",
    "Web-ISAPI-Filter",
    "Web-Security",
    "Web-Basic-Auth",
    "Web-Filtering",
    "Web-Mgmt-Tools",
    "Web-Mgmt-Console",
    "Web-WHC",                # World Wide Web Services
    "Web-WMI",                # Ferramentas de Gerenciamento da Web
    "WAS",                    # Windows Process Activation Service
    "WAS-Process-Model",
    "WAS-NET-Environment",
    "WAS-Config-APIs"
)

# Função para instalar features
function Install-FeatureList {
    param(
        [string[]]$featuresList
    )

    $total = $featuresList.Count
    $current = 0

    foreach ($feature in $featuresList) {
        $current++
        Write-Progress -Activity "Instalando features" -Status "Instalando $feature ($current/$total)" -PercentComplete (($current/$total)*100)

        $result = Install-WindowsFeature -Name $feature -ErrorAction SilentlyContinue
        if ($result.Success -eq $false) {
            Write-Warning "Falha ao instalar a feature: $feature"
        } else {
            Write-Host "Feature instalada com sucesso: $feature" -ForegroundColor Green
        }
    }
}

# Instalação das features
try {
    Install-FeatureList -featuresList $features
    Write-Host "Processo de instalação finalizado!" -ForegroundColor Cyan
    Write-Host "Algumas features podem requerer reinicialização." -ForegroundColor Yellow
} catch {
    Write-Error "Erro durante a instalação: $_"
    exit 1
}
