# ===========================
# TestOracleMonitor.ps1
# Monitora o tempo de resposta de uma query Oracle a cada 3 segundos
# ===========================

# Connection string (substitua conforme necessário)
$connectionString = 'persist security info=true;data source=(DESCRIPTION=(SDU=65535)(SEND_BUF_SIZE=10485760)(RECV_BUF_SIZE=10485760)(ADDRESS=(PROTOCOL=TCP)(HOST=172.25.200.45)(PORT=1521))(CONNECT_DATA=(SERVER=DEDICATED)(SERVICE_NAME=DBPASA01_DBPASA01.subnet172252000.vcnclientesdrsp.oraclevcn.com)));user id=OCI200_AG_PROD;password="@@200agP#t4c202*";'

# Query a ser monitorada
$query = 'SELECT * FROM z_sistema'

# Arquivo de saída
$outputFile = "C:\temp\oracle_query_timings.txt"

# Caminho da DLL Oracle Managed Data Access
$assemblyPath = Join-Path -Path (Split-Path -Parent $PSCommandPath) -ChildPath "Oracle.ManagedDataAccess.dll"

# Carrega a DLL Oracle
if (-not (Test-Path $assemblyPath)) {
    try {
        Add-Type -AssemblyName "Oracle.ManagedDataAccess"
    } catch {
        Write-Host "Não foi possível carregar Oracle.ManagedDataAccess automaticamente." -ForegroundColor Yellow
        if (Test-Path $assemblyPath) {
            [Reflection.Assembly]::LoadFrom($assemblyPath) | Out-Null
        } else {
            throw "Oracle.ManagedDataAccess.dll não encontrado. Instale o pacote 'Oracle.ManagedDataAccess' ou copie a DLL para $assemblyPath"
        }
    }
} else {
    [Reflection.Assembly]::LoadFrom($assemblyPath) | Out-Null
}

# Função que executa a query e mede o tempo
function Measure-OracleQueryTime {
    param(
        [string]$ConnStr,
        [string]$Sql
    )

    $result = [ordered]@{
        Timestamp     = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        TTFB_ms       = $null
        TotalFetch_ms = $null
        RowCount      = 0
        Error         = $null
    }

    $conn = $null
    try {
        $conn = New-Object Oracle.ManagedDataAccess.Client.OracleConnection($ConnStr)
        $conn.Open()

        $cmd = $conn.CreateCommand()
        $cmd.CommandText = $Sql
        $cmd.CommandTimeout = 300

        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        $reader = $cmd.ExecuteReader()
        $ttfbMeasured = $false
        $firstRowTime = 0

        while ($reader.Read()) {
            if (-not $ttfbMeasured) {
                $firstRowTime = $sw.ElapsedMilliseconds
                $ttfbMeasured = $true
            }
            $result.RowCount++
        }

        if (-not $ttfbMeasured) {
            $firstRowTime = $sw.ElapsedMilliseconds
        }

        $sw.Stop()
        $result.TTFB_ms = $firstRowTime
        $result.TotalFetch_ms = $sw.ElapsedMilliseconds

        $reader.Close()
        $cmd.Dispose()
    } catch {
        $result.Error = $_.Exception.Message
    } finally {
        if ($conn -ne $null) {
            try { $conn.Close() } catch {}
            $conn.Dispose()
        }
    }

    return $result
}

# Garante que o diretório exista
$dir = Split-Path -Parent $outputFile
if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }

Write-Host "Iniciando monitoramento Oracle. Intervalo: 3 segundos (Ctrl + C para parar)" -ForegroundColor Cyan
Write-Host "Gravando logs em: $outputFile"
Write-Host "-------------------------------------------------------------"

# Loop infinito
while ($true) {
    $measurement = Measure-OracleQueryTime -ConnStr $connectionString -Sql $query

    $line = "Timestamp: $($measurement.Timestamp) | TTFB_ms: $($measurement.TTFB_ms) | TotalFetch_ms: $($measurement.TotalFetch_ms) | RowCount: $($measurement.RowCount) | Error: $($measurement.Error)"
    
    Add-Content -Path $outputFile -Value $line

    if ($measurement.Error) {
        Write-Host "[$($measurement.Timestamp)] ERRO: $($measurement.Error)" -ForegroundColor Red
    } else {
        Write-Host "[$($measurement.Timestamp)] OK - TTFB: $($measurement.TTFB_ms) ms | Total: $($measurement.TotalFetch_ms) ms | Rows: $($measurement.RowCount)" -ForegroundColor Green
    }

    Start-Sleep -Seconds 3
}
