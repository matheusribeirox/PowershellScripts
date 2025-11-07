# Defina os nomes ou endereços dos servidores que deseja testar a comunicação
$Server2 = "172.16.0.160"

# Caminho para o arquivo de log
$LogFile = "C:\Benner\logTestConnection.txt"

# Loop infinito para testar a comunicação a cada 1 segundo
while ($true) {
    $CurrentDateTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    
    # Tente testar a comunicação usando o comando Test-Connection
    try {
        $TestResult = Test-Connection -ComputerName $Server2 -Count 1 -ErrorAction Stop
        $IsReachable = $true
        $ResponseTimeMs = $TestResult.ResponseTime
    } catch {
        $IsReachable = $false
        $ErrorMessage = $_.Exception.Message
        $ResponseTimeMs = "N/A"
    }
    
    # Registre o resultado no arquivo de log
    if ($IsReachable) {
        Add-Content -Path $LogFile -Value "$CurrentDateTime - A comunicação com $Server2 está OK. Tempo de resposta: $ResponseTimeMs ms"
        Write-Host "$CurrentDateTime | Connection Success. $ResponseTimeMs ms" -ForegroundColor Green
    } else {
        Add-Content -Path $LogFile -Value "$CurrentDateTime - A comunicação com $Server2 falhou. Erro: $ErrorMessage"
        Write-Host "$CurrentDateTime | Connection Failed." -ForegroundColor Red
    }
    
    # Espere 1 segundo antes da próxima verificação
    Start-Sleep -Seconds 1
}
