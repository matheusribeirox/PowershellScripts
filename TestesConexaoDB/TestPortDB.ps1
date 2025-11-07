# Parâmetros de conexão ao SQL Server
$serverName = "172.16.0.160"
$port = 1433

# Caminho e nome do arquivo de log
$logPath = "C:\Benner\logPortaBD.txt"

function TestarConexaoELogin {

# Testar a conexão usando o comando Test-NetConnection
$connectionTest = Test-NetConnection -ComputerName $serverName -Port $port
if ($connectionTest.TcpTestSucceeded) {
    Write-Output "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | Teste de conexão bem-sucedido para $serverName na porta $port."
    # Registrar o teste de conexão bem-sucedido no log
    Add-Content -Path $logPath -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') Teste de conexão bem-sucedido para $serverName na porta $port."
}
else {
    $errorMessage = $connectionTest.Exception.Message
    Write-Output "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | Erro ao testar a conexão para $serverName na porta $port : $errorMessage"
    # Registrar o erro no log
    Add-Content -Path $logPath -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')-----------------------------------Erro ao testar a conexão para $serverName na porta $port : $errorMessage"
}
}


# Executar o teste de conexão e login a cada 1 segundos
While ($true) {
    TestarConexaoELogin
    Start-Sleep -Seconds 2 # aguardar 1 segundos

}