$connectionString = "persist security info=true;packet size=8000;data source=10.175.32.202, 1433;initial catalog=RH_BENNER_PROD;user id=sa_acessosm;password=#smapl2014@;connect timeout=30;trusted_connection=no;Max Pool Size=200;"
$logFile = "C:\Benner\logBD20241210.txt"
$resultFile = "C:\Benner\resultBD20241210.txt"

while ($true) {
    try {
        $connection = New-Object System.Data.SqlClient.SqlConnection($connectionString)
        $connection.Open()

        if ($connection.State -eq 'Open') {
            Write-Host "Conexão bem-sucedida com o banco de dados."
            Add-Content -Path $logFile -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Conexão bem-sucedida com o banco de dados."

            # Teste de seleção simples na tabela z_sistemas
            $query = "SELECT handle, licclienteid, licclientenome, versaodosistema, bserversistema, bserverhost FROM z_sistema"
            $command = $connection.CreateCommand()
            $command.CommandText = $query

            # Início do cronômetro
            $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

            $result = $command.ExecuteReader()

            # Parar o cronômetro e calcular o tempo
            $stopwatch.Stop()
            $elapsedTime = $stopwatch.Elapsed.TotalMilliseconds

            # Log do tempo de execução
            Write-Host "Tempo de execução da query: $elapsedTime ms"
            Add-Content -Path $resultFile -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Tempo de execução da query: $elapsedTime ms"

            # Processar os resultados e gravar no arquivo de resultado
            $resultText = @()
            while ($result.Read()) {
                $resultRow = @()
                for ($i = 0; $i -lt $result.FieldCount; $i++) {
                    $resultRow += $result[$i]
                }
                $resultText += ($resultRow -join ",")
            }
            $result.Close()
            $resultText | Out-File -FilePath $resultFile -Append
        }
        else {
            Write-Host "Não foi possível abrir a conexão com o banco de dados."
            Add-Content -Path $logFile -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Não foi possível abrir a conexão com o banco de dados."
        }
    }
    catch {
        Write-Host "Ocorreu um erro ao tentar conectar ao banco de dados: $_"
        Add-Content -Path $logFile -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Ocorreu um erro ao tentar conectar ao banco de dados: $_"
    }
    finally {
        if ($connection.State -eq 'Open') {
            $connection.Close()
        }
    }

    Start-Sleep -Seconds 2
}
