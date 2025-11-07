# Configurações de e-mail para Outlook.com ou Microsoft 365 sem autenticacao moderna
$smtpServer = "smtp.office365.com"
$smtpPort = 587  # Pode ser 25 também
$smtpFrom = "matheus.oqw@hotmail.com"
$smtpTo = "matheus.ribeiro@benner.com.br"
$messageSubject = "Assunto do E-mail"
$messageBody = "TestEnvio de Emails do Office365"

# Configurações de credenciais
$smtpUsername = "matheus.oqw@hotmail.com"
$smtpPassword = "Skyline@157"

# Configuração do cliente SMTP
$smtp = New-Object Net.Mail.SmtpClient($smtpServer, $smtpPort)
$smtp.EnableSsl = $true  # Habilita STARTTLS ou TLS
$smtp.Credentials = New-Object System.Net.NetworkCredential($smtpUsername, $smtpPassword)

# Cria a mensagem
$message = New-Object system.net.mail.mailmessage
$message.from = ($smtpFrom)
$message.To.add($smtpTo)
$message.Subject = $messageSubject
$message.Body = $messageBody

# Tenta enviar o e-mail
try {
    $smtp.Send($message)
    Write-Host "E-mail enviado com sucesso!" -ForegroundColor Green
}
catch {
    Write-Host "Erro ao enviar o e-mail: $_"
}
