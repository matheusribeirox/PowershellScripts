# Script para realizar envio de emails com a autenticação moderna (OAuth 2.0) com uma conta GMail. 

Install-Module -Name Mailozaurr -AllowClobber -Force    # Importar modulo necessário com os métodos para o envio de emails com autenticação moderna

# Update-Module -Name Mailozaurr

$ClientID = '57492806606-inq623pbvp1flg5oofpefhsgqen5kjq7.apps.googleusercontent.com'    # ID do Cliente
$ClientSecret = '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'      # Secret ID

$CredentialOAuth2 = Connect-oAuthGoogle -ClientID $ClientID -ClientSecret $ClientSecret -GmailAccount 'matheus.oqw@gmail.com'   # Conta do Gmail

Send-EmailMessage -From @{ Name = 'Matheus Silva Ribeiro'; Email = 'matheus.oqw@gmail.com' } -To 'matheus.ribeiro@benner.com.br' `
    -Server 'smtp.gmail.com' -HTML $Body -Text $Text -DeliveryNotificationOption OnSuccess -Priority High `
    -Subject 'This is another test email' -SecureSocketOptions Auto -Credential $CredentialOAuth2 -oAuth