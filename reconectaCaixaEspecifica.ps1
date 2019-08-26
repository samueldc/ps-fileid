Set-ExecutionPolicy RemoteSigned
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://urca1.redecamara.camara.gov.br/PowerShell/ -Authentication Kerberos
Import-PSSession $Session -DisableNameChecking
Connect-Mailbox -Identity "Dacila Araci Schmitt" -Database USR012 -User P_122005 -WhatIf
Remove-PSSession $Session