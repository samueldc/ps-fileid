#Requires -Version 5.0
"Running PowerShell $($PSVersionTable.PSVersion)."

<#
Programa: listaCaixasDesconectadas
Objetivo: Lista caixas postais desconectadas de um usuário do AD
Autor: Samuel Diniz Casimiro - P_7029
Setor: Ditec/Coaus/Satus
Versões: 1.0 - Criação do script
#>

Set-ExecutionPolicy RemoteSigned
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://urca1.redecamara.camara.gov.br/PowerShell/ -Authentication Kerberos
Import-PSSession $Session -DisableNameChecking
Get-MailboxDatabase | foreach {Get-MailboxStatistics -Database $_.Name} | where {$_.DisconnectReason.Value -eq "Disabled"} | ft displayname,database,disconnectreason,DisconnectDate -auto
Remove-PSSession $Session