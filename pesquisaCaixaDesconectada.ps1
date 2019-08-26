<#
Programa: pesquisaCaixaDesconectada.ps1
Objetivo: Pesquisa caixa a partir do nome da pessoa e indica
          se está desativada ou não
Autor: Samuel Diniz Casimiro - P_7029
Setor: Ditec/Coaus/Satus
Versões: 1.0 - Criação do script
#>

Set-ExecutionPolicy RemoteSigned
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://urca1.redecamara.camara.gov.br/PowerShell/ -Authentication Kerberos
Import-PSSession $Session -DisableNameChecking
$displayName = Read-Host -Prompt "Informe uma parte do nome para pesquisa"
Get-MailboxDatabase | foreach {Get-MailboxStatistics –Database $_.Name} | Where { $_.DisplayName -like "*$displayName*" } | ft DisplayName,Database,DisconnectReason,DisconnectDate
Remove-PSSession $Session