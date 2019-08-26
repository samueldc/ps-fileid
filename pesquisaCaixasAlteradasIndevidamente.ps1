<#
Programa: pesquisaCaixasAlteradasIndevidamente.ps1
Objetivo: Pesquisa caixas alteradas a partir de alguns parâmetros
Autor: Samuel Diniz Casimiro - P_7029
Setor: Ditec/Coaus/Satus
Versões: 1.0 - Criação do script
#>

Set-ExecutionPolicy RemoteSigned
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://urca1.redecamara.camara.gov.br/PowerShell/ -Authentication Kerberos
Import-PSSession $Session -DisableNameChecking
Get-Mailbox -OrganizationalUnit "OU=Legislatura55,OU=CaixasPostaisInstitucionais,OU=Correio,OU=Usuarios,DC=redecamara,DC=camara,DC=gov,DC=br" -ResultSize 2000 | 
    Where-Object { ( $_.WhenChanged -gt "2019-01-24" ) -and ( $_.PrimarySmtpAddress -like "XX*" ) } | 
    ft Name, PrimarySmtpAddress, WhenChanged
Remove-PSSession $Session