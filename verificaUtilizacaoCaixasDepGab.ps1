Set-ExecutionPolicy RemoteSigned
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://urca1.redecamara.camara.gov.br/PowerShell/ -Authentication Kerberos
Import-PSSession $Session -DisableNameChecking

$OUcaixa = "OU=Legislatura56,OU=CaixasPostaisInstitucionais,OU=Correio,OU=Usuarios,DC=redecamara,DC=camara,DC=gov,DC=br"

Get-Mailbox -OrganizationalUnit $OUcaixa -Filter ' Name -like "dep.*" ' | Get-MailboxStatistics | Select-Object DisplayName, TotalItemSize | Export-Csv C:\Temp\MailboxSizesDep.CSV –NoTypeInformation –Encoding UTF8
#Get-Mailbox -OrganizationalUnit $OUcaixa -Filter ' Name -like "gab.*" ' | Get-MailboxStatistics | Format-Table DisplayName, @{Label="TotalItemSize";Expression={“{0:N2}” -f $_.TotalItemSize.Value.ToKB()}}
#Get-Mailbox -OrganizationalUnit $OUcaixa -Filter ' Name -like "gab.*" ' | Select-Object Alias, ProhibitSendQuota, @{label="TotalItemSize(MB)";expression={(get-mailboxstatistics $_).TotalItemSize.Value.ToMB()}}, @{label="ItemCount";expression={(get-mailboxstatistics $_).ItemCount}}, Database | Format-Table

Remove-PSSession $Session