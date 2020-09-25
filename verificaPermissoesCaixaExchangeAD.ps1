# Verifica permissões nas urcas de 1 a 8
Set-ExecutionPolicy RemoteSigned
for (($i = 1); $i -lt 9; $i++) {
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://urca$i.redecamara.camara.gov.br/PowerShell/" -Authentication Kerberos
    $echo = Import-PSSession $session -Verbose $false
    "urca$i"
    Get-MailboxPermission -Identity sdr.lid.maioria -User sdr.lid.maioria-u
    Remove-PSSession $Session
}
# Verifica permissões nos controladores de domínio
$controladores = "CALIFORNIO", "CARBONO", "CERIO", "CHUMBO", "CLORO", "COBALTO", "COBRE", "CROMO"
$controladores | ForEach-Object {
    $_
    Get-ADPermission -Identity sdr.lid.maioria -User sdr.lid.maioria-u -DomainController "$_.redecamara.camara.gov.br"
}


