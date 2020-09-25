#Set-ExecutionPolicy RemoteSigned
#$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://urca2.redecamara.camara.gov.br/PowerShell/ -Authentication Kerberos
#Import-PSSession $Session -DisableNameChecking
Get-Mailbox -Filter " Alias -like 'sdr.lid*' " | 
    ForEach-Object {
        Set-Mailbox -Identity $_.SamAccountName -DisplayName "[SDR] $($_.DisplayName)"
    }
