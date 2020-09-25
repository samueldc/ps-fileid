Set-ExecutionPolicy RemoteSigned
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://urca2.redecamara.camara.gov.br/PowerShell/ -Authentication Kerberos
Import-PSSession $Session -DisableNameChecking
Get-ADUser -Filter " Name -like 'SDR_56001' " | 
    ForEach-Object { 
        $mailbox = Get-Mailbox $_.SamAccountName
            if ($?) {
                Set-Mailbox -Identity $mailbox.SamAccountName -DisplayName "SDR $($mailbox.DisplayName) [$($mailbox.PrimarySmtpAddress)]"
                $mailbox = $null
            }
    }