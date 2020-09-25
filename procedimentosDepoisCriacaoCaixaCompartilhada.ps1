Set-ExecutionPolicy RemoteSigned
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://urca8.redecamara.camara.gov.br/PowerShell/ -Authentication Kerberos
Import-PSSession $session

#P/ permissão de enviar como e receber como (Caixas Institucional)
  #Add-MailboxPermission -Identity sdr.lid.cidadania -User sdr.lid.cidadania-u -AccessRights:FullAccess -AutoMapping $true
  #Add-ADPermission -Identity "sdr.lid.cidadania" -User sdr.lid.cidadania-u -ExtendedRights "Send As"

#P/ mover caixas institucionais de "CAIXAS DE CORREIO"  para "COMPARTILHADOS"
    #Get-Mailbox agendadopresidente | Set-Mailbox -Type shared

#Remove-PSSession $Session