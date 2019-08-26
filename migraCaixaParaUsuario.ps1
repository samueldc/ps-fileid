<#
Programa: migraCaixaParaUsuario.ps1
Objetivo: Conecta uma caixa desconectada a um ponto
          É necessário antes desconectar a caixa manualmente
          utilizando a interface web do Exchange
Autor: Samuel Diniz Casimiro - P_7029
Setor: Ditec/Coaus/Satus
Versões: 1.0 - Criação do script
#>

#Import-Module ActiveDirectory
Set-ExecutionPolicy RemoteSigned
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://urca1.redecamara.camara.gov.br/PowerShell/ -Authentication Kerberos
Import-PSSession $Session -DisableNameChecking
"======================================================================================="
$displayName = Read-Host -Prompt "Informe o nome completo (nome de exibição) exatamente como estava no ponto antigo"
#$antiga = (Get-MailboxDatabase | foreach {Get-MailboxStatistics –Database $_.Name} | Where { $_.DisplayName -eq $displayName } | ft DisplayName,Database,DisconnectReason,DisconnectDate)
$antiga = (Get-MailboxDatabase | foreach {Get-MailboxStatistics –Database $_.Name} | Where { $_.DisplayName -eq $displayName } | Select-Object -First 1)
if ($?) {
    $confirmacao = Read-Host -Prompt ("Encontrei a caixa desconectada de " + $antiga.DisplayName + " no banco " +  $antiga.DatabaseName + "; confirma? (S/N)")
    If ($confirmacao -eq "S") {
        $confirmacao = $null
        $ponto = Read-Host -Prompt "Informe o novo ponto no formato P_9999"
        Get-Mailbox -Identity $ponto
        if ($?) {
            "O novo ponto informado já possui uma caixa, desabilite a caixa antes ou informe outro ponto."
        } else {
            "Conectando caixa ao novo ponto..."
            Connect-Mailbox -Identity $antiga.DisplayName -Database $antiga.DatabaseName -User $ponto -WhatIf
            If ($?) {
                $confirmacao = Read-Host -Prompt "A verificação deu certo; confirma a operação? (S/N)"
                If ($confirmacao -eq "S") {
                    Connect-Mailbox -Identity $antiga.DisplayName -Database $antiga.DatabaseName -User $ponto
                        If ($?) {
                            "Caixa conectada com sucesso ao novo ponto."
                        } else {
                            "Ocorreu um erro ao tentar reconectar a caixa ao novo ponto."
                        }
                } else {
                    "Operação cancelada"
                }
            }
        }
    }
} else {
    "Não foi encontrada nenhuma caixa desconectada em nome de $displayName; tente novamente."
}
"Encerrando script..."
Remove-PSSession $Session
"Script encerrado"