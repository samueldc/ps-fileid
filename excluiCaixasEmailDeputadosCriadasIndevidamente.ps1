<#
Programa: excluiCaixasEmailDeputadosCriadasIndevidamente
Objetivo: Exclui caixas de e-mail de deputados criadas indevidamente
Autor: Samuel Diniz Casimiro - P_7029
Setor: Ditec/Coaus/Satus
Versões: 1.0 - Criação do script
#>

Set-ExecutionPolicy RemoteSigned
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://urca1.redecamara.camara.gov.br/PowerShell/ -Authentication Kerberos
Import-PSSession $Session -DisableNameChecking
Import-Module ActiveDirectory

$Total = 0
$TotalDepExcluido = 0
$TotalGabExcluido = 0

$ArqLista = $null
$ArqLista = Import-Csv -Path "C:\Users\p_7029\ownCloud\_trabalho\coaus-satus\Posse 2019\excluiCaixasEmailDeputados.csv" -Delimiter ";" -Encoding Default
$ArqLista | ForEach-Object {
    $iddep = $PSItem.iddep.Trim()
    $logindep = "dep." + $iddep
    $grupodep = $logindep + "-u"
    $logingab = "gab." + $iddep
    $grupogab = $logingab + "-u"
    $ponto56 = $PSItem.ponto56.Trim()
    $ponto55 = ""
    
    # Remove caixa dep
    $caixadep = $null
    $caixadep = Get-Mailbox -Identity "$logindep@camara.leg.br" -OrganizationalUnit "OU=Legislatura56,OU=CaixasPostaisInstitucionais,OU=Correio,OU=Usuarios,DC=redecamara,DC=camara,DC=gov,DC=br"
    Remove-ADUser -Identity $caixadep.SamAccountName -Confirm:$false
    If ($?) {
        $TotalDepExcluido++
    }

    # Remove caixa gab
    $caixagab = $null
    $caixagab = Get-Mailbox -Identity "$logingab@camara.leg.br" -OrganizationalUnit "OU=Legislatura56,OU=CaixasPostaisInstitucionais,OU=Correio,OU=Usuarios,DC=redecamara,DC=camara,DC=gov,DC=br"
    Remove-ADUser -Identity $caixagab.SamAccountName -Confirm:$false
    If ($?) {
        $TotalGabExcluido++
    }
    $Total++
}

"Total: $Total"
"Total Dep Excluido: $TotalDepExcluido"
"Total Gab Excluido: $TotalGabExcluido"

Remove-PSSession $Session
