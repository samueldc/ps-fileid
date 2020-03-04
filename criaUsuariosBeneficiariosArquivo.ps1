<#
Programa: criaUsuariosBeneficiariosArquivo
Objetivo: Facilitar a criação de usuários beneficiários no AD com base em arquivo
Autor: Samuel Diniz Casimiro - P_7029
Setor: Ditec/Coaus/Sepac
#>

# ------- Definição de variáveis -------

# Se true, ativa o mode de teste (dry-run) nos comandos que utilizam este parâmetro
$WhatIf = $true
$Confirm = $false

$PathLog = "c:\Temp\LogCriaUsuarioBeneficiarioArquivo-" + (Get-Date -UFormat %Y-%m-%d-%H-%M) + ".log"
$PathLista = "C:\Temp\ListaBeneficios.csv"
$OU = 'OU=Deputados,OU=Beneficios,OU=Usuarios,DC=redecamara,DC=camara,DC=gov,DC=br'
$UsrCriados = 0
$UsrNaoCriados = 0
$UsrGrupos = 0
$UsrNaoGrupos = 0
$UsrTotal = 0

# ------- Início do Script -------

"Início do script: " + (Get-Date) *> $PathLog
$ArqLista = import-csv -Path $PathLista -Delimiter ";" -Encoding Default

$ArqLista | ForEach-Object {

    $UsrTotal ++

    # Cria o usuário no AD
    New-ADUser -Name $PSItem.ponto -DisplayName $PSItem.nome -GivenName $PSItem.primeironome -Surname $PSItem.sobrenome -UserPrincipalName ($PSItem.ponto + "@redecamara.camara.gov.br") -Path $OU -ChangePasswordAtLogon $false -Enabled $true -AccountPassword (ConvertTo-SecureString -AsPlainText '0123456789' -Force) -verbose -Confirm:$Confirm -WhatIf:$WhatIf
        if ($?) { 
            $UsrCriados++
            "O usuário " + $PSItem.ponto + " foi criado. Registro nº $UsrTotal da Lista." *>> $PathLog
            "-------------------------------------" *>> $PathLog
        } else { 
            $UsrNaoCriados++
            "O usuário " + $PSItem.ponto + " não foi criado. Registro nº $UsrTotal da Lista." *>> $PathLog
            "-------------------------------------" *>> $PathLog
        }    
        
    # Adiciona o usuário nos grupos padrão
    Add-ADGroupMember -Identity Internet -Members $PSItem.ponto -verbose -Confirm:$Confirm -WhatIf:$WhatIf
        if ($?) { 
            $UsrGrupos++
            "O usuário " + $PSItem.ponto + " foi incluido nos grupos padrão. Registro nº $UsrTotal da Lista." *>> $PathLog
            "-------------------------------------" *>> $PathLog
        } else { 
            $UsrNaoGrupos++
            "O usuário " + $PSItem.ponto + " não incluido nos grupos padrão. Registro nº $UsrTotal da Lista." *>> $PathLog
            "-------------------------------------" *>> $PathLog
        }    

}

# Relatório da execução do script
"Relatório de usuários:" *>> $PathLog
"Total de Usuários na lista: $UsrTotal" *>> $PathLog
"Usuários criados: $UsrCriados" *>> $PathLog
"Usuários NÃO criados: $UsrNaoCriados" *>> $PathLog
"Usuários incluídos nos grupos padrão: $UsrGrupos" *>> $PathLog
"Usuários NÃO incluídos nos grupos padrão: $UsrNaoGrupos" *>> $PathLog
"Fim do script:" + (Get-Date) *>> $PathLog

# ------- Fim do Script -------