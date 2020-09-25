<#
Programa: criaUsuariosCaixasEmailDeputadosSDR
Objetivo: Facilitar a criação de usuários e caixas postais de Deputados para uso no SDR
Autor: Samuel Diniz Casimiro - P_7029
Setor: Ditec/Coaus/Satus
Versões: 1.0 - Criação do script
#>

# Variáveis globais -----------------------------------------------------------------------

# Se true, ativa o mode de teste (dry-run) nos comandos que utilizam este parâmetro
$WhatIf = $true
$Confirm = $false

# Caminho e nome dos arquivos de log
$PathLog = "C:\Users\p_7029\ownCloud\_trabalho\teletrabalho\sdr\criaUsuariosCaixasEmailDeputadosSDR-" + (Get-Date -UFormat %Y-%m-%d-%H-%M) + ".log"
$PathLogErro = "C:\Users\p_7029\ownCloud\_trabalho\teletrabalho\sdr\criaUsuariosCaixasEmailDeputadosSDR-" + (Get-Date -UFormat %Y-%m-%d-%H-%M) + ".logerro"

# Caminho e nome do arquivo CSV com a lista de usuarios a ser importada e exportada
$PathLista = "C:\Users\p_7029\ownCloud\_trabalho\teletrabalho\sdr\usuariosSDR2.csv"
$PathListaFinal = "C:\Users\p_7029\ownCloud\_trabalho\teletrabalho\sdr\usuariosSDR2Final.csv"

# DistinguishedName das "OU"s envolvidas
$OU = "OU=56Legislatura,OU=Deputados,OU=Usuarios,DC=redecamara,DC=camara,DC=gov,DC=br"
$OUsdr = "OU=56LegislaturaSDR,OU=Deputados,OU=Usuarios,DC=redecamara,DC=camara,DC=gov,DC=br"

<#
Os contadores a seguir servem para gerar um resumo das ações ao final da execução do script.
Esse resumo é incluído no final do arquivo de log gerado.
#>

# Contadores de usuários
$UsuariosCriados = 0
$UsuariosNaoCriados = 0
$UsuariosNaoEncontrados = 0
$UsuariosTotal = 0

# Aqui começam as ações do script

# Solicita credencial para executar o script (não precisa pq por enquanto estamos utilizando as credenciais do usuário logado na máquina)
##$Credential = Get-Credential

# Cria sessão com o Exchange; a autenticação é feita com as credenciais do usuário logado na máquina
Set-ExecutionPolicy RemoteSigned
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://urca2.redecamara.camara.gov.br/PowerShell/ -Authentication Kerberos
Import-PSSession $Session -DisableNameChecking

# Cria arquivo de log e já registra a data e a hora do inicio da execução do script.
"Início do script: " + (Get-Date) *> $PathLog

# Cria arquivos
"" *> $PathLogErro
"ponto_sdr;nome_sdr;email_sdr;senha_sdr" *> $PathListaFinal

# Registra no log a credencial usada para execução do script
"O script foi executado por: $env:COMPUTERNAME\$env:UserName" *>> $PathLog

# Importar arquivo de CSV
$ArqLista = $null
$ArqLista = Import-Csv -Path $PathLista -Delimiter ";" -Encoding Default
# Verifica se a importacao foi bem sucedida
if (!$?) { # Se não foi...
    # Loga o erro
    "[Import-Csv] Ocorreu um erro ao abrir o arquivo: " + $PathLista *>> $PahLogErro    
    # Finaliza a sessão com o Exchange
    Remove-PSSession $Session
    "Fim do script:" + (Get-Date) *>> $PathLog
    # Encerra a execução do script
    Pause
    Exit
}

# Se chegou até aqui é pq a importação do arquivo csv foi bem sucedida

# Passa os registros da lista para estrutura de repetição
$ArqLista | ForEach-Object {

    # Povoa as variáveis
    
    # Variáveis gerais

    # Redefine para não correr o risco de viciar os valores (trata-se de uma característica do PowerShell que em caso de erro em comando o valor anterior da variável é mantido)
    $ponto = ""
    $pontosdr = ""
    $senhasdr = ""

    # Povoa
    $ponto = $PSItem.ponto.Trim()
    $pontosdr = $PSItem.ponto_sdr.Trim()
    $senhasdr = $PSItem.senha_sdr.Trim()

    $usuario = $null
    
    $pontosdr
    # Verifica se o usuario do deputado existe
    $usuario = Get-ADUser -Identity $ponto -Properties *

    # Se o usuário do deputado existe, e tem email, cria a caixa e o usuario sdr com base nos dados do usuario e email do deputado
    if ($? -and $usuario -and $usuario.mail) {

        # converte o nome da caixa dep para sdr
        if ($usuario.mail -match ".*(dep|gab)\.(.*)") {
            $emaildep1 = $usuario.mail -replace ".*(dep|gab)\.(.*)",'sdr.$2'
        } else {
            $emaildep1 = "sdr.$($usuario.mail)"
        }
        $emaildep2 = $emaildep1 -replace "(.*)\.leg\.br",'$1.gov.br'

        # cria a nova caixa sdr que já cria automaticamente o usuario sdr
        $mailbox = New-Mailbox -Name $pontosdr -SamAccountName $pontosdr -DisplayName $usuario.DisplayName -UserPrincipalName ("$pontosdr" + "@redecamara.camara.gov.br") -PrimarySmtpAddress $emaildep1 -Alias $pontosdr -FirstName $usuario.GivenName -LastName $usuario.sn -OrganizationalUnit $OUsdr -Password (ConvertTo-SecureString -AsPlainText $senhasdr -Force) -ResetPasswordOnNextLogon $false -Confirm:$Confirm -WhatIf:$WhatIf
        $mailbox.DistinguishedName
        
        # verifica se a conta foi criada
        if ($?) {

            "[New-Mailbox] Caixa " + $pontosdr + " criada." *>> $PathLog
            Set-Mailbox -Identity $mailbox.DistinguishedName -EmailAddresses @{Add="smtp:$emaildep2"} -EmailAddressPolicyEnabled $false -Confirm:$Confirm -WhatIf:$WhatIf
                if($?){"[Set-Mailbox] Caixa $pontosdr com endereços alternativos incluidos." *>> $PathLog}
                    else{"[Set-Mailbox] Ocorreu um erro ao incluir endereços alternativos na caixa $pontosdr." *>> $PahLogErro}
            Set-ADUser -Identity $mailbox.DistinguishedName -EmailAddress $emaildep1 -Description $usuario.Description -Title $usuario.Title -Department $usuario.Department -ChangePasswordAtLogon $false -Enabled $true -Confirm:$Confirm -WhatIf:$WhatIf
                if($?){"[Set-ADUser] Caixa $pontosdr com titulo e descricao incluidos." *>> $PathLog}
                    else{"[Set-ADUser] Ocorreu um erro ao incluir titulo e descricao na caixa $pontosdr." *>> $PahLogErro}

            # Adiciona a caixa nos grupos padrão
            Add-ADGroupMember -Identity Deputados -Members $mailbox.DistinguishedName -Confirm:$Confirm -WhatIf:$WhatIf
            Add-ADGroupMember -Identity DeputadosSDR -Members $mailbox.DistinguishedName -Confirm:$Confirm -WhatIf:$WhatIf
            Add-ADGroupMember -Identity ExchangePerfil2 -Members $mailbox.DistinguishedName -Confirm:$Confirm -WhatIf:$WhatIf

            # Adiciona o usuário nos grupos padrão
            #Add-ADGroupMember -Identity InfolegParlamentarUsuarioDeputado -Members $mailbox.DistinguishedName -Confirm:$Confirm -WhatIf:$WhatIf
            #Add-ADGroupMember -Identity InfolegPlenarioVirtual -Members $mailbox.DistinguishedName -Confirm:$Confirm -WhatIf:$WhatIf
            Add-ADGroupMember -Identity InfolegAppSDR -Members $mailbox.DistinguishedName -Confirm:$Confirm -WhatIf:$WhatIf
            Add-ADGroupMember -Identity AppParlamentar -Members $mailbox.DistinguishedName -Confirm:$Confirm -WhatIf:$WhatIf
            Add-ADGroupMember -Identity deputadologon -Members $mailbox.DistinguishedName -Confirm:$Confirm -WhatIf:$WhatIf
            Add-ADGroupMember -Identity Internet -Members $mailbox.DistinguishedName -Confirm:$Confirm -WhatIf:$WhatIf
            Add-ADGroupMember -Identity LeginetConsulta -Members $mailbox.DistinguishedName -Confirm:$Confirm -WhatIf:$WhatIf
            Add-ADGroupMember -Identity SilegAutenticadorGab -Members $mailbox.DistinguishedName -Confirm:$Confirm -WhatIf:$WhatIf
            Add-ADGroupMember -Identity UsuariosCorreioEletronicoSeguro -Members $mailbox.DistinguishedName -Confirm:$Confirm -WhatIf:$WhatIf
                if($?){"[Add-ADGroupMember] Caixa $pontosdr com permissoes incluidas." *>> $PathLog}
                    else{"[Add-ADGroupMember] Ocorreu um erro ao incluir permissoes na caixa $pontosdr." *>> $PahLogErro}

            #New-ADUser -Name $pontosdr -DisplayName $usuario.DisplayName -GivenName $usuario.GivenName -Surname $usuario.sn -Description $usuario.Description -Title $usuario.Title -Department $usuario.Department -UserPrincipalName ($pontosdr + "@redecamara.camara.gov.br") -Path $OUsdr -ChangePasswordAtLogon $false -Enabled $true -AccountPassword (ConvertTo-SecureString -AsPlainText $senhasdr -Force) -Confirm:$Confirm -WhatIf:$WhatIf
                #if($?){"[New-ADUser] Usuario " + $pontosdr + " criado." *>> $PathLog}
                    #else{"[New-ADUser] Ocorreu um erro ao criar o usuario " + $pontosdr + "." *>> $PahLogErro}

            "$pontosdr;$($usuario.DisplayName);$emaildep1;$senhasdr" *>> $PathListaFinal
            $UsuariosCriados++

        } else {

            "[New-Mailbox] Ocorreu um erro ao criar o usuário e caixa $pontosdr." *>> $PahLogErro
            $UsuariosNaoCriados++

        }

    } else {

        "[New-Mailbox] Usuário não encontrado $pontosdr." *>> $PahLogErro
        $UsuariosNaoEncontrados++

    }

    $UsuariosTotal++

}

# Registra no Log um relatório da execução do Script
"Relatório de usuários:" *>> $PathLog
"Total de usuários na lista: $UsuariosTotal" *>> $PathLog
"Usuários criados: $UsuariosCriados" *>> $PathLog
"Usuários com erro na criação: $UsuariosNaoCriados" *>> $PathLog
"Usuários não encontrados: $UsuariosNaoEncontrados" *>> $PathLog

if ($UsuariosTotal -ne ($UsuariosCriados + $UsuariosNaoEncontrados)) {
    "ATENÇÃO: Existe uma diferença entre a qtde. de usuários total e a qtde. de usuários criados mais os não encontrados." *>> $PathLog
    ">>>>>>>> Sugere-se revisar os dados do arquivo csv." *>> $PathLog
}

# Finaliza sessão

Remove-PSSession $Session

# Registra no Log a data e a hora do FIM da execução do script.
"Fim do script: $(Get-Date)" *>> $PathLog

# Fim do Script ------------------------------------------------------------------