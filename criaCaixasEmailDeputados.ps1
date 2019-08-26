<#
Programa: criaCaixasEmailDeputados
Objetivo: Facilitar a criação de caixas postais de Deputados recém empossados
Autor: Samuel Diniz Casimiro - P_7029
Setor: Ditec/Coaus/Satus
Versões: 1.0 - Criação do script
#>

# Variáveis globais -----------------------------------------------------------------------

# Se true, ativa o mode de teste (dry-run) nos comandos que utilizam este parâmetro
$WhatIf = $true
$Confirm = $false

# Caminho e nome dos arquivos de log
$PathLog = "C:\Users\p_7029\ownCloud\_trabalho\coaus-satus\Posse 2019\criaCaixasEmailDeputados-" + (Get-Date -UFormat %Y-%m-%d-%H-%M) + ".log"
$PathLogErro = "C:\Users\p_7029\ownCloud\_trabalho\coaus-satus\Posse 2019\criaCaixasEmailDeputados-" + (Get-Date -UFormat %Y-%m-%d-%H-%M) + ".logerro"

# Caminho e nome do arquivo CSV com a lista de usuarios a ser importada
$PathLista = "C:\Users\p_7029\ownCloud\_trabalho\coaus-satus\Posse 2019\criaCaixasEmailDeputados.csv"

# DistinguishedName da "OU" onde os usuários serão criados
$OU = "OU=56Legislatura,OU=Deputados,OU=Usuarios,DC=redecamara,DC=camara,DC=gov,DC=br"
$OUcaixa = "OU=Legislatura56,OU=CaixasPostaisInstitucionais,OU=Correio,OU=Usuarios,DC=redecamara,DC=camara,DC=gov,DC=br"

<#
Os contadores a seguir servem para gerar um resumo das ações ao final da execução do script.
Esse resumo é incluído no final do arquivo de log gerado.
#>

# Contadores de usuários
$UsuariosCriados = 0
$UsuariosRenomeados = 0
$UsuariosNaoCriados = 0
$UsuariosTotal = 0

# Contador de caixas dep
$CaixasDepCriadas = 0
$CaixasDepNaoCriadas = 0
$CaixasDepTotal = 0
$CaixasDepDesativadas = 0

# Contador de caixas gab
$CaixasGabCriadas = 0
$CaixasGabNaoCriadas = 0
$CaixasGabTotal = 0
$CaixasGabDesativadas = 0

# Aqui começam as ações do script

# Solicita credencial para executar o script (não precisa pq por enquanto estamos utilizando as credenciais do usuário logado na máquina)
##$Credential = Get-Credential

# Cria sessão com o Exchange; a autenticação é feita com as credenciais do usuário logado na máquina
Set-ExecutionPolicy RemoteSigned
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://urca1.redecamara.camara.gov.br/PowerShell/ -Authentication Kerberos
Import-PSSession $Session -DisableNameChecking

# Cria arquivo de log e já registra a data e a hora do inicio da execução do script.
"Início do script: " + (Get-Date) *> $PathLog

# Cria arquivo de log de erros
"" *> $PathLogErro

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
    $iddep = ""
    $ponto56 = ""
    $ponto55 = ""
    $nome = ""
    $primeironome = ""
    $sobrenome = ""
    $senha = ""

    # Povoa
    $iddep = $PSItem.iddep.Trim()
    $ponto56 = $PSItem.ponto56.Trim()
    # Verifica se tem ponto 55
    if ($PSItem.ponto55) { # Se tem...
        $ponto55 = $PSItem.ponto55.Trim()
    } # Se não tem, o ponto ficam em branco
    $descricao = "Disabled Windows user account"
    $nome = $PSItem.nome.Trim()
    $primeironome = $PSItem.primeironome.Trim()
    # Verifica se tem sobrenome
    if ($PSItem.segundonome) { # Se tem...
        $sobrenome = $PSItem.segundonome.Trim()
    } # Se não tem, o sobrenome fica em branco
    $senha = $PSItem.senha.Trim()

    # Variáveis dep
    $logindep = "dep." + $iddep
    $samaccountnamedep = $logindep
    # O atributo SamAccountName do AD só admite 20 caracteres; então é necessário truncar quando for maior
    if ($logindep.Length -gt 20) {$samaccountnamedep = $logindep.Substring(0, 20)}
    $logindepdn = "cn=" + $logindep + ",OU=Legislatura56,OU=CaixasPostaisInstitucionais,OU=Correio,OU=Usuarios,DC=redecamara,DC=camara,DC=gov,DC=br"
    $nomedep = "Dep. " + $nome
    $emaildep1 = $logindep + "@camara.leg.br"
    $emaildep2 = $logindep + "@camara.gov.br"
    $emaildep3 = $ponto56 + "@camara.leg.br"
    $emaildep4 = $ponto56 + "@camara.gov.br"
    $grupodep = $logindep + "-u"
    $grupodepdn = "cn=" + $grupodep + ",ou=IDEA,ou=Grupos,ou=Usuarios,dc=redecamara,dc=camara,dc=gov,dc=br"
    $titulo = "Deputado"

    # Variáveis gab
    $logingab = "gab." + $iddep
    $samaccountnamegab = $logingab
    if ($logingab.Length -gt 20) {$samaccountnamegab = $logingab.Substring(0, 20)}
    $logingabdn = "cn=" + $logingab + ",OU=Legislatura56,OU=CaixasPostaisInstitucionais,OU=Correio,OU=Usuarios,DC=redecamara,DC=camara,DC=gov,DC=br"
    $nomegab = "Gab. " + $nome
    $emailgab1 = $logingab + "@camara.leg.br"
    $emailgab2 = $logingab + "@camara.gov.br"
    $grupogab = $logingab + "-u"
    $grupogabdn = "cn=" + $grupogab + "-u,ou=IDEA,ou=Grupos,ou=Usuarios,dc=redecamara,dc=camara,dc=gov,dc=br"
    $titulogab = "Institucional"
    $empresagab = "Gabinetes Parlamentares"

    # Verifica se o novo ponto já existe
    $usuario = $null
    # Tenta obter o usuario novo
    $usuario = Get-ADUser -Identity $ponto56 -Properties *
    # Verifica se o comando acima foi bem sucedido e se o objeto $usuario foi retornado pelo comando
    if ($? -and $usuario) { # Se o usuário foi encontrado...
        # Usuario novo encontrado
        "Usuario " + $ponto56 + " já existente e não foi criado." *>> $PathLog
        # Atualiza os dados para ficar padronizado
        Set-ADUser -Identity $ponto56 -DisplayName $nomedep -GivenName $primeironome -Surname $sobrenome -Description $nomegab -Title $titulo -Department $nomegab -UserPrincipalName ($ponto56 + "@redecamara.camara.gov.br") -ChangePasswordAtLogon $true -Enabled $true -Confirm:$Confirm -WhatIf:$WhatIf
            if($?){"[Set-ADUser] Usuario " + $ponto56 + " atualizado." *>> $PathLog}
                else{"[Set-ADUser] Ocorreu um erro ao atualizar o usuario " + $ponto56 + "." *>> $PahLogErro}
        $UsuariosNaoCriados++
    }
    else { # Se o novo ponto ainda não existe...
        # Verifica se o antigo ponto já existe
        if ($ponto55) { # Se o antigo ponto existe...
            # Obtem o usuário antigo
            $usuario = Get-ADUser -Identity $ponto55 -Properties *
        }
        # Verifica se o usuário antigo foi encontrado
        if ($? -and $usuario) { # Se o usuário antigo (reeleito) já existe, então é necessário apenas renomeá-lo e move-lo para a nova "OU"
            # Renomea o usuário antigo
            "Renomeando usuário " + $ponto55 + "..." *>> $PathLog
            Rename-ADObject -Identity ("CN=" + $ponto55 + ",OU=55Legislatura,OU=Deputados,OU=Usuarios,DC=redecamara,DC=camara,DC=gov,DC=br") -NewName $ponto56 -Confirm:$Confirm -WhatIf:$WhatIf
                if($?){"[Rename-ADObject] Usuario " + $ponto55 + " renomeado para " + $ponto56 + "." *>> $PathLog}
                    else{"[Rename-ADObject] Ocorreu um erro ao renomear o usuario " + $ponto55 + " para " + $ponto56 + "." *>> $PahLogErro}
            Set-ADUser -Identity $ponto55 -SamAccountName $ponto56 -UserPrincipalName ($ponto56 + "@redecamara.camara.gov.br") -Confirm:$Confirm -WhatIf:$WhatIf
                if($?){"[Set-ADUser] Usuario " + $ponto55 + " renomeado para " + $ponto56 + "." *>> $PathLog}
                    else{"[Set-ADUser] Ocorreu um erro ao renomear o usuario " + $ponto55 + " para " + $ponto56 + "." *>> $PahLogErro}
            Move-ADObject -Identity ("CN=" + $ponto56 + ",OU=55Legislatura,OU=Deputados,OU=Usuarios,DC=redecamara,DC=camara,DC=gov,DC=br") -TargetPath $OU -Confirm:$Confirm -WhatIf:$WhatIf
                if($?){"[Move-ADObject] Usuario " + $ponto56 + " movido para " + $OU + "." *>> $PathLog}
                    else{"[Move-ADObject] Ocorreu um erro ao mover o usuario " + $ponto56 + " para " + $OU + "." *>> $PahLogErro}
            $UsuariosRenomeados++
        }
        else { # Se usuario antigo também não existe, então é necessário criar um novo
            # Inclui novo usuario
            "Adicionando usuário " + $ponto56 + "..." *>> $PathLog
            New-ADUser -Name $ponto56 -DisplayName $nomedep -GivenName $primeironome -Surname $sobrenome -Description $nomegab -Title $titulo -Department $nomegab -UserPrincipalName ($ponto56 + "@redecamara.camara.gov.br") -Path $OU -ChangePasswordAtLogon $true -Enabled $true -AccountPassword (ConvertTo-SecureString -AsPlainText $senha -Force) -Confirm:$Confirm -WhatIf:$WhatIf
                if($?){"[New-ADUser] Usuario " + $ponto56 + " criado." *>> $PathLog}
                    else{"[New-ADUser] Ocorreu um erro ao criar o usuario " + $ponto56 + "." *>> $PahLogErro}
            # Adiciona o usuário nos grupos padrão
            Add-ADGroupMember -Identity AppParlamentar -Members $ponto56 -Confirm:$Confirm -WhatIf:$WhatIf
            Add-ADGroupMember -Identity deputadologon -Members $ponto56 -Confirm:$Confirm -WhatIf:$WhatIf
            Add-ADGroupMember -Identity Internet -Members $ponto56 -Confirm:$Confirm -WhatIf:$WhatIf
            Add-ADGroupMember -Identity LeginetConsulta -Members $ponto56 -Confirm:$Confirm -WhatIf:$WhatIf
            Add-ADGroupMember -Identity SilegAutenticadorGab -Members $ponto56 -Confirm:$Confirm -WhatIf:$WhatIf
            Add-ADGroupMember -Identity UsuariosCorreioEletronicoSeguro -Members $ponto56 -Confirm:$Confirm -WhatIf:$WhatIf
            Add-ADGroupMember -Identity InfolegParlamentarUsuarioDeputado -Members $ponto56 -Confirm:$Confirm -WhatIf:$WhatIf
                if($?){"[Add-ADGroupMember] Usuario " + $ponto56 + " adicionado aos grupos pre-definidos." *>> $PathLog}
                    else{"[Add-ADGroupMember] Ocorreu um erro ao adicionar o usuario " + $ponto56 + " aos grupos pre-definidos." *>> $PahLogErro}
            $UsuariosCriados++
        }
    }
    $UsuariosTotal++

    # Estando tudo ok com o usuário, passa a tratar as caixas postais

    <#
    Já existindo uma caixa postal de um deputado, ela deve ser reaproveitada.
    Mas é necessário garantir que aquela caixa realmente é da mesma pessoa, pois alguns deputados são homônimos.
    Para garantir isso, a única forma confiável é pesquisar a caixa pelo ponto, ou no endereço ou no alias.
    No caso da caixa dep isso é tranquilo, pois normalmente ela vai conter um ou dois endereços que utilizam o ponto (ex.: D_55001@camara.leg.br).
    O problema é a caixa gab, que não contém referências ao ponto. 
    A solução é só reaproveitar a caixa gab quando a dep for encontrada, e obter a caixa gab a partir de uma composição do nome da caixa dep.
    E mesmo que uma caixa não seja reaproveitada, antes de criar a caixa nova é preciso garantir que não há uma caixa antiga com o mesmo nome.
    Para garantir isso, é feita uma pesquisa pelo nome, ao invés do ponto. Se uma caixa com o mesmo nome for encontrada, ela é "desativada".
    A "desativação" envolve tão somente invalidar o endereço anterior com um "XX" e retirar a caixa dos grupos de lista de distribuição.
    Também é necessário renomear o objeto AD correspondente a caixa desativada para incluir o XX também no nome do objeto (Name, UserPrincipalName e SamAccountName).
    #>

    # Cria caixa dep

    # Verifica antes se a caixa dep já existe
    $caixadep = $null
    $caixadep = Get-Mailbox -Filter { EmailAddresses -like "*$ponto56@*" } # procura pelo endereço no formato '<ponto56>@camara.leg.br', que é mais seguro, pois o nome pode ter sido alterado
    if (!$caixadep -and $ponto55) { # Se a caixa dep nova não existe, procura pelo ponto antigo, se houver
        $caixadep = Get-Mailbox -Filter { EmailAddresses -like "*$ponto55@*" } # busca pelo endereço no formato '<ponto55>@camara.leg.br', que é mais seguro, pois o nome pode ter sido alterado
    }
    if (!$caixadep -and $ponto55) { # Se não encontrar, procura uma última vez pelo alias (que geralmente contém o ponto do parlamentar)
        $caixadep = Get-Mailbox -Identity $ponto55 # busca pelo alias
    }
    <#
    Trecho descontinuado
    if (!$caixadep) { # Se não encontrou pelos pontos, procura pelo SamAccountName
        $caixadep = Get-Mailbox -Identity $samaccountnamedep -OrganizationalUnit "OU=Legislatura55,OU=CaixasPostaisInstitucionais,OU=Correio,OU=Usuarios,DC=redecamara,DC=camara,DC=gov,DC=br"
            if($? -and $caixadep){"ATENÇÃO: Verifique se esta caixa $samaccountnamedep realmente é da conta $ponto55 ." *>> $PathLog}
    }
    if (!$caixadep) { # Se não encontrou pelo SamAccountName, procura pelo email mesmo
        $caixadep = Get-Mailbox -Identity $emaildep1 -OrganizationalUnit "OU=Legislatura55,OU=CaixasPostaisInstitucionais,OU=Correio,OU=Usuarios,DC=redecamara,DC=camara,DC=gov,DC=br"
            if($? -and $caixadep){"ATENÇÃO: Verifique se esta caixa $emaildep1 realmente é da conta $ponto55 ." *>> $PathLog}
    }
    Fim do trecho descontinuado
    #>
    if (!$caixadep) { # Se a caixa dep ainda não existe...
        # Se assegura que qualquer caixa antiga com o mesmo endereço seja "desativada" 
        # A desativação, na verdade, envolve apenas renomear o usuário da caixa e endereço para XX<nomedacaixa> e retirar as permissões
        $caixadepXX = $null
        $caixadepXX = Get-Mailbox -Filter { EmailAddresses -like "*$logindep@*" } # Pesquisa a caixa pelo endereço
        if (!$caixadepXX) { # Se não encontrar pelo endereço, pesquisa pelo nome
            $caixadepXX = Get-Mailbox -Filter { Name -like "*$logindep" } # Pesquisa a caixa pelo nome
        }
        if ($caixadepXX) { # Se encontrar uma caixa com o mesmo endereço ou nome
            # "Desativa" caixa antiga
            "Desativando caixa $emaildep1 antiga." *>> $PathLog
            $enderecos = $null
            # Coloca o prefixo "XX" apenas nos endereços nominais; não altera os endereços baseados no ponto
            $enderecos = $caixadepXX.EmailAddresses.ForEach( { $_ -replace "(smtp:).*(dep|gab)\.(.*)",'$1XX$2.$3' } )
            Set-Mailbox -Identity $caixadepXX.SamAccountName -Name "XX$logindep" -UserPrincipalName ("XX$logindep" + "@redecamara.camara.gov.br") -EmailAddresses $enderecos -EmailAddressPolicyEnabled $false -HiddenFromAddressListsEnabled $true -Confirm:$Confirm -WhatIf:$WhatIf
                if($?){"[Set-Mailbox] Caixa $logindep renomeada para XX$logindep ." *>> $PathLog}
                    else{"[Set-Mailbox] Ocorreu um erro ao renomear a caixa $logindep para XX$logindep ." *>> $PahLogErro}
            Set-ADUser -Identity $caixadepXX.SamAccountName -SamAccountName ("XX" + $caixadepXX.SamAccountName) -Confirm:$Confirm -WhatIf:$WhatIf
                if($?){"[Set-ADUser] Usuario $logindep renomeado para XX$logindep ." *>> $PathLog}
                    else{"[Set-ADUser] Ocorreu um erro ao renomear o usuario $logindep para XX$logindep ." *>> $PahLogErro}
            # Retira a caixa desativada das listas de distribuição
            Remove-ADGroupMember -Identity Deputados -Members ("XX" + $caixadepXX.SamAccountName) -Confirm:$Confirm -WhatIf:$WhatIf
            <#
            Trecho descontinuado; isso não é mais necessário pois o IDEA (em conjunto com o SisDelagações) é quem gerencia essas permissões
            # Remove as permissões do grupo ...-u
            Get-MailboxPermission -Identity "XX$logindep" | Where-Object { $_.User -Match ".*\-u" } | ForEach-Object { 
                Remove-MailboxPermission -Identity $_.Identity -User $_.User -AccessRights FullAccess -InheritanceType All -Confirm:$Confirm -WhatIf:$WhatIf
                    if($?){"[Remove-MailboxPermission] Caixa XX$logindep com permissões FullAccess excluídas." *>> $PathLog}
                        else{"[Remove-MailboxPermission] Ocorreu um erro ao excluir permissões da caixa XX$logindep ." *>> $PahLogErro}
                Remove-ADPermission -Identity $_.Identity -User $_.User -ExtendedRights Send-As -Confirm:$Confirm -WhatIf:$WhatIf
                    if($?){"[Remove-ADPermission] Caixa XX$logindep com permissões Send-As excluídas." *>> $PathLog}
                        else{"[Remove-ADPermission] Ocorreu um erro ao excluir permissões da caixa XX$logindep ." *>> $PahLogErro}
            }
            Fim do trecho descontinuado
            #>
            $CaixasDepDesativadas++
        }
        # Cria nova caixa dep
        $samaccountnamedep
        New-Mailbox -Shared -Name $logindep -SamAccountName $samaccountnamedep -DisplayName $nomedep -UserPrincipalName ("$logindep" + "@redecamara.camara.gov.br") -PrimarySmtpAddress $emaildep1 -Alias $ponto56 -FirstName $primeironome -LastName $sobrenome -OrganizationalUnit $OUcaixa -Confirm:$Confirm -WhatIf:$WhatIf
            if($?){"[New-Mailbox] Caixa " + $logindep + " criada." *>> $PathLog}
                else{"[New-Mailbox] Ocorreu um erro ao criar a caixa " + $logindep + "." *>> $PahLogErro}
        Set-Mailbox -Identity $logindep -EmailAddresses @{Add="smtp:$emaildep2","smtp:$emaildep3","smtp:$emaildep4"} -EmailAddressPolicyEnabled $false -Confirm:$Confirm -WhatIf:$WhatIf
            if($?){"[Set-Mailbox] Caixa " + $logindep + " com endereços alternativos incluidos." *>> $PathLog}
                else{"[Set-Mailbox] Ocorreu um erro ao incluir endereços alternativos na caixa " + $logindep + "." *>> $PahLogErro}
        Add-MailboxPermission -Identity $logindep -User $grupodep -AccessRights FullAccess -InheritanceType All -Confirm:$Confirm -WhatIf:$WhatIf
            if($?){"[Add-MailboxPermission] Caixa " + $logindep + " com permissões incluidas." *>> $PathLog}
                else{"[Add-MailboxPermission] Ocorreu um erro ao incluir permissões na caixa " + $logindep + "." *>> $PahLogErro}
        Add-ADPermission -Identity $logindep -User $grupodep -ExtendedRights Send-As -Confirm:$Confirm -WhatIf:$WhatIf
            if($?){"[Add-ADPermission] Caixa " + $logindep + " com permissões incluidas." *>> $PathLog}
                else{"[Add-ADPermission] Ocorreu um erro ao incluir permissões na caixa " + $logindep + "." *>> $PahLogErro}
        Set-ADUser -Identity $samaccountnamedep -Title $titulo -Description $descricao -Confirm:$Confirm -WhatIf:$WhatIf
            if($?){"[Set-ADUser] Caixa " + $samaccountnamedep + " com titulo e descricao incluidos." *>> $PathLog}
                else{"[Set-ADUser] Ocorreu um erro ao incluir titulo e descricao na caixa " + $samaccountnamedep + "." *>> $PahLogErro}
        Set-ADUser -Identity $ponto56 -EmailAddress $emaildep1 -Confirm:$Confirm -WhatIf:$WhatIf
            if($?){"[Set-ADUser] Usuário " + $ponto56 + " com endereco dep incluído." *>> $PathLog}
                else{"[Set-ADUser] Ocorreu um erro ao incluir endereco dep no usuario " + $ponto56 + "." *>> $PahLogErro}
        Add-ADGroupMember -Identity Deputados -Members $samaccountnamedep -Confirm:$Confirm -WhatIf:$WhatIf
        Add-ADGroupMember -Identity ExchangePerfil1 -Members $samaccountnamedep -Confirm:$Confirm -WhatIf:$WhatIf
            if($?){"[Add-ADGroupMember] Caixa " + $logindep + " com permissoes incluidas." *>> $PathLog}
                else{"[Add-ADGroupMember] Ocorreu um erro ao incluir permissoes na caixa " + $logindep + "." *>> $PahLogErro}
        # Loga que a caixa dep foi criada
        "Caixa " + $emaildep1 + " criada." *>> $PathLog
        $CaixasDepCriadas++
    }
    else { # Se a caixa dep já existe...
        # Redefine os endereços de e-mail e muda a "OU"
        Set-Mailbox -Identity $caixadep.Identity -EmailAddresses "SMTP:$emaildep1","smtp:$emaildep2","smtp:$emaildep3","smtp:$emaildep4" -EmailAddressPolicyEnabled $false -Alias $ponto56 -Confirm:$Confirm -WhatIf:$WhatIf
            if($?){"[Set-Mailbox] Caixa " + $logindep + " com endereços alternativos incluidos." *>> $PathLog}
                else{"[Set-Mailbox] Ocorreu um erro ao incluir endereços alternativos na caixa " + $logindep + " (talvez eles já existissem)." *>> $PahLogErro}
        Add-ADGroupMember -Identity Deputados -Members $caixadep.Identity -Confirm:$Confirm -WhatIf:$WhatIf
        Add-ADGroupMember -Identity ExchangePerfil1 -Members $caixadep.Identity -Confirm:$Confirm -WhatIf:$WhatIf
        if ($caixadep.OrganizationalUnit -ne $OUcaixa) {
            Move-ADObject -Identity $caixadep.Identity -TargetPath $OUcaixa -Confirm:$Confirm -WhatIf:$WhatIf
                if($?){"[Move-ADObject] Caixa " + $caixadep.Identity + " movida para " + $OUcaixa + "." *>> $PathLog}
                    else{"[Move-ADObject] Ocorreu um erro ao mover a caixa " + $caixadep.Identity + " para " + $OUcaixa + "." *>> $PahLogErro}
        }
        # Loga que a caixa dep já existe
        "Caixa " + $emaildep1 + " já existente e não foi criada." *>> $PathLog
        # Atualiza o SamAccountName da caixa gab com base na caixa dep para manter a rastreabilidade entre as duas caixas
        $samaccountnamegab = $caixadep.SamAccountName.Replace("dep.", "gab.")
        $CaixasDepNaoCriadas++
    }
    $CaixasDepTotal++

    # Cria caixa gab

    # Verifica se a caixa gab já existe
    $caixagab = $null
    # Conforme já dito acima, somente tenta reaproveitar a caixa gab se a dep já foi encontrada
    if ($caixadep) { # Se a caixa dep foi encontrada...
        # Pesquisa a caixa gab com base em uma derivação do SamAccountName da caixa dep
        $caixagab = Get-Mailbox -Identity $samaccountnamegab
        if (!$caixagab) { # Se não achou...
            # Busca pelo email mesmo
            $caixagab = Get-Mailbox -Identity $emailgab1
        }
    }
    if (!$caixagab) { # Se a caixa gab ainda não existe...
        # Se assegura que qualquer caixa antiga com o mesmo endereço seja "desativada" 
        # A desativação, na verdade, envolve apenas renomear o usuário da caixa e endereço para XX<nomedacaixa> e retirar as permissões
        $caixagabXX = $null
        $caixagabXX = Get-Mailbox -Identity $emailgab1 # Pesquisa a caixa pelo endereço
        if (!$caixagabXX) { # Se não encontrar pelo endereço, pesquisa pelo nome
            $caixagabXX = Get-Mailbox -Identity $logingab # Pesquisa a caixa pelo nome
        }
        if ($caixagabXX) { # Se encontrar uma caixa com o mesmo endereço ou nome
            # "Desativa" caixa antiga
            "Desativando caixa $emailgab1 antiga." *>> $PathLog
            $enderecos = $null
            # Coloca o prefixo "XX" apenas nos endereços nominais; não altera os endereços baseados no ponto
            $enderecos = $caixadepXX.EmailAddresses.ForEach({$_ -replace "(smtp:).*(dep|gab)\.(.*)",'$1XX$2.$3'})
            Set-Mailbox -Identity $caixagabXX.SamAccountName -Name "XX$logingab" -UserPrincipalName ("XX$logingab" + "@redecamara.camara.gov.br") -EmailAddresses $enderecos -EmailAddressPolicyEnabled $false -HiddenFromAddressListsEnabled $true -Confirm:$Confirm -WhatIf:$WhatIf
                if($?){"[Set-Mailbox] Caixa $logingab renomeada para XX$logingab ." *>> $PathLog}
                    else{"[Set-Mailbox] Ocorreu um erro ao renomear a caixa $logingab para XX$logingab ." *>> $PahLogErro}
            Set-ADUser -Identity $caixagabXX.SamAccountName -SamAccountName ("XX" + $caixagabXX.SamAccountName) -Confirm:$Confirm -WhatIf:$WhatIf
                if($?){"[Set-ADUser] Usuario $logingab renomeado para XX$logingab ." *>> $PathLog}
                    else{"[Set-ADUser] Ocorreu um erro ao renomear o usuario $logingab para XX$logingab ." *>> $PahLogErro}
            # Retira a caixa desativada das listas de distribuição
            Remove-ADGroupMember -Identity GabDep -Members ("XX" + $caixagabXX.SamAccountName) -Confirm:$Confirm -WhatIf:$WhatIf
            <#
            Trecho descontinuado; isso não é mais necessário pois o IDEA (em conjunto com o SisDelagações) é quem gerencia essas permissões
            # Remove as permissões do grupo ...-u
            Get-MailboxPermission -Identity "XX$logingab" | Where-Object { $_.User -Match ".*\-u" } | ForEach-Object { 
                Remove-MailboxPermission -Identity $_.Identity -User $_.User -AccessRights FullAccess -InheritanceType All -Confirm:$Confirm -WhatIf:$WhatIf
                    if($?){"[Remove-MailboxPermission] Caixa XX$logingab com permissões FullAccess excluídas." *>> $PathLog}
                        else{"[Remove-MailboxPermission] Ocorreu um erro ao excluir permissões da caixa XX$logingab ." *>> $PahLogErro}
                Remove-ADPermission -Identity $_.Identity -User $_.User -ExtendedRights Send-As -Confirm:$Confirm -WhatIf:$WhatIf
                    if($?){"[Remove-ADPermission] Caixa XX$logingab com permissões Send-As excluídas." *>> $PathLog}
                        else{"[Remove-ADPermission] Ocorreu um erro ao excluir permissões da caixa XX$logingab ." *>> $PahLogErro}
            }
            Fim do trecho descontinuado
            #>
            $CaixasGabDesativadas++
        }
        # Cria nova caixa gab
        $samaccountnamegab
        New-Mailbox -Shared -Name $logingab -SamAccountName $samaccountnamegab -DisplayName $nomegab -UserPrincipalName ("$logingab" + "@redecamara.camara.gov.br") -PrimarySmtpAddress $emailgab1 -Alias $logingab -FirstName $primeironome -LastName $sobrenome -OrganizationalUnit $OUcaixa -Confirm:$Confirm -WhatIf:$WhatIf
            if($?){"[New-Mailbox] Caixa " + $logingab + " criada." *>> $PathLog}
                else{"[New-Mailbox] Ocorreu um erro ao criar a caixa " + $logingab + "." *>> $PahLogErro}
        Set-Mailbox -Identity $logingab -EmailAddresses @{Add="smtp:$emailgab2"} -EmailAddressPolicyEnabled $false -Confirm:$Confirm -WhatIf:$WhatIf
            if($?){"[Set-Mailbox] Caixa " + $logingab + " com endereços alternativos incluidos." *>> $PathLog}
                else{"[Set-Mailbox] Ocorreu um erro ao incluir endereços alternativos na caixa " + $logingab + "." *>> $PahLogErro}
        Add-MailboxPermission -Identity $logingab -User $grupogab -AccessRights FullAccess -InheritanceType All -Confirm:$Confirm -WhatIf:$WhatIf
            if($?){"[Add-MailboxPermission] Caixa " + $logingab + " com permissões incluidas." *>> $PathLog}
                else{"[Add-MailboxPermission] Ocorreu um erro ao incluir permissões na caixa " + $logingab + "." *>> $PahLogErro}
        Add-ADPermission -Identity $logingab -User $grupogab -ExtendedRights Send-As -Confirm:$Confirm -WhatIf:$WhatIf
            if($?){"[Add-ADPermission] Caixa " + $logingab + " com permissões incluidas." *>> $PathLog}
                else{"[Add-ADPermission] Ocorreu um erro ao incluir permissões na caixa " + $logingab + "." *>> $PahLogErro}
        Set-ADUser -Identity $samaccountnamegab -Title $titulogab -Company $empresagab -Department $ponto56 -Description $descricao -Confirm:$Confirm -WhatIf:$WhatIf
            if($?){"[Set-ADUser] Caixa " + $samaccountnamegab + " com titulo e descricao incluidos." *>> $PathLog}
                else{"[Set-ADUser] Ocorreu um erro ao incluir titulo e descricao na caixa " + $samaccountnamegab + "." *>> $PahLogErro}
        Add-ADGroupMember -Identity GabDep -Members $samaccountnamegab -Confirm:$Confirm -WhatIf:$WhatIf
        Add-ADGroupMember -Identity ExchangePerfil1 -Members $samaccountnamegab -Confirm:$Confirm -WhatIf:$WhatIf
        Add-ADGroupMember -Identity MailboxSemSMTP -Members $samaccountnamegab -Confirm:$Confirm -WhatIf:$WhatIf
            if($?){"[Add-ADGroupMember] Caixa " + $logingab + " com permissoes incluidas." *>> $PathLog}
                else{"[Add-ADGroupMember] Ocorreu um erro ao incluir permissoes na caixa " + $logingab + "." *>> $PahLogErro}
        # Loga que a caixa gab foi criada
        "Caixa " + $emailgab1 + " criada." *>> $PathLog
        $CaixasGabCriadas++
    }
    else { # Se a caixa gab já existe...
        # Atualiza definições e muda a OU
        Set-Mailbox -Identity $caixagab.Identity -EmailAddresses "SMTP:$emailgab1","smtp:$emailgab2" -EmailAddressPolicyEnabled $false -Confirm:$Confirm -WhatIf:$WhatIf
            if($?){"[Set-Mailbox] Caixa " + $logingab + " com endereços alternativos incluidos." *>> $PathLog}
                else{"[Set-Mailbox] Ocorreu um erro ao incluir endereços alternativos na caixa " + $logindep + " (talvez eles já existissem)." *>> $PahLogErro}
        Add-MailboxPermission -Identity $caixagab.Identity -User $grupogab -AccessRights FullAccess -InheritanceType All -Confirm:$Confirm -WhatIf:$WhatIf
            if($?){"[Add-MailboxPermission] Caixa " + $logingab + " com permissões incluidas." *>> $PathLog}
                else{"[Add-MailboxPermission] Ocorreu um erro ao incluir permissões na caixa " + $logingab + "." *>> $PahLogErro}
        Add-ADPermission -Identity $caixagab.SamAccountName -User $grupogab -ExtendedRights Send-As -Confirm:$Confirm -WhatIf:$WhatIf
            if($?){"[Add-ADPermission] Caixa " + $logingab + " com permissões incluidas." *>> $PathLog}
                else{"[Add-ADPermission] Ocorreu um erro ao incluir permissões na caixa " + $logingab + "." *>> $PahLogErro}
        Set-ADUser -Identity $caixagab.SamAccountName -Title $titulogab -Company $empresagab -Department $ponto56 -Description $descricao -Confirm:$Confirm -WhatIf:$WhatIf
            if($?){"[Set-ADUser] Caixa " + $samaccountnamegab + " com titulo e descricao incluidos." *>> $PathLog}
                else{"[Set-ADUser] Ocorreu um erro ao incluir titulo e descricao na caixa " + $samaccountnamegab + "." *>> $PahLogErro}
        Add-ADGroupMember -Identity GabDep -Members $caixagab.SamAccountName -Confirm:$Confirm -WhatIf:$WhatIf
        Add-ADGroupMember -Identity ExchangePerfil1 -Members $caixagab.SamAccountName -Confirm:$Confirm -WhatIf:$WhatIf
        Add-ADGroupMember -Identity MailboxSemSMTP -Members $caixagab.SamAccountName -Confirm:$Confirm -WhatIf:$WhatIf
        if ($caixagab.OrganizationalUnit -ne $OUcaixa) {
            Move-ADObject -Identity $caixagab.Identity -TargetPath $OUcaixa -Confirm:$Confirm -WhatIf:$WhatIf
                if($?){"[Move-ADObject] Caixa " + $caixagab.Identity + " movida para " + $OUcaixa + "." *>> $PathLog}
                    else{"[Move-ADObject] Ocorreu um erro ao mover a caixa " + $caixagab.Identity + " para " + $OUcaixa + "." *>> $PahLogErro}
        }
        # Loga que a caixa gab já existe
        "Caixa " + $emailgab1 + " já existente e não foi criada." *>> $PathLog
        $CaixasGabNaoCriadas++
    }
    $CaixasGabTotal++

}

# Registra no Log um relatório da execução do Script
"Relatório de usuários:" *>> $PathLog
"Total de usuários na lista: $UsuariosTotal" *>> $PathLog
"Usuários criados: $UsuariosCriados" *>> $PathLog
"Usuários renomeados: $UsuariosRenomeados" *>> $PathLog
"Usuários que já existiam: $UsuariosNaoCriados" *>> $PathLog

"Relatório de caixas dep:" *>> $PathLog
"Total de caixas dep na lista: $CaixasDepTotal" *>> $PathLog
"Caixas dep desativadas: $CaixasDepDesativadas" *>> $PathLog
"Caixas dep criadas: $CaixasDepCriadas" *>> $PathLog
"Caixas dep que já existiam: $CaixasDepNaoCriadas" *>> $PathLog

"Relatório de caixas gab:" *>> $PathLog
"Total de caixas gab na lista: $CaixasGabTotal" *>> $PathLog
"Caixas gab desativadas: $CaixasGabDesativadas" *>> $PathLog
"Caixas gab criadas: $CaixasGabCriadas" *>> $PathLog
"Caixas gab que já existiam: $CaixasGabNaoCriadas" *>> $PathLog

if ($UsuariosRenomeados -ne $CaixasDepNaoCriadas) {
    "ATENÇÃO: Existe uma diferença entre a qtde. de usuários renomeados e a qtde. de caixas que já existiam (deveriam ser iguais)." *>> $PathLog
    ">>>>>>>> Sugere-se revisar os dados do arquivo csv." *>> $PathLog
}

# Finaliza sessão

Remove-PSSession $Session

# Registra no Log a data e a hora do FIM da execução do script.
"Fim do script:" + (Get-Date) *>> $PathLog

# Fim do Script ------------------------------------------------------------------