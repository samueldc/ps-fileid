# Script RemovePermissoesServidoresInativos.ps1
<# .SYNOPSIS
     Script para limpar permissões de usuarios inativos
.DESCRIPTION
     Verifica as OU definidas abaixo em busca de usuarios com a informacao 'Servidor Inativo'.
     Em seguida move esses usuarios para uma OU especifica de usuarios inativos do ano corrente.
     Depois remove todos os grupos desses usuarios e por fim os inclue apenas nos grupos minimos.
.NOTES
     Author: Samuel Diniz Casimiro - samuel.casimiro@camara.leg.br
.LINK
     https://git.camara.gov.br/coaus-satus/scripts-powershell-sepac
#>

# Se true, ativa o mode de teste (dry-run) nos comandos que utilizam este parâmetro
$WhatIf = $false
$Confirm = $false

# Lista de OUs
$OUInativos = "OU=Inativos,OU=Usuarios,DC=redecamara,DC=camara,DC=gov,DC=br"
$OUCenin = "OU=Cenin,OU=Usuarios,DC=redecamara,DC=camara,DC=gov,DC=br"
$OUFuncionarios = "OU=Funcionarios,OU=Usuarios,DC=redecamara,DC=camara,DC=gov,DC=br"
$OUSECOM = "OU=SECOM,OU=Usuarios,DC=redecamara,DC=camara,DC=gov,DC=br"
$OUSuporteTecnico = "OU=Suporte Tecnico,OU=Usuarios,DC=redecamara,DC=camara,DC=gov,DC=br"

# Array de OUs
$OUs = $OUCenin, $OUFuncionarios, $OUSECOM, $OUSuporteTecnico

# Obtem ano corrente
$Ano = (Get-Date).Year

# Verifica se OU do ano já existe
Get-ADOrganizationalUnit -Identity $OUInativos
    If (!$?) { # Caso não exista, cria uma nova
        "Criando nova OU para o ano " + $Ano
        New-ADOrganizationalUnit -Name $Ano -Path $OUInativos -ProtectedFromAccidentalDeletion $true -WhatIf:$WhatIf -Confirm:$Confirm
        if ($?) { # Caso a OU seja criada com sucesso, atualiza o caminho completo da OU de Inativos
            "OU crida com sucesso"
            $OUInativos = "OU=" + $Ano + "," + $OUInativos
        } else { # Caso contrário, interrompe a execução do script
            "Não foi possível criar a OU"
            Pause
            Exit
        }
    } else {
        "OU do ano corrente ja existente"
        $OUInativos = "OU=" + $Ano + "," + $OUInativos
    }

# Localiza as contas com o cargo "Servidor Inativo" nas OUs desejadas e move para a OU de Inativos 
"Localizando e movendo as caixas inativas para a nova OU"
$OUs | ForEach-Object {
    Get-ADUser -Filter {Title -like "Servidor Inativo"} -SearchBase $_ |  Move-ADObject -TargetPath $OUInativos -WhatIf:$WhatIf -Confirm:$Confirm
}

# Seleciona os usuarios inativados
"Selecionando os usuarios inativados"
$Usuarios = Get-ADUser -Filter * -SearchBase $OUInativos -Properties MemberOf
    if (!$?) {
        "Nenhum usuario encontrado!"
        Pause
        Exit
    }

# Lista usuários
"Listando usuarios"
$Usuarios | ft -Property Name, GivenName, Surname

# Pede confirmação para prosseguir
$R = Read-Host -Prompt ("Foram encontrados " + ([array]$Usuarios).Length + " usuários; deseja prosseguir (S/N)?")
    if ($R -eq "N") {
        # Script interrompido
        Pause
        Exit
    } elseif ($R -eq "S") {
        # Remove todos os grupos dos usuarios inativados
        " Removendo grupos"
        $Usuarios | ForEach-Object { 
            $_.MemberOf | # Repare que o nome do membro eh repassado via pipeline para o parametro -Identity do comando Remove-ADGroupMember
                Remove-ADGroupMember -Members $_.DistinguishedName -WhatIf:$WhatIf -Confirm:$Confirm 
        }
        if ($?) {
            "Grupos removidos"
        } else {
            "Ocorreu um erro ao remover os grupos; verifique!"
        }

        # Adiciona os grupos básicos
        "Adicionando grupos básicos"
        Add-ADGroupMember -Identity ExchangePerfil4 -Members $Usuarios -WhatIf:$WhatIf -Confirm:$Confirm
        Add-ADGroupMember -Identity FuncInativos -Members $Usuarios -WhatIf:$WhatIf -Confirm:$Confirm
        Add-ADGroupMember -Identity Inativos -Members $Usuarios -WhatIf:$WhatIf -Confirm:$Confirm
        Add-ADGroupMember -Identity Internet -Members $Usuarios -WhatIf:$WhatIf -Confirm:$Confirm
        if ($?) {
            "Grupos básicos adicionados"
        } else {
            "Ocorreu um erro ao adicionar os grupos básicos; verifique!"
        }

    }