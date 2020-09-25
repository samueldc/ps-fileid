<#
Programa: criaUsuariosTesteSenadoSDR
Objetivo: Criação de usuários de teste para uso no SDR
Autor: Samuel Diniz Casimiro - P_7029
Setor: Ditec/Coaus/Satus
Versões: 1.0 - Criação do script
#>

# Variáveis globais -----------------------------------------------------------------------

# Se true, ativa o mode de teste (dry-run) nos comandos que utilizam este parâmetro
$WhatIf = $false
$Confirm = $false

# Caminho e nome dos arquivos de log
$PathLog = "C:\Users\p_7029\ownCloud\_trabalho\teletrabalho\sdr\criaUsuariosTesteSenadoSDR-" + (Get-Date -UFormat %Y-%m-%d-%H-%M) + ".log"
$PathLogErro = "C:\Users\p_7029\ownCloud\_trabalho\teletrabalho\sdr\criaUsuariosTesteSenadoSDR-" + (Get-Date -UFormat %Y-%m-%d-%H-%M) + ".logerro"

# Caminho e nome do arquivo CSV com a lista de usuarios a ser importada
$PathLista = "C:\Users\p_7029\ownCloud\_trabalho\teletrabalho\sdr\usuariosTesteSenadoSDR.csv"

# DistinguishedName das "OU"s envolvidas
$OUsdr = "OU=TesteStressSDR,OU=Deputados,OU=Usuarios,DC=redecamara,DC=camara,DC=gov,DC=br"

<#
Os contadores a seguir servem para gerar um resumo das ações ao final da execução do script.
Esse resumo é incluído no final do arquivo de log gerado.
#>

# Contadores de usuários
$UsuariosCriados = 0
$UsuariosNaoCriados = 0
$UsuariosTotal = 0

# Aqui começam as ações do script

# Solicita credencial para executar o script (não precisa pq por enquanto estamos utilizando as credenciais do usuário logado na máquina)
##$Credential = Get-Credential

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
    $pontosdr = ""
    $senhasdr = ""
    $nomesdr = ""

    # Povoa
    $pontosdr = $PSItem.ponto_sdr.Trim()
    $senhasdr = $PSItem.senha_sdr.Trim()
    $nomesdr = $PSItem.nome_sdr.Trim()

    $usuario = $null
    
    $pontosdr
    # Verifica se o usuario do deputado existe
    $usuario = Get-ADUser -Identity $pontosdr -Properties *

    # Se o usuário não existe, cria o usuário de teste
    if (!$?) {

        # Cria a conta
        $usuario = New-ADUser -Name $pontosdr -DisplayName $nomesdr -GivenName "Teste Senador" -Surname "SDR" -Description "Solicitação do Fernando Torres" -Title "Sistema" -Department "DIRETORIA DE INOVAÇÃO E TECNOLOGIA DA INFORMAÇÃO" -UserPrincipalName ($pontosdr + "@redecamara.camara.gov.br") -Path $OUsdr -ChangePasswordAtLogon $false -Enabled $true -AccountPassword (ConvertTo-SecureString -AsPlainText $senhasdr -Force) -PassThru:$true -Confirm:$Confirm -WhatIf:$WhatIf
            if($?){"[New-ADUser] Usuario $pontosdr criado." *>> $PathLog}
                else{"[New-ADUser] Ocorreu um erro ao criar o usuario $pontosdr." *>> $PahLogErro}

            # Adiciona o usuário nos grupos padrão
            Add-ADGroupMember -Identity InfolegSimulacaoSDR -Members $usuario.DistinguishedName -Confirm:$Confirm -WhatIf:$WhatIf
            Add-ADGroupMember -Identity NegarAcessoWifi -Members $usuario.DistinguishedName -Confirm:$Confirm -WhatIf:$WhatIf
            Add-ADGroupMember -Identity Negar_Logon_local_RDP -Members $usuario.DistinguishedName -Confirm:$Confirm -WhatIf:$WhatIf
                if($?){"[Add-ADGroupMember] Usuário $pontosdr com permissoes incluidas." *>> $PathLog}
                    else{"[Add-ADGroupMember] Ocorreu um erro ao incluir permissoes para o usuário $pontosdr." *>> $PahLogErro}
            $UsuariosCriados++
    } else {

        "[New-Mailbox] Usuário $pontosdr já existente." *>> $PahLogErro
        $UsuariosNaoCriados++

    }

    $UsuariosTotal++

}

# Registra no Log um relatório da execução do Script
"Relatório de usuários:" *>> $PathLog
"Total de usuários na lista: $UsuariosTotal" *>> $PathLog
"Usuários criados: $UsuariosCriados" *>> $PathLog
"Usuários com erro na criação: $UsuariosNaoCriados" *>> $PathLog

if ($UsuariosTotal -ne ($UsuariosCriados + $UsuariosNaoCriados)) {
    "ATENÇÃO: Existe uma diferença entre a qtde. de usuários total e a qtde. de usuários criados mais os não encontrados." *>> $PathLog
    ">>>>>>>> Sugere-se revisar os dados do arquivo csv." *>> $PathLog
}

# Registra no Log a data e a hora do FIM da execução do script.
"Fim do script: $(Get-Date)" *>> $PathLog

# Fim do Script ------------------------------------------------------------------