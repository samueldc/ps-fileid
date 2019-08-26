<#
Programa: corrigePendenciasCaixasEmailDeputados
Objetivo: Facilitar a correção de pendências de caixas postais de Deputados recém empossados
Autor: Samuel Diniz Casimiro - P_7029
Setor: Ditec/Coaus/Satus
Versões: 1.0 - Criação do script
#>

# Variáveis globais -----------------------------------------------------------------------

# Se true, ativa o mode de teste (dry-run) nos comandos que utilizam este parâmetro
$WhatIf = $false
$Confirm = $false

# Caminho e nome dos arquivos de log
$PathLog = "C:\Users\p_7029\ownCloud\_trabalho\coaus-satus\Posse 2019\corrigePendenciasCaixasEmailDeputados-" + (Get-Date -UFormat %Y-%m-%d-%H-%M) + ".log"
$PathLogErro = "C:\Users\p_7029\ownCloud\_trabalho\coaus-satus\Posse 2019\corrigePendenciasCaixasEmailDeputados-" + (Get-Date -UFormat %Y-%m-%d-%H-%M) + ".logerro"

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

    # Corrige ponto
    Set-ADUser -Identity $ponto56 -DisplayName $nomedep -GivenName $primeironome -Surname $sobrenome -Description $nomegab -Title $titulo -Department $nomegab -Enabled $true -Confirm:$Confirm -WhatIf:$WhatIf
    if($?){"[Set-ADUser] Usuario " + $ponto56 + " atualizado." *>> $PathLog}
        else{"[Set-ADUser] Ocorreu um erro ao atualizar o usuario " + $ponto56 + "." *>> $PahLogErro}

    # Corrige caixa dep
    $caixadep = $null
    $caixadep = Get-ADUser -Identity ("CN=" + $logindep + ",OU=Legislatura55,OU=CaixasPostaisInstitucionais,OU=Correio,OU=Usuarios,DC=redecamara,DC=camara,DC=gov,DC=br")
    if(!$caixadep) {
        $caixadep = Get-ADUser -Identity ("CN=" + $logindep + ",OU=Legislatura56,OU=CaixasPostaisInstitucionais,OU=Correio,OU=Usuarios,DC=redecamara,DC=camara,DC=gov,DC=br")
    }
    Set-ADUser -Identity $caixadep.DistinguishedName -Title $titulo -GivenName $primeironome -Description $descricao -Confirm:$Confirm -WhatIf:$WhatIf
        if($?){"[Set-ADUser] Caixa " + $samaccountnamedep + " com titulo e descricao incluidos." *>> $PathLog}
            else{"[Set-ADUser] Ocorreu um erro ao incluir titulo e descricao na caixa " + $samaccountnamedep + "." *>> $PahLogErro}
    Add-ADGroupMember -Identity Deputados -Members $caixadep.DistinguishedName -Confirm:$Confirm -WhatIf:$WhatIf
    Add-ADGroupMember -Identity ExchangePerfil1 -Members $caixadep.DistinguishedName -Confirm:$Confirm -WhatIf:$WhatIf
        if($?){"[Add-ADGroupMember] Caixa " + $samaccountnamedep + " com permissoes incluidas." *>> $PathLog}
            else{"[Add-ADGroupMember] Ocorreu um erro ao incluir permissoes na caixa " + $samaccountnamedep + "." *>> $PahLogErro}
    Move-ADObject -Identity $caixadep.DistinguishedName -TargetPath $OUcaixa -Confirm:$Confirm -WhatIf:$WhatIf
        if($?){"[Move-ADObject] Caixa " + $samaccountnamedep + " movida para " + $OUcaixa + "." *>> $PathLog}
            else{"[Move-ADObject] Ocorreu um erro ao mover a caixa " + $samaccountnamedep + " para " + $OUcaixa + "." *>> $PahLogErro}
    
    # Corrige caixa gab
    $caixagab = $null
    $caixagab = Get-ADUser -Identity ("CN=" + $logingab + ",OU=Legislatura55,OU=CaixasPostaisInstitucionais,OU=Correio,OU=Usuarios,DC=redecamara,DC=camara,DC=gov,DC=br")
    if(!$caixagab) {
        $caixagab = Get-ADUser -Identity ("CN=" + $logingab + ",OU=Legislatura56,OU=CaixasPostaisInstitucionais,OU=Correio,OU=Usuarios,DC=redecamara,DC=camara,DC=gov,DC=br")
    }
    Set-ADUser -Identity $caixagab.DistinguishedName -Title $titulogab -GivenName $primeironome -Company $empresagab -Department $ponto56 -Description $descricao -Confirm:$Confirm -WhatIf:$WhatIf
        if($?){"[Set-ADUser] Caixa " + $samaccountnamegab + " com titulo e descricao incluidos." *>> $PathLog}
            else{"[Set-ADUser] Ocorreu um erro ao incluir titulo e descricao na caixa " + $samaccountnamegab + "." *>> $PahLogErro}
    Add-ADGroupMember -Identity GabDep -Members $caixagab.DistinguishedName -Confirm:$Confirm -WhatIf:$WhatIf
    Add-ADGroupMember -Identity ExchangePerfil1 -Members $caixagab.DistinguishedName -Confirm:$Confirm -WhatIf:$WhatIf
    Add-ADGroupMember -Identity MailboxSemSMTP -Members $caixagab.DistinguishedName -Confirm:$Confirm -WhatIf:$WhatIf
        if($?){"[Add-ADGroupMember] Caixa " + $logingab + " com permissoes incluidas." *>> $PathLog}
            else{"[Add-ADGroupMember] Ocorreu um erro ao incluir permissoes na caixa " + $logingab + "." *>> $PahLogErro}
    Move-ADObject -Identity $caixagab.DistinguishedName -TargetPath $OUcaixa -Confirm:$Confirm -WhatIf:$WhatIf
        if($?){"[Move-ADObject] Caixa " + $samaccountnamegab + " movida para " + $OUcaixa + "." *>> $PathLog}
            else{"[Move-ADObject] Ocorreu um erro ao mover a caixa " + $samaccountnamegab + " para " + $OUcaixa + "." *>> $PahLogErro}

}

# Finaliza sessão

Remove-PSSession $Session

# Registra no Log a data e a hora do FIM da execução do script.
"Fim do script:" + (Get-Date) *>> $PathLog

# Fim do Script ------------------------------------------------------------------