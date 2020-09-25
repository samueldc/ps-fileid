<#
Programa: corrigeEnderecoPrimarioCaixasEmailDeputadosSDR
Objetivo: Facilitar a correção do endereço primário de caixas postais de Deputados para uso no SDR
Autor: Samuel Diniz Casimiro - P_7029
Setor: Ditec/Coaus/Satus
Versões: 1.0 - Criação do script
#>

# Variáveis globais -----------------------------------------------------------------------

# Se true, ativa o mode de teste (dry-run) nos comandos que utilizam este parâmetro
$WhatIf = $false
$Confirm = $false

# Caminho e nome do arquivo CSV com a lista de usuarios a ser importada e exportada
$PathLista = "C:\Users\p_7029\ownCloud\_trabalho\teletrabalho\sdr\usuariosSDRFinal.csv"

# DistinguishedName das "OU"s envolvidas
$OU = "OU=56Legislatura,OU=Deputados,OU=Usuarios,DC=redecamara,DC=camara,DC=gov,DC=br"
$OUsdr = "OU=56LegislaturaSDR,OU=Deputados,OU=Usuarios,DC=redecamara,DC=camara,DC=gov,DC=br"

# Aqui começam as ações do script

# Solicita credencial para executar o script (não precisa pq por enquanto estamos utilizando as credenciais do usuário logado na máquina)
##$Credential = Get-Credential

# Cria sessão com o Exchange; a autenticação é feita com as credenciais do usuário logado na máquina
Set-ExecutionPolicy RemoteSigned
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://urca2.redecamara.camara.gov.br/PowerShell/ -Authentication Kerberos
Import-PSSession $Session -DisableNameChecking

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
    $emailsdr = ""

    # Povoa
    $pontosdr = $PSItem.ponto_sdr.Trim()
    $emailsdr = $PSItem.email_sdr.Trim()

    $pontosdr

    Set-Mailbox -Identity $emailsdr -PrimarySmtpAddress $emailsdr -EmailAddressPolicyEnabled $false -Confirm:$Confirm -WhatIf:$WhatIf

}
# Finaliza sessão

#Remove-PSSession $Session

# Fim do Script ------------------------------------------------------------------