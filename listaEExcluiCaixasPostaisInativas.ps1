<#
Programa: listaEExcluiCaixasPostaisInativas
Objetivo: Facilitar exclusão de caixas postais inativas
Autor: Samuel Diniz Casimiro - P_7029
Setor: Ditec/Coaus/Satus
Versões: 1.0 - Criação do script
#>

# Variáveis globais -----------------------------------------------------------------------

# Se true, ativa o mode de teste (dry-run) nos comandos que utilizam este parâmetro
$WhatIf = $true
$Confirm = $false

# Caminho e nome dos arquivos de log
$PathLog = "C:\Users\p_7029\Downloads\listaEExcluiCaixasPostaisInativas-" + (Get-Date -UFormat %Y-%m-%d-%H-%M) + ".log"
$PathLogErro = "C:\Users\p_7029\Downloads\listaEExcluiCaixasPostaisInativas-" + (Get-Date -UFormat %Y-%m-%d-%H-%M) + ".logerro"

# Caminho e nome do arquivo CSV com a lista de usuarios a ser importada e exportada
# $PathLista = "C:\Users\p_7029\ownCloud\_trabalho\teletrabalho\sdr\usuariosSDR2.csv"
$PathListaFinal = "C:\Users\p_7029\Downloads\listaEExcluiCaixasPostaisInativasFinal.csv"

<#
Os contadores a seguir servem para gerar um resumo das ações ao final da execução do script.
Esse resumo é incluído no final do arquivo de log gerado.
#>

# Contadores de usuários
$CaixasExcluidas = 0
$CaixasNaoExcluidas = 0
$CaixasTotal = 0

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
"" *> $PathListaFinal

# Registra no log a credencial usada para execução do script
"O script foi executado por: $env:COMPUTERNAME\$env:UserName" *>> $PathLog

# Variáveis gerais

# Redefine para não correr o risco de viciar os valores (trata-se de uma característica do PowerShell que em caso de erro em comando o valor anterior da variável é mantido)

# Povoa as variáveis

$estatisticas = $null
    
# Obtém estatísticas
#Get-MailboxDatabase -Server URCA2 | Get-MailboxStatistics -Filter "(LastLogonTime -lt '06/10/2015')" | Select DataBase, Identity, DisplayName, LastLogoffTime, LastLogonTime, LastLoggedOnUserAccount | ForEach-Object {
Get-MailboxDatabase | Get-MailboxStatistics -Filter "(LastLogonTime -lt '06/10/2015') -and (LastLogonTime -like '*')" | Select DataBase, Identity, DisplayName, LastLogoffTime, LastLogonTime, LastLoggedOnUserAccount | ForEach-Object {

    "$($PSItem.Database);$($PSItem.Identity);$($PSItem.DisplayName);$($PSItem.LastLogoffTime);$($PSItem.LastLogonTime);$($PSItem.LastLoggedOnUserAccount)" *>> $PathListaFinal

    $CaixasTotal++

}

$CaixasTotal

# Registra no Log um relatório da execução do Script
"Relatório de caixas:" *>> $PathLog
"Total de caixas na lista: $CaixasTotal" *>> $PathLog
"Caixas excluídas: $CaixasExcluidas" *>> $PathLog
"Caixas não excluídas: $CaixasNaoExcluidas" *>> $PathLog

if ($CaixasTotal -ne ($CaixasExcluidas + $CaixasNaoExcluidas)) {
    "ATENÇÃO: Existe uma diferença entre a qtde. de caixas total e a qtde. de caixas excluídas e as não excluídas." *>> $PathLog
    ">>>>>>>> Sugere-se revisar os dados do arquivo csv." *>> $PathLog
}

# Finaliza sessão

Remove-PSSession $Session

# Registra no Log a data e a hora do FIM da execução do script.
"Fim do script: $(Get-Date)" *>> $PathLog

# Fim do Script ------------------------------------------------------------------