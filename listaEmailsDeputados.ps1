<#
Programa: listaEmailsDeputados
Objetivo: Lista e-mails de deputados
Autor: Samuel Diniz Casimiro - P_7029
Setor: Ditec/Coaus/Satus
Versões: 1.0 - Criação do script
Dependências: Módulo https://www.powershellgallery.com/packages/FileSystemForms/
#>

# Seleciona a pasta
$File = Select-FileSystemForm
# Busca os objetos e escreve os e-mails em um arquivo na pasta selecionada
Get-ADGroupMember -Identity Deputados | ForEach-Object { ( $_.name + "@camara.leg.br" ) *>> ( $File + "\ListaEmailsDeputados.txt" ) }
