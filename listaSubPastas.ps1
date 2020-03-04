<#
Programa: listaSubPastas
Objetivo: Lista subpastas a partir de uma localização específica
Autor: Samuel Diniz Casimiro - P_7029
Setor: Ditec/Coaus/Satus
Versões: 1.0 - Criação do script
#>
Set-Location \\redecamara\DfsData
Get-ChildItem | ft FullName