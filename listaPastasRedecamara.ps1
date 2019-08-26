<#
Programa: listaPastasRedecamara
Objetivo: Lista pastas da redecamara
Autor: Samuel Diniz Casimiro - P_7029
Setor: Ditec/Coaus/Satus
Versões: 1.0 - Criação do script
#>

Set-Location \\redecamara\dfsdata
Get-ChildItem | ft FullName
Set-Location \\redecamara\Gabinetes
Get-ChildItem | ft FullName