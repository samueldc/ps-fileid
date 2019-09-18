<#
Programa: listaGruposNomeData
Objetivo: Lista grupos do AD por data de criação
Autor: Samuel Diniz Casimiro - P_7029
Setor: Ditec/Coaus/Satus
Versões: 1.0 - Criação do script
#>
$grupos = Get-ADGroup -Filter "Name -like '*localadmin*'" | Get-ADObject -Properties * | Where-Object { $_.WhenCreated -lt '2016-01-01' }
$grupos | Format-Table Name,WhenCreated,WhenChanged
$grupos.Count