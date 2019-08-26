<#
Programa: listaUsuariosComScriptLogon.ps1
Objetivo: Lista usuários com atributo ScriptPath preenchido
          Útil para saber quem ainda tem script de logon ativado
Autor: Samuel Diniz Casimiro - P_7029
Setor: Ditec/Coaus/Satus
Versões: 1.0 - Criação do script
#>

$WhatIf = $true # Modo teste (dry-run) ativado
$qtdeTotal = 0
$usuarios = Get-ADUser -Filter "ScriptPath -like '*'" -Properties DistinguishedName,Name,ScriptPath,HomeDrive,HomeDirectory | Where-Object {$_.Name -Match "^\w_.+$"}
#$usuarios = Get-ADUser -Filter 'Name -Match "^\w_.+$" -and ScriptPath -like "*"' -Properties DistinguishedName,Name,ScriptPath,HomeDrive,HomeDirectory
$usuarios | ForEach-Object {
    $_.Name
    $qtdeTotal++    
}
"Total de registros processados: $qtdeTotal"
