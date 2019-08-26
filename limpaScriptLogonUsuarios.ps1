<#
Programa: limpaScriptLogonUsuarios
Objetivo: Limpa atributo ScriptPath de todos os usuários
Autor: Samuel Diniz Casimiro - P_7029
Setor: Ditec/Coaus/Satus
Versões: 1.0 - Criação do script
#>

$WhatIf = $true # Modo teste (dry-run) ativado
$qtdeOk = 0
$qtdeTotal = 0
#$usuarios = Get-ADUser -Filter 'Name -Match "^\w_.+$" -and ScriptPath -like "*"' -Properties DistinguishedName,Name,ScriptPath,HomeDrive,HomeDirectory
$usuarios = Get-ADUser -Filter "ScriptPath -like '*'" -Properties DistinguishedName,Name,ScriptPath,HomeDrive,HomeDirectory | Where-Object {$_.Name -Match "^\w_.+$"}
$usuarios | ForEach-Object {
    if ($_.ScriptPath) {
        $_.Name
        Set-ADUser -Identity $_.distinguishedName -ScriptPath $null -Confirm:$false -WhatIf:$WhatIf
        if ($?) {
            $qtdeOk++
        }
    }
    $qtdeTotal++    
}
"Registros alterados: $qtdeOk"
"Registros NÃO alterados: " + ($qtdeTotal - $qtdeOk)
"Total de registros processados: $qtdeTotal"
