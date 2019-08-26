<#
Programa: criaMapeamentoPastaUsuariosGabinentes
Objetivo: Utilizar o atributo HomeDrive do AD para mapear automaticamente a pasta compartilhada do gabinete
Autor: Samuel Diniz Casimiro - P_7029
Setor: Ditec/Coaus/Satus
Versões: 1.0 - Criação do script
#>

$WhatIf = $false # Modo teste (dry-run) ativado
$carteiras = 1..550 # Execução por faixas; próxima faixa 101..200 e assim por diante
$qtdeOk = 0
$qtdeTotal = 0
# Para cada número de carteira...
foreach ($carteira in $carteiras) {
    $c = $carteira.ToString("000")
    # ... obtém os usuários membros do grupo de permissão na pasta
    # Já é esperado que algumas iterações deem erro, pois nem todas as carteiras são utilizadas por motivos culturais (ex.: 013, 024)
    Get-ADGroupMember -Identity "Dep-$c-F" | ForEach-Object {
        # Exibe os nomes em tela para acompanhamento da execução do script
        $_.name + " Z:\\redecamara\Gabinetes\Dep-56$c" 
        # Define o atributo HomeDrive para cada usuário membro do grupo
        Set-ADUser -Identity $_.distinguishedName -ScriptPath $null -HomeDrive "Z:" -HomeDirectory "\\redecamara\Gabinetes\Dep-56$c" -Confirm:$false -WhatIf:$WhatIf
        if ($?) {
            $qtdeOk++
        }
        $qtdeTotal++    
    }
}
"Registros alterados: $qtdeOk"
"Registros NÃO alterados: " + ($qtdeTotal - $qtdeOk)
"Total de registros processados: $qtdeTotal"


