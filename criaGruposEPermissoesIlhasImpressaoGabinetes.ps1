<#
Programa: criaGruposEPermissoesIlhasImpressaoGabinetes
Objetivo: Facilitar a criação dos grupos de permissão das filas das ilhas de impressão e atribuir as permissões aos gabinetes
Autor: P_7029 Samuel Diniz Casimiro
Setor: Ditec/Coaus/Satus
Versões: 1.0 - Criação do script
#>

Set-ExecutionPolicy RemoteSigned
Import-Module ImportExcel

# Variáveis globais

# Se true, ativa o mode de teste (dry-run) nos comandos que utilizam este parâmetro
$WhatIf = $false
$Confirm = $false

# Caminho e nome dos arquivos de log
$PathLog = "C:\Temp\criaGruposEPermissoesIlhasImpressaoGabinetes-" + (Get-Date -UFormat %Y-%m-%d-%H-%M) + ".log"
$PathLogErro = "C:\Temp\criaGruposEPermissoesIlhasImpressaoGabinetes-" + (Get-Date -UFormat %Y-%m-%d-%H-%M) + ".logerro"

# DistinguishedName da "OU" onde os usuários serão criados
$OUGrupoFila = "OU=Printers-Ilhas,OU=Grupos,OU=Usuarios,DC=redecamara,DC=camara,DC=gov,DC=br"

<#
Os contadores a seguir servem para gerar um resumo das ações ao final da execução do script.
Esse resumo é incluído no final do arquivo de log gerado.
#>

# Contadores de grupos
$GruposCriados = 0
$GruposNaoCriados = 0
$GruposTotal = 0

# Contador de permissoes
$PermissoesAtribuidas = 0
$PermissoesNaoAtribuidas = 0
$PermissoesTotal = 0

# Cria arquivo de log e já registra a data e a hora do inicio da execução do script.
"Início do script: " + (Get-Date) *> $PathLog

# Cria arquivo de log de erros
"" *> $PathLogErro

# Registra no log a credencial usada para execução do script
"O script foi executado por: $env:COMPUTERNAME\$env:UserName" *>> $PathLog

# Declaracao de funcoes

function atribuiPermissoes {
    param (
        [string]$NomeGrupoFila, [string]$Anexo, [string]$Andar
    )
    
    $PermissoesTotal = 0

    $gabinetes = $null
    $gabinetes = Import-Excel -Path "c:/Temp/gabinetes.xlsx" -WorkSheetname todos

    $gabinetes | ForEach-Object {
        if ( $PSItem -and $PSItem.carteira -and ( $PSItem.anexo -eq $Anexo ) -and ( $PSItem.andar -eq $Andar ) ) {
            Add-ADGroupMember -Identity $NomeGrupoFila -Members "Dep-$($PSItem.carteira)-P" -Confirm:$Confirm -WhatIf:$WhatIf
                if($?){
                    "[Add-ADGroupMember] Permissao no grupo " + $NomeGrupoFila + " atribuida para " + "Dep-$($PSItem.carteira)-P" + "." *>> $PathLog
                    $PermissoesAtribuidas++
                } else {
                    "[Add-ADGroupMember] Ocorreu um erro ao atribuir permissao no grupo " + $NomeGrupoFila + " para " + "Dep-$($PSItem.carteira)-P" + "." *>> $PahLogErro
                    $PermissoesNaoAtribuidas++
                }
        }
        $PermissoesTotal++ # Será zerado a cada chamada desta funcao, mas no final o valor dele sera sempre atualizado para o total na ultima chamada
    }
}

# Aqui começam as ações do script

$filas = Import-Excel -Path "c:/Temp/ilhasImpressao.xlsx" -WorkSheetname filas

$filas | ForEach-Object {
    if ( $PSItem -and $PSItem.fila ) {
        New-ADGroup -Name $PSItem.fila -SamAccountName $PSItem.fila -GroupCategory Security -GroupScope Global -DisplayName $PSItem.fila -Path $OUGrupoFila -Description $PSItem.fila -Confirm:$Confirm -WhatIf:$WhatIf
            if($?){
                "[New-ADGroup] Grupo " + $PSItem.fila + " criado." *>> $PathLog
                $GruposCriados++
            } else {
                "[New-ADGroup] Ocorreu um erro ao criar o grupo " + $PSItem.fila + "." *>> $PahLogErro
                $GruposNaoCriados++
            }
        atribuiPermissoes -NomeGrupoFila $PSItem.fila -Anexo $PSItem.anexo -Andar $PSItem.andar
        $GruposTotal++
    }
}

# Registra no Log um relatório da execução do Script
"Relatório de grupos:" *>> $PathLog
"Total de grupos na lista: $GruposTotal" *>> $PathLog
"Grupos criados: $GruposCriados" *>> $PathLog
"Grupos não criados ou que já existiam: $GruposNaoCriados" *>> $PathLog

"Relatório de permissoes:" *>> $PathLog
"Total de permissoes na lista: $PermissoesTotal" *>> $PathLog
"Permissoes atribuidas: $PermissoesAtribuidas" *>> $PathLog
"Permissoes nao atribuidas: $PermissoesNaoAtribuidas" *>> $PathLog

if ($PermissoesAtribuidas -ne $PermissoesTotal) {
    "ATENÇÃO: Existe uma diferença entre a qtde. de permissoes na lista e a qtde. de permissoes atribuidas (deveriam ser iguais)." *>> $PathLog
    ">>>>>>>> Sugere-se revisar os dados dos arquivos e logs de erro." *>> $PathLog
}

# Finaliza sessão

#Remove-PSSession $Session

# Registra no Log a data e a hora do FIM da execução do script.
"Fim do script:" + (Get-Date) *>> $PathLog

# Fim do Script ------------------------------------------------------------------