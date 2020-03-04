#New-ADGroup -Name "DitecW10" -SamAccountName DitecW10 -GroupCategory Security -GroupScope Global -DisplayName "DitecW10" -Path "ou=Usuarios,dc=redecamara,dc=camara,dc=gov,dc=br" -Description "Estacoes da Ditec com Windows 10"
$ArqLista = $null
$ArqLista = Import-Csv -Path $PathLista -Delimiter ";" -Encoding Default
Add-ADGroupMember -Identity DitecW10 -Members 