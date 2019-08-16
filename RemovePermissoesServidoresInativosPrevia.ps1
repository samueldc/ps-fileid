<#
   Script para verificar servidores inativos. 
   12/08/2019

   - Localizar as contas com o cargo "Servidor Inativo" em 
        * OU aonde possui contas de servidores /Usuarios/Cenin/
        * OU aonde possui contas de servidores /Usuarios/Funcionarios/
        * OU aonde possui contas de servidores /Usuarios/SECOM/
        * OU aonde possui contas de servidores /Usuarios/Suporte Tecnico/

    - Mover os pontos para ou /Usuarios/Inativos/Funcionarios/

    - Deixar somente os grupos de permissão: 
		Domain Users 
		ExchangePerfil4
		FuncInativos
		Inativos
		Internet
#>


$Contas = 0

# Localiza as contas com o cargo "Servidor Inativo" nas OUs: /Cenin /Funcionarios /SECOM Suporte Tecnico/ e move para OU /Inativos/0 
Get-ADUser -Filter {Title -like 'Teste'} -Properties * -SearchBase "OU=Cenin,OU=Usuarios,DC=redecamara,DC=camara,DC=gov,DC=br" |  Move-ADObject -TargetPath "OU=0,OU=Inativos,OU=Usuarios,DC=redecamara,DC=camara,DC=gov,DC=br"
Get-ADUser -Filter {Title -like 'Teste'} -Properties * -SearchBase "OU=Funcionarios,OU=Usuarios,DC=redecamara,DC=camara,DC=gov,DC=br" |  Move-ADObject -TargetPath "OU=0,OU=Inativos,OU=Usuarios,DC=redecamara,DC=camara,DC=gov,DC=br"
Get-ADUser -Filter {Title -like 'Teste'} -Properties * -SearchBase "OU=SECOM,OU=Usuarios,DC=redecamara,DC=camara,DC=gov,DC=br" |  Move-ADObject -TargetPath "OU=0,OU=Inativos,OU=Usuarios,DC=redecamara,DC=camara,DC=gov,DC=br"
Get-ADUser -Filter {Title -like 'Teste'} -Properties * -SearchBase "OU=Suporte Tecnico,OU=Usuarios,DC=redecamara,DC=camara,DC=gov,DC=br" |  Move-ADObject -TargetPath "OU=0,OU=Inativos,OU=Usuarios,DC=redecamara,DC=camara,DC=gov,DC=br"


#Inclui grupos nas contas que foram movidas para OU /Usuarios/Inativos/0/

$Contas = Get-ADUser -Filter * -SearchBase "OU=0,OU=Inativos,OU=Usuarios,DC=redecamara,DC=camara,DC=gov,DC=br" -Properties * 

##########Remove groups dos usuários na OU Especificada ########################

Get-ADUser -Filter * -SearchBase "OU=0,OU=Inativos,OU=Usuarios,DC=redecamara,DC=camara,DC=gov,DC=br" -Properties * | select -Expand MemberOf |%{Remove-ADGroupMember $_ -Member $Contas} 

################################################################################

Add-ADGroupMember -Identity ExchangePerfil4 -Members $Contas
Add-ADGroupMember -Identity FuncInativos -Members $Contas
Add-ADGroupMember -Identity Inativos -Members $Contas
Add-ADGroupMember -Identity Internet -Members $Contas


Remove-ADGroupMember -Identity internet -Members $Contas  #?Como remover todos os grupos de uma conta? 







# Mover os pontos para ou /Usuarios/Inativos/Funcionarios/

# Deixar somente os grupos: Domain Users, ExchangePerfil4, FuncInativos, Inativos, Internet


#>

