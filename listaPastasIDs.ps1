Remove-Module PSFileId
Import-Module  C:\desenv\scripts-powershell-sepac\PSFileId\PSFileId
$identifier = Get-ItemId $(Get-Item "C:\Users\p_7029\Nextcloud\_trabalho\copad\otrs-sustentacao-projetos\termoReferenciaSigmasRevisado.pdf")
$identifier.FileId.Identifier.FixedElementField