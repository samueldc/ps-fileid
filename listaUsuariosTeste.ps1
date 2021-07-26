Set-ExecutionPolicy RemoteSigned
Import-Module ImportExcel
$usuarios = Get-ADUser -Filter "Name -like '*test*'" -Properties *
$usuarios | ForEach-Object {
    If ($_.MemberOf -like "*Internet*") {
        Add-Member -InputObject $_ -NotePropertyName Internet -NotePropertyValue $true -Force
    } Else {
        Add-Member -InputObject $_ -NotePropertyName Internet -NotePropertyValue $false -Force
    }
}
$usuarios | ForEach-Object {
    If ($_.MemberOf -like "*Negar_Logon_local_RDP*") {
        Add-Member -InputObject $_ -NotePropertyName NegarLogon -NotePropertyValue $true -Force
    } Else {
        Add-Member -InputObject $_ -NotePropertyName NegarLogon -NotePropertyValue $false -Force
    }
}
$usuarios | Select Name,DistinguishedName,Department,Description,EmailAddress,Enabled,lastLogoff,lastLogon,LastLogonDate,mail,Organization,Title,whenCreated,whenChanged,Internet,NegarLogon |
Export-Csv -Path "c:/temp/usuariosTeste.csv" -Delimiter ";" -Encoding UTF8