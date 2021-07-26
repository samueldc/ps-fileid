Set-ExecutionPolicy RemoteSigned
Import-Module ImportExcel
$gabinetes = Import-Excel -Path "c:/temp/criaPastasBackupGabinetes.xlsx" -WorkSheetname gabinetes
$pathRoot = "\\redecamara\DfsData\Mig_gabinete01"
$gabinetes | ForEach-Object {
    if ( $PSItem -and $PSItem.carteira ) {
        $PSItem
        $Path = "$pathRoot\$($PSItem.anexo)\$($PSItem.andar)\TempBackupDep-56$($PSItem.carteira)"
        $Path
        if ($PSItem.acao -eq "C") { # Create directory and add permissions
            New-Item -ItemType "directory" -Path $Path
            $ACL = Get-ACL -Path $Path
            $AccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule("Dep-$($PSItem.carteira)-F", "Modify", "ContainerInherit, ObjectInherit", "None", "Allow")
            $ACL.AddAccessRule($AccessRule)
            $ACL | Set-Acl -Path $Path
        } elseif ($PSItem.acao -eq "E") { # Remove permissions
            # New-Item -ItemType "directory" -Path $Path
            $ACL = Get-ACL -Path $Path
            $AccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule("Dep-$($PSItem.carteira)-F", "Modify", "ContainerInherit, ObjectInherit", "None", "Deny")
            $ACL.AddAccessRule($AccessRule)
            $ACL | Set-Acl -Path $Path
        } elseif ($PSItem.acao -eq "D") { # Remove directory
            Remove-Item -Path $Path
            # $ACL = Get-ACL -Path $Path
            # $AccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule("Dep-$($PSItem.carteira)-F", "Modify", "ContainerInherit, ObjectInherit", "None", "Allow")
            # $ACL.AddAccessRule($AccessRule)
            # $ACL | Set-Acl -Path $Path
        } else {
            # Do nothing
        }
    }
}
