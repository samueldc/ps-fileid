## SYNOPSIS
     PowerShell module that uses Windows API GetFileInformationByHandleEx
     function to get a file ou folder filesystem id. Useful to asset files
     and folders and detect name changes, for example.
## SINTAXE
     Get-ItemId <System.IO.FileSystemInfo>
## DESCRIPTION
     Compile the Windows API and exposes by a PowerShell function called Get-ItemId.
## USAGE
     Install module using Install-Module -Name PSFileId
     Call Get-ItemId function
     Example: Get-ItemId $(Get-Item "C:\Test")
