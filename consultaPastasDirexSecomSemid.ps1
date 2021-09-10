$path = "\\redecamara\DfsData"
$pastas = Get-ChildItem -Path $path -Attributes Directory
$pastas | Where-Object Name -Match "secom|semid|direx|comunicacao|imprensa|midia" | Format-Table Name