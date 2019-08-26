# This function maps printers from an array
function Map-Printers($Printers) {
  # Loop over the array
  foreach ($Printer in $Printers) {
    # Map the printer
    (New-Object -ComObject WScript.Network).AddWindowsPrinterConnection($Printer)
  }
}

$PathLista = "\\redecamara\dfsdata\Satus\Publico - Central de Atendimento\Softwares\computadoresGabinetes.csv"
$ArqLista = $null
$ArqLista = Import-Csv -Path $PathLista -Delimiter "," -Encoding Default
$ArqLista | ForEach-Object {
    $computador = $PSItem.name
    $gabinete = ([int] $PSItem.gabinete).ToString("000")
    if ($env:COMPUTERNAME -eq $computador) {
        #$env:COMPUTERNAME
        # Define a printer array
        $Printers = @("\\chillan\LP15-GAB-$gabinete", "\\elbrus\MF15-GAB-$gabinete")
        # Call our map printers function and pass in the printers array.
        Map-Printers -Printers $Printers
        Pause
        Exit
    }
}

