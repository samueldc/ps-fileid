# This function maps printers from an array
function Map-Printers($Printers) {
  # Loop over the array
  foreach ($Printer in $Printers) {
    # Map the printer
    (New-Object -ComObject WScript.Network).AddWindowsPrinterConnection($Printer)
  }
}

# Define a printer array
$Printers = @("\\lanin\ditec027", "\\lanin\ditec001")

# Call our map printers function and pass in the printers array.
Map-Printers -Printers $Printers
