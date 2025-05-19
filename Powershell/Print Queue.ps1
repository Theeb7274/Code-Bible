$Printer = "Resources-Colour"


function ClearPrintQueue {
    Get-PrintJob -PrinterName $Printer
        ForEach-Object {
            Try {
                Remove-PrintJob
                }
            Catch {
                Write-Host "Error removing Job"
            }
        }
}

ClearPrintQueue