function ClearPrintQueue {
    param(string)[Mandatory=$true()]$Printer
    
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