<#
    Printer & Driver Install V1.0

    Overview:
    
        - Installs a printer driver and adds a printer using it.
        - I settled on using DISM due to environmental restrictions surrounding pnputil
        - DISM also allows unsigned drivers to be installed where they otherwise may be restricted

    Notes

        - Currently, -NoNewWindows does nothing, as DISM will in fact create a new window
        - There are some basic if -not checks to avoid creating duplicates
        - Query driver name step added as the name can be inconsistent from the .inf file

#>

# Variables
$PrinterName   = ""
$PrinterIP     = ""
$DriverINFPath = ""
$DriverSearch  = "" # Used to find the driver, use your .inf as reference

# Install the driver package into the driver store 
Start-Process -FilePath "dism.exe" -ArgumentList "/online /add-driver /driver:`"$DriverINFPath`" /forceunsigned" -Wait -NoNewWindow

# Query driver name 
$DriverName = (Get-PrinterDriver | Where-Object { $_.Name -match "" } | Sort-Object Name -Descending | Select-Object -First 1).Name

# Step 3: Create TCP/IP port if missing
if (-not (Get-PrinterPort -Name "IP_$PrinterIP" -ErrorAction SilentlyContinue)) {
    Add-PrinterPort -Name "IP_$PrinterIP" -PrinterHostAddress $PrinterIP -PortNumber 9100
}

# Step 4: Add printer if missing
if (-not (Get-Printer -Name $PrinterName -ErrorAction SilentlyContinue)) {
    Add-Printer -Name $PrinterName -DriverName $DriverName -PortName "IP_$PrinterIP"
}

Write-Host "Printer $PrinterName installed using driver: $DriverName"
