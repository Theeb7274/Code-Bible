# Variables
$LocalDestinationPath = "" # Installation folder on endpoints
$ExecutablePath = "" # Network location of app to install
$LogFilePath = "" # Specify the log file path on endpoints

# Function to create $LocalDestinationPath directory if it doesn't exist
function Create-Path {
    if (-not (Test-Path -Path $LocalDestinationPath)) {
        try {
            New-Item -ItemType Directory -Path $LocalDestinationPath | Out-Null        
        }
        catch {
            $ErrorMessage = "Failed to create local directory: $($_.Exception.Message)"
            return
            Write-Log $ErrorMessage
        }
    }
}
# Function to write to the log file
function Write-Log {
    param (
        [string]$Message
    )
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogEntry = "$Timestamp - $Message"
    Add-Content -Path $LogFilePath -Value $LogEntry
}

function CheckRegKeyExists {
    param (
            [parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]$Path,

            [parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]$Value
          )

    try {
        Get-ItemProperty -Path $Path | Select-Object -ExpandProperty $Value -ErrorAction Stop | Out-Null
        #Write-Host "Key Exists"
        return $true         
    }
     catch 
     {
        #Write-Host "Key Does not Exist"
        return $false
    }                        
}
# Optional function for cleaning up installation directory ($LocalDestinationPath)
function Remove-File {
    param ([String]$File)
    #Remove-File -FileToRemove "$LocalDestinationPath"

    # Delete local zip file
    try {
            Remove-Item -LiteralPath $File -Force
            Write-Log "Local zip '$File' deleted successfully or not present."
        }
    catch {
            $WarningMessage = "Failed to delete local zip file: $($_.Exception.Message)"
            Write-Log $WarningMessage
        }
}

# Run the executable
function Install-App {
    try {
        if 
        Start-Process msiexec -ArgumentList "/i `"$ExecutablePath`"","/qn""/li $LogFilePath"  -Wait;
        Write-Log "Executable ran successfully."
    }
    catch {
        $ErrorMessage = "Failed to run msi, refer to $LogFilePath for details"
        Write-Log $ErrorMessage
        return;
    }
}

Write-Log "######### Starting Script #########"
Create-Path
Install-App
Write-Log "######### Script Completed #########"
Write-Log "######### Closing Script #########"
