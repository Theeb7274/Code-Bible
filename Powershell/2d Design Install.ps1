# Variables
$LocalDestinationPath = "C:\EDUIT" # Installation folder on endpoints
$ExecutablePath = "\\sfet-bhcs-mgt01\software$\TechSoft2DDesign v2.17" # Network location of app to install
$LogFilePath = "C:\EDUIT\2DD_script_log.txt" # Specify the log file path on endpoints

# Function to create the local EDUIT destination directory if it doesn't exist
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

function Remove-FolderRecursive {
    param ([string]$RecursiveFolderPath)

    #Remove-FolderRecursive -RecursiveFolderPath "C:\EDUIT\ExtractedFiles"

    try {
        if (Test-Path -Path $RecursiveFolderPath -PathType Container) {
            Remove-Item -Path $RecursiveFolderPath -Recurse -Force -ErrorAction Stop
            Write-Log "Folder '$RecursiveFolderPath' and its contents deleted successfully."
        } else {
            Write-Log "Folder '$RecursiveFolderPath' does not exist."
        }
    } catch {
        Write-Error "Error deleting folder '$FolderPath': $($_.Exception.Message)"
    }
}

function Remove-File {
    param ([String]$File)
    #Remove-File -FileToRemove "C:\EDUIT"

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