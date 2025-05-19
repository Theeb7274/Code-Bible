# Configuration
$NetworkZipPath = "\\bws-sr-001\EDUSoftwareDeploy$\EDU-CCS-[ID-0331467]_en_GB_WIN_64.zip" # Replace with your network path
$LocalDestinationPath = "C:\EDUIT" # Replace with your desired local path
$LocalZipPath = "C:\EDUIT\EDU-CCS-[ID-0331467]_en_GB_WIN_64.zip"
$ExtractedFolderPath = "$LocalDestinationPath\ExtractedFiles" # Folder where zip will be extracted
$ExecutablePath = "$ExtractedFolderPath\build\setup.exe" # Replace with the actual path to your executable
$LogFilePath = "C:\EDUIT\CCS_script_log.txt" # Specify the log file path

# Define the registry pa and value name
$RegistryPath = "HKLM:\Software\eduthing"
$ValueName1 = "CreativeCloudSuiteInstalled"
$ValueName2 = "CreativeCloudSuiteExtracted"

#Computer\HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{AC76BA86-1033-FFFF-7760-BC15014EA700}

# Create the local EDUIT destination directory if it doesn't exist
if (-not (Test-Path -Path $LocalDestinationPath)) {
    try {
        New-Item -ItemType Directory -Path $LocalDestinationPath | Out-Null        
    }
    catch {
        $ErrorMessage = "Failed to create local directory: $($_.Exception.Message)"
        return
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

function CheckRegKeyExists
{
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

Write-Log "######### Starting Script #########"

# Check if the eduthing registry key exists
if (!(Test-Path -Path $RegistryPath)) {
    try {
        # Create the registry key
        New-Item -Path $RegistryPath -Force | Out-Null
        Write-Log "Registry key '$RegistryPath' created successfully."
    } catch {
        Write-Log "Error creating registry key: $($_.Exception.Message)"
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

#################################
if (Get-ItemProperty -Path $RegistryPath | Select-Object -ExpandProperty $ValueName1 -ErrorAction SilentlyContinue)
    {
        Write-Log "$ValueName1 Exists, exiting"
        Write-Log "######### Closing Script #########"
        exit
    } 


if (Get-ItemProperty -Path $RegistryPath | Select-Object -ExpandProperty $ValueName2 -ErrorAction SilentlyContinue) 
    { 
        Write-Log "$ValueName2 Exists, skipping copy and extract" 
    } 
else     
    { 
        Write-Log "Cleaning up any previous failed attempt as copying and extracting"

        Remove-FolderRecursive -RecursiveFolderPath $ExtractedFolderPath
        Remove-File -File $LocalZipPath
    
        # Copy the zip file using robocopy
        try {
                $NetworkSourceDirectory = Split-Path -Path $NetworkZipPath
                $ZipFileName = Split-Path -Path $NetworkZipPath -Leaf

                Write-Log "Starting robocopy of Zip file"

                robocopy "$NetworkSourceDirectory" "$LocalDestinationPath" "$ZipFileName" /COPY:DAT /R:3 /W:5 /NP /NFL /NDL

                if ($LASTEXITCODE -ge 8) 
            {
                $ErrorMessage = "Robocopy failed with exit code: $LASTEXITCODE"
                throw $ErrorMessage
            }

            Write-Log "Zip file copied successfully using robocopy."
        }
        catch {
            $ErrorMessage = "Failed to copy zip file using robocopy: $($_.Exception.Message)"
            Write-Log $ErrorMessage
            return
        }

        $LocalZipFilePath = Join-Path -Path $LocalDestinationPath -ChildPath (Split-Path -Path $NetworkZipPath -Leaf)

        # Extract the zip file
        try {
            if (-not (Test-Path -Path $ExtractedFolderPath)) {
                New-Item -ItemType Directory -Path $ExtractedFolderPath | Out-Null
                Write-Log "Created directory: $ExtractedFolderPath"
            }
            Write-Log "Starting zip file extraction."
            Expand-Archive -LiteralPath $LocalZipFilePath -DestinationPath $ExtractedFolderPath -Force
            Write-Log "Zip file extracted successfully."

            New-ItemProperty -Path $RegistryPath -Name $ValueName2 -Value "1" -PropertyType String -Force | Out-Null 
            Write-Log "Registry value '$ValueName2' created successfully."

        }
        catch {
            $ErrorMessage = "Failed to extract zip file: $($_.Exception.Message)"
            Write-Log $ErrorMessage
            return
        }
    }

# Run the executable
try {
    #Start-Process -FilePath $ExecutablePath -ArgumentList "--silent" -Wait
    Write-Log "Executable ran successfully."

    # If the registry value does not exist, create it
    try {
        # Create the registry key if it doesn't exist
        if (!(Test-Path -Path $RegistryPath)) {
            New-Item -Path $RegistryPath -Force | Out-Null
        }

        # Create the registry value
        New-ItemProperty -Path $RegistryPath -Name $ValueName1 -Value "1" -PropertyType String -Force | Out-Null 

        Write-Log "Registry value '$ValueName1' created successfully."

        Remove-File -FileToRemove "C:\EDUIT\EDU-CCS-[ID-0331467]_en_GB_WIN_64.zip"
        Remove-FolderRecursive -RecursiveFolderPath "C:\EDUIT\ExtractedFiles"
        Remove-ItemProperty -Path $RegistryPath -Name $ValueName2

    } catch {
        Write-Error "Error creating registry value: $($_.Exception.Message)"
        Write-Log "Error creating '$ValueName1' the script will attempt to reinstall next startup"
    }
}
catch {
    $ErrorMessage = "Failed to run executable: $($_.Exception.Message)"
    Write-Log $ErrorMessage
    return;
}

Write-Host "######### Script Completed #########"
Write-Log "######### Closing Script #########"