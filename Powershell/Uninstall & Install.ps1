# --- Version 1.0 CM ---

# Arguments for function
$SoftwareName = "" # Searches for the Display name key, not always the same as the application name
$OldVersion = "" # If found this version will be uninstalled
$MSIPath = "" # Path to the installer 
$InstallLog = "" # Path to log install progress
$UninstallLog = "" # Path to log uninstall progress



function Script {
    param(
        [Parameter(Mandatory=$true)] # Called display name in the registry
        [string]$SoftwareName,

        [Parameter(Mandatory=$true)] # Version which if found, will be uninstalled
        [string]$OldVersionString, 

        [Parameter(Mandatory=$true)]
        [ValidateScript({Test-Path $_ -PathType Leaf})] # Ensure the MSI file exists
        [string]$NewVersionMsiPath,

        [Parameter(Mandatory=$false)]
        [string]$UninstallLogFile,

        [Parameter(Mandatory=$false)]
        [string]$InstallLogFile
    )

    # --- Script Body ---
    Write-Host "--- Starting Software Upgrade Process ---"
    Write-Host "Target Software: '$SoftwareName'"
    Write-Host "Target Old Version: '$OldVersionString'"
    Write-Host "New Version MSI: '$NewVersionMsiPath'"

    # Check for Admin privileges
    if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-Error "This script requires Administrator privileges. Please re-run PowerShell as Administrator."
        exit 1
    }

    # Registry paths to search, use HKCU for user only installations
    $UninstallRegPaths = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
        'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
    )

    $foundApp = $null
    $productCodeToUninstall = $null

    # Search for the old version
    Write-Host "Searching for installed version $OldVersionString..."
    foreach ($regPath in $UninstallRegPaths) {
        Write-Verbose "Checking registry path: $regPath"
        $regKeys = Get-ChildItem -Path $regPath -ErrorAction SilentlyContinue
        if (-not $regKeys) { continue } # Skips if the path doesn't exist or is empty

        foreach ($key in $regKeys) {
            $appName = $key.GetValue('DisplayName') -as [string]
            $appVersion = $key.GetValue('DisplayVersion') -as [string]
            $uninstallString = $key.GetValue('UninstallString') -as [string] # Often contains Product Code for MSI

            if (($appName -ne $null) -and ($appVersion -ne $null)) {
                # Check if DisplayName matches the pattern AND DisplayVersion matches exactly
                if (($appName -like $SoftwareName) -and ($appVersion -eq $OldVersionString)) {
                    Write-Host "Found matching application:"
                    Write-Host "  Name: $appName"
                    Write-Host "  Version: $appVersion"
                    Write-Host "  Registry Key: $($key.Name)"

                    # Attempt to extract MSI Product Code (often the registry key name itself for MSI)
                    if ($key.PSChildName -match '^\{[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}\}$') {
                         $productCodeToUninstall = $key.PSChildName
                         Write-Host "  Detected Product Code: $productCodeToUninstall"
                    }
                    # Fallback: Try parsing UninstallString if it looks like msiexec /x {GUID}
                    elseif ($uninstallString -match 'msiexec.*(/x|/I|/uninstall)\s*(\{.*?\})') {
                        $productCodeToUninstall = $Matches[2]
                        Write-Host "  Detected Product Code from UninstallString: $productCodeToUninstall"
                    }
                     else {
                        Write-Warning "  Could not reliably determine MSI Product Code for uninstall from key name or UninstallString."
                        # You might need more specific logic here based on the app if the above fails
                     }

                    $foundApp = $key # Store the found key object
                    break # Stop searching this registry path
                }
            }
        }
        if ($foundApp) { break } # Stop searching other registry paths if found
    }

    # --- Uninstall Process ---
    $uninstallSuccessful = $false
    if ($foundApp -and $productCodeToUninstall) {
        Write-Host "Attempting to uninstall '$($foundApp.GetValue('DisplayName'))' (Version: $OldVersionString)..."

        # Array for MSI Arguments
        $msiArgsUninstall = @(
            "/x"                     
            $productCodeToUninstall  
            "/qn"                    
            "/norestart"             
        )

        if (Ensure-LogDirectoryExists($UninstallLogFile)) {
             if (-not [string]::IsNullOrWhiteSpace($UninstallLogFile)) {
                $msiArgsUninstall += "/L*v" 
                $msiArgsUninstall += "`"$UninstallLogFile`"" # Path to log file (quoted)
                Write-Host "Uninstall log: $UninstallLogFile"
             }
        } else {
            $UninstallLogFile = "" # Disable logging if directory creation failed
        }


        Write-Verbose "Running: msiexec.exe $($msiArgsUninstall -join ' ')"
        try {
            $process = Start-Process msiexec.exe -ArgumentList $msiArgsUninstall -Wait -PassThru -ErrorAction Stop

            $exitCode = $process.ExitCode
            Write-Host "Uninstall process completed with Exit Code: $exitCode"

            # Treat 0 (success) and 3010 (success, reboot required) as successful for proceeding
            if ($exitCode -eq 0 -or $exitCode -eq 3010) {
                Write-Host "Uninstall successful."
                $uninstallSuccessful = $true
                if ($exitCode -eq 3010) {
                    Write-Warning "A system restart is recommended to complete the uninstall."
                }
            } else {
                Write-Error "Uninstall failed (Exit Code: $exitCode). Check the uninstall log if enabled."
                # Optional: Decide whether to stop the script here or still attempt install
                # exit $exitCode # Uncomment to stop script on uninstall failure
            }
        } catch {
            Write-Error "Failed to start uninstall process. Error: $($_.Exception.Message)"
              exit 1
        }
    } elseif ($foundApp -and !$productCodeToUninstall) {
         Write-Warning "Found the application entry but could not determine the Product Code to uninstall it. Exiting script."
         exit 1
    }
    else {
        Write-Host "Older version ($OldVersionString) not found installed. Proceeding directly to installation."
        # Set uninstallSuccessful to true, thus enabling the installation. 
        $uninstallSuccessful = $true
    }

    # --- Installation Process ---
    if ($uninstallSuccessful) {
        Write-Host "Attempting to install new version from '$NewVersionMsiPath'..."

        # New array for the install arguments
        $msiArgsInstall = @(
            "/i"                   
            "`"$NewVersionMsiPath`"" 
            "/qn"                  
            "/norestart"           
        )

         if (Ensure-LogDirectoryExists($InstallLogFile)) {
            if (-not [string]::IsNullOrWhiteSpace($InstallLogFile)) {
                $msiArgsInstall += "/L*v" 
                $msiArgsInstall += "`"$InstallLogFile`"" 
                Write-Host "Install log: $InstallLogFile"
            }
        } else {
            $InstallLogFile = "" # Disable logging if directory creation failed
        }

        Write-Verbose "Running: msiexec.exe $($msiArgsInstall -join ' ')"
        try {
            $process = Start-Process msiexec.exe -ArgumentList $msiArgsInstall -Wait -PassThru -ErrorAction Stop

            $exitCode = $process.ExitCode
            Write-Host "Install process completed with Exit Code: $exitCode"

            if ($exitCode -eq 0) {
                Write-Host "Installation of new version successful."
            } elseif ($exitCode -eq 3010) {
                 Write-Host "Installation of new version successful, but a system restart is required."
            }
             else {
                Write-Error "Installation of new version failed (Exit Code: $exitCode). Check the install log."
            }
        } catch {
            Write-Error "Failed to start install process. Error: $($_.Exception.Message)"
        }
    } else {
        Write-Warning "Skipping installation of new version because the uninstall of the old version did not succeed or was skipped."
    }

    Write-Host "--- Software Upgrade Process Finished ---"
}

# Logging is optional, if not defined below the script will still run. All other arguments are required

Script -SoftwareName $SoftwareName -OldVersionString $OldVersion -NewVersionMsiPath $MSIPath -InstallLogFile $InstallLog -UninstallLogFile $UninstallLog