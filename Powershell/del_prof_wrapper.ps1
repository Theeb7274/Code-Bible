<#
.SYNOPSIS
    Runs delprof2.exe to delete user profiles locally or remotely, optionally logging output.

.DESCRIPTION
    This script is a wrapper around Helge Klein's delprof2.exe utility.
    It provides a PowerShell-friendly interface to delprof2.exe's common options.
    The script requires delprof2.exe to be in the same directory, in the system PATH,
    or specified via the -DelprofPath parameter.
    Requires Administrator privileges to run.
    Output from delprof2.exe can be redirected to a specified log file.

.PARAMETER ComputerName
    Specifies the computer on which to delete profiles. Defaults to the local computer.
    Maps to delprof2.exe /c:<ComputerName>

.PARAMETER DaysInactive
    Delete profiles inactive for more than this number of days.
    Maps to delprof2.exe /d:<Days>

.PARAMETER ListOnly
    Only list profiles that would be deleted. No changes are made.
    Maps to delprof2.exe /l

.PARAMETER Unattended
    Runs delprof2.exe in unattended mode (no prompts).
    Maps to delprof2.exe /u

.PARAMETER PromptBeforeDeleting
    Prompts before deleting each profile (delprof2.exe's own prompt).
    Maps to delprof2.exe /p
    This is mutually exclusive with -Unattended. If both are specified, -Unattended takes precedence.

.PARAMETER IgnoreErrors
    Ignores errors and continues processing.
    Maps to delprof2.exe /i

.PARAMETER DeleteRoaming
    Delete cached copies of roaming profiles.
    Maps to delprof2.exe /r

.PARAMETER DeleteNtUserIni
    Delete ntuser.ini for roaming profiles without deleting the profile. Requires -DeleteRoaming.
    Maps to delprof2.exe /ntuserini

.PARAMETER ExcludeUsers
    An array of user names (e.g., "DOMAIN\user", "user") to exclude from deletion.
    Maps to multiple /ed:<username> parameters.

.PARAMETER IncludeUsers
    An array of user names (e.g., "DOMAIN\user", "user") to explicitly include for deletion (overrides inactivity).
    Maps to multiple /id:<username> parameters.

.PARAMETER DelprofPath
    Specifies the full path to delprof2.exe if it's not in the script's directory or system PATH.

.PARAMETER LogFile
    Specifies the full path to a file where the standard output and standard error from
    delprof2.exe will be redirected. The file will be overwritten if it exists.
    If omitted, output goes to the console as usual.

.PARAMETER Force
    Suppresses the script's own confirmation prompt before running delprof2 (unless -ListOnly or -Unattended is used).
    Delprof2.exe might still prompt if -Unattended or -PromptBeforeDeleting is not used appropriately.

.NOTES
    By Cameron Meylan
    Version: 1.2
    Requires: delprof2.exe from Helge Klein (https://helgeklein.com/free-tools/delprof2-user-profile-deletion-tool/)
              Administrator privileges.
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [string]$ComputerName,

    [ValidateRange(0, 36500)] # Up to 100 years
    [int]$DaysInactive,

    [switch]$ListOnly,

    [switch]$Unattended,

    [switch]$PromptBeforeDeleting,

    [switch]$IgnoreErrors,

    [switch]$DeleteRoaming,

    [switch]$DeleteNtUserIni,

    [string[]]$ExcludeUsers,

    [string[]]$IncludeUsers,

    [string]$DelprofPath,

    [ValidateScript({
        if ([string]::IsNullOrWhiteSpace($_)) {
            throw "LogFile path cannot be empty or whitespace."
        }
        return $true
    })]
    [string]$LogFile,

    [switch]$Force
)

begin {
    # Check for Administrator Privileges
    if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-Error "This script must be run as Administrator."
        exit 1
    }

    # Find delprof2.exe
    $delprofExeName = "delprof2.exe"
    $resolvedDelprofPath = $null

    if ($DelprofPath) {
        if (Test-Path $DelprofPath -PathType Leaf) {
            $resolvedDelprofPath = $DelprofPath
        } else {
            Write-Error "Specified DelprofPath '$DelprofPath' not found or is not a file."
            exit 1
        }
    } else {
        # Check script's directory
        $scriptDirDelprof = Join-Path $PSScriptRoot $delprofExeName
        if (Test-Path $scriptDirDelprof -PathType Leaf) {
            $resolvedDelprofPath = $scriptDirDelprof
        } else {
            # Check PATH
            $resolvedDelprofPath = (Get-Command $delprofExeName -ErrorAction SilentlyContinue).Source
            if (-not $resolvedDelprofPath) {
                Write-Error "$delprofExeName not found in script directory ('$PSScriptRoot') or system PATH. Use -DelprofPath to specify its location."
                exit 1
            }
        }
    }
    Write-Verbose "Using delprof2.exe at: $resolvedDelprofPath"

    # Build arguments for delprof2.exe
    $arguments = @()

    if ($PSBoundParameters.ContainsKey('ComputerName')) {
        $arguments += "/c:$ComputerName"
    }
    if ($PSBoundParameters.ContainsKey('DaysInactive')) {
        $arguments += "/d:$DaysInactive"
    }
    if ($ListOnly) {
        $arguments += "/l"
    }
    if ($Unattended) {
        $arguments += "/u"
    } elseif ($PromptBeforeDeleting) { # Only add /p if /u is not specified
        $arguments += "/p"
    }
    if ($IgnoreErrors) {
        $arguments += "/i"
    }
    if ($DeleteRoaming) {
        $arguments += "/r"
        if ($DeleteNtUserIni) {
            $arguments += "/ntuserini"
        }
    } elseif ($DeleteNtUserIni) {
        Write-Warning "-DeleteNtUserIni requires -DeleteRoaming to be specified. Ignoring -DeleteNtUserIni."
    }


    if ($ExcludeUsers) {
        foreach ($user in $ExcludeUsers) {
            $arguments += "/ed:$user"
        }
    }
    if ($IncludeUsers) {
        foreach ($user in $IncludeUsers) {
            $arguments += "/id:$user"
        }
    }

    # Prepare Log File Path if specified
    $logFilePathResolved = $null
    if ($PSBoundParameters.ContainsKey('LogFile')) {
        try {
            # Resolve the path (handles relative paths)
            $logFilePathResolved = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($LogFile)
            Write-Verbose "Attempting to log output to: $logFilePathResolved"

            # Ensure the directory exists
            $logDir = Split-Path -Path $logFilePathResolved -Parent
            if ($logDir -and (-not (Test-Path -Path $logDir -PathType Container))) {
                Write-Verbose "Creating log directory: $logDir"
                $null = New-Item -ItemType Directory -Path $logDir -Force -ErrorAction Stop
            }
        } catch {
            Write-Error "Error preparing log file path '$LogFile': $($_.Exception.Message)"
            # Exit if we can't set up logging as requested
            exit 1
        }
    }

}

process {
    $targetComputer = if ($ComputerName) { $ComputerName } else { $env:COMPUTERNAME }
    $actionMessage = "Run delprof2.exe on '$targetComputer'"
    if ($ListOnly) {
        $actionMessage += " (ListOnly)"
    }
    if ($logFilePathResolved) {
         $actionMessage += " and log output to '$logFilePathResolved'"
    } else {
         $actionMessage += " and display output"
    }

    # Build the argument string for display and potential direct execution
    $argumentStringForDisplay = $arguments -join ' '
    # Escape arguments properly for direct execution if they contain spaces
    $argumentsForDirectExec = $arguments | ForEach-Object {
        if ($_ -match '\s') { "'$_'" } else { $_ }
    }


    if ($ListOnly -or $Unattended -or ($PromptBeforeDeleting -and -not $Unattended) -or $Force -or $PSCmdlet.ShouldProcess($targetComputer, $actionMessage)) {
        Write-Verbose "Command line: `"$resolvedDelprofPath`" $argumentStringForDisplay" # Display the arguments as delprof2 expects them
        try {
            $exitCode = -1 # Initialize exit code

            if ($logFilePathResolved) {
                # Use PowerShell's redirection for external commands when logging
                Write-Host "Executing delprof2.exe and redirecting output to '$logFilePathResolved'..."
                # & : Call operator
                # $resolvedDelprofPath : Path to the executable
                # $arguments : Array of arguments passed directly
                # *> : Redirects ALL output streams (stdout, error, warning, verbose, etc. from the *external command's perspective*, mainly stdout and stderr)
                # This is generally safer than '> file 2>&1' for capturing everything an external console app might output.
                # Using Out-File ensures consistent encoding (default UTF8 without BOM in PS 7+, UTF16-LE in PS 5.1) - specify if needed.
                # Note: For PS 5.1, consider -Encoding Default or ASCII if delprof2 output is expected to be non-Unicode

                # PowerShell 7+ prefer *>, PS 5.1 might behave better with > file 2>&1
                if ($PSVersionTable.PSVersion.Major -ge 6) {
                     # PS Core/7+ : *> is generally preferred for merging streams to file
                     & $resolvedDelprofPath $arguments *> $logFilePathResolved
                } else {
                     # PS 5.1: > file 2>&1 is the classic way
                     & $resolvedDelprofPath $arguments > $logFilePathResolved 2>&1
                }

                $exitCode = $LASTEXITCODE # Capture exit code *immediately* after execution

                # Verify log file creation
                if (Test-Path $logFilePathResolved) {
                     Write-Host "Log file operation completed. Log: $logFilePathResolved"
                } else {
                     # This might happen if the command failed very early or permissions are wrong
                     Write-Warning "Log file '$logFilePathResolved' was specified but not found after execution attempt. Check permissions or delprof2 execution errors."
                }

            } else {
                # Execute directly to console (no redirection requested)
                # Start-Process is still good here if we want -NoNewWindow behavior explicitly
                # Or just call directly & and let output go to host console. Let's keep Start-Process for consistency.
                Write-Host "Executing delprof2.exe and displaying output..."
                 $startProcessParams = @{
                    FilePath     = $resolvedDelprofPath
                    ArgumentList = $arguments # Pass as array
                    Wait         = $true
                    PassThru     = $true
                    ErrorAction  = 'Stop'
                    NoNewWindow  = $true
                 }
                $process = Start-Process @startProcessParams
                $exitCode = $process.ExitCode
            }

            # --- Check Exit Code ---
            if ($exitCode -ne 0) {
                # Still write warnings to host even if logging, as it indicates non-standard completion
                Write-Warning "delprof2.exe exited with code $exitCode. Check output for details."
                if ($logFilePathResolved){
                     Write-Warning "Details should be in the log file: $logFilePathResolved"
                }

                # Specific exit code messages (can still be useful on console)
                if ($exitCode -eq 1) {
                    Write-Host "Delprof2 Status: No profiles found that match the criteria."
                } elseif ($exitCode -eq 2) {
                    Write-Host "Delprof2 Status: User cancelled operation."
                }
                 # Optionally, make the script exit code reflect delprof2's failure
                 # exit $exitCode
            } else {
                Write-Verbose "delprof2.exe completed successfully (Exit Code 0)."
                 if ($logFilePathResolved){
                     Write-Host "delprof2.exe completed successfully. Output logged to '$logFilePathResolved'."
                 } else {
                      Write-Host "delprof2.exe completed successfully." # Added message for non-logging case
                 }
            }

        } catch {
             # This catch block now primarily catches errors from Start-Process (if used) or path resolution/setup errors.
             # Errors *during* the direct execution (&) might not be caught here unless they are PowerShell-level errors.
             # The exit code check handles delprof2's own errors.
            Write-Error "Error during script execution or launching delprof2.exe: $($_.Exception.Message)"
            # If logging was intended, mention it might be incomplete
             if ($logFilePathResolved) {
                 Write-Warning "Logging to '$logFilePathResolved' may be incomplete due to the error."
             }
             # exit 1 # Optionally exit script on error
        }
    } else {
        Write-Host "Operation cancelled by user (script's confirmation prompt)."
    }
}

end {
    Write-Verbose "Script finished."
}