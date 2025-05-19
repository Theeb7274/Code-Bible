<#
.SYNOPSIS
    Sets the Out of Office (Automatic Reply) configuration for multiple specified mailboxes.

.DESCRIPTION
    This script connects to Exchange Online, reads a list of user mailboxes from a CSV file,
    and applies the specified Out of Office settings to each mailbox.

.NOTES
    Author: CM V1.2
    Date:   2025-05-15
    Requires: ExchangeOnlineManagement module.
    Ensure the CSV file path and user identity column name are correct.
    Modify the OOF settings variables as needed.

.CHANGELOG
    V1.1 - Added a check for the contents of $csvPath, which is outputted in the shell which called the script
    V1.2 - Modified varibles for testing
    V1.3 - General cleanup of notes & variables from testing
#>

#Requires -Modules ExchangeOnlineManagement

# --- Configuration ---
$csvPath = "" # Local csv of users to modify
$identityColumnName = "UserPrincipalName"   # Column header in CSV for user identities, ID can instead be used

# --- Check Contents of CSV ---
$usersToProcess = Import-Csv -Path $csvPath
 if ($null -eq $usersToProcess -or $usersToProcess.Count -eq 0) {
     Write-Warning "No users found in the CSV file or the file is empty."
     exit 1 # Came for the debugging, stayed because useful
 } else {
     Write-Host "First few rows from CSV:"
     $usersToProcess | Select-Object -First 5 | Format-Table -AutoSize # Display first 5 rows, you can extend this if you wish
     Write-Host "Properties (column headers) found in the CSV:"
     $usersToProcess[0] | Get-Member -MemberType NoteProperty | Select-Object Name
 }

# --- OOF Settings ---
# Choose one of the following $oofState options:
# "Disabled" - Turns OOF off.
# "Enabled"  - Turns OOF on indefinitely (until manually disabled).
# "Scheduled"- Turns OOF on for a specific period.
$oofState = "" 

# If $oofState is "Scheduled", set StartTime and EndTime
# Ensure date format is understood by Get-Date or use specific culture parsing.
# Best to use ISO 8601 format (YYYY-MM-DDTHH:MM:SS) or let Get-Date parse friendly dates.
$oofStartTime = (Get-Date "2023-12-20 17:00:00") # Example: December 20, 2023, 5 PM
$oofEndTime   = (Get-Date "2024-01-05 09:00:00") # Example: January 5, 2024, 9 AM

# --- OOF Messages (HTML is allowed) ---
$internalMessage = @"
"@ # Leaving this blank will allow only external replies, despite the GUI preventing this

$externalMessage = @"
<p>Example Message.</p>
"@ # <p> begins the paragraph, .</p> ends it

# --- External Audience ---
# Acceptable options here are:
# "None"   - External auto-reply is not sent.
# "Known"  - External auto-reply is sent only to Senders in the Mailbox user's Contacts folder.
# "All"    - External auto-reply is sent to all external Senders.
$externalAudience = ""

# --- Script Execution ---

# Function to connect to Exchange Online, I recommend using a regular Powershell instance, rather than ISE as the GUI to authenticate has issues in ISE
Function Connect-ToExchange {
    Write-Host "Attempting to connect to Exchange Online..."
    try {
        # Check if already connected
        $currentSession = Get-PSSession | Where-Object { $_.ConfigurationName -eq 'Microsoft.Exchange' }
        if ($currentSession.Count -gt 0) {
            Write-Host "Already connected to Exchange Online."
        } else {
            Connect-ExchangeOnline -ShowBanner:$false
            Write-Host "Successfully connected to Exchange Online." -ForegroundColor Green
        }
    }
    catch {
        Write-Error "Failed to connect to Exchange Online. Please ensure the ExchangeOnlineManagement module is installed and you have the necessary permissions."
        Write-Error $_.Exception.Message
        exit 1
    }
}

# Function to disconnect from Exchange Online, this is called regardless of where the script terminates to ensure disconnection
Function Disconnect-FromExchange {
    Write-Host "Disconnecting from Exchange Online..."
    $session = Get-PSSession | Where-Object { $_.ConfigurationName -eq 'Microsoft.Exchange' }
    if ($session) {
        Remove-PSSession $session
        Write-Host "Disconnected."
    } else {
        Write-Host "No active Exchange Online session found to disconnect."
    }
}

# --- Main Script Logic ---
try {
    # 1. Connect to Exchange Online
    Connect-ToExchange

    # 2. Import users from CSV
    if (-not (Test-Path $csvPath)) {
        Write-Error "CSV file not found at '$csvPath'. Please check the path."
        exit 1
    }
    $usersToProcess = Import-Csv -Path $csvPath
    if ($null -eq $usersToProcess -or $usersToProcess.Count -eq 0) {
        Write-Warning "No users found in the CSV file or the file is empty."
        exit 1
    }

    Write-Host "Processing $($usersToProcess.Count) users from CSV..." -ForegroundColor Yellow

    # 3. Loop through each user and set OOF
    foreach ($userRow in $usersToProcess) {
        $userIdentity = $userRow.$identityColumnName
        if ([string]::IsNullOrWhiteSpace($userIdentity)) {
            Write-Warning "Skipping row with empty identity in CSV."
            continue
        }

        Write-Host "Processing user: $userIdentity"

        try {
            $params = @{
                Identity         = $userIdentity
                AutoReplyState   = $oofState
                InternalMessage  = $internalMessage
                ExternalMessage  = $externalMessage
                ExternalAudience = $externalAudience
            }

            if ($oofState -eq "Scheduled") {
                $params.Add("StartTime", $oofStartTime)
                $params.Add("EndTime", $oofEndTime)
            }

            # Optinal removal of message when disabling, this isn't required and can be changed if desired
            if ($oofState -eq "Disabled") {
                $params.InternalMessage = ""
                $params.ExternalMessage = ""
            }

            Set-MailboxAutoReplyConfiguration @params

            Write-Host "Successfully set OOF for $userIdentity. State: $oofState" -ForegroundColor Green
            if ($oofState -eq "Scheduled") {
                Write-Host "  Scheduled from: $($oofStartTime.ToString('yyyy-MM-dd HH:mm')) to $($oofEndTime.ToString('yyyy-MM-dd HH:mm'))" -ForegroundColor Green
            }
        }
        catch {
            Write-Error "Failed to set OOF for $userIdentity. Error: $($_.Exception.Message)"
            # Optionally, you could log this to a file
        }
        Write-Host "---"
    }

    Write-Host "Script finished." -ForegroundColor Cyan
}
catch {
    Write-Error "An unexpected error occurred: $($_.Exception.Message)"
}
finally {
    # 4. Disconnect from Exchange Online, due to finally being used, this is always ran regardless of how/when the script terminated 
    Disconnect-FromExchange
}
