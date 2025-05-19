<#
.SYNOPSIS
    Sets the Out of Office (Automatic Reply) configuration for multiple specified mailboxes.

.DESCRIPTION
    This script connects to Exchange Online, reads a list of user mailboxes from a CSV file,
    and applies the specified Out of Office settings to each mailbox.

.NOTES
    Author: CM V1.2
    V1.1 - Added a check for the contents of $csvPath
    V1.2 - Modified $oofState to "Enabled"
    Date:   2025-05-15
    Requires: ExchangeOnlineManagement module.
    Ensure the CSV file path and user identity column name are correct.
    Modify the OOF settings variables as needed.
#>

#Requires -Modules ExchangeOnlineManagement

# --- Configuration ---
$csvPath = "C:\Users\Admin\Desktop\Shtuff\BHCSOOF.csv" # IMPORTANT: Update this path
$identityColumnName = "UserPrincipalName"   # Column header in CSV for user identities

# --- Check Contents of CSV ---
$usersToProcess = Import-Csv -Path $csvPath
 if ($null -eq $usersToProcess -or $usersToProcess.Count -eq 0) {
     Write-Warning "No users found in the CSV file or the file is empty."
     exit 1 # Used for debugging
 } else {
     Write-Host "First few rows from CSV:"
     $usersToProcess | Select-Object -First 5 | Format-Table -AutoSize # Display first 5 rows
     Write-Host "Properties (column headers) found in the CSV:"
     $usersToProcess[0] | Get-Member -MemberType NoteProperty | Select-Object Name
 }

# --- OOF Settings ---
# Choose one of the following $oofState options:
# "Disabled" - Turns OOF off.
# "Enabled"  - Turns OOF on indefinitely (until manually disabled).
# "Scheduled"- Turns OOF on for a specific period.
$oofState = "Enabled" # Or "Enabled" or "Disabled"

# If $oofState is "Scheduled", set StartTime and EndTime
# Ensure date format is understood by Get-Date or use specific culture parsing.
# Best to use ISO 8601 format (YYYY-MM-DDTHH:MM:SS) or let Get-Date parse friendly dates.
$oofStartTime = (Get-Date "2023-12-20 17:00:00") # Example: December 20, 2023, 5 PM
$oofEndTime   = (Get-Date "2024-01-05 09:00:00") # Example: January 5, 2024, 9 AM

# --- OOF Messages (HTML is allowed) ---
$internalMessage = @"
"@

$externalMessage = @"
<p>Thank you for contacting Brighton Hill Community School.</p>
<p>This is an automated response to acknowledge safe receipt of your email.</p>
<p>Except in extenuating circumstances, we endeavor to respond to emails within 3 school working days of receipt. This is because the matter may need consideration and because colleagues spend a high proportion of their working day teaching and supporting students.</p>
<p>If you don't receive a reply within 3 working days, it may be that a colleague is absent or urgent matters have prevented them from replying within the usual time frame; please re-send the email to the intended recipient and to the admin@bhcs.sfet.org.uk</p>
<p>email address. Please also check your 'junk mail' folder as some email accounts have been found to automatically divert school emails to this folder.</p>
<p>If your email relates to a safeguarding matter, please don’t hesitate to forward your email to safeguarding@bhcs.sfet.org.uk and we will deal with the concern as a matter of urgent priority.</p>
<p>If you are contacting us to raise a safeguarding concern and it is out of school hours or term time, please contact Hampshire County Council's safeguarding team on: 0300 555 1384 or the police on 101.</p>
<p>If a child is in immediate danger, please contact the police using 999.</p>
"@

# --- External Audience ---
# Acceptable options here are:
# "None"   - External auto-reply is not sent.
# "Known"  - External auto-reply is sent only to Senders in the Mailbox user's Contacts folder.
# "All"    - External auto-reply is sent to all external Senders.
$externalAudience = "All"

# --- Script Execution ---

# Function to connect to Exchange Online
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

# Function to disconnect from Exchange Online
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

            # If disabling, we can clear the messages, though it's not strictly necessary
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
    # 4. Disconnect from Exchange Online
    Disconnect-FromExchange
}