<#
.SYNOPSIS
    Sets the Out of Office (Automatic Reply) configuration for multiple specified mailboxes.

.DESCRIPTION
    This script connects to Exchange Online and Microsoft Graph. It prompts the user to either provide a CSV
    of users or specify an Azure AD Group. It then applies the specified Out of Office settings to each
    user in the chosen source.

.NOTES
    Author: CM V2.1
    Date:   2025-05-16
    Requires: ExchangeOnlineManagement, Microsoft.Graph.Users, Microsoft.Graph.Groups modules.
    When using a CSV, ensure it has a column header named 'UserPrincipalName'.
    Modify the OOF settings variables as needed.

.CHANGELOG
    V1.0 - Script created
    V1.1 - Added a check for the contents of $csvPath, which is outputted in the shell which called the script
    V1.2 - Modified variables for testing
    V1.3 - General cleanup of notes & variables from testing
    
    V2.0 - Added prompt for user source (CSV or Azure Group).
         - Implemented functionality to pull users from an Azure AD Group using Microsoft Graph.
         - Added Connect/Disconnect functions for MS Graph.
         - Updated required modules and comments.
         - $usersToProcess created to be used throughout the main processing loop, regardless of CSV/Group being used
    V2.1 - Added some further comments throughout to go with the new logic
#>

#Requires -Modules ExchangeOnlineManagement, Microsoft.Graph.Users, Microsoft.Graph.Groups

# --- OOF Settings ---
# Choose one of the following $oofState options:
# "Disabled" - Turns OOF off.
# "Enabled"  - Turns OOF on indefinitely (until manually disabled).
# "Scheduled"- Turns OOF on for a specific period.
$oofState = "Enabled" 

# If $oofState is "Scheduled", set StartTime and EndTime
# Ensure date format is understood by Get-Date or use specific culture parsing.
# Best to use ISO 8601 format (YYYY-MM-DDTHH:MM:SS) or let Get-Date parse friendly dates.
$oofStartTime = (Get-Date "2023-12-20 17:00:00") # Example: December 20, 2023, 5 PM
$oofEndTime   = (Get-Date "2024-01-05 09:00:00") # Example: January 5, 2024, 9 AM

# --- OOF Messages (HTML formatting is required (begin paragraphs with <p> and end them with </p>)) ---
$internalMessage = @"
<p>Thank you for your message. I am currently out of the office and will respond upon my return.</p>
"@ # Leaving this blank will allow only external replies, despite the GUI preventing blank entries.

$externalMessage = @"
<p>Thank you for your message. Our office is currently closed for the holidays. I will respond to your email upon my return.</p>
"@ 

# --- External Audience ---
# Acceptable options here are:
# "None"   - External auto-reply is not sent. (I.e you only want internal responses)
# "Known"  - External auto-reply is sent only to Senders in the Mailbox user's Contacts folder.
# "All"    - External auto-reply is sent to all external Senders.
$externalAudience = "All"

# --- Global Variables ---
$usersToProcess = @() # LIst of users to proccess which is used throughout the main loop
$sourceType = "" # This is defined by the user when running 

# --- Connection Functions ---

# Function to connect to Exchange Online
Function Connect-ToExchange {
    Write-Host "Attempting to connect to Exchange Online..."
    try {
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
    $session = Get-PSSession | Where-Object { $_.ConfigurationName -eq 'Microsoft.Exchange' }
    if ($session) {
        Write-Host "Disconnecting from Exchange Online..."
        Remove-PSSession $session
        Write-Host "Disconnected."
    }
}

# Function to connect to Microsoft Graph (Required to pull info from groups)
Function Connect-ToGraph {
    Write-Host "Attempting to connect to Microsoft Graph..."
    try {
        # Check if already connected with the required scopes
        $token = Get-MgContext
        if ($token -and ($token.Scopes -contains "GroupMember.Read.All") -and ($token.Scopes -contains "User.Read.All")) {
            Write-Host "Already connected to Microsoft Graph with required permissions."
        } else {
            # Scopes needed to read group members and user details
            $scopes = @("GroupMember.Read.All", "User.Read.All")
            Connect-MgGraph -Scopes $scopes
            Write-Host "Successfully connected to Microsoft Graph." -ForegroundColor Green
        }
    }
    catch {
        Write-Error "Failed to connect to Microsoft Graph. Please ensure the Microsoft.Graph modules are installed and you have the necessary permissions."
        Write-Error $_.Exception.Message
        exit 1
    }
}

# Function to disconnect from Microsoft Graph
Function Disconnect-FromGraph {
    if (Get-MgContext) {
        Write-Host "Disconnecting from Microsoft Graph..."
        Disconnect-MgGraph
        Write-Host "Disconnected."
    }
}


# --- Script Execution ---

# 1. Prompt user for the source of users
while ($true) {
    Write-Host "How would you like to specify the users?" -ForegroundColor Yellow
    Write-Host "[1] From a CSV file"
    Write-Host "[2] From an Azure AD Group"
    $choice = Read-Host "Please enter your choice (1 or 2)"

    if ($choice -eq '1') {
        $sourceType = "CSV"
        $csvPath = Read-Host "Please enter the full path to your CSV file"
        if (-not (Test-Path $csvPath)) {
            Write-Warning "File not found at '$csvPath'. Please try again."
            continue
        }
        $identityColumnName = "UserPrincipalName" # Standardize on UPN for simplicity
        $usersToProcess = Import-Csv -Path $csvPath
        Write-Host "Properties (column headers) found in the CSV:"
        $usersToProcess[0] | Get-Member -MemberType NoteProperty | Select-Object Name
        Write-Warning "Please ensure your CSV has a column named '$identityColumnName'."
        Read-Host "Press Enter to continue if the column name is correct..."
        break
    }
    elseif ($choice -eq '2') {
        $sourceType = "Azure Group"
        break
    }
    else {
        Write-Warning "Invalid choice. Please enter 1 or 2." # Simple error catch here to stop the script breaking if an input is recieved that it doesn't expect
    }
}


# --- Main Script Logic ---
try {
    # 2. Connect to required services
    Connect-ToExchange

    # 3. Get the list of users based on the chosen source
    if ($sourceType -eq "CSV") {
        if ($null -eq $usersToProcess -or $usersToProcess.Count -eq 0) {
            Write-Error "No users found in the CSV file or the file is empty." # Check the contents of the CSV and exit script if empty
            exit 1
        }
        Write-Host "Loaded $($usersToProcess.Count) users from CSV." -ForegroundColor Yellow
    }
    elseif ($sourceType -eq "Azure Group") {
        Connect-ToGraph
        $groupName = Read-Host "Please enter the Azure AD Group Display Name or Object ID"
        
        Write-Host "Searching for Azure AD Group: $groupName..."
        $group = Get-MgGroup -Filter "displayName eq '$groupName'" -ErrorAction SilentlyContinue
        if (-not $group) {
            # If not found by name, try by ID
            $group = Get-MgGroup -GroupId $groupName -ErrorAction SilentlyContinue
        }

        if (-not $group) {
            Write-Error "Azure AD Group '$groupName' not found. Please check the name/ID and permissions." # Group not found, exit to prevent running into errors later
            exit 1
        }

        Write-Host "Found group: $($group.DisplayName) ($($group.Id)). Fetching members..."
        # Get members and filter for only User objects. The -All parameter handles paging for large groups.
        $groupMembers = Get-MgGroupMember -GroupId $group.Id -All | Where-Object { $_.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.user' }

        if ($null -eq $groupMembers -or $groupMembers.Count -eq 0) {
            Write-Error "No user objects found in the group '$($group.DisplayName)'."
            exit 1
        }
        
        # The main loop expects an object with a 'UserPrincipalName' property, which Get-MgGroupMember provides.
        $usersToProcess = $groupMembers
        $identityColumnName = "UserPrincipalName"
        Write-Host "Found $($usersToProcess.Count) users in the Azure group." -ForegroundColor Yellow
    }

    # 4. Loop through each user and set OOF (This is the famous "Main Loop")
    Write-Host "Starting to process users..." -ForegroundColor Cyan
    foreach ($userRow in $usersToProcess) {
        $userIdentity = $userRow.$identityColumnName
        if ([string]::IsNullOrWhiteSpace($userIdentity)) {
            Write-Warning "Skipping row with empty identity."
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

            # Optional: Clear messages when disabling OOF. 
            # This one is important to be aware of, as bulk clearing lengthy OOF messages could cause some headaches.
            # If you have a copy of the message, you could use this script to re-instate the OOF.
            if ($oofState -eq "Disabled") {
                $params.InternalMessage = ""
                $params.ExternalMessage = ""
            }

            Set-MailboxAutoReplyConfiguration @params -ErrorAction Stop

            # Provide some live logging on the script execution back to the user
            Write-Host "Successfully set OOF for $userIdentity. State: $oofState" -ForegroundColor Green
            if ($oofState -eq "Scheduled") {
                Write-Host "  Scheduled from: $($oofStartTime.ToString('yyyy-MM-dd HH:mm')) to $($oofEndTime.ToString('yyyy-MM-dd HH:mm'))" -ForegroundColor Green
            }
        }
        catch {
            Write-Error "Failed to set OOF for $userIdentity. Error: $($_.Exception.Message)" # Restarts the loop after failure, meaning we don't exit just becuase one mailbox has an issue
        }
        Write-Host "---"
    }

    Write-Host "Script finished." -ForegroundColor Cyan
}
catch {
    Write-Error "An unexpected error occurred in the main script block: $($_.Exception.Message)"
}
finally {
    # 5. Disconnect from all services
    Disconnect-FromExchange # This always runs, as we always make the connection
    Disconnect-FromGraph # This will only run if a connection was made
}
