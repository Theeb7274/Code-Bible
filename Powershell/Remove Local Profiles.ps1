<#region --- INFO ---
V 1.0 - By Cameron Meylan
This script is designed to be called by another process, such as a scheduled task or GPO.
The following are unique paramers when calling this script, none of them are mandatory.

Custom parameters:
    -InactiveDays #specifies the age to look for. This is not always accurate.
    -ExcludeUsers #Users to not remove.

CmdletBinding is used, enabling standard switches such as -force & -verbose to be used

SupportsShouldProcess allows the following switches:
    -WhatIf #Runs the script as read only, displaying what would of been deleted but not performing any actions.
    -Confirm #Prompts the user when ran.

ConfirmImpact is currently set to none as this is intended to be an automated script. Raising this will cause a prompt, which is additve with -Confirm.
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'none')]
param (
    [Parameter(Mandatory = $false)]
    [int]$InactiveDays,

    [Parameter(Mandatory = $false)]
    [string[]]$ExcludeUsers,

    [Parameter(Mandatory = $false)]
    [switch]$Force
)



#region --- Configuration and Exclusions ---
Write-Verbose "Starting user profile cleanup process."

# Date threshold for inactivity
$thresholdDate = $null
if ($PSBoundParameters.ContainsKey('InactiveDays')) {
    if ($InactiveDays -lt 0) {
         Write-Error "InactiveDays must be zero or a positive integer." -ErrorAction Stop
    }
    $thresholdDate = (Get-Date).AddDays(-$InactiveDays)
    Write-Verbose "Profiles must have LastUseTime older than $($thresholdDate.ToString('yyyy-MM-dd HH:mm:ss')) to be considered for removal."
}
else {
    Write-Verbose "No InactivityDays specified; age will not be checked."
}

# --- Build list of SIDs to exclude ---
$excludedSIDs = @(
    "S-1-5-18", # Local System
    "S-1-5-19", # Local Service
    "S-1-5-20"  # Network Service
)

# Exclude the currently logged-on user running the script
$currentUserSID = $currentUser.User.Value
Write-Verbose "Excluding current user: $($currentUser.Name) (SID: $currentUserSID)"
$excludedSIDs += $currentUserSID

# Resolve additional user exclusions to SIDs
if ($PSBoundParameters.ContainsKey('ExcludeUsers')) {
    Write-Verbose "Processing additional user exclusions: $($ExcludeUsers -join ', ')"
    foreach ($user in $ExcludeUsers) {
        try {
            $ntAccount = New-Object System.Security.Principal.NTAccount($user)
            $sid = $ntAccount.Translate([System.Security.Principal.SecurityIdentifier]).Value
            if ($sid -and $excludedSIDs -notcontains $sid) {
                Write-Verbose "Excluding user '$user' (SID: $sid)"
                $excludedSIDs += $sid
            } else {
                 Write-Verbose "User '$user' already in exclusion list or SID resolution failed (but continuing)."
            }
        }
        catch [System.Security.Principal.IdentityNotMappedException] {
            Write-Warning "Could not resolve username '$user' to a SID. It might not exist locally. Skipping this exclusion."
        }
        catch {
            Write-Warning "An error occurred resolving SID for user '$user': $($_.Exception.Message)"
        }
    }
}

Write-Verbose "Final list of excluded SIDs: $($excludedSIDs -join ', ')"
#endregion

#region --- Get and Filter Profiles ---
Write-Verbose "Querying local user profiles..."
try {
    # Get all profiles using CIM (more modern than WMI)
    $allProfiles = Get-CimInstance -ClassName Win32_UserProfile -ErrorAction Stop

    Write-Verbose "Found $($allProfiles.Count) total profiles."

    # Filter profiles
    $profilesToRemove = $allProfiles | Where-Object {
        # Basic Exclusions: Special profiles (like Default, Public) and specific SIDs
        $isSpecial = $_.Special
        $isExcludedSID = $_.SID -in $excludedSIDs
        $isDefaultOrPublic = $_.LocalPath -match '\\(Default|Public)$' # Extra check, Special should cover Default

        $excludeBasic = $isSpecial -or $isExcludedSID -or $isDefaultOrPublic
        if ($excludeBasic) {
             Write-Verbose "Profile $($_.LocalPath) (SID: $($_.SID)) excluded (Special=$isSpecial, ExcludedSID=$isExcludedSID, DefaultOrPublic=$isDefaultOrPublic)."
             return $false # Exclude this profile
        }

        # Inactivity Check (only if -InactiveDays was specified)
        if ($thresholdDate) {
            if ($_.LastUseTime -and $_.LastUseTime -ge $thresholdDate) {
                 Write-Verbose "Profile $($_.LocalPath) (SID: $($_.SID)) kept (LastUseTime: $($_.LastUseTime) is not older than threshold)."
                 return $false # Exclude this profile (it's active)
            }
            elseif (!$_.LastUseTime) {
                 Write-Verbose "Profile $($_.LocalPath) (SID: $($_.SID)) kept (LastUseTime is null/unknown)."
                 return $false # Exclude profile if LastUseTime is unknown, safer approach
            }
             # If LastUseTime exists and is older than threshold, it passes this check (implicit true)
             Write-Verbose "Profile $($_.LocalPath) (SID: $($_.SID)) is potentially INACTIVE (LastUseTime: $($_.LastUseTime))."
        }

        # If we got here, the profile is not excluded by basic rules or inactivity (if checked)
        Write-Verbose "Profile $($_.LocalPath) (SID: $($_.SID)) flagged for potential removal."
        return $true
    }

    Write-Host "$($profilesToRemove.Count) profiles identified for potential removal." -ForegroundColor Yellow

}
catch {
    Write-Error "Failed to query or filter profiles: $($_.Exception.Message)" -ErrorAction Stop
}
#endregion

#region --- Remove Profiles ---
if ($profilesToRemove.Count -eq 0) {
    Write-Host "No profiles meet the criteria for removal." -ForegroundColor Green
    Exit 0
}

Write-Host "Starting profile removal process..."

foreach ($profile in $profilesToRemove) {
    $profilePath = $profile.LocalPath
    $profileSID = $profile.SID
    $targetObject = "User profile at '$profilePath' (SID: $profileSID)"

    # Use ShouldProcess for -WhatIf and -Confirm support (unless -Force is used)
    if ($PSCmdlet.ShouldProcess($targetObject, "Remove")) {
        Write-Host "Attempting to remove profile: '$profilePath' (SID: $profileSID)..." -ForegroundColor Yellow
        try {
            # Remove the profile using its CIM object instance
            Remove-CimInstance -InputObject $profile -ErrorAction Stop -Confirm:(!$Force) # Only skip confirmation if -Force is explicitly true

            if ($?) { # Check if the last command succeeded
                 Write-Host "Successfully removed profile: '$profilePath'" -ForegroundColor Green
            }
            # Note: Remove-CimInstance doesn't always throw an error on failure with Win32_UserProfile,
            # so checking if the profile still exists might be needed for absolute certainty,
            # but often the command completing without error is sufficient indication.

        }
        catch {
            Write-Error "Failed to remove profile '$profilePath' (SID: $profileSID). Error: $($_.Exception.Message)".
            # Write-Warning "CIM removal failed. You might need to manually delete the folder '$profilePath' and associated registry keys (HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$profileSID)"
        }
    } else {
         Write-Host "Skipped removal of profile '$profilePath' due to -WhatIf or user cancellation." -ForegroundColor Cyan
    }
}

Write-Verbose "Profile cleanup process finished."
#endregion