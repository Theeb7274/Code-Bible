<#
Description
    - Applies security filtering to the specified GPO's for the specified group, and sets it to deny
    - This prevents the specified GPO's from inheriting onto the group.

 By Cam M V1.0

 Parameters
    - GpoNames is mandatory, and is stored as an array. Specify your GPOS to block inheritiance of here
    - GroupName is mandatory, it is the group you wish these deny rules to apply to
    - DomainName is not mandatory, the script will attempt to find the domain by itself. If this fails, consider using -DomainName instead
    - SupportsShouldProccess allows this script to be called using generic parameters, such as -WhatIf

Notes
    - A tracker is used for ccompleted and failed ACE changes
    - The parameter -WhatIf is great for testing, as no changes are made, but what would of been changed is displayed
    - There is a check for the applying ACE rule to prevent a duplicate change. This is often unecessary, as it should be able to handle this fine, but is there just in case

Changelog
    - V0.1 Untested
    - v0.2 Corrected a few syntax errors
    - V1.0 Tested & working.

#>

# Parameters Definitions
[CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='Low')]
param (
    [Parameter(Mandatory=$true)]
    [string[]]$GpoNames, # [] used to initialise an empty array, ensuring this is always handled as an array

    [Parameter(Mandatory=$true)]
    [string]$GroupName,

    [string]$DomainName
)

# Functions to test modules, and import them if found
function Test-ADModuleAvailable {
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        Write-Error "Active Directory Powershell module was not found"
        return $false # Avoids attempting to import a module which isn't found
    }
    Import-Module ActiveDirectory -ErrorAction SilentlyContinue
    return $true
}

function Test-GroupPolicyModuleAvailable {
    # Repeat of above function but for a different module
    if (-not (Get-Module -ListAvailable -Name GroupPolicy)) {
        Write-Error "Group Policy Powershell module was not found"
        return $false
    }
    Import-Module GroupPolicy -ErrorAction SilentlyContinue
    return $true
}

# Main Logic

# Exits if the functions return $false
if (-not (Test-ADModuleAvailable)) { exit 1 }
if (-not (Test-GroupPolicyModuleAvailable)) { exit 1 }

# Find Target Domain
$TargetDomain = $DomainName
if ([string]::IsNullOrWhiteSpace($TargetDomain)) { # Checks if $DomainName was specified
    try {
        $TargetDomain = (Get-ADDomain).DNSRoot
        Write-Verbose "Current domain detected as $TargetDomain"
    } catch {
        # Exit if we can't determine the domain
        Write-Error "Could not determine current domain"
        exit 1
    }
}

# Get SID of group
try {
    $AdGroup = Get-ADGroup -Identity $GroupName -Server $TargetDomain -ErrorAction Stop
    $GroupSid = New-Object System.Security.Principal.SecurityIdentifier ($AdGroup.SID.Value)
    Write-Verbose "Found group '$($AdGroup.Name)' with SID '$($GroupSid.Value)'."
} catch {
    # Exit if we cannot find the group
    Write-Error "Could not find AD group '$GroupName' in domain '$TargetDomain'. $_"
    exit 1
}

# GUID for "Apply Group Policy" Extended right, which is what Group Policy engine uses to determine if a GPO applies to a user or computer 
$applyGroupPolicyGuid = New-Object Guid "EDACFD8F-FFB3-11D1-B41D-00A0C968F939" 

# Trackers
$gposProcessed = 0
$gposFailed = 0

foreach ($GpoName in $GpoNames) { # Loop through $GpoNames array
    Write-Host "`nProcessing GPO: '$GpoName'..."
    try {
        $Gpo = Get-GPO -Name $GpoName -Domain $TargetDomain -ErrorAction Stop
        Write-Verbose "Found GPO '$($Gpo.DisplayName)' with ID '$($Gpo.Id)'."

        # Construct AD path to GPO object
        $GpoADPath = "AD:\CN={$($Gpo.Id)},CN=Policies,CN=System,$((Get-ADDomain -Server $TargetDomain).DistinguishedName)"
        Write-Verbose "GPO AD Path: $GpoADPath"

        if (-not (Test-Path $GpoADPath)) {
            Write-Error "GPO not found at $GpoADPath"
            $gposFailed++ # Tally the failure
            continue # But carry on
        }

        # Get the ACL of the GPO
        $Acl = Get-Acl -Path $GpoADPath -ErrorAction Stop

        # Check if the deny is already in place, even though it should handle a duplicate without issues
        $existingDenyRule = $Acl.Access | Where-Object {
            $_.IdentityReference -eq $GroupSid -and
            $_.AccessControlType -eq [System.Security.AccessControl.AccessControlType]::Deny -and
            $_.ObjectType -eq $applyGroupPolicyGuid -and
            $_.ActiveDirectoryRights -band [System.DirectoryServices.ActiveDirectoryRights]::ExtendedRight
        }

        if ($existingDenyRule) {
            Write-Warning "A Deny 'Apply Group Policy' for group '$GroupName' already exists on GPO '$GpoName'. No changes made."
            $gposProcessed++ # We still count it as processed
            continue
        }

        <#
        - Create the Deny ACE for "Apply Group Policy"
        - ActiveDirectoryRights for ExtendedRight needs to be specified
        - InheritanceFlags and PropagationFlags are tpyically None for direct permissions
        #>

        $Ace = New-Object System.DirectoryServices.ActiveDirectoryAccessRule(
            $GroupSid,
            [System.DirectoryServices.ActiveDirectoryRights]::ExtendedRight, # The right itself
            [System.Security.AccessControl.AccessControlType]::Deny,
            $applyGroupPolicyGuid, # The specific extended right GUID
            [System.Security.AccessControl.InheritanceFlags]::None,
            [System.Guid]::Empty # No inherited object type
        )

        if ($PSCmdlet.ShouldProcess("GPO '$($Gpo.Displayname)' (AD Path: $GpoADPath)", "Add Deny 'Apply Group Policy' ACE for group '$($AdGroup.Name)'")) {
            $Acl.AddAccessRule($Ace)
            Set-Acl -Path $GpoADPath -AclObject $Acl -ErrorAction Stop
            Write-Host "Successfully added Deny 'Apply GroupPolicy' ACE for group '$($AdGroup.Name)' to GPO '$($Gpo.DisplayName)'."
            $gposProcessed++
        } else { 
            Write-Warning "Skipped modifications of GPO '$($Gpo.Displayname)' due to -WhatIf or user cancellation."
        }
    }

    catch { 
        Write-Error "Failed to process GPO '$GpoName'. Error: $($_.Exception.Message)"
        $gposFailed++

    }
}




Write-Host "`n--- Summary ---"
Write-Host "GPOs successfully processed or already set: $gposProcessed"
Write-Host "GPOs failed to process: $gposFailed"

if ($gposFailed -gt 0) {
    Write-Warning "Some GPOs could not be processed."
}

