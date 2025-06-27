<#
 =================================================================================
 PowerShell Script to Fix OneDrive "Site User ID Mismatch"
 This script automatically determines the user's OneDrive URL from their UPN.
 =================================================================================

 --- Step 1: Set the Required Variables ---
 IMPORTANT: You ONLY need to update these two variables.

 The User Principal Name (UPN) of the user having the issue.
 $UserUPN = "user.experiencing.issue@yourdomain.com" 

 The URL to your SharePoint Admin Center.
 It usually follows this format: https://[tenant]-admin.sharepoint.com
 $AdminURL = "https://yourtenant-admin.sharepoint.com"


 --- Step 2: Connect to SharePoint Online Admin Center ---
 A login prompt will appear. Sign in with your SharePoint Admin or Global Admin account.
 Write-Host "Connecting to SharePoint Online Admin Center at $AdminURL..." -ForegroundColor Yellow
 Connect-SPOService -Url $AdminURL


 --- Step 3: Automatically Construct the OneDrive URL ---
 Write-Host "Automatically determining OneDrive URL for $UserUPN..." -ForegroundColor Cyan
#>

# Get the tenant's MySite Host URL (the base URL for all OneDrives)
try {
    $mySiteHost = (Get-SPOTenant).MySiteHostUrl
    Write-Host $mySiteHosto
    if (-not $mySiteHost) {
        throw "Could not retrieve MySite Host URL. Please ensure you are connected."
    }
}
catch {
    Write-Host "Error: Could not retrieve the MySite Host URL from your tenant. Please check your connection and permissions." -ForegroundColor Red
    # Disconnect before exiting
    Disconnect-SPOService
    return
}

# Convert the user's UPN to the format used in OneDrive URLs
# Example: "jane.smith@contoso.com" becomes "jane_smith_contoso_com"
$formattedUser = $UserUPN -replace '@', '_' -replace '\.', '_'

# Combine the parts to create the full OneDrive URL
$OneDriveURL = "$mySiteHost/personal/$formattedUser"

Write-Host "Derived OneDrive URL: $OneDriveURL" -ForegroundColor Green


# --- Step 4: Run the Fix ---
# The script will now use the automatically generated $OneDriveURL.
# By default, it runs Method 1. To run Method 2, comment out the first line and uncomment the second.

# === METHOD 1: Re-sync User from Azure AD (Recommended) ===
Write-Host "Attempting to fix user ID mismatch for '$UserUPN' by re-syncing from Azure AD." -ForegroundColor Green
Set-SPOUser -Site $OneDriveURL -LoginName $UserUPN -SyncFromAD


# === METHOD 2: Remove and Re-create User Profile (More forceful, use if Method 1 fails) ===
# NOTE: After running this, the user MUST navigate to their OneDrive to have their profile recreated.
# Remove the '#' from the line below to run this method.
# Write-Host "Attempting to fix user ID mismatch for '$UserUPN' by REMOVING the user profile from the site collection." -ForegroundColor Magenta
# Remove-SPOUser -Site $OneDriveURL -LoginName $UserUPN


# --- Step 5: Confirmation ---
Write-Host "Process complete. Please have the user wait 5-10 minutes, then clear their browser cache and try accessing their OneDrive again." -ForegroundColor Cyan
Write-Host "If the issue persists, try running Method 2." -ForegroundColor Yellow

# Disconnect the session
Disconnect-SPOService