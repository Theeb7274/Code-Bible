# ===============================
# Block all sign-ins except specified accounts
# Includes function to connect to Exchange Online
# Tested working 21/10/2025
# Note that the account used won't be able to block accounts of the same permission set or higher, I.E global admin accounts.
# The output of the accounts being blocked is printed live into the shell output. It will continue on errors
# ===============================

# Function: Check for required modules, install them if missing
function Install-Modules {

    $modules = @(
        "Microsoft.Graph.Users",
        "ExchangeOnlineManagement"
)

    foreach ($module in $modules) {
        if (-not (Get-Module -ListAvailable -Name $module)) {
            Write-Host "Installing module: $module..."
            Install-Module $module -Scope CurrentUser -Force
     } else {
          Write-Host "Module already installed: $module"
      }
     Import-Module $module -Force
    }
}

# Define allowed accounts (UPNs)
$allowedUsers = @(
    ""
)

# Function: Connect to Microsoft Graph and Exchange Online
function Connect-Services {
    Write-Host "Connecting to Microsoft Graph..."
    Connect-MgGraph -Scopes "User.ReadWrite.All" -ErrorAction Stop

    Write-Host "Connecting to Exchange Online..."
    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop

    Write-Host "Connected to both Microsoft Graph and Exchange Online."
}

# Function: Block all users except those in $allowedUsers
function Block-UsersExcept {
    param (
        [string[]]$AllowedUsers
    )

    Write-Host "Fetching all Azure AD users..."
    $allUsers = Get-MgUser -All

    foreach ($user in $allUsers) {
        if ($AllowedUsers -notcontains $user.UserPrincipalName) {
            Update-MgUser -UserId $user.Id -AccountEnabled:$false
            Write-Host "Blocked: $($user.UserPrincipalName)"
        } else {
            Write-Host "Allowed: $($user.UserPrincipalName)"
        }
    }
}

# Call Functions
Install-Modules
Connect-Services
Block-UsersExcept -AllowedUsers $allowedUsers
