# A function to check the script has been ran with administrator rights, and exits if not
function AdminCheck {

    if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-Error "This script requires Administrator privileges. Please re-run PowerShell as Administrator."
        exit 1
    }
}
