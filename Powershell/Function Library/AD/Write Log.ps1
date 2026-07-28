# Logs output with date & time

$LogFilePath # path defined outside of function as it's designed to be called numerous times throughout a script

function Write-Log {
    param (
        [string]$Message
    )
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogEntry = "$Timestamp - $Message"
    Add-Content -Path $LogFilePath -Value $LogEntry
}