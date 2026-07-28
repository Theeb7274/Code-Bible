# Attempts to delete $File
function Remove-File {
    param ([String]$File)

    # Delete local zip file
    try {
            Remove-Item -LiteralPath $File -Force
            Write-Log "Local zip '$File' deleted successfully or not present."
        }
    catch {
            $WarningMessage = "Failed to delete local zip file: $($_.Exception.Message)"
            Write-Log $WarningMessage
        }
}