# Creates a folder path if it doesn't exist
function Create-Path {
    if (-not (Test-Path -Path $LocalDestinationPath)) {
        try {
            New-Item -ItemType Directory -Path $LocalDestinationPath | Out-Null        
        }
        catch {
            $ErrorMessage = "Failed to create local directory: $($_.Exception.Message)"
            return
            Write-Log $ErrorMessage
        }
    }
}