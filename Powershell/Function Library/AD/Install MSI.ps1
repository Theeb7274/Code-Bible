# Execute MSI 
function Install-App {
    param(
        [parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]$ExecutablePath,

        [parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]$LogFilePath
    )

    try {
        if 
        Start-Process msiexec -ArgumentList "/i `"$ExecutablePath`"","/qn""/li $LogFilePath"  -Wait;
        Write-Log "Executable ran successfully."
    }
    catch {
        $ErrorMessage = "Failed to run msi, refer to $LogFilePath for details"
        Write-Log $ErrorMessage
        return;
    }
}