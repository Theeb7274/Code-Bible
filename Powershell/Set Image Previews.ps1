# Ticks "Always show icons, never thumbnails" for the current user.
# Must run in the USER context (HKCU), not as SYSTEM.

$regPath = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced'
$name    = 'IconsOnly'
$value   = 1   # 1 = icons only, 0 = thumbnails

try {
    # Create the key if it does not exist
    if (-not (Test-Path -Path $regPath)) {
        New-Item -Path $regPath -Force | Out-Null
    }

    Set-ItemProperty -Path $regPath -Name $name -Value $value -Type DWord -Force

    Write-Output "IconsOnly set to $value"
    exit 0
}
catch {
    # Non-zero exit code makes Intune report the script as failed
    Write-Error $_.Exception.Message
    exit 1
}