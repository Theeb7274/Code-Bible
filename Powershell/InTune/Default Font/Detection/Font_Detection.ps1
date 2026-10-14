$targetVersion = 3
$markerKey = 'HKCU:\Software\Contoso\OfficeFont'
$apps = 'Excel', 'Word', 'Outlook', 'PowerPoint'

$missing = foreach ($app in $apps) {
    $value = (Get-ItemProperty -Path $markerKey -Name $app -ErrorAction SilentlyContinue).$app
    if (-not $value -or $value -lt $targetVersion) { $app }
}

if ($missing) {
    Write-Output "Not applied: $($missing -join ', ')"
    exit 1
}
else {
    Write-Output 'All four apps applied'
    exit 0
}