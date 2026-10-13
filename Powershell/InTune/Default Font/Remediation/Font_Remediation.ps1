$ErrorActionPreference = 'Stop'

$targetVersion = 3
$fontName = 'Century Gothic'
$fontSize = 11
$markerKey = 'HKCU:\Software\EDUIT\OfficeFont'
$templates = Join-Path $env:APPDATA 'Microsoft\Templates'

function Test-AppRunning {
    param([string[]]$ProcessName)
    return [bool](Get-Process -Name $ProcessName -ErrorAction SilentlyContinue)
}

function Test-Applied {
    param([string]$App)
    $value = (Get-ItemProperty -Path $markerKey -Name $App -ErrorAction SilentlyContinue).$App
    return ($value -ge $targetVersion)
}

function Set-Applied {
    param([string]$App)
    if (-not (Test-Path $markerKey)) { New-Item -Path $markerKey -Force | Out-Null }
    Set-ItemProperty -Path $markerKey -Name $App -Value $targetVersion -Type DWord
}

function Set-WordTemplateFont {
    param([string]$Font, [int]$Size, [string]$Path)

    $word = New-Object -ComObject Word.Application
    try {
        $word.Visible = $false
        $word.DisplayAlerts = 0

        if ($Path) { $doc = $word.Documents.Open($Path, $false, $false, $false) }
        else       { $doc = $word.NormalTemplate.OpenAsDocument() }

        $normal = $doc.Styles.Item(-1)
        $normal.Font.Name = $Font
        $normal.Font.Size = $Size
        $doc.Save()
        $doc.Close(0)
    }
    finally {
        $word.Quit()
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
    }
}

function Set-PowerPointDefaultFont {
    param([string]$Font, [string]$Path)

    $ppt = New-Object -ComObject PowerPoint.Application
    try {
        if (Test-Path $Path) { $pres = $ppt.Presentations.Open($Path, 0, 0, 0) }
        else                 { $pres = $ppt.Presentations.Add(0) }

        $scheme = $pres.SlideMaster.Theme.ThemeFontScheme
        $scheme.MajorFont.Item(1).Name = $Font
        $scheme.MinorFont.Item(1).Name = $Font
        $pres.SaveAs($Path, 26)
        $pres.Close()
    }
    finally {
        $ppt.Quit()
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($ppt)
    }
}

# Excel: registry only, safe while Excel is open
if (-not (Test-Applied 'Excel')) {
    $excelKey = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Options'
    if (-not (Test-Path $excelKey)) { New-Item -Path $excelKey -Force | Out-Null }
    Set-ItemProperty -Path $excelKey -Name 'Font' -Value "$fontName,$fontSize" -Type String
    Set-Applied 'Excel'
    Write-Output 'Excel applied'
}

# Word: edits Normal.dotm through COM
if (-not (Test-Applied 'Word')) {
    if (Test-AppRunning 'WINWORD') { Write-Output 'Word is open, skipping' }
    else {
        Set-WordTemplateFont -Font $fontName -Size $fontSize
        Set-Applied 'Word'
        Write-Output 'Word applied'
    }
}

# Outlook: edits NormalEmail.dotm through a hidden Word instance
if (-not (Test-Applied 'Outlook')) {
    $emailTemplate = Join-Path $templates 'NormalEmail.dotm'
    if (Test-AppRunning 'OUTLOOK', 'WINWORD') { Write-Output 'Outlook or Word is open, skipping' }
    elseif (-not (Test-Path $emailTemplate)) { Write-Output 'NormalEmail.dotm not found, skipping' }
    else {
        Set-WordTemplateFont -Font $fontName -Size $fontSize -Path $emailTemplate
        Set-Applied 'Outlook'
        Write-Output 'Outlook applied'
    }
}

# PowerPoint: edits the theme fonts in the template file
if (-not (Test-Applied 'PowerPoint')) {
    if (Test-AppRunning 'POWERPNT') { Write-Output 'PowerPoint is open, skipping' }
    else {
        Set-PowerPointDefaultFont -Font $fontName -Path (Join-Path $templates 'Blank.potx')
        Set-Applied 'PowerPoint'
        Write-Output 'PowerPoint applied'
    }
}

exit 0