#Requires -RunAsAdministrator

$ErrorActionPreference = 'Stop'

$InstallDir = 'C:\EDUIT\Scripts'
$WorkerPath = Join-Path $InstallDir 'FontInstall.ps1'
$TaskName   = 'RunScriptAtLogon'

# ---- The worker script. Everything between @' and '@ is written to disk verbatim. ----
$WorkerContent = @'
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
'@
# ---------------------------------------------------------------------------------------

function Install-Worker {
    param(
        [Parameter(Mandatory)][string]$Directory,
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Content
    )

    # New-Item -Path $Directory -ItemType Directory -Force | Out-Null

    $current = $null
    if (Test-Path -Path $Path -PathType Leaf) {
        $current = [System.IO.File]::ReadAllText($Path)
    }

    if ($current -ceq $Content) {
        Write-Output "Worker script already up to date at $Path"
        return
    }

    $utf8Bom = New-Object System.Text.UTF8Encoding($true)
    [System.IO.File]::WriteAllText($Path, $Content, $utf8Bom)
    Write-Output "Worker script written to $Path"
}

function Set-WorkerAcl {
    param([Parameter(Mandatory)][string]$Directory)

    & icacls.exe $Directory /inheritance:r `
        /grant:r '*S-1-5-18:(OI)(CI)F' `
                 '*S-1-5-32-544:(OI)(CI)F' `
                 '*S-1-5-32-545:(OI)(CI)RX' | Out-Null

    if ($LASTEXITCODE -ne 0) {
        throw "icacls failed with exit code $LASTEXITCODE"
    }
    Write-Output "Permissions set on $Directory"
}

function New-LogonTask {
    param(
        [Parameter(Mandatory)][string]$TaskName,
        [Parameter(Mandatory)][string]$ScriptPath
    )

    $arguments = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$ScriptPath`""

    $action    = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument $arguments
    $trigger   = New-ScheduledTaskTrigger -AtLogOn
    $principal = New-ScheduledTaskPrincipal -GroupId 'BUILTIN\Users' -RunLevel Limited

    Register-ScheduledTask -TaskName $TaskName `
                           -Action $action `
                           -Trigger $trigger `
                           -Principal $principal `
                           -Force | Out-Null
    Write-Output "Scheduled task '$TaskName' registered"
}

try {
    Install-Worker -Directory $InstallDir -Path $WorkerPath -Content $WorkerContent
    Set-WorkerAcl  -Directory $InstallDir
    New-LogonTask  -TaskName $TaskName -ScriptPath $WorkerPath
    exit 0
}
catch {
    Write-Error $_
    exit 1
}