# Remediate-OfficeDefaultFont.ps1

$FontName = "Century Gothic"
$FontSize = 11

$ErrorActionPreference = "SilentlyContinue"

# ------------------------------------------------------------
# Check whether Office applications are running
# ------------------------------------------------------------

$OfficeProcesses = @(
    "WINWORD",
    "EXCEL",
    "POWERPNT",
    "OUTLOOK"
)

$Running = Get-Process -Name $OfficeProcesses -ErrorAction SilentlyContinue

#if ($Running) {
#    Write-Output "Office application(s) currently running. Remediation will retry later."
#    exit 1
#}

# ------------------------------------------------------------
# Excel
# ------------------------------------------------------------

$ExcelReg = "HKCU:\Software\Microsoft\Office\16.0\Excel\Options"

if (-not (Test-Path $ExcelReg)) {
    New-Item -Path $ExcelReg -Force | Out-Null
}

Set-ItemProperty `
    -Path $ExcelReg `
    -Name "Font" `
    -Value $FontName `
    -Type String

Set-ItemProperty `
    -Path $ExcelReg `
    -Name "FontSize" `
    -Value $FontSize `
    -Type String

Write-Output "Excel configured."

# ------------------------------------------------------------
# Word - Normal.dotm
# ------------------------------------------------------------

$TemplateFolder = Join-Path $env:APPDATA "Microsoft\Templates"

if (-not (Test-Path $TemplateFolder)) {
    New-Item -Path $TemplateFolder -ItemType Directory -Force | Out-Null
}

$WordTemplate = Join-Path $TemplateFolder "Normal.dotm"

try {

    $Word = New-Object -ComObject Word.Application
    $Word.Visible = $false
    $Word.DisplayAlerts = 0

    if (Test-Path $WordTemplate) {

        $Template = $Word.Documents.Open(
            $WordTemplate,
            $false,
            $false
        )
    }
    else {

        $Template = $Word.Documents.Add()

        $Template.SaveAs2(
            $WordTemplate,
            13
        )
    }

    $Style = $Template.Styles.Item("Normal")

    $Style.Font.Name = $FontName
    $Style.Font.NameAscii = $FontName
    $Style.Font.NameOther = $FontName
    $Style.Font.NameFarEast = $FontName
    $Style.Font.NameBi = $FontName
    $Style.Font.Size = $FontSize

    $Template.Save()

    $Template.Close($true)
    $Word.Quit()

    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Style) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Template) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Word) | Out-Null

    Write-Output "Word configured."
}
catch {
    Write-Output "Word configuration failed: $($_.Exception.Message)"
}

# ------------------------------------------------------------
# Outlook - NormalEmail.dotm
# ------------------------------------------------------------

$OutlookTemplate = Join-Path $TemplateFolder "NormalEmail.dotm"

try {

    $Word = New-Object -ComObject Word.Application
    $Word.Visible = $false
    $Word.DisplayAlerts = 0

    if (Test-Path $OutlookTemplate) {

        $Template = $Word.Documents.Open(
            $OutlookTemplate,
            $false,
            $false
        )
    }
    else {

        $Template = $Word.Documents.Add()

        $Template.SaveAs2(
            $OutlookTemplate,
            13
        )
    }

    $Style = $Template.Styles.Item("Normal")

    $Style.Font.Name = $FontName
    $Style.Font.NameAscii = $FontName
    $Style.Font.NameOther = $FontName
    $Style.Font.NameFarEast = $FontName
    $Style.Font.NameBi = $FontName
    $Style.Font.Size = $FontSize

    $Template.Save()

    $Template.Close($true)
    $Word.Quit()

    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Style) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Template) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Word) | Out-Null

    Write-Output "Outlook configured."
}
catch {
    Write-Output "Outlook configuration failed: $($_.Exception.Message)"
}

# ------------------------------------------------------------
# PowerPoint - Blank.potx
# ------------------------------------------------------------

$PowerPointTemplate = Join-Path $TemplateFolder "Blank.potx"

try {

    $PowerPoint = New-Object -ComObject PowerPoint.Application

    $Presentation = $PowerPoint.Presentations.Add()

    # Configure the slide master
    $Master = $Presentation.SlideMaster

    foreach ($Shape in $Master.Shapes) {

        if ($Shape.HasTextFrame -eq -1) {

            try {
                $Shape.TextFrame2.TextRange.Font.Name = $FontName
                $Shape.TextFrame2.TextRange.Font.Size = $FontSize
            }
            catch {}
        }
    }

    # Configure each layout
    foreach ($Layout in $Master.CustomLayouts) {

        foreach ($Shape in $Layout.Shapes) {

            if ($Shape.HasTextFrame -eq -1) {

                try {
                    $Shape.TextFrame2.TextRange.Font.Name = $FontName
                    $Shape.TextFrame2.TextRange.Font.Size = $FontSize
                }
                catch {}
            }
        }
    }

    # 26 = ppSaveAsOpenXMLTemplate
    $Presentation.SaveAs(
        $PowerPointTemplate,
        26
    )

    $Presentation.Close()
    $PowerPoint.Quit()

    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Master) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Presentation) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($PowerPoint) | Out-Null

    Write-Output "PowerPoint configured."
}
catch {
    Write-Output "PowerPoint configuration failed: $($_.Exception.Message)"
}

Write-Output "Office default font remediation completed."
exit 0