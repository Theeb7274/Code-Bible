#Requires -RunAsAdministrator

param(
    [string]$ScriptPath = '',
    [string]$TaskName   = ''
)

function New-LogonTask {
    param(
        [Parameter(Mandatory)][string]$TaskName,
        [Parameter(Mandatory)][string]$ScriptPath
    )

    if (-not (Test-Path -Path $ScriptPath -PathType Leaf)) {
        Write-Warning "Script not found at '$ScriptPath'. The task will still be created."
    }

    $arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$ScriptPath`""

    $action    = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument $arguments
    $trigger   = New-ScheduledTaskTrigger -AtLogOn
    $principal = New-ScheduledTaskPrincipal -GroupId 'BUILTIN\Users' -RunLevel Limited

    Register-ScheduledTask -TaskName $TaskName `
                           -Action $action `
                           -Trigger $trigger `
                           -Principal $principal `
                           -Force
}

New-LogonTask -TaskName $TaskName -ScriptPath $ScriptPath