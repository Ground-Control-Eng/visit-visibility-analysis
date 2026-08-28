<#
.SYNOPSIS
    Registers (or re-registers) the daily "Visit Reconciliation" Windows Scheduled Task.

.NOTES
    Run "only when user is logged on" is required because the script drives Outlook via COM,
    which needs an interactive desktop session. This means the task will not run if the
    machine is fully logged off (a locked session should still be fine, but test this on
    this specific machine before relying on it unattended).
#>

param(
    [string]$TaskName = "Visit Reconciliation Daily Check",
    [string]$RunTime = "21:15"
)

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$BatPath = Join-Path $ScriptDir "run_daily_check.bat"

if (-not (Test-Path $BatPath)) {
    throw "Cannot find run_daily_check.bat at $BatPath"
}

$Action = New-ScheduledTaskAction -Execute $BatPath -WorkingDirectory (Split-Path -Parent $ScriptDir)
$Trigger = New-ScheduledTaskTrigger -Daily -At $RunTime
$Principal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType Interactive -RunLevel Limited
$Settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -DontStopOnIdleEnd -ExecutionTimeLimit (New-TimeSpan -Minutes 30)

Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue

Register-ScheduledTask -TaskName $TaskName -Action $Action -Trigger $Trigger -Principal $Principal -Settings $Settings

Write-Host "Registered scheduled task '$TaskName' to run daily at $RunTime as $env:USERDOMAIN\$env:USERNAME (interactive logon required)."
