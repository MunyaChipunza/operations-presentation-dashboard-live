[CmdletBinding()]
param(
    [string]$TaskName = "Operations PPT Dashboard Auto Publish",
    [string]$WorkbookPath = "..\\..\\..\\PPT presentation source data.xlsx"
)

$ErrorActionPreference = "Stop"

$runnerScriptPath = (Resolve-Path (Join-Path $PSScriptRoot "run_local_autopublish.pyw")).Path
$workbookFullPath = (Resolve-Path (Join-Path $PSScriptRoot $WorkbookPath)).Path
$pythonPath = (Get-Command python -ErrorAction Stop).Source
$pythonwPath = Join-Path (Split-Path $pythonPath) "pythonw.exe"
$runnerExe = if (Test-Path $pythonwPath) { $pythonwPath } else { $pythonPath }
$triggerTime = (Get-Date).AddMinutes(1)
$taskArgs = '"' + $runnerScriptPath + '" --workbook "' + $workbookFullPath + '"'
$action = New-ScheduledTaskAction -Execute $runnerExe -Argument $taskArgs -WorkingDirectory $PSScriptRoot
$trigger = New-ScheduledTaskTrigger -Once -At $triggerTime -RepetitionInterval (New-TimeSpan -Minutes 1) -RepetitionDuration (New-TimeSpan -Days 3650)
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -MultipleInstances IgnoreNew

Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Settings $settings -Description "Publishes the operations presentation dashboard from the local workbook every minute." -Force | Out-Null

Write-Host "Scheduled task created:"
Write-Host "  Name: $TaskName"
Write-Host "  Workbook: $workbookFullPath"
Write-Host "  Runs every minute on this PC."
