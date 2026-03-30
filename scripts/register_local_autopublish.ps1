[CmdletBinding()]
param(
    [string]$TaskName = "Operations PPT Dashboard Auto Publish",
    [string]$WorkbookPath = ""
)

$ErrorActionPreference = "Stop"

$runnerScriptPath = (Resolve-Path (Join-Path $PSScriptRoot "run_local_autopublish.pyw")).Path
$pythonPath = (Get-Command python -ErrorAction Stop).Source
$pythonwPath = Join-Path (Split-Path $pythonPath) "pythonw.exe"
$runnerExe = if (Test-Path $pythonwPath) { $pythonwPath } else { $pythonPath }
$triggerTime = (Get-Date).AddMinutes(1)

if ($WorkbookPath) {
    $candidateWorkbook = (Resolve-Path (Join-Path $PSScriptRoot $WorkbookPath)).Path
} else {
    $searchRoots = @(
        (Resolve-Path (Join-Path $PSScriptRoot "..\\..\\..")).Path,
        (Resolve-Path (Join-Path $PSScriptRoot "..\\..")).Path
    )
    $workbookMatch = $null
    foreach ($root in $searchRoots) {
        $preferred = Get-ChildItem -Path $root -File -Filter "*Operations Data*.xlsx" -ErrorAction SilentlyContinue | Select-Object -First 1
        if (-not $preferred) {
            $preferred = Get-ChildItem -Path $root -File -Filter "*PPT presentation source data*.xlsx" -ErrorAction SilentlyContinue | Select-Object -First 1
        }
        if (-not $preferred) {
            $preferred = Get-ChildItem -Path $root -File -ErrorAction SilentlyContinue | Where-Object { $_.Extension -in ".xlsx", ".xlsm" } | Select-Object -First 1
        }
        if ($preferred) {
            $workbookMatch = $preferred.FullName
            break
        }
    }
    if (-not $workbookMatch) {
        throw "Could not find a workbook to wire into the auto-publish task."
    }
    $candidateWorkbook = $workbookMatch
}

$taskArgs = '"' + $runnerScriptPath + '" --workbook "' + $candidateWorkbook + '"'
$action = New-ScheduledTaskAction -Execute $runnerExe -Argument $taskArgs -WorkingDirectory $PSScriptRoot
$trigger = New-ScheduledTaskTrigger -Once -At $triggerTime -RepetitionInterval (New-TimeSpan -Minutes 1) -RepetitionDuration (New-TimeSpan -Days 3650)
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -MultipleInstances IgnoreNew

Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Settings $settings -Description "Publishes the operations presentation dashboard from the local workbook every minute." -Force | Out-Null

Write-Host "Scheduled task created:"
Write-Host "  Name: $TaskName"
Write-Host "  Workbook: $candidateWorkbook"
Write-Host "  Runs every minute on this PC."
