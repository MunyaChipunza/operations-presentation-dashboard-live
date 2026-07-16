[CmdletBinding()]
param(
    [string]$TaskName = "Operations PPT Dashboard Auto Publish",
    [string]$WorkbookPath = ""
)

$ErrorActionPreference = "Stop"

function Test-PythonOpenpyxl {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PythonPath
    )

    if (-not (Test-Path $PythonPath)) {
        return $false
    }

    $result = & $PythonPath -c "import openpyxl" 2>$null
    return $LASTEXITCODE -eq 0
}

function Resolve-PythonRunner {
    $candidates = @()

    $candidates += @(
        "C:\Users\Dell\AppData\Local\Python\pythoncore-3.14-64\python.exe",
        "C:\Users\Dell\AppData\Local\Programs\Python\Python313\python.exe"
    )

    try {
        $pyList = & py -0p 2>$null
        foreach ($line in $pyList) {
            if ($line -match '^\s*-V:[^ ]+\s+\*?\s*(.+python\.exe)\s*$') {
                $candidates += $Matches[1].Trim()
            }
        }
    }
    catch {
    }

    try {
        $commandPython = (Get-Command python -ErrorAction Stop).Source
        if ($commandPython) {
            $candidates += $commandPython
        }
    }
    catch {
    }

    $seen = @{}
    foreach ($candidate in $candidates) {
        if (-not $candidate) {
            continue
        }
        $normalized = [System.IO.Path]::GetFullPath($candidate)
        if ($seen.ContainsKey($normalized)) {
            continue
        }
        $seen[$normalized] = $true

        if ($normalized -like '*\WindowsApps\*') {
            continue
        }

        if (-not (Test-PythonOpenpyxl -PythonPath $normalized)) {
            continue
        }

        $pythonwCandidate = Join-Path (Split-Path $normalized) "pythonw.exe"
        if (Test-Path $pythonwCandidate) {
            return $pythonwCandidate
        }
        return $normalized
    }

    throw "Could not find a Python interpreter with openpyxl installed."
}

$runnerScriptPath = (Resolve-Path (Join-Path $PSScriptRoot "run_local_autopublish.pyw")).Path
$runnerExe = Resolve-PythonRunner
$triggerTime = (Get-Date).AddMinutes(1)

if ($WorkbookPath) {
    if ([System.IO.Path]::IsPathRooted($WorkbookPath)) {
        $candidateWorkbook = (Resolve-Path $WorkbookPath).Path
    }
    else {
        $candidateWorkbook = (Resolve-Path (Join-Path $PSScriptRoot $WorkbookPath)).Path
    }
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
