[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$WorkbookPath
)

$ErrorActionPreference = "Stop"

$publishScript = (Resolve-Path (Join-Path $PSScriptRoot "publish_dashboard_data.py")).Path
$workbookFullPath = (Resolve-Path $WorkbookPath).Path
$pythonPath = (Get-Command python -ErrorAction Stop).Source
$pythonwPath = Join-Path (Split-Path $pythonPath) "pythonw.exe"
$runnerPath = if (Test-Path $pythonwPath) { $pythonwPath } else { $pythonPath }

Push-Location (Split-Path $publishScript)
try {
    $env:PYTHONDONTWRITEBYTECODE = "1"
    & $runnerPath $publishScript --workbook $workbookFullPath
    if ($LASTEXITCODE -ne 0) {
        throw "Local dashboard publish failed with exit code $LASTEXITCODE."
    }
} finally {
    Pop-Location
}
