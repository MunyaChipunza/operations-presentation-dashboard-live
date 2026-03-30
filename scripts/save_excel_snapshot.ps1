[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$SourcePath,

    [Parameter(Mandatory = $true)]
    [string]$TargetPath
)

$ErrorActionPreference = "Stop"

$excel = $null
$workbook = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.DisplayAlerts = $false
    $excel.Visible = $false
    $excel.ScreenUpdating = $false
    $excel.EnableEvents = $false

    Start-Sleep -Milliseconds 250
    $workbook = $excel.Workbooks.Open($SourcePath)

    $targetDir = Split-Path -Parent $TargetPath
    if ($targetDir -and -not (Test-Path -LiteralPath $targetDir)) {
        New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
    }

    $workbook.SaveCopyAs($TargetPath)
}
finally {
    if ($workbook -ne $null) {
        $workbook.Close($false) | Out-Null
    }

    if ($workbook -ne $null) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
    }

    if ($excel -ne $null) {
        $excel.Quit() | Out-Null
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
