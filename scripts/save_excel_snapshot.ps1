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
$openedHere = $false
$createdInstance = $false

try {
    try {
        $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
    }
    catch {
        $excel = New-Object -ComObject Excel.Application
        $createdInstance = $true
    }

    $excel.DisplayAlerts = $false
    $excel.Visible = $false

    foreach ($candidate in @($excel.Workbooks)) {
        if ([string]::Equals($candidate.FullName, $SourcePath, [System.StringComparison]::OrdinalIgnoreCase)) {
            $workbook = $candidate
            break
        }
    }

    if ($null -eq $workbook) {
        Start-Sleep -Milliseconds 250
        $workbook = $excel.Workbooks.Open($SourcePath)
        $openedHere = $true
    }

    $targetDir = Split-Path -Parent $TargetPath
    if ($targetDir -and -not (Test-Path -LiteralPath $targetDir)) {
        New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
    }

    $workbook.SaveCopyAs($TargetPath)
}
finally {
    if ($openedHere -and $workbook -ne $null) {
        $workbook.Close($false) | Out-Null
    }

    if ($workbook -ne $null) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
    }

    if ($excel -ne $null) {
        if ($createdInstance) {
            $excel.Quit() | Out-Null
        }
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
