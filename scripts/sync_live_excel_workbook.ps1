[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$SourcePath
)

$ErrorActionPreference = "Stop"

function Normalize-Path([string]$PathValue) {
    return [System.IO.Path]::GetFullPath($PathValue).TrimEnd('\').ToLowerInvariant()
}

$excel = $null
$workbook = $null

try {
    try {
        $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
    }
    catch {
        exit 3
    }

    $sourceFullName = Normalize-Path $SourcePath
    foreach ($candidate in @($excel.Workbooks)) {
        if (-not $candidate.FullName) {
            continue
        }

        if ((Normalize-Path $candidate.FullName) -eq $sourceFullName) {
            $workbook = $candidate
            break
        }
    }

    if ($workbook -eq $null) {
        exit 3
    }

    if ($workbook.ReadOnly) {
        Write-Output "Workbook is open read-only; skipping live save."
        exit 0
    }

    if (-not $workbook.Saved) {
        $workbook.Save()
        Start-Sleep -Milliseconds 250
        Write-Output "Workbook saved from active Excel session."
    }
    else {
        Write-Output "Workbook already saved in active Excel session."
    }
}
finally {
    if ($workbook -ne $null) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
    }

    if ($excel -ne $null) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
