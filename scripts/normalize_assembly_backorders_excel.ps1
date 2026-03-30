param(
    [Parameter(Mandatory = $true)]
    [string]$WorkbookPath,

    [string]$SheetName = "DATA",

    [string]$BackupDirectory
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Parse-ExcelNumber {
    param([object]$Value)

    if ($null -eq $Value) {
        return $null
    }
    if ($Value -is [double] -or $Value -is [int] -or $Value -is [decimal]) {
        return [double]$Value
    }

    $text = [string]$Value
    $text = $text.Trim()
    if (-not $text) {
        return $null
    }
    $text = $text.Replace(",", "")
    $number = 0.0
    if ([double]::TryParse($text, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$number)) {
        return $number
    }
    return $null
}

$resolvedWorkbook = (Resolve-Path -LiteralPath $WorkbookPath).ProviderPath
if (-not $BackupDirectory) {
    $BackupDirectory = Join-Path (Split-Path -Parent $resolvedWorkbook) "Backups"
}
$resolvedBackup = [System.IO.Path]::GetFullPath($BackupDirectory)
New-Item -ItemType Directory -Path $resolvedBackup -Force | Out-Null

$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$backupPath = Join-Path $resolvedBackup ("PPT presentation source data - before assembly normalize $timestamp.xlsx")
Copy-Item -LiteralPath $resolvedWorkbook -Destination $backupPath -Force

$excel = $null
$workbook = $null
$sheet = $null
$converted = 0

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false

    $workbook = $excel.Workbooks.Open($resolvedWorkbook)
    $sheet = $workbook.Worksheets.Item($SheetName)
    try {
        $excel.Calculation = -4105
    }
    catch {
        # Some Excel installs reject calculation mode writes through COM; the full rebuild below still forces recalc.
    }

    $sheet.Range("AC2").Value2 = "Month"
    $sheet.Range("AD2").Value2 = "Assembled"
    $sheet.Range("AE2").Value2 = "Backorders"
    $sheet.Range("AF2").Value2 = "Fill Rate"

    foreach ($row in 3..14) {
        foreach ($column in "AD", "AE") {
            $cell = $sheet.Range("$column$row")
            $parsed = Parse-ExcelNumber $cell.Value2
            if ($null -ne $parsed) {
                if ($cell.Value2 -is [string]) {
                    $converted += 1
                }
                $cell.Value2 = $parsed
                $cell.NumberFormat = "#,##0"
            }
        }

        $formulaCell = $sheet.Range("AF$row")
        $formulaCell.FormulaR1C1 = '=IF(OR(RC[-2]="",RC[-1]="",RC[-2]=0),"",1-RC[-1]/RC[-2])'
        $formulaCell.NumberFormat = "0.0%"
    }

    $workbook.Application.CalculateFullRebuild()
    $workbook.Save()
}
finally {
    if ($workbook) {
        $workbook.Close($true)
    }
    if ($excel) {
        $excel.Quit()
    }
    foreach ($item in $sheet, $workbook, $excel) {
        if ($item) {
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($item)
        }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Write-Output "Normalized assemblies/backorders block in $resolvedWorkbook"
Write-Output "Converted text numbers: $converted"
Write-Output "Backup created at $backupPath"
