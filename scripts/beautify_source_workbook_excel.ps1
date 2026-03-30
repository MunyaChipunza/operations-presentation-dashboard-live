param(
    [Parameter(Mandatory = $true)]
    [string]$WorkbookPath,

    [string]$SheetName = "DATA",

    [string]$BackupDirectory
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Get-ExcelColor {
    param(
        [int]$Red,
        [int]$Green,
        [int]$Blue
    )

    return $Red + ($Green * 256) + ($Blue * 65536)
}

function Set-Borders {
    param(
        $Range,
        [int]$Color,
        [int]$Weight = 2
    )

    foreach ($borderId in 7, 8, 9, 10, 11, 12) {
        $border = $Range.Borders.Item($borderId)
        $border.LineStyle = 1
        $border.Weight = $Weight
        $border.Color = $Color
    }
}

function Format-TitleRange {
    param(
        $Range,
        [int]$FillColor,
        [int]$BorderColor
    )

    $Range.MergeCells = $false
    $Range.Interior.Color = $FillColor
    $Range.Font.Name = "Aptos Display"
    $Range.Font.Size = 12
    $Range.Font.Bold = $true
    $Range.Font.Color = (Get-ExcelColor 255 255 255)
    $Range.HorizontalAlignment = 7
    $Range.VerticalAlignment = -4108
    Set-Borders -Range $Range -Color $BorderColor
}

function Format-HeaderRange {
    param(
        $Range,
        [int]$FillColor,
        [int]$BorderColor
    )

    $Range.Interior.Color = $FillColor
    $Range.Font.Name = "Aptos"
    $Range.Font.Size = 10
    $Range.Font.Bold = $true
    $Range.Font.Color = (Get-ExcelColor 255 255 255)
    $Range.HorizontalAlignment = -4108
    $Range.VerticalAlignment = -4108
    Set-Borders -Range $Range -Color $BorderColor
}

function Format-CellRange {
    param(
        $Range,
        [int]$FillColor,
        [int]$BorderColor,
        [string]$Alignment = "center",
        [string]$NumberFormat = "",
        [bool]$Bold = $false,
        [int]$FontColor = -1
    )

    $Range.Interior.Color = $FillColor
    $Range.Font.Name = "Aptos"
    $Range.Font.Size = 10
    $Range.Font.Bold = $Bold
    $Range.Font.Color = $(if ($FontColor -ge 0) { $FontColor } else { Get-ExcelColor 21 31 44 })
    $Range.VerticalAlignment = -4108
    switch ($Alignment.ToLowerInvariant()) {
        "left" { $Range.HorizontalAlignment = -4131 }
        "right" { $Range.HorizontalAlignment = -4152 }
        default { $Range.HorizontalAlignment = -4108 }
    }
    if ($NumberFormat) {
        $Range.NumberFormat = $NumberFormat
    }
    Set-Borders -Range $Range -Color $BorderColor
}

function Apply-StripeBlock {
    param(
        $Sheet,
        [string]$Address,
        [string]$RowLabelRange,
        [string[]]$IntegerRanges = @(),
        [string[]]$PercentRanges = @(),
        [string[]]$DecimalRanges = @(),
        [int[]]$HighlightRows = @()
    )

    $range = $Sheet.Range($Address)
    $firstRow = $range.Row
    $lastRow = $firstRow + $range.Rows.Count - 1
    $startCol = $range.Column
    $lastCol = $startCol + $range.Columns.Count - 1
    $border = (Get-ExcelColor 206 214 224)
    $rowA = (Get-ExcelColor 249 250 251)
    $rowB = (Get-ExcelColor 239 244 248)
    $highlight = (Get-ExcelColor 255 247 214)

    for ($row = $firstRow; $row -le $lastRow; $row++) {
        $rowRange = $Sheet.Range($Sheet.Cells.Item($row, $startCol), $Sheet.Cells.Item($row, $lastCol))
        $fill = if ($HighlightRows -contains $row) { $highlight } elseif (($row - $firstRow) % 2 -eq 0) { $rowA } else { $rowB }
        Format-CellRange -Range $rowRange -FillColor $fill -BorderColor $border -Bold ($HighlightRows -contains $row)
    }

    if ($RowLabelRange) {
        $labelRange = $Sheet.Range($RowLabelRange)
        for ($row = $labelRange.Row; $row -lt ($labelRange.Row + $labelRange.Rows.Count); $row++) {
            $cell = $Sheet.Cells.Item($row, $labelRange.Column)
            $fill = if ($HighlightRows -contains $row) { $highlight } elseif (($row - $firstRow) % 2 -eq 0) { $rowA } else { $rowB }
            Format-CellRange -Range $cell -FillColor $fill -BorderColor $border -Alignment "left" -Bold ($HighlightRows -contains $row)
        }
    }

    for ($row = $firstRow; $row -le $lastRow; $row++) {
        $fill = if ($HighlightRows -contains $row) { $highlight } elseif (($row - $firstRow) % 2 -eq 0) { $rowA } else { $rowB }
        foreach ($block in $IntegerRanges) {
            $blockRange = $Sheet.Range($block)
            $blockFirstRow = $blockRange.Row
            $blockLastRow = $blockFirstRow + $blockRange.Rows.Count - 1
            if ($row -ge $blockFirstRow -and $row -le $blockLastRow) {
                $blockStartCol = $blockRange.Column
                $blockEndCol = $blockStartCol + $blockRange.Columns.Count - 1
                $rowRange = $Sheet.Range($Sheet.Cells.Item($row, $blockStartCol), $Sheet.Cells.Item($row, $blockEndCol))
                Format-CellRange -Range $rowRange -FillColor $fill -BorderColor $border -NumberFormat "#,##0" -Bold ($HighlightRows -contains $row)
            }
        }
        foreach ($block in $PercentRanges) {
            $blockRange = $Sheet.Range($block)
            $blockFirstRow = $blockRange.Row
            $blockLastRow = $blockFirstRow + $blockRange.Rows.Count - 1
            if ($row -ge $blockFirstRow -and $row -le $blockLastRow) {
                $blockStartCol = $blockRange.Column
                $blockEndCol = $blockStartCol + $blockRange.Columns.Count - 1
                $rowRange = $Sheet.Range($Sheet.Cells.Item($row, $blockStartCol), $Sheet.Cells.Item($row, $blockEndCol))
                Format-CellRange -Range $rowRange -FillColor $fill -BorderColor $border -NumberFormat "0.0%" -Bold ($HighlightRows -contains $row)
            }
        }
        foreach ($block in $DecimalRanges) {
            $blockRange = $Sheet.Range($block)
            $blockFirstRow = $blockRange.Row
            $blockLastRow = $blockFirstRow + $blockRange.Rows.Count - 1
            if ($row -ge $blockFirstRow -and $row -le $blockLastRow) {
                $blockStartCol = $blockRange.Column
                $blockEndCol = $blockStartCol + $blockRange.Columns.Count - 1
                $rowRange = $Sheet.Range($Sheet.Cells.Item($row, $blockStartCol), $Sheet.Cells.Item($row, $blockEndCol))
                Format-CellRange -Range $rowRange -FillColor $fill -BorderColor $border -NumberFormat "0.0" -Bold ($HighlightRows -contains $row)
            }
        }
    }
}

$resolvedWorkbook = (Resolve-Path -LiteralPath $WorkbookPath).ProviderPath
if (-not $BackupDirectory) {
    $BackupDirectory = Join-Path (Split-Path -Parent $resolvedWorkbook) "Backups"
}
$resolvedBackup = [System.IO.Path]::GetFullPath($BackupDirectory)
New-Item -ItemType Directory -Path $resolvedBackup -Force | Out-Null

$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$backupPath = Join-Path $resolvedBackup ("PPT presentation source data - before native formatting $timestamp.xlsx")
Copy-Item -LiteralPath $resolvedWorkbook -Destination $backupPath -Force

$accentRed = Get-ExcelColor 239 68 68
$accentCyan = Get-ExcelColor 0 183 255
$accentAmber = Get-ExcelColor 245 158 11
$accentGreen = Get-ExcelColor 16 185 129
$accentCoral = Get-ExcelColor 255 107 116
$accentRose = Get-ExcelColor 225 29 72
$darkHeader = Get-ExcelColor 17 24 39
$borderDark = Get-ExcelColor 55 65 81

$excel = $null
$workbook = $null
$sheet = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false

    $workbook = $excel.Workbooks.Open($resolvedWorkbook)
    $sheet = $workbook.Worksheets.Item($SheetName)
    $sheet.Activate() | Out-Null
    $excel.ActiveWindow.DisplayGridlines = $false
    $excel.ActiveWindow.Zoom = 92
    $excel.ActiveWindow.SplitRow = 2
    $excel.ActiveWindow.SplitColumn = 0
    $excel.ActiveWindow.FreezePanes = $true

    $sheet.Cells.Font.Name = "Aptos"
    $sheet.Cells.Font.Size = 10
    $sheet.Rows("1:1").RowHeight = 26
    $sheet.Rows("2:2").RowHeight = 22
    $sheet.Rows("3:15").RowHeight = 20
    $sheet.Rows("17:18").RowHeight = 22
    $sheet.Rows("19:31").RowHeight = 20
    $sheet.Tab.Color = $accentRed

    $widths = @{
        "A" = 16; "B" = 11; "C" = 11; "D" = 3;  "E" = 3;
        "F" = 11; "G" = 14; "H" = 3;
        "I" = 11; "J" = 14; "K" = 3;
        "L" = 17; "M" = 10; "N" = 12; "O" = 3;
        "P" = 8;  "Q" = 10; "R" = 12; "S" = 3;
        "T" = 15; "U" = 11; "V" = 11; "W" = 3;
        "X" = 10; "Y" = 12; "Z" = 12; "AA" = 12
    }
    foreach ($column in $widths.Keys) {
        $sheet.Columns.Item($column).ColumnWidth = $widths[$column]
    }

    foreach ($spacer in "D:E", "H:H", "K:K", "O:O", "S:S", "W:W") {
        $spacerRange = $sheet.Range($spacer)
        $spacerRange.Interior.Color = (Get-ExcelColor 255 255 255)
        $spacerRange.Borders.LineStyle = 0
    }

    Format-TitleRange -Range $sheet.Range("A1:C1") -FillColor $accentGreen -BorderColor $borderDark
    Format-TitleRange -Range $sheet.Range("F1:G1") -FillColor $accentCyan -BorderColor $borderDark
    Format-TitleRange -Range $sheet.Range("I1:J1") -FillColor $accentCyan -BorderColor $borderDark
    Format-TitleRange -Range $sheet.Range("L1:N1") -FillColor $accentAmber -BorderColor $borderDark
    Format-TitleRange -Range $sheet.Range("P1:R1") -FillColor $accentRed -BorderColor $borderDark
    Format-TitleRange -Range $sheet.Range("T1:V1") -FillColor $accentCoral -BorderColor $borderDark
    Format-TitleRange -Range $sheet.Range("X1:AA1") -FillColor $accentRose -BorderColor $borderDark
    Format-TitleRange -Range $sheet.Range("A17:H17") -FillColor $darkHeader -BorderColor $borderDark

    foreach ($headerRange in "A2:C2", "F2:G2", "I2:J2", "L2:N2", "P2:R2", "T2:V2", "X2:AA2", "A18:H18") {
        Format-HeaderRange -Range $sheet.Range($headerRange) -FillColor $darkHeader -BorderColor $borderDark
    }

    Apply-StripeBlock -Sheet $sheet -Address "A3:C15" -RowLabelRange "A3:A15" -PercentRanges @("B3:C15") -HighlightRows @(3)
    Apply-StripeBlock -Sheet $sheet -Address "F3:G15" -RowLabelRange "F3:F15" -PercentRanges @("G3:G15")
    Apply-StripeBlock -Sheet $sheet -Address "I3:J15" -RowLabelRange "I3:I15" -PercentRanges @("J3:J15")
    Apply-StripeBlock -Sheet $sheet -Address "M3:N15" -RowLabelRange "M3:M15" -PercentRanges @("N3:N15")
    Apply-StripeBlock -Sheet $sheet -Address "Q3:R15" -RowLabelRange "Q3:Q15" -PercentRanges @("R3:R15")
    Apply-StripeBlock -Sheet $sheet -Address "T3:V7" -RowLabelRange "T3:T7" -IntegerRanges @("U3:U7") -PercentRanges @("V3:V7")
    Apply-StripeBlock -Sheet $sheet -Address "X3:AA14" -RowLabelRange "X3:X14" -IntegerRanges @("Y3:AA14")
    Apply-StripeBlock -Sheet $sheet -Address "A19:H31" -RowLabelRange "A19:A31" -DecimalRanges @("B19:H31") -HighlightRows @(31)

    Format-CellRange -Range $sheet.Range("A3") -FillColor $darkHeader -BorderColor $borderDark -Alignment "left" -Bold $true -FontColor (Get-ExcelColor 255 255 255)
    Format-CellRange -Range $sheet.Range("A31") -FillColor $darkHeader -BorderColor $borderDark -Alignment "left" -Bold $true -FontColor (Get-ExcelColor 255 255 255)

    $sheet.Range("A1").Select() | Out-Null
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

Write-Output "Styled workbook saved to $resolvedWorkbook"
Write-Output "Backup created at $backupPath"
