# Probe: which format does a cell inserted by Excel (Format Same As Left / Above) get?
# Builds rows of (left cell B, right cell C) border / fill combinations on sheet Cols and (upper cell, lower cell)
# combinations on sheet Rows, saves before.xlsx, inserts a column at C / a row, saves after.xlsx.
param([string]$OutDir)
$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\common.ps1"
$OutDir = (Resolve-Path $OutDir).Path

# side codes: L=7 T=8 B=9 R=10 ; spec = "sides:style" e.g. "LTBR:1/2" (linestyle/weight)
function ApplyB($rng, [string]$sides, [int]$ls, [int]$wt, [int]$color) {
  $map = @{ 'L' = 7; 'T' = 8; 'B' = 9; 'R' = 10 }
  foreach ($ch in $sides.ToCharArray()) { $b = $rng.Borders.Item($map[[string]$ch]); $b.LineStyle = $ls; $b.Weight = $wt; if ($color -ge 0) { $b.Color = $color } }
}

$combos = @(
  @('LTBR', '', 0),      # 1 left all thin, right none
  @('', 'LTBR', 0),      # 2
  @('T', 'T', 0),        # 3
  @('T', 'B', 0),        # 4
  @('LTBR', 'LTBR', 0),  # 5
  @('LTBR', 'TB', 0),    # 6
  @('LR', 'LR', 0),      # 7
  @('R', 'L', 0),        # 8
  @('TB', 'TB', 1),      # 9  left medium, right thin
  @('LTBR', '', 2),      # 10 left has fill
  @('B', 'B', 3),        # 11 right has fill
  @('TB', 'LTBR', 0),    # 12
  @('L', '', 0),         # 13
  @('R', '', 0),         # 14
  @('', 'R', 0),         # 15
  @('TB', 'T', 0)        # 16
)

$xl = $null
try {
  $xl = New-GExcel
  $wb = $xl.Workbooks.Add()
  while ($wb.Worksheets.Count -lt 2) { [void]$wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)) }
  $wc = $wb.Worksheets.Item(1); $wc.Name = 'Cols'
  $wr = $wb.Worksheets.Item(2); $wr.Name = 'Rows'
  for ($i = 0; $i -lt $combos.Count; $i++) {
    $r = 2 + 2 * $i
    $l = $combos[$i][0]; $rt = $combos[$i][1]; $flag = $combos[$i][2]
    SetF $wc.Cells.Item($r, 1) ($i + 1)
    $lw = 2; if ($flag -eq 1) { $lw = -4138 }
    if ($l) { ApplyB $wc.Cells.Item($r, 2) $l 1 $lw -1 }
    if ($rt) { ApplyB $wc.Cells.Item($r, 3) $rt 1 2 -1 }
    if ($flag -eq 2) { $wc.Cells.Item($r, 2).Interior.Color = 0x00FFFF }
    if ($flag -eq 3) { $wc.Cells.Item($r, 3).Interior.Color = 0x00FFFF }
    SetF $wc.Cells.Item($r, 2) 'L'; SetF $wc.Cells.Item($r, 3) 'R'
    # rows sheet: upper cell row 2, lower cell row 3, in column 2+2i
    $c = 2 + 2 * $i
    SetF $wr.Cells.Item(1, $c) ($i + 1)
    if ($l) { ApplyB $wr.Cells.Item(2, $c) $l 1 $lw -1 }
    if ($rt) { ApplyB $wr.Cells.Item(3, $c) $rt 1 2 -1 }
    if ($flag -eq 2) { $wr.Cells.Item(2, $c).Interior.Color = 0x00FFFF }
    if ($flag -eq 3) { $wr.Cells.Item(3, $c).Interior.Color = 0x00FFFF }
    SetF $wr.Cells.Item(2, $c) 'U'; SetF $wr.Cells.Item(3, $c) 'D'
  }
  $wr.Rows.Item(2).RowHeight = 30
  $wc.Columns.Item(2).ColumnWidth = 15
  $b = Join-Path $OutDir 'probe-before.xlsx'
  if (Test-Path $b) { Remove-Item $b -Force }
  $wb.SaveAs($b, 51)
  [void]$wc.Range('C:C').EntireColumn.Insert()
  [void]$wr.Range('3:3').EntireRow.Insert()
  $a = Join-Path $OutDir 'probe-after.xlsx'
  if (Test-Path $a) { Remove-Item $a -Force }
  $wb.SaveAs($a, 51)
  $wb.Close($false)
  Write-Host 'probe saved'
}
finally { Close-GExcel $xl }
