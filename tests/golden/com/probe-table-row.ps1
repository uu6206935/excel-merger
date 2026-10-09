# Probe: Excel's reference adjustment when inserting a sheet row at / inside / below a table (vs a plain range).
# Saves tr-<table|plain>-<row>.xlsx; sheet O holds formulas over T.
param([string]$OutDir)
$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\common.ps1"
$OutDir = (Resolve-Path $OutDir).Path
$M = $script:GMissing
function Build($xl, [bool]$withTable) {
  $wb = $xl.Workbooks.Add(); $ws = $wb.Worksheets.Item(1); $ws.Name = 'T'
  $o = $wb.Worksheets.Add([Type]::Missing, $ws); $o.Name = 'O'
  SetF $ws.Range('B3') 'k1'; SetF $ws.Range('C3') 'k2'; SetF $ws.Range('D3') 'k3'; SetF $ws.Range('E3') 'k4'
  for ($r = 4; $r -le 10; $r++) { for ($c = 1; $c -le 8; $c++) { SetF $ws.Cells.Item($r, $c) ($r * 10 + $c) } }
  $ws.Range('B3:E3').Font.Bold = $true; $ws.Range('B3:E3').Interior.Color = 49407
  $ws.Range('C4:C10').NumberFormat = '0.0'
  if ($withTable) { $lo = $ws.ListObjects.Add(1, $ws.Range('B3:E10'), $M, 1); $lo.Name = 'tb' }
  $f = @('SUM(T!C4:C10)', 'SUM(T!C4:C8)', 'SUM(T!H4:H10)', 'SUM(T!C4:C12)', 'SUM(T!A4:A10)', 'T!C4', 'SUM(T!C5:C10)', 'SUM(T!B3:E10)', 'SUM(T!B4:E10)', 'SUM(T!C4:D4)', 'SUM(T!A4:H10)')
  for ($i = 0; $i -lt $f.Count; $i++) { $o.Cells.Item($i + 1, 1).Formula = '=' + $f[$i]; SetF $o.Cells.Item($i + 1, 2) $f[$i] }
  return $wb
}
$xl = $null
try {
  $xl = New-GExcel
  foreach ($tb in @($true, $false)) {
    foreach ($at in @(4, 5, 11)) {
      $wb = @(Build $xl $tb)[-1]
      $ws = $wb.Worksheets.Item('T')
      [void]$ws.Range("$($at):$($at)").EntireRow.Insert()
      $name = 'tr-' + $(if ($tb) { 'table' } else { 'plain' }) + "-$at.xlsx"
      $p = Join-Path $OutDir $name
      if (Test-Path $p) { Remove-Item $p -Force }
      $wb.SaveAs($p, 51); $wb.Close($false)
    }
  }
  Write-Host 'saved'
}
finally { Close-GExcel $xl }
