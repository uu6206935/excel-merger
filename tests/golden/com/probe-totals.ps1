# Probe (owner: patcher P): what Excel stores when a formula / constant is TYPED into a table's totals row
# (r2-table.xlsx: Tbl A1:C5 with totals row 5). Saves <Out>\totals-<n>.xlsx per case and prints TotalsCalculation.
#   powershell -NoProfile -ExecutionPolicy Bypass -File tests\golden\com\probe-totals.ps1 -Fixture tests\fixtures\golden\r2-table.xlsx -Out <dir>
# stdout: ASCII JSON lines.
param([string]$Fixture, [string]$Out)
$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\common.ps1"
$Fixture = (Resolve-Path $Fixture).Path
$Out = (Resolve-Path $Out).Path
$xl = $null
try {
  $xl = New-GExcel
  $cases = @(
    @{ n = 1; cell = 'A5'; f = '=SUBTOTAL(109,Tbl[Score])' },
    @{ n = 2; cell = 'B5'; f = '=SUBTOTAL(101,Tbl[Score])' },
    @{ n = 3; cell = 'C5'; f = '=SUBTOTAL(103,Tbl[Twice])' },
    @{ n = 4; cell = 'B5'; f = '=1+1' },
    @{ n = 5; cell = 'B5'; f = '=SUM(B2:B4)' },
    @{ n = 6; cell = 'B5'; f = '=SUBTOTAL(109,Tbl[Twice])' },
    @{ n = 7; cell = 'A5'; f = '2026/9/30' },
    @{ n = 8; cell = 'A5'; f = 'TRUE' },
    @{ n = 9; cell = 'A5'; f = '=Tbl[[#This Row],[Score]]' },
    @{ n = 10; cell = 'B5'; f = '=MAX(Tbl[Score])' },
    @{ n = 11; cell = 'A5'; f = '=SUBTOTAL(104,Tbl[Score])' },
    @{ n = 12; cell = 'A5'; f = '=SUBTOTAL(102,Tbl[Score])' }
  )
  foreach ($c in $cases) {
    $wb = $xl.Workbooks.Open($Fixture, 0, $true)
    $ws = $wb.Worksheets.Item('T')
    $cell = $ws.Range($c.cell)
    $cell.Formula = $c.f
    $lo = $ws.ListObjects.Item(1)
    $calc = @(); foreach ($lc in $lo.ListColumns) { $calc += ($lc.Name + '=' + $lc.TotalsCalculation) }
    Write-Host ('{"n":' + $c.n + ',"cell":' + (JStr $c.cell) + ',"typed":' + (JStr $c.f) + ',"formula":' + (JStr $cell.Formula) + ',"v2":' + (JStr ([string]$cell.Value2)) + ',"type":' + (JStr ($cell.Value2.GetType().Name)) + ',"nf":' + (JStr $cell.NumberFormat) + ',"calc":' + (JStr ($calc -join '|')) + '}')
    $p = Join-Path $Out ('totals-' + $c.n + '.xlsx')
    if (Test-Path $p) { Remove-Item $p -Force }
    $wb.SaveAs($p, 51)
    $wb.Close($false)
  }
}
finally { Close-GExcel $xl }
