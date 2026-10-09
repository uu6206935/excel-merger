# Probe (owner: patcher P): what Excel stores when values are TYPED into General cells (Range.Formula = text is
# what the formula bar does): number format, column width; also typing over table headers (number / formula / blank).
#   powershell -NoProfile -ExecutionPolicy Bypass -File tests\golden\com\probe-typed.ps1 -Out <dir>
# stdout: ASCII JSON lines; saves <dir>\typed.xlsx (inspect numFmtId / cols).
param([string]$Out, [string]$Table = '')
$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\common.ps1"
$Out = (Resolve-Path $Out).Path
$xl = $null
try {
  $xl = New-GExcel
  $wb = $xl.Workbooks.Add()
  $ws = $wb.Worksheets.Item(1)
  $inputs = @('2026/9/30', '50%', ([regex]::Unescape('¥') + '1,000'), '1,234', '9:30', '2026/9/30 9:30', '12.5%', '$1,000', '3/4', '1.5E+3', ([regex]::Unescape('￥') + '2,000'))
  for ($i = 0; $i -lt $inputs.Count; $i++) {
    $c = $ws.Cells.Item($i + 1, 1)
    $c.Formula = $inputs[$i]
    Write-Host ('{"in":' + (JStr $inputs[$i]) + ',"nf":' + (JStr $c.NumberFormat) + ',"nfl":' + (JStr $c.NumberFormatLocal) + ',"v":' + (JStr ([string]$c.Value2)) + ',"text":' + (JStr $c.Text) + ',"w":' + (JStr ([string]$ws.Columns.Item(1).ColumnWidth)) + '}')
  }
  # same values written through Value2 + NumberFormat (the Excel runner's way): does the width change?
  $ws.Cells.Item(1, 3).Value2 = 46295
  $ws.Cells.Item(1, 3).NumberFormat = 'yyyy/m/d'
  Write-Host ('{"runner":"C1","w":' + (JStr ([string]$ws.Columns.Item(3).ColumnWidth)) + ',"text":' + (JStr $ws.Cells.Item(1, 3).Text) + '}')
  $p = Join-Path $Out 'typed.xlsx'
  if (Test-Path $p) { Remove-Item $p -Force }
  $wb.SaveAs($p, 51)
  $wb.Close($false)
  if ($Table) {
    # table header typing: number, formula, blank (g-table.xlsx tblMain header row 3)
    foreach ($case in @(@('A3', '2024'), @('D3', '=1+1'), @('E3', ''), @('B3', '2026/9/30'))) {
      $wb = $xl.Workbooks.Open($Table, 0, $true)
      $ws = $wb.Worksheets.Item(1)
      $c = $ws.Range($case[0])
      if ($case[1] -eq '') { [void]$c.ClearContents() } else { $c.Formula = $case[1] }
      $lo = $ws.ListObjects.Item(1)
      $names = @(); foreach ($lc in $lo.ListColumns) { $names += $lc.Name }
      Write-Host ('{"hdr":' + (JStr $case[0]) + ',"typed":' + (JStr $case[1]) + ',"formula":' + (JStr $c.Formula) + ',"v2type":' + (JStr ($c.Value2.GetType().Name)) + ',"text":' + (JStr $c.Text) + ',"cols":' + (JStr ($names -join '|')) + ',"E15":' + (JStr $ws.Range('E15').Formula) + '}')
      $wb.Close($false)
    }
  }
}
finally { Close-GExcel $xl }
