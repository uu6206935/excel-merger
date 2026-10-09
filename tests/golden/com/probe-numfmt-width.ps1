# Probe (owner: patcher P): when does Range.NumberFormat widen a column (the Excel runner writes Value2 +
# NumberFormat for CellWrite.numFmt), and to which width?
#   powershell -NoProfile -ExecutionPolicy Bypass -File tests\golden\com\probe-numfmt-width.ps1
# stdout: ASCII JSON lines {"v":..,"nf":..,"text":..,"w":ColumnWidth after,"w0":before}
$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\common.ps1"
$xl = $null
try {
  $xl = New-GExcel
  $wb = $xl.Workbooks.Add()
  $ws = $wb.Worksheets.Item(1)
  $cases = @(
    @(46023, 'yyyy/m/d'), @(46295, 'yyyy/m/d'), @(46305, 'yyyy/m/d'), @(46295.3958333333, 'yyyy/m/d h:mm'),
    @(0.5238425925925926, 'h:mm:ss'), @(123456789, '#,##0'), @(12345678, '#,##0'), @(1234567, '#,##0'),
    @(123.456, '0.00%'), @(1234567, '"$"#,##0'), @(46295, 'yyyy"Y"m"M"d"D"'), @(46295, 'ge.m.d'),
    @(1234567.891, '#,##0.00'), @(0.5, '0%'), @(46295, 'm/d'), @(12345678901, '0'), @(46295, 'yyyy-mm-dd'),
    @(111111111, '#,##0'), @(1111111, '#,##0'), @(88888888, '#,##0'), @(-1234567, '#,##0'), @(46295, 'dddd')
  )
  $col = 1
  foreach ($c in $cases) {
    $cell = $ws.Cells.Item(1, $col)
    $w0 = $ws.Columns.Item($col).ColumnWidth
    SetF $cell $c[0]
    $cell.NumberFormat = [string]$c[1]
    $w = $ws.Columns.Item($col).ColumnWidth
    Write-Host ('{"col":' + $col + ',"v":' + (JStr ([string]$c[0])) + ',"nf":' + (JStr $c[1]) + ',"text":' + (JStr $cell.Text) + ',"w0":' + (JStr ([string]$w0)) + ',"w":' + (JStr ([string]$w)) + '}')
    $col++
  }
  # a column with a custom width: does NumberFormat still widen it?
  $ws.Columns.Item($col).ColumnWidth = 5
  $cell = $ws.Cells.Item(1, $col)
  SetF $cell 46295
  $cell.NumberFormat = 'yyyy/m/d'
  Write-Host ('{"custom":' + $col + ',"text":' + (JStr $cell.Text) + ',"w":' + (JStr ([string]$ws.Columns.Item($col).ColumnWidth)) + '}')
  $col++
  # value written AFTER the number format (format first, then Value2)
  $cell = $ws.Cells.Item(1, $col)
  $cell.NumberFormat = 'yyyy/m/d'
  SetF $cell 46295
  Write-Host ('{"fmtFirst":' + $col + ',"text":' + (JStr $cell.Text) + ',"w":' + (JStr ([string]$ws.Columns.Item($col).ColumnWidth)) + '}')
  $col++
  # a cell with a larger font
  $cell = $ws.Cells.Item(1, $col)
  $cell.Font.Size = 16
  SetF $cell 46295
  $cell.NumberFormat = 'yyyy/m/d'
  Write-Host ('{"font16":' + $col + ',"text":' + (JStr $cell.Text) + ',"w":' + (JStr ([string]$ws.Columns.Item($col).ColumnWidth)) + '}')
  $col++
  # AutoFit of a TEXT cell with the same characters (is it the same measure?)
  $cell = $ws.Cells.Item(1, $col)
  $cell.NumberFormat = '@'
  $cell.Value2 = '2026/9/30'
  [void]$ws.Columns.Item($col).AutoFit()
  Write-Host ('{"autofitText":' + $col + ',"w":' + (JStr ([string]$ws.Columns.Item($col).ColumnWidth)) + '}')
  $wb.Close($false)
}
finally { Close-GExcel $xl }
