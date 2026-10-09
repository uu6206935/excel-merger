# Builds Excel-made reader fixtures (owner: reader R) in -Work (ASCII path) and records what Excel shows:
#   legacy.xls       FileFormat 56 (BIFF8): header row, number / date / text-with-leading-zero / formula / bool / error,
#                    merged cell, hidden 2nd sheet
#   password.xlsx    SaveAs with a password (CFB EncryptedPackage)
#   password.xls     BIFF8 with a password (FILEPASS)
#   strict.xlsx      torture.xlsx saved as Strict Open XML (FileFormat 61)
#   date1904.xlsx    1904 date system with a date, a time and a plain number
# Output: <Work>\excel-fixtures.json (Text / Value2 per cell for legacy.xls and date1904.xlsx).
# Run: powershell -NoProfile -ExecutionPolicy Bypass -File make-fixtures.ps1 -Work <dir>
param([Parameter(Mandatory = $true)][string]$Work)
$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\common.ps1"

function Get-Cells($ws, [int]$rows, [int]$cols) {
  $out = @()
  for ($r = 1; $r -le $rows; $r++) {
    for ($c = 1; $c -le $cols; $c++) {
      $cell = $ws.Cells.Item($r, $c)
      $v2 = $cell.Value2
      if ($null -eq $v2) { continue }
      $vt = if ($v2 -is [double]) { 'n' } elseif ($v2 -is [bool]) { 'b' } elseif ($v2 -is [int]) { 'e' } else { 's' }
      $out += [ordered]@{ addr = [string]$cell.Address($false, $false); text = [string]$cell.Text; vt = $vt; v = $(if ($vt -eq 'e') { [int]$v2 } else { $v2 }); nf = [string]$cell.NumberFormat }
    }
  }
  return , $out
}

$result = [ordered]@{}
$xl = $null
try {
  $xl = New-RdrExcel

  # ---------------- legacy.xls
  $wb = $xl.Workbooks.Add()
  $ws = $wb.Worksheets.Item(1)
  $ws.Name = 'Data'
  $ws.Range('A1').Formula = 'ID'
  $ws.Range('B1').Formula = '氏名'
  $ws.Range('C1').Formula = '日付'
  $ws.Range('D1').Formula = '金額'
  $ws.Range('E1').Formula = 'コード'
  $ws.Range('F1').Formula = '計算'
  $ws.Range('G1').Formula = '判定'
  $names = @('山田 太郎', '佐藤 花子', '髙橋 一郎')
  for ($i = 0; $i -lt 3; $i++) {
    $r = $i + 2
    $ws.Range("A$r").Formula = [string](1001 + $i)
    $ws.Range("B$r").Formula = $names[$i]
    $ws.Range("C$r").Formula = ('2026/10/{0}' -f (20 + $i))
    $ws.Range("D$r").NumberFormat = '#,##0.0'
    $ws.Range("D$r").Formula = [string](1234.5 + $i * 1000)
    $ws.Range("E$r").NumberFormat = '@'
    $ws.Range("E$r").Formula = ('0012{0}' -f $i)
    $ws.Range("F$r").Formula = "=D$r*2"
  }
  $ws.Range('G2').Formula = '=TRUE'
  $ws.Range('G3').Formula = '=NA()'
  $ws.Range('G4').Formula = '=1/0'
  $ws.Range('A6').Formula = '結合セル'
  $ws.Range('A6:C6').Merge()
  $ws.Range('C2:C4').NumberFormatLocal = 'yyyy/m/d'
  $ws.Columns.Item(2).ColumnWidth = 14
  $ws2 = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $ws)
  $ws2.Name = '隠し'
  $ws2.Range('A1').Formula = 'secret'
  $ws2.Visible = 0
  $p = Join-Path $Work 'legacy.xls'
  if (Test-Path $p) { Remove-Item $p -Force }
  $wb.SaveAs($p, 56)
  $wb.Close($false)
  $wb = $xl.Workbooks.Open($p, 0, $true)
  $result.legacyXls = Get-Cells $wb.Worksheets.Item('Data') 8 8
  $wb.Close($false)
  Write-Host 'OK legacy.xls'

  # ---------------- password.xlsx / password.xls
  foreach ($fmt in @(@{ n = 'password.xlsx'; f = 51 }, @{ n = 'password.xls'; f = 56 })) {
    $wb = $xl.Workbooks.Add()
    $wb.Worksheets.Item(1).Range('A1').Formula = 'secret'
    $p = Join-Path $Work $fmt.n
    if (Test-Path $p) { Remove-Item $p -Force }
    $wb.SaveAs($p, $fmt.f, 'pass1234')
    $wb.Close($false)
    Write-Host ('OK ' + $fmt.n)
  }

  # ---------------- strict.xlsx (from torture.xlsx)
  $src = Join-Path $Work 'torture.xlsx'
  if (Test-Path $src) {
    $wb = $xl.Workbooks.Open($src, 0, $true)
    $p = Join-Path $Work 'strict.xlsx'
    if (Test-Path $p) { Remove-Item $p -Force }
    $wb.SaveAs($p, 61)
    $wb.Close($false)
    Write-Host 'OK strict.xlsx'
  }

  # ---------------- date1904.xlsx
  $wb = $xl.Workbooks.Add()
  $wb.Date1904 = $true
  $ws = $wb.Worksheets.Item(1)
  $ws.Range('A1').Formula = '2026/10/20'
  $ws.Range('A2').Formula = '12:30'
  $ws.Range('A3').Formula = '1234'
  $ws.Range('A4').Formula = '1900/1/1'
  $ws.Range('B1').Formula = '=A1+1'
  $ws.Range('B1').NumberFormatLocal = 'yyyy"年"m"月"d"日"'
  $p = Join-Path $Work 'date1904.xlsx'
  if (Test-Path $p) { Remove-Item $p -Force }
  $wb.SaveAs($p, 51)
  $wb.Close($false)
  $wb = $xl.Workbooks.Open($p, 0, $true)
  $result.date1904 = Get-Cells $wb.Worksheets.Item(1) 5 3
  $wb.Close($false)
  Write-Host 'OK date1904.xlsx'

  Save-Json $result (Join-Path $Work 'excel-fixtures.json')
  Write-Host 'DONE'
}
finally { Close-RdrExcel $xl }
