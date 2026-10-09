# Round-2 fixtures (owner: patcher P), authored by Excel itself:
#   powershell -NoProfile -ExecutionPolicy Bypass -File tests\golden\com\make-fixtures-r2.ps1 -OutDir tests\fixtures\golden
#  r2-arrays.xlsx : sheet A: A2:A4 = 1..3; B2:B4 legacy CSE array {=A2:A4*10}; D2 =SEQUENCE(3,2) (spills D2:E4);
#                   G2 =A2:A4&"x" (string spill G2:G4); I1:I4 General cells; K2 = SUM(B2:B4); sheet B refers to A.
#  r2-rich.xlsx   : sheet S (a premerge source): header ID / NAME / MEMO, rows 2..3; B2 has rich text runs
#                   (bold red first run), B3 has a phonetic reading (furigana), C2 plain text.
#  r2-table.xlsx  : sheet T: ListObject Tbl A1:C5 (ID / Score / Twice, calculated column Twice = [@Score]*2, totals
#                   row 5 with a custom formula =SUM(Tbl[Score])*2 under Score and SUBTOTAL under Twice); workbook name
#                   ScoreCol = Tbl[Score]; sheet U: =SUM(Tbl[Score]), =Tbl[[#Headers],[Score]], =SUM(Tbl[[Score]:[Twice]]),
#                   =COUNT(ScoreCol), =Tbl[@Score] cannot live outside the table (not used).
# -Only rich|arrays|table builds just that file.
# Output (ASCII stdout): one line per saved file.
param([string]$OutDir, [string]$Only = '')
$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\common.ps1"
$OutDir = (Resolve-Path $OutDir).Path

function Save-Book($wb, [string]$name) {
  $p = Join-Path $OutDir $name
  if (Test-Path $p) { Remove-Item $p -Force }
  $wb.SaveAs($p, 51)
  $wb.Close($false)
  Write-Host "saved $name"
}

$xl = $null
try {
  $xl = New-GExcel
  Write-Host "excel pid $($script:GExcelPid)"

  # ---- r2-arrays.xlsx ---------------------------------------------------------------------------
  if (-not $Only -or $Only -eq 'arrays') {
  $wb = $xl.Workbooks.Add()
  while ($wb.Worksheets.Count -lt 2) { [void]$wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)) }
  $ws = $wb.Worksheets.Item(1); $ws.Name = 'A'
  $wb.Worksheets.Item(2).Name = 'B'
  SetF $ws.Range('A1') 'val'
  for ($i = 1; $i -le 3; $i++) { SetF $ws.Cells.Item($i + 1, 1) $i }
  $ws.Range('B2:B4').FormulaArray = '=A2:A4*10'
  $ws.Range('D2').Formula2 = '=SEQUENCE(3,2)'
  $ws.Range('G2').Formula2 = '=A2:A4&"x"'
  SetF $ws.Range('I1') 'typed'
  $ws.Range('K2').Formula = '=SUM(B2:B4)'
  $ws.Range('K3').Formula = '=SUM(D2#)'
  $wb.Worksheets.Item(2).Range('A1').Formula = '=A!B3+A!E4'
  $wb.Worksheets.Item(2).Range('A2').Formula2 = '=A!G2#'
  Save-Book $wb 'r2-arrays.xlsx'
  }

  # ---- r2-rich.xlsx -------------------------------------------------------------------------------
  if (-not $Only -or $Only -eq 'rich') {
  $wb = $xl.Workbooks.Add()
  $ws = $wb.Worksheets.Item(1); $ws.Name = 'S'
  SetF $ws.Range('A1') 'ID'; SetF $ws.Range('B1') 'NAME'; SetF $ws.Range('C1') 'MEMO'
  SetF $ws.Range('A2') 'S001'; SetF $ws.Range('A3') 'S002'
  $ws.Range('B2').Value2 = 'Bold normal'
  $ch = $ws.Range('B2').Characters(1, 4)
  $ch.Font.Bold = $true
  $ch.Font.Color = 255
  $ws.Range('B3').Value2 = [regex]::Unescape('山田')          # yamada (kanji)
  [void]$ws.Range('B3').Phonetics.Add(1, 2, [regex]::Unescape('ヤマダ'))  # katakana reading
  $ws.Range('B3').Phonetic.Visible = $true
  SetF $ws.Range('C2') 'plain'
  SetF $ws.Range('C3') 'x'
  Save-Book $wb 'r2-rich.xlsx'
  }

  # ---- r2-table.xlsx ------------------------------------------------------------------------------
  if (-not $Only -or $Only -eq 'table') {
  $wb = $xl.Workbooks.Add()
  while ($wb.Worksheets.Count -lt 2) { [void]$wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)) }
  $ws = $wb.Worksheets.Item(1); $ws.Name = 'T'
  $wb.Worksheets.Item(2).Name = 'U'
  SetF $ws.Range('A1') 'ID'; SetF $ws.Range('B1') 'Score'; SetF $ws.Range('C1') 'Twice'
  for ($i = 1; $i -le 3; $i++) { SetF $ws.Cells.Item($i + 1, 1) ('S00' + $i); SetF $ws.Cells.Item($i + 1, 2) ($i * 10) }
  $lo = $ws.ListObjects.Add(1, $ws.Range('A1:C4'), [Type]::Missing, 1)
  $lo.Name = 'Tbl'
  $ws.Range('C2').Formula = '=Tbl[[#This Row],[Score]]*2'   # fills the calculated column
  $lo.ShowTotals = $true                       # totals row 5
  $ws.Range('B5').Formula = '=SUM(Tbl[Score])*2'
  $lo.ListColumns.Item(3).TotalsCalculation = 1   # xlTotalsCalculationSum
  [void]$wb.Names.Add('ScoreCol', '=Tbl[Score]')
  $u = $wb.Worksheets.Item(2)
  $u.Range('A1').Formula = '=SUM(Tbl[Score])'
  $u.Range('A2').Formula = '=Tbl[[#Headers],[Score]]'
  $u.Range('A3').Formula = '=SUM(Tbl[[Score]:[Twice]])'
  $u.Range('A4').Formula = '=COUNT(ScoreCol)'
  $u.Range('A5').Formula = '=SUM(Tbl[Twice])'
  Save-Book $wb 'r2-table.xlsx'
  }
}
finally { Close-GExcel $xl }
