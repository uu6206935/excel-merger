# Builds the Excel-authored golden / premerge fixtures (owner: patcher P).
#   powershell -NoProfile -ExecutionPolicy Bypass -File tests\golden\com\make-fixtures.ps1 -OutDir tests\fixtures\golden
# Output (ASCII stdout): one line per saved file.
param([string]$OutDir)
$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\common.ps1"
$M = $script:GMissing
$OutDir = (Resolve-Path $OutDir).Path

function Save-Book($wb, [string]$name) {
  $p = Join-Path $OutDir $name
  if (Test-Path $p) { Remove-Item $p -Force }
  $wb.SaveAs($p, 51)
  $wb.Close($false)
  Write-Host "saved $name"
}

function U([string]$s) { return [regex]::Unescape($s) }   # write Japanese text from \u escapes (ASCII-only script)

$xl = $null
try {
  $xl = New-GExcel
  Write-Host "excel pid $($script:GExcelPid)"

  # =============================================================================================
  # g-table.xlsx : sheet MEIBO (U+540D U+7C3F) with ListObject tblMain A3:F13 + total row 15
  # =============================================================================================
  $wb = $xl.Workbooks.Add()
  while ($wb.Worksheets.Count -lt 3) { [void]$wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)) }
  $ws = $wb.Worksheets.Item(1); $ws.Name = (U '名簿')          # 名簿
  $sm = $wb.Worksheets.Item(2); $sm.Name = 'Summary'
  $ls = $wb.Worksheets.Item(3); $ls.Name = 'Lists'
  $unis = @((U '東大'), (U '京大'), 'SFC')               # 東大 京大 SFC
  for ($i = 0; $i -lt 3; $i++) { SetF $ls.Cells.Item($i + 1, 1) $unis[$i] }

  SetF $ws.Range('A1') (U '成績一覧')                   # 成績一覧
  $ws.Range('A1').Font.Bold = $true; $ws.Range('A1').Font.Size = 14
  $hdr = @('ID', (U '氏名'), (U '大学'), (U '点数1'), (U '点数2'), (U '合計'))  # ID 氏名 大学 点数1 点数2 合計
  for ($c = 0; $c -lt 6; $c++) { SetF $ws.Cells.Item(3, $c + 1) $hdr[$c] }
  $names = @('A1','B2','C3','D4','E5','F6','G7','H8','I9','J10')
  for ($i = 0; $i -lt 10; $i++) {
    $r = 4 + $i
    SetF $ws.Cells.Item($r, 1) (1001 + $i)
    SetF $ws.Cells.Item($r, 2) ((U '学生') + $names[$i])       # 学生..
    SetF $ws.Cells.Item($r, 3) $unis[$i % 3]
    SetF $ws.Cells.Item($r, 4) (50 + ($i * 7) % 50)
    SetF $ws.Cells.Item($r, 5) (40 + ($i * 11) % 60)
  }
  $lo = $ws.ListObjects.Add(1, $ws.Range('A3:F13'), $M, 1)
  $lo.Name = 'tblMain'
  $ws.Range('F4').Formula = '=[@' + (U '点数1') + ']+[@' + (U '点数2') + ']'
  $ws.Range('D4:E13').NumberFormat = '0"' + (U '点') + '"'
  # total row below the table (not a ListObject totals row)
  SetF $ws.Range('B15') (U '合計')
  $ws.Range('D15').Formula = '=SUM(D4:D13)'
  $ws.Range('E15').Formula = '=SUM(tblMain[' + (U '点数2') + '])'
  $ws.Range('F15').Formula = '=SUBTOTAL(9,F4:F13)'
  $ws.Range('G4').Formula = '=D4*2'; $ws.Range('G4:G13').FillDown()     # shared formulas right of the table
  $ws.Range('B15:F15').Font.Bold = $true
  $ws.Range('B15:F15').Borders.Item(8).LineStyle = 1
  # CF: relative expression over the table rows, cellIs on E
  $ws.Activate(); $ws.Range('A4').Select()
  $fc = $ws.Range('A4:F13').FormatConditions.Add(2, $M, '=$D4>=80'); $fc.Interior.Color = 0xCCFFCC
  $ws.Range('E4').Select()
  $fc2 = $ws.Range('E4:E13').FormatConditions.Add(1, 5, '=90'); $fc2.Font.Bold = $true
  $ws.Range('G4').Select()
  $fc3 = $ws.Range('G4:G13').FormatConditions.Add(2, $M, '=G4>F4'); $fc3.Font.Color = 0x0000FF
  # DV: list from another sheet (x14), whole numbers on points
  $ws.Range('C4:C13').Validation.Add(3, 1, 1, '=Lists!$A$1:$A$3')
  $ws.Range('D4:E13').Validation.Add(1, 1, 1, '0', '100')
  # names
  [void]$wb.Names.Add((U '点数'), ('=' + $ws.Name + '!$D$4:$E$13'))        # 点数
  [void]$wb.Names.Add('AllCols', ('=' + $ws.Name + '!$A:$F'))
  [void]$ws.Names.Add('Local', ('=' + $ws.Name + '!$B$4:$B$13'))
  $ws.Range('A4').Select()
  [void]$wb.Names.Add('NextCell', ('=' + $ws.Name + '!B4'))                          # relative name
  # freeze B4, hidden H, grouped J:K, widths, merge H1:K1, note
  $ws.Range('B4').Select(); $xl.ActiveWindow.FreezePanes = $true
  $ws.Columns.Item(8).Hidden = $true
  [void]$ws.Range('J:K').Columns.Group()
  $ws.Columns.Item(1).ColumnWidth = 6; $ws.Columns.Item(2).ColumnWidth = 14; $ws.Columns.Item(3).ColumnWidth = 9
  $ws.Columns.Item(4).ColumnWidth = 8; $ws.Columns.Item(5).ColumnWidth = 8.5; $ws.Columns.Item(6).ColumnWidth = 10
  $ws.Columns.Item(9).ColumnWidth = 16
  SetF $ws.Range('I1') 'memo'
  $ws.Range('I1:K1').Merge()
  SetF $ws.Range('I4') 'x'; SetF $ws.Range('J4') 1; SetF $ws.Range('K4') 2
  [void]$ws.Range('B5').AddComment('note')
  $ws.Rows.Item(3).RowHeight = 24
  [void]$ws.Hyperlinks.Add($ws.Range('B16'), '', 'Summary!A1', $M, 'to summary')
  # Summary sheet: references into the table sheet
  $q = "'" + $ws.Name + "'!"
  $sm.Range('A1').Formula = '=SUM(' + $ws.Name + '!D4:D13)'
  $sm.Range('A2').Formula = '=AVERAGE(tblMain[' + (U '点数1') + '])'
  $sm.Range('A3').Formula = '=VLOOKUP(1003,' + $ws.Name + '!A4:F13,6,FALSE)'
  $sm.Range('A4').Formula = '=COUNTIF(' + $ws.Name + '!C:C,"SFC")'
  $sm.Range('A5').Formula = '=SUM(' + (U '点数') + ')'
  $sm.Range('A6').Formula = '=' + $ws.Name + '!F15'
  $sm.Range('A7').Formula = '=SUM(' + $ws.Name + '!E:E)'
  $sm.Range('A8').Formula = '=' + $ws.Name + '!$C$4&' + $ws.Name + '!H4'
  $sm.Range('A9').Formula = '=SUM(tblMain[[#This Row],[ID]:[' + (U '点数1') + ']])'
  $sm.Range('A10').Formula = '=SUM(' + $ws.Name + '!D4:' + $ws.Name + '!E13)'
  $sm.Range('A11').Formula = '=ROWS(tblMain)'
  $co = $sm.ChartObjects().Add(200, 20, 320, 180)
  $co.Chart.ChartType = 51
  $co.Chart.SetSourceData($ws.Range('B3:B13,D3:D13'))
  $ws.Activate(); $ws.Range('A1').Select()
  Save-Book $wb 'g-table.xlsx'

  # =============================================================================================
  # premerge fixtures
  # =============================================================================================
  $gakusei = U '学生'                                     # 学生
  $shimei = U '氏名'; $daigaku = U '大学'; $tensu = U '点数'; $goukei = U '合計'
  foreach ($variant in @('plain', 'table')) {
    $wb = $xl.Workbooks.Add()
    while ($wb.Worksheets.Count -lt 2) { [void]$wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)) }
    $ws = $wb.Worksheets.Item(1); $ws.Name = $gakusei
    $ot = $wb.Worksheets.Item(2); $ot.Name = 'Report'
    $h = @('ID', $shimei, $daigaku, $tensu)
    for ($c = 0; $c -lt 4; $c++) { SetF $ws.Cells.Item(1, $c + 1) $h[$c] }
    $ids = @('S001', 'S002', 'S003', 'S004', 'S005')
    for ($i = 0; $i -lt 5; $i++) {
      $ws.Cells.Item(2 + $i, 1).NumberFormat = '@'
      SetF $ws.Cells.Item(2 + $i, 1) $ids[$i]
      SetF $ws.Cells.Item(2 + $i, 2) ($gakusei + [char](65 + $i))
      SetF $ws.Cells.Item(2 + $i, 3) $unis[$i % 3]
      SetF $ws.Cells.Item(2 + $i, 4) (60 + $i * 5)
    }
    $ws.Range('A1:D1').Font.Bold = $true
    $ws.Range('A1:D1').Interior.Color = 0xF2E6DD
    $ws.Range('D2:D6').NumberFormat = '0.0'
    $ws.Range('A1:D6').Borders.Item(7).LineStyle = 1; $ws.Range('A1:D6').Borders.Item(8).LineStyle = 1
    $ws.Range('A1:D6').Borders.Item(9).LineStyle = 1; $ws.Range('A1:D6').Borders.Item(10).LineStyle = 1
    $ws.Range('A1:D6').Borders.Item(11).LineStyle = 1; $ws.Range('A1:D6').Borders.Item(12).LineStyle = 1
    if ($variant -eq 'table') {
      $lo = $ws.ListObjects.Add(1, $ws.Range('A1:D6'), $M, 1)
      $lo.Name = 'tblStudents'
    } else {
      [void]$ws.Range('A1:D6').AutoFilter(1)
    }
    SetF $ws.Range('C8') $goukei
    $ws.Range('D8').Formula = '=SUM(D2:D6)'
    $ws.Range('B8').Formula = '=COUNTA(B2:B6)'
    $ws.Range('D8').NumberFormat = '0.0'
    $ws.Activate(); $ws.Range('D2').Select()
    $fc = $ws.Range('D2:D6').FormatConditions.Add(1, 5, '=70'); $fc.Font.Bold = $true
    [void]$wb.Names.Add('Scores', ('=' + $ws.Name + '!$D$2:$D$6'))
    $ot.Range('A1').Formula = '=AVERAGE(' + $ws.Name + '!D2:D6)'
    $ot.Range('A2').Formula = '=SUM(Scores)'
    $ot.Range('A3').Formula = '=COUNTIF(' + $ws.Name + '!C2:C6,"SFC")'
    $co = $ot.ChartObjects().Add(100, 20, 300, 160)
    $co.Chart.ChartType = 51
    $co.Chart.SetSourceData($ws.Range('B1:B6,D1:D6'))
    $ws.Columns.Item(1).ColumnWidth = 8; $ws.Columns.Item(2).ColumnWidth = 12
    $ws.Activate(); $ws.Range('A1').Select()
    Save-Book $wb ('pm-base-' + $variant + '.xlsx')
  }

  # source: sheet KAITOU with styled columns and formulas
  $wb = $xl.Workbooks.Add()
  $ws = $wb.Worksheets.Item(1); $ws.Name = (U '回答')                        # 回答
  $h = @((U '学生ID'), (U '名前'), (U '志望'), (U '評価'), (U '備考'), (U '点数'), 'Calc', 'Total', $shimei)
  # 学生ID 名前 志望 評価 備考 点数 Calc Total 氏名
  for ($c = 0; $c -lt $h.Count; $c++) { SetF $ws.Cells.Item(1, $c + 1) $h[$c] }
  $rows = @(
    @('S003', 'C', 'A', 4.5, 'ok', 80),
    @('S001', 'A', 'B', 3.25, 'multi-line', 91),
    @('S009', 'Z', 'C', 2, 'new', 55),
    @('S002', 'B', 'A', 5, '', 70),
    @('S005', 'E', 'B', 1.5, 'late', 66)
  )
  for ($i = 0; $i -lt $rows.Count; $i++) {
    $r = 2 + $i
    $ws.Cells.Item($r, 1).NumberFormat = '@'
    for ($c = 0; $c -lt 6; $c++) { if ("$($rows[$i][$c])" -ne '') { SetF $ws.Cells.Item($r, $c + 1) $rows[$i][$c] } }
    SetF $ws.Cells.Item($r, 9) ($gakusei + $rows[$i][1])
  }
  $ws.Range('G2').Formula = '=F2*2+D2'; $ws.Range('G2:G6').FillDown()      # same-row refs -> keep
  $ws.Range('H2').Formula = '=SUM($F$2:$F$6)'; $ws.Range('H2:H6').FillDown() # column total -> value
  $ws.Range('A1:I1').Font.Bold = $true
  $ws.Range('A1:I1').Interior.Color = 0x66CCFF
  $ws.Range('C2:C6').Font.Color = 0x0000C0; $ws.Range('C2:C6').Font.Italic = $true
  $ws.Range('C2:C6').Interior.Color = 0xCCFFFF
  $ws.Range('D2:D6').NumberFormat = '0.00"' + (U '点') + '"'
  $ws.Range('F2:F6').NumberFormat = '#,##0"pt"'
  $ws.Range('E2:E6').WrapText = $true
  $ws.Range('E2:E6').Borders.Item(9).LineStyle = -4115; $ws.Range('E2:E6').Borders.Item(9).Weight = -4138
  $ws.Range('E3').Formula = "=""line1""&CHAR(10)&""line2"""
  $ws.Columns.Item(3).ColumnWidth = 18; $ws.Columns.Item(5).ColumnWidth = 22; $ws.Columns.Item(4).ColumnWidth = 11
  Save-Book $wb 'pm-src.xlsx'
}
finally { Close-GExcel $xl }
