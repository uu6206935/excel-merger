# Builds the Excel-authored fixtures of the sheet copier (PremergeOptions.includeSourceSheets, round 9; owner: patcher P).
#   powershell -NoProfile -ExecutionPolicy Bypass -File tests\golden\com\make-fixtures-addsheets.ps1 -OutDir tests\fixtures\addsheets
# ex-formats.xlsx     名簿 + マスタ: custom number formats (text, 0.0"点", Japanese dates / eras, %, yen suffix), theme-coloured header with
#                     tint, merged title, frozen panes, hidden column, grouped rows, tab colour, column widths / row heights, conditional formats
#                     with dxfs (cell value, expression that reads マスタ, colour scale, data bar, icon set), data validations (list via a
#                     name, list on another sheet, whole number, date), hyperlinks (external, to another sheet), sparklines, names
#                     (workbook scope, sheet scope, print area), formulas that read マスタ and a name, a spilling dynamic array.
# ex-theme-other.xlsx a workbook whose theme colours / fonts differ from the default Office theme: cells with theme colours + tint (font, fill,
#                     border), rich text, a conditional format with a theme colour; second sheet 参照 read by the first.
# ex-two-sheets.xlsx  売上 2025 (sheet name with a space) + Summary: cross-sheet formulas (direct, range, VLOOKUP, XLOOKUP, LET, a name),
#                     a sheet-scoped name, a hidden sheet.
# ex-1904.xlsx        the 1904 date system: dates, date-times, times of day, durations, a formula over them.
# ex-macro.xlsm      a macro-enabled workbook (no VBA project: the content type is what counts) with two sheets and a name.
# Output (ASCII stdout): one line per saved file, "warn ..." for optional features this Excel could not create.
# Own Excel instance only (tests/golden/com/common.ps1); no clipboard, no PageSetup.
param([string]$OutDir, [switch]$OnlyMacro)
$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\common.ps1"
$M = $script:GMissing
$OutDir = (Resolve-Path $OutDir).Path

function Save-Book($wb, [string]$name, [int]$format = 51) {
  $p = Join-Path $OutDir $name
  if (Test-Path $p) { Remove-Item $p -Force }
  $wb.SaveAs($p, $format)
  $wb.Close($false)
  Write-Host "saved $name"
}
function Try-Do([string]$label, [scriptblock]$sb) {
  try { & $sb } catch { Write-Host ('warn ' + $label + ': ' + (($_.Exception.Message) -replace '[^\x20-\x7e]', '?')) }
}
function NewSheet($wb, [string]$name) {
  $ws = $wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count))
  $ws.Name = $name
  return $ws
}
function Rgb([int]$r, [int]$g, [int]$b) { return $r + 256 * $g + 65536 * $b }

$xl = $null
try {
  $xl = New-GExcel
  Write-Host "excel pid $($script:GExcelPid)"

  if (-not $OnlyMacro) {
  # =============================================================================================
  # ex-formats.xlsx
  # =============================================================================================
  $wb = $xl.Workbooks.Add()
  $ws = $wb.Worksheets.Item(1); $ws.Name = '名簿'
  $mst = NewSheet $wb 'マスタ'
  SetF $mst.Range('A1') '区分'
  $kinds = @('正社員', '契約', '派遣', 'アルバイト')
  for ($i = 0; $i -lt 4; $i++) { SetF $mst.Cells.Item($i + 2, 1) $kinds[$i] }
  SetF $mst.Range('B1') 0.1; SetF $mst.Range('C1') '基準日'; SetF $mst.Range('D1') 45658; $mst.Range('D1').NumberFormat = 'yyyy/m/d'
  [void]$wb.Names.Add('税率', '=マスタ!$B$1')
  [void]$wb.Names.Add('候補', '=マスタ!$A$2:$A$5')
  [void]$wb.Names.Add('基準日', '=マスタ!$D$1')
  [void]$ws.Names.Add('局所', '=名簿!$C$3:$C$12')

  # title (merged), header
  $ws.Range('A1').Value2 = '社員名簿（フォーマット見本）'
  $ws.Range('A1:H1').Merge()
  $ws.Range('A1').Font.Size = 14; $ws.Range('A1').Font.Bold = $true
  $ws.Range('A1').HorizontalAlignment = -4108
  $h = @('ID', '氏名', '点数', '入社日', '比率', '給与', '判定', '税込')
  for ($c = 0; $c -lt 8; $c++) { $ws.Cells.Item(2, $c + 1).Value2 = $h[$c] }
  $hd = $ws.Range('A2:H2')
  $hd.Font.Bold = $true; $hd.Font.ThemeColor = 1
  $hd.Interior.ThemeColor = 5; $hd.Interior.TintAndShade = 0.6
  $hd.Borders.LineStyle = 1; $hd.Borders.Weight = 2
  $ws.Range('A3:A12').NumberFormat = '@'
  $ws.Range('C3:C12').NumberFormat = '0.0"点"'
  $ws.Range('D3:D12').NumberFormat = 'yyyy"年"m"月"d"日"'
  $ws.Range('E3:E12').NumberFormat = '0.0%'
  $ws.Range('F3:F12').NumberFormat = '#,##0"円";-#,##0"円"'
  $ws.Range('H3:H12').NumberFormat = '#,##0.00'
  for ($i = 0; $i -lt 10; $i++) {
    $r = $i + 3
    $ws.Cells.Item($r, 1).Value2 = ('000' + (100 + $i))
    $ws.Cells.Item($r, 2).Value2 = ('社員' + [char](65 + $i))
    SetF $ws.Cells.Item($r, 3) (55 + ($i * 7) % 45)
    SetF $ws.Cells.Item($r, 4) (44927 + $i * 31)
    SetF $ws.Cells.Item($r, 5) (0.05 * ($i + 1))
    SetF $ws.Cells.Item($r, 6) ((-1) * (($i % 3) - 1) * (200000 + 15000 * $i))
  }
  $ws.Range('G3').Formula = '=IF(C3>=80,"優",IF(C3>=60,"良","可"))'; $ws.Range('G3:G12').FillDown()
  $ws.Range('H3').Formula = '=F3*(1+税率)'; $ws.Range('H3:H12').FillDown()
  $ws.Range('J2').Value2 = '区分'; $ws.Range('K2').Value2 = '基準との差'; $ws.Range('L2').Value2 = '推移'; $ws.Range('M2').Value2 = '並び'
  for ($i = 0; $i -lt 10; $i++) { $ws.Cells.Item($i + 3, 10).Value2 = $kinds[$i % 4] }
  $ws.Range('K3').Formula = '=D3-基準日'; $ws.Range('K3:K12').FillDown(); $ws.Range('K3:K12').NumberFormat = '0'
  Try-Do 'dynamic array' { $ws.Range('M3').Formula2 = '=SORT(C3:C12)' }
  $ws.Columns.Item(1).ColumnWidth = 9; $ws.Columns.Item(2).ColumnWidth = 14; $ws.Columns.Item(4).ColumnWidth = 16; $ws.Columns.Item(6).ColumnWidth = 13
  $ws.Columns.Item(5).Hidden = $true
  $ws.Rows.Item(2).RowHeight = 24
  Try-Do 'row group' { $ws.Rows('9:11').Group() | Out-Null }
  $ws.Tab.ThemeColor = 7
  # conditional formats with dxfs
  Try-Do 'cf cellvalue' { $fc = $ws.Range('C3:C12').FormatConditions.Add(1, 7, '=80'); $fc.Interior.Color = (Rgb 198 239 206); $fc.Font.Color = (Rgb 0 97 0) }
  Try-Do 'cf expression' { $fc = $ws.Range('B3:B12').FormatConditions.Add(2, [Type]::Missing, '=マスタ!$B$1>0.05'); $fc.Font.Bold = $true }
  Try-Do 'cf scale' { [void]$ws.Range('F3:F12').FormatConditions.AddColorScale(2) }
  Try-Do 'cf databar' { [void]$ws.Range('K3:K12').FormatConditions.AddDatabar() }
  Try-Do 'cf iconset' { [void]$ws.Range('E3:E12').FormatConditions.AddIconSetCondition() }
  # data validations
  Try-Do 'dv list name' { $ws.Range('J3:J12').Validation.Add(3, 1, 1, '=候補') }
  Try-Do 'dv list other sheet' { $ws.Range('I3:I12').Validation.Add(3, 1, 1, '=マスタ!$A$2:$A$5') }
  Try-Do 'dv whole' { $ws.Range('C3:C12').Validation.Add(1, 1, 1, '0', '100') }
  Try-Do 'dv date' { $ws.Range('D3:D12').Validation.Add(4, 1, 1, '44562', '47118') }
  # hyperlinks
  Try-Do 'hyperlinks' {
    [void]$ws.Hyperlinks.Add($ws.Range('B3'), 'https://example.com/a?b=1&c=2', '', 'ツールチップ', '社員A')
    [void]$ws.Hyperlinks.Add($ws.Range('B4'), '', "'マスタ'!A1", '', 'マスタへ')
  }
  Try-Do 'sparkline' { [void]$ws.Range('L3').SparklineGroups.Add(1, '名簿!C3:E3') }
  Try-Do 'freeze' { $ws.Activate(); $xl.ActiveWindow.FreezePanes = $false; $ws.Range('C3').Select(); $xl.ActiveWindow.FreezePanes = $true }
  Try-Do 'print area' { [void]$ws.Names.Add('Print_Area', '=名簿!$A$1:$H$12') }
  Try-Do 'select' { $ws.Activate(); $ws.Range('A1').Select() | Out-Null }
  Save-Book $wb 'ex-formats.xlsx'

  # =============================================================================================
  # ex-theme-other.xlsx
  # =============================================================================================
  $wb = $xl.Workbooks.Add()
  $ws = $wb.Worksheets.Item(1); $ws.Name = '配色'
  $ref = NewSheet $wb '参照'
  Try-Do 'theme colours' {
    $cs = $wb.Theme.ThemeColorScheme
    $cs.Colors(5).RGB = (Rgb 192 80 77)    # accent 1
    $cs.Colors(6).RGB = (Rgb 155 187 89)   # accent 2
    $cs.Colors(7).RGB = (Rgb 128 100 162)  # accent 3
    $cs.Colors(8).RGB = (Rgb 75 172 198)   # accent 4
    $cs.Colors(9).RGB = (Rgb 247 150 70)   # accent 5
    $cs.Colors(10).RGB = (Rgb 44 77 117)   # accent 6
    $cs.Colors(3).RGB = (Rgb 31 73 125)    # dark 2
    $cs.Colors(4).RGB = (Rgb 238 236 225)  # light 2
  }
  Try-Do 'theme fonts' { $wb.Theme.ThemeFontScheme.MinorFont.Item(1).Name = 'Arial'; $wb.Theme.ThemeFontScheme.MajorFont.Item(1).Name = 'Georgia' }
  SetF $ref.Range('A1') 100; SetF $ref.Range('A2') 200
  for ($c = 1; $c -le 6; $c++) {
    $cell = $ws.Cells.Item(1, $c)
    $cell.Value2 = ('色' + $c)
    $cell.Interior.ThemeColor = 4 + $c       # accent 1..6 (xlThemeColorAccent1 = 5)
    $cell.Interior.TintAndShade = 0.4 - 0.2 * ($c % 3)
    $cell.Font.ThemeColor = 1
    $cell.Borders.Item(9).LineStyle = 1; $cell.Borders.Item(9).ThemeColor = 4 + $c; $cell.Borders.Item(9).Weight = -4138
  }
  for ($r = 2; $r -le 8; $r++) {
    $ws.Cells.Item($r, 1).Value2 = ('項目' + $r)
    SetF $ws.Cells.Item($r, 2) ($r * 11)
    $ws.Cells.Item($r, 2).Font.ThemeColor = 6
    $ws.Cells.Item($r, 3).Formula = '=B' + $r + '+参照!$A$1'
  }
  $ws.Range('A2:A8').Font.Name = '+Body'
  Try-Do 'rich text' { $ws.Range('E2').Value2 = '赤と青'; $ws.Range('E2').Characters(1, 1).Font.Color = (Rgb 255 0 0); $ws.Range('E2').Characters(3, 1).Font.ThemeColor = 5; $ws.Range('E2').Characters(2, 1).Font.Bold = $true }
  Try-Do 'cf theme' { $fc = $ws.Range('B2:B8').FormatConditions.Add(1, 5, '=50'); $fc.Interior.ThemeColor = 6; $fc.Interior.TintAndShade = 0.4 }
  $ws.Columns.Item(1).ColumnWidth = 12
  Try-Do 'select' { $ws.Activate(); $ws.Range('A1').Select() | Out-Null }
  Save-Book $wb 'ex-theme-other.xlsx'

  # =============================================================================================
  # ex-two-sheets.xlsx
  # =============================================================================================
  $wb = $xl.Workbooks.Add()
  $ws = $wb.Worksheets.Item(1); $ws.Name = 'Summary'
  $sl = NewSheet $wb '売上 2025'
  $hid = NewSheet $wb '作業'
  $hid.Visible = 0
  $h = @('商品', '数量', '単価')
  for ($c = 0; $c -lt 3; $c++) { $sl.Cells.Item(1, $c + 1).Value2 = $h[$c] }
  $items = @('りんご', 'みかん', 'ぶどう', 'もも', 'なし')
  for ($i = 0; $i -lt 5; $i++) { $sl.Cells.Item($i + 2, 1).Value2 = $items[$i]; SetF $sl.Cells.Item($i + 2, 2) (3 + $i); SetF $sl.Cells.Item($i + 2, 3) (100 + 20 * $i) }
  SetF $sl.Range('E1') 0.08
  SetF $hid.Range('A1') 7
  [void]$wb.Names.Add('消費税', "='売上 2025'!`$E`$1")
  [void]$wb.Names.Add('隠し値', '=作業!$A$1')
  [void]$sl.Names.Add('合計行', "='売上 2025'!`$B`$2:`$B`$6")
  $ws.Range('A1').Value2 = '商品'; $ws.Range('B1').Value2 = '金額'; $ws.Range('C1').Value2 = '検索'; $ws.Range('D1').Value2 = '合計'; $ws.Range('E1').Value2 = '税込'; $ws.Range('F1').Value2 = '隠し'
  for ($i = 0; $i -lt 5; $i++) {
    $r = $i + 2
    $ws.Cells.Item($r, 1).Formula = "='売上 2025'!A$r"
    $ws.Cells.Item($r, 2).Formula = "='売上 2025'!B$r*'売上 2025'!C$r"
    $ws.Cells.Item($r, 3).Formula = "=VLOOKUP(A$r,'売上 2025'!`$A`$2:`$C`$6,3,FALSE)"
    $ws.Cells.Item($r, 5).Formula = "=B$r*(1+消費税)"
    $ws.Cells.Item($r, 6).Formula = "=隠し値*B$r"
  }
  $ws.Range('D2').Formula = "=SUM('売上 2025'!B2:B6)"
  Try-Do 'xlookup' { $ws.Range('D3').Formula2 = "=XLOOKUP(""ぶどう"",'売上 2025'!A2:A6,'売上 2025'!C2:C6)" }
  Try-Do 'let' { $ws.Range('D4').Formula2 = "=LET(x,SUM(合計行),x*2)" }
  Try-Do 'spill' { $ws.Range('H2').Formula2 = "=SORT('売上 2025'!C2:C6,,-1)" }
  $sl.Range('A8').Formula = '=Summary!B2+1'
  Try-Do 'select' { $ws.Activate(); $ws.Range('A1').Select() | Out-Null }
  Save-Book $wb 'ex-two-sheets.xlsx'

  # =============================================================================================
  # ex-1904.xlsx
  # =============================================================================================
  $wb = $xl.Workbooks.Add()
  try { $wb.Date1904 = $true } catch { Write-Host 'warn date1904' }
  $ws = $wb.Worksheets.Item(1); $ws.Name = '日付'
  $ws.Range('A1').Value2 = '日付'; $ws.Range('B1').Value2 = '日時'; $ws.Range('C1').Value2 = '時刻'; $ws.Range('D1').Value2 = '経過'; $ws.Range('E1').Value2 = '差'
  $ws.Range('A2:A6').NumberFormat = 'yyyy/m/d'; $ws.Range('B2:B6').NumberFormat = 'yyyy/m/d h:mm'; $ws.Range('C2:C6').NumberFormat = 'h:mm'; $ws.Range('D2:D6').NumberFormat = '[h]:mm'
  for ($i = 0; $i -lt 5; $i++) {
    $r = $i + 2
    SetF $ws.Cells.Item($r, 1) (41000 + 40 * $i)
    SetF $ws.Cells.Item($r, 2) (41000.5 + 40 * $i)
    SetF $ws.Cells.Item($r, 3) (0.25 + 0.1 * $i)
    SetF $ws.Cells.Item($r, 4) (1.5 + $i)
  }
  $ws.Range('E2').Formula = '=A3-A2'; $ws.Range('E2:E5').FillDown()
  $ws.Columns.Item(1).ColumnWidth = 12; $ws.Columns.Item(2).ColumnWidth = 18
  Save-Book $wb 'ex-1904.xlsx'
  }
  # =============================================================================================
  # ex-macro.xlsm (xlOpenXMLWorkbookMacroEnabled = 52)
  # =============================================================================================
  $wb = $xl.Workbooks.Add()
  $ws = $wb.Worksheets.Item(1); $ws.Name = '台帳'
  $cf = NewSheet $wb '設定'
  SetF $cf.Range('A1') 0.1
  [void]$wb.Names.Add('率', '=設定!$A$1')
  $ws.Range('A1').Value2 = 'ID'; $ws.Range('B1').Value2 = '氏名'; $ws.Range('C1').Value2 = '金額'; $ws.Range('D1').Value2 = '税込'
  for ($i = 0; $i -lt 6; $i++) {
    $r = $i + 2
    $ws.Cells.Item($r, 1).Value2 = ('M' + (100 + $i)); $ws.Cells.Item($r, 2).Value2 = ('担当' + [char](65 + $i)); SetF $ws.Cells.Item($r, 3) (1000 * ($i + 1))
    $ws.Cells.Item($r, 4).Formula = '=C' + $r + '*(1+率)'
  }
  Save-Book $wb 'ex-macro.xlsm' 52
}
finally { Close-GExcel $xl }
