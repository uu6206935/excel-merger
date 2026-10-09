# Builds the sample / fixture workbooks with OUR OWN hidden Excel instance (owner: T).
# Normally started by `npm run samples` (tools/make-samples.mjs), which generates the data (tools/lib/sample-data.mjs),
# writes it as JSON, runs this script with ASCII temp paths and copies the results into samples/.
#
#   powershell -NoProfile -ExecutionPolicy Bypass -File tools\make-samples.ps1 -SpecJson <spec.json> -OutDir <dir>
#
# Output (in -OutDir): s01.xlsx (01_学生名簿), s02.xlsx (02_面接評価), s04.xls (04_適性検査), s05.xlsx (05_研修出欠).
# stdout: ASCII JSON lines ({"type":"excel","pid":..} first, then {"type":"step"...}, {"type":"saved"...}, {"type":"done"}).
param(
  [Parameter(Mandatory = $true)][string]$SpecJson,
  [Parameter(Mandatory = $true)][string]$OutDir
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'lib\xlcom.ps1')

$M = [Type]::Missing
$spec = [IO.File]::ReadAllText($SpecJson, [Text.Encoding]::UTF8) | ConvertFrom-Json
if (-not (Test-Path $OutDir)) { [void](New-Item -ItemType Directory -Path $OutDir -Force) }

# ---- constants (Excel enums) ------------------------------------------------------------------------------------
$xlCenter = -4108; $xlLeft = -4131; $xlRight = -4152
$xlContinuous = 1; $xlThin = 2; $xlMedium = -4138; $xlDouble = -4119
$xlEdgeTop = 8; $xlEdgeBottom = 9
$xlCellValue = 1; $xlExpression = 2; $xlGreaterEqual = 7; $xlEqual = 3
$xlValidateList = 3; $xlValidAlertStop = 1; $xlBetween = 1
$xlOpenXMLWorkbook = 51; $xlExcel8 = 56
$xlColumnClustered = 51
$xlThemeColorDark1 = 1; $xlThemeColorLight1 = 2; $xlThemeColorAccent1 = 5; $xlThemeColorAccent6 = 10; $xlThemeColorAccent4 = 8

function Step([string]$msg) { Write-EmLine ([ordered]@{ type = 'step'; message = $msg }) }

function ColLetter([int]$c) {
  $s = ''
  while ($c -gt 0) { $m = ($c - 1) % 26; $s = [char](65 + $m) + $s; $c = [int][math]::Floor(($c - 1) / 26) }
  return $s
}

# Write a JSON 2-D block (array of rows) through Range.Formula (strings starting with '=' are formulas,
# "'" forces text, numbers stay numbers, null = empty).
function Write-Block($ws, [int]$row, [int]$col, $rows) {
  $list = @($rows)
  $n = $list.Count
  if ($n -eq 0) { return }
  $w = 0
  foreach ($r in $list) { $c = @($r).Count; if ($c -gt $w) { $w = $c } }
  $arr = New-Object 'object[,]' $n, $w
  for ($i = 0; $i -lt $n; $i++) {
    $r = @($list[$i])
    for ($j = 0; $j -lt $r.Count; $j++) {
      $v = $r[$j]
      if ($null -eq $v) { continue }
      # cast explicitly: PSObject-wrapped values in an object[,] make COM fail with E_OUTOFMEMORY
      if ($v -is [string]) { $arr[$i, $j] = [string]$v }
      elseif ($v -is [bool]) { $arr[$i, $j] = [bool]$v }
      else { $arr[$i, $j] = [double]$v }
    }
  }
  $rng = $ws.Range($ws.Cells.Item($row, $col), $ws.Cells.Item($row + $n - 1, $col + $w - 1))
  $rng.Formula = $arr
}

function Write-Row($ws, [int]$row, [int]$col, $values) { Write-Block $ws $row $col @(, @($values)) }

function Set-Borders($rng) {
  $rng.Borders.LineStyle = $xlContinuous
  $rng.Borders.Weight = $xlThin
  $rng.Borders.Color = 0xBFBFBF   # BGR: light grey
}

function Style-Header($rng, [int]$theme, [double]$tint) {
  $rng.Font.Bold = $true
  $rng.Interior.ThemeColor = $theme
  $rng.Interior.TintAndShade = $tint
  $rng.HorizontalAlignment = $xlCenter
  $rng.VerticalAlignment = $xlCenter
  $rng.WrapText = $true
}

# New workbook with exactly these sheets. Uses Worksheets.Add() without arguments (the PowerShell COM binder
# sometimes refuses Add(Before, After) after the first workbook) and never touches the user's
# SheetsInNewWorkbook option (Excel would persist it).
function New-Book([string[]]$names) {
  $wb = $xl.Workbooks.Add()
  while ($wb.Worksheets.Count -lt $names.Count) { [void]$wb.Worksheets.Add() }
  while ($wb.Worksheets.Count -gt $names.Count) { $wb.Worksheets.Item($wb.Worksheets.Count).Delete() }
  # temporary unique names first (a fresh sheet may already be called like one of the targets)
  for ($i = 1; $i -le $names.Count; $i++) { $wb.Worksheets.Item($i).Name = "tmp_$i" }
  for ($i = 0; $i -lt $names.Count; $i++) { $wb.Worksheets.Item($i + 1).Name = $names[$i] }
  return $wb
}

function Save-Book($wb, [string]$name, [int]$format) {
  $p = Join-Path $OutDir $name
  if (Test-Path $p) { Remove-Item $p -Force }
  $wb.SaveAs($p, $format)
  Write-EmLine ([ordered]@{ type = 'saved'; file = $name })
}

$xl = $null
try {
  $xl = New-EmExcel

  # ==============================================================================================================
  # 01_学生名簿.xlsx (BASE)
  # ==============================================================================================================
  Step 'base: sheets'
  $b = $spec.base
  $R = $b.roster
  $nStu = @($b.rows).Count
  $hr = [int]$R.headerRow                      # 3
  $fr = [int]$R.firstDataRow                   # 4
  $lr = $fr + $nStu - 1                        # 63
  $nCol = @($R.headers).Count                  # 12 (L = hidden helper)
  $lastL = ColLetter $nCol                     # L
  $footer = $lr + 2                            # 65 (one blank row between the table and the footer)

  $wb = New-Book @('名簿', '集計', 'ピボット', 'マスタ')
  $ws = $wb.Worksheets.Item('名簿')
  $sum = $wb.Worksheets.Item('集計')
  $piv = $wb.Worksheets.Item('ピボット')
  $mst = $wb.Worksheets.Item('マスタ')

  # ---- マスタ ----------------------------------------------------------------------------------------------------
  $m = $b.master
  Write-Row $mst 1 1 @('大学一覧', '学部一覧', '学年', $null, '志望度', '意味', $null, '基準日')
  Write-Block $mst 2 1 (@($m.universities) | ForEach-Object { , @($_) })
  Write-Block $mst 2 2 (@($m.faculties) | ForEach-Object { , @($_) })
  Write-Block $mst 2 3 (@($m.grades) | ForEach-Object { , @($_) })
  Write-Block $mst 2 5 @($m.aspiration)
  Write-Row $mst 2 8 @($m.refDate)
  $mst.Range('H2').NumberFormat = 'yyyy/m/d'
  Style-Header $mst.Range('A1:C1') $xlThemeColorAccent4 0.6
  Style-Header $mst.Range('E1:F1') $xlThemeColorAccent4 0.6
  Style-Header $mst.Range('H1') $xlThemeColorAccent4 0.6
  $mst.Columns.Item(1).ColumnWidth = 16; $mst.Columns.Item(2).ColumnWidth = 26; $mst.Columns.Item(3).ColumnWidth = 10
  $mst.Columns.Item(4).ColumnWidth = 3; $mst.Columns.Item(5).ColumnWidth = 8; $mst.Columns.Item(6).ColumnWidth = 10
  $mst.Columns.Item(7).ColumnWidth = 3; $mst.Columns.Item(8).ColumnWidth = 12
  $mst.Range('E2:E6').HorizontalAlignment = $xlCenter
  $mst.Tab.ThemeColor = $xlThemeColorAccent4

  # ---- names (before the formulas that use them) ----------------------------------------------------------------
  [void]$wb.Names.Add('大学一覧', '=マスタ!$A$2:$A$20')
  [void]$wb.Names.Add('基準日', '=マスタ!$H$2')
  [void]$wb.Names.Add('志望度列', '=名簿!$K:$K')

  # ---- 名簿 --------------------------------------------------------------------------------------------------------
  Step 'base: roster'
  $ws.Range('A1').Formula = [string]$b.title
  $ws.Range('A1:K1').Merge()
  $t = $ws.Range('A1')
  $t.Font.Size = 14; $t.Font.Bold = $true
  $t.Font.ThemeColor = $xlThemeColorDark1
  $t.Interior.ThemeColor = $xlThemeColorAccent1; $t.Interior.TintAndShade = 0
  $t.HorizontalAlignment = $xlLeft; $t.VerticalAlignment = $xlCenter; $t.IndentLevel = 1
  $ws.Rows.Item(1).RowHeight = 26
  $ws.Range('A2').Formula = [string]$b.note
  $ws.Range('A2').Font.Size = 9; $ws.Range('A2').Font.Color = 0x595959

  Write-Row $ws $hr 1 @($R.headers)
  Style-Header $ws.Range("A$($hr):$lastL$($hr)") $xlThemeColorAccent1 0.6
  $ws.Rows.Item($hr).RowHeight = 22
  # 学生ID as text: format the column BEFORE writing so '000123' keeps its zeros
  $ws.Range("A$($fr):A$($lr)").NumberFormat = '@'
  Write-Block $ws $fr 1 $b.rows
  $ws.Range("I$($fr):I$($lr)").NumberFormat = 'yyyy/m/d'
  $ws.Range("J$($fr):J$($lr)").NumberFormat = '0"歳"'
  $ws.Range("A$($fr):A$($lr)").HorizontalAlignment = $xlCenter
  $ws.Range("F$($fr):F$($lr)").HorizontalAlignment = $xlCenter
  $ws.Range("J$($fr):K$($lr)").HorizontalAlignment = $xlCenter
  Set-Borders $ws.Range("A$($hr):$lastL$($lr)")
  $hb = $ws.Range("A$($hr):$lastL$($hr)").Borders.Item($xlEdgeBottom)
  $hb.LineStyle = $xlContinuous; $hb.Weight = $xlMedium; $hb.Color = 0x7F3F1F

  # footer below the table (one blank row): 人数 / 平均
  Write-Row $ws $footer 1 @('人数', "=COUNTA(B$($fr):B$($lr))", $null, $null, $null, $null, $null, $null, '平均', "=AVERAGE(J$($fr):J$($lr))", "=AVERAGE(K$($fr):K$($lr))")
  Write-Row $ws ($footer + 1) 1 @('志望度4以上', "=COUNTIF(K$($fr):K$($lr),"">=4"")")
  $ws.Range("B$($footer):B$($footer + 1)").NumberFormat = '0"名"'
  $ws.Range("J$($footer)").NumberFormat = '0.0"歳"'
  $ws.Range("K$($footer)").NumberFormat = '0.00'
  $fz = $ws.Range("A$($footer):K$($footer + 1)")
  $fz.Font.Bold = $true
  $fz.Interior.ThemeColor = $xlThemeColorAccent1; $fz.Interior.TintAndShade = 0.8
  $ft = $ws.Range("A$($footer):K$($footer)").Borders.Item($xlEdgeTop)
  $ft.LineStyle = $xlDouble; $ft.Color = 0x7F3F1F

  # widths (varied) + hidden helper column
  $widths = @(9, 14, 18, 15, 22, 9, 30, 14, 11, 7, 8, 10)
  for ($c = 1; $c -le $widths.Count; $c++) { $ws.Columns.Item($c).ColumnWidth = $widths[$c - 1] }
  $ws.Columns.Item([int]$R.helperCol).Hidden = $true

  # conditional formatting (CF formulas are relative to the ACTIVE cell -> select the top-left cell first)
  Step 'base: CF / DV'
  $ws.Activate()
  [void]$ws.Range("K$fr").Select()
  $fc1 = $ws.Range("K$($fr):K$($lr)").FormatConditions.Add($xlCellValue, $xlGreaterEqual, '=4')
  $fc1.Interior.Color = 0x9CE0FF      # BGR -> RGB FFE09C (light orange)
  $fc1.Font.Color = 0x00569C          # BGR -> RGB 9C5600
  $fc1.Font.Bold = $true
  [void]$ws.Range("B$fr").Select()
  $fc2 = $ws.Range("B$($fr):B$($lr)").FormatConditions.Add($xlExpression, $M, "=`$L$fr>1")
  $fc2.Font.Color = 0x0000C0          # red
  $fc2.Interior.Color = 0xE6E6FF      # BGR -> RGB FFE6E6 (pale red)

  # data validation
  $dv1 = $ws.Range("K$($fr):K$($lr)").Validation
  $dv1.Delete(); $dv1.Add($xlValidateList, $xlValidAlertStop, $xlBetween, '1,2,3,4,5')
  $dv1.ErrorTitle = '志望度'; $dv1.ErrorMessage = '志望度は 1〜5 から選んでください。'
  $dv2 = $ws.Range("D$($fr):D$($lr)").Validation
  $dv2.Delete(); $dv2.Add($xlValidateList, $xlValidAlertStop, $xlBetween, '=マスタ!$A$2:$A$20')
  $dv2.ErrorTitle = '大学'; $dv2.ErrorMessage = 'マスタシートの大学一覧から選んでください。'

  # autofilter + freeze panes at C4
  [void]$ws.Range("A$($hr):$lastL$($lr)").AutoFilter(1)
  [void]$ws.Range("C$fr").Select()
  $xl.ActiveWindow.FreezePanes = $true
  $ws.Tab.ThemeColor = $xlThemeColorAccent1

  # ---- 集計 --------------------------------------------------------------------------------------------------------
  Step 'base: summary'
  $s = $b.summary
  $sum.Range('A1').Formula = [string]$s.title
  $sum.Range('A1').Font.Size = 14; $sum.Range('A1').Font.Bold = $true
  $sum.Range('A2').Formula = [string]$s.note
  $sum.Range('A2').Font.Size = 9; $sum.Range('A2').Font.Color = 0x595959
  $sh = [int]$s.headerRow   # 4
  Write-Row $sum $sh 1 @($s.headers)
  Style-Header $sum.Range("A$($sh):F$($sh)") $xlThemeColorAccent6 0.6
  Write-Block $sum ($sh + 1) 1 $s.uniRows
  $tr = [int]$s.totalRow
  Write-Row $sum $tr 1 @($s.total)
  $sum.Range("A$($tr):F$($tr)").Font.Bold = $true
  $tb = $sum.Range("A$($tr):F$($tr)").Borders.Item($xlEdgeTop); $tb.LineStyle = $xlDouble
  Set-Borders $sum.Range("A$($sh):F$($tr)")
  $sum.Range("B$($sh + 1):B$($tr)").NumberFormat = '0"名"'
  $sum.Range("C$($sh + 1):C$($tr)").NumberFormat = '0.0"歳"'
  $sum.Range("D$($sh + 1):D$($tr)").NumberFormat = '0"名"'
  $sum.Range("E$($sh + 1):E$($tr)").NumberFormat = '0.00'
  $sum.Range("F$($sh + 1):F$($tr)").NumberFormat = '0.0%'

  $gh = [int]$s.gradeHeaderRow
  Write-Row $sum $gh 1 @('学年', '人数')
  Style-Header $sum.Range("A$($gh):B$($gh)") $xlThemeColorAccent6 0.6
  Write-Block $sum ($gh + 1) 1 $s.gradeRows
  $gl = $gh + @($s.gradeRows).Count
  Set-Borders $sum.Range("A$($gh):B$($gl)")
  $sum.Range("B$($gh + 1):B$($gl)").NumberFormat = '0"名"'

  $oh = [int]$s.otherHeaderRow
  Write-Row $sum $oh 1 @('項目', '値')
  Style-Header $sum.Range("A$($oh):B$($oh)") $xlThemeColorAccent6 0.6
  Write-Block $sum ($oh + 1) 1 $s.others
  $ol = $oh + @($s.others).Count
  Set-Borders $sum.Range("A$($oh):B$($ol)")
  $sum.Range("B$($oh + 1):B$($ol)").HorizontalAlignment = $xlRight
  $sum.Columns.Item(1).ColumnWidth = 22; $sum.Columns.Item(2).ColumnWidth = 14
  for ($c = 3; $c -le 6; $c++) { $sum.Columns.Item($c).ColumnWidth = 12 }

  # chart: 人数 per 大学
  $left = [double]$sum.Range('H4').Left; $top = [double]$sum.Range('H4').Top
  $co = $sum.ChartObjects().Add($left, $top, 420, 260)
  $co.Name = '大学別人数グラフ'
  $ch = $co.Chart
  $ch.ChartType = $xlColumnClustered
  $ch.SetSourceData($sum.Range("A$($sh):B$($sh + @($s.uniRows).Count)"))
  $ch.HasTitle = $true; $ch.ChartTitle.Text = '大学別 応募者数'
  $ch.HasLegend = $false
  $sum.Tab.ThemeColor = $xlThemeColorAccent6

  # ---- ピボット ------------------------------------------------------------------------------------------------------
  Step 'base: pivot'
  $piv.Range('A1').Formula = '大学×学年 応募者数（ピボットテーブル）'
  $piv.Range('A1').Font.Bold = $true; $piv.Range('A1').Font.Size = 12
  $srcAddr = "名簿!R$($hr)C1:R$($lr)C$($nCol)"
  $pc = $wb.PivotCaches().Create(1, $srcAddr, 5)
  $pt = $pc.CreatePivotTable('ピボット!R3C1', '大学別学年ピボット', $true, 5)
  $pt.PivotFields('大学').Orientation = 1          # xlRowField
  $pt.PivotFields('学年').Orientation = 2          # xlColumnField
  [void]$pt.AddDataField($pt.PivotFields('学生ID'), '人数', -4112)   # xlCount
  $piv.Columns.Item(1).ColumnWidth = 16

  $ws.Activate()
  [void]$ws.Range('A1').Select()
  $xl.ActiveWindow.ScrollRow = 1; $xl.ActiveWindow.ScrollColumn = 1
  [void]$ws.Range("C$fr").Select()
  Save-Book $wb 's01.xlsx' $xlOpenXMLWorkbook
  $wb.Close($false)

  # ==============================================================================================================
  # 02_面接評価.xlsx
  # ==============================================================================================================
  Step 'interview'
  $iv = $spec.interview
  $n2 = @($iv.rows).Count
  $l2 = 1 + $n2
  $wb = New-Book @([string]$iv.sheet, [string]$iv.tableSheet)
  $w1 = $wb.Worksheets.Item(1)
  $w2 = $wb.Worksheets.Item(2)
  foreach ($w in @($w1, $w2)) {
    Write-Row $w 1 1 @($iv.headers)
    Write-Block $w 2 1 $iv.rows
    $w.Range("G2:G$l2").NumberFormat = '0.0'
    $w.Range("D2:H$l2").HorizontalAlignment = $xlCenter
    $widths2 = @(11, 16, 9, 7, 7, 7, 7, 8, 30)
    for ($c = 1; $c -le $widths2.Count; $c++) { $w.Columns.Item($c).ColumnWidth = $widths2[$c - 1] }
  }
  # sheet 1: plain range with colour-scale CF on 平均
  Style-Header $w1.Range('A1:I1') $xlThemeColorAccent6 0.4
  Set-Borders $w1.Range("A1:I$l2")
  $w1.Activate(); [void]$w1.Range('G2').Select()
  [void]$w1.Range("G2:G$l2").FormatConditions.AddColorScale(3)
  [void]$w1.Range('A2').Select()
  $xl.ActiveWindow.FreezePanes = $true
  # sheet 2: the same data as an Excel table (ListObject) with calculated columns
  $lo = $w2.ListObjects.Add(1, $w2.Range("A1:I$l2"), $M, 1)
  $lo.Name = [string]$iv.tableName
  $lo.TableStyle = 'TableStyleMedium7'
  $w2.Range("G2:G$l2").Formula = "=AVERAGE($($iv.tableName)[@[評価1]:[評価3]])"
  $w2.Range("H2:H$l2").Formula = '=IF([@平均]>=3.5,"合格","保留")'
  $w1.Activate(); [void]$w1.Range('A2').Select()
  Save-Book $wb 's02.xlsx' $xlOpenXMLWorkbook
  $wb.Close($false)

  # ==============================================================================================================
  # 04_適性検査.xls (Excel 97-2003)
  # ==============================================================================================================
  Step 'aptitude (.xls)'
  $ap = $spec.aptitude
  $n4 = @($ap.rows).Count
  $l4 = 1 + $n4
  $wb = New-Book @([string]$ap.sheet)
  $w = $wb.Worksheets.Item(1)
  Write-Row $w 1 1 @($ap.headers)
  Write-Block $w 2 1 $ap.rows
  $w.Range("A2:A$l4").NumberFormat = '0'
  $w.Range("C2:C$l4").NumberFormat = 'yyyy/m/d'
  $w.Range("D2:H$l4").HorizontalAlignment = $xlCenter
  $h = $w.Range('A1:H1')
  $h.Font.Bold = $true; $h.Interior.Color = 0xDCE6F2; $h.HorizontalAlignment = $xlCenter  # BGR -> F2E6DC
  Set-Borders $w.Range("A1:H$l4")
  $widths4 = @(11, 20, 11, 7, 7, 7, 10, 9)
  for ($c = 1; $c -le $widths4.Count; $c++) { $w.Columns.Item($c).ColumnWidth = $widths4[$c - 1] }
  $wb.CheckCompatibility = $false
  Save-Book $wb 's04.xls' $xlExcel8
  $wb.Close($false)

  # ==============================================================================================================
  # 05_研修出欠.xlsx
  # ==============================================================================================================
  Step 'attendance'
  $at = $spec.attendance
  $n5 = @($at.rows).Count
  $h5 = [int]$at.headerRow
  $f5 = $h5 + 1
  $l5 = $h5 + $n5
  $wb = New-Book @([string]$at.sheet)
  $w = $wb.Worksheets.Item(1)
  $w.Range('A1').Formula = [string]$at.title
  $w.Range('A1:G1').Merge()
  $w.Range('A1').Font.Bold = $true; $w.Range('A1').Font.Size = 13
  $w.Range('A1').HorizontalAlignment = $xlCenter
  Write-Row $w $h5 1 @($at.headers)
  Style-Header $w.Range("A$($h5):G$($h5)") $xlThemeColorAccent4 0.6
  Write-Block $w $f5 1 $at.rows
  $w.Range("C$($f5):F$($l5)").HorizontalAlignment = $xlCenter
  $w.Range("F$($f5):F$($l5)").NumberFormat = '0"回"'
  Set-Borders $w.Range("A$($h5):G$($l5)")
  $w.Activate(); [void]$w.Range("C$f5").Select()
  $fc = $w.Range("C$($f5):E$($l5)").FormatConditions.Add($xlCellValue, $xlEqual, '="×"')
  $fc.Interior.Color = 0xD9D9FF       # BGR -> FFD9D9
  $fc.Font.Color = 0x00009C
  $widths5 = @(14, 16, 11, 11, 11, 9, 26)
  for ($c = 1; $c -le $widths5.Count; $c++) { $w.Columns.Item($c).ColumnWidth = $widths5[$c - 1] }
  [void]$w.Range("A$f5").Select()
  Save-Book $wb 's05.xlsx' $xlOpenXMLWorkbook
  $wb.Close($false)

  Write-EmLine ([ordered]@{ type = 'done' })
}
catch {
  Write-EmLine ([ordered]@{ type = 'error'; message = [string]$_.Exception.Message; at = [string]$_.InvocationInfo.PositionMessage })
  exit 1
}
finally {
  Close-EmExcel $xl
}
