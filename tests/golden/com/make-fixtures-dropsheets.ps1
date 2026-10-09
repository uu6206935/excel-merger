# Builds the Excel-authored fixtures of the single-sheet output (PremergeOptions.dropOtherSheets; owner: patcher P).
#   powershell -NoProfile -ExecutionPolicy Bypass -File tests\golden\com\make-fixtures-dropsheets.ps1 -OutDir tests\fixtures\dropsheets
# ds-features.xlsx : merged sheet 名簿 with formulas that read the other sheets (direct, whole column, 3D, names, a table,
#                    dynamic / legacy arrays, a spill reference), validations (list on another sheet, via a name), conditional
#                    formats that read another sheet (x14), links, charts (own data / another sheet), sparklines, text boxes
#                    linked to cells, a note + a threaded comment, pivot tables on the merged sheet (fed by the sheet itself /
#                    by another sheet), slicers bound to a pivot table of another sheet and to the table of another sheet;
#                    sheets 設定 マスタ 売上 (ListObject) 集計 (formulas + chart) ピボット (pivot table).
# ds-table.xlsx    : merged sheet 名簿 = an Excel Table with a calculated column that uses a name that points at 設定, a totals
#                    row, a table slicer, a pivot table fed by the table + its slicer, a slicer bound to the table of 売上;
#                    sheets 設定 売上.
# Output (ASCII stdout): one line per saved file, "warn ..." for optional features that this Excel could not create.
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
function Try-Do([string]$label, [scriptblock]$sb) {
  try { & $sb } catch { Write-Host ('warn ' + $label + ': ' + (($_.Exception.Message) -replace '[^\x20-\x7e]', '?')) }
}
function NewSheet($wb, [string]$name) {
  $ws = $wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count))
  $ws.Name = $name
  return $ws
}

$xl = $null
try {
  $xl = New-GExcel
  Write-Host "excel pid $($script:GExcelPid)"

  # =============================================================================================
  # ds-features.xlsx
  # =============================================================================================
  $wb = $xl.Workbooks.Add()
  $ws = $wb.Worksheets.Item(1); $ws.Name = '名簿'
  $set = NewSheet $wb '設定'
  $mst = NewSheet $wb 'マスタ'
  $uri = NewSheet $wb '売上'
  $agg = NewSheet $wb '集計'
  $pvs = NewSheet $wb 'ピボット'

  # 設定: 税率 B1, list A2:B4
  SetF $set.Range('A1') '税率'; SetF $set.Range('B1') 0.1
  SetF $set.Range('A2') 'A'; SetF $set.Range('A3') 'B'; SetF $set.Range('A4') 'C'
  SetF $set.Range('B2') 10; SetF $set.Range('B3') 20; SetF $set.Range('B4') 30
  # マスタ: 大学 / コード / 地域, 基準日 in H2
  $h = @('大学', 'コード', '地域')
  for ($c = 0; $c -lt 3; $c++) { SetF $mst.Cells.Item(1, $c + 1) $h[$c] }
  $unis = @('東大', '京大', '阪大', '東北大', '九大'); $regs = @('関東', '関西', '関西', '東北', '九州')
  for ($i = 0; $i -lt 5; $i++) { SetF $mst.Cells.Item($i + 2, 1) $unis[$i]; SetF $mst.Cells.Item($i + 2, 2) ($i + 1); SetF $mst.Cells.Item($i + 2, 3) $regs[$i] }
  SetF $mst.Range('G2') '基準日'; SetF $mst.Range('H2') 46113; $mst.Range('H2').NumberFormat = 'yyyy/m/d'
  # 売上: ListObject 売上
  $h = @('商品', '数量', '単価', '金額')
  for ($c = 0; $c -lt 4; $c++) { SetF $uri.Cells.Item(1, $c + 1) $h[$c] }
  $items = @('りんご', 'みかん', 'ぶどう', 'もも', 'なし')
  for ($i = 0; $i -lt 5; $i++) { SetF $uri.Cells.Item($i + 2, 1) $items[$i]; SetF $uri.Cells.Item($i + 2, 2) (3 + $i); SetF $uri.Cells.Item($i + 2, 3) (100 + 20 * $i) }
  $lo = $uri.ListObjects.Add(1, $uri.Range('A1:D6'), $M, 1); $lo.Name = '売上'
  $uri.Range('D2').Formula = '=[@数量]*[@単価]'
  # names
  [void]$wb.Names.Add('税率', '=設定!$B$1')
  [void]$wb.Names.Add('税込係数', '=1+税率')
  [void]$wb.Names.Add('大学一覧', '=マスタ!$A$2:$A$6')
  [void]$wb.Names.Add('基準日', '=マスタ!$H$2')
  [void]$ws.Names.Add('局所', '=名簿!$C$2:$C$9')
  [void]$ws.Names.Add('局所他', '=設定!$B$2:$B$4')

  # 名簿
  $h = @('ID', '氏名', '点数', '大学')
  for ($c = 0; $c -lt 4; $c++) { SetF $ws.Cells.Item(1, $c + 1) $h[$c] }
  for ($i = 0; $i -lt 8; $i++) {
    $r = $i + 2
    SetF $ws.Cells.Item($r, 1) (1000 + $r); SetF $ws.Cells.Item($r, 2) ('学生' + [char](65 + $i))
    SetF $ws.Cells.Item($r, 3) (50 + ($i * 9) % 50); SetF $ws.Cells.Item($r, 4) $unis[$i % 5]
  }
  $ws.Range('A1:D1').Font.Bold = $true
  $ws.Range('E1').Value2 = '税込'; $ws.Range('F1').Value2 = 'コード'; $ws.Range('G1').Value2 = '係数'; $ws.Range('H1').Value2 = '売上計'
  $ws.Range('I1').Value2 = '税込2'; $ws.Range('J1').Value2 = '倍'; $ws.Range('K1').Value2 = 'マスタ順'; $ws.Range('L1').Value2 = '件数'
  $ws.Range('M1').Value2 = '大学一意'; $ws.Range('N1').Value2 = '配列'; $ws.Range('O1').Value2 = '3D'; $ws.Range('P1').Value2 = '推移'
  $ws.Range('E2').Formula = '=C2*税率'; $ws.Range('E2:E9').FillDown()
  $ws.Range('F2').Formula = '=VLOOKUP(D2,マスタ!$A$2:$C$6,2,FALSE)'; $ws.Range('F2:F9').FillDown()
  $ws.Range('G2').Formula = '=IF(C2>=60,設定!$B$1,0)'; $ws.Range('G2:G9').FillDown()
  $ws.Range('H2').Formula = '=SUM(売上[金額])'; $ws.Range('H2:H9').FillDown()
  $ws.Range('I2').Formula = '=C2*税込係数'; $ws.Range('I2:I9').FillDown()
  $ws.Range('J2').Formula = '=C2*2'; $ws.Range('J2:J9').FillDown()
  Try-Do 'dynamic K2' { $ws.Range('K2').Formula2 = '=SORT(マスタ!A2:A6)' }
  Try-Do 'anchor L2' { $ws.Range('L2').Formula2 = '=COUNTA(K2#)' }
  Try-Do 'dynamic M2' { $ws.Range('M2').Formula2 = '=SORT(UNIQUE(D2:D9))' }
  Try-Do 'cse N2' { $ws.Range('N2:N4').FormulaArray = '=設定!B2:B4*2' }
  $ws.Range('O2').Formula = '=SUM(設定:マスタ!B1)'
  $ws.Range('O3').Formula = '=SUM(名簿!C2:C9)'
  SetF $ws.Range('A11') '合計'
  $ws.Range('C11').Formula = '=SUM(C2:C9)'
  $ws.Range('D11').Formula = '=COUNTIF(マスタ!A:A,"東大")'
  # validations / conditional formats
  Try-Do 'dv list other sheet' { $ws.Range('D2:D9').Validation.Add(3, 1, 1, '=マスタ!$A$2:$A$6') }
  Try-Do 'dv list name' { $ws.Range('B2:B9').Validation.Add(3, 1, 1, '=大学一覧') }
  Try-Do 'dv whole' { $ws.Range('C2:C9').Validation.Add(1, 1, 1, '0', '100') }
  Try-Do 'dv whole other sheet' { $ws.Range('J2:J9').Validation.Add(1, 1, 7, '=設定!$B$2') }
  Try-Do 'cf cross sheet' { $fc = $ws.Range('C2:C9').FormatConditions.Add(2, $M, '=$C2>設定!$B$2'); $fc.Interior.Color = 0xCCFFCC }
  Try-Do 'cf own' { $fc2 = $ws.Range('E2:E9').FormatConditions.Add(1, 5, '=5'); $fc2.Font.Bold = $true }
  Try-Do 'cf bar' { $db = $ws.Range('F2:F9').FormatConditions.AddDatabar() }
  Try-Do 'cf bar other' { $db2 = $ws.Range('G2:G9').FormatConditions.AddDatabar(); $db2.MinPoint.Modify(4, '=設定!$B$1') }
  Try-Do 'cf icons' { $ic = $ws.Range('J2:J9').FormatConditions.AddIconSetCondition(); $ic.IconCriteria.Item(2).Type = 4; $ic.IconCriteria.Item(2).Value = '=設定!$B$3' }
  Try-Do 'cf name' { $fc3 = $ws.Range('H2:H9').FormatConditions.Add(2, $M, '=$H2>税率'); $fc3.Font.Color = 0x0000FF }
  # links, notes
  Try-Do 'link to sheet' { [void]$ws.Hyperlinks.Add($ws.Range('A12'), '', '設定!A1', $M, '設定へ') }
  Try-Do 'link to own' { [void]$ws.Hyperlinks.Add($ws.Range('A13'), '', '名簿!B2', $M, '先頭へ') }
  Try-Do 'link url' { [void]$ws.Hyperlinks.Add($ws.Range('A14'), 'https://example.com/', $M, $M, 'web') }
  Try-Do 'note' { [void]$ws.Range('A1').AddComment('メモ') }
  Try-Do 'threaded' { [void]$ws.Range('B2').AddCommentThreaded('スレッド') }
  # charts
  Try-Do 'chart own' { $co = $ws.ChartObjects().Add(900, 10, 260, 160); $co.Chart.ChartType = 51; $co.Chart.SetSourceData($ws.Range('B1:C9')) }
  Try-Do 'chart other' { $co2 = $ws.ChartObjects().Add(900, 190, 260, 160); $co2.Chart.ChartType = 51; $co2.Chart.SetSourceData($mst.Range('A1:B6')) }
  Try-Do 'chart mixed' { $co3 = $ws.ChartObjects().Add(900, 370, 260, 160); $co3.Chart.ChartType = 51; $s = $co3.Chart.SeriesCollection().NewSeries(); $s.Values = $ws.Range('C2:C9'); $s2 = $co3.Chart.SeriesCollection().NewSeries(); $s2.Values = $set.Range('B2:B4') }
  # sparklines, text boxes
  Try-Do 'sparkline other' { [void]$ws.Range('P2').SparklineGroups.Add(1, 'マスタ!B2:B6') }
  Try-Do 'sparkline own' { [void]$ws.Range('P3').SparklineGroups.Add(1, '名簿!C2:C9') }
  Try-Do 'textbox other' { $tb = $ws.Shapes.AddTextbox(1, 420, 300, 120, 24); $tb.DrawingObject.Formula = '=設定!$A$1' }
  Try-Do 'textbox own' { $tb2 = $ws.Shapes.AddTextbox(1, 420, 340, 120, 24); $tb2.DrawingObject.Formula = '=名簿!$A$1' }
  # page setup free: print area / titles through names
  [void]$ws.Names.Add('Print_Area', '=名簿!$A$1:$P$14')
  [void]$ws.Names.Add('Print_Titles', '=名簿!$1:$1')
  [void]$ws.Range('A1:D9').AutoFilter(1)
  $ws.Activate(); $ws.Range('B2').Select(); $xl.ActiveWindow.FreezePanes = $true

  # 集計: formulas + chart
  $agg.Range('A1').Formula = '=SUM(名簿!C2:C9)'; $agg.Range('A2').Formula = '=SUM(売上[金額])'; $agg.Range('A3').Formula = '=税率*100'
  Try-Do 'chart agg' { $co4 = $agg.ChartObjects().Add(200, 20, 300, 180); $co4.Chart.ChartType = 51; $co4.Chart.SetSourceData($ws.Range('B1:C9')) }

  # pivot tables
  $pcache = $null
  Try-Do 'pivot on ピボット' {
    $script:pcache = $wb.PivotCaches().Create(1, '売上')
    $script:pt1 = $script:pcache.CreatePivotTable($pvs.Range('A3'), 'PT売上')
    $script:pt1.PivotFields('商品').Orientation = 1
    [void]$script:pt1.AddDataField($script:pt1.PivotFields('金額'), '合計金額', -4157)
  }
  Try-Do 'pivot on merged (other source)' {
    $pcb = $wb.PivotCaches().Create(1, 'マスタ!R1C1:R6C3')
    $ptb = $pcb.CreatePivotTable($ws.Range('A30'), 'PT他')
    $ptb.PivotFields('地域').Orientation = 1
    [void]$ptb.AddDataField($ptb.PivotFields('コード'), '合計コード', -4157)
  }
  Try-Do 'pivot on merged (own source)' {
    $pco = $wb.PivotCaches().Create(1, '名簿!R1C1:R9C4')
    $pto = $pco.CreatePivotTable($ws.Range('H30'), 'PT自')
    $pto.PivotFields('大学').Orientation = 1
    [void]$pto.AddDataField($pto.PivotFields('点数'), '合計点', -4157)
  }
  # slicers on the merged sheet: one bound to the pivot table of ピボット, one to the table 売上
  Try-Do 'slicer pivot' {
    $sc = $wb.SlicerCaches.Add2($script:pt1, '商品')
    [void]$sc.Slicers.Add($ws, $M, 'スライサー商品', '商品', 560, 700, 120, 120)
  }
  Try-Do 'slicer table' {
    $sc2 = $wb.SlicerCaches.Add2($lo, '商品')
    [void]$sc2.Slicers.Add($ws, $M, 'スライサー表', '表の商品', 560, 840, 120, 120)
  }
  $ws.Activate(); $ws.Range('A1').Select()
  Save-Book $wb 'ds-features.xlsx'

  # =============================================================================================
  # ds-table.xlsx
  # =============================================================================================
  $wb = $xl.Workbooks.Add()
  $ws = $wb.Worksheets.Item(1); $ws.Name = '名簿'
  $set = NewSheet $wb '設定'
  $uri = NewSheet $wb '売上'
  SetF $set.Range('A1') '税率'; SetF $set.Range('B1') 0.1
  $h = @('商品', '数量', '単価', '金額')
  for ($c = 0; $c -lt 4; $c++) { SetF $uri.Cells.Item(1, $c + 1) $h[$c] }
  for ($i = 0; $i -lt 4; $i++) { SetF $uri.Cells.Item($i + 2, 1) $items[$i]; SetF $uri.Cells.Item($i + 2, 2) (3 + $i); SetF $uri.Cells.Item($i + 2, 3) (100 + 20 * $i) }
  $lo2 = $uri.ListObjects.Add(1, $uri.Range('A1:D5'), $M, 1); $lo2.Name = '売上'
  $uri.Range('D2').Formula = '=[@数量]*[@単価]'
  [void]$wb.Names.Add('税率', '=設定!$B$1')
  $h = @('ID', '氏名', '大学', '点数')
  for ($c = 0; $c -lt 4; $c++) { SetF $ws.Cells.Item(1, $c + 1) $h[$c] }
  for ($i = 0; $i -lt 6; $i++) {
    $r = $i + 2
    SetF $ws.Cells.Item($r, 1) ('S00' + ($i + 1)); SetF $ws.Cells.Item($r, 2) ('学生' + [char](65 + $i)); SetF $ws.Cells.Item($r, 3) $unis[$i % 5]; SetF $ws.Cells.Item($r, 4) (55 + $i * 6)
  }
  $lot = $ws.ListObjects.Add(1, $ws.Range('A1:D7'), $M, 1); $lot.Name = 'tbl名簿'
  $ws.Range('E1').Value2 = '税込'
  $ws.Range('E2').Formula = '=[@点数]*(1+税率)'
  $ws.Range('F1').Value2 = '売上計'
  $ws.Range('F2').Formula = '=SUM(売上[金額])'
  $lot.ShowTotals = $true
  Try-Do 'totals' { $lot.ListColumns.Item('点数').TotalsCalculation = 1; $lot.ListColumns.Item('税込').TotalsCalculation = 1 }
  Try-Do 'table slicer own' {
    $sc = $wb.SlicerCaches.Add2($lot, '大学')
    [void]$sc.Slicers.Add($ws, $M, 'スライサー大学', '大学', 300, 700, 120, 120)
  }
  Try-Do 'pivot own' {
    $pc = $wb.PivotCaches().Create(1, 'tbl名簿')
    $pt = $pc.CreatePivotTable($ws.Range('A14'), 'PT名簿')
    $pt.PivotFields('大学').Orientation = 1
    [void]$pt.AddDataField($pt.PivotFields('点数'), '合計点', -4157)
    $sc3 = $wb.SlicerCaches.Add2($pt, '大学')
    [void]$sc3.Slicers.Add($ws, $M, 'スライサー大学P', '大学P', 450, 700, 120, 120)
  }
  Try-Do 'table slicer other' {
    $sc2 = $wb.SlicerCaches.Add2($lo2, '商品')
    [void]$sc2.Slicers.Add($ws, $M, 'スライサー商品', '商品', 300, 840, 120, 120)
  }
  $ws.Activate(); $ws.Range('A1').Select()
  Save-Book $wb 'ds-table.xlsx'
}
finally { Close-GExcel $xl }
