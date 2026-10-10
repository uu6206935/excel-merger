# Builds the Excel-authored fixtures of the sheet copier's OBJECT phase (PremergeOptions.includeSourceSheets, round 9, task A2):
#   powershell -NoProfile -ExecutionPolicy Bypass -File tests\golden\com\make-fixtures-addsheets-objects.ps1 -OutDir tests\fixtures\addsheets-objects [-Only regex]
# ob-charts.xlsx     データ (month table) with a column + line combo chart (secondary axis, title linked to a cell, series names linked to
#                    cells) embedded on it; グラフ sheet with a combo chart reading データ, a pie chart and two new-style charts (AddChart2: explicit
#                    theme colours, chart style / colour parts); グラフ用シート (a chart sheet).
# ob-shapes.xlsx     図形: a picture (also as a second copy of the same image), a picture with a hyperlink to another sheet, a text box with a
#                    cell link, a rectangle with a hyperlink to another sheet, a group of two shapes, a connector, a line; 背景 sheet with a
#                    sheet background picture.
# ob-comments.xlsx   コメント: notes (one always visible, one with a long text and another author) + a threaded comment with a reply;
#                    メモのみ: notes only.
# ob-table.xlsx      テーブル: ListObject T売上 (calculated column, totals row with SUM, custom table style), a second table T小計; 集計 sheet
#                    with structured references (column, [#Totals], [#This Row] style, INDEX), a defined name over a table column, a chart
#                    reading the table.
# ob-pivot.xlsx      元データ (range) + 元テーブル (ListObject T元): pivot on a range source (+ a second pivot sharing the cache, number formats),
#                    pivot on a Table source, a pivot chart.
# ob-slicer-control.xlsx  a table with a slicer, a pivot with a slicer, form controls (check box, button), an ActiveX check box, notes on the
#                    same sheet as the controls (VML shared by notes and controls).
# ob-strict.xlsx    the same kinds of objects saved as Strict Open XML Spreadsheet (purl.oclc.org namespaces).
# Output (ASCII stdout): one line per saved file, "warn ..." for optional features this Excel could not create.
# Own Excel instance only (tests/golden/com/common.ps1); no clipboard, no PageSetup.
param([string]$OutDir, [string]$Only = '')
$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\common.ps1"
$M = $script:GMissing
$OutDir = (Resolve-Path $OutDir).Path
$Scratch = Join-Path $env:TEMP ('ob-fixtures-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Force $Scratch | Out-Null

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
function Want([string]$name) { return (-not $Only) -or ($name -match $Only) }

# a small PNG made with System.Drawing (gradient + text), so that no binary asset has to be kept in the repository
function Make-Png([string]$path, [int]$w, [int]$h, [string]$text) {
  Add-Type -AssemblyName System.Drawing
  $bmp = New-Object System.Drawing.Bitmap $w, $h
  $g = [System.Drawing.Graphics]::FromImage($bmp)
  $br = New-Object System.Drawing.Drawing2D.LinearGradientBrush ((New-Object System.Drawing.Point 0, 0), (New-Object System.Drawing.Point $w, $h), [System.Drawing.Color]::FromArgb(255, 60, 120, 200), [System.Drawing.Color]::FromArgb(255, 240, 160, 60))
  $g.FillRectangle($br, 0, 0, $w, $h)
  $f = New-Object System.Drawing.Font 'Arial', 14
  $g.DrawString($text, $f, [System.Drawing.Brushes]::White, 6, 6)
  $g.Dispose(); $bmp.Save($path, [System.Drawing.Imaging.ImageFormat]::Png); $bmp.Dispose()
}

$xl = $null
try {
  $xl = New-GExcel
  Write-Host "excel pid $($script:GExcelPid) version $($xl.Version)"
  $png1 = Join-Path $Scratch 'logo.png'; Make-Png $png1 160 90 'LOGO'
  $png2 = Join-Path $Scratch 'banner.png'; Make-Png $png2 240 60 'BANNER'

  # =============================================================================================
  # ob-charts.xlsx
  # =============================================================================================
  if (Want 'ob-charts') {
    $wb = $xl.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1); $ws.Name = 'データ'
    $h = @('月', '売上', '利益', '利益率')
    for ($c = 0; $c -lt 4; $c++) { $ws.Cells.Item(1, $c + 1).Value2 = $h[$c] }
    $months = @('1月', '2月', '3月', '4月', '5月', '6月')
    $sales = @(120, 135, 128, 160, 172, 190)
    $profit = @(24, 31, 22, 40, 45, 57)
    for ($i = 0; $i -lt 6; $i++) {
      $r = $i + 2
      $ws.Cells.Item($r, 1).Value2 = $months[$i]
      SetF $ws.Cells.Item($r, 2) $sales[$i]
      SetF $ws.Cells.Item($r, 3) $profit[$i]
      $ws.Cells.Item($r, 4).Formula = "=C$r/B$r"
    }
    $ws.Range('D2:D7').NumberFormat = '0.0%'
    $ws.Range('F1').Value2 = '売上と利益の推移'
    $gr = NewSheet $wb 'グラフ'
    $gr.Range('A1').Value2 = '別シートのグラフ'
    # (1) embedded on the data sheet: column + line on the secondary axis
    Try-Do 'embedded combo chart' {
      $co = $ws.ChartObjects().Add(260, 20, 380, 230)
      $co.Name = '埋め込みグラフ'
      $ch = $co.Chart
      $ch.ChartType = 51
      $ch.SetSourceData($ws.Range('A1:C7'))
      $s3 = $ch.SeriesCollection().NewSeries()
      $s3.Name = '=データ!$D$1'
      $s3.Values = $ws.Range('D2:D7')
      $s3.XValues = $ws.Range('A2:A7')
      $s3.ChartType = 65
      $s3.AxisGroup = 2
      $ch.HasTitle = $true
      $ch.ChartTitle.Text = '売上と利益'
      Try-Do 'linked title' { $ch.ChartTitle.Formula = '=データ!$F$1' }
      $ch.HasLegend = $true
    }
    # (2) on its own worksheet, reading データ: column + line combo and a pie chart
    Try-Do 'own-sheet combo chart' {
      $co = $gr.ChartObjects().Add(10, 40, 420, 250)
      $co.Name = '別シートのコンボ'
      $ch = $co.Chart
      $ch.ChartType = 51
      while ($ch.SeriesCollection().Count -gt 0) { $ch.SeriesCollection(1).Delete() | Out-Null }
      $a = $ch.SeriesCollection().NewSeries()
      $a.Name = '=データ!$B$1'; $a.Values = $ws.Range('B2:B7'); $a.XValues = $ws.Range('A2:A7')
      $b = $ch.SeriesCollection().NewSeries()
      $b.Name = '=データ!$C$1'; $b.Values = $ws.Range('C2:C7'); $b.XValues = $ws.Range('A2:A7'); $b.ChartType = 4
      $ch.HasTitle = $true; $ch.ChartTitle.Text = 'コンボ（別シート）'
    }
    # new-style charts (AddChart2): explicit theme colours on every series, chart style / colour parts
    Try-Do 'chart2 column' {
      $s = $gr.Shapes.AddChart2(201, 51, 10, 320, 380, 230)
      $s.Name = '新スタイル縦棒'
      $s.Chart.SetSourceData($ws.Range('A1:C7'))
      $s.Chart.HasTitle = $true; $s.Chart.ChartTitle.Text = '新スタイルの縦棒'
    }
    Try-Do 'chart2 pie' {
      $s = $gr.Shapes.AddChart2(251, 5, 420, 320, 280, 230)
      $s.Name = '新スタイル円'
      $s.Chart.SetSourceData($ws.Range('A1:B7'))
    }
    Try-Do 'pie chart' {
      $co = $gr.ChartObjects().Add(450, 40, 280, 250)
      $co.Name = '円グラフ'
      $ch = $co.Chart
      $ch.ChartType = 5
      $ch.SetSourceData($ws.Range('A1:B7'))
      $ch.HasTitle = $true; $ch.ChartTitle.Text = '売上の構成'
    }
    # (3) a chart sheet
    Try-Do 'chart sheet' {
      $cs = $wb.Charts.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count))
      $cs.Name = 'グラフ用シート'
      $cs.ChartType = 57
      while ($cs.SeriesCollection().Count -gt 0) { $cs.SeriesCollection(1).Delete() | Out-Null }
      $a = $cs.SeriesCollection().NewSeries()
      $a.Name = '=データ!$B$1'; $a.Values = $ws.Range('B2:B7'); $a.XValues = $ws.Range('A2:A7')
      $cs.HasTitle = $true; $cs.ChartTitle.Text = 'グラフシート'
    }
    Try-Do 'select' { $ws.Activate(); $ws.Range('A1').Select() | Out-Null }
    Save-Book $wb 'ob-charts.xlsx'
  }

  # =============================================================================================
  # ob-shapes.xlsx
  # =============================================================================================
  if (Want 'ob-shapes') {
    $wb = $xl.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1); $ws.Name = '図形'
    $tg = NewSheet $wb '飛び先'
    $bg = NewSheet $wb '背景'
    $tg.Range('A1').Value2 = '飛び先シート'
    $ws.Range('A1').Value2 = '図形の見本'; $ws.Range('A3').Value2 = 100; $ws.Range('A3').NumberFormat = '#,##0'
    Try-Do 'picture' { $p = $ws.Shapes.AddPicture($png1, 0, -1, 20, 40, 160, 90); $p.Name = 'ロゴ' }
    Try-Do 'picture again' { $p = $ws.Shapes.AddPicture($png1, 0, -1, 200, 40, 80, 45); $p.Name = 'ロゴ小' }
    Try-Do 'picture 2' { $p = $ws.Shapes.AddPicture($png2, 0, -1, 300, 40, 240, 60); $p.Name = 'バナー' }
    Try-Do 'picture hyperlink' {
      $p = $ws.Shapes.AddPicture($png2, 0, -1, 20, 150, 120, 30); $p.Name = '飛ぶ画像'
      [void]$ws.Hyperlinks.Add($p, '', "'飛び先'!A1", 'ツールチップ')
    }
    Try-Do 'text box' {
      $t = $ws.Shapes.AddTextbox(1, 20, 200, 180, 50); $t.Name = 'テキストボックス'
      $t.TextFrame2.TextRange.Text = 'メモ：これはテキストボックス'
      $t.Fill.ForeColor.RGB = (Rgb 255 242 204)
    }
    Try-Do 'linked text box' {
      $t = $ws.Shapes.AddTextbox(1, 220, 200, 120, 30); $t.Name = 'リンクテキスト'
      $t.DrawingObject.Formula = '=図形!$A$3'
    }
    Try-Do 'rectangle with link' {
      $r = $ws.Shapes.AddShape(5, 360, 150, 120, 40); $r.Name = '角丸ボタン'
      $r.TextFrame2.TextRange.Text = '飛び先へ'
      [void]$ws.Hyperlinks.Add($r, '', "'飛び先'!B2", '飛び先へ移動')
    }
    Try-Do 'group' {
      $a = $ws.Shapes.AddShape(1, 20, 280, 60, 40); $a.Name = '群A'
      $b = $ws.Shapes.AddShape(9, 90, 280, 60, 40); $b.Name = '群B'
      $g = $ws.Shapes.Range([object[]]@('群A', '群B')).Group(); $g.Name = 'グループ1'
    }
    Try-Do 'connector' {
      $a = $ws.Shapes.AddShape(1, 200, 280, 60, 40); $a.Name = '接続A'
      $b = $ws.Shapes.AddShape(1, 320, 330, 60, 40); $b.Name = '接続B'
      $c = $ws.Shapes.AddConnector(2, 0, 0, 10, 10); $c.Name = 'コネクタ'
      $c.ConnectorFormat.BeginConnect($a, 4); $c.ConnectorFormat.EndConnect($b, 2)
    }
    Try-Do 'line' { $l = $ws.Shapes.AddLine(400, 280, 520, 320); $l.Name = '直線'; $l.Line.Weight = 3 }
    Try-Do 'background' { $bg.SetBackgroundPicture($png2); $bg.Range('A1').Value2 = '背景画像のシート' }
    Try-Do 'select' { $ws.Activate(); $ws.Range('A1').Select() | Out-Null }
    Save-Book $wb 'ob-shapes.xlsx'
  }

  # =============================================================================================
  # ob-comments.xlsx
  # =============================================================================================
  if (Want 'ob-comments') {
    $wb = $xl.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1); $ws.Name = 'コメント'
    $mo = NewSheet $wb 'メモのみ'
    $ws.Range('A1').Value2 = '氏名'; $ws.Range('B1').Value2 = '点数'
    for ($i = 0; $i -lt 6; $i++) { $ws.Cells.Item($i + 2, 1).Value2 = ('受講者' + [char](65 + $i)); SetF $ws.Cells.Item($i + 2, 2) (60 + 5 * $i) }
    Try-Do 'note 1' { $c = $ws.Range('B2').AddComment('確認済み（メモ）'); $c.Visible = $false }
    Try-Do 'note visible' { $c = $ws.Range('B3').AddComment('常に表示されるメモ'); $c.Visible = $true; $c.Shape.Width = 150; $c.Shape.Height = 40 }
    Try-Do 'note long' { $c = $ws.Range('A4').AddComment("長いメモ`r`n2行目`r`n3行目"); $c.Shape.Height = 80 }
    Try-Do 'threaded' {
      [void]$ws.Range('B5').AddCommentThreaded('スレッドコメントの本文')
      [void]$ws.Range('B5').CommentThreaded.AddReply('返信です')
    }
    Try-Do 'threaded 2' { [void]$ws.Range('B6').AddCommentThreaded('もう一つのスレッド') }
    $mo.Range('A1').Value2 = 'メモのみ'
    Try-Do 'note m1' { [void]$mo.Range('A1').AddComment('メモだけのシート') }
    Try-Do 'note m2' { [void]$mo.Range('C3').AddComment('二つ目') }
    Try-Do 'select' { $ws.Activate(); $ws.Range('A1').Select() | Out-Null }
    Save-Book $wb 'ob-comments.xlsx'
  }

  # =============================================================================================
  # ob-table.xlsx
  # =============================================================================================
  if (Want 'ob-table') {
    $wb = $xl.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1); $ws.Name = 'テーブル'
    $sm = NewSheet $wb '集計'
    $h = @('商品', '数量', '単価', '金額')
    for ($c = 0; $c -lt 4; $c++) { $ws.Cells.Item(1, $c + 1).Value2 = $h[$c] }
    $items = @('りんご', 'みかん', 'ぶどう', 'もも', 'なし', 'かき')
    for ($i = 0; $i -lt 6; $i++) { $r = $i + 2; $ws.Cells.Item($r, 1).Value2 = $items[$i]; SetF $ws.Cells.Item($r, 2) (3 + $i); SetF $ws.Cells.Item($r, 3) (100 + 20 * $i) }
    Try-Do 'custom table style' {
      $ts = $wb.TableStyles.Add('マイ表スタイル')
      $ts.ShowAsAvailableTableStyle = $true
      $ts.TableStyleElements.Item(0).Interior.Color = (Rgb 235 245 255)
      $ts.TableStyleElements.Item(1).Interior.Color = (Rgb 31 78 121)
      $ts.TableStyleElements.Item(1).Font.Color = (Rgb 255 255 255)
      $ts.TableStyleElements.Item(1).Font.Bold = $true
    }
    $lo = $ws.ListObjects.Add(1, $ws.Range('A1:D7'), $null, 1)
    $lo.Name = 'T売上'
    $lo.ListColumns.Item(4).DataBodyRange.Formula = '=[@数量]*[@単価]'
    $lo.ShowTotals = $true
    $lo.ListColumns.Item(2).TotalsCalculation = 1
    $lo.ListColumns.Item(4).TotalsCalculation = 1
    Try-Do 'custom style on table' { $lo.TableStyle = 'マイ表スタイル' }
    $ws.Range('F1').Value2 = '小計用'; $ws.Range('F2').Value2 = 'キー'; $ws.Range('G2').Value2 = '値'; $ws.Range('F3').Value2 = 'a'; SetF $ws.Range('G3') 1; $ws.Range('F4').Value2 = 'b'; SetF $ws.Range('G4') 2
    $l2 = $ws.ListObjects.Add(1, $ws.Range('F2:G4'), $null, 1)
    $l2.Name = 'T小計'
    $l2.TableStyle = 'TableStyleLight9'
    [void]$wb.Names.Add('金額列', '=T売上[金額]')
    $sm.Range('A1').Formula = '=SUM(T売上[金額])'
    $sm.Range('A2').Formula = '=T売上[[#Totals],[金額]]'
    $sm.Range('A3').Formula = '=INDEX(T売上[商品],2)'
    $sm.Range('A4').Formula = '=SUM(金額列)'
    $sm.Range('A5').Formula = '=ROWS(T売上)'
    $sm.Range('A6').Formula = '=SUM(T小計[値])'
    $sm.Range('A7').Formula = '=COUNTA(テーブル!A:A)'
    Try-Do 'chart on a table' {
      $co = $sm.ChartObjects().Add(150, 10, 300, 180)
      $co.Name = 'テーブルのグラフ'
      $ch = $co.Chart
      $ch.ChartType = 51
      $ch.SetSourceData($ws.Range('A1:A7,D1:D7'))
      $ch.HasTitle = $true; $ch.ChartTitle.Text = '金額'
    }
    Try-Do 'select' { $ws.Activate(); $ws.Range('A1').Select() | Out-Null }
    Save-Book $wb 'ob-table.xlsx'
  }

  # =============================================================================================
  # ob-pivot.xlsx
  # =============================================================================================
  if (Want 'ob-pivot') {
    $wb = $xl.Workbooks.Add()
    $src = $wb.Worksheets.Item(1); $src.Name = '元データ'
    $tsrc = NewSheet $wb '元テーブル'
    $pv = NewSheet $wb '集計1'
    $pv2 = NewSheet $wb '集計2'
    $h = @('店舗', '商品', '日付', '売上', '数量')
    $shops = @('東京', '大阪', '名古屋')
    $goods = @('りんご', 'みかん', 'ぶどう', 'もも')
    foreach ($sh in @($src, $tsrc)) {
      for ($c = 0; $c -lt 5; $c++) { $sh.Cells.Item(1, $c + 1).Value2 = $h[$c] }
      for ($i = 0; $i -lt 24; $i++) {
        $r = $i + 2
        $sh.Cells.Item($r, 1).Value2 = $shops[$i % 3]
        $sh.Cells.Item($r, 2).Value2 = $goods[$i % 4]
        SetF $sh.Cells.Item($r, 3) (45658 + $i * 9)
        SetF $sh.Cells.Item($r, 4) (1000 + 130 * (($i * 7) % 11))
        SetF $sh.Cells.Item($r, 5) (1 + ($i * 5) % 7)
      }
      $sh.Range('C2:C25').NumberFormat = 'yyyy/m/d'
      $sh.Range('D2:D25').NumberFormat = '#,##0'
    }
    $lo = $tsrc.ListObjects.Add(1, $tsrc.Range('A1:E25'), $null, 1)
    $lo.Name = 'T元'
    # (1) range source, two pivots sharing one cache
    $pc = $wb.PivotCaches().Create(1, '元データ!R1C1:R25C5')
    $pt = $pc.CreatePivotTable('集計1!R3C1', 'ピボット範囲')
    $pt.PivotFields('店舗').Orientation = 1
    $pt.PivotFields('商品').Orientation = 2
    $df = $pt.AddDataField($pt.PivotFields('売上'), '売上合計', -4157)
    $df.NumberFormat = '#,##0"円"'
    $pt2 = $pc.CreatePivotTable('集計1!R3C9', 'ピボット範囲2')
    $pt2.PivotFields('商品').Orientation = 1
    $d2 = $pt2.AddDataField($pt2.PivotFields('数量'), '数量合計', -4157)
    $d2.NumberFormat = '0.0'
    # (2) Table source
    $pc3 = $wb.PivotCaches().Create(1, 'T元')
    $pt3 = $pc3.CreatePivotTable('集計2!R3C1', 'ピボット表')
    $pt3.PivotFields('商品').Orientation = 1
    [void]$pt3.AddDataField($pt3.PivotFields('売上'), '売上の合計', -4157)
    [void]$pt3.AddDataField($pt3.PivotFields('数量'), '数量の平均', -4106)
    # (3) a pivot chart
    Try-Do 'pivot chart' {
      $co = $pv.ChartObjects().Add(20, 200, 360, 220)
      $co.Name = 'ピボットグラフ'
      $co.Chart.SetSourceData($pt.TableRange1)
      $co.Chart.ChartType = 51
    }
    Try-Do 'select' { $pv.Activate(); $pv.Range('A1').Select() | Out-Null }
    Save-Book $wb 'ob-pivot.xlsx'
  }

  # =============================================================================================
  # ob-slicer-control.xlsx
  # =============================================================================================
  if (Want 'ob-slicer-control') {
    $wb = $xl.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1); $ws.Name = 'スライサー'
    $pvs = NewSheet $wb 'ピボット'
    $h = @('店舗', '商品', '売上')
    for ($c = 0; $c -lt 3; $c++) { $ws.Cells.Item(1, $c + 1).Value2 = $h[$c] }
    $shops = @('東京', '大阪', '名古屋')
    for ($i = 0; $i -lt 9; $i++) { $r = $i + 2; $ws.Cells.Item($r, 1).Value2 = $shops[$i % 3]; $ws.Cells.Item($r, 2).Value2 = ('商品' + [char](65 + $i % 4)); SetF $ws.Cells.Item($r, 3) (100 * ($i + 1)) }
    $lo = $ws.ListObjects.Add(1, $ws.Range('A1:C10'), $null, 1)
    $lo.Name = 'T店舗'
    Try-Do 'table slicer' {
      $sc = $wb.SlicerCaches.Add2($lo, '店舗')
      [void]$sc.Slicers.Add($ws, $M, 'スライサー店舗', '店舗', 10, 300, 120, 100)
    }
    $pc = $wb.PivotCaches().Create(1, 'T店舗')
    $pt = $pc.CreatePivotTable('ピボット!R3C1', 'ピボット店舗')
    $pt.PivotFields('商品').Orientation = 1
    [void]$pt.AddDataField($pt.PivotFields('売上'), '売上計', -4157)
    Try-Do 'pivot slicer' {
      $sc = $wb.SlicerCaches.Add2($pt, '店舗')
      [void]$sc.Slicers.Add($pvs, $M, 'スライサーP', '店舗', 10, 300, 120, 100)
    }
    Try-Do 'note' { [void]$ws.Range('A12').AddComment('コントロールと同じシートのメモ') }
    Try-Do 'check box' { $cb = $ws.CheckBoxes().Add(20, 280, 90, 20); $cb.Caption = 'チェック'; $cb.LinkedCell = '$E$1' }
    Try-Do 'button' { $bt = $ws.Buttons().Add(20, 320, 90, 28); $bt.Caption = 'ボタン' }
    Try-Do 'activex' { $o = $ws.Shapes.AddOLEObject('Forms.CheckBox.1', '', $false, $false, '', 0, '', 140, 280, 100, 24) }
    Try-Do 'select' { $ws.Activate(); $ws.Range('A1').Select() | Out-Null }
    Save-Book $wb 'ob-slicer-control.xlsx'
  }

  # =============================================================================================
  # ob-strict.xlsx (Strict Open XML Spreadsheet: other namespaces for the sheet, the drawing, the chart ...)
  # =============================================================================================
  if (Want 'ob-strict') {
    $wb = $xl.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1); $ws.Name = 'データ'
    $pv = NewSheet $wb '集計'
    $h = @('月', '売上', '数量')
    for ($c = 0; $c -lt 3; $c++) { $ws.Cells.Item(1, $c + 1).Value2 = $h[$c] }
    for ($i = 0; $i -lt 6; $i++) { $r = $i + 2; $ws.Cells.Item($r, 1).Value2 = ('' + ($i + 1) + '月'); SetF $ws.Cells.Item($r, 2) (100 + 20 * $i); SetF $ws.Cells.Item($r, 3) (3 + $i) }
    $lo = $ws.ListObjects.Add(1, $ws.Range('A1:C7'), $null, 1)
    $lo.Name = 'T厳密'
    Try-Do 'strict chart' {
      $co = $ws.ChartObjects().Add(260, 20, 320, 200)
      $co.Chart.ChartType = 51
      $co.Chart.SetSourceData($ws.Range('A1:B7'))
      $co.Chart.HasTitle = $true; $co.Chart.ChartTitle.Text = '厳密形式のグラフ'
    }
    Try-Do 'strict new-style chart' {
      $s = $ws.Shapes.AddChart2(201, 51, 260, 240, 320, 200)
      $s.Name = '厳密新スタイル'
      $s.Chart.SetSourceData($ws.Range('A1:C7'))
    }
    Try-Do 'strict picture' { $p = $ws.Shapes.AddPicture($png1, 0, -1, 20, 200, 120, 68); $p.Name = '厳密ロゴ' }
    Try-Do 'strict text box' { $t = $ws.Shapes.AddTextbox(1, 150, 200, 150, 40); $t.Name = '厳密テキスト'; $t.TextFrame2.TextRange.Text = '厳密形式' }
    Try-Do 'strict note' { [void]$ws.Range('B2').AddComment('厳密形式のメモ') }
    $pc = $wb.PivotCaches().Create(1, 'T厳密')
    $pt = $pc.CreatePivotTable('集計!R3C1', 'ピボット厳密')
    $pt.PivotFields('月').Orientation = 1
    [void]$pt.AddDataField($pt.PivotFields('売上'), '売上計', -4157)
    Try-Do 'select' { $ws.Activate(); $ws.Range('A1').Select() | Out-Null }
    # xlStrictOpenXMLSpreadsheet = 61
    Save-Book $wb 'ob-strict.xlsx' 61
  }
}
finally {
  Close-GExcel $xl
  try { Remove-Item $Scratch -Recurse -Force -ErrorAction SilentlyContinue } catch {}
}
