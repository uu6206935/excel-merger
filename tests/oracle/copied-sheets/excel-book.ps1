# Excel half of the copied-sheets oracle (task V9): ONE hidden Excel instance of our own for a whole batch of workbooks.
#   powershell -NoProfile -ExecutionPolicy Bypass -File excel-book.ps1 -JobsFile <jobs.json> -OutDir <dir>
# jobs.json: [{id, file, deep}] where deep = "all" | "none" | ["sheet name", ...] (the sheets that also get the "deep" dump below)
# For every job a file <OutDir>\<id>.book.json is written:
#   {"open":"normal|repair|failed","log":"..","date1904":bool,"names":[{"name","refers","visible"}],
#    "sheets":[{"name","type","visible","rows","cols","trunc","circ","v":[[..]],"f":[[..]],"deep":{..}}]}
#   type    : -4167 worksheet, -4109 chart sheet, 3 / 4 macro sheet, -4116 dialog sheet (Excel's XlSheetType)
#   visible : -1 visible, 0 hidden, 2 very hidden
#   v / f   : Value2 / Formula of the used range from A1 after CalculateFull (cell = null | true/false | number | "text" | {"e":<excel error code>}),
#             capped (trunc = true) at 150000 cells
#   deep    : tab colour, column widths (points) of the first 40 columns, hidden columns / rows, conditional-format count, validation cells,
#             hyperlink count, 40 sampled cells (number format, font, fill, alignment, wrap, merge, borders), shapes, charts (series / values),
#             tables, notes, pivot tables, sheet-scoped names
# The cell-by-cell comparison itself is done in Node (judge.ts). ASCII-only stdout. Only our own Excel PID is ever killed (golden\com\common.ps1).
# No clipboard, no PageSetup.
param([Parameter(Mandatory = $true)][string]$JobsFile, [Parameter(Mandatory = $true)][string]$OutDir)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '..\..\golden\com\common.ps1')
$M = $script:GMissing
$inv = [Globalization.CultureInfo]::InvariantCulture
$jobs = Get-Content -Raw -Encoding UTF8 $JobsFile | ConvertFrom-Json
# Windows PowerShell 5.1 hands a JSON array over as ONE object[]: a single job comes back as a bare object
if ($jobs -isnot [System.Array]) { $jobs = @($jobs) }
$CAP = 150000

function Cell-Json($x) {
  if ($null -eq $x) { return 'null' }
  if ($x -is [bool]) { if ($x) { return 'true' } else { return 'false' } }
  if ($x -is [int]) { return ('{"e":' + $x + '}') }
  if ($x -is [double]) {
    if ([double]::IsNaN($x) -or [double]::IsInfinity($x)) { return 'null' }
    return $x.ToString('R', $inv)
  }
  return (JStr ([string]$x))
}

# $v = the cell's Value2: a TEXT cell whose text starts with '=' (typed into a Text-formatted cell) reports that text as its Formula too: its value is the very same
# string, a real formula never has its own text as its result
function Formula-Json($x, $v) {
  if ($null -eq $x) { return '""' }
  $s = [string]$x
  if ($s.Length -gt 0 -and $s[0] -eq '=') {
    if ($v -is [string] -and $v -ceq $s) { return '""' }
    return (JStr $s)
  }
  return '""'
}

function Num($x) {
  try { if ($null -eq $x) { return 'null' }; $d = [double]$x; if ([double]::IsNaN($d)) { return 'null' }; return $d.ToString('R', $inv) } catch { return 'null' }
}
function Bool($x) { try { if ([bool]$x) { return 'true' } else { return 'false' } } catch { return 'null' } }
function Str($x) { try { return (JStr ([string]$x)) } catch { return '""' } }
function Cnt($o) { try { return [int]$o.Count } catch { return 0 } }
function Flat($v) {
  if ($null -eq $v) { return '' }
  if ($v -is [array]) { $o = @(); foreach ($e in $v) { $o += [string]$e }; return ($o -join ';') }
  return [string]$v
}

# ---- one cell's look: what Excel renders (number format, font, fill, alignment, wrap, merge, borders)
function Cell-Look($ws, [int]$r, [int]$c) {
  $x = $ws.Cells.Item($r, $c)
  $f = $x.Font
  $i = $x.Interior
  $nf = ''; try { $nf = [string]$x.NumberFormat } catch {}
  $bb = 0; $bl = 0; $bt = 0; $br = 0
  try { $bb = [int]$x.Borders.Item(9).LineStyle } catch {}
  try { $bl = [int]$x.Borders.Item(7).LineStyle } catch {}
  try { $bt = [int]$x.Borders.Item(8).LineStyle } catch {}
  try { $br = [int]$x.Borders.Item(10).LineStyle } catch {}
  return ('{"r":' + $r + ',"c":' + $c + ',"nf":' + (Str $nf) + ',"fn":' + (Str $f.Name) + ',"fs":' + (Num $f.Size) + ',"fb":' + (Bool $f.Bold) + ',"fi":' + (Bool $f.Italic) + ',"fc":' + (Num $f.Color) + ',"ic":' + (Num $i.Color) + ',"ip":' + (Num $i.Pattern) + ',"ha":' + (Num $x.HorizontalAlignment) + ',"va":' + (Num $x.VerticalAlignment) + ',"wr":' + (Bool $x.WrapText) + ',"mg":' + (Bool $x.MergeCells) + ',"bb":' + $bb + ',"bl":' + $bl + ',"bt":' + $bt + ',"br":' + $br + '}')
}

# the coordinates sampled from a sheet (the same for the source and its copy): the top-left 6 x 8 block plus a stride over the rest
function Sample-Cells([int]$maxR, [int]$maxC) {
  $pts = New-Object System.Collections.Generic.List[object]
  $seen = @{}
  $add = { param($r, $c) $k = "$r,$c"; if (-not $seen.ContainsKey($k) -and $r -ge 1 -and $c -ge 1 -and $r -le $maxR -and $c -le $maxC) { $seen[$k] = 1; $pts.Add(@($r, $c)) } }
  for ($r = 1; $r -le [math]::Min(4, $maxR); $r++) { for ($c = 1; $c -le [math]::Min(8, $maxC); $c++) { & $add $r $c } }
  $rs = [math]::Max(1, [int][math]::Ceiling($maxR / 7)); $cs = [math]::Max(1, [int][math]::Ceiling($maxC / 5))
  for ($r = 1; $r -le $maxR -and $pts.Count -lt 42; $r += $rs) { for ($c = 1; $c -le $maxC -and $pts.Count -lt 42; $c += $cs) { & $add $r $c } }
  & $add $maxR $maxC
  return $pts
}

function Deep-Sheet($wb, $ws, [int]$type, [int]$maxR, [int]$maxC) {
  $p = New-Object System.Collections.Generic.List[string]
  $tab = -1; try { if ($ws.Tab.ColorIndex -ne -4142) { $tab = [int]$ws.Tab.Color } } catch {}
  $p.Add('"tab":' + $tab)
  if ($type -eq -4109) {
    # chart sheet: the chart itself
    $series = New-Object System.Collections.Generic.List[string]
    $n = 0; try { $n = [int]$ws.SeriesCollection().Count } catch {}
    for ($j = 1; $j -le $n; $j++) {
      $s = $ws.SeriesCollection($j)
      $v = ''; try { $v = Flat $s.Values } catch {}
      $series.Add('{"v":' + (Str $v) + '}')
    }
    $ct = ''; try { $ct = [string]$ws.ChartType } catch {}
    $p.Add('"chartSheet":{"type":' + (Str $ct) + ',"series":[' + ($series -join ',') + ']}')
    return '{' + ($p -join ',') + '}'
  }
  # column widths (points = what the column HAS: stored widths are characters of each workbook's default font) and hidden rows / columns
  $w = New-Object System.Collections.Generic.List[string]
  $hc = New-Object System.Collections.Generic.List[string]
  $nc = [math]::Min([math]::Max($maxC, 1), 120)
  for ($c = 1; $c -le [math]::Min($nc, 40); $c++) { $w.Add((Num ([math]::Round([double]$ws.Columns.Item($c).Width, 1)))) }
  # Range.Hidden is True / False when uniform and Null when mixed: only a mixed range needs the loop
  $hh = $null; try { $hh = $ws.Range($ws.Columns.Item(1), $ws.Columns.Item($nc)).Hidden } catch {}
  if ($hh -is [bool]) { if ($hh) { for ($c = 1; $c -le $nc; $c++) { $hc.Add([string]$c) } } }
  else { for ($c = 1; $c -le $nc; $c++) { try { if ($ws.Columns.Item($c).Hidden) { $hc.Add([string]$c) } } catch {} } }
  $p.Add('"widths":[' + ($w -join ',') + ']')
  $p.Add('"hiddenCols":[' + ($hc -join ',') + ']')
  $hr = New-Object System.Collections.Generic.List[string]
  $nr = [math]::Min([math]::Max($maxR, 1), 400)
  $hh = $null; try { $hh = $ws.Range($ws.Rows.Item(1), $ws.Rows.Item($nr)).Hidden } catch {}
  if ($hh -is [bool]) { if ($hh) { for ($r = 1; $r -le $nr; $r++) { $hr.Add([string]$r) } } }
  else { for ($r = 1; $r -le $nr; $r++) { try { if ($ws.Rows.Item($r).Hidden) { $hr.Add([string]$r) } } catch {} } }
  $p.Add('"hiddenRows":[' + ($hr -join ',') + ']')
  # sampled cells
  $cells = New-Object System.Collections.Generic.List[string]
  foreach ($pt in (Sample-Cells $maxR $maxC)) { try { $cells.Add((Cell-Look $ws $pt[0] $pt[1])) } catch {} }
  $p.Add('"look":[' + ($cells -join ',') + ']')
  # conditional formats / validation / hyperlinks / sheet-scoped names
  $cf = -1; try { $cf = [int]$ws.Cells.FormatConditions.Count } catch {}
  $p.Add('"cf":' + $cf)
  $dv = ''; try { $dv = [string]$ws.Cells.SpecialCells(-4174).Address($false, $false) } catch {}
  $p.Add('"dv":' + (Str $dv))
  $p.Add('"links":' + (Cnt $ws.Hyperlinks))
  # shapes: pictures, text boxes, shapes, groups, connectors, charts (notes, form controls, ActiveX, slicers are counted apart)
  $sh = New-Object System.Collections.Generic.List[string]
  $ns = 0; try { $ns = [int]$ws.Shapes.Count } catch {}
  $other = 0
  for ($i = 1; $i -le $ns; $i++) {
    $s = $ws.Shapes.Item($i)
    $t = [int]$s.Type
    if ($t -eq 4 -or $t -eq 8 -or $t -eq 12 -or $t -eq 24 -or $t -eq 25) { $other++; continue }
    $txt = ''; try { $txt = [string]$s.TextFrame2.TextRange.Text } catch {}
    $sh.Add('{"t":' + $t + ',"n":' + (Str $s.Name) + ',"x":' + (Str $txt) + '}')
  }
  $p.Add('"shapes":[' + ($sh -join ',') + ']')
  $p.Add('"shapesOther":' + $other)
  # charts (series count + the values of every series)
  $ch = New-Object System.Collections.Generic.List[string]
  $nch = 0; try { $nch = [int]$ws.ChartObjects().Count } catch {}
  for ($i = 1; $i -le $nch; $i++) {
    $co = $ws.ChartObjects($i)
    $c = $co.Chart
    $series = New-Object System.Collections.Generic.List[string]
    $n = 0; try { $n = [int]$c.SeriesCollection().Count } catch {}
    for ($j = 1; $j -le $n; $j++) {
      $s = $c.SeriesCollection($j)
      $v = ''; try { $v = Flat $s.Values } catch {}
      $x = ''; try { $x = Flat $s.XValues } catch {}
      $series.Add('{"v":' + (Str $v) + ',"x":' + (Str $x) + '}')
    }
    $ct = ''; try { $ct = [string]$c.ChartType } catch {}
    $ch.Add('{"type":' + (Str $ct) + ',"series":[' + ($series -join ',') + ']}')
  }
  $p.Add('"charts":[' + ($ch -join ',') + ']')
  # tables
  $tb = New-Object System.Collections.Generic.List[string]
  $nt = Cnt $ws.ListObjects
  for ($i = 1; $i -le $nt; $i++) {
    $lo = $ws.ListObjects.Item($i)
    $cols = @(); $ncol = Cnt $lo.ListColumns
    for ($c = 1; $c -le $ncol; $c++) { $cols += [string]$lo.ListColumns.Item($c).Name }
    $tb.Add('{"n":' + (Str $lo.Name) + ',"a":' + (Str $lo.Range.Address($false, $false)) + ',"cols":' + (Str ($cols -join '|')) + ',"totals":' + (Bool $lo.ShowTotals) + '}')
  }
  $p.Add('"tables":[' + ($tb -join ',') + ']')
  # notes
  $nn = New-Object System.Collections.Generic.List[string]
  $ncm = Cnt $ws.Comments
  for ($i = 1; $i -le $ncm; $i++) {
    $cm = $ws.Comments.Item($i)
    $t = ''; try { $t = [string]$cm.Text() } catch {}
    $nn.Add('{"a":' + (Str $cm.Parent.Address($false, $false)) + ',"t":' + (Str $t) + '}')
  }
  $p.Add('"notes":[' + ($nn -join ',') + ']')
  # pivot tables
  $pv = New-Object System.Collections.Generic.List[string]
  $npv = 0; try { $npv = Cnt ($ws.PivotTables()) } catch {}
  for ($i = 1; $i -le $npv; $i++) {
    $pt = $ws.PivotTables($i)
    $src = ''; try { $src = [string]$pt.SourceData } catch {}
    $vals = ''; try { $vals = Flat $pt.TableRange2.Value2 } catch {}
    $pv.Add('{"n":' + (Str $pt.Name) + ',"a":' + (Str $pt.TableRange1.Address($false, $false)) + ',"src":' + (Str $src) + ',"v":' + (Str $vals) + '}')
  }
  $p.Add('"pivots":[' + ($pv -join ',') + ']')
  return '{' + ($p -join ',') + '}'
}

function Dump-Sheet($wb, $ws, [bool]$deep) {
  # (Worksheet.Type is the XlSheetType, but Chart.Type of a chart sheet is its CHART type: ask the collections instead)
  $type = 3; if ($script:kinds.ContainsKey([string]$ws.Name)) { $type = [int]$script:kinds[[string]$ws.Name] }
  $vis = -1; try { $vis = [int]$ws.Visible } catch {}
  $head = '{"name":' + (JStr ([string]$ws.Name)) + ',"type":' + $type + ',"visible":' + $vis
  if ($type -ne -4167) {
    $d = ''
    if ($deep) { try { $d = ',"deep":' + (Deep-Sheet $wb $ws $type 0 0) } catch { $d = '' } }
    return ($head + ',"rows":0,"cols":0,"circ":"","v":[],"f":[]' + $d + '}')
  }
  $ur = $ws.UsedRange
  $maxR = $ur.Row + $ur.Rows.Count - 1
  $maxC = $ur.Column + $ur.Columns.Count - 1
  $trunc = $false
  $dr = $maxR; $dc = $maxC
  if ($dc -gt 150) { $dc = 150; $trunc = $true }
  if ($dr * $dc -gt $CAP) { $dr = [int][math]::Floor($CAP / $dc); $trunc = $true }
  $rng = $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item($dr, $dc))
  $vals = $rng.Value2
  $fms = $rng.Formula
  $vrows = New-Object System.Collections.Generic.List[string]
  $frows = New-Object System.Collections.Generic.List[string]
  if ($vals -isnot [System.Array]) {
    $vrows.Add('[' + (Cell-Json $vals) + ']')
    $frows.Add('[' + (Formula-Json $fms $vals) + ']')
  } else {
    for ($r = 1; $r -le $dr; $r++) {
      $vc = New-Object System.Collections.Generic.List[string]
      $fc = New-Object System.Collections.Generic.List[string]
      for ($c = 1; $c -le $dc; $c++) {
        $vc.Add((Cell-Json $vals[$r, $c]))
        $fc.Add((Formula-Json $fms[$r, $c] $vals[$r, $c]))
      }
      $vrows.Add('[' + ($vc -join ',') + ']')
      $frows.Add('[' + ($fc -join ',') + ']')
    }
  }
  # a circular reference has no defined value in Excel (the result depends on the order of evaluation): the judge must know
  $circ = ''
  try { $cr = $ws.CircularReference; if ($null -ne $cr) { $circ = [string]$cr.Address($false, $false) } } catch { $circ = '' }
  $d = ''
  if ($deep) { try { $d = ',"deep":' + (Deep-Sheet $wb $ws $type $maxR $maxC) } catch { $d = ',"deepError":' + (JStr ([string]$_.Exception.Message)) } }
  return ($head + ',"rows":' + $dr + ',"cols":' + $dc + ',"used":[' + $maxR + ',' + $maxC + '],"trunc":' + ([string]$trunc).ToLower() + ',"circ":' + (JStr $circ) + ',"v":[' + ($vrows -join ',') + '],"f":[' + ($frows -join ',') + ']' + $d + '}')
}

function Dump-Names($wb) {
  $o = New-Object System.Collections.Generic.List[string]
  $n = Cnt $wb.Names
  for ($i = 1; $i -le [math]::Min($n, 600); $i++) {
    try {
      $x = $wb.Names.Item($i)
      $o.Add('{"name":' + (JStr ([string]$x.Name)) + ',"refers":' + (JStr ([string]$x.RefersTo)) + ',"visible":' + (Bool $x.Visible) + '}')
    } catch {}
  }
  return ($o -join ',')
}

function Alive { try { return ($null -ne $script:xl -and $null -ne $script:xl.Workbooks -and $null -ne $script:xl.Version) } catch { return $false } }
function Restart-Excel {
  [Console]::Error.WriteLine('excel instance lost, starting another')
  Close-GExcel $script:xl
  $script:xl = New-GExcel
  Write-Host ('{"type":"excel","pid":' + [int]$script:GExcelPid + ',"restart":true}')
}

$xl = $null
try {
  $xl = New-GExcel
  Write-Host ('{"type":"excel","pid":' + $script:GExcelPid + '}')
  foreach ($j in $jobs) {
    # a transient COM hiccup (the first jobs of a fresh instance) must not look like a broken package: a failed job is tried again;
    # an Excel that died (it does, now and then, for no reason of the file) is replaced by another instance of our own
    for ($attempt = 1; $attempt -le 3; $attempt++) {
      if (-not (Alive)) { Restart-Excel }
      $xl = $script:xl
      $mode = 'failed'; $wb = $null; $log = ''; $body = ''; $names = ''; $d1904 = 'false'
      $t0 = Get-Date
      try { $wb = $xl.Workbooks.Open($j.file, 0, $true); $mode = 'normal' }
      catch {
        try { $wb = $xl.Workbooks.Open($j.file, 0, $true, 5, '', '', $true, 2, '', $false, $false, $M, $false, $false, 1); $mode = 'repair' } catch { $mode = 'failed'; $log = [string]$_.Exception.Message + ' @line ' + [string]$_.InvocationInfo.ScriptLineNumber }
      }
      if ($mode -eq 'repair') {
        Start-Sleep -Milliseconds 200
        foreach ($l in @(Get-ChildItem $env:TEMP -Filter 'error*.xml' -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime -ge $t0.AddSeconds(-1) })) {
          $x = [IO.File]::ReadAllText($l.FullName)
          $log += ([regex]::Matches($x, '<(removed|repaired)[A-Za-z]*>[^<]*') | ForEach-Object { $_.Value }) -join ' ; '
        }
      }
      if ($wb) {
        try {
          $xl.CalculateFull()
          try { $d1904 = ([string][bool]$wb.Date1904).ToLower() } catch {}
          $script:kinds = @{}
          foreach ($x in @($wb.Worksheets)) { $script:kinds[[string]$x.Name] = -4167 }
          foreach ($x in @($wb.Charts)) { $script:kinds[[string]$x.Name] = -4109 }
          try { foreach ($x in @($wb.DialogSheets)) { $script:kinds[[string]$x.Name] = -4116 } } catch {}
          $deepSet = $null; $deepAll = $false
          if ($j.deep -is [string]) { if ($j.deep -eq 'all') { $deepAll = $true } }
          elseif ($null -ne $j.deep) { $deepSet = @{}; foreach ($n in @($j.deep)) { $deepSet[[string]$n] = $true } }
          $parts = New-Object System.Collections.Generic.List[string]
          foreach ($ws in @($wb.Sheets)) {
            $deep = $deepAll -or ($null -ne $deepSet -and $deepSet.ContainsKey([string]$ws.Name))
            $parts.Add((Dump-Sheet $wb $ws $deep))
          }
          $body = $parts -join ','
          $names = Dump-Names $wb
        } catch { $mode = 'failed'; $log = [string]$_.Exception.Message + ' @line ' + [string]$_.InvocationInfo.ScriptLineNumber }
        try { $wb.Close($false) } catch {}
      }
      if ($mode -ne 'failed') { break }
      [Console]::Error.WriteLine('job ' + $j.id + ' attempt ' + $attempt + ' failed: ' + $log)
      Start-Sleep -Milliseconds 1500
    }
    $json = '{"open":"' + $mode + '","log":' + (JStr $log) + ',"date1904":' + $d1904 + ',"names":[' + $names + '],"sheets":[' + $body + ']}'
    [IO.File]::WriteAllText((Join-Path $OutDir ($j.id + '.book.json')), $json, (New-Object System.Text.UTF8Encoding($false)))
    Write-Host ('{"type":"job","id":' + (JStr $j.id) + ',"open":"' + $mode + '"}')
  }
} finally {
  Close-GExcel $xl
}
