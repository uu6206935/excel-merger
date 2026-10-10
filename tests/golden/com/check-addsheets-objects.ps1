# Opens the outputs of tests/golden/tools/make-addsheets-objects-outputs.ts (premerged workbooks whose COPIED sheets own objects: charts, pictures, shapes,
# notes, tables, pivot tables, chart sheets) in ONE hidden Excel instance of our own and compares every copied sheet with the source sheet (task A2):
#   powershell -NoProfile -ExecutionPolicy Bypass -File tests\golden\com\check-addsheets-objects.ps1 -Dir <folder with out\ manifest.json> [-Only regex]
# Per manifest entry (one ASCII JSON line each, then a summary):
#   open        normal | repair | failed   (a plain Open that throws means Excel wants to repair the file; then it is opened with CorruptLoad and the
#                                          repair log is collected, like tests/golden/com/opentest.ps1)
#   names       the sheet names, in order, are the planned ones
#   shapeDiffs  pictures / text boxes / shapes / groups / connectors: type, name, position and size (1 pt), text of text boxes, link of linked text boxes,
#               the target of shape hyperlinks (modulo the renamed sheets); the placeholders of the things that are not copied (slicers, controls) are not shapes
#   chartDiffs  per chart: series count, series formula (modulo sheet prefixes), values, category values, name, chart type, axis group, series colour,
#               title text, legend; chart sheets likewise
#   noteDiffs   notes: count (a threaded comment is a note in the copy), text, author, visibility
#   tableDiffs  tables: count, names (a table of the second copy of a file is named <name>_2 ...), range, totals row, style, column names, first data
#               cell formula of every column (modulo renamed sheets / tables)
#   pivotDiffs  pivot tables: count, name, range, source data (modulo the sheet name), cell values; then PivotCache.Refresh on the copy: no error,
#               the values are the source's again
#   newErr      #REF! / #NAME? cells of the copied sheets that the source sheets did not have (after CalculateFull of both)
# Own instance only: its PID comes from Application.Hwnd and only that PID is ever killed. No clipboard, no PageSetup.
param([string]$Dir, [string]$Only = '')
$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\common.ps1"
$M = $script:GMissing
$Dir = (Resolve-Path $Dir).Path
$man = Get-Content -Raw -Encoding UTF8 (Join-Path $Dir 'manifest.json') | ConvertFrom-Json
$REF = -2146826265; $NAME = -2146826259
$PREFIX = "('([^']|'')*'|[^\s'!=+\-*/^&<>(),:;{}\[\]]+)!"
function Strip-Prefix([string]$f) { return [regex]::Replace($f, $PREFIX, '') }
function Norm-Name([string]$n) { return ($n -replace '_\d+$', '') }
# a formula / text modulo the renamed sheets and the "_2" table names
function Norm-Formula([string]$f) {
  $x = Strip-Prefix $f
  return ([regex]::Replace($x, '_\d+(?=\[|\b)', ''))
}

function Open-Book($xl, [string]$path) {
  $t0 = Get-Date
  $mode = 'failed'; $wb = $null; $log = ''
  try { $wb = $xl.Workbooks.Open($path, 0, $true); $mode = 'normal' }
  catch {
    try { $wb = $xl.Workbooks.Open($path, 0, $true, 5, '', '', $true, 2, '', $false, $false, $M, $false, $false, 1); $mode = 'repair' } catch { $mode = 'failed'; $log = $_.Exception.Message }
  }
  Start-Sleep -Milliseconds 150
  foreach ($l in @(Get-ChildItem $env:TEMP -Filter 'error*.xml' -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime -ge $t0.AddSeconds(-1) })) {
    $x = [IO.File]::ReadAllText($l.FullName)
    $log += ([regex]::Matches($x, '<(removed|repaired)[A-Za-z]*>[^<]*') | ForEach-Object { $_.Value }) -join ' ; '
  }
  return @{ wb = $wb; mode = $mode; log = $log }
}

function Cnt($o) { try { return [int]$o.Count } catch { return 0 } }
function J($a) { return (@($a | ForEach-Object { [string]$_ }) -join ';') }
function Flat($v) {
  if ($null -eq $v) { return '' }
  if ($v -is [array]) { $o = @(); foreach ($e in $v) { $o += [string]$e }; return ($o -join ';') }
  return [string]$v
}

# ---- shapes (pictures, text boxes, shapes, groups, connectors, charts as shapes) ----
function Shape-List($ws) {
  $out = @()
  $n = 0; try { $n = [int]$ws.Shapes.Count } catch {}
  for ($i = 1; $i -le $n; $i++) {
    $s = $ws.Shapes.Item($i)
    $t = [int]$s.Type
    if ($t -eq 4 -or $t -eq 8 -or $t -eq 12 -or $t -eq 24 -or $t -eq 25) { continue }   # note, form control, OLE / ActiveX control, slicer (msoSlicer)
    # (HasTextFrame is not reachable through this COM binding: ask the text range itself; pictures and lines have none)
    $text = ''; try { $text = [string]$s.TextFrame2.TextRange.Text } catch {}
    $link = ''; try { $link = [string]$s.Hyperlink.SubAddress } catch {}
    $fml = ''; try { if ($t -eq 17) { $fml = [string]$s.DrawingObject.Formula } } catch {}
    $kids = 0; try { if ($t -eq 6) { $kids = [int]$s.GroupItems.Count } } catch {}
    $out += [pscustomobject]@{ type = $t; name = [string]$s.Name; l = [math]::Round([double]$s.Left, 1); t = [math]::Round([double]$s.Top, 1); w = [math]::Round([double]$s.Width, 1); h = [math]::Round([double]$s.Height, 1); text = $text; link = $link; fml = $fml; kids = $kids }
  }
  return $out
}
function Compare-Shapes($a, $b, $first, $geom) {
  $d = 0
  if (@($a).Count -ne @($b).Count) { $first.Value += ('shapes ' + @($a).Count + ' != ' + @($b).Count); return 1 }
  for ($i = 0; $i -lt @($a).Count; $i++) {
    $x = @($a)[$i]; $y = @($b)[$i]
    $bad = ($x.type -ne $y.type) -or ($x.name -ne $y.name) -or ($x.text -ne $y.text) -or ($x.kids -ne $y.kids)
    if (-not $bad -and $x.link) { $bad = ((Norm-Formula $x.link) -ne (Norm-Formula $y.link)) }
    if (-not $bad -and $x.fml) { $bad = ((Norm-Formula $x.fml) -ne (Norm-Formula $y.fml)) }
    if ($bad) { $d++; $first.Value += ('shape ' + $x.name + ' ' + $x.type + '/' + $y.type + ' ' + $x.text + ' | ' + $y.text) }
    # position and size: they follow the rows / columns the object is anchored to (a source whose default font differs from the base's has other default row heights)
    if (([math]::Abs($x.l - $y.l) -gt 1.1) -or ([math]::Abs($x.t - $y.t) -gt 1.1) -or ([math]::Abs($x.w - $y.w) -gt 1.1) -or ([math]::Abs($x.h - $y.h) -gt 1.1)) {
      $geom.Value += 1
      $first.Value += ('geometry ' + $x.name + ' ' + $x.l + ',' + $x.t + ',' + $x.w + ',' + $x.h + ' | ' + $y.l + ',' + $y.t + ',' + $y.w + ',' + $y.h)
    }
  }
  return $d
}

# ---- charts ----
function Chart-Sig($ch) {
  $o = [ordered]@{ type = ''; pie = $false; title = ''; legend = ''; series = @() }
  try { $o['type'] = [string]$ch.ChartType } catch {}
  # pie / doughnut charts: one colour per point, the series colour says nothing
  $o['pie'] = (@('5', '68', '69', '70', '71', '-4102', '-4120', '80') -contains $o['type'])
  try { if ($ch.HasTitle) { $o['title'] = [string]$ch.ChartTitle.Text } } catch {}
  try { $o['legend'] = [string]$ch.HasLegend } catch {}
  $n = 0; try { $n = [int]$ch.SeriesCollection().Count } catch {}
  for ($j = 1; $j -le $n; $j++) {
    $s = $ch.SeriesCollection($j)
    $f = ''; try { $f = [string]$s.Formula } catch {}
    $v = ''; try { $v = Flat $s.Values } catch {}
    $x = ''; try { $x = Flat $s.XValues } catch {}
    $nm = ''; try { $nm = [string]$s.Name } catch {}
    $ty = ''; try { $ty = [string]$s.ChartType } catch {}
    $ax = ''; try { $ax = [string]$s.AxisGroup } catch {}
    $col = ''; try { $col = [string]$s.Format.Fill.ForeColor.RGB } catch {}
    $line = ''; try { $line = [string]$s.Format.Line.ForeColor.RGB } catch {}
    # the colour of the first points (a pie chart has one colour per point: the series colour alone says nothing)
    $pts = @(); $np = 0; try { $np = [int]$s.Points().Count } catch {}
    for ($k = 1; $k -le [math]::Min(6, $np); $k++) { try { $pts += [string]$s.Points($k).Format.Fill.ForeColor.RGB } catch { $pts += '?' } }
    $o['series'] += [pscustomobject]@{ f = $f; v = $v; x = $x; name = $nm; type = $ty; axis = $ax; color = $col; line = $line; pts = ($pts -join ',') }
  }
  return $o
}
function Compare-Chart($a, $b, $first, [string]$what, [bool]$skipColor, [bool]$skipFormula) {
  $d = 0
  if ($a['type'] -ne $b['type'] -or $a['title'] -ne $b['title'] -or $a['legend'] -ne $b['legend']) { $d++; $first.Value += ($what + ' chart ' + $a['type'] + '/' + $a['title'] + ' | ' + $b['type'] + '/' + $b['title']) }
  if (@($a['series']).Count -ne @($b['series']).Count) { $first.Value += ($what + ' series ' + @($a['series']).Count + ' != ' + @($b['series']).Count); return $d + 1 }
  for ($j = 0; $j -lt @($a['series']).Count; $j++) {
    $x = @($a['series'])[$j]; $y = @($b['series'])[$j]
    $bad = ((-not $skipFormula) -and ((Norm-Formula $x.f) -ne (Norm-Formula $y.f))) -or ($x.v -ne $y.v) -or ($x.x -ne $y.x) -or ($x.name -ne $y.name) -or ($x.type -ne $y.type) -or ($x.axis -ne $y.axis) -or ((-not $skipColor) -and ($(if ($a['pie']) { $x.pts -ne $y.pts } else { ($x.color -ne $y.color) -or ($x.line -ne $y.line) })))
    if ($bad) { $d++; $first.Value += ($what + ' series ' + ($j + 1) + ': ' + $x.f + ' | ' + $y.f + ' ; v ' + $x.v + ' | ' + $y.v + ' ; c ' + $x.color + ' | ' + $y.color) }
  }
  return $d
}

# ---- notes ----
function Note-List($ws) {
  $out = @()
  $n = Cnt $ws.Comments
  for ($i = 1; $i -le $n; $i++) {
    $c = $ws.Comments.Item($i)
    $t = ''; try { $t = [string]$c.Text() } catch {}
    $vis = ''; try { $vis = [string]$c.Visible } catch {}
    $out += [pscustomobject]@{ addr = [string]$c.Parent.Address($false, $false); text = $t; vis = $vis }
  }
  return $out
}

# ---- tables ----
function Table-List($ws) {
  $out = @()
  $n = Cnt $ws.ListObjects
  for ($i = 1; $i -le $n; $i++) {
    $lo = $ws.ListObjects.Item($i)
    $cols = @(); $fml = @()
    $nc = Cnt $lo.ListColumns
    for ($c = 1; $c -le $nc; $c++) {
      $cols += [string]$lo.ListColumns.Item($c).Name
      $f = ''; try { $f = [string]$lo.ListColumns.Item($c).DataBodyRange.Cells.Item(1, 1).Formula } catch {}
      $fml += $f
    }
    $style = ''; try { $style = [string]$lo.TableStyle.Name } catch {}
    $out += [pscustomobject]@{ name = [string]$lo.Name; addr = [string]$lo.Range.Address($false, $false); totals = [bool]$lo.ShowTotals; style = $style; cols = ($cols -join '|'); fml = ($fml -join '|') }
  }
  return $out
}

# ---- pivot tables ----
function Pivot-List($ws) {
  $out = @()
  $n = Cnt ($ws.PivotTables())
  for ($i = 1; $i -le $n; $i++) {
    $pt = $ws.PivotTables($i)
    $src = ''; try { $src = [string]$pt.SourceData } catch {}
    $vals = ''; try { $vals = Flat $pt.TableRange2.Value2 } catch {}
    $out += [pscustomobject]@{ name = [string]$pt.Name; addr = [string]$pt.TableRange1.Address($false, $false); src = $src; vals = $vals; pt = $pt }
  }
  return $out
}

function Count-Errors($ws) {
  $e = 0
  try {
    $u = $ws.UsedRange
    $v = $u.Value2
    if ($v -is [array]) { foreach ($c in $v) { if ($c -is [int] -and ($c -eq $REF -or $c -eq $NAME)) { $e++ } } }
    elseif ($v -is [int] -and ($v -eq $REF -or $v -eq $NAME)) { $e++ }
  } catch {}
  return $e
}

function Alive { try { return ($null -ne $script:xl -and $null -ne $script:xl.Workbooks -and $null -ne $script:xl.Version) } catch { return $false } }
function Restart-Excel {
  [Console]::Error.WriteLine('excel instance lost, starting another')
  $script:srcCache = @{}
  Close-GExcel $script:xl
  $script:xl = New-GExcel
  Write-Host ('{"type":"excel","pid":' + [int]$script:GExcelPid + ',"restart":true}')
}

function Run-Entry($e) {
  $xl = $script:xl; $srcCache = $script:srcCache
  $t0 = Get-Date
  $outPath = Join-Path $Dir $e.out
  $line = [ordered]@{ name = $e.name; open = 'failed'; log = ''; sheets = -1; names = $false; lint = [string]$e.lint; shapeDiffs = 0; geomDiffs = 0; chartDiffs = 0; noteDiffs = 0; tableDiffs = 0; pivotDiffs = 0; newErr = 0; compared = 0; pivotsRefreshed = 0; charts = 0; shapes = 0; notes = 0; tables = 0; pivots = 0 }
  $o = Open-Book $xl $outPath
  [Console]::Error.WriteLine('opened ' + $e.name + ' ' + $o.mode)
  $line['open'] = $o.mode; $line['log'] = $o.log
  $info = @()
  if ($o.wb) {
    $wb = $o.wb
    try { $xl.CalculateFull() } catch {}
    $got = @(); foreach ($s in @($wb.Sheets)) { $got += $s.Name }
    $line['sheets'] = $got.Count
    $line['names'] = (($got -join '|') -ceq (@($e.expected) -join '|'))
    if (-not $line['names']) { $info += ('names ' + ($got -join '|') + ' != ' + (@($e.expected) -join '|')) }
    foreach ($src in @($e.sources)) {
      $sp = [string]$src.file
      if (-not $srcCache.ContainsKey($sp)) {
        $sb = Open-Book $xl $sp
        if ($sb.wb) { try { $xl.CalculateFull() } catch {} }
        $srcCache[$sp] = $sb
      }
      $sb = $srcCache[$sp]
      if (-not $sb.wb) { $info += ('source ' + $src.label + ' did not open: ' + $sb.mode); continue }
      foreach ($sh in @($src.sheets)) {
        $wsO = $null; try { $wsO = $wb.Sheets.Item([string]$sh.name) } catch {}
        $wsS = $null
        foreach ($cand in @($sb.wb.Sheets)) { if ($cand.Name -ceq $sh.original) { $wsS = $cand; break } }
        if (-not $wsO -or -not $wsS) { $info += ('missing pair ' + $sh.original + ' -> ' + $sh.name); $line['names'] = $false; continue }
        $line['compared'] = $line['compared'] + 1
        $first = @()
        $isChartSheet = ([string]$sh.objects.kind -eq 'chartsheet')
        try {
          if ($isChartSheet -and $src.drawingsSkipped) {
            if ($wsO.GetType().Name -eq 'Chart') { $line['chartDiffs']++; $first += 'chart sheet that must become an empty sheet' }
          } elseif ($isChartSheet) {
            $line['charts'] += 1
            $line['chartDiffs'] += (Compare-Chart (Chart-Sig $wsS) (Chart-Sig $wsO) ([ref]$first) 'chartsheet' ([bool]($src.themeDiffers -and @($sh.objects.chartAuto)[0])) ([bool]$e.literalCharts))
          } else {
            if ($src.drawingsSkipped) {
              # a Strict package and a transitional one differ in the values of DrawingML: the drawing is left out (and named in the warning)
              if ((Cnt ($wsO.ChartObjects())) -ne 0 -or @(Shape-List $wsO).Count -ne 0) { $line['shapeDiffs']++; $first += 'shapes / charts of a drawing that must be left out' }
            } else {
            # shapes
            $a = Shape-List $wsS; $b = Shape-List $wsO
            $line['shapes'] += @($b).Count
            $g = 0
            $line['shapeDiffs'] += (Compare-Shapes $a $b ([ref]$first) ([ref]$g))
            $line['geomDiffs'] += $g
            # charts
            $nco = Cnt ($wsS.ChartObjects())
            if ($nco -ne (Cnt ($wsO.ChartObjects()))) { $line['chartDiffs']++; $first += ('chart objects ' + $nco + ' != ' + (Cnt ($wsO.ChartObjects()))) }
            else {
              $line['charts'] += $nco
              for ($i = 1; $i -le $nco; $i++) { $line['chartDiffs'] += (Compare-Chart (Chart-Sig ($wsS.ChartObjects($i).Chart)) (Chart-Sig ($wsO.ChartObjects($i).Chart)) ([ref]$first) ('chart ' + $i) ([bool]($src.themeDiffers -and @($sh.objects.chartAuto)[$i - 1])) ([bool]$e.literalCharts)) }
            }
            }
            # notes: a threaded comment of the source is a note of the copy
            $na = @(Note-List $wsS); $nb = @(Note-List $wsO)
            $nt = Cnt $wsS.CommentsThreaded
            $line['notes'] += $nb.Count
            if ($nb.Count -ne $na.Count + $nt) { $line['noteDiffs']++; $first += ('notes ' + $nb.Count + ' != ' + $na.Count + '+' + $nt) }
            else {
              foreach ($x in $na) {
                $y = $nb | Where-Object { $_.addr -eq $x.addr } | Select-Object -First 1
                if (-not $y -or $y.text -ne $x.text -or $y.vis -ne $x.vis) { $line['noteDiffs']++; $first += ('note ' + $x.addr + ': ' + $x.text + ' | ' + $(if ($y) { $y.text } else { '-' })) }
              }
              if ((Cnt $wsO.CommentsThreaded) -ne 0) { $line['noteDiffs']++; $first += 'threaded comments in the copy' }
            }
            # tables
            $ta = @(Table-List $wsS); $tb = @(Table-List $wsO)
            $line['tables'] += $tb.Count
            if ($ta.Count -ne $tb.Count) { $line['tableDiffs']++; $first += ('tables ' + $ta.Count + ' != ' + $tb.Count) }
            else {
              for ($i = 0; $i -lt $ta.Count; $i++) {
                $x = $ta[$i]; $y = $tb | Where-Object { (Norm-Name $_.name) -eq (Norm-Name $x.name) } | Select-Object -First 1
                if (-not $y) { $line['tableDiffs']++; $first += ('table ' + $x.name + ' missing'); continue }
                if ((Norm-Name $x.name) -ne (Norm-Name $y.name) -or $x.addr -ne $y.addr -or $x.totals -ne $y.totals -or $x.style -ne $y.style -or $x.cols -ne $y.cols -or (Norm-Formula $x.fml) -ne (Norm-Formula $y.fml)) {
                  $line['tableDiffs']++; $first += ('table ' + $x.name + ' ' + $x.addr + ' ' + $x.style + ' | ' + $y.name + ' ' + $y.addr + ' ' + $y.style + ' ; ' + $x.fml + ' | ' + $y.fml)
                }
              }
            }
            # pivot tables
            $pa = @(Pivot-List $wsS); $pb = @(Pivot-List $wsO)
            $line['pivots'] += $pb.Count
            if ($pa.Count -ne $pb.Count) { $line['pivotDiffs']++; $first += ('pivots ' + $pa.Count + ' != ' + $pb.Count) }
            else {
              for ($i = 0; $i -lt $pa.Count; $i++) {
                $x = $pa[$i]; $y = $pb | Where-Object { (Norm-Name $_.name) -eq (Norm-Name $x.name) } | Select-Object -First 1
                if (-not $y) { $line['pivotDiffs']++; $first += ('pivot ' + $x.name + ' missing'); continue }
                $srcOk = ((Norm-Formula $x.src) -eq (Norm-Formula $y.src))
                if ((Norm-Name $x.name) -ne (Norm-Name $y.name) -or $x.addr -ne $y.addr -or -not $srcOk -or $x.vals -ne $y.vals) {
                  $line['pivotDiffs']++; $first += ('pivot ' + $x.name + ' ' + $x.addr + ' ' + $x.src + ' | ' + $y.name + ' ' + $y.addr + ' ' + $y.src + $(if ($x.vals -ne $y.vals) { ' ; values differ' } else { '' }))
                  continue
                }
                # refresh the copy's cache from the copied data: no error, the same values again
                try {
                  $y.pt.PivotCache().Refresh() | Out-Null
                  $line['pivotsRefreshed'] += 1
                  $after = Flat $y.pt.TableRange2.Value2
                  if ($after -ne $x.vals) { $line['pivotDiffs']++; $first += ('pivot ' + $y.name + ' differs after Refresh') }
                } catch { $line['pivotDiffs']++; $first += ('pivot ' + $y.name + ' Refresh failed: ' + (($_.Exception.Message) -replace '[^\x20-\x7e]', '?')) }
              }
            }
            $line['newErr'] += ((Count-Errors $wsO) - (Count-Errors $wsS))
          }
        } catch { $info += ('compare failed for ' + $sh.name + ': ' + (($_.Exception.Message) -replace '[^\x20-\x7e]', '?')) }
        if ($first.Count) { $info += ($sh.name + ' => ' + (($first | Select-Object -First 5) -join ' ;; ')) }
      }
    }
    $line['info'] = $info
    try { $wb.Close($false) } catch {}
  }
  $line['ms'] = [int]((Get-Date) - $t0).TotalMilliseconds
  return $line
}

$xl = $null
$srcCache = @{}
$bad = 0; $count = 0
try {
  $xl = New-GExcel
  Write-Host ('{"type":"excel","pid":' + [int]$script:GExcelPid + ',"version":' + (JStr ([string]$xl.Version)) + '}')
  foreach ($e in @($man)) {
    if ($Only -and ($e.name -notmatch $Only)) { continue }
    $count++
    $line = $null
    for ($attempt = 1; $attempt -le 3; $attempt++) {
      if (-not (Alive)) { Restart-Excel }
      $line = Run-Entry $e
      if (Alive) { break }
      [Console]::Error.WriteLine('excel lost during ' + $e.name + ' (attempt ' + $attempt + ')')
      $line = $null
    }
    if ($null -eq $line) { $line = [ordered]@{ name = $e.name; open = 'failed'; log = 'excel kept dying'; sheets = -1; names = $false } }
    $ok = ($line['open'] -eq 'normal') -and $line['names'] -and (-not $line['lint']) -and ($line['shapeDiffs'] -eq 0) -and ($line['chartDiffs'] -eq 0) -and ($line['noteDiffs'] -eq 0) -and ($line['tableDiffs'] -eq 0) -and ($line['pivotDiffs'] -eq 0) -and ($line['newErr'] -le 0)
    if (-not $ok) { $bad++ }
    $parts = @()
    foreach ($k in $line.Keys) {
      $v = $line[$k]
      if ($v -is [string]) { $parts += ((JStr $k) + ':' + (JStr $v)) }
      elseif ($v -is [bool]) { $parts += ((JStr $k) + ':' + ([string]$v).ToLower()) }
      elseif ($v -is [array]) { $parts += ((JStr $k) + ':[' + (($v | ForEach-Object { JStr ([string]$_) }) -join ',') + ']') }
      else { $parts += ((JStr $k) + ':' + [string]$v) }
    }
    Write-Host ('{' + ($parts -join ',') + ',"ok":' + ([string]$ok).ToLower() + '}')
  }
  Write-Host ('{"type":"summary","files":' + $count + ',"notOk":' + $bad + '}')
}
finally {
  foreach ($k in @($srcCache.Keys)) { try { if ($srcCache[$k].wb) { $srcCache[$k].wb.Close($false) } } catch {} }
  Close-GExcel $xl
}
