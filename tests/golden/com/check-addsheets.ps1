# Opens the outputs of tests/golden/tools/make-addsheets-outputs.ts (premerged workbooks WITH copied sheets, includeSourceSheets) in ONE hidden Excel
# instance of our own and checks them against the files they were built from (owner: patcher P):
#   powershell -NoProfile -ExecutionPolicy Bypass -File tests\golden\com\check-addsheets.ps1 -Dir <folder with out\ manifest.json> [-Only regex]
# Per manifest entry:
#   open        normal | repair | failed   (a plain Open that throws means Excel wants to repair the file; then it is opened with CorruptLoad and
#                                          Excel's repair log is collected, like tests/golden/com/opentest.ps1)
#   names       the sheet names, in order, equal the expected ones (planSheetNames) and their visibility equals the source's
#   newErr      #REF! / #NAME? cells + names of the copied sheets that the source sheets did not have (after CalculateFull of both)
#   valueDiffs  cells of every copied sheet whose value differs from the source sheet (Value2; a cell with a date format may differ by the 1462 days
#               of the 1904 <-> 1900 shift when the source has the other date system)
#   formulaDiffs  cells whose formula text differs, ignoring sheet prefixes (the references are renamed) — only sheets that were copied completely
#   formatDiffs sampled cells whose number format / font / fill / alignment / merge differs, column widths, tab colour, conditional-format count,
#               validation presence, hyperlink count (full-fidelity sources only)
# stdout: one ASCII JSON line per entry, a last line {"type":"summary",...}. Own instance only: its PID comes from Application.Hwnd and only that PID
# is ever killed.
param([string]$Dir, [string]$Only = '')
$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\common.ps1"
$M = $script:GMissing
$Dir = (Resolve-Path $Dir).Path
$man = Get-Content -Raw -Encoding UTF8 (Join-Path $Dir 'manifest.json') | ConvertFrom-Json
$ERRMIN = -2146826281; $ERRMAX = -2146826239
$REF = -2146826265; $NAME = -2146826259

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

function Last-Cell($ws) {
  $u = $ws.UsedRange
  return @(($u.Row + $u.Rows.Count - 1), ($u.Column + $u.Columns.Count - 1))
}
function Grid($ws, [int]$rows, [int]$cols, [string]$prop) {
  $rg = $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item($rows, $cols))
  $v = $rg.$prop
  if ($rows -eq 1 -and $cols -eq 1) { $a = New-Object 'object[,]' 1, 1; $a[0, 0] = $v; return , $a }
  return , $v
}
function Same($a, $b) {
  if ($null -eq $a -and $null -eq $b) { return $true }
  if ($null -eq $a) { return ($b -is [string] -and $b -eq '') }
  if ($null -eq $b) { return ($a -is [string] -and $a -eq '') }
  if (($a -is [double] -or $a -is [int]) -and ($b -is [double] -or $b -is [int])) { return ([math]::Abs([double]$a - [double]$b) -le 1e-9 * [math]::Max(1.0, [math]::Abs([double]$a))) }
  return ($a -eq $b)
}
function ColL([int]$n) { $s = ''; while ($n -gt 0) { $m = ($n - 1) % 26; $s = [char](65 + $m) + $s; $n = [math]::Floor(($n - 1) / 26) }; return $s }
$PREFIX = "('([^']|'')*'|[^\s'!=+\-*/^&<>(),:;{}\[\]]+)!"
function Strip-Prefix([string]$f) { return [regex]::Replace($f, $PREFIX, '') }
function IsErr($v) { return ($v -is [int] -and $v -le $ERRMAX -and $v -ge $ERRMIN) }

# compare one copied sheet with its source sheet
# a number format that shows a date or a time (outside quoted text and [colour] / [h] / [$-411] parts)
function Is-DateFmt([string]$nf) {
  $t = [regex]::Replace($nf, '"[^"]*"|\[[^\]]*\]|\\.', '')
  return ($t -match '[ymdhs]')
}
# $shift: days that a date moves between the date systems of the source and the output (0 = same system, 1462 = 1904 -> 1900, -1462 = 1900 -> 1904)
function Compare-Pair($xl, $wsS, $wsO, [bool]$formulas, [bool]$formats, [int]$shift) {
  [Console]::Error.WriteLine('  compare ' + $wsO.Name)
  $res = @{ values = 0; formulas = 0; formats = 0; fonts = 0; newErr = 0; first = @(); cells = 0 }
  $ls = Last-Cell $wsS; $lo = Last-Cell $wsO
  $nR = [math]::Min(400, [math]::Max($ls[0], $lo[0])); $nC = [math]::Min(80, [math]::Max($ls[1], $lo[1]))
  $res['cells'] = $nR * $nC
  $vS = Grid $wsS $nR $nC 'Value2'; $vO = Grid $wsO $nR $nC 'Value2'
  $fS = $null; $fO = $null
  if ($formulas) { $fS = Grid $wsS $nR $nC 'Formula'; $fO = Grid $wsO $nR $nC 'Formula' }
  $errS = 0; $errO = 0
  for ($r = 1; $r -le $nR; $r++) {
    for ($c = 1; $c -le $nC; $c++) {
      $a = $vS[$r, $c]; $b = $vO[$r, $c]
      if ($null -eq $a -and $null -eq $b) {
        if ($formulas -and ([string]$fS[$r, $c]).StartsWith('=') -ne ([string]$fO[$r, $c]).StartsWith('=')) { $res['formulas']++ }
        continue
      }
      if ($a -is [int] -and ($a -eq $REF -or $a -eq $NAME)) { $errS++ }
      if ($b -is [int] -and ($b -eq $REF -or $b -eq $NAME)) { $errO++ }
      $same = Same $a $b
      if (-not $same -and $shift -ne 0 -and ($a -is [double] -or $a -is [int]) -and ($b -is [double] -or $b -is [int])) {
        # another date system: a cell that shows a date / time may have moved by the shift (times of day stay), any other number must be equal
        if (Is-DateFmt ([string]$wsS.Cells.Item($r, $c).NumberFormat)) { $same = Same ([double]$a + $shift) $b }
      }
      if (-not $same) {
        $res['values']++
        if ($res['first'].Count -lt 6) { $res['first'] += ((ColL $c) + $r + ': v ' + [string]$a + ' | ' + [string]$b) }
      }
      if ($formulas) {
        $x = [string]$fS[$r, $c]; $y = [string]$fO[$r, $c]
        if ($x.StartsWith('=') -or $y.StartsWith('=')) {
          if ((Strip-Prefix $x) -ne (Strip-Prefix $y)) {
            $res['formulas']++
            if ($res['first'].Count -lt 6) { $res['first'] += ((ColL $c) + $r + ': f ' + $x + ' | ' + $y) }
          }
        }
      }
    }
  }
  $res['newErr'] = $errO - $errS
  [Console]::Error.WriteLine('   values done')
  if ($formats) {
    # sampled cells: number format, font, fill, alignment, merge, borders
    $picked = 0; $stride = [math]::Max(1, [int](($nR * $nC) / 300))
    $k = 0
    for ($r = 1; $r -le $nR -and $picked -lt 60; $r++) {
      for ($c = 1; $c -le $nC -and $picked -lt 60; $c++) {
        $k++
        if ($k % $stride -ne 0) { continue }
        if ($null -eq $vS[$r, $c]) { continue }
        $picked++
        $cs = $wsS.Cells.Item($r, $c); $co = $wsO.Cells.Item($r, $c)
        $keyS = @($cs.NumberFormat, $cs.Font.Bold, $cs.Font.Italic, $cs.Font.Size, $cs.Font.Color, $cs.Interior.Color, $cs.Interior.Pattern, $cs.HorizontalAlignment, $cs.WrapText, $cs.MergeCells, $cs.Borders.Item(9).LineStyle, $cs.Borders.Item(7).LineStyle) -join '|'
        $keyO = @($co.NumberFormat, $co.Font.Bold, $co.Font.Italic, $co.Font.Size, $co.Font.Color, $co.Interior.Color, $co.Interior.Pattern, $co.HorizontalAlignment, $co.WrapText, $co.MergeCells, $co.Borders.Item(9).LineStyle, $co.Borders.Item(7).LineStyle) -join '|'
        # the font NAME of a cell whose font is a theme (scheme) font differs when the two themes name other fonts for the script of the text
        # (Calibri + a Jpan script font in openpyxl's stock theme): counted apart, not a failure of the copy
        if ([string]$cs.Font.Name -ne [string]$co.Font.Name) { $res['fonts']++ }
        if ($keyS -ne $keyO) {
          $res['formats']++
          if ($res['first'].Count -lt 6) { $res['first'] += ((ColL $c) + $r + ': fmt ' + $keyS + ' | ' + $keyO) }
        }
      }
    }
    for ($c = 1; $c -le [math]::Min($nC, 20); $c++) {
      # the width a column HAS (points = pixels x 0.75): stored widths are characters of each workbook's default font, the copy keeps the pixels
      if ([math]::Abs([double]$wsS.Columns.Item($c).Width - [double]$wsO.Columns.Item($c).Width) -gt 1.6) { $res['formats']++; if ($res['first'].Count -lt 6) { $res['first'] += ('col ' + (ColL $c) + ' width ' + $wsS.Columns.Item($c).Width + 'pt | ' + $wsO.Columns.Item($c).Width + 'pt') } }
      if ($wsS.Columns.Item($c).Hidden -ne $wsO.Columns.Item($c).Hidden) { $res['formats']++ }
    }
    try { if ($wsS.Tab.Color -ne $wsO.Tab.Color) { $res['formats']++; if ($res['first'].Count -lt 6) { $res['first'] += ('tab ' + $wsS.Tab.Color + ' | ' + $wsO.Tab.Color) } } } catch {}
    try { if ($wsS.UsedRange.FormatConditions.Count -ne $wsO.UsedRange.FormatConditions.Count) { $res['formats']++; if ($res['first'].Count -lt 6) { $res['first'] += ('cf ' + $wsS.UsedRange.FormatConditions.Count + ' | ' + $wsO.UsedRange.FormatConditions.Count) } } } catch {}
    try { if ($wsS.Hyperlinks.Count -ne $wsO.Hyperlinks.Count) { $res['formats']++; if ($res['first'].Count -lt 6) { $res['first'] += ('links ' + $wsS.Hyperlinks.Count + ' | ' + $wsO.Hyperlinks.Count) } } } catch {}
    # validations: the cells that carry one (SpecialCells(xlCellTypeAllValidation) throws once per sheet when there is none)
    $va = ''; $vo = ''
    try { $va = [string]$wsS.UsedRange.SpecialCells(-4174).Address($false, $false) } catch {}
    try { $vo = [string]$wsO.UsedRange.SpecialCells(-4174).Address($false, $false) } catch {}
    if ($va -ne $vo) { $res['formats']++; if ($res['first'].Count -lt 6) { $res['first'] += ('validation cells ' + $va + ' | ' + $vo) } }
  }
  return $res
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
  $xl = $script:xl; $srcCache = $script:srcCache; $sw = $script:sw
  $t0 = Get-Date
  $outPath = Join-Path $Dir $e.out
  $line = [ordered]@{ name = $e.name; open = 'failed'; log = ''; sheets = -1; names = $false; newErr = 0; valueDiffs = 0; formulaDiffs = 0; formatDiffs = 0; compared = 0; lenientValueDiffs = 0; fontNameDiffs = 0 }
  $o = Open-Book $xl $outPath
  [Console]::Error.WriteLine('opened ' + $e.name + ' ' + $o.mode + ' at ' + $sw.Elapsed.TotalSeconds)
  $line['open'] = $o.mode; $line['log'] = $o.log
  if ($o.wb) {
    $wb = $o.wb
    try { $xl.CalculateFull() } catch {}
    $got = @(); foreach ($s in @($wb.Sheets)) { $got += $s.Name }
    $line['sheets'] = $got.Count
    $exp = @($e.expected)
    $line['names'] = (($got -join '|') -ceq ($exp -join '|'))
    if (-not $line['names']) { $line['gotNames'] = ($got -join '|'); $line['expNames'] = ($exp -join '|') }
    $info = @()
    foreach ($src in @($e.sources)) {
      $sp = [string]$src.file
      if (-not $srcCache.ContainsKey($sp)) {
        $sb = Open-Book $xl $sp
        if ($sb.wb) { try { $xl.CalculateFull() } catch {} }
        $srcCache[$sp] = $sb
      }
      $sb = $srcCache[$sp]
      [Console]::Error.WriteLine(' source ' + $sp + ' ' + $sb.mode)
      if (-not $sb.wb) { $info += ('source ' + $src.label + ' did not open: ' + $sb.mode); continue }
      if ($sb.mode -ne 'normal') { $info += ('source ' + $src.label + ' opened with ' + $sb.mode) }
      foreach ($sh in @($src.sheets)) {
        $wsO = $null; try { $wsO = $wb.Worksheets.Item($sh.name) } catch {}
        $wsS = $null
        foreach ($cand in @($sb.wb.Sheets)) { if ($cand.Name -ceq $sh.original) { $wsS = $cand; break } }
        if (-not $wsO -or -not $wsS) { $line['names'] = $false; $info += ('missing pair ' + $sh.original + ' -> ' + $sh.name); continue }
        # visibility
        if ([int]$wsS.Visible -ne [int]$wsO.Visible) { $line['names'] = $false; $info += ('visibility differs ' + $sh.name) }
        if (-not $src.compare) { continue }
        $formulas = [bool]$sh.complete
        $formats = [bool]$sh.complete
        $shift = 0
        if ([bool]$sb.wb.Date1904 -ne [bool]$wb.Date1904) { if ([bool]$sb.wb.Date1904) { $shift = 1462 } else { $shift = -1462 } }
        try {
          $r = Compare-Pair $xl $wsS $wsO $formulas $formats $shift
          $line['compared'] = $line['compared'] + 1
          if ($src.lenientValues) { $line['lenientValueDiffs'] += $r['values'] } else { $line['valueDiffs'] += $r['values'] }
          $line['fontNameDiffs'] += $r['fonts']
          $line['formulaDiffs'] += $r['formulas']; $line['formatDiffs'] += $r['formats']; $line['newErr'] += $r['newErr']
          if ($r['first'].Count) { $info += ($sh.name + ' => ' + ($r['first'] -join ' ; ')) }
        } catch { $info += ('compare failed for ' + $sh.name + ': ' + (($_.Exception.Message) -replace '[^\x20-\x7e]', '?')) }
      }
    }
    # names of the output that read nothing (#REF!)
    $nameErr = 0
    foreach ($nm in @($wb.Names)) { try { if ([string]$nm.RefersTo -match '#REF!') { $nameErr++ } } catch {} }
    $line['nameErrors'] = $nameErr
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
  $sw = [Diagnostics.Stopwatch]::StartNew()
  Write-Host ('{"type":"excel","pid":' + [int]$script:GExcelPid + ',"version":' + (JStr ([string]$xl.Version)) + '}')
  foreach ($e in @($man)) {
    if ($Only -and ($e.name -notmatch $Only)) { continue }
    $count++
    $line = $null
    for ($attempt = 1; $attempt -le 3; $attempt++) {
      # Excel is ours alone, but if it dies (it did, twice, for no reason of the files) start another instance of our own and redo the entry
      if (-not (Alive)) { Restart-Excel }
      $line = Run-Entry $e
      if (Alive) { break }
      [Console]::Error.WriteLine('excel lost during ' + $e.name + ' (attempt ' + $attempt + ')')
      $line = $null
    }
    if ($null -eq $line) { $line = [ordered]@{ name = $e.name; open = 'failed'; log = 'excel kept dying'; sheets = -1; names = $false; newErr = 0; valueDiffs = 0; formulaDiffs = 0; formatDiffs = 0; compared = 0 } }
    $ok = ($line['open'] -eq 'normal') -and $line['names'] -and ($line['newErr'] -le 0) -and ($line['valueDiffs'] -eq 0) -and ($line['formulaDiffs'] -eq 0) -and ($line['formatDiffs'] -eq 0)
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
