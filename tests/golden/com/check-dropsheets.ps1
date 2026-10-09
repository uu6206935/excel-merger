# Opens the single-sheet outputs of tests/golden/tools/make-dropsheets-outputs.ts in ONE hidden Excel instance of our own (owner: patcher P).
#   powershell -NoProfile -ExecutionPolicy Bypass -File tests\golden\com\check-dropsheets.ps1 -Dir <folder with in\ out\ manifest.json> [-Only name1,name2]
# Per manifest entry (the input and the output are both opened read-only, fully recalculated):
#   open   normal | repair | failed   (a plain Open that throws means Excel wants to repair the file; then it is opened with CorruptLoad
#                                       and Excel's repair log is collected, like tests/golden/com/opentest.ps1)
#   sheets number of worksheets + chart sheets of the output (must be 1)
#   newErr #REF! / #NAME? error cells or names of the output that the input did not have (must be 0)
#   diffs  (compare = true: the output has the layout of its input) cells of the merged sheet whose value differs after recalculation
# stdout: one ASCII JSON line per entry, a last line {"type":"summary",...}. Own instance only: its PID comes from Application.Hwnd and only
# that PID is ever killed.
param([string]$Dir, [string]$Only = '')
$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\common.ps1"
$M = $script:GMissing
$Dir = (Resolve-Path $Dir).Path
$man = Get-Content -Raw -Encoding UTF8 (Join-Path $Dir 'manifest.json') | ConvertFrom-Json
$onlyList = @(); if ($Only) { $onlyList = $Only.Split(',') }

$ERR = @{ -2146826281 = '#DIV/0!'; -2146826246 = '#N/A'; -2146826259 = '#NAME?'; -2146826288 = '#NULL!'; -2146826252 = '#NUM!'; -2146826265 = '#REF!'; -2146826273 = '#VALUE!'; -2146826240 = '#SPILL!'; -2146826239 = '#CALC!' }

function Open-Book($xl, [string]$path) {
  $t0 = Get-Date
  $mode = 'failed'; $wb = $null; $log = ''
  try { $wb = $xl.Workbooks.Open($path, 0, $true); $mode = 'normal' }
  catch {
    try { $wb = $xl.Workbooks.Open($path, 0, $true, 5, '', '', $true, 2, '', $false, $false, $M, $false, $false, 1); $mode = 'repair' } catch { $mode = 'failed'; $log = $_.Exception.Message }
  }
  Start-Sleep -Milliseconds 200
  foreach ($l in @(Get-ChildItem $env:TEMP -Filter 'error*.xml' -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime -ge $t0.AddSeconds(-1) })) {
    $x = [IO.File]::ReadAllText($l.FullName)
    $log += ([regex]::Matches($x, '<(removed|repaired)[A-Za-z]*>[^<]*') | ForEach-Object { $_.Value }) -join ' ; '
  }
  return @{ wb = $wb; mode = $mode; log = $log }
}

# error cells of one sheet: name -> count
function Sheet-Errors($ws) {
  $counts = @{}
  foreach ($kind in @(-4123, 2)) {
    $rg = $null
    try { $rg = $ws.UsedRange.SpecialCells($kind, 16) } catch { $rg = $null }
    if ($rg) {
      $n = 0
      foreach ($c in $rg.Cells) {
        $v = $c.Value2
        $name = '?'
        if ($v -is [int] -and $ERR.ContainsKey([int]$v)) { $name = $ERR[[int]$v] } elseif ($v -is [int]) { $name = '#E' + $v }
        if (-not $counts.ContainsKey($name)) { $counts[$name] = 0 }
        $counts[$name] = $counts[$name] + 1
        $n++; if ($n -ge 5000) { break }
      }
    }
  }
  return $counts
}

function Last-Cell($ws) {
  $u = $ws.UsedRange
  return @(($u.Row + $u.Rows.Count - 1), ($u.Column + $u.Columns.Count - 1))
}

function Grid($ws, [int]$rows, [int]$cols) {
  $rg = $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item($rows, $cols))
  $v = $rg.Value2
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

function Name-Errors($wb) {
  $n = 0
  foreach ($nm in @($wb.Names)) { try { if ([string]$nm.RefersTo -match '#REF!') { $n++ } } catch {} }
  return $n
}

function Total($h) { $s = 0; foreach ($k in @('#REF!', '#NAME?')) { if ($h.ContainsKey($k)) { $s += $h[$k] } }; return $s }

$xl = $null
$bad = 0; $count = 0
try {
  $xl = New-GExcel
  Write-Host ('{"type":"excel","pid":' + [int]$script:GExcelPid + ',"version":' + (JStr ([string]$xl.Version)) + '}')
  foreach ($e in @($man)) {
    if ($onlyList.Count -and ($onlyList -notcontains $e.name)) { continue }
    $count++
    $t0 = Get-Date
    $inPath = Join-Path $Dir ('in\' + $e.name + '.xlsx'); $outPath = Join-Path $Dir ('out\' + $e.name + '.xlsx')
    $line = [ordered]@{ file = $e.name; sheet = $e.sheet; open = 'failed'; log = ''; sheets = -1; newErr = -1; diffs = -1 }
    $inInfo = Open-Book $xl $inPath
    $inErr = @{}; $inNameErr = 0; $inGrid = $null; $inLast = @(0, 0)
    if ($inInfo.wb) {
      try { $xl.CalculateFull() } catch {}
      $ws = $null; try { $ws = $inInfo.wb.Worksheets.Item($e.sheet) } catch {}
      if ($ws) {
        $inErr = Sheet-Errors $ws
        $inLast = Last-Cell $ws
        if ($e.compare) { $inGrid = Grid $ws $inLast[0] $inLast[1] }
      }
      $inNameErr = Name-Errors $inInfo.wb
      # (the grid is kept in memory; the book can be closed)
      try { $inInfo.wb.Close($false) } catch {}
    } else { $line['inOpen'] = $inInfo.mode }
    $o = Open-Book $xl $outPath
    $line['open'] = $o.mode; $line['log'] = $o.log
    if ($o.wb) {
      try { $xl.CalculateFull() } catch {}
      $line['sheets'] = [int]$o.wb.Sheets.Count
      $names = @(); foreach ($s in @($o.wb.Sheets)) { $names += $s.Name }
      $line['sheetNames'] = ($names -join '|')
      $ws2 = $null; try { $ws2 = $o.wb.Worksheets.Item($e.sheet) } catch {}
      $outErr = @{}
      if ($ws2) { $outErr = Sheet-Errors $ws2 }
      $line['inErrors'] = (($inErr.Keys | Sort-Object | ForEach-Object { $_ + '=' + $inErr[$_] }) -join ',')
      $line['outErrors'] = (($outErr.Keys | Sort-Object | ForEach-Object { $_ + '=' + $outErr[$_] }) -join ',')
      $newErr = (Total $outErr) - (Total $inErr) + (Name-Errors $o.wb) - $inNameErr
      $line['newErr'] = $newErr
      $line['names'] = [int]$o.wb.Names.Count
      if ($e.compare -and $inGrid -and $ws2) {
        $oLast = Last-Cell $ws2
        $rows = [math]::Max($inLast[0], $oLast[0]); $cols = [math]::Max($inLast[1], $oLast[1])
        $go = Grid $ws2 $oLast[0] $oLast[1]
        $d = 0; $first = @()
        for ($r = 0; $r -lt $rows; $r++) {
          for ($c = 0; $c -lt $cols; $c++) {
            $a = $null; $b = $null
            if ($r -lt $inLast[0] -and $c -lt $inLast[1]) { $a = $inGrid[($r + 1), ($c + 1)] }
            if ($r -lt $oLast[0] -and $c -lt $oLast[1]) { $b = $go[($r + 1), ($c + 1)] }
            if (-not (Same $a $b)) {
              $d++
              if ($first.Count -lt 8) { $first += ((ColL ($c + 1)) + ($r + 1) + ': in=' + [string]$a + ' out=' + [string]$b) }
            }
          }
        }
        $line['diffs'] = $d
        $line['first'] = $first
      }
      try { $o.wb.Close($false) } catch {}
    }
    $line['ms'] = [int]((Get-Date) - $t0).TotalMilliseconds
    $ok = ($line['open'] -eq 'normal') -and ($line['sheets'] -eq 1) -and ($line['newErr'] -le 0)
    if (-not $ok) { $bad++ }
    $parts = @()
    foreach ($k in $line.Keys) {
      $v = $line[$k]
      if ($v -is [string]) { $parts += ((JStr $k) + ':' + (JStr $v)) }
      elseif ($v -is [array]) { $parts += ((JStr $k) + ':[' + (($v | ForEach-Object { JStr ([string]$_) }) -join ',') + ']') }
      else { $parts += ((JStr $k) + ':' + [string]$v) }
    }
    Write-Host ('{' + ($parts -join ',') + ',"ok":' + ([string]$ok).ToLower() + '}')
  }
  Write-Host ('{"type":"summary","files":' + $count + ',"notOk":' + $bad + '}')
}
finally { Close-GExcel $xl }
