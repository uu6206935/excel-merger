# Excel half of the consistency oracle (task V1): ONE hidden Excel instance of our own for a whole batch of workbooks.
#   powershell -NoProfile -ExecutionPolicy Bypass -File excel-batch.ps1 -JobsFile <jobs.json> -OutDir <dir>
# jobs.json: [{id, file, kind: 'dump'|'reduced'|'resave', sheet?, deleteRows?, saveAs?}] (written by oracle.ts).
#   dump   : opens <file> (a package) - a file that needs repair is reported as "repair" - recalculates fully (CalculateFull) and dumps
#            EVERY worksheet (values and formulas of the used range, from A1) to <OutDir>\<id>.values.json
#   reduced: opens <file> (a source workbook), deletes the listed data rows of <sheet> in memory (Excel shrinks the ranges that cover them),
#            recalculates and dumps the same way (nothing is saved)
#   resave : opens <file>, recalculates fully and saves it as <saveAs> (xlsx) - authors the "cached" twin of a generated workbook (results stored)
# Dump format: {"open":"normal|repair|failed","log":"..","sheets":[{"name":"..","rows":R,"cols":C,"v":[[cell,..],..],"f":[["=formula" or "",..],..]}]}
#   cell = null | true/false | number | "text" | {"e":<excel error code>}
# ASCII-only stdout; only our own Excel PID is ever killed (golden\com\common.ps1). No clipboard, no PageSetup.
param([Parameter(Mandatory = $true)][string]$JobsFile, [Parameter(Mandatory = $true)][string]$OutDir)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '..\..\golden\com\common.ps1')
$M = $script:GMissing
$inv = [Globalization.CultureInfo]::InvariantCulture
$jobs = Get-Content -Raw -Encoding UTF8 $JobsFile | ConvertFrom-Json
# Windows PowerShell 5.1 hands a JSON array over as ONE object[]: a single job comes back as a bare object
if ($jobs -isnot [System.Array]) { $jobs = @($jobs) }

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

function Formula-Json($x) {
  if ($null -eq $x) { return '""' }
  $s = [string]$x
  if ($s.Length -gt 0 -and $s[0] -eq '=') { return (JStr $s) }
  return '""'
}

function Dump-Sheet($ws) {
  $ur = $ws.UsedRange
  $maxR = $ur.Row + $ur.Rows.Count - 1
  $maxC = $ur.Column + $ur.Columns.Count - 1
  $rng = $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item($maxR, $maxC))
  $vals = $rng.Value2
  $fms = $rng.Formula
  $vrows = New-Object System.Collections.Generic.List[string]
  $frows = New-Object System.Collections.Generic.List[string]
  if ($vals -isnot [System.Array]) {
    $vrows.Add('[' + (Cell-Json $vals) + ']')
    $frows.Add('[' + (Formula-Json $fms) + ']')
  } else {
    for ($r = 1; $r -le $maxR; $r++) {
      $vc = New-Object System.Collections.Generic.List[string]
      $fc = New-Object System.Collections.Generic.List[string]
      for ($c = 1; $c -le $maxC; $c++) {
        $vc.Add((Cell-Json $vals[$r, $c]))
        $fc.Add((Formula-Json $fms[$r, $c]))
      }
      $vrows.Add('[' + ($vc -join ',') + ']')
      $frows.Add('[' + ($fc -join ',') + ']')
    }
  }
  # a circular reference has no defined value in Excel (the result depends on the order of evaluation): the judge must know
  $circ = ''
  try { $cr = $ws.CircularReference; if ($null -ne $cr) { $circ = [string]$cr.Address($false, $false) } } catch { $circ = '' }
  return ('{"name":' + (JStr ([string]$ws.Name)) + ',"rows":' + $maxR + ',"cols":' + $maxC + ',"circ":' + (JStr $circ) + ',"v":[' + ($vrows -join ',') + '],"f":[' + ($frows -join ',') + ']}')
}

function Dump-Book($wb) {
  $parts = New-Object System.Collections.Generic.List[string]
  foreach ($ws in @($wb.Worksheets)) { $parts.Add((Dump-Sheet $ws)) }
  return ($parts -join ',')
}

$xl = $null
try {
  $xl = New-GExcel
  Write-Host ('{"type":"excel","pid":' + $script:GExcelPid + '}')
  foreach ($j in $jobs) {
    # a transient COM hiccup (the first jobs of a fresh instance) must not look like a broken package: a failed job is tried again
    for ($attempt = 1; $attempt -le 3; $attempt++) {
      $mode = 'failed'; $wb = $null; $log = ''; $body = ''
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
          if ($j.kind -eq 'reduced') {
            $ws0 = $wb.Worksheets.Item([string]$j.sheet)
            foreach ($r in @($j.deleteRows)) { [void]$ws0.Rows.Item([int]$r).Delete() }
          }
          $xl.CalculateFull()
          if ($j.kind -eq 'resave') { $wb.SaveAs([string]$j.saveAs, 51); $body = '' } else { $body = Dump-Book $wb }
        } catch { $mode = 'failed'; $log = [string]$_.Exception.Message + ' @line ' + [string]$_.InvocationInfo.ScriptLineNumber }
        try { $wb.Close($false) } catch {}
      }
      if ($mode -ne 'failed') { break }
      Start-Sleep -Milliseconds 1500
    }
    $json = '{"open":"' + $mode + '","log":' + (JStr $log) + ',"sheets":[' + $body + ']}'
    [IO.File]::WriteAllText((Join-Path $OutDir ($j.id + '.values.json')), $json, (New-Object System.Text.UTF8Encoding($false)))
    Write-Host ('{"type":"job","id":' + (JStr $j.id) + ',"open":"' + $mode + '"}')
  }
} finally {
  Close-GExcel $xl
}
