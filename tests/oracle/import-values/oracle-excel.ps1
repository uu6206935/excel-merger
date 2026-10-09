# Excel half of the import-values oracle (task F1): ONE hidden Excel instance of our own for all jobs.
#   powershell -NoProfile -ExecutionPolicy Bypass -File oracle-excel.ps1 -JobsFile <jobs.json> -OutDir <dir>
# jobs.json: [{id, file, kind: 'output'|'reduced', sheet, deleteRows?}] (written by oracle.ts).
#   output : opens <file> (a premerged package) - a file that needs repair is reported as "repair" - recalculates fully (CalculateFull) and
#            dumps the used range of <sheet> to <OutDir>\<id>.values.json
#   reduced: opens <file> (the Excel-saved twin of a source), deletes the listed data rows of <sheet> in memory (Excel shrinks the ranges
#            that cover them), recalculates and dumps the same way (nothing is saved)
# Dump format: {"open":"normal|repair|failed","log":"..","rows":[[cell,..],..]} ; cell = null | true/false | number | "text" | {"e":<excel error code>}
# ASCII-only stdout; only our own Excel PID is ever killed (common.ps1).
param([Parameter(Mandatory = $true)][string]$JobsFile, [Parameter(Mandatory = $true)][string]$OutDir)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '..\..\golden\com\common.ps1')
$M = $script:GMissing
$inv = [Globalization.CultureInfo]::InvariantCulture
$jobs = Get-Content -Raw -Encoding UTF8 $JobsFile | ConvertFrom-Json

function Cell-Json($x) {
  if ($null -eq $x) { return 'null' }
  if ($x -is [bool]) { if ($x) { return 'true' } else { return 'false' } }
  if ($x -is [int]) { return ('{"e":' + $x + '}') }
  if ($x -is [double]) { return $x.ToString('R', $inv) }
  return (JStr ([string]$x))
}

function Dump-Sheet($ws) {
  $ur = $ws.UsedRange
  $maxR = $ur.Row + $ur.Rows.Count - 1
  $maxC = $ur.Column + $ur.Columns.Count - 1
  $vals = $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item($maxR, $maxC)).Value2
  $rows = New-Object System.Collections.Generic.List[string]
  if ($vals -isnot [System.Array]) {
    $rows.Add('[' + (Cell-Json $vals) + ']')
  } else {
    for ($r = 1; $r -le $maxR; $r++) {
      $cells = New-Object System.Collections.Generic.List[string]
      for ($c = 1; $c -le $maxC; $c++) { $cells.Add((Cell-Json $vals[$r, $c])) }
      $rows.Add('[' + ($cells -join ',') + ']')
    }
  }
  return ($rows -join ',')
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
        try { $wb = $xl.Workbooks.Open($j.file, 0, $true, 5, '', '', $true, 2, '', $false, $false, $M, $false, $false, 1); $mode = 'repair' } catch { $mode = 'failed'; $log = ($_.Exception.Message -replace '[^\x20-\x7E]', '?') }
      }
      if ($mode -eq 'repair') {
        Start-Sleep -Milliseconds 200
        foreach ($l in @(Get-ChildItem $env:TEMP -Filter 'error*.xml' -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime -ge $t0.AddSeconds(-1) })) {
          $x = [IO.File]::ReadAllText($l.FullName)
          $log += ([regex]::Matches($x, '<(removed|repaired)[A-Za-z]*>[^<]*') | ForEach-Object { $_.Value }) -join ' ; '
        }
        $log = $log -replace '[^\x20-\x7E]', '?'
      }
      if ($wb) {
        try {
          $ws = $wb.Worksheets.Item([string]$j.sheet)
          if ($j.kind -eq 'reduced') { foreach ($r in @($j.deleteRows)) { [void]$ws.Rows.Item([int]$r).Delete() } }
          $xl.CalculateFull()
          $body = Dump-Sheet $ws
        } catch { $mode = 'failed'; $log = ($_.Exception.Message -replace '[^\x20-\x7E]', '?') }
        try { $wb.Close($false) } catch {}
      }
      if ($mode -ne 'failed') { break }
      Start-Sleep -Milliseconds 1500
    }
    $json = '{"open":"' + $mode + '","log":' + (JStr $log) + ',"rows":[' + $body + ']}'
    [IO.File]::WriteAllText((Join-Path $OutDir ($j.id + '.values.json')), $json, (New-Object System.Text.UTF8Encoding($false)))
    Write-Host ('{"type":"job","id":' + (JStr $j.id) + ',"open":"' + $mode + '"}')
  }
} finally {
  Close-GExcel $xl
}
