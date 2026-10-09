# Opens workbooks in ONE hidden Excel instance of our own and reports, per file, how they open and which error cells they hold
# (round 5, module H1; same repair detection as tests/golden/com/opentest.ps1).
#   powershell -NoProfile -ExecutionPolicy Bypass -File tests\fixtures\foreign\excel-check.ps1 -Dir <folder> [-Filter *.xlsx] [-Recalc]
# Own instance only: its PID comes from Application.Hwnd, only that PID is ever killed (never by process name); Quit + release
# in finally. Open: a plain Open first (ReadOnly, no link update); with DisplayAlerts=false a file that needs repair throws, then
# it is opened again with CorruptLoad=xlRepairFile and Excel's repair log (%TEMP%\error*.xml) is collected.
# stdout: one ASCII JSON line per file {"file","open":"normal|repair|failed","log","sheets":[{"name","used","errors":{"#REF!":n,..},"first":[..]}]}
param([string]$Dir, [string]$Filter = '*.xlsx', [switch]$Recalc)
$ErrorActionPreference = 'Stop'
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public static class H1Win32 {
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);
}
"@ -ErrorAction SilentlyContinue

function JStr([string]$s) {
  if ($null -eq $s) { return 'null' }
  $sb = New-Object System.Text.StringBuilder
  [void]$sb.Append('"')
  foreach ($ch in $s.ToCharArray()) {
    $c = [int]$ch
    if ($ch -eq '"') { [void]$sb.Append('\"') }
    elseif ($ch -eq '\') { [void]$sb.Append('\\') }
    elseif ($c -lt 32 -or $c -gt 126) { [void]$sb.Append(('\u{0:x4}' -f $c)) }
    else { [void]$sb.Append($ch) }
  }
  [void]$sb.Append('"')
  return $sb.ToString()
}

$ERR = @{ -2146826281 = '#DIV/0!'; -2146826246 = '#N/A'; -2146826259 = '#NAME?'; -2146826288 = '#NULL!'; -2146826252 = '#NUM!'; -2146826265 = '#REF!'; -2146826273 = '#VALUE!'; -2146826240 = '#SPILL!'; -2146826239 = '#CALC!' }
$M = [System.Reflection.Missing]::Value
$xl = $null
$xlPid = 0
try {
  $xl = New-Object -ComObject Excel.Application
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  $xl.ScreenUpdating = $false
  $xl.EnableEvents = $false
  try { $xl.AutomationSecurity = 3 } catch {}
  [uint32]$p = 0
  [void][H1Win32]::GetWindowThreadProcessId([IntPtr]$xl.Hwnd, [ref]$p)
  $xlPid = $p
  Write-Host ('{"type":"excel","pid":' + [int]$xlPid + ',"version":' + (JStr ([string]$xl.Version)) + '}')
  foreach ($f in @(Get-ChildItem -Path $Dir -Filter $Filter | Where-Object { $_.Name -notlike '~$*' } | Sort-Object Name)) {
    $t0 = Get-Date
    $before = @{}
    foreach ($l in @(Get-ChildItem $env:TEMP -Filter 'error*.xml' -ErrorAction SilentlyContinue)) { $before[$l.FullName] = $l.LastWriteTime.Ticks }
    $mode = 'failed'; $wb = $null; $log = ''
    try { $wb = $xl.Workbooks.Open($f.FullName, 0, $true); $mode = 'normal' }
    catch {
      try { $wb = $xl.Workbooks.Open($f.FullName, 0, $true, 5, '', '', $true, 2, '', $false, $false, $M, $false, $false, 1); $mode = 'repair' } catch { $mode = 'failed'; $log = $_.Exception.Message }
    }
    Start-Sleep -Milliseconds 150
    foreach ($l in @(Get-ChildItem $env:TEMP -Filter 'error*.xml' -ErrorAction SilentlyContinue | Where-Object { -not $before.ContainsKey($_.FullName) -or $before[$_.FullName] -ne $_.LastWriteTime.Ticks })) {
      $x = [IO.File]::ReadAllText($l.FullName)
      $log += ([regex]::Matches($x, '<(removed|repaired)[A-Za-z]*>[^<]*') | ForEach-Object { $_.Value }) -join ' ; '
    }
    $sheetJson = @()
    if ($wb) {
      if ($Recalc) { try { $xl.CalculateFull() } catch {} }
      foreach ($ws in @($wb.Worksheets)) {
        $used = ''
        $counts = @{}
        $first = @()
        try {
          $used = [string]$ws.UsedRange.Address($false, $false)
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
                if ($first.Count -lt 6) { $first += ($c.Address($false, $false) + '=' + $name) }
                $n++
                if ($n -ge 2000) { break }
              }
            }
          }
        } catch { $used = 'ERR ' + $_.Exception.Message }
        $cj = ($counts.Keys | Sort-Object | ForEach-Object { (JStr $_) + ':' + $counts[$_] }) -join ','
        $fj = ($first | ForEach-Object { JStr $_ }) -join ','
        $sheetJson += ('{"name":' + (JStr $ws.Name) + ',"used":' + (JStr $used) + ',"errors":{' + $cj + '},"first":[' + $fj + ']}')
      }
      try { $wb.Close($false) } catch {}
    }
    Write-Host ('{"file":' + (JStr $f.Name) + ',"open":"' + $mode + '","log":' + (JStr $log) + ',"sheets":[' + ($sheetJson -join ',') + ']}')
  }
}
finally {
  if ($xl) {
    try { foreach ($wb in @($xl.Workbooks)) { try { $wb.Close($false) } catch {} } } catch {}
    try { $xl.Quit() } catch {}
    try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl) } catch {}
    [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect()
    Start-Sleep -Milliseconds 800
    if ($xlPid) {
      $pr = Get-Process -Id $xlPid -ErrorAction SilentlyContinue
      # only OUR instance, identified by PID (never by name)
      if ($pr -and $pr.ProcessName -eq 'EXCEL') { Stop-Process -Id $xlPid -Force -ErrorAction SilentlyContinue }
    }
  }
}
