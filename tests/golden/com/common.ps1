# Golden harness COM helpers (owner: patcher P).
# Creates OUR OWN hidden Excel instance, records its PID (from Application.Hwnd) and, on cleanup, kills only that
# PID if it is still alive. Never touches other EXCEL.EXE processes (the user and other tools run Excel too).
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public static class GWin32 {
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);
}
"@ -ErrorAction SilentlyContinue

$script:GMissing = [System.Reflection.Missing]::Value
$script:GExcelPid = 0

function New-GExcel {
  $xl = New-Object -ComObject Excel.Application
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  $xl.ScreenUpdating = $false
  $xl.EnableEvents = $false
  try { $xl.AutomationSecurity = 3 } catch {}
  [uint32]$procId = 0
  [void][GWin32]::GetWindowThreadProcessId([IntPtr]$xl.Hwnd, [ref]$procId)
  $script:GExcelPid = $procId
  try { [IO.File]::AppendAllText((Join-Path $env:TEMP 'excel-merger-golden-pids.txt'), ("{0} {1:o}`r`n" -f $procId, (Get-Date))) } catch {}
  return $xl
}

function Close-GExcel($xl) {
  if ($null -eq $xl) { return }
  try { foreach ($wb in @($xl.Workbooks)) { try { $wb.Close($false) } catch {} } } catch {}
  try { $xl.Quit() } catch {}
  try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl) } catch {}
  [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect()
  Start-Sleep -Milliseconds 800
  if ($script:GExcelPid) {
    $p = Get-Process -Id $script:GExcelPid -ErrorAction SilentlyContinue
    if ($p -and $p.ProcessName -eq 'EXCEL') { Stop-Process -Id $script:GExcelPid -Force -ErrorAction SilentlyContinue }
  }
}

function ColL([int]$n) {
  $s = ''
  while ($n -gt 0) { $m = ($n - 1) % 26; $s = [char](65 + $m) + $s; $n = [math]::Floor(($n - 1) / 26) }
  return $s
}

# Set a cell through .Formula with an invariant-culture string (Value2 = Int32 fails in this setup, R1 §11).
function SetF($rng, $val) {
  $rng.Formula = [string]([System.Convert]::ToString($val, [Globalization.CultureInfo]::InvariantCulture))
}

# Minimal JSON string escaping for ASCII-only stdout / result files.
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
