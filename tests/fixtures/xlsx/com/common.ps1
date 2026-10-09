# Shared COM helpers for the reader fixtures (owner: reader R).
# Creates OUR OWN hidden Excel instance, records its PID, and only ever kills that PID.
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public static class RdrWin32 {
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);
}
"@ -ErrorAction SilentlyContinue

$script:RdrMissing = [System.Reflection.Missing]::Value
$script:RdrPid = 0

function New-RdrExcel {
  $xl = New-Object -ComObject Excel.Application
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  $xl.ScreenUpdating = $false
  $xl.EnableEvents = $false
  try { $xl.AskToUpdateLinks = $false } catch {}
  try { $xl.AutomationSecurity = 3 } catch {}
  [uint32]$procId = 0
  [void][RdrWin32]::GetWindowThreadProcessId([IntPtr]$xl.Hwnd, [ref]$procId)
  $script:RdrPid = $procId
  try { [IO.File]::AppendAllText((Join-Path $env:TEMP 'reader-excel-pids.txt'), ("{0} {1:o}`r`n" -f $procId, (Get-Date))) } catch {}
  return $xl
}

function Close-RdrExcel($xl) {
  if ($null -eq $xl) { return }
  try { foreach ($wb in @($xl.Workbooks)) { try { $wb.Close($false) } catch {} } } catch {}
  try { $xl.Quit() } catch {}
  try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl) } catch {}
  [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect()
  Start-Sleep -Milliseconds 800
  if ($script:RdrPid) {
    $p = Get-Process -Id $script:RdrPid -ErrorAction SilentlyContinue
    if ($p -and $p.ProcessName -eq 'EXCEL') { Stop-Process -Id $script:RdrPid -Force -ErrorAction SilentlyContinue }
  }
}

# BGR long (Excel .Color) -> 'RRGGBB'
function ConvertTo-Rgb($bgr) {
  if ($null -eq $bgr -or $bgr -is [DBNull]) { return $null }
  $n = [int64][double]$bgr
  if ($n -lt 0) { return $null }
  $r = $n -band 0xFF; $g = ($n -shr 8) -band 0xFF; $b = ($n -shr 16) -band 0xFF
  return ('{0:X2}{1:X2}{2:X2}' -f $r, $g, $b)
}

function Get-Safe([scriptblock]$sb) {
  try { $v = (& $sb); if ($null -eq $v -or $v -is [DBNull]) { return $null }; return $v } catch { return ('ERR: ' + $_.Exception.Message) }
}

# Write JSON as UTF-8 without BOM (Node reads it with JSON.parse)
function Save-Json($obj, [string]$path) {
  $json = $obj | ConvertTo-Json -Depth 10
  [IO.File]::WriteAllText($path, $json, (New-Object System.Text.UTF8Encoding($false)))
}
