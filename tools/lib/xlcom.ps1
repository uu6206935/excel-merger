# Shared Excel COM helpers for tools/*.ps1 (owner: T).
# Rules (docs/ARCHITECTURE.md section 7): own hidden instance only, remember its PID, never kill EXCEL.EXE by
# name, Quit + ReleaseComObject in finally, ASCII-only stdout (JSON lines, non-ASCII escaped as \uXXXX).

Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public static class EmToolsWin32 {
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);
}
"@ -ErrorAction SilentlyContinue

$script:EmMissing = [System.Reflection.Missing]::Value
$script:EmExcelPid = 0

# JSON string escape that keeps stdout ASCII-only.
function ConvertTo-EmJsonString([string]$s) {
  if ($null -eq $s) { return 'null' }
  $sb = New-Object System.Text.StringBuilder ($s.Length + 2)
  [void]$sb.Append('"')
  foreach ($ch in $s.ToCharArray()) {
    $code = [int]$ch
    if ($ch -eq '"') { [void]$sb.Append('\"') }
    elseif ($ch -eq '\') { [void]$sb.Append('\\') }
    elseif ($code -lt 32 -or $code -gt 126) { [void]$sb.Append(('\u{0:x4}' -f $code)) }
    else { [void]$sb.Append($ch) }
  }
  [void]$sb.Append('"')
  return $sb.ToString()
}

# Write one ASCII JSON line to stdout. $pairs: ordered hashtable of scalar values.
function Write-EmLine($pairs) {
  $parts = @()
  foreach ($k in $pairs.Keys) {
    $v = $pairs[$k]
    if ($null -eq $v) { $j = 'null' }
    elseif ($v -is [bool]) { $j = $(if ($v) { 'true' } else { 'false' }) }
    elseif ($v -is [int] -or $v -is [long] -or $v -is [double] -or $v -is [uint32]) { $j = [System.Convert]::ToString($v, [Globalization.CultureInfo]::InvariantCulture) }
    else { $j = ConvertTo-EmJsonString ([string]$v) }
    $parts += ((ConvertTo-EmJsonString ([string]$k)) + ':' + $j)
  }
  [Console]::Out.WriteLine('{' + ($parts -join ',') + '}')
  [Console]::Out.Flush()
}

function New-EmExcel {
  $xl = New-Object -ComObject Excel.Application
  [uint32]$procId = 0
  try { [void][EmToolsWin32]::GetWindowThreadProcessId([IntPtr]$xl.Hwnd, [ref]$procId) } catch {}
  $script:EmExcelPid = $procId
  # announce the PID first, so a Node watchdog can kill exactly this process if we hang
  Write-EmLine ([ordered]@{ type = 'excel'; pid = [int]$procId; version = [string]$xl.Version })
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  $xl.ScreenUpdating = $false
  $xl.EnableEvents = $false
  try { $xl.AskToUpdateLinks = $false } catch {}
  try { $xl.AutomationSecurity = 3 } catch {}   # msoAutomationSecurityForceDisable (no macros)
  return $xl
}

function Close-EmExcel($xl) {
  if ($null -eq $xl) { return }
  try { foreach ($wb in @($xl.Workbooks)) { try { $wb.Close($false) } catch {} } } catch {}
  try { $xl.Quit() } catch {}
  try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl) } catch {}
  [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect()
  if ($script:EmExcelPid) {
    for ($i = 0; $i -lt 20; $i++) {
      $p = Get-Process -Id $script:EmExcelPid -ErrorAction SilentlyContinue
      if (-not $p) { break }
      Start-Sleep -Milliseconds 250
    }
    $p = Get-Process -Id $script:EmExcelPid -ErrorAction SilentlyContinue
    # only OUR instance, identified by PID (never by name)
    if ($p -and $p.ProcessName -eq 'EXCEL') { Stop-Process -Id $script:EmExcelPid -Force -ErrorAction SilentlyContinue }
  }
}

function Release-EmCom($o) {
  if ($null -ne $o -and [System.Runtime.InteropServices.Marshal]::IsComObject($o)) {
    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($o) } catch {}
  }
}

# Invariant-culture number text (Range.Formula = [string] is the robust way to write values via COM).
function Format-EmNumber($n) { return [System.Convert]::ToString([double]$n, [Globalization.CultureInfo]::InvariantCulture) }

# Temporary ASCII-only working directory (Excel/COM and PowerShell are most reliable with ASCII paths).
function New-EmTempDir([string]$prefix) {
  $d = Join-Path $env:TEMP ('{0}-{1}-{2}' -f $prefix, $PID, [Guid]::NewGuid().ToString('N').Substring(0, 8))
  [void](New-Item -ItemType Directory -Path $d -Force)
  return $d
}
