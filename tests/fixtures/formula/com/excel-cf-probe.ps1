param(
  [Parameter(Mandatory = $true)][string]$Spec,
  [Parameter(Mandatory = $true)][string]$OutDir,
  [string]$PidFile = ''
)
# Generic conditional-format / data-validation probe (round 2 of the formula engine).
# -Spec   JSON: { "scenarios": [ { "id": "s1",
#                   "rules": [ { "type": "expr|cellIs|colorScale|dvList|dvCustom|dvWhole", "addr": "B2:E3" or "B2:B3,F2:F3",
#                                "active": "B2", "f": "=C2>0" } ],
#                   "ops":   [ { "kind": "ins", "at": 4, "count": 1 } | { "kind": "del", "at": 4, "count": 1 } |
#                              { "kind": "mov", "from": 2, "count": 1, "to": 6 } | { "kind": "insr"|"delr", "at": 5, "count": 1 } ] } ] }
#   Column ops are whole-column ops; "mov" is the Excel runner's clipboard-free sequence (insert blank columns at `to`,
#   Range.Cut(Destination), delete the emptied source); "tmov" is Excel's native Cut() + Insert (for tables only).
# Every scenario gets a fresh workbook (sheet P, numbers in A1:T40); after the ops it is saved as <OutDir>\<id>.xlsx.
# extract-cf-probe.mjs reads the saved files.
# Safety: own hidden Excel instance only (PID from Application.Hwnd; only that PID is ever killed); Quit in finally;
# Range.Cut(Destination) empties the Windows clipboard, so the user's clipboard is saved first and put back at the
# end. No PageSetup. Stdout is ASCII only.
$ErrorActionPreference = 'Stop'
Add-Type @"
using System; using System.Runtime.InteropServices;
public static class W32CFP { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid); }
"@
Add-Type -AssemblyName System.Windows.Forms
$M = [System.Reflection.Missing]::Value

function Save-Clip {
  try {
    $d = [System.Windows.Forms.Clipboard]::GetDataObject()
    if ($null -eq $d) { return $null }
    $snap = New-Object System.Windows.Forms.DataObject
    $n = 0
    foreach ($f in $d.GetFormats($false)) { try { $v = $d.GetData($f, $false); if ($null -ne $v) { $snap.SetData($f, $false, $v); $n++ } } catch {} }
    if ($n -eq 0) { return $null }
    return $snap
  } catch { return $null }
}
function Restore-Clip($snap) {
  if ($null -eq $snap) { return }
  for ($i = 0; $i -lt 30; $i++) { try { [System.Windows.Forms.Clipboard]::SetDataObject($snap, $true); Write-Output 'clipboard restored'; return } catch { Start-Sleep -Milliseconds 100 } }
  Write-Output 'clipboard restore FAILED'
}
$L = { param($i) $s = ''; $n = [int]$i; while ($n -gt 0) { $m = ($n - 1) % 26; $s = [string][char](65 + $m) + $s; $n = [math]::Floor(($n - 1) / 26) }; $s }
$A = { param($i, $n) (& $L $i) + ':' + (& $L ($i + $n - 1)) }

$specData = Get-Content -Raw -Encoding UTF8 $Spec | ConvertFrom-Json
$clip = Save-Clip
$xl = $null; $xlPid = 0
try {
  $xl = New-Object -ComObject Excel.Application
  $p = [uint32]0; [W32CFP]::GetWindowThreadProcessId([IntPtr]$xl.Hwnd, [ref]$p) | Out-Null; $xlPid = $p
  if ($PidFile) { [System.IO.File]::WriteAllText($PidFile, [string]$xlPid) }
  Write-Output "excel pid $xlPid"
  $xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false; $xl.EnableEvents = $false
  $color = 1000
  foreach ($sc in $specData.scenarios) {
    $wb = $xl.Workbooks.Add()
    try {
      $ws = $wb.Worksheets.Item(1); $ws.Name = 'P'
      $vals = New-Object 'object[,]' 40, 20
      for ($r = 1; $r -le 40; $r++) { for ($c = 1; $c -le 20; $c++) { $vals[($r - 1), ($c - 1)] = [double]($r * 100 + $c) } }
      $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item(40, 20)).Value2 = $vals
      [void]$ws.Activate()
      foreach ($r in $sc.rules) {
        [void]$ws.Range([string]$r.active).Select()
        $rg = $ws.Range([string]$r.addr)
        $color += 7919
        try {
          switch ([string]$r.type) {
            'expr' { $f = $rg.FormatConditions.Add(2, $M, [string]$r.f); $f.Interior.Color = $color }
            'cellIs' { $f = $rg.FormatConditions.Add(1, 5, [string]$r.f); $f.Interior.Color = $color }
            'colorScale' { [void]$rg.FormatConditions.AddColorScale(3) }
            'dvList' { [void]$rg.Validation.Add(3, 1, 1, [string]$r.f) }
            'dvCustom' { [void]$rg.Validation.Add(7, 1, 1, [string]$r.f) }
            'dvWhole' { [void]$rg.Validation.Add(1, 1, 1, '0', '1000') }
          }
        } catch { Write-Output ('rule failed ' + $sc.id + ' ' + ($_.Exception.Message -replace '[^\x20-\x7e]', '?')) }
      }
      [void]$ws.Range('A1').Select()
      foreach ($o in $sc.ops) {
        switch ([string]$o.kind) {
          'ins' { [void]$ws.Range((& $A ([int]$o.at) ([int]$o.count))).Insert(-4161) }
          'del' { [void]$ws.Range((& $A ([int]$o.at) ([int]$o.count))).Delete(-4159) }
          'insr' { [void]$ws.Range(([string]$o.at + ':' + ([int]$o.at + [int]$o.count - 1))).Insert(-4121) }
          'delr' { [void]$ws.Range(([string]$o.at + ':' + ([int]$o.at + [int]$o.count - 1))).Delete(-4162) }
          'mov' {
            $from = [int]$o.from; $n = [int]$o.count; $to = [int]$o.to
            [void]$ws.Range((& $A $to $n)).Insert(-4161)
            $src = if ($to -lt $from) { $from + $n } else { $from }
            [void]$ws.Range((& $A $src $n)).Cut($ws.Range((& $A $to $n)))
            [void]$ws.Range((& $A $src $n)).Delete(-4159)
          }
        }
      }
      $out = Join-Path $OutDir ([string]$sc.id + '.xlsx')
      if (Test-Path $out) { Remove-Item $out -Force }
      $wb.SaveAs($out, 51)
    } catch {
      Write-Output ('ERROR ' + $sc.id + ' line ' + $_.InvocationInfo.ScriptLineNumber + ' ' + ($_.Exception.Message -replace '[^\x20-\x7e]', '?'))
    } finally {
      $wb.Close($false)
    }
  }
  Write-Output ('saved ' + @($specData.scenarios).Count)
} catch {
  Write-Output ('ERROR line ' + $_.InvocationInfo.ScriptLineNumber + ' ' + ($_.Exception.Message -replace '[^\x20-\x7e]', '?'))
} finally {
  if ($xl) { try { $xl.Quit() } catch {} ; [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl) }
  $xl = $null
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  Start-Sleep -Milliseconds 1500
  if ($xlPid) { $pr = Get-Process -Id $xlPid -ErrorAction SilentlyContinue; if ($pr -and $pr.ProcessName -eq 'EXCEL') { Stop-Process -Id $xlPid -Force -ErrorAction SilentlyContinue } }
  Restore-Clip $clip
}
Write-Output 'done'
