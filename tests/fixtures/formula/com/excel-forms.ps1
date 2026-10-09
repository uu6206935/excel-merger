param(
  [Parameter(Mandatory = $true)][string]$InputJson,
  [Parameter(Mandatory = $true)][string]$OutXlsx,
  [Parameter(Mandatory = $true)][string]$OutJson,
  [string]$PidFile = ''
)
# Types the formulas of forms-input.json into a fresh workbook through Range.Formula2 (what a user types), saves it
# as .xlsx and records Formula2 / Formula read back. extract-forms.mjs then reads the stored <f> texts.
# Safety: own hidden Excel instance only (PID recorded; never touches other Excel processes); Quit in finally;
# no clipboard, no PageSetup. Stdout is ASCII only.
$ErrorActionPreference = 'Stop'
Add-Type @"
using System; using System.Runtime.InteropServices;
public static class W32F { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid); }
"@
$M = [System.Reflection.Missing]::Value
$in = [System.IO.File]::ReadAllText($InputJson, [System.Text.Encoding]::UTF8) | ConvertFrom-Json
$xl = $null; $xlPid = 0
$results = New-Object System.Collections.ArrayList
try {
  $xl = New-Object -ComObject Excel.Application
  $p = [uint32]0; [W32F]::GetWindowThreadProcessId([IntPtr]$xl.Hwnd, [ref]$p) | Out-Null; $xlPid = $p
  if ($PidFile) { [System.IO.File]::WriteAllText($PidFile, [string]$xlPid) }
  Write-Output "excel pid $xlPid"
  $xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false; $xl.AskToUpdateLinks = $false
  $wb = $xl.Workbooks.Add()
  try {
    $F = $wb.Worksheets.Item(1); $F.Name = 'F'
    $S = $wb.Worksheets.Add($M, $F); $S.Name = 'S'
    $J = $wb.Worksheets.Add($M, $S); $J.Name = [regex]::Unescape('\u96c6\u8a08 \u30b7\u30fc\u30c8')
    $T = $wb.Worksheets.Add($M, $J); $T.Name = 'T'
    # table Tbl on T!A1:E4 with awkward column names; E = calculated column
    $hdr = @('a', 'b c', 'd#e', [regex]::Unescape('\u5358\u4fa1'), 'calc')
    for ($k = 0; $k -lt 5; $k++) { $T.Range([string][char](65 + $k) + '1').Value2 = [string]$hdr[$k] }
    for ($rr = 2; $rr -le 4; $rr++) { for ($k = 0; $k -lt 4; $k++) { $T.Range([string][char](65 + $k) + [string]$rr).Value2 = [double]($rr * 10 + $k) } }
    $lo = $T.ListObjects.Add(1, $T.Range('A1:E4'), $M, 1); $lo.Name = 'Tbl'
    $T.Range('E2').Formula2 = '=[@a]*2'
    $sheets = @{ 'F' = $F; 'S' = $S; 'T' = $T }
    $sheets[$J.Name] = $J
    foreach ($c in $in.cells) {
      $ws = $sheets[$c.sheet]
      $rg = $ws.Range($c.cell)
      $used = $null; $err = $null
      foreach ($cand in $c.candidates) {
        try { $rg.Formula2 = $cand; $used = $cand; $err = $null; break } catch { $err = $_.Exception.Message }
      }
      $o = [ordered]@{ sheet = $c.sheet; cell = $c.cell; fn = $c.fn; input = $used; formula2 = $null; formula = $null; err = $err }
      if ($used) { $o.formula2 = [string]$rg.Formula2; $o.formula = [string]$rg.Formula }
      [void]$results.Add($o)
    }
    foreach ($a in @('E2', 'E3')) {
      [void]$results.Add([ordered]@{ sheet = 'T'; cell = $a; fn = $null; input = '=[@a]*2'; formula2 = [string]$T.Range($a).Formula2; formula = [string]$T.Range($a).Formula; err = $null })
    }
    $xl.Calculate()
    if (Test-Path $OutXlsx) { Remove-Item $OutXlsx -Force }
    $wb.SaveAs($OutXlsx, 51)
  } finally {
    $wb.Close($false)
  }
} catch {
  Write-Output ('ERROR line ' + $_.InvocationInfo.ScriptLineNumber + ' ' + ($_.Exception.Message -replace '[^\x20-\x7e]', '?'))
} finally {
  if ($xl) { try { $xl.Quit() } catch {} ; [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl) }
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  Start-Sleep -Milliseconds 1500
  if ($xlPid) { $pr = Get-Process -Id $xlPid -ErrorAction SilentlyContinue; if ($pr) { Stop-Process -Id $xlPid -Force -ErrorAction SilentlyContinue } }
}
$json = $results | ConvertTo-Json -Depth 5
[System.IO.File]::WriteAllText($OutJson, $json, (New-Object System.Text.UTF8Encoding($false)))
Write-Output "done $($results.Count) cells"
