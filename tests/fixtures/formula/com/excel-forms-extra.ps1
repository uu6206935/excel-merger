param(
  [Parameter(Mandatory = $true)][string]$InputJson,
  [Parameter(Mandatory = $true)][string]$OutXlsx,
  [Parameter(Mandatory = $true)][string]$OutJson,
  [string]$PidFile = ''
)
# Second forms probe (see make-forms-extra-input.mjs): types the formulas of forms-extra-input.json into sheet X of a
# fresh workbook through Range.Formula2, saves it as .xlsx and records Formula2 / Formula read back.
# extract-forms.mjs then reads the stored <f> texts.
# Safety: own hidden Excel instance only (PID recorded; never touches other Excel processes); Quit in finally;
# no clipboard, no PageSetup. Stdout is ASCII only.
$ErrorActionPreference = 'Stop'
Add-Type @"
using System; using System.Runtime.InteropServices;
public static class W32FX { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid); }
"@
$M = [System.Reflection.Missing]::Value
$in = [System.IO.File]::ReadAllText($InputJson, [System.Text.Encoding]::UTF8) | ConvertFrom-Json
$xl = $null; $xlPid = 0
$results = New-Object System.Collections.ArrayList
try {
  $xl = New-Object -ComObject Excel.Application
  $p = [uint32]0; [W32FX]::GetWindowThreadProcessId([IntPtr]$xl.Hwnd, [ref]$p) | Out-Null; $xlPid = $p
  if ($PidFile) { [System.IO.File]::WriteAllText($PidFile, [string]$xlPid) }
  Write-Output "excel pid $xlPid"
  $xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false; $xl.AskToUpdateLinks = $false
  $wb = $xl.Workbooks.Add()
  try {
    $X = $wb.Worksheets.Item(1); $X.Name = 'X'
    $X.Range('B1').Value2 = [double]1; $X.Range('B2').Value2 = [double]2; $X.Range('B3').Value2 = [double]3; $X.Range('B4').Value2 = [double]4
    $X.Range('C1').Value2 = [double]5; $X.Range('C2').Value2 = [double]6; $X.Range('C3').Value2 = [double]7
    $X.Range('Z1').Formula2 = '=SEQUENCE(3)'
    # table TblSales on D1:G4 (+ totals row 5): a | b | kingaku | calc
    $hdr = @('a', 'b', [regex]::Unescape('金額'), 'calc')
    for ($k = 0; $k -lt 4; $k++) { $X.Range([string][char](68 + $k) + '1').Value2 = [string]$hdr[$k] }
    for ($rr = 2; $rr -le 4; $rr++) { for ($k = 0; $k -lt 3; $k++) { $X.Range([string][char](68 + $k) + [string]$rr).Value2 = [double]($rr * 10 + $k) } }
    $lo = $X.ListObjects.Add(1, $X.Range('D1:G4'), $M, 1); $lo.Name = 'TblSales'
    $lo.ShowTotals = $true
    $lo.ListColumns.Item(3).TotalsCalculation = 1
    [void]$wb.Names.Add('nmSpill', '=X!$Z$1')
    [void]$wb.Names.Add('nmRange', '=X!$B$1:$B$3')
    [void]$wb.Names.Add('nmRate', '=0.1')
    foreach ($c in $in.cells) {
      $rg = $X.Range($c.cell)
      $used = $null; $err = $null
      foreach ($cand in $c.candidates) {
        try { $rg.Formula2 = $cand; $used = $cand; $err = $null; break } catch { $err = $_.Exception.Message }
      }
      $o = [ordered]@{ sheet = $c.sheet; cell = $c.cell; fn = $c.fn; input = $used; formula2 = $null; formula = $null; err = $err }
      if ($used) { $o.formula2 = [string]$rg.Formula2; $o.formula = [string]$rg.Formula }
      [void]$results.Add($o)
    }
    # auto-filled calculated column cells and the totals row cell
    foreach ($a in @('G3', 'G4', 'F5')) {
      $rg = $X.Range($a)
      [void]$results.Add([ordered]@{ sheet = 'X'; cell = $a; fn = $null; input = $null; formula2 = [string]$rg.Formula2; formula = [string]$rg.Formula; err = $null })
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
