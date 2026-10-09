param([string]$InFile, [string]$OutJson, [string]$OutXlsx)
# Types every input into cells of every number format with Range.FormulaLocal (= typing, ja-JP) in an OWN hidden
# Excel instance; records Value2 / NumberFormat / HasFormula and saves the book for readBook. ASCII-only stdout.
# Used for ../typed-excel.json: run with -InFile probe-input.json (UTF-8: {formats, inputs}), read OutXlsx back with
# readBook and store [input, format index, t, v, resulting numFmt code or 0] per cell (see tests/doc/typed.test.ts).
# Usage: powershell -NoProfile -ExecutionPolicy Bypass -File probe-typed.ps1 -InFile <json> -OutJson <json> -OutXlsx <ascii path>
$ErrorActionPreference = 'Stop'
Add-Type -Namespace D2P -Name U -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(System.IntPtr h, out uint p);'
$cfg = [IO.File]::ReadAllText($InFile, [Text.Encoding]::UTF8) | ConvertFrom-Json
$xl = New-Object -ComObject Excel.Application
$p = [uint32]0
[void][D2P.U]::GetWindowThreadProcessId([IntPtr][int64]$xl.Hwnd, [ref]$p)
$xpid = [int]$p
Write-Output "pid $xpid"
try {
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  $xl.ScreenUpdating = $false
  $wb = $xl.Workbooks.Add()
  $ws = $wb.Worksheets.Item(1)
  $formats = @($cfg.formats)
  $inputs = @($cfg.inputs)
  $nI = $inputs.Count
  $results = New-Object System.Collections.ArrayList
  for ($fi = 0; $fi -lt $formats.Count; $fi++) {
    $col = $fi + 1
    $rng = $ws.Range($ws.Cells.Item(1, $col), $ws.Cells.Item($nI, $col))
    if ([string]$formats[$fi] -ne "General") { try { $rng.NumberFormat = [string]$formats[$fi] } catch { Write-Output ("fmt $fi NumberFormat failed, trying Local"); try { $rng.NumberFormatLocal = [string]$formats[$fi] } catch { Write-Output ("fmt $fi failed"); continue } } }
    for ($ii = 0; $ii -lt $nI; $ii++) {
      $c = $ws.Cells.Item($ii + 1, $col)
      $err = $null
      try { $c.FormulaLocal = [string]$inputs[$ii] } catch { $err = $_.Exception.Message }
      $v = $c.Value2
      $type = if ($null -eq $v) { 'null' } else { $v.GetType().Name }
      $val = $v
      if ($type -eq 'Int32') { $val = [string]$c.Text }
      elseif ($type -eq 'Double') { $val = $v.ToString('R', [Globalization.CultureInfo]::InvariantCulture) }
      [void]$results.Add([ordered]@{ f = $fi; i = $ii; type = $type; v = $val; nf = [string]$c.NumberFormat; hf = [bool]$c.HasFormula; fl = [string]$c.Formula; err = $err })
    }
  }
  $wb.SaveAs($OutXlsx, 51)
  $wb.Close($false)
  $json = ConvertTo-Json -InputObject $results -Depth 3 -Compress
  [IO.File]::WriteAllText($OutJson, $json, (New-Object Text.UTF8Encoding $false))
  Write-Output "done $($results.Count)"
} finally {
  try { $xl.Quit() } catch {}
  [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl)
  $xl = $null
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
  $proc = Get-Process -Id $xpid -ErrorAction SilentlyContinue
  if ($proc -and $proc.ProcessName -eq 'EXCEL' -and -not $proc.WaitForExit(8000)) { $proc.Kill() }
}
