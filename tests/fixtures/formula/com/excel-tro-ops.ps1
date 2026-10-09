param(
  [Parameter(Mandatory = $true)][string]$OutDir,
  [string]$PidFile = ''
)
# Trim-range references (A1:.B9 / A1.:B9 / A1.:.B9) under structural ops: for every op a fresh workbook gets the same
# formulas (typed through Range.Formula2), the op is applied (insert / delete / clip-free move of columns, insert /
# delete of rows on sheet Data), Formula2 is read back and the workbook is saved as <OutDir>\tro-<op>.xlsx so that
# extract-tro-ops.mjs can read the stored <f> texts. Result: tests/fixtures/formula/tro-ops-excel-results.json.
# Safety: own hidden Excel instance only (PID recorded; never touches other Excel processes); Quit in finally;
# no clipboard (clip-free move = Insert blank + Range.Cut(Destination) + Delete), no PageSetup. ASCII-only stdout.
$ErrorActionPreference = 'Stop'
Add-Type @"
using System; using System.Runtime.InteropServices;
public static class W32TRO { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid); }
"@
$cells = @(
  @('Other', 'A1', '=SUM(Data!B1:.D10)'),
  @('Other', 'A2', '=SUM(Data!C:.C)'),
  @('Other', 'A3', '=SUM(Data!B1.:.E5)'),
  @('Other', 'A4', '=ROWS(Data!$C$2:.$C$9)'),
  @('Other', 'A5', '=SUM(Data!B2.:F2)'),
  @('Other', 'A6', '=SUM(Data!C1:.C3)'),
  @('Other', 'A7', '=SUM(Data!$2:.$4)'),
  @('Other', 'A8', '=SUM(Data!B.:D)'),
  @('Data', 'K11', '=SUM(B1:.D3)'),
  @('Data', 'K12', '=SUM(C:.E)')
)
$ops = @('insC', 'delC', 'delCD', 'delBF', 'movDtoB', 'movBtoF', 'insRow3', 'delRow2', 'delRows1to10')
$results = New-Object System.Collections.ArrayList
$xl = $null; $xlPid = 0
try {
  $xl = New-Object -ComObject Excel.Application
  $p = [uint32]0; [W32TRO]::GetWindowThreadProcessId([IntPtr]$xl.Hwnd, [ref]$p) | Out-Null; $xlPid = $p
  if ($PidFile) { [System.IO.File]::WriteAllText($PidFile, [string]$xlPid) }
  Write-Output "excel pid $xlPid"
  $xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false
  foreach ($op in $ops) {
    $wb = $xl.Workbooks.Add()
    try {
      $D = $wb.Worksheets.Item(1); $D.Name = 'Data'
      $O = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $D); $O.Name = 'Other'
      for ($r = 1; $r -le 6; $r++) { for ($c = 2; $c -le 6; $c++) { $D.Cells.Item($r, $c).Value2 = [double]($r * 10 + $c) } }
      foreach ($x in $cells) { $wb.Worksheets.Item($x[0]).Range($x[1]).Formula2 = $x[2] }
      switch ($op) {
        'insC' { [void]$D.Columns.Item(3).Insert() }
        'delC' { [void]$D.Columns.Item(3).Delete() }
        'delCD' { [void]$D.Range('C:D').Delete() }
        'delBF' { [void]$D.Range('B:F').Delete() }
        'delRows1to10' { [void]$D.Range('1:10').Delete() }
        'movDtoB' { [void]$D.Range('B:B').Insert(); [void]$D.Range('E:E').Cut($D.Range('B:B')); [void]$D.Range('E:E').Delete() }
        'movBtoF' { [void]$D.Range('F:F').Insert(); [void]$D.Range('B:B').Cut($D.Range('F:F')); [void]$D.Range('B:B').Delete() }
        'insRow3' { [void]$D.Rows.Item(3).Insert() }
        'delRow2' { [void]$D.Rows.Item(2).Delete() }
      }
      # every formula cell after the op (cells on Data move with the op; the test maps them with mapIndex)
      foreach ($sn in @('Other', 'Data')) {
        $sh = $wb.Worksheets.Item($sn)
        foreach ($rg in $sh.UsedRange.Cells) {
          if ($rg.HasFormula) { [void]$results.Add([ordered]@{ op = $op; sheet = $sn; cell = $rg.Address($false, $false); formula2 = [string]$rg.Formula2 }) }
        }
      }
      $out = Join-Path $OutDir ('tro-' + $op + '.xlsx')
      if (Test-Path $out) { Remove-Item $out -Force }
      $wb.SaveAs($out, 51)
    } finally {
      $wb.Close($false)
    }
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
[System.IO.File]::WriteAllText((Join-Path $OutDir 'tro-ops-com.json'), $json, (New-Object System.Text.UTF8Encoding($false)))
Write-Output "done $($results.Count) cells"
