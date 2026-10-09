param(
  [Parameter(Mandatory = $true)][string]$OutFile,
  [string]$PidFile = ''
)
# Moves of whole columns around a ListObject (table) with Excel's NATIVE "Insert Cut Cells" (Range.Cut() + Range.Insert),
# the only sequence that does not corrupt tables (the clipboard-free one deletes / re-creates the table column):
# which moves does Excel refuse, and what happens to the table range / column order in the others? The question of
# round 2: a move whose destination is right AFTER the last table column (to = c2 + 1) - reorder inside the table,
# refused, or something else?  Table Tbl on D!C3:H10 (headers tC..tH), row-1 markers hA..hT, a value in I3 and K3.
# Every scenario gets a fresh workbook; a failed Insert is retried up to 3 times (the native sequence can fail
# because the clipboard is busy; a refusal of Excel repeats with the same message).
# Output (UTF-8 JSON): { scenarios: [ { id, cut, at, error, tries, table, headers, markers } ] }
# Safety: own hidden Excel instance only (PID from Application.Hwnd; only that PID is ever killed); Quit in finally;
# the user's clipboard is saved first and put back at the end. No PageSetup. Stdout is ASCII only.
$ErrorActionPreference = 'Stop'
Add-Type @"
using System; using System.Runtime.InteropServices;
public static class W32TBL { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid); }
"@
Add-Type -AssemblyName System.Windows.Forms

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

# id, cut range, insert-before column range (whole columns)
$scen = @(
  @('in_to_I',      'E:E', 'I:I'),   # inside -> right after the last column (to = c2 + 1)
  @('in2_to_I',     'D:E', 'I:I'),
  @('first_to_I',   'C:C', 'I:I'),
  @('in_to_H',      'E:E', 'H:H'),   # inside -> the last column (to = c2): reorder
  @('in_to_C',      'E:E', 'C:C'),   # inside -> the first column (to = c1): reorder
  @('whole_to_K',   'C:H', 'K:K'),   # the whole table moves right
  @('cover_to_L',   'C:I', 'L:L'),   # a block covering the table and one more column
  @('out_to_I',     'J:J', 'I:I'),   # from the right, landing right after the table (to = c2 + 1)
  @('outL_to_C',    'A:A', 'C:C')    # from the left, landing at the first table column (to = c1)
)

function Build($xl) {
  $wb = $xl.Workbooks.Add()
  $ws = $wb.Worksheets.Item(1); $ws.Name = 'D'
  for ($c = 1; $c -le 20; $c++) { $ws.Cells.Item(1, $c).Value2 = 'h' + (& $L $c) }
  for ($c = 3; $c -le 8; $c++) { $ws.Cells.Item(3, $c).Value2 = 't' + (& $L $c) }
  for ($r = 4; $r -le 10; $r++) { for ($c = 3; $c -le 8; $c++) { $ws.Cells.Item($r, $c).Value2 = [double]($r * 100 + $c) } }
  $ws.Cells.Item(3, 9).Value2 = 'after'; $ws.Cells.Item(3, 11).Value2 = 'K3'
  $lo = $ws.ListObjects.Add(1, $ws.Range('C3:H10'), [Type]::Missing, 1)
  $lo.Name = 'Tbl'
  [void]$ws.Activate(); [void]$ws.Range('A1').Select()
  return $wb
}

$clip = Save-Clip
$xl = $null; $xlPid = 0
$results = @()
try {
  $xl = New-Object -ComObject Excel.Application
  $p = [uint32]0; [W32TBL]::GetWindowThreadProcessId([IntPtr]$xl.Hwnd, [ref]$p) | Out-Null; $xlPid = $p
  if ($PidFile) { [System.IO.File]::WriteAllText($PidFile, [string]$xlPid) }
  Write-Output "excel pid $xlPid"
  $xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false; $xl.EnableEvents = $false
  foreach ($sc in $scen) {
    $err = $null; $prev = $null; $tries = 0; $res = $null
    while ($tries -lt 3) {
      $tries++
      $wb = Build $xl
      try {
        $ws = $wb.Worksheets.Item('D')
        $err = $null
        try {
          [void]$ws.Range($sc[1]).Cut()
          [void]$ws.Range($sc[2]).Insert(-4161)
        } catch { $err = ($_.Exception.Message -replace '\s+', ' ') }
        try { [void]$xl.GetType().InvokeMember('CutCopyMode', [Reflection.BindingFlags]::SetProperty, $null, $xl, @($false)) } catch {}
        $lo = $null
        try { $lo = $ws.ListObjects.Item('Tbl') } catch {}
        $markers = @(); for ($c = 1; $c -le 14; $c++) { $markers += [string]$ws.Cells.Item(1, $c).Value2 }
        $res = [ordered]@{
          id = $sc[0]; cut = $sc[1]; at = $sc[2]; error = $err; tries = $tries
          table = $(if ($lo) { $lo.Range.Address($false, $false) } else { $null })
          headers = $(if ($lo) { @($lo.HeaderRowRange.Cells | ForEach-Object { [string]$_.Value2 }) } else { @() })
          markers = $markers
          after = [string]$ws.Cells.Item(3, 9).Value2
        }
      } finally { $wb.Close($false) }
      # a refusal of Excel repeats with the same message: stop after two identical errors
      if (-not $err) { break }
      if ($err -eq $prev) { break }
      $prev = $err
    }
    $results += $res
    Write-Output ('done ' + $sc[0] + ' tries ' + $tries + $(if ($res.error) { ' (error)' } else { '' }))
  }
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
$json = (@{ scenarios = $results } | ConvertTo-Json -Depth 6)
[System.IO.File]::WriteAllText($OutFile, $json, (New-Object System.Text.UTF8Encoding($false)))
Write-Output 'done'
