#Requires -Version 5.1
<#
  Dumps a workbook as Excel sees it (hidden Excel instance of its own, read-only open):
  every non-empty cell of every worksheet {r, c, f (Range.Formula when HasFormula), f2 (Range.Formula2), v (Value2;
  errors as 'ERR:#N/A'), t (Text), nf (NumberFormat), ha (HorizontalAlignment), arr (CurrentArray address),
  spill (SpillingToRange of a spill anchor)}, column widths of the first 30 columns, defined names {name, refersTo},
  merged areas, tables {name, range, columns}.
  Writes UTF-8 JSON to -Out. ASCII-only stdout: {"type":"excel","pid":N} then result / error.
#>
param([Parameter(Mandatory = $true)][string]$Path, [Parameter(Mandatory = $true)][string]$Out)
$ErrorActionPreference = 'Stop'
Add-Type -Namespace EmTest -Name User32 -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(System.IntPtr h, out uint p);'
function Say([string]$json) { [Console]::Out.WriteLine($json); [Console]::Out.Flush() }
function Esc([string]$s) {
  $sb = New-Object System.Text.StringBuilder
  foreach ($ch in $s.ToCharArray()) {
    $n = [int]$ch
    if ($ch -eq '"' -or $ch -eq '\') { [void]$sb.Append('\').Append($ch) }
    elseif ($n -lt 32 -or $n -gt 126) { [void]$sb.Append(('\u{0:x4}' -f $n)) }
    else { [void]$sb.Append($ch) }
  }
  return $sb.ToString()
}

$xl = $null; $xlPid = 0; $code = 0
try {
  $xl = New-Object -ComObject Excel.Application
  $p = [uint32]0
  [void][EmTest.User32]::GetWindowThreadProcessId([IntPtr][int64]$xl.Hwnd, [ref]$p)
  $xlPid = [int]$p
  Say ('{"type":"excel","pid":' + $xlPid + '}')
  $xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false; $xl.EnableEvents = $false
  $xl.AutomationSecurity = 3
  # Open(Filename, UpdateLinks, ReadOnly); a repair prompt would fail the open (DisplayAlerts off -> error)
  $wb = $xl.Workbooks.Open($Path, 0, $true)
  $sheets = New-Object System.Collections.ArrayList
  foreach ($ws in @($wb.Worksheets)) {
    $ur = $ws.UsedRange
    $r0 = [int]$ur.Row; $c0 = [int]$ur.Column; $nr = [int]$ur.Rows.Count; $nc = [int]$ur.Columns.Count
    $cells = New-Object System.Collections.ArrayList
    for ($r = 0; $r -lt $nr; $r++) {
      for ($c = 0; $c -lt $nc; $c++) {
        $cell = $ws.Cells.Item($r0 + $r, $c0 + $c)
        $f = [string]$cell.Formula
        $v = $cell.Value2
        # spilled cells of a dynamic array have no Formula of their own but a value: keep them too
        if ($f -eq '' -and $null -eq $v) { continue }
        $hf = [bool]$cell.HasFormula
        if ($v -is [int]) { $v = 'ERR:' + [string]$cell.Text }   # COM error values arrive as Int32 codes
        $o = [ordered]@{ r = $r0 + $r; c = $c0 + $c; hf = $hf; v = $v; t = [string]$cell.Text; nf = [string]$cell.NumberFormat; ha = [int]$cell.HorizontalAlignment }
        if ($hf) { $o['f'] = $f; $o['f2'] = [string]$cell.Formula2 }
        if ([bool]$cell.HasArray) { $o['arr'] = [string]$cell.CurrentArray.Address($false, $false) }
        try { if ([bool]$cell.HasSpill -and $cell.SpillParent.Address() -eq $cell.Address()) { $o['spill'] = [string]$cell.SpillingToRange.Address($false, $false) } } catch {}
        [void]$cells.Add($o)
      }
    }
    $widths = New-Object System.Collections.ArrayList
    for ($c = 1; $c -le 30; $c++) { [void]$widths.Add([double]$ws.Columns.Item($c).ColumnWidth) }
    $tables = New-Object System.Collections.ArrayList
    foreach ($lo in @($ws.ListObjects)) {
      $cols = New-Object System.Collections.ArrayList
      foreach ($lc in @($lo.ListColumns)) { [void]$cols.Add([string]$lc.Name) }
      [void]$tables.Add([ordered]@{ name = [string]$lo.Name; range = [string]$lo.Range.Address($false, $false); columns = $cols })
    }
    $merges = New-Object System.Collections.ArrayList
    if ($ur.MergeCells -ne $false) {
      foreach ($cell in @($ur.Cells)) {
        if ($cell.MergeCells -eq $true -and $cell.MergeArea.Cells.Item(1, 1).Address() -eq $cell.Address()) { [void]$merges.Add([string]$cell.MergeArea.Address($false, $false)) }
      }
    }
    [void]$sheets.Add([ordered]@{ name = [string]$ws.Name; cells = $cells; widths = $widths; tables = $tables; merges = $merges })
  }
  $names = New-Object System.Collections.ArrayList
  foreach ($n in @($wb.Names)) { [void]$names.Add([ordered]@{ name = [string]$n.Name; refersTo = [string]$n.RefersTo }) }
  $json = ([ordered]@{ sheets = $sheets; names = $names }) | ConvertTo-Json -Depth 8 -Compress
  [IO.File]::WriteAllText($Out, $json, (New-Object System.Text.UTF8Encoding $false))
  $wb.Close($false)
  Say '{"type":"result","ok":true}'
} catch {
  $code = 1
  Say ('{"type":"error","ok":false,"message":"' + (Esc ([string]$_.Exception.Message + ' (line ' + $_.InvocationInfo.ScriptLineNumber + ')')) + '"}')
} finally {
  if ($xl) {
    try { foreach ($w in @($xl.Workbooks)) { $w.Close($false) } } catch {}
    try { $xl.Quit() } catch {}
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl)
    $xl = $null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  }
  if ($xlPid -gt 0) {
    $proc = Get-Process -Id $xlPid -ErrorAction SilentlyContinue
    if ($proc -and $proc.ProcessName -eq 'EXCEL' -and -not $proc.WaitForExit(8000)) { try { $proc.Kill() } catch {} }
  }
}
exit $code
