#Requires -Version 5.1
<#
  Builds an Excel-authored test workbook from a JSON spec (UTF-8), in a hidden Excel instance of its own.
  ASCII-only file on purpose (Japanese names come from the spec file).
    -Spec  spec.json : { sheets: [{ name,
                           cells: [[addr, kind, value]]   kind: 'n' number, 's' text, 'f' formula (Range.Formula),
                                                          'd' dynamic-array formula (Range.Formula2),
                                                          'a' legacy array formula over the range addr (FormulaArray)
                           merges: [addr], tables: [{name, range}], widths: [[col, width]], hidden: [col],
                           autoFilter: addr,
                           formats: [[addr, numberFormat]], align: [[addr, 'center'|'left'|'right'|'centerAcross']],
                           cf: [{ range, type: 'expression'|'cellIs'|'colorScale'|'dataBar', formula?, operator? }],
                           dv: [{ range, type: 'list'|'whole', formula, formula2? }] }],
                 names: [{ name, refersTo }] }
    -Out   output .xlsx
  stdout: ASCII JSON lines ({"type":"excel","pid":N} first, then {"type":"result"...} or {"type":"error"...}).
#>
param([Parameter(Mandatory = $true)][string]$Spec, [Parameter(Mandatory = $true)][string]$Out)
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
# unary comma: a function's returned array is unrolled into the pipeline, so a ONE-element list of pairs
# ([["I3:I12","yyyy/m/d"]]) would come back as the pair itself and be iterated item by item
function Items($x) { if ($null -eq $x) { return , @() }; return , @($x) }

$xl = $null; $xlPid = 0; $code = 0
try {
  $doc = [IO.File]::ReadAllText($Spec, [Text.Encoding]::UTF8) | ConvertFrom-Json
  $xl = New-Object -ComObject Excel.Application
  $p = [uint32]0
  [void][EmTest.User32]::GetWindowThreadProcessId([IntPtr][int64]$xl.Hwnd, [ref]$p)
  $xlPid = [int]$p
  Say ('{"type":"excel","pid":' + $xlPid + '}')
  $xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false; $xl.EnableEvents = $false
  $wb = $xl.Workbooks.Add()
  $want = @($doc.sheets).Count
  while ($wb.Worksheets.Count -lt $want) { [void]$wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)) }
  while ($wb.Worksheets.Count -gt $want) { $wb.Worksheets.Item($wb.Worksheets.Count).Delete() }
  for ($i = 0; $i -lt $want; $i++) { $wb.Worksheets.Item($i + 1).Name = [string]$doc.sheets[$i].name }
  # order: values -> formats / merges / tables / widths -> names -> formulas (formulas may use tables, names,
  # other sheets) -> CF / DV
  function Put-Cells([string]$pass) {
    foreach ($s in @($doc.sheets)) {
      $ws = $wb.Worksheets.Item([string]$s.name)
      foreach ($c in (Items $s.cells)) {
        $kind = [string]$c[1]
        $isF = ($kind -eq 'f' -or $kind -eq 'd' -or $kind -eq 'a')
        if (($pass -eq 'f') -ne $isF) { continue }
        $r = $ws.Range([string]$c[0])
        if ($kind -eq 'f') { $r.Formula = [string]$c[2] }
        elseif ($kind -eq 'd') { $r.Formula2 = [string]$c[2] }
        elseif ($kind -eq 'a') { $r.FormulaArray = [string]$c[2] }
        elseif ($kind -eq 'n') { $r.Value2 = [double]$c[2] }
        else { $r.Value2 = [string]$c[2] }
      }
    }
  }
  Put-Cells 'v'
  $alignCodes = @{ center = -4108; left = -4131; right = -4152; centerAcross = 7 }
  foreach ($s in @($doc.sheets)) {
    $ws = $wb.Worksheets.Item([string]$s.name)
    foreach ($f in (Items $s.formats)) { $ws.Range([string]$f[0]).NumberFormat = [string]$f[1] }
    foreach ($m in (Items $s.merges)) { if ($m) { $ws.Range([string]$m).Merge() } }
    foreach ($a in (Items $s.align)) { $ws.Range([string]$a[0]).HorizontalAlignment = $alignCodes[[string]$a[1]] }
    foreach ($t in (Items $s.tables)) {
      if ($t) { $lo = $ws.ListObjects.Add(1, $ws.Range([string]$t.range), [Type]::Missing, 1); $lo.Name = [string]$t.name }
    }
    foreach ($w in (Items $s.widths)) { if ($w) { $ws.Columns.Item([int]$w[0]).ColumnWidth = [double]$w[1] } }
    foreach ($h in (Items $s.hidden)) { if ($h) { $ws.Columns.Item([int]$h).Hidden = $true } }
    if ($s.autoFilter) { [void]$ws.Range([string]$s.autoFilter).AutoFilter(1) }
  }
  foreach ($n in (Items $doc.names)) { if ($n) { [void]$wb.Names.Add([string]$n.name, [string]$n.refersTo) } }
  Put-Cells 'f'
  foreach ($s in @($doc.sheets)) {
    $ws = $wb.Worksheets.Item([string]$s.name)
    foreach ($r in (Items $s.cf)) {
      $rng = $ws.Range([string]$r.range)
      switch ([string]$r.type) {
        'expression' { $fc = $rng.FormatConditions.Add(2, [Type]::Missing, [string]$r.formula); $fc.Interior.Color = 13551615 }
        'cellIs' { $fc = $rng.FormatConditions.Add(1, [int]$r.operator, [string]$r.formula); $fc.Font.Bold = $true }
        'colorScale' { [void]$rng.FormatConditions.AddColorScale(3) }
        'dataBar' { [void]$rng.FormatConditions.AddDatabar() }
      }
    }
    foreach ($d in (Items $s.dv)) {
      $rng = $ws.Range([string]$d.range)
      $rng.Validation.Delete()
      if ([string]$d.type -eq 'list') { $rng.Validation.Add(3, 1, 1, [string]$d.formula) }
      else { $rng.Validation.Add(1, 1, 1, [string]$d.formula, [string]$d.formula2) }
    }
  }
  $wb.Worksheets.Item(1).Activate()
  if (Test-Path -LiteralPath $Out) { Remove-Item -LiteralPath $Out -Force }
  $wb.SaveAs($Out, 51)
  $wb.Close($false)
  Say ('{"type":"result","ok":true,"out":"' + (Esc $Out) + '"}')
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
