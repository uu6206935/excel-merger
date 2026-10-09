# Round 3 probes (owner: patcher P): one hidden Excel instance.
#   powershell -NoProfile -ExecutionPolicy Bypass -File tests\golden\com\probe-r3.ps1 -Raw <r3-misc-raw.xlsx> -FixtureOut tests\fixtures\golden\r3-misc.xlsx -ProbeDir tests\fixtures\golden\probe
# 1. normalizes the hand-built fixture (tools/make-r3-fixture.ts): open + SaveAs, so that Excel's XML (fonts with family /
#    charset / scheme, dimension ...) is what both the replay and the golden start from;
# 2. probe-r3-native-unmerge.xlsx : Range.UnMerge ONLY (no look copy) on the fixture's merged areas;
# 3. probe-r3-merged.xlsx / probe-r3-unmerged.xlsx : merges made by Excel itself (outlined / top-left-only styled areas)
#    before and after UnMerge, and single border writes (does Excel mirror an edge into the neighbour cell?);
# 4. stdout: how a quote prefix can be removed through COM (PrefixCharacter after each technique).
#    (Worksheet.ExportAsFixedFormat / PDF hangs in a hidden instance on this PC: not used.)
# stdout: ASCII JSON lines.
param([string]$Raw, [string]$FixtureOut, [string]$ProbeDir, [switch]$QuoteOnly)
$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\common.ps1"
$Raw = (Resolve-Path $Raw).Path
$FixtureOut = [IO.Path]::GetFullPath($FixtureOut)
$ProbeDir = (Resolve-Path $ProbeDir).Path

function Out-Json($name, $obj) { Write-Host ('{"probe":' + (JStr $name) + ',' + $obj + '}') }

$xl = $null
try {
  $xl = New-GExcel

  if (-not $QuoteOnly) {
  # ---- 1. normalize --------------------------------------------------------------------------------------------
  $wb = $xl.Workbooks.Open($Raw, 0, $true)
  if (Test-Path $FixtureOut) { Remove-Item $FixtureOut -Force }
  $wb.SaveAs($FixtureOut, 51)
  $wb.Close($false)
  Out-Json 'normalize' '"ok":true'

  # ---- 2. native UnMerge of the fixture's areas --------------------------------------------------------------
  $wb = $xl.Workbooks.Open($FixtureOut, 0, $true)
  $ws = $wb.Worksheets.Item('Band')
  foreach ($a in 'A1:F1', 'A3:C5', 'E7:G7', 'A9:C9', 'A11:C11') { [void]$ws.Range($a).UnMerge() }
  $p = Join-Path $ProbeDir 'probe-r3-native-unmerge.xlsx'
  if (Test-Path $p) { Remove-Item $p -Force }
  $wb.SaveAs($p, 51)
  $wb.Close($false)
  Out-Json 'native-unmerge' '"ok":true'

  # ---- 3. merges made by Excel + single border writes --------------------------------------------------------
  $wb = $xl.Workbooks.Open($FixtureOut, 0, $true)
  $ws = $wb.Worksheets.Item('Band')
  # (a) outline around A20:F20, yellow fill on all cells, then merge
  $ws.Range('A20').Value2 = 'band'
  $ws.Range('A20:F20').Interior.Color = 65535
  $ws.Range('A20:F20').BorderAround(1, 2) | Out-Null
  # (b) all four sides + fill on A22 only, then merge A22:F22
  $ws.Range('A22').Value2 = 'tl only'
  $ws.Range('A22').Interior.Color = 65535
  foreach ($e in 7, 8, 9, 10) { $b = $ws.Range('A22').Borders.Item($e); $b.LineStyle = 1; $b.Weight = 2 }
  # (c) outline around A24:C26 (+ fill), merge
  $ws.Range('A24').Value2 = 'area'
  $ws.Range('A24:C26').Interior.Color = 65535
  $ws.Range('A24:C26').BorderAround(1, 2) | Out-Null
  # (d) single border writes on plain cells: right of B30, bottom of B32, all four of E30
  $ws.Range('B30').Borders.Item(10).LineStyle = 1
  $ws.Range('B32').Borders.Item(9).LineStyle = 1
  foreach ($e in 7, 8, 9, 10) { $ws.Range('E30').Borders.Item($e).LineStyle = 1 }
  foreach ($a in 'A20:F20', 'A22:F22', 'A24:C26') { [void]$ws.Range($a).Merge() }
  $p = Join-Path $ProbeDir 'probe-r3-merged.xlsx'
  if (Test-Path $p) { Remove-Item $p -Force }
  $wb.SaveCopyAs($p)
  foreach ($a in 'A20:F20', 'A22:F22', 'A24:C26') { [void]$ws.Range($a).UnMerge() }
  $p = Join-Path $ProbeDir 'probe-r3-unmerged.xlsx'
  if (Test-Path $p) { Remove-Item $p -Force }
  $wb.SaveAs($p, 51)
  $wb.Close($false)
  Out-Json 'excel-made-merges' '"ok":true'
  }

  # ---- 4. removing a quote prefix through COM ---------------------------------------------------------------
  $wb = $xl.Workbooks.Open($FixtureOut, 0, $true)
  $ws = $wb.Worksheets.Item('Q')
  $techniques = [ordered]@{
    'Value2=x'              = { param($c) $c.Value2 = 'x' }
    'Formula=x'             = { param($c) $c.Formula = 'x' }
    'Value2=5'              = { param($c) $c.Value2 = 5 }
    'ClearContents+Value2'  = { param($c) [void]$c.ClearContents(); $c.Value2 = 'x' }
    'General+Value2'        = { param($c) $c.NumberFormat = 'General'; $c.Value2 = 'x' }
    'Text+Value2+General'   = { param($c) $c.NumberFormat = '@'; $c.Value2 = 'x'; $c.NumberFormat = 'General' }
    'ClearFormats+Value2'   = { param($c) [void]$c.ClearFormats(); $c.Value2 = 'x' }
    'Value2=Value2'         = { param($c) $c.Value2 = $c.Value2 }
    'Formula=Formula'       = { param($c) $c.Formula = $c.Formula }
    'Clear+Value2'          = { param($c) [void]$c.Clear(); $c.Value2 = 'x' }
    'Style=Style+Value2'    = { param($c) $c.Style = $c.Style; $c.Value2 = 'x' }
    'TextToColumns'         = { param($c) $c.TextToColumns($c, 1, -4142, $false, $false, $false, $false, $false, $false) | Out-Null }
    'Value2=$null+Value2'   = { param($c) $c.Value2 = $null; $c.Value2 = 'x' }
  }
  $row = 40
  foreach ($k in $techniques.Keys) {
    $c = $ws.Cells.Item($row, 1)
    $c.Formula = "'00123"
    $before = [string]$c.PrefixCharacter
    $err = ''
    try { & $techniques[$k] $c } catch { $err = $_.Exception.Message }
    $after = [string]$c.PrefixCharacter
    $fmt = [string]$c.NumberFormat
    Out-Json ('quote-' + $k) ('"before":' + (JStr $before) + ',"after":' + (JStr $after) + ',"value":' + (JStr ([string]$c.Value2)) + ',"numfmt":' + (JStr $fmt) + ',"err":' + (JStr $err))
    $row++
  }
  # an EMPTY cell that still carries the prefix, then a number typed: does Formula = 5 keep it?
  $c = $ws.Cells.Item($row, 1); $c.Formula = "'7"; [void]$c.ClearContents()
  Out-Json 'quote-cleared-cell' ('"prefix":' + (JStr ([string]$c.PrefixCharacter)))
  $wb.Close($false)
}
finally { Close-GExcel $xl }
