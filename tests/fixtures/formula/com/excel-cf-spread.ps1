param(
  [Parameter(Mandatory = $true)][string]$OutDir,
  [string]$PidFile = ''
)
# Which conditional-format rule types / data validations grow over inserted columns and spread onto a moved block
# (round 2 of the formula engine; complements tests/golden/com/probe-cf-leak.ps1 of the patcher).
# One band of 2 rows per (rule type x width): N = area C<r>:C<r+1>, W = area C<r>:E<r+1>; plus two replicas of R2b
# cases (cellIs "=$C$5" on J:K, list validation "=$B$2:$H$2" on a single cell of column B). For every op a fresh
# workbook is built, the op is applied (moves = the runner's clipboard-free sequence: insert blank + Range.Cut
# (Destination) + delete) and the workbook is saved as <OutDir>\cfs-<op>.xlsx (cfs-before.xlsx = no op).
# extract-cf-spread.mjs reads the saved files into ../cf-spread-excel-results.json (tests/formula/cfSpread.test.ts).
# Safety: own hidden Excel instance only (PID from Application.Hwnd; only that PID is ever killed); Quit in finally;
# Range.Cut(Destination) empties the Windows clipboard, so the user's clipboard is saved first and put back at the
# end. No PageSetup. Stdout is ASCII only.
$ErrorActionPreference = 'Stop'
Add-Type @"
using System; using System.Runtime.InteropServices;
public static class W32CFS { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid); }
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

# band layout: types x widths, 3 rows per band starting at row 2
$types = @('cellIs', 'cellIsRef', 'expression', 'colorScale', 'dataBar', 'iconSet', 'top10', 'uniqueValues', 'duplicateValues',
  'aboveAverage', 'containsText', 'containsBlanks', 'notContainsBlanks', 'containsErrors', 'notContainsErrors', 'beginsWith',
  'timePeriod', 'dvWhole', 'dvList', 'dvCustom')
$widths = @('N', 'W')
# rules whose formulas point at / around the moved blocks (does a rule spread onto a block its formula refers to?)
# relative references that point across the op boundary from some cells of a wide area (W = C:E, V = C:G): does
# Excel adjust them per cell or once per fragment?
$straddle = @(
  @('expr', '=E{r}>0'), @('expr', '=B{r}>0'), @('expr', '=D{r}>E{r}'), @('expr', '=SUM(C{r}:E{r})>0'), @('expr', '=$F{r}>F{r}'),
  @('dvCustom', '=F{r}>0')
)
$extras = @(
  @('cellIs', '=$E$1'), @('cellIs', '=SUM($A$1:$Z$1)'), @('cellIs', '=SUM($F$1:$H$1)'), @('expr', '=G{r}>0'),
  @('expr', '=$G{r}>0'), @('cellIs', '=$G$1'), @('cellIs', '=$D$1'), @('dvList', '=$G$1:$G$5'), @('dvCustom', '=$E{r}>0'),
  @('cellIs', '=$B$1'), @('expr', '=$H{r}>0'), @('dvList', '=$H$1:$H$5'), @('cellIs', '=$I$1'), @('expr', '=SUM($A{r}:$Z{r})>0'),
  @('expr', '=COUNT($F{r}:$G{r})>0'), @('cellIs', '=SUM($C$1:$D$1)'), @('cellIs', '=SUM($E$1:$F$1)'), @('cellIs', '=SUM($1:$1)'),
  @('cellIs', '=Other!$G$1'), @('cellIs', '=nmG'), @('expr', '=E{r}>0'), @('cellIs', '=$G{r}'), @('expr', '=G$1>0')
)

function Add-Rule($ws, [string]$t, [string]$addr, [int]$r) {
  $rg = $ws.Range($addr)
  switch ($t) {
    'cellIs' { $f = $rg.FormatConditions.Add(1, 5, '=5'); $f.Interior.Color = 255 }
    'cellIsRef' { $f = $rg.FormatConditions.Add(1, 5, '=$A$1'); $f.Interior.Color = 33023 }
    'expression' { $f = $rg.FormatConditions.Add(2, $M, "=C$($r)>0"); $f.Interior.Color = 65280 }
    'colorScale' { [void]$rg.FormatConditions.AddColorScale(3) }
    'dataBar' { [void]$rg.FormatConditions.AddDatabar() }
    'iconSet' { [void]$rg.FormatConditions.AddIconSetCondition() }
    'top10' { $f = $rg.FormatConditions.AddTop10(); $f.Interior.Color = 16711680 }
    'uniqueValues' { $f = $rg.FormatConditions.AddUniqueValues(); $f.DupeUnique = 0; $f.Interior.Color = 65535 }
    'duplicateValues' { $f = $rg.FormatConditions.AddUniqueValues(); $f.DupeUnique = 1; $f.Interior.Color = 10092543 }
    'aboveAverage' { $f = $rg.FormatConditions.AddAboveAverage(); $f.Interior.Color = 16776960 }
    'containsText' { $f = $rg.FormatConditions.Add(9, $M, $M, $M, 'x', 0); $f.Interior.Color = 12632256 }
    'containsBlanks' { $f = $rg.FormatConditions.Add(10); $f.Interior.Color = 8421504 }
    'notContainsBlanks' { $f = $rg.FormatConditions.Add(13); $f.Interior.Color = 8421376 }
    'containsErrors' { $f = $rg.FormatConditions.Add(16); $f.Interior.Color = 32896 }
    'notContainsErrors' { $f = $rg.FormatConditions.Add(17); $f.Interior.Color = 8388736 }
    'beginsWith' { $f = $rg.FormatConditions.Add(9, $M, $M, $M, 'x', 2); $f.Interior.Color = 16744576 }
    'timePeriod' { $f = $rg.FormatConditions.Add(11, $M, $M, $M, $M, $M, 2); $f.Interior.Color = 13421619 }
    'dvWhole' { [void]$rg.Validation.Add(1, 1, 1, '0', '1000') }
    'dvList' { [void]$rg.Validation.Add(3, 1, 1, '=$A$1:$A$5') }
    'dvCustom' { [void]$rg.Validation.Add(7, 1, 1, "=C$($r)>0") }
  }
}

function Build($xl) {
  $wb = $xl.Workbooks.Add()
  $ws = $wb.Worksheets.Item(1)
  $ws.Name = 'P'
  $oth = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $ws); $oth.Name = 'Other'
  [void]$wb.Names.Add('nmG', '=P!$G$1')
  $last = 2 + 3 * ($types.Count * $widths.Count + 2 + $extras.Count * $widths.Count + $straddle.Count * 2)
  $vals = New-Object 'object[,]' $last, 20
  for ($r = 1; $r -le $last; $r++) { for ($c = 1; $c -le 20; $c++) { $vals[($r - 1), ($c - 1)] = [double]($r * 100 + $c) } }
  $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item($last, 20)).Value2 = $vals
  [void]$ws.Activate()
  $k = 0
  foreach ($w in $widths) {
    foreach ($t in $types) {
      $r = 2 + 3 * $k
      [void]$ws.Range("C$r").Select()   # relative CF formulas are interpreted relative to the active cell
      $addr = if ($w -eq 'N') { "C$($r):C$($r + 1)" } else { "C$($r):E$($r + 1)" }
      try { Add-Rule $ws $t $addr $r } catch { Write-Host ("rule failed " + $t + ' ' + $w + ': ' + ($_.Exception.Message -replace '[^\x20-\x7e]', '?')) }
      $k++
    }
  }
  # R2b replicas
  $r = 2 + 3 * $k; [void]$ws.Range("J$r").Select(); $f = $ws.Range("J$($r):K$($r + 1)").FormatConditions.Add(1, 5, '=$C$5'); $f.Interior.Color = 6697881; $k++
  $r = 2 + 3 * $k; [void]$ws.Range("B$r").Validation.Add(3, 1, 1, '=$B$2:$H$2'); $k++
  $color = 1000
  foreach ($w in $widths) {
    foreach ($x in $extras) {
      $r = 2 + 3 * $k
      [void]$ws.Range("C$r").Select()
      $rg = $ws.Range($(if ($w -eq 'N') { "C$($r):C$($r + 1)" } else { "C$($r):E$($r + 1)" }))
      $fx = $x[1].Replace('{r}', [string]$r)
      try {
        switch ($x[0]) {
          'cellIs' { $f = $rg.FormatConditions.Add(1, 5, $fx); $f.Interior.Color = $color }
          'expr' { $f = $rg.FormatConditions.Add(2, $M, $fx); $f.Interior.Color = $color }
          'dvList' { [void]$rg.Validation.Add(3, 1, 1, $fx) }
          'dvCustom' { [void]$rg.Validation.Add(7, 1, 1, $fx) }
        }
      } catch { Write-Host ("extra rule failed " + $fx + ': ' + ($_.Exception.Message -replace '[^ -~]', '?')) }
      $color += 7919; $k++
    }
  }
  foreach ($w in @('W', 'V')) {
    foreach ($x in $straddle) {
      $r = 2 + 3 * $k
      [void]$ws.Range("C$r").Select()
      $rg = $ws.Range($(if ($w -eq 'W') { "C$($r):E$($r + 1)" } else { "C$($r):G$($r + 1)" }))
      $fx = $x[1].Replace('{r}', [string]$r)
      try {
        switch ($x[0]) {
          'expr' { $f = $rg.FormatConditions.Add(2, $M, $fx); $f.Interior.Color = $color }
          'dvCustom' { [void]$rg.Validation.Add(7, 1, 1, $fx) }
        }
      } catch { Write-Host ("straddle rule failed " + $fx) }
      $color += 7919; $k++
    }
  }
  [void]$ws.Range('A1').Select()
  return $wb
}

# op name -> [kind, a, b, c]: ins(at, count) / del(at, count) / mov(from, count, to) on columns
$ops = [ordered]@{
  'before'   = @('none', 0, 0, 0)
  'insD'     = @('ins', 4, 1, 0)      # right after N (C), inside W (C:E)
  'insF2'    = @('ins', 6, 2, 0)      # right after W (C:E)
  'insC'     = @('ins', 3, 1, 0)      # at the first column (shift only)
  'delD'     = @('del', 4, 1, 0)      # inside W
  'movGtoD'  = @('mov', 7, 1, 4)      # from the right: lands right after N, inside W
  'movAtoD'  = @('mov', 1, 1, 4)      # from the left: N shifts to B, block lands right after it; inside W
  'movHtoF'  = @('mov', 8, 1, 6)      # lands right after W
  'movIJtoF' = @('mov', 9, 2, 6)      # 2 columns, right after W
  'movBHtoL' = @('mov', 2, 7, 12)     # R2b cf_m15 (the whole N / W areas travel with the block)
  'movFGtoC' = @('mov', 6, 2, 3)      # R2b dv m06 (right after the B validation; right before N / W)
  'movDtoH'  = @('mov', 4, 1, 8)      # a column of W moves away (per-cell split)
  'insF'     = @('ins', 6, 1, 0)      # inside V, right after W
  'delF'     = @('del', 6, 1, 0)      # inside V, right after W
  'delEF'    = @('del', 5, 2, 0)      # end of W, inside V
  'insA'     = @('ins', 1, 1, 0)      # before everything (every area shifts)
  'delB'     = @('del', 2, 1, 0)      # before every area (every area shifts left)
}

$clip = Save-Clip
$xl = $null; $xlPid = 0
try {
  $xl = New-Object -ComObject Excel.Application
  $p = [uint32]0; [W32CFS]::GetWindowThreadProcessId([IntPtr]$xl.Hwnd, [ref]$p) | Out-Null; $xlPid = $p
  if ($PidFile) { [System.IO.File]::WriteAllText($PidFile, [string]$xlPid) }
  Write-Output "excel pid $xlPid"
  $xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false; $xl.EnableEvents = $false
  foreach ($name in $ops.Keys) {
    $o = $ops[$name]
    $wb = @(Build $xl)[-1]
    try {
      $ws = $wb.Worksheets.Item(1)
      # column ranges by address (a scriptblock returning a multi-column Range would enumerate it into columns)
      $L = { param($i) $s = ''; $n = [int]$i; while ($n -gt 0) { $m = ($n - 1) % 26; $s = [string][char](65 + $m) + $s; $n = [math]::Floor(($n - 1) / 26) }; $s }
      $A = { param($i, $n) (& $L $i) + ':' + (& $L ($i + $n - 1)) }
      switch ($o[0]) {
        'ins' { [void]$ws.Range((& $A $o[1] $o[2])).Insert(-4161) }
        'del' { [void]$ws.Range((& $A $o[1] $o[2])).Delete(-4159) }
        'mov' {
          $from = [int]$o[1]; $n = [int]$o[2]; $to = [int]$o[3]
          [void]$ws.Range((& $A $to $n)).Insert(-4161)
          $src = if ($to -lt $from) { $from + $n } else { $from }
          [void]$ws.Range((& $A $src $n)).Cut($ws.Range((& $A $to $n)))
          [void]$ws.Range((& $A $src $n)).Delete(-4159)
        }
      }
      $out = Join-Path $OutDir ('cfs-' + $name + '.xlsx')
      if (Test-Path $out) { Remove-Item $out -Force }
      $wb.SaveAs($out, 51)
      Write-Output "saved $name"
    } catch {
      Write-Output ('ERROR ' + $name + ' line ' + $_.InvocationInfo.ScriptLineNumber + ' ' + ($_.Exception.Message -replace '[^\x20-\x7e]', '?'))
    } finally {
      $wb.Close($false)
    }
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
Write-Output 'done'
