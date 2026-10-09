# Replays golden op sequences in real Excel and saves Excel's result workbooks (owner: patcher P).
#   powershell -NoProfile -ExecutionPolicy Bypass -File tests\golden\com\run-ops.ps1 -Cases cases.json -FixtureDir tests\fixtures\golden -OutDir tests\fixtures\golden\out [-Only name1,name2]
# cases.json = [{ name, fixture, ops: LoggedOp[] }]  (written by `npx tsx tests/golden/tools/cli.ts cases <file>`)
# Moves use the clipboard-free sequence (insert blank at the target, Range.Cut(Destination), delete the source);
# never Cut()+Insert() (the Windows clipboard is shared).
# stdout: one ASCII JSON line per case {"name":..,"ok":true|false,"error":..}
#   Round 3 extras (same Excel instance): -NormalizeRaw <raw.xlsx> -NormalizedName r3-misc.xlsx  (open + SaveAs the hand-built
#   fixture into FixtureDir first), -MergeProbe -ProbeDir <dir>  (what Excel considers the borders of merged areas, which COM
#   write moves a border), -OpenDir <dir>  (afterwards open every .xlsx of that folder: opens normally or needs repair?)
param([string]$Cases, [string]$FixtureDir, [string]$OutDir, [string[]]$Only, [string]$NormalizeRaw, [string]$NormalizedName, [switch]$MergeProbe, [string]$ProbeDir, [string]$OpenDir)
$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\common.ps1"
$FixtureDir = (Resolve-Path $FixtureDir).Path
$OutDir = (Resolve-Path $OutDir).Path
$list = [IO.File]::ReadAllText((Resolve-Path $Cases).Path, [Text.Encoding]::UTF8) | ConvertFrom-Json
if ($Only) { $Only = @($Only | ForEach-Object { $_ -split ',' } | ForEach-Object { $_.Trim() } | Where-Object { $_ }) }

function AxisRange($ws, [string]$axis, [int]$a, [int]$b) {
  # the comma keeps PowerShell from unrolling a multi-column Range into an array
  if ($axis -eq 'col') { return ,$ws.Range("$(ColL $a):$(ColL $b)").EntireColumn }
  return ,$ws.Range("$($a):$($b)").EntireRow
}

function Invoke-Struct($ws, $op) {
  $axis = [string]$op.axis
  if ($op.type -eq 'insert') {
    [void](AxisRange $ws $axis ([int]$op.at) ([int]$op.at + [int]$op.count - 1)).Insert()
  } elseif ($op.type -eq 'delete') {
    [void](AxisRange $ws $axis ([int]$op.at) ([int]$op.at + [int]$op.count - 1)).Delete()
  } else {
    $from = [int]$op.from; $n = [int]$op.count; $to = [int]$op.to
    if ($to -ge $from -and $to -le $from + $n) { return }
    $blk = AxisRange $ws $axis $to ($to + $n - 1)
    [void]$blk.Insert()
    $blk = AxisRange $ws $axis $to ($to + $n - 1)
    if ($to -lt $from) { $src = AxisRange $ws $axis ($from + $n) ($from + 2 * $n - 1) } else { $src = AxisRange $ws $axis $from ($from + $n - 1) }
    [void]$src.Cut($blk)
    if ($to -lt $from) { $src = AxisRange $ws $axis ($from + $n) ($from + 2 * $n - 1) } else { $src = AxisRange $ws $axis $from ($from + $n - 1) }
    [void]$src.Delete()
  }
}

# --- unmerge look (types.ts LoggedOp 'unmerge') -----------------------------------------------------------------
# The top-left cell's FILL and its border sides spread over the other cells of each former area (Excel's own UnMerge
# leaves them empty): top sides on the first row, bottom on the last row, left on the first column, right on the last
# column. xlEdgeLeft = 7, xlEdgeTop = 8, xlEdgeBottom = 9, xlEdgeRight = 10; xlNone = -4142; xlColorIndexAutomatic = -4105.
function Get-MergeLook($tl) {
  $look = @{ fill = $null; sides = @{} }
  $i = $tl.Interior
  if ($i.Pattern -ne -4142) {
    $f = @{ pattern = $i.Pattern; color = $i.Color; theme = 0; tint = 0.0 }
    try { $t = $i.ThemeColor; if ($t) { $f.theme = [int]$t; $f.tint = [double]$i.TintAndShade } } catch {}
    $look.fill = $f
  }
  foreach ($e in 7, 8, 9, 10) {
    $b = $tl.Borders.Item($e)
    if ($b.LineStyle -ne -4142) {
      $s = @{ lineStyle = $b.LineStyle; weight = $b.Weight; color = $b.Color; colorIndex = $b.ColorIndex; theme = 0; tint = 0.0 }
      try { $t = $b.ThemeColor; if ($t) { $s.theme = [int]$t; $s.tint = [double]$b.TintAndShade } } catch {}
      $look.sides[$e] = $s
    }
  }
  return $look
}
function Set-MergeLook($c, $look, $edges) {
  if ($look.fill) {
    $f = $look.fill
    $c.Interior.Pattern = $f.pattern
    if ($f.theme) { $c.Interior.ThemeColor = $f.theme; $c.Interior.TintAndShade = $f.tint } else { $c.Interior.Color = $f.color }
  }
  foreach ($e in $edges) {
    $s = $look.sides[$e]
    if (-not $s) { continue }
    $b = $c.Borders.Item($e)
    $b.LineStyle = $s.lineStyle
    $b.Weight = $s.weight
    if ($s.colorIndex -eq -4105) { $b.ColorIndex = -4105 }
    elseif ($s.theme) { $b.ThemeColor = $s.theme; $b.TintAndShade = $s.tint }
    else { $b.Color = $s.color }
  }
}

function Bval($v) { if ($null -eq $v -or $v -is [DBNull]) { return 'null' } else { return [string]$v } }

$xl = $null
try {
  $xl = New-GExcel
  if ($NormalizeRaw) {
    # the hand-built fixture as Excel writes it (fonts with family / charset / scheme, dimensions ...)
    $wbn = $xl.Workbooks.Open((Resolve-Path $NormalizeRaw).Path, 0, $true)
    $np = Join-Path $FixtureDir $NormalizedName
    if (Test-Path $np) { Remove-Item $np -Force }
    $wbn.SaveAs($np, 51)
    $wbn.Close($false)
    Write-Host ('{"normalized":' + (JStr $NormalizedName) + '}')
  }
  if ($MergeProbe) {
    $pd = (Resolve-Path $ProbeDir).Path
    $wbp = $xl.Workbooks.Open((Join-Path $FixtureDir $NormalizedName), 0, $true)
    $wsp = $wbp.Worksheets.Item('Band')
    # borders Excel reports for the merged area (MergeArea) next to those of its top-left cell
    foreach ($a in 'A1:F1', 'A3:C5', 'A11:C11') {
      $tl = $wsp.Range($a).Cells.Item(1, 1)
      $ma = $tl.MergeArea
      $o = @()
      foreach ($e in 7, 8, 9, 10) { $o += ('"e' + $e + '":[' + (JStr (Bval ($tl.Borders.Item($e).LineStyle))) + ',' + (JStr (Bval ($ma.Borders.Item($e).LineStyle))) + ']') }
      Write-Host ('{"mergearea":' + (JStr $a) + ',' + ($o -join ',') + '}')
    }
    # which COM write moves the top-left cell's right / bottom border to its neighbour? (A1:F1: fill on B1 only;
    # A3:C5: top border on B3 only; A11:C11: bottom border on B11 only)
    [void]$wsp.Range('A1:F1').UnMerge(); $wsp.Range('B1').Interior.Color = 65535
    [void]$wsp.Range('A3:C5').UnMerge(); $pb = $wsp.Range('B3').Borders.Item(8); $pb.LineStyle = 1; $pb.Weight = 2
    [void]$wsp.Range('A11:C11').UnMerge(); $pb = $wsp.Range('B11').Borders.Item(9); $pb.LineStyle = 1; $pb.Weight = 2
    $pp = Join-Path $pd 'probe-r3-which.xlsx'
    if (Test-Path $pp) { Remove-Item $pp -Force }
    $wbp.SaveAs($pp, 51)
    $wbp.Close($false)
    Write-Host '{"mergeprobe":"done"}'
  }
  foreach ($c in $list) {
    if ($Only -and -not ($Only -contains $c.name)) { continue }
    $err = $null
    $wb = $null
    try {
      $src = Join-Path $FixtureDir $c.fixture
      $wb = $xl.Workbooks.Open($src, 0, $true)
      foreach ($lop in $c.ops) {
        $ws = $wb.Worksheets.Item([string]$lop.sheet)
        if ($lop.kind -eq 'struct') { Invoke-Struct $ws $lop.op }
        elseif ($lop.kind -eq 'colWidth') {
          # LoggedOp width = the <col width> FILE value; ColumnWidth excludes the 5 px padding (0.625 with the
          # default 11pt font, max digit width 8 px)
          $ws.Columns.Item([int]$lop.col).ColumnWidth = [double]$lop.width - 0.625
        }
        elseif ($lop.kind -eq 'cells') {
          foreach ($w in $lop.writes) {
            $cell = $ws.Cells.Item([int]$w.row, [int]$w.col)
            if ($lop.reason -eq 'toValues') {
              # an array / spill master: the whole range becomes its values (like Copy + Paste Values)
              $arr = $null
              if ($cell.HasArray) { $arr = $cell.CurrentArray }
              elseif ($cell.HasSpill -and $cell.SpillingToRange -ne $null -and $cell.SpillingToRange.Cells.Item(1, 1).Address() -eq $cell.Address()) { $arr = $cell.SpillingToRange }
              if ($arr) { $vals = $arr.Value2; $arr.Value2 = $vals; continue }
            }
            if ($null -eq $w.cell) { [void]$cell.ClearContents(); continue }
            # types.ts CellWrite: the text format '@' is applied BEFORE the value, every other format after it
            if ($w.cell.numFmt -eq '@') { $cell.NumberFormat = '@' }
            if ($w.cell.f) {
              # FILE form -> UI form for the few prefixes the cases use
              $uf = [regex]::Replace([string]$w.cell.f, '_xlfn\.ANCHORARRAY\(([A-Za-z0-9!$]+)\)', '$1#')
              $uf = $uf -replace '_xlfn\._xlws\.', '' -replace '_xlfn\.', '' -replace '_xlws\.', ''
              if ($w.cell.dyn) { $cell.Formula2 = '=' + $uf } else { $cell.Formula = '=' + $uf }
            }
            elseif ($w.cell.t -eq 's') {
              # quote === true: the apostrophe prefix (quotePrefix); anything else: the text as is
              if ($w.cell.quote -eq $true) { $cell.Formula = "'" + [string]$w.cell.v }
              else {
                # NOTE: through COM a prefixed cell KEEPS its quotePrefix flag when a value is written (Value2 / Formula /
                # ClearContents; PrefixCharacter reads '' for a number / empty cell but the flag stays in the xf and acts
                # again when text is written; only ClearFormats / Clear drop it, with every other format). The UI resets
                # it on typing and so does the contract (types.ts CellWrite.cell.quote): goldens q-text / q-overwrite
                # list this as a known deviation (tests/golden/cases.ts)
                $cell.Value2 = [string]$w.cell.v
              }
            }
            elseif ($w.cell.t -eq 'b') { $cell.Formula = $(if ($w.cell.v) { 'TRUE' } else { 'FALSE' }) }
            else { SetF $cell $w.cell.v }
            if ($w.cell.numFmt -and $w.cell.numFmt -ne '@') { $cell.NumberFormat = [string]$w.cell.numFmt }
          }
        }
        elseif ($lop.kind -eq 'unmerge') {
          foreach ($ref in $lop.refs) {
            $rg = $ws.Range([string]$ref)
            $r1 = [int]$rg.Row; $c1 = [int]$rg.Column; $nr = [int]$rg.Rows.Count; $nc = [int]$rg.Columns.Count
            [void]$rg.UnMerge()
            # the top-left cell keeps its look through UnMerge; read it, spread it over the rest of the area
            $look = Get-MergeLook $ws.Cells.Item($r1, $c1)
            if ($look.fill -or $look.sides.Count) {
              for ($dr = 0; $dr -lt $nr; $dr++) {
                for ($dc = 0; $dc -lt $nc; $dc++) {
                  if ($dr -eq 0 -and $dc -eq 0) { continue }
                  $edges = @()
                  if ($dr -eq 0) { $edges += 8 }
                  if ($dr -eq $nr - 1) { $edges += 9 }
                  if ($dc -eq 0) { $edges += 7 }
                  if ($dc -eq $nc - 1) { $edges += 10 }
                  Set-MergeLook $ws.Cells.Item($r1 + $dr, $c1 + $dc) $look $edges
                }
              }
            }
            if ($lop.centerAcross) { $rg.HorizontalAlignment = 7 }   # xlCenterAcrossSelection
          }
        }
      }
      $out = Join-Path $OutDir ($c.name + '.xlsx')
      if (Test-Path $out) { Remove-Item $out -Force }
      $wb.SaveAs($out, 51)
    } catch {
      $err = $_.Exception.Message
    } finally {
      if ($wb) { try { $wb.Close($false) } catch {} }
    }
    $ok = if ($err) { 'false' } else { 'true' }
    Write-Host ('{"name":' + (JStr $c.name) + ',"ok":' + $ok + ',"error":' + (JStr $err) + '}')
  }

  if ($OpenDir) {
    # the files of OpenDir (our outputs) open without repair?
    $M = $script:GMissing
    foreach ($f in @(Get-ChildItem -Path $OpenDir -Filter '*.xlsx' | Sort-Object Name)) {
      $t0 = Get-Date
      $mode = 'failed'; $wbo = $null; $log = ''
      try { $wbo = $xl.Workbooks.Open($f.FullName, 0, $true); $mode = 'normal' }
      catch {
        try { $wbo = $xl.Workbooks.Open($f.FullName, 0, $true, 5, '', '', $true, 2, '', $false, $false, $M, $false, $false, 1); $mode = 'repair' } catch { $mode = 'failed'; $log = $_.Exception.Message }
      }
      Start-Sleep -Milliseconds 100
      foreach ($l in @(Get-ChildItem $env:TEMP -Filter 'error*.xml' -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime -ge $t0.AddSeconds(-1) })) {
        $x = [IO.File]::ReadAllText($l.FullName)
        $log += ([regex]::Matches($x, '<(removed|repaired)[A-Za-z]*>[^<]*') | ForEach-Object { $_.Value }) -join ' ; '
      }
      if ($wbo) { try { $wbo.Close($false) } catch {} }
      Write-Host ('{"file":' + (JStr $f.Name) + ',"open":"' + $mode + '","log":' + (JStr $log) + '}')
    }
  }
}
finally { Close-GExcel $xl }
