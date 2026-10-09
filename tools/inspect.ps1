# Opens workbooks in OUR OWN hidden Excel instance and dumps what Excel sees as JSON (owner: T).
# Based on research/r1/inspect.ps1 + opentest.ps1 (R1 section 2: repair detection via CorruptLoad retry).
#
#   one file : powershell -NoProfile -ExecutionPolicy Bypass -File tools\inspect.ps1 -Path <in.xlsx> -Out <out.json> [-Recalc]
#   many     : ... -File tools\inspect.ps1 -JobJson <job.json>
#              job.json = {"files":[{"in":"...","out":"..."}], "recalc":true, "maxCells":300000,
#                          "probeCells":["Sheet!A1", ...]}   (optional: cellProbes = numFmt/text/value/formula/align/merge/spill)
#   (Node wrapper: tools/lib/inspect.mjs; CLI: node tools/inspect.mjs <file>)
#
# Every input is first copied to an ASCII temp file with a unique name (Japanese paths + repair-log matching).
# openMode: 'normal' | 'repair' (opened only with CorruptLoad=xlRepairFile) | 'extract' (xlExtractData) | 'failed'.
# The JSON file is ASCII (non-ASCII escaped); stdout = ASCII JSON lines.
param(
  [string]$Path,
  [string]$Out,
  [string]$JobJson,
  [switch]$Recalc,
  [int]$MaxCells = 300000
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'lib\xlcom.ps1')

Add-Type -TypeDefinition @"
using System;
using System.Text;
using System.Globalization;
using System.Collections.Generic;

public static class EmInspect {
  public static void Esc(StringBuilder sb, string s) {
    sb.Append('"');
    foreach (char ch in s) {
      if (ch == '"') sb.Append("\\\"");
      else if (ch == '\\') sb.Append("\\\\");
      else if (ch < 32 || ch > 126) sb.Append("\\u").Append(((int)ch).ToString("x4"));
      else sb.Append(ch);
    }
    sb.Append('"');
  }
  public static string Str(string s) {
    if (s == null) return "null";
    StringBuilder sb = new StringBuilder(s.Length + 2); Esc(sb, s); return sb.ToString();
  }
  public static string Col(int c) {
    string s = "";
    while (c > 0) { int m = (c - 1) % 26; s = (char)('A' + m) + s; c = (c - 1) / 26; }
    return s;
  }
  public static string ErrName(int v) {
    int code = v + 2146828288;
    switch (code) {
      case 2000: return "#NULL!"; case 2007: return "#DIV/0!"; case 2015: return "#VALUE!";
      case 2023: return "#REF!"; case 2029: return "#NAME?"; case 2036: return "#NUM!";
      case 2042: return "#N/A"; case 2043: return "#GETTING_DATA"; case 2045: return "#SPILL!";
      case 2046: return "#CONNECT!"; case 2047: return "#BLOCKED!"; case 2048: return "#UNKNOWN!";
      case 2049: return "#FIELD!"; case 2050: return "#CALC!";
      default: return "#ERR" + code.ToString(CultureInfo.InvariantCulture);
    }
  }
  public static string Num(double d) {
    if (double.IsNaN(d) || double.IsInfinity(d)) return "null";
    return d.ToString("R", CultureInfo.InvariantCulture);
  }

  public class Block {
    public StringBuilder Formulas = new StringBuilder();
    public StringBuilder Values = new StringBuilder();
    public StringBuilder Errors = new StringBuilder();
    public int FormulaCount;
    public int ValueCount;
    public int RefTextCount;
    public Dictionary<string, int> ErrorCounts = new Dictionary<string, int>();
    public List<string> FormulaCells = new List<string>();
    public int[] RowTextCounts;
  }

  static object At(object arr, object scalar, int i, int j) {
    object[,] a = arr as object[,];
    if (a != null) return a[a.GetLowerBound(0) + i, a.GetLowerBound(1) + j];
    return (i == 0 && j == 0) ? scalar : null;
  }

  // values = Range.Value2, formulas = Range.Formula2 (object[,] with 1-based bounds, or a scalar for one cell)
  public static Block Dump(object values, object formulas, int row0, int col0, int nRows, int nCols) {
    Block b = new Block();
    b.RowTextCounts = new int[Math.Min(nRows, 40)];
    bool fF = true, fV = true, fE = true;
    b.Formulas.Append('{'); b.Values.Append('{'); b.Errors.Append('{');
    for (int i = 0; i < nRows; i++) {
      string rowText = (row0 + i).ToString(CultureInfo.InvariantCulture);
      for (int j = 0; j < nCols; j++) {
        object v = At(values, values, i, j);
        object f = At(formulas, formulas, i, j);
        string addr = null;
        string fs = f as string;
        if (fs != null && fs.Length > 1 && fs[0] == '=') {
          addr = Col(col0 + j) + rowText;
          if (!fF) b.Formulas.Append(','); fF = false;
          Esc(b.Formulas, addr); b.Formulas.Append(':'); Esc(b.Formulas, fs);
          b.FormulaCount++;
          if (fs.IndexOf("#REF!", StringComparison.Ordinal) >= 0) b.RefTextCount++;
          b.FormulaCells.Add(addr);
        }
        if (v == null || v is DBNull) continue;
        if (addr == null) addr = Col(col0 + j) + rowText;
        if (v is int) {
          string e = ErrName((int)v);
          if (!fE) b.Errors.Append(','); fE = false;
          Esc(b.Errors, addr); b.Errors.Append(':'); Esc(b.Errors, e);
          int n; b.ErrorCounts.TryGetValue(e, out n); b.ErrorCounts[e] = n + 1;
          continue;
        }
        string json;
        if (v is double) json = Num((double)v);
        else if (v is bool) json = ((bool)v) ? "true" : "false";
        else if (v is string) {
          string s = (string)v;
          if (s.Length == 0) continue;
          json = Str(s);
          if (i < b.RowTextCounts.Length && s.Trim().Length > 0) b.RowTextCounts[i]++;
        }
        else if (v is decimal) json = Num((double)(decimal)v);
        else if (v is DateTime) json = Num(((DateTime)v).ToOADate());
        else json = Str(Convert.ToString(v, CultureInfo.InvariantCulture));
        if (!fV) b.Values.Append(','); fV = false;
        Esc(b.Values, addr); b.Values.Append(':'); b.Values.Append(json);
        b.ValueCount++;
      }
    }
    b.Formulas.Append('}'); b.Values.Append('}'); b.Errors.Append('}');
    return b;
  }
}
"@ -ErrorAction SilentlyContinue

$M = [System.Reflection.Missing]::Value

# ---- tiny JSON writer (ASCII) -------------------------------------------------------------------------------------
function J($v) {
  if ($null -eq $v -or $v -is [DBNull]) { return 'null' }
  if ($v -is [System.Text.StringBuilder]) { return $v.ToString() }       # raw JSON fragment
  if ($v -is [bool]) { if ($v) { return 'true' } else { return 'false' } }
  if ($v -is [string]) { return [EmInspect]::Str($v) }
  if ($v -is [int] -or $v -is [long] -or $v -is [double] -or $v -is [single] -or $v -is [decimal] -or $v -is [uint32] -or $v -is [int16] -or $v -is [byte]) {
    return [EmInspect]::Num([double]$v)
  }
  if ($v -is [System.Collections.IDictionary]) {
    $parts = New-Object System.Collections.Generic.List[string]
    foreach ($k in $v.Keys) { $parts.Add([EmInspect]::Str([string]$k) + ':' + (J $v[$k])) }
    return '{' + ($parts -join ',') + '}'
  }
  if ($v -is [System.Collections.IEnumerable]) {
    $parts = New-Object System.Collections.Generic.List[string]
    foreach ($x in $v) { $parts.Add((J $x)) }
    return '[' + ($parts -join ',') + ']'
  }
  return [EmInspect]::Str([string]$v)
}

# ---- "Excel is busy" handling ---------------------------------------------------------------------------------------
# A COM call into an Excel that is loading / calculating / competing with another instance is REJECTED (RPC_E_CALL_REJECTED,
# RPC_E_SERVERCALL_RETRYLATER, VBA_E_IGNORE). A swallowed rejection used to read as "nothing there" (no names, no CF / DV / chart
# series ...) and turned a good file into a FAIL now and then. Such calls are now retried with growing pauses; one that never
# gets through is COUNTED (transientFailures in the dump) so that tools/lib/dump-check.mjs can never mistake it for a fact.
$script:TransientFailures = 0

function Get-EmHResults($ex) {
  $list = New-Object System.Collections.Generic.List[int]
  $e = $ex
  for ($i = 0; $i -lt 6 -and $null -ne $e; $i++) {
    try { $list.Add([int]$e.HResult) } catch {}
    $e = $e.InnerException
  }
  return , $list
}

function Test-EmTransient($errRecord) {
  $transient = @(-2147418111, -2147417846, -2146777998)    # 0x80010001, 0x8001010A, 0x800AC472
  foreach ($h in (Get-EmHResults $errRecord.Exception)) { if ($transient -contains $h) { return $true } }
  return ([string]$errRecord.Exception.Message) -match '0x80010001|0x8001010A|0x800AC472|RPC_E_CALL_REJECTED|SERVERCALL_RETRYLATER'
}

# Runs $sb and returns its value (`, $v` keeps COM objects from being unrolled). Transient errors are retried; anything else,
# and a transient error that survives every attempt (counted), is rethrown.
function Invoke-EmRetry([scriptblock]$sb, [int]$attempts = 8) {
  for ($i = 1; $i -le $attempts; $i++) {
    try {
      $v = & $sb
      if ($null -eq $v -or $v -is [DBNull]) { return $null }
      return , $v
    } catch {
      if (Test-EmTransient $_) {
        if ($i -lt $attempts) { Start-Sleep -Milliseconds (150 * $i); continue }
        $script:TransientFailures++
      }
      throw
    }
  }
}

# Evaluate a scriptblock, $null on error (transient "busy" errors are retried first). `return , $v` keeps COM collections /
# Ranges from being unrolled into arrays by the pipeline (FormatConditions.Item() would then hit the 0-based array indexer).
function TryS([scriptblock]$sb) {
  try { $v = Invoke-EmRetry $sb } catch { return $null }
  if ($null -eq $v -or $v -is [DBNull]) { return $null }
  return , $v
}

# Waits until Excel reports Ready and finished calculating (Application.Ready, CalculationState = xlDone) twice in a row.
# Returns $false when it never settled within $timeoutMs (the dump is then flagged notReady).
function Wait-EmExcelReady($xl, [int]$timeoutMs = 60000) {
  $t0 = [DateTime]::UtcNow
  $stable = 0
  while (([DateTime]::UtcNow - $t0).TotalMilliseconds -lt $timeoutMs) {
    $ok = $false
    try { $ok = ([bool]$xl.Ready) -and ([int]$xl.CalculationState -eq 0) } catch { $ok = $false }
    if ($ok) { $stable++; if ($stable -ge 2) { return $true } } else { $stable = 0 }
    Start-Sleep -Milliseconds 120
  }
  return $false
}

function Addr($rng) { if ($null -eq $rng) { return $null }; return [string]$rng.Address($false, $false) }

# ---- open with repair detection (R1 section 2) --------------------------------------------------------------------
function Open-Checked($xl, [string]$file, $res) {
  $t0 = (Get-Date).AddSeconds(-2)
  $leaf = [IO.Path]::GetFileName($file)
  $wb = $null
  $modes = @(@('normal', 0), @('repair', 1), @('extract', 2))
  $fmt = 5
  if ($leaf -match '[.]csv$') { $fmt = 2 }   # text files: comma delimited
  # HRESULTs that mean "Excel is busy, try again" (RPC_E_CALL_REJECTED, RPC_E_SERVERCALL_RETRYLATER, VBA_E_IGNORE):
  # they must not be mistaken for "the file needs repair" (Test-EmTransient looks through the wrapped exceptions too)
  foreach ($md in $modes) {
    for ($attempt = 1; $attempt -le 8; $attempt++) {
      try {
        # pass real values for every positional argument (with Missing, the CorruptLoad open fails too - R1)
        $wb = $xl.Workbooks.Open($file, 0, $true, $fmt, '', '', $true, 2, '', $false, $false, $M, $false, $false, $md[1])
        $res['openMode'] = $md[0]
        break
      } catch {
        $wb = $null
        if ((Test-EmTransient $_) -and $attempt -lt 8) { Start-Sleep -Milliseconds 500; continue }
        if (-not $res.Contains('openError')) { $res['openError'] = [string]$_.Exception.Message }
        break
      }
    }
    if ($null -ne $wb) { break }
  }
  if ($null -eq $wb) { $res['openMode'] = 'failed' }
  Start-Sleep -Milliseconds 300
  $logs = New-Object System.Collections.Generic.List[string]
  foreach ($l in @(Get-ChildItem $env:TEMP -Filter 'error*.xml' -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime -ge $t0 })) {
    try {
      $x = [IO.File]::ReadAllText($l.FullName)
      if ($x.IndexOf($leaf, [StringComparison]::OrdinalIgnoreCase) -lt 0) { continue }   # another process' log
      $items = @([regex]::Matches($x, '<(removedRecord|repairedRecord|removedPart|repairedPart|removedFeature|repairedFeature)(?:\s[^>]*)?>([^<]+)') | ForEach-Object { $_.Groups[1].Value + ': ' + $_.Groups[2].Value })
      if ($items.Count -eq 0) { $items = @('log: ' + $l.Name) }
      foreach ($it in $items) { $logs.Add([string]$it) }
    } catch {}
  }
  $res['repairLog'] = $logs
  # a log that mentions our file means Excel repaired something even when the plain open "succeeded"
  if ($res['openMode'] -eq 'normal' -and $logs.Count -gt 0) { $res['openMode'] = 'repair' }
  return $wb
}

# ---- one sheet ------------------------------------------------------------------------------------------------------
function Inspect-Sheet($xl, $wb, $ws, [int]$index, [int]$maxCells) {
  $s = [ordered]@{}
  $s['name'] = [string]$ws.Name
  $s['index'] = $index
  $s['visible'] = [int]$ws.Visible
  $ur = $ws.UsedRange
  $row0 = [int]$ur.Row; $col0 = [int]$ur.Column
  $nRows = [int]$ur.Rows.Count; $nCols = [int]$ur.Columns.Count
  $s['usedRange'] = Addr $ur
  $s['maxRow'] = $row0 + $nRows - 1
  $s['maxCol'] = $col0 + $nCols - 1
  $dumpRows = $nRows
  if ([double]$nRows * $nCols -gt $maxCells) { $dumpRows = [int][math]::Max(1, [math]::Floor($maxCells / $nCols)); $s['capped'] = $true } else { $s['capped'] = $false }
  $rng = $ws.Range($ws.Cells.Item($row0, $col0), $ws.Cells.Item($row0 + $dumpRows - 1, $col0 + $nCols - 1))
  $v2 = $rng.Value2
  $f2 = $null
  try { $f2 = $rng.Formula2 } catch { $f2 = $rng.Formula }
  $blk = [EmInspect]::Dump($v2, $f2, $row0, $col0, $dumpRows, $nCols)
  $s['formulaCount'] = $blk.FormulaCount
  $s['refTextCount'] = $blk.RefTextCount
  $ec = [ordered]@{}; foreach ($k in $blk.ErrorCounts.Keys) { $ec[$k] = $blk.ErrorCounts[$k] }
  $s['errorCounts'] = $ec
  $rt = @(); for ($i = 0; $i -lt $blk.RowTextCounts.Length; $i++) { $rt += $blk.RowTextCounts[$i] }
  $s['rowTextCounts'] = $rt       # text cells per row for the first 40 rows (header-row guess)
  $s['firstRow'] = $row0
  $s['firstCol'] = $col0
  $s['formulas'] = $blk.Formulas
  $s['values'] = $blk.Values
  $s['errors'] = $blk.Errors

  # style probes: row 1, the guessed header row and the row below it (first 40 columns)
  $best = -1; $bestN = 0
  for ($i = 0; $i -lt $blk.RowTextCounts.Length; $i++) { if ($blk.RowTextCounts[$i] -gt $bestN) { $bestN = $blk.RowTextCounts[$i]; $best = $i } }
  $probeRows = @(1)
  if ($best -ge 0) { $hg = $row0 + $best; $probeRows += $hg; $probeRows += ($hg + 1) }
  $probes = [ordered]@{}
  foreach ($r in ($probeRows | Select-Object -Unique)) {
    if ($r -gt $s['maxRow']) { continue }
    for ($c = $col0; $c -le [math]::Min($s['maxCol'], $col0 + 39); $c++) {
      $cell = $ws.Cells.Item($r, $c)
      try {
        $fill = 'none'
        if ([int]$cell.Interior.ColorIndex -ne -4142) { $fill = [string][int]$cell.Interior.Color }
        $probes[[EmInspect]::Col($c) + $r] = ('{0}|b{1}|fc{2}|bg{3}|h{4}|sz{5}' -f [string]$cell.NumberFormat, [int][bool]$cell.Font.Bold, [int]$cell.Font.Color, $fill, [int]$cell.HorizontalAlignment, [double]$cell.Font.Size)
      } catch {}
    }
  }
  $s['probes'] = $probes

  # array / spill formulas (per formula cell, only when the sheet has any)
  $arrays = New-Object System.Collections.Generic.List[string]
  $spills = New-Object System.Collections.Generic.List[string]
  $hasArr = TryS { $ur.HasArray }
  $hasSpill = TryS { $ur.HasSpill }
  if (($hasArr -ne $false -or $hasSpill -ne $false) -and $blk.FormulaCells.Count -gt 0) {
    $n = 0
    foreach ($a in $blk.FormulaCells) {
      if (++$n -gt 5000) { $s['arraysCapped'] = $true; break }
      $c = $ws.Range($a)
      if ($hasArr -ne $false -and $c.HasArray) { $ca = Addr $c.CurrentArray; if ($ca.StartsWith($a)) { $arrays.Add("$a=>$ca") } }
      if ($hasSpill -ne $false) { try { if ($c.HasSpill -and (Addr $c.SpillParent) -eq $a) { $spills.Add("$a=>" + (Addr $c.SpillingToRange)) } } catch {} }
    }
  }
  $s['arrays'] = $arrays
  $s['spills'] = $spills

  # merged areas
  $merges = New-Object System.Collections.Generic.List[string]
  $mc = TryS { $ur.MergeCells }
  if ($mc -ne $false) {
    $lastRow = [math]::Min($row0 + $nRows - 1, $row0 + 1999)
    for ($r = $row0; $r -le $lastRow; $r++) {
      $rowRng = $ws.Range($ws.Cells.Item($r, $col0), $ws.Cells.Item($r, $col0 + $nCols - 1))
      if ((TryS { $rowRng.MergeCells }) -eq $false) { continue }
      for ($c = $col0; $c -le $col0 + $nCols - 1; $c++) {
        $cell = $ws.Cells.Item($r, $c)
        if ($cell.MergeCells) { $ma = $cell.MergeArea; if ([int]$ma.Row -eq $r -and [int]$ma.Column -eq $c) { $merges.Add((Addr $ma)) } }
      }
    }
  }
  $s['merges'] = $merges

  # columns: width / hidden
  $cols = New-Object System.Collections.Generic.List[object]
  $colMax = [math]::Min($s['maxCol'] + 3, 100)
  for ($c = 1; $c -le $colMax; $c++) {
    $co = $ws.Columns.Item($c)
    $cols.Add([ordered]@{ col = $c; letter = [EmInspect]::Col($c); width = [math]::Round([double]$co.ColumnWidth, 2); hidden = [bool]$co.Hidden })
  }
  $s['cols'] = $cols
  $s['hiddenColumns'] = @($cols | Where-Object { $_.hidden } | ForEach-Object { $_.letter })

  # sheet view: needs the sheet to be active
  $ws.Activate()
  $w = $xl.ActiveWindow
  $s['freeze'] = [ordered]@{ frozen = [bool]$w.FreezePanes; splitRow = [int]$w.SplitRow; splitCol = [int]$w.SplitColumn }
  $s['autoFilter'] = TryS { if ($ws.AutoFilterMode) { Addr $ws.AutoFilter.Range } else { $null } }

  # conditional formats (Formula1 is relative to the active cell -> select the top-left cell of AppliesTo first)
  $cfs = New-Object System.Collections.Generic.List[object]
  # (the leading comma keeps the COM collection from being unrolled into an array by the pipeline)
  $fcs = TryS { , $ws.Cells.FormatConditions }
  if ($null -ne $fcs) {
    for ($i = 1; $i -le $fcs.Count; $i++) {
      $fc = $fcs.Item($i)
      $at = $null; try { $at = $fc.AppliesTo } catch {}
      if ($null -ne $at) { try { [void]$at.Cells.Item(1, 1).Select() } catch {} }
      $cfs.Add([ordered]@{
          type = TryS { [int]$fc.Type }
          appliesTo = TryS { Addr $fc.AppliesTo }
          operator = TryS { [int]$fc.Operator }
          formula1 = TryS { [string]$fc.Formula1 }
          formula2 = TryS { [string]$fc.Formula2 }
          priority = TryS { [int]$fc.Priority }
        })
    }
  }
  $s['cf'] = $cfs

  # data validation: one entry per column segment of every validated area (consecutive equal columns merged)
  $dvs = New-Object System.Collections.Generic.List[object]
  # xlCellTypeAllValidation: throws "no cells found" when the sheet has no validation (not transient: not retried)
  $vr = $null; try { $vr = Invoke-EmRetry { , $ws.Cells.SpecialCells(-4174) } } catch {}
  if ($null -ne $vr) {
    foreach ($area in $vr.Areas) {
      $prev = $null
      for ($c = 1; $c -le $area.Columns.Count; $c++) {
        $colRng = $area.Columns.Item($c)
        $top = $colRng.Cells.Item(1, 1)
        try { [void]$top.Select() } catch {}
        $v = $top.Validation
        $d = [ordered]@{
          range = Addr $colRng
          type = TryS { [int]$v.Type }
          operator = TryS { [int]$v.Operator }
          formula1 = TryS { [string]$v.Formula1 }
          formula2 = TryS { [string]$v.Formula2 }
          alertStyle = TryS { [int]$v.AlertStyle }
        }
        if ($null -ne $prev -and $prev.type -eq $d.type -and $prev.formula1 -eq $d.formula1 -and $prev.formula2 -eq $d.formula2 -and $prev.operator -eq $d.operator) {
          $prev.range = Addr $ws.Range($ws.Range($prev.range.Split(':')[0]), $colRng)
        } else {
          $dvs.Add($d); $prev = $d
        }
      }
    }
  }
  $s['dv'] = $dvs
  try { [void]$ws.Range('A1').Select() } catch {}

  # tables
  $tables = New-Object System.Collections.Generic.List[object]
  foreach ($lo in $ws.ListObjects) {
    $tcols = @(); foreach ($lc in $lo.ListColumns) { $tcols += [string]$lc.Name }
    $tables.Add([ordered]@{ name = [string]$lo.Name; range = Addr $lo.Range; showTotals = [bool]$lo.ShowTotals; columns = $tcols })
  }
  $s['tables'] = $tables

  # charts
  $charts = New-Object System.Collections.Generic.List[object]
  foreach ($co in $ws.ChartObjects()) {
    $ch = $co.Chart
    $series = New-Object System.Collections.Generic.List[object]
    $sc = $ch.SeriesCollection()
    for ($i = 1; $i -le $sc.Count; $i++) {
      $se = $sc.Item($i)
      $series.Add([ordered]@{ name = TryS { [string]$se.Name }; formula = TryS { [string]$se.Formula } })
    }
    $charts.Add([ordered]@{
        name = [string]$co.Name
        type = TryS { [int]$ch.ChartType }
        title = TryS { if ($ch.HasTitle) { [string]$ch.ChartTitle.Text } else { $null } }
        topLeft = TryS { Addr $co.TopLeftCell }
        series = $series
      })
  }
  $s['charts'] = $charts

  # pivot tables
  $pivots = New-Object System.Collections.Generic.List[object]
  foreach ($pt in $ws.PivotTables()) {
    # RowFields etc. are parameterized properties: call them like methods to get the collections
    $rf = @(); try { foreach ($f in $pt.RowFields()) { $rf += [string]$f.Name } } catch {}
    $cf = @(); try { foreach ($f in $pt.ColumnFields()) { $cf += [string]$f.Name } } catch {}
    $df = @(); try { foreach ($f in $pt.DataFields()) { $df += [string]$f.Name } } catch {}
    $pivots.Add([ordered]@{
        name = [string]$pt.Name
        range = TryS { Addr $pt.TableRange1 }
        source = TryS { [string]$pt.SourceData }
        rowFields = $rf; columnFields = $cf; dataFields = $df
      })
  }
  $s['pivots'] = $pivots

  $s['shapes'] = [int]$ws.Shapes.Count
  $s['comments'] = [int]$ws.Comments.Count
  $s['threadedComments'] = TryS { [int]$ws.CommentsThreaded.Count }
  $s['hyperlinks'] = [int]$ws.Hyperlinks.Count
  return $s
}

# Detailed look at single cells ("Sheet!A1"): number format, displayed text, value, formula, alignment, merge, spill.
function Probe-Cells($wb, $cells) {
  $out = [ordered]@{}
  foreach ($pc in @($cells)) {
    $key = [string]$pc
    $d = [ordered]@{}
    try {
      $bang = $key.LastIndexOf('!')
      if ($bang -lt 1) { throw "expected Sheet!A1: $key" }
      $sn = $key.Substring(0, $bang)
      if ($sn.StartsWith("'") -and $sn.EndsWith("'")) { $sn = $sn.Substring(1, $sn.Length - 2).Replace("''", "'") }
      $c = $wb.Worksheets.Item($sn).Range($key.Substring($bang + 1)).Cells.Item(1, 1)
      $d['numFmt'] = [string]$c.NumberFormat
      $d['numFmtLocal'] = TryS { [string]$c.NumberFormatLocal }
      $d['text'] = [string]$c.Text
      $v = $c.Value2
      if ($v -is [int]) { $d['error'] = [EmInspect]::ErrName([int]$v); $d['value'] = $null } else { $d['value'] = $v }
      $f2 = TryS { [string]$c.Formula2 }
      $d['formula'] = $(if ($c.HasFormula) { $f2 } else { $null })
      $d['hAlign'] = [int]$c.HorizontalAlignment
      $d['bold'] = [bool]$c.Font.Bold
      $d['fontSize'] = [double]$c.Font.Size
      $d['fontColor'] = TryS { [int]$c.Font.Color }
      # Interior.Pattern -4142 = xlNone -> no fill; otherwise the BGR colour as Excel shows it
      $d['fill'] = TryS { if ([int]$c.Interior.Pattern -eq -4142) { $null } else { [int]$c.Interior.Color } }
      $d['merge'] = $(if ($c.MergeCells) { Addr $c.MergeArea } else { $null })
      $d['hasSpill'] = TryS { [bool]$c.HasSpill }
      $d['spill'] = TryS { if ($c.HasSpill -and (Addr $c.SpillParent) -eq (Addr $c)) { Addr $c.SpillingToRange } else { $null } }
      $d['prefix'] = TryS { [string]$c.PrefixCharacter }
    } catch { $d['probeError'] = [string]$_.Exception.Message }
    $out[$key] = $d
  }
  return $out
}

function Inspect-File($xl, [string]$inPath, [string]$outPath, [bool]$recalc, [int]$maxCells, $probeCells) {
  $res = [ordered]@{ tool = 'excel-merger/inspect'; version = 1; file = [IO.Path]::GetFileName($inPath); path = $inPath }
  if (-not (Test-Path -LiteralPath $inPath)) { throw "file not found: $inPath" }
  $script:TransientFailures = 0
  $tmpDir = New-EmTempDir 'em-insp'
  $ext = [IO.Path]::GetExtension($inPath)
  $tmpFile = Join-Path $tmpDir ('insp-' + [Guid]::NewGuid().ToString('N').Substring(0, 12) + $ext)
  Copy-Item -LiteralPath $inPath -Destination $tmpFile
  $wb = $null
  try {
    # one Excel start per batch, every file opened ONCE and read-only; nothing is read before Excel has settled
    [void](Wait-EmExcelReady $xl 30000)
    $wb = Open-Checked $xl $tmpFile $res
    if ($null -ne $wb) {
      if (-not (Wait-EmExcelReady $xl 60000)) { $res['notReady'] = $true }
      # the workbook Excel opened must be the file we asked for (unique leaf name)
      $opened = TryS { [string]$wb.FullName }
      if ($opened -and -not ([IO.Path]::GetFileName([string]$opened) -ieq [IO.Path]::GetFileName($tmpFile))) { $res['wrongWorkbook'] = [string]$opened }
      if ($recalc) {
        $xl.CalculateFull()
        # CalculateFull returns while multi-threaded calculation may still run: wait for xlDone before anything is read
        if (-not (Wait-EmExcelReady $xl 120000)) { $res['notReady'] = $true }
        $res['recalculated'] = $true
      } else { $res['recalculated'] = $false }
      $res['excelVersion'] = [string]$xl.Version
      $res['date1904'] = [bool]$wb.Date1904
      $res['fileFormat'] = [int]$wb.FileFormat
      # defined names: read until the number of names read equals Names.Count (a half-loaded name table reads short)
      $names = $null
      for ($pass = 1; $pass -le 4; $pass++) {
        $names = New-Object System.Collections.Generic.List[object]
        $expectNames = [int]$wb.Names.Count
        foreach ($n in $wb.Names) {
          # sheet-scoped names come back as  Sheet!Name  /  'My Sheet'!Name
          $nm = [string]$n.Name
          $scope = $null
          $bang = $nm.LastIndexOf('!')
          if ($bang -gt 0) {
            $scope = $nm.Substring(0, $bang)
            if ($scope.StartsWith("'") -and $scope.EndsWith("'")) { $scope = $scope.Substring(1, $scope.Length - 2).Replace("''", "'") }
            $nm = $nm.Substring($bang + 1)
          }
          $names.Add([ordered]@{ name = $nm; scope = $scope; refersTo = TryS { [string]$n.RefersTo }; visible = TryS { [bool]$n.Visible } })
        }
        if ($names.Count -eq $expectNames) { break }
        Start-Sleep -Milliseconds (300 * $pass)
        [void](Wait-EmExcelReady $xl 30000)
      }
      if ($names.Count -ne $expectNames) { $res['nameCountMismatch'] = ('{0}/{1}' -f $names.Count, $expectNames) }
      $res['names'] = $names
      $sheets = New-Object System.Collections.Generic.List[object]
      $i = 0
      foreach ($ws in $wb.Worksheets) {
        $i++
        Write-EmLine ([ordered]@{ type = 'progress'; file = $res.file; sheet = [string]$ws.Name })
        # a sheet is read without side effects on the file, so a rejected call simply reads the sheet again
        $sheets.Add((Invoke-EmRetry { Inspect-Sheet $xl $wb $ws $i $maxCells } 4))
      }
      # a text file's only sheet is named after the (temp) file: report the original name instead
      if ($ext -match '^[.](csv|txt)$' -and $sheets.Count -eq 1) {
        $orig = [IO.Path]::GetFileNameWithoutExtension($inPath)
        if ($orig.Length -gt 31) { $orig = $orig.Substring(0, 31) }
        $sheets[0]['name'] = $orig
      }
      $res['sheets'] = $sheets
      $res['chartSheets'] = @(foreach ($c in $wb.Charts) { [string]$c.Name })
      if ($null -ne $probeCells -and @($probeCells).Count -gt 0) { $res['cellProbes'] = Probe-Cells $wb $probeCells }
      # reads that Excel kept rejecting (busy): tools/lib/dump-check.mjs treats the dump as unreliable and inspects again
      if ($script:TransientFailures -gt 0) { $res['transientFailures'] = $script:TransientFailures }
      # Test hook EM_INSPECT_FAULT_ONCE=<marker file>: while the marker does not exist, the first dump of this process is
      # damaged the way the old flake looked (no defined names, no cell values) and the marker is created, so that the
      # retry in a fresh instance can be tested end to end (tests/qa/dump-check.test.ts).
      if ($env:EM_INSPECT_FAULT_ONCE -and -not (Test-Path -LiteralPath $env:EM_INSPECT_FAULT_ONCE)) {
        [IO.File]::WriteAllText($env:EM_INSPECT_FAULT_ONCE, 'fault injected', (New-Object System.Text.UTF8Encoding($false)))
        $res['names'] = @()
        foreach ($fs in $sheets) { $fs['values'] = New-Object System.Text.StringBuilder '{}' }
      }
    }
  } finally {
    if ($null -ne $wb) { try { $wb.Close($false) } catch {} }
    Remove-Item -LiteralPath $tmpDir -Recurse -Force -ErrorAction SilentlyContinue
  }
  [IO.File]::WriteAllText($outPath, (J $res), (New-Object System.Text.UTF8Encoding($false)))
  Write-EmLine ([ordered]@{ type = 'inspected'; file = $res.file; out = $outPath; openMode = [string]$res.openMode })
}

# ---- main -------------------------------------------------------------------------------------------------------------
$jobs = @()
$doRecalc = [bool]$Recalc
$probe = $null
if ($JobJson) {
  $job = [IO.File]::ReadAllText($JobJson, [Text.Encoding]::UTF8) | ConvertFrom-Json
  foreach ($f in @($job.files)) { $jobs += , @([string]$f.in, [string]$f.out) }
  if ($null -ne $job.recalc) { $doRecalc = [bool]$job.recalc }
  if ($null -ne $job.maxCells) { $MaxCells = [int]$job.maxCells }
  if ($null -ne $job.probeCells) { $probe = @($job.probeCells | ForEach-Object { [string]$_ }) }
} elseif ($Path -and $Out) {
  $jobs += , @($Path, $Out)
} else {
  Write-EmLine ([ordered]@{ type = 'error'; message = 'usage: inspect.ps1 -Path <file> -Out <json> [-Recalc] | -JobJson <job.json>' })
  exit 2
}

$xl = $null
$failed = 0
try {
  $xl = New-EmExcel
  foreach ($jb in $jobs) {
    try { Inspect-File $xl $jb[0] $jb[1] $doRecalc $MaxCells $probe }
    catch {
      $failed++
      Write-EmLine ([ordered]@{ type = 'error'; file = $jb[0]; message = [string]$_.Exception.Message; at = [string]$_.InvocationInfo.PositionMessage })
    }
  }
  Write-EmLine ([ordered]@{ type = 'done'; failed = $failed })
}
finally {
  Close-EmExcel $xl
}
if ($failed -gt 0) { exit 1 }
