#Requires -Version 5.1
<#
  Exercises XlRunnerHelper.cs (the C# half of excel-runner.ps1) WITHOUT Excel: compiles it with Add-Type and runs the cases
  of a JSON file. ASCII-only file. Output: one JSON array (ASCII-escaped, written by XlRunner.Cells.ToJson) in -Out.
    -Cs     path of XlRunnerHelper.cs
    -Cases  JSON [{ kind: 'build', rows, mode?, read? } | { kind: 'localize', fmt, general } | { kind: 'unsafe', f } ]
      build    -> { cells: [[{ kind, val, prefix }]], prefixed, riskyPlain: [flat idx], bad: [flat idx of Verify(read)], neutralized: [flat idx] }
      localize -> { fmt }
      unsafe   -> { unsafe: bool }  (XlRunner.Cells.UnsafeFormula: DDE / external workbook / WEBSERVICE ...)
#>
param([Parameter(Mandatory = $true)][string]$Cs, [Parameter(Mandatory = $true)][string]$Cases, [Parameter(Mandatory = $true)][string]$Out)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Web.Extensions
Add-Type -TypeDefinition ([IO.File]::ReadAllText($Cs, [Text.Encoding]::UTF8))
$js = New-Object System.Web.Script.Serialization.JavaScriptSerializer
$js.MaxJsonLength = [int]::MaxValue
$caseList = $js.DeserializeObject([IO.File]::ReadAllText($Cases, [Text.Encoding]::UTF8))
$results = New-Object System.Collections.ArrayList
foreach ($c in $caseList) {
  if ($c['kind'] -eq 'localize') {
    [void]$results.Add(@{ fmt = [XlRunner.Cells]::Localize([string]$c['fmt'], [string]$c['general']) })
    continue
  }
  if ($c['kind'] -eq 'unsafe') {
    [void]$results.Add(@{ unsafe = [bool][XlRunner.Cells]::UnsafeFormula([string]$c['f']) })
    continue
  }
  $mode = 0
  if ($c.ContainsKey('mode')) { $mode = [int]$c['mode'] }
  $b = [XlRunner.Cells]::Build($c['rows'], $null, $mode)
  $cells = New-Object System.Collections.ArrayList
  for ($r = 0; $r -lt $b.Rows; $r++) {
    $row = New-Object System.Collections.ArrayList
    for ($k = 0; $k -lt $b.Cols; $k++) {
      $kd = [int]$b.Kind[$r, $k]; $vl = $b.Data[$r, $k]; $pf = [bool]$b.Prefix[$r, $k]
      [void]$row.Add(@{ kind = $kd; val = $vl; prefix = $pf })
    }
    [void]$cells.Add($row)
  }
  $res = @{ cells = $cells; prefixed = $b.Prefixed; riskyPlain = @([XlRunner.Cells]::RiskyPlain($b)); bad = @(); neutralized = @([XlRunner.Cells]::NeutralizedCells($b)) }
  if ($c.ContainsKey('read')) {
    $rd = $c['read']
    $m = New-Object 'object[,]' $b.Rows, $b.Cols
    for ($r = 0; $r -lt $b.Rows; $r++) { for ($k = 0; $k -lt $b.Cols; $k++) { $m[$r, $k] = $rd[$r][$k] } }
    $res['bad'] = @([XlRunner.Cells]::Verify($m, $b))
  }
  [void]$results.Add($res)
}
[IO.File]::WriteAllText($Out, [XlRunner.Cells]::ToJson($results), (New-Object System.Text.UTF8Encoding $false))
