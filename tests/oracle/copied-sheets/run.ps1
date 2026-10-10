# Re-runnable driver of the copied-sheets oracle (task V9).
#   powershell -NoProfile -ExecutionPolicy Bypass -File tests\oracle\copied-sheets\run.ps1 [-Out <dir>] [-Only <regex>] [-Limit N] [-CheckOutput N] [-SeedBase 0] [-NoRunner] [-NoExcel] [-Reuse]
# -CheckOutput N additionally runs the QA tools (tools/lib/verdict.mjs with --single-sheet / --with-copies --source ..., and the two engines' compare) over N
# scenarios of every group in ONE more Excel instance. -Negative plants one fault at a time into the premerged packages of four scenarios (negative.ts) and
# requires the judge to find every one of them.
# Needs: Node and Excel (own hidden instances only: one per batch of ~24 workbooks for the recalculation and the dumps, one per scenario for
# the app's Excel engine). Every Excel start writes ~80 MB of Office logs: the oracle gives each its own TEMP below -Out and deletes it afterwards.
# -Reuse re-judges the scenarios / dumps of an earlier run in -Out (no new Excel start unless something is missing).
param([string]$Out = '', [string]$Only = '', [int]$Limit = 0, [int]$CheckOutput = 0, [int]$SeedBase = 0, [switch]$Negative, [switch]$NoRunner, [switch]$NoExcel, [switch]$Reuse)
$ErrorActionPreference = 'Stop'
$repo = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..')).Path
$tmp = Join-Path (Split-Path $repo -Parent) '.tmp'
if (-not (Test-Path $tmp)) { New-Item -ItemType Directory -Path $tmp | Out-Null }
$env:TEMP = $tmp; $env:TMP = $tmp   # C: has little space
if (-not $Out) { $Out = Join-Path $tmp 'v9\oracle' }
New-Item -ItemType Directory -Path $Out -Force | Out-Null
Push-Location $repo
try {
  $a = @('tests/oracle/copied-sheets/run.mjs', '--out', $Out, '--seed-base', "$SeedBase")
  if ($Only) { $a += @('--only', $Only) }
  if ($Limit -gt 0) { $a += @('--limit', "$Limit") }
  if ($CheckOutput -gt 0) { $a += @('--check-output', "$CheckOutput") }
  if ($Negative) { $a += '--negative' }
  if ($NoRunner) { $a += '--no-runner' }
  if ($NoExcel) { $a += '--no-excel' }
  if ($Reuse) { $a += '--reuse' }
  & node @a
  exit $LASTEXITCODE
} finally { Pop-Location }
