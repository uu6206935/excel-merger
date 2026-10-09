# Re-runnable driver of the consistency oracle (task V1).
#   powershell -NoProfile -ExecutionPolicy Bypass -File tests\oracle\consistency\run.ps1 [-Count 320] [-Out <dir>] [-SeedBase 1] [-Only <regex>] [-NoRunner] [-NoExcel] [-Reuse]
# Needs: Node and Excel (own hidden instances only: one per batch of ~40 workbooks for the recalculation, one per scenario for the app's
# Excel engine). Every Excel start writes ~80 MB of Office logs: the oracle gives each its own TEMP below -Out and deletes it afterwards.
# -Reuse re-judges the scenarios / dumps of an earlier run in -Out (no new Excel start unless something is missing).
param([int]$Count = 320, [string]$Out = '', [int]$SeedBase = 1, [string]$Only = '', [switch]$NoRunner, [switch]$NoExcel, [switch]$Reuse)
$ErrorActionPreference = 'Stop'
$repo = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..')).Path
$tmp = Join-Path (Split-Path $repo -Parent) '.tmp'
if (-not (Test-Path $tmp)) { New-Item -ItemType Directory -Path $tmp | Out-Null }
$env:TEMP = $tmp; $env:TMP = $tmp   # C: has little space
if (-not $Out) { $Out = Join-Path $tmp 'v1\oracle' }
New-Item -ItemType Directory -Path $Out -Force | Out-Null
Push-Location $repo
try {
  $a = @('tests/oracle/consistency/run.mjs', '--out', $Out, '--count', "$Count", '--seed-base', "$SeedBase")
  if ($Only) { $a += @('--only', $Only) }
  if ($NoRunner) { $a += '--no-runner' }
  if ($NoExcel) { $a += '--no-excel' }
  if ($Reuse) { $a += '--reuse' }
  & node @a
  exit $LASTEXITCODE
} finally { Pop-Location }
