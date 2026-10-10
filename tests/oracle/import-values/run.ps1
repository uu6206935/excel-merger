# Re-runnable driver of the import-values oracle (task F1, extended by F2 for includeSourceSheets).
#   powershell -NoProfile -ExecutionPolicy Bypass -File tests\oracle\import-values\run.ps1 [-Corpus <dir>] [-Out <dir>] [-Only <regex>] [-Configs all|quick] [-Regenerate] [-NoExcel] [-Reuse]
#                                                                                         [-Sources] [-Random <n> [-Seed <s>]]
# Needs: Node, a Python 3.11 venv with openpyxl / XlsxWriter / pandas (created under <repo>\..\.tmp\f1\venv when missing) and Excel (own hidden
# instance, see oracle-excel.ps1 / author.ps1). -Regenerate rebuilds the corpus (gen_corpus.py + author.ps1) first.
#   -Sources  every package is built with PremergeOptions.includeSourceSheets (the original sheets of the source are copied; the formulas
#             the merged table cannot keep read the copies): the same comparison with the source's own values (round 9, task F2).
#   -Random n the corpus is replaced by n random source workbooks (random.ts: 3 row orders each, rows missing on both sides, join left /
#             full, half-imported columns; Excel computes the source itself): the property test "a remapped formula evaluated by Excel
#             equals the source value". Implies -Sources. -Seed picks the workbooks (default 1). No Python needed.
param([string]$Corpus = '', [string]$Out = '', [string]$Only = '', [string]$Configs = 'all', [switch]$Regenerate, [switch]$NoExcel, [switch]$Reuse, [switch]$Sources, [int]$Random = 0, [int]$Seed = 1)
$ErrorActionPreference = 'Stop'
$repo = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..')).Path
$tmp = Join-Path (Split-Path $repo -Parent) '.tmp'
if (-not (Test-Path $tmp)) { New-Item -ItemType Directory -Path $tmp | Out-Null }
$env:TEMP = $tmp; $env:TMP = $tmp   # C: has little space (every Excel start writes ~80 MB of Office logs)
if (-not $Corpus) { $Corpus = Join-Path $tmp 'f1\corpus' }
if (-not $Out) { $Out = Join-Path $tmp 'f1\oracle-out' }
if (-not (Test-Path $Out)) { New-Item -ItemType Directory -Path $Out -Force | Out-Null }
if ($Random -le 0 -and ($Regenerate -or -not (Test-Path (Join-Path $Corpus 'manifest.json')) -or -not (Test-Path (Join-Path $Corpus 'cached')))) {
  $venv = Join-Path $tmp 'f1\venv'
  $py = Join-Path $venv 'Scripts\python.exe'
  if (-not (Test-Path $py)) {
    & py -3.11 -m venv $venv
    & $py -m pip install --no-cache-dir openpyxl XlsxWriter pandas
  }
  New-Item -ItemType Directory -Path $Corpus -Force | Out-Null
  $env:PYTHONIOENCODING = 'utf-8'
  & $py (Join-Path $PSScriptRoot 'gen_corpus.py') $Corpus
  & powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'author.ps1') -CorpusDir $Corpus
}
Push-Location $repo
try {
  $a = @('tests/oracle/import-values/run.mjs', '--corpus', $Corpus, '--out', $Out, '--configs', $Configs)
  if ($Only) { $a += @('--only', $Only) }
  if ($NoExcel) { $a += '--no-excel' }
  if ($Reuse) { $a += '--reuse' }
  if ($Sources -or $Random -gt 0) { $a += '--sources' }
  if ($Random -gt 0) { $a += @('--random', "$Random", '--seed', "$Seed") }
  & node @a
  exit $LASTEXITCODE
} finally { Pop-Location }
