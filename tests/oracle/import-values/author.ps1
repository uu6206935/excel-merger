# Corpus authoring with real Excel (task F1): ONE hidden Excel instance of our own for the whole batch.
#   powershell -NoProfile -ExecutionPolicy Bypass -File author.ps1 -CorpusDir <dir>
# Reads <dir>\manifest.json (gen_corpus.py), opens every <dir>\nocache\<file>, applies the comPost steps (a formula assigned to a whole
# range is stored as a SHARED formula), recalculates (CalculateFull) and saves <dir>\cached\<file> (Excel-written cached results).
# comSaveXlsm additionally saves <name>.xlsm (macro-enabled format, no macros). ASCII-only stdout (JSON lines).
param([Parameter(Mandatory = $true)][string]$CorpusDir, [string]$Only = '')
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '..\..\golden\com\common.ps1')
$M = $script:GMissing
$CorpusDir = (Resolve-Path $CorpusDir).Path
$cached = Join-Path $CorpusDir 'cached'
if (-not (Test-Path $cached)) { New-Item -ItemType Directory -Path $cached | Out-Null }
$manifest = Get-Content -Raw -Encoding UTF8 (Join-Path $CorpusDir 'manifest.json') | ConvertFrom-Json

$xl = $null
try {
  $xl = New-GExcel
  Write-Host ('{"type":"excel","pid":' + $script:GExcelPid + '}')
  foreach ($e in $manifest) {
    if ($Only -and ($e.file -notmatch $Only)) { continue }
    $src = Join-Path (Join-Path $CorpusDir 'nocache') $e.file
    $dst = Join-Path $cached $e.file
    if (Test-Path $dst) { Remove-Item $dst -Force }
    $ok = $true; $msg = ''
    try {
      $wb = $xl.Workbooks.Open($src, 0)
      $ws = $wb.Worksheets.Item(1)
      foreach ($p in @($e.comPost)) {
        if ($null -eq $p) { continue }
        if ($p.op -eq 'formula') { $ws.Range($p.range).Formula = $p.formula }
      }
      $xl.CalculateFull()
      $wb.SaveAs($dst, 51)
      if ($e.comSaveXlsm) {
        $x2 = Join-Path $cached (([IO.Path]::GetFileNameWithoutExtension($e.file)) + '.xlsm')
        if (Test-Path $x2) { Remove-Item $x2 -Force }
        $wb.SaveAs($x2, 52)
      }
      $wb.Close($false)
    } catch { $ok = $false; $msg = $_.Exception.Message -replace '[^\x20-\x7E]', '?' }
    Write-Host ('{"type":"file","file":' + (JStr $e.file) + ',"ok":' + ($ok.ToString().ToLower()) + ',"msg":' + (JStr $msg) + '}')
  }
} finally {
  Close-GExcel $xl
}
