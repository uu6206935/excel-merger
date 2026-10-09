# Reads rich-text formatting runs and phonetic readings of cells as Excel shows them (owner: patcher P).
#   powershell -NoProfile -ExecutionPolicy Bypass -File tests\golden\com\probe-runs.ps1 -File <x.xlsx> -Cells "Sheet!E2,Sheet!E3"
# stdout: one ASCII JSON line per cell {"cell":..,"text":..,"runs":"1-4:b=True c=255 s=11 | 5-11:...","phonetic":..}
param([string]$File, [string]$Cells)
$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\common.ps1"
$xl = $null
try {
  $xl = New-GExcel
  $wb = $xl.Workbooks.Open((Resolve-Path $File).Path, 0, $true)
  foreach ($spec in $Cells.Split(',')) {
    $i = $spec.LastIndexOf('!')
    $sh = $spec.Substring(0, $i).Trim("'"); $addr = $spec.Substring($i + 1)
    $rg = $wb.Worksheets.Item($sh).Range($addr)
    $text = [string]$rg.Text
    $runs = @(); $cur = $null; $start = 1
    for ($k = 1; $k -le $text.Length; $k++) {
      $f = $rg.Characters($k, 1).Font
      $sig = 'b=' + $f.Bold + ' i=' + $f.Italic + ' c=' + $f.Color + ' s=' + $f.Size
      if ($null -ne $cur -and $sig -ne $cur) { $runs += ('' + $start + '-' + ($k - 1) + ':' + $cur); $start = $k }
      $cur = $sig
    }
    if ($null -ne $cur) { $runs += ('' + $start + '-' + $text.Length + ':' + $cur) }
    $ph = ''
    try { if ($rg.Phonetics.Count -gt 0) { $ph = [string]$rg.Phonetics.Item(1).Text } } catch {}
    Write-Host ('{"cell":' + (JStr $spec) + ',"text":' + (JStr $text) + ',"runs":' + (JStr ($runs -join ' | ')) + ',"phonetic":' + (JStr $ph) + '}')
  }
  $wb.Close($false)
}
finally { Close-GExcel $xl }
