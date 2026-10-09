# Builds tests/fixtures/xlsx/theme-probe.xlsx from torture.xlsx (custom theme: accent1 = CC9933) by adding a sheet
# "Probe" with theme colours 1..12 x tints (fill + font) and indexed colours 1..56, saves it, reopens the SAVED file
# and records the colours Excel reports (Interior.Color / Font.Color) into excel-colors.json.
# Run: powershell -NoProfile -ExecutionPolicy Bypass -File make-theme-probe.ps1 -Dir <tests\fixtures\xlsx>
param([Parameter(Mandatory = $true)][string]$Work)
$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\common.ps1"
$tints = @(-0.9, -0.5, -0.25, -0.1, 0, 0.05, 0.15, 0.25, 0.35, 0.4, 0.5, 0.6, 0.8, 0.95)
$src = Join-Path $Work 'torture.xlsx'
$tmpIn = Join-Path $Work 'probe-src.xlsx'
$tmpOut = Join-Path $Work 'theme-probe.xlsx'
Copy-Item $src $tmpIn -Force
if (Test-Path $tmpOut) { Remove-Item $tmpOut -Force }
$xl = $null
$result = [ordered]@{ tints = $tints; theme = @(); indexed = @() }
try {
  $xl = New-RdrExcel
  $wb = $xl.Workbooks.Open($tmpIn)
  $ws = $wb.Worksheets.Add($script:RdrMissing, $wb.Worksheets.Item($wb.Worksheets.Count))
  $ws.Name = 'Probe'
  for ($k = 1; $k -le 12; $k++) {
    for ($j = 0; $j -lt $tints.Count; $j++) {
      $c = $ws.Cells.Item($k, $j + 1)
      $c.Value2 = 'x'
      $c.Interior.Pattern = 1
      $c.Interior.ThemeColor = $k
      $c.Interior.TintAndShade = [double]$tints[$j]
      $c.Font.ThemeColor = $k
      $c.Font.TintAndShade = [double]$tints[$j]
    }
  }
  for ($i = 1; $i -le 56; $i++) {
    $c = $ws.Cells.Item([int](14 + [math]::Floor(($i - 1) / 14)), [int]((($i - 1) % 14) + 1))
    $c.Value2 = 'i'
    $c.Interior.ColorIndex = $i
    $c.Font.ColorIndex = [int]((($i + 5) % 56) + 1)
  }
  $wb.SaveAs($tmpOut, 51)
  $wb.Close($false)
  # reopen the saved file and read what Excel reports
  $wb = $xl.Workbooks.Open($tmpOut, 0, $true)
  $ws = $wb.Worksheets.Item('Probe')
  for ($k = 1; $k -le 12; $k++) {
    for ($j = 0; $j -lt $tints.Count; $j++) {
      $c = $ws.Cells.Item($k, $j + 1)
      $result.theme += [ordered]@{ addr = [string]$c.Address($false, $false); comTheme = $k; tint = $tints[$j]; fill = (ConvertTo-Rgb $c.Interior.Color); font = (ConvertTo-Rgb $c.Font.Color) }
    }
  }
  for ($i = 1; $i -le 56; $i++) {
    $c = $ws.Cells.Item([int](14 + [math]::Floor(($i - 1) / 14)), [int]((($i - 1) % 14) + 1))
    $result.indexed += [ordered]@{ addr = [string]$c.Address($false, $false); colorIndex = $i; fill = (ConvertTo-Rgb $c.Interior.Color); font = (ConvertTo-Rgb $c.Font.Color) }
  }
  $wb.Close($false)
  # output: $tmpOut
  Save-Json $result (Join-Path $Work 'excel-colors.json')
  Write-Host ('OK theme=' + $result.theme.Count + ' indexed=' + $result.indexed.Count)
}
finally { Close-RdrExcel $xl }
