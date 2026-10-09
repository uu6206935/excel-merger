# Probe: CF expression rule A2:F10 ($D2>=80) on a plain range (no table); move D before B with the
# clipboard-free sequence; saves cfm-before.xlsx / cfm-after.xlsx
param([string]$OutDir)
$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\common.ps1"
$OutDir = (Resolve-Path $OutDir).Path
$M = $script:GMissing
$xl = $null
try {
  $xl = New-GExcel
  $wb = $xl.Workbooks.Add(); $ws = $wb.Worksheets.Item(1)
  for ($r = 1; $r -le 12; $r++) { for ($c = 1; $c -le 8; $c++) { SetF $ws.Cells.Item($r, $c) ($r * 10 + $c) } }
  $ws.Activate(); [void]$ws.Range('A2').Select()
  $f = $ws.Range('A2:F10').FormatConditions.Add(2, $M, '=$D2>=80'); $f.Interior.Color = 65280
  [void]$ws.Range('E2').Select()
  $f = $ws.Range('E2:E10').FormatConditions.Add(1, 5, '=90'); $f.Font.Bold = $true
  $p = Join-Path $OutDir 'cfm-before.xlsx'; if (Test-Path $p) { Remove-Item $p -Force }; $wb.SaveAs($p, 51)
  [void]$ws.Range('B:B').EntireColumn.Insert()
  [void]$ws.Range('E:E').EntireColumn.Cut($ws.Range('B:B').EntireColumn)
  [void]$ws.Range('E:E').EntireColumn.Delete()
  $p = Join-Path $OutDir 'cfm-after.xlsx'; if (Test-Path $p) { Remove-Item $p -Force }; $wb.SaveAs($p, 51)
  $wb.Close($false)
  Write-Host 'saved'
}
finally { Close-GExcel $xl }
