# Probe: which conditional-format rule types grow over an inserted column / leak onto a moved block?
# Each rule type lives on its own rows in column C (C<r>:C<r+1>). Saves cf-before.xlsx, then
#   cf-ins.xlsx  : insert a column at D
#   cf-mov.xlsx  : move column F before D (clipboard-free: insert blank at D, Range.Cut(Destination), delete source)
param([string]$OutDir)
$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\common.ps1"
$OutDir = (Resolve-Path $OutDir).Path
$M = $script:GMissing

function Build($xl) {
  $wb = $xl.Workbooks.Add()
  $ws = $wb.Worksheets.Item(1)
  $ws.Name = 'P'
  for ($r = 1; $r -le 40; $r++) { for ($c = 1; $c -le 8; $c++) { SetF $ws.Cells.Item($r, $c) ($r * 10 + $c) } }
  $ws.Activate()
  $i = 0
  $row = { param($k) return 2 + 3 * $k }
  $r = & $row 0; $ws.Range("C$($r)").Select(); $f = $ws.Range("C$($r):C$($r+1)").FormatConditions.Add(1, 5, '=5'); $f.Interior.Color = 255          # cellIs
  $r = & $row 1; $ws.Range("C$($r)").Select(); $f = $ws.Range("C$($r):C$($r+1)").FormatConditions.Add(2, $M, "=C$($r)>0"); $f.Interior.Color = 65280  # expression
  $r = & $row 2; [void]$ws.Range("C$($r):C$($r+1)").FormatConditions.AddColorScale(3)
  $r = & $row 3; [void]$ws.Range("C$($r):C$($r+1)").FormatConditions.AddDatabar()
  $r = & $row 4; [void]$ws.Range("C$($r):C$($r+1)").FormatConditions.AddIconSetCondition()
  $r = & $row 5; $f = $ws.Range("C$($r):C$($r+1)").FormatConditions.AddTop10(); $f.Interior.Color = 16711680                                               # top10
  $r = & $row 6; $f = $ws.Range("C$($r):C$($r+1)").FormatConditions.AddUniqueValues(); $f.Interior.Color = 65535                                        # duplicate/unique
  $r = & $row 7; $f = $ws.Range("C$($r):C$($r+1)").FormatConditions.AddAboveAverage(); $f.Interior.Color = 16776960                                     # aboveAverage
  $r = & $row 8; $f = $ws.Range("C$($r):C$($r+1)").FormatConditions.Add(9, $M, $M, $M, 'x', 0); $f.Interior.Color = 12632256                                        # containsText (xlTextString)
  $r = & $row 9; $f = $ws.Range("C$($r):C$($r+1)").FormatConditions.Add(10, $M, $M, $M, $M, $M, $M, $M); $f.Interior.Color = 8421504                                                # blanks
  $r = & $row 10; [void]$ws.Range("C$($r):C$($r+1)").Validation.Add(1, 1, 1, '0', '1000')                                                            # DV whole
  return $wb
}

$xl = $null
try {
  $xl = New-GExcel
  $wb = @(Build $xl)[-1]
  $p = Join-Path $OutDir 'cf-before.xlsx'; if (Test-Path $p) { Remove-Item $p -Force }; $wb.SaveAs($p, 51); $wb.Close($false)
  $wb = @(Build $xl)[-1]; $ws = $wb.Worksheets.Item(1)
  [void]$ws.Range('D:D').EntireColumn.Insert()
  $p = Join-Path $OutDir 'cf-ins.xlsx'; if (Test-Path $p) { Remove-Item $p -Force }; $wb.SaveAs($p, 51); $wb.Close($false)
  $wb = @(Build $xl)[-1]; $ws = $wb.Worksheets.Item(1)
  [void]$ws.Range('D:D').EntireColumn.Insert()
  [void]$ws.Range('G:G').EntireColumn.Cut($ws.Range('D:D').EntireColumn)
  [void]$ws.Range('G:G').EntireColumn.Delete()
  $p = Join-Path $OutDir 'cf-mov.xlsx'; if (Test-Path $p) { Remove-Item $p -Force }; $wb.SaveAs($p, 51); $wb.Close($false)
  Write-Host 'cf probe saved'
}
finally { Close-GExcel $xl }
