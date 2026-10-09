# Saves legacy.xls (in -Work) with OUR OWN hidden Excel instance (owner: reader R) as
#   legacy.ods       FileFormat 60 (OpenDocument Spreadsheet)
#   legacy.xlsb      FileFormat 50 (Excel binary workbook)
#   legacy2003.xls   FileFormat 46 (XML Spreadsheet 2003) under a .xls name, as business systems export it
#   legacy-excel.csv FileFormat 6 (CSV, Shift_JIS), legacy-utf8.csv FileFormat 62 (CSV UTF-8 with BOM),
#   legacy-unicode.txt FileFormat 42 (Unicode text: UTF-16LE with BOM, TAB separated) — first sheet only
# Run: powershell -NoProfile -ExecutionPolicy Bypass -File make-ods-xlsb.ps1 -Work <dir>
param([Parameter(Mandatory = $true)][string]$Work)
$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\common.ps1"
$xl = $null
try {
  $xl = New-RdrExcel
  foreach ($t in @(@{ n = 'legacy.ods'; f = 60 }, @{ n = 'legacy.xlsb'; f = 50 }, @{ n = 'legacy2003.xls'; f = 46 }, @{ n = 'legacy-excel.csv'; f = 6 }, @{ n = 'legacy-utf8.csv'; f = 62 }, @{ n = 'legacy-unicode.txt'; f = 42 })) {
    $wb = $xl.Workbooks.Open((Join-Path $Work 'legacy.xls'), 0, $true)
    $p = Join-Path $Work $t.n
    if (Test-Path $p) { Remove-Item $p -Force }
    $wb.SaveAs($p, $t.f)
    $wb.Close($false)
    Write-Host ('OK ' + $t.n)
  }
}
finally { Close-RdrExcel $xl }
