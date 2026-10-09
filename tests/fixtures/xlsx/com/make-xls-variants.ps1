# Saves Excel-made .xls (BIFF8, FileFormat 56) variants of reader fixtures with OUR OWN hidden Excel instance
# (owner: reader R), and records what Excel shows for them (Text / Value2 of every used cell of every sheet):
#   torture.xls    torture.xlsx (shared / array formulas, names, CF, DV, charts, comments, rich text, many formats)
#   date1904.xls   date1904.xlsx (1904 date system)
# Inputs are copied to an ASCII-only -Work directory first. Output: <Work>\xls-variants.json (UTF-8, no BOM).
# Run: powershell -NoProfile -ExecutionPolicy Bypass -File make-xls-variants.ps1 -Work <dir>
param([Parameter(Mandatory = $true)][string]$Work)
$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\common.ps1"
$result = [ordered]@{}
$xl = $null
try {
  $xl = New-RdrExcel
  foreach ($n in @('torture', 'date1904')) {
    $wb = $xl.Workbooks.Open((Join-Path $Work ($n + '.xlsx')), 0, $true)
    $p = Join-Path $Work ($n + '.xls')
    if (Test-Path $p) { Remove-Item $p -Force }
    $wb.SaveAs($p, 56)
    $wb.Close($false)
    $wb = $xl.Workbooks.Open($p, 0, $true)
    $sheets = @()
    foreach ($ws in @($wb.Worksheets)) {
      $ur = $ws.UsedRange
      $r0 = [int]$ur.Row; $c0 = [int]$ur.Column
      $rmax = [math]::Min($r0 + [int]$ur.Rows.Count - 1, 80)
      $cmax = [math]::Min($c0 + [int]$ur.Columns.Count - 1, 40)
      $cells = @()
      for ($r = $r0; $r -le $rmax; $r++) {
        for ($c = $c0; $c -le $cmax; $c++) {
          $cell = $ws.Cells.Item($r, $c)
          $v2 = $cell.Value2
          if ($null -eq $v2) { continue }
          $vt = if ($v2 -is [double]) { 'n' } elseif ($v2 -is [bool]) { 'b' } elseif ($v2 -is [int]) { 'e' } else { 's' }
          $cells += [ordered]@{ addr = [string]$cell.Address($false, $false); text = [string]$cell.Text; vt = $vt; v = $(if ($vt -eq 'e') { [int]$v2 } else { $v2 }) }
        }
      }
      $sheets += [ordered]@{ name = [string]$ws.Name; visible = [int]$ws.Visible; cells = $cells }
    }
    $result[$n] = [ordered]@{ date1904 = [bool]$wb.Date1904; sheets = $sheets }
    $wb.Close($false)
    Write-Host ('OK ' + $n + '.xls')
  }
  Save-Json $result (Join-Path $Work 'xls-variants.json')
  Write-Host 'DONE'
}
finally { Close-RdrExcel $xl }
