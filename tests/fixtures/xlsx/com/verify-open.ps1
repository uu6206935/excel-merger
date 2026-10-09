# Opens every .xlsx in -Dir in OUR OWN hidden Excel and records whether it opens WITHOUT repair (owner: reader R).
# With DisplayAlerts=$false a plain Workbooks.Open(path, 0, $true) throws when the file needs repair (R1 §2), so
# "opened": true means no repair. Also records Text / Value2 of the first sheet's used range (max 20x10).
# Output: <Dir>\verify-open.json. Run: powershell -NoProfile -ExecutionPolicy Bypass -File verify-open.ps1 -Dir <dir>
param([Parameter(Mandatory = $true)][string]$Dir)
$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\common.ps1"
$result = @()
$xl = $null
try {
  $xl = New-RdrExcel
  foreach ($f in (Get-ChildItem -Path $Dir -Filter '*.xlsx' | Sort-Object Name)) {
    $rec = [ordered]@{ file = $f.Name; opened = $false; error = $null; sheets = @(); cells = @() }
    $wb = $null
    try {
      $wb = $xl.Workbooks.Open($f.FullName, 0, $true)
      $rec.opened = $true
      foreach ($ws in @($wb.Worksheets)) { $rec.sheets += [ordered]@{ name = [string]$ws.Name; visible = [int]$ws.Visible } }
      $ws = $wb.Worksheets.Item(1)
      $ur = $ws.UsedRange
      $rmax = [math]::Min([int]$ur.Row + [int]$ur.Rows.Count - 1, 20)
      $cmax = [math]::Min([int]$ur.Column + [int]$ur.Columns.Count - 1, 10)
      for ($r = 1; $r -le $rmax; $r++) {
        for ($c = 1; $c -le $cmax; $c++) {
          $cell = $ws.Cells.Item($r, $c)
          $v2 = $cell.Value2
          if ($null -eq $v2) { continue }
          $vt = if ($v2 -is [double]) { 'n' } elseif ($v2 -is [bool]) { 'b' } elseif ($v2 -is [int]) { 'e' } else { 's' }
          $rec.cells += [ordered]@{ addr = [string]$cell.Address($false, $false); text = [string]$cell.Text; vt = $vt; v = $(if ($vt -eq 'e') { [int]$v2 } else { $v2 }) }
        }
      }
    }
    catch { $rec.error = [string]$_.Exception.Message }
    finally { if ($wb) { try { $wb.Close($false) } catch {} } }
    $result += $rec
    Write-Host ('{0} opened={1}' -f $f.Name, $rec.opened)
  }
  Save-Json $result (Join-Path $Dir 'verify-open.json')
  Write-Host 'DONE'
}
finally { Close-RdrExcel $xl }
