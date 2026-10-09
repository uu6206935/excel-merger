# Records what Excel itself reports for the reader fixtures (owner: reader R) into excel-probes.json.
#  - "cells": for every cell of the used range of the listed workbooks/sheets: Text (display text), Value2, NumberFormat
#  - "styles": for style-probe sheets: font / fill / alignment / borders as Excel reports them
# Inputs must be copied to an ASCII-only directory first (-Work). Output: <Work>\excel-probes.json (UTF-8, no BOM).
# Run: powershell -NoProfile -ExecutionPolicy Bypass -File record-probes.ps1 -Work <dir>
param([Parameter(Mandatory = $true)][string]$Work)
$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\common.ps1"

# file -> @{ sheets = @(...) or '*'; styles = $true/$false; maxRows = n }
$plan = @(
  @{ file = 'torture.xlsx'; styles = @('Data', 'Tbl', '集計 シート'); maxRows = 60 },
  @{ file = 'apply-probe.xlsx'; styles = @('Probe'); maxRows = 20 },
  @{ file = 'interview1.xlsx'; styles = @(); maxRows = 25 },
  @{ file = 'interview2.xlsx'; styles = @(); maxRows = 25 },
  @{ file = 'interview3.xlsx'; styles = @(); maxRows = 25 },
  @{ file = 'interview4.xlsx'; styles = @(); maxRows = 25 },
  @{ file = 'storage-test.xlsx'; styles = @(); maxRows = 60 }
)

function Get-Border($c, [int]$idx) {
  $b = $c.Borders.Item($idx)
  $ls = $b.LineStyle
  if ($null -eq $ls -or $ls -is [DBNull] -or [int]$ls -eq -4142) { return $null }
  return [ordered]@{ lineStyle = [int]$ls; weight = [int]$b.Weight; color = (ConvertTo-Rgb $b.Color) }
}

$result = [ordered]@{ excelVersion = ''; files = @() }
$xl = $null
try {
  $xl = New-RdrExcel
  $result.excelVersion = [string]$xl.Version
  foreach ($p in $plan) {
    $path = Join-Path $Work $p.file
    if (-not (Test-Path $path)) { Write-Host ('SKIP ' + $p.file); continue }
    $wb = $xl.Workbooks.Open($path, 0, $true)
    $fileRec = [ordered]@{ file = $p.file; date1904 = [bool]$wb.Date1904; sheets = @() }
    foreach ($ws in @($wb.Worksheets)) {
      $name = [string]$ws.Name
      $ur = $ws.UsedRange
      $r0 = [int]$ur.Row; $c0 = [int]$ur.Column
      $nr = [int]$ur.Rows.Count; $nc = [int]$ur.Columns.Count
      $rmax = [math]::Min($r0 + $nr - 1, $p.maxRows)
      $cmax = [math]::Min($c0 + $nc - 1, 60)
      $doStyles = $p.styles -contains $name
      $cells = @()
      for ($r = $r0; $r -le $rmax; $r++) {
        for ($k = $c0; $k -le $cmax; $k++) {
          $c = $ws.Cells.Item($r, $k)
          $v2 = $c.Value2
          $hasF = [bool]$c.HasFormula
          if (($null -eq $v2) -and (-not $hasF) -and (-not $doStyles)) { continue }
          $vt = if ($null -eq $v2) { 'null' } elseif ($v2 -is [double]) { 'n' } elseif ($v2 -is [bool]) { 'b' } elseif ($v2 -is [int]) { 'e' } else { 's' }
          $rec = [ordered]@{ addr = [string]$c.Address($false, $false); text = [string]$c.Text; vt = $vt; v = $(if ($vt -eq 'e') { [int]$v2 } elseif ($null -eq $v2) { $null } else { $v2 }); nf = [string]$c.NumberFormat }
          if ($doStyles) {
            $f = $c.Font
            $rec.font = [ordered]@{ name = [string]$f.Name; size = [double]$f.Size; bold = [bool]$f.Bold; italic = [bool]$f.Italic; underline = [int]$f.Underline; strike = [bool]$f.Strikethrough; color = (ConvertTo-Rgb $f.Color) }
            $pat = [int]$c.Interior.Pattern
            $rec.fill = [ordered]@{ pattern = $pat; color = $(if ($pat -eq -4142) { $null } else { ConvertTo-Rgb $c.Interior.Color }) }
            $rec.hAlign = [int]$c.HorizontalAlignment
            $rec.vAlign = [int]$c.VerticalAlignment
            $rec.wrap = [bool]$c.WrapText
            $rec.indent = [int]$c.IndentLevel
            $rec.borders = [ordered]@{ left = (Get-Border $c 7); top = (Get-Border $c 8); bottom = (Get-Border $c 9); right = (Get-Border $c 10) }
          }
          $cells += $rec
        }
      }
      $fileRec.sheets += [ordered]@{ name = $name; usedRange = [string]$ur.Address($false, $false); cells = $cells }
    }
    $wb.Close($false)
    $result.files += $fileRec
    Write-Host ('OK ' + $p.file)
  }
  Save-Json $result (Join-Path $Work 'excel-probes.json')
  Write-Host 'DONE'
}
finally { Close-RdrExcel $xl }
