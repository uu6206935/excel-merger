# Builds "exported workbook" variants of samples/01_学生名簿.xlsx with OUR OWN hidden Excel, to test
# tools/check-output.mjs (tests/qa/check-output.test.ts). Owner: T.
#   good.xlsx      3 imported columns appended after the table (M:O), then 判定 (O) moved to column A
#                  with the clipboard-free 3-step move (R2 section 7) -> PASS
#   converted.xlsx good + 集計!B5 converted to its value                          -> PASS (info)
#   inserted.xlsx  an empty column inserted right of 志望度 by Excel (CF / DV grow over it)      -> PASS (info)
#   ref.xlsx       column D (大学) deleted -> 集計 formulas become #REF!             -> FAIL
#   name.xlsx      集計!B24 uses an unknown function -> #NAME?                      -> FAIL
#   wrongcol.xlsx  集計!B5 points to 学部 instead of 大学 (a buggy engine)          -> FAIL
param(
  [Parameter(Mandatory = $true)][string]$Base,
  [Parameter(Mandatory = $true)][string]$OutDir
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '..\..\..\tools\lib\xlcom.ps1')

function L([int]$c) { $s = ''; while ($c -gt 0) { $m = ($c - 1) % 26; $s = [char](65 + $m) + $s; $c = [int][math]::Floor(($c - 1) / 26) }; return $s }

# clipboard-free move of columns [from .. from+n-1] before column `to` (original coordinates)
function Move-Cols($ws, [int]$from, [int]$n, [int]$to) {
  $blk = "$(L $to):$(L ($to + $n - 1))"
  [void]$ws.Range($blk).EntireColumn.Insert()
  if ($to -lt $from) { $src = "$(L ($from + $n)):$(L ($from + 2 * $n - 1))" } else { $src = "$(L $from):$(L ($from + $n - 1))" }
  [void]$ws.Range($src).EntireColumn.Cut($ws.Range($blk).EntireColumn)
  [void]$ws.Range($src).EntireColumn.Delete()
}

function Open-Base($xl) {
  $tmp = Join-Path $OutDir ('base-' + [Guid]::NewGuid().ToString('N').Substring(0, 8) + '.xlsx')
  Copy-Item -LiteralPath $Base -Destination $tmp
  $wb = $xl.Workbooks.Open($tmp)
  return @($wb, $tmp)
}

function Save-As($wb, [string]$name) {
  $p = Join-Path $OutDir $name
  if (Test-Path $p) { Remove-Item $p -Force }
  $wb.SaveAs($p, 51)
  Write-EmLine ([ordered]@{ type = 'saved'; file = $name })
}

function Add-Imported($ws) {
  # imported columns after the table (M:O), header row 3, rows 4..63
  $ws.Range('M3:O3').Formula = [object[]]@('面接官', '面接平均', '判定')
  $arr = New-Object 'object[,]' 60, 3
  for ($i = 0; $i -lt 60; $i++) {
    $arr[$i, 0] = [string](@('田村', '松井', '大野', '杉山')[$i % 4])
    $arr[$i, 1] = [double](2.5 + ($i % 5) * 0.5)
    $arr[$i, 2] = [string]$(if ((2.5 + ($i % 5) * 0.5) -ge 3.5) { '合格' } else { '保留' })
  }
  $ws.Range('M4:O63').Value2 = $arr
  $ws.Range('M3:O3').Font.Bold = $true
}

$xl = $null
try {
  $xl = New-EmExcel
  if (-not (Test-Path $OutDir)) { [void](New-Item -ItemType Directory -Path $OutDir -Force) }

  # good + converted
  $o = Open-Base $xl; $wb = $o[0]
  $ws = $wb.Worksheets.Item('名簿')
  Add-Imported $ws
  # the title merge A1:K1 blocks moving any of A..K (Excel refuses, and so does the app): move the imported 判定
  # column (O) to the front instead -> every 名簿 column shifts right by one
  Move-Cols $ws 15 1 1
  Save-As $wb 'good.xlsx'
  $sum = $wb.Worksheets.Item('集計')
  $v = $sum.Range('B5').Value2
  $sum.Range('B5').Value2 = $v
  Save-As $wb 'converted.xlsx'
  $wb.Close($false); Remove-Item $o[1] -Force

  # inserted: an empty column inserted right of 志望度 (K) by Excel itself: Excel extends the CF / DV of 志望度 over it
  $o = Open-Base $xl; $wb = $o[0]
  [void]$wb.Worksheets.Item('名簿').Range('L:L').EntireColumn.Insert()
  Save-As $wb 'inserted.xlsx'
  $wb.Close($false); Remove-Item $o[1] -Force

  # ref: delete 大学
  $o = Open-Base $xl; $wb = $o[0]
  [void]$wb.Worksheets.Item('名簿').Range('D:D').EntireColumn.Delete()
  Save-As $wb 'ref.xlsx'
  $wb.Close($false); Remove-Item $o[1] -Force

  # name: unknown function
  $o = Open-Base $xl; $wb = $o[0]
  $wb.Worksheets.Item('集計').Range('B24').Formula = '=COUNTBLANKX(名簿!G4:G63)'
  Save-As $wb 'name.xlsx'
  $wb.Close($false); Remove-Item $o[1] -Force

  # wrongcol: 集計!B5 counts 学部 instead of 大学
  $o = Open-Base $xl; $wb = $o[0]
  $wb.Worksheets.Item('集計').Range('B5').Formula = '=COUNTIF(名簿!$E:$E,$A5)'
  Save-As $wb 'wrongcol.xlsx'
  $wb.Close($false); Remove-Item $o[1] -Force

  Write-EmLine ([ordered]@{ type = 'done' })
}
catch {
  Write-EmLine ([ordered]@{ type = 'error'; message = [string]$_.Exception.Message; at = [string]$_.InvocationInfo.PositionMessage })
  exit 1
}
finally {
  Close-EmExcel $xl
}
