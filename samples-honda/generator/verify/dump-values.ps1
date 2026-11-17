param([string]$File, [string]$Out, [string]$EditsJson = '')
# 自前の非表示 Excel で開く → （任意）セルを書き換え → 全再計算 → 1 枚目の値・数式・シート一覧を JSON へ
$ErrorActionPreference = 'Stop'
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
try {
  $wb = $xl.Workbooks.Open($File, 0, ($EditsJson -eq ''))
  try {
    if ($EditsJson -ne '') {
      $edits = Get-Content -LiteralPath $EditsJson -Raw -Encoding UTF8 | ConvertFrom-Json
      foreach ($e in $edits) { $sh = $wb.Worksheets.Item([string]$e.sheet); $sh.Range([string]$e.addr).Value2 = [double]$e.value }
    }
    $xl.CalculateFull()
    $names = @(); foreach ($ws in $wb.Worksheets) { $names += $ws.Name }
    $ws1 = $wb.Worksheets.Item(1); $ur = $ws1.UsedRange
    $rows = $ur.Rows.Count; $cols = $ur.Columns.Count
    $vv = $ur.Value2; $fm = $ur.Formula
    $vals = New-Object System.Collections.Generic.List[object]; $fors = New-Object System.Collections.Generic.List[object]
    for ($i = 1; $i -le $rows; $i++) {
      $rv = New-Object object[] $cols; $rf = New-Object object[] $cols
      for ($j = 1; $j -le $cols; $j++) { $rv[$j-1] = $vv[$i,$j]; $f = $fm[$i,$j]; if ($f -is [string] -and $f.StartsWith('=')) { $rf[$j-1] = $f } }
      $vals.Add($rv); $fors.Add($rf)
    }
    # 他のシートの数式セル数
    $fc = @{}
    foreach ($ws in $wb.Worksheets) { $n = 0; try { $n = $ws.UsedRange.SpecialCells(-4123).Count } catch { $n = 0 }; $fc[$ws.Name] = $n }
    $obj = [pscustomobject]@{ sheets = $names; rows = $rows; cols = $cols; values = $vals; formulas = $fors; formulaCounts = $fc }
    ($obj | ConvertTo-Json -Depth 6 -Compress) | Out-File -LiteralPath $Out -Encoding utf8
    'ok'
  } finally { $wb.Close($false) }
} finally { $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
