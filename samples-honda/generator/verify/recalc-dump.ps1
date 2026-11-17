param([string]$Dir, [string]$Out)
# 各 xlsx を自前の非表示 Excel で開き、全再計算した後の「数式セルの値」を JSON に書き出す
$ErrorActionPreference = 'Stop'
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$result = @()
try {
  foreach ($f in Get-ChildItem -LiteralPath $Dir -Filter '*.xlsx' | Sort-Object Name) {
    $wb = $xl.Workbooks.Open($f.FullName, 0, $true)
    try {
      $xl.CalculateFull()
      foreach ($ws in $wb.Worksheets) {
        $ur = $ws.UsedRange
        $rows = $ur.Rows.Count; $cols = $ur.Columns.Count
        $r0 = $ur.Row; $c0 = $ur.Column
        $fm = $ur.Formula; $vv = $ur.Value2
        for ($i = 1; $i -le $rows; $i++) {
          for ($j = 1; $j -le $cols; $j++) {
            $fx = $fm[$i, $j]
            if ($fx -is [string] -and $fx.StartsWith('=')) {
              $v = $vv[$i, $j]
              $result += [pscustomobject]@{ file = $f.Name; sheet = $ws.Name; row = $r0 + $i - 1; col = $c0 + $j - 1; f = $fx; v = $v }
            }
          }
        }
      }
    } finally { $wb.Close($false) }
  }
} finally { $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
$result | ConvertTo-Json -Depth 4 -Compress | Out-File -LiteralPath $Out -Encoding utf8
"formula cells: $($result.Count)"
