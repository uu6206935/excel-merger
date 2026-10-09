#Requires -Version 5.1
<#
  Prints the current clipboard TEXT (read-only; never changes the clipboard) as one ASCII JSON line:
  {"type":"clip","text":"..."|null,"seq":N}. Uses Win32 directly (no STA/WinForms requirement).
#>
$ErrorActionPreference = 'Stop'
Add-Type -Namespace EmTest -Name Clip -MemberDefinition @'
[DllImport("user32.dll", SetLastError = true)] public static extern bool OpenClipboard(System.IntPtr h);
[DllImport("user32.dll")] public static extern bool CloseClipboard();
[DllImport("user32.dll")] public static extern System.IntPtr GetClipboardData(uint f);
[DllImport("user32.dll")] public static extern bool IsClipboardFormatAvailable(uint f);
[DllImport("user32.dll")] public static extern uint GetClipboardSequenceNumber();
[DllImport("kernel32.dll")] public static extern System.IntPtr GlobalLock(System.IntPtr h);
[DllImport("kernel32.dll")] public static extern bool GlobalUnlock(System.IntPtr h);
public static string Text() {
  if (!IsClipboardFormatAvailable(13)) return null;
  for (int i = 0; i < 40; i++) { if (OpenClipboard(System.IntPtr.Zero)) break; System.Threading.Thread.Sleep(50); if (i == 39) throw new System.Exception("clipboard busy"); }
  try {
    System.IntPtr h = GetClipboardData(13);
    if (h == System.IntPtr.Zero) return null;
    System.IntPtr p = GlobalLock(h);
    try { return System.Runtime.InteropServices.Marshal.PtrToStringUni(p); } finally { GlobalUnlock(h); }
  } finally { CloseClipboard(); }
}
'@
function Esc([string]$s) {
  $sb = New-Object System.Text.StringBuilder
  foreach ($ch in $s.ToCharArray()) {
    $n = [int]$ch
    if ($ch -eq '"' -or $ch -eq '\') { [void]$sb.Append('\').Append($ch) }
    elseif ($n -lt 32 -or $n -gt 126) { [void]$sb.Append(('\u{0:x4}' -f $n)) }
    else { [void]$sb.Append($ch) }
  }
  return $sb.ToString()
}
$t = [EmTest.Clip]::Text()
$seq = [EmTest.Clip]::GetClipboardSequenceNumber()
if ($null -eq $t) { [Console]::Out.WriteLine('{"type":"clip","text":null,"seq":' + $seq + '}') }
else { [Console]::Out.WriteLine('{"type":"clip","text":"' + (Esc $t) + '","seq":' + $seq + '}') }
