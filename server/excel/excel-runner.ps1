#Requires -Version 5.1
<#
  excel-runner.ps1 - "Excel本体" export engine of Excel 統合ツール.
  SOURCE file: build-runner.mjs merges XlRunnerHelper.cs into it and writes excel-runner.ps1
  (UTF-8 WITH BOM, CRLF). Do not run this .src.ps1 directly.

  Usage (the server uses the ScriptBlock form, which also works under a Restricted execution policy):
    powershell.exe -NoProfile -NonInteractive -Command "& ([ScriptBlock]::Create([IO.File]::ReadAllText('<ps1>', [Text.Encoding]::UTF8))) -JobPath '<job.json>'"
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File excel-runner.ps1 <job.json>
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File excel-runner.ps1 -Probe

  Opens a COPY of job.input in a NEW, hidden Excel instance of its own (never the user's Excel), replays
  job.ops (insertCols / deleteCols / moveCols / insertRows / deleteRows / moveRows / setCells /
  setColumnWidth / unmerge), so Excel itself adjusts every reference (formulas on all sheets, names, CF, DV, charts,
  pivots, tables), and saves job.output (.xlsx / .xlsm). Job format: see server/excel/README.md.

  stdout: one ASCII-only JSON object per line:
    {"type":"start"} {"type":"excel","pid":N} {"type":"progress",...} {"type":"warn",...}
    and finally {"type":"result","ok":true,...} or {"type":"error","ok":false,"code":...,"message":...}
  exit codes: 0 ok, 2 bad job, 3 Excel unavailable, 4 open failed, 5 op failed / refused, 6 save failed,
              7 output locked, 9 internal error.
  Environment: XLRUNNER_CACHE (folder for the compiled helper DLL), XLRUNNER_ALLOW_DEBUG=1 (test ops).
#>
param(
  [Parameter(Position = 0)][string]$JobPath,
  [switch]$Probe,
  [int]$QuitWaitMs = 0,        # how long to wait for Excel to exit before returning. 0 (default, round 5): not at all - a hidden Excel
                               # that has been told to Quit() does NOT exit before this process lets go of it (measured: never within 2 s),
                               # and waiting only delayed the result; the caller (server) watches the PID and kills ONLY it after a few seconds
  [switch]$KillOnExit          # standalone use: kill our own Excel if it has not exited after QuitWaitMs
)
$script:QuitWaitMs = $QuitWaitMs; $script:KillOnExit = [bool]$KillOnExit

$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding $false } catch {}
$script:Clock = [Diagnostics.Stopwatch]::StartNew()
$script:M = [Type]::Missing
$script:Warnings = New-Object System.Collections.ArrayList
$script:Timings = New-Object System.Collections.ArrayList
$script:ExitCode = 0
$script:Final = $null
$script:xl = $null
$script:XlPid = 0
$script:Wb = $null
$script:ClipSnapshot = $null
$script:ClipTouched = $false
$script:ClipPending = $false
$script:ClipMode = 'restoreEachOp'   # or 'restoreAtEnd' (job.options.clipboard)
$script:ExcelExited = $true
$script:OwnWork = $null      # scratch folder this runner created itself (job has no workDir): removed at the end
$script:TempFiles = New-Object System.Collections.ArrayList   # copy / saved workbook in the work folder: removed at the end, also after a failure
$script:HelperMs = 0
$script:RunnerVersion = '1.1.0'
$script:MaxCols = 16384
$script:MaxRows = 1048576

Add-Type -AssemblyName System.Web.Extensions
$script:Json = New-Object System.Web.Script.Serialization.JavaScriptSerializer
$script:Json.MaxJsonLength = [int]::MaxValue
$script:Json.RecursionLimit = 256

$HelperSource = @'
// Helper types for excel-runner.ps1 (merged into it by build-runner.mjs).
// C# 5 only (Windows PowerShell 5.1 Add-Type uses the .NET Framework csc): no $"", ?., =>, out var.
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;

namespace XlRunner
{
    public static class Native
    {
        [DllImport("user32.dll")]
        public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
        [DllImport("user32.dll")]
        public static extern uint GetClipboardSequenceNumber();
    }

    [ComImport, Guid("00000016-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IOleMessageFilter
    {
        [PreserveSig] int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo);
        [PreserveSig] int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType);
        [PreserveSig] int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType);
    }

    /// Retries calls that Excel rejects while busy (RPC_E_CALL_REJECTED / RPC_E_SERVERCALL_RETRYLATER).
    public class MessageFilter : IOleMessageFilter
    {
        [DllImport("ole32.dll")]
        private static extern int CoRegisterMessageFilter(IOleMessageFilter newFilter, out IOleMessageFilter oldFilter);
        public static int RetryLimitMs = 120000;
        public static int Retries = 0;
        public static int Register() { IOleMessageFilter old; return CoRegisterMessageFilter(new MessageFilter(), out old); }
        public static int Revoke() { IOleMessageFilter old; return CoRegisterMessageFilter(null, out old); }
        public int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo) { return 0; }
        public int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType)
        {
            if (dwRejectType == 2 && dwTickCount < RetryLimitMs) { Retries++; return 200; }
            return -1;
        }
        public int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType) { return 2; }
    }

    /// Byte-level clipboard snapshot/restore through Win32 (no WinForms DataObject: asking a foreign
    /// Excel for every clipboard format crashed powershell.exe with 0xC000041D in our tests).
    /// Only an allow-list of plain HGLOBAL formats is kept: text, locale, HTML, RTF, CSV, DIB/PNG images
    /// and copied files. GDI handles (bitmaps/metafiles) and OLE/Excel-private formats are skipped.
    public static class Clip
    {
        [DllImport("user32.dll", SetLastError = true)] static extern bool OpenClipboard(IntPtr hWndNewOwner);
        [DllImport("user32.dll", SetLastError = true)] static extern bool CloseClipboard();
        [DllImport("user32.dll", SetLastError = true)] static extern bool EmptyClipboard();
        [DllImport("user32.dll")] static extern uint EnumClipboardFormats(uint format);
        [DllImport("user32.dll")] static extern IntPtr GetClipboardData(uint uFormat);
        [DllImport("user32.dll")] static extern IntPtr SetClipboardData(uint uFormat, IntPtr hMem);
        [DllImport("user32.dll")] static extern int CountClipboardFormats();
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetClipboardFormatName(uint format, StringBuilder name, int cch);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern uint RegisterClipboardFormat(string name);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern IntPtr CreateWindowEx(int exStyle, string cls, string name, int style, int x, int y, int w, int h, IntPtr parent, IntPtr menu, IntPtr inst, IntPtr param);
        [DllImport("kernel32.dll")] static extern IntPtr GlobalLock(IntPtr h);
        [DllImport("kernel32.dll")] static extern bool GlobalUnlock(IntPtr h);
        [DllImport("kernel32.dll")] static extern UIntPtr GlobalSize(IntPtr h);
        [DllImport("kernel32.dll")] static extern IntPtr GlobalAlloc(uint flags, UIntPtr bytes);
        [DllImport("kernel32.dll")] static extern IntPtr GlobalFree(IntPtr h);

        static IntPtr _hwnd = IntPtr.Zero;
        static IntPtr Owner()
        {
            if (_hwnd == IntPtr.Zero) _hwnd = CreateWindowEx(0, "STATIC", "xlrunner-clip", 0, 0, 0, 0, 0, new IntPtr(-3), IntPtr.Zero, IntPtr.Zero, IntPtr.Zero); // HWND_MESSAGE
            return _hwnd;
        }
        static readonly uint[] Std = { 1, 7, 13, 16, 8, 17, 15 };   // TEXT, OEMTEXT, UNICODETEXT, LOCALE, DIB, DIBV5, HDROP
        static readonly string[] Named = { "HTML Format", "Rich Text Format", "CSV", "Csv", "PNG", "image/png", "FileName", "FileNameW", "Preferred DropEffect", "Shell IDList Array", "text/html", "text/plain" };

        public sealed class Snapshot
        {
            public List<uint> Formats = new List<uint>();
            public List<string> Names = new List<string>();
            public List<byte[]> Data = new List<byte[]>();
            public int Total, Skipped;
            public bool Empty { get { return Total == 0; } }
        }
        static bool TryOpen(int tries, int sleepMs)
        {
            for (int i = 0; i < tries; i++) { if (OpenClipboard(Owner())) return true; System.Threading.Thread.Sleep(sleepMs); }
            return false;
        }
        static string NameOf(uint f)
        {
            StringBuilder sb = new StringBuilder(256);
            int n = GetClipboardFormatName(f, sb, sb.Capacity);
            return n > 0 ? sb.ToString() : null;
        }
        /// null = clipboard could not be opened (another program keeps it open).
        public static Snapshot Save(long maxBytesPerFormat)
        {
            if (!TryOpen(60, 50)) return null;
            Snapshot s = new Snapshot();
            try
            {
                uint f = 0;
                while ((f = EnumClipboardFormats(f)) != 0)
                {
                    s.Total++;
                    string name = f >= 0xC000 ? NameOf(f) : null;
                    bool ok = Array.IndexOf(Std, f) >= 0 || (name != null && Array.IndexOf(Named, name) >= 0);
                    if (!ok) { s.Skipped++; continue; }
                    IntPtr h = GetClipboardData(f);                 // may ask the owner to render
                    if (h == IntPtr.Zero) { s.Skipped++; continue; }
                    long size = (long)GlobalSize(h).ToUInt64();
                    if (size <= 0 || size > maxBytesPerFormat) { s.Skipped++; continue; }
                    IntPtr p = GlobalLock(h);
                    if (p == IntPtr.Zero) { s.Skipped++; continue; }
                    try { byte[] b = new byte[size]; Marshal.Copy(p, b, 0, (int)size); s.Formats.Add(f); s.Names.Add(name); s.Data.Add(b); }
                    finally { GlobalUnlock(h); }
                }
            }
            finally { CloseClipboard(); }
            return s;
        }
        /// Puts the snapshot back (or empties the clipboard when the snapshot was empty).
        public static bool Restore(Snapshot s, int tries)
        {
            if (!TryOpen(tries, 50)) return false;
            try
            {
                if (!EmptyClipboard()) return false;
                for (int i = 0; i < s.Formats.Count; i++)
                {
                    uint f = s.Names[i] != null ? RegisterClipboardFormat(s.Names[i]) : s.Formats[i];
                    byte[] b = s.Data[i];
                    IntPtr h = GlobalAlloc(0x0002, new UIntPtr((ulong)b.Length));   // GMEM_MOVEABLE
                    if (h == IntPtr.Zero) continue;
                    IntPtr p = GlobalLock(h);
                    Marshal.Copy(b, 0, p, b.Length);
                    GlobalUnlock(h);
                    if (SetClipboardData(f, h) == IntPtr.Zero) GlobalFree(h);   // on success the system owns h
                }
                return true;
            }
            finally { CloseClipboard(); }
        }
        public static int Count() { return CountClipboardFormats(); }

        /// Keeps the clipboard open (without emptying it) so another program cannot put anything on it; Excel's
        /// Range.Cut() then keeps the cut internal. false = could not open it (someone else holds it).
        public static bool Hold(int tries) { return TryOpen(tries, 20); }
        public static bool Release() { return CloseClipboard(); }
    }

    public sealed class Block
    {
        public int Rows, Cols;
        public object[,] Data;          // what is assigned to Range.Formula / Value2
        public byte[,] Kind;            // Cells.K*
        public string[,] Text;          // literal text wanted (Kind == KText)
        public bool[,] Prefix;          // Kind == KText and the string is written with a leading apostrophe (quotePrefix)
        public bool HasFormulas;        // legacy formulas (Range.Formula) in Data
        public int Prefixed;            // strings written with a leading apostrophe up front
        public int Formulas;            // legacy + dynamic
        public int DynFormulas;         // dynamic-array formulas (Range.Formula2), see TakeDynamic
        public List<int> Neutralized = new List<int>();   // flat indexes of cells turned into plain text because they would call out (see Cells.UnsafeFormula)
    }

    /// A vertical run of cells (0-based, inclusive) of one block column that get the same number format.
    public sealed class FormatRun
    {
        public int R1, R2, C;
        public string Fmt;
    }

    public static class Cells
    {
        // KFormula = legacy formula (Range.Formula), KFormula2 = dynamic-array formula (Range.Formula2)
        public const byte KEmpty = 0, KText = 1, KNumber = 2, KBool = 3, KFormula = 4, KFormula2 = 5, KError = 6, KSkip = 7;
        static readonly Regex ErrRx = new Regex(@"^#(N/A|DIV/0!|REF!|NAME\?|NULL!|NUM!|VALUE!|GETTING_DATA|SPILL!|CALC!|FIELD!|BLOCKED!|CONNECT!|BUSY!|UNKNOWN!)$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        static readonly Regex EraRx = new Regex(@"^(明治|大正|昭和|平成|令和|[MTSHRmtshr])\s*(\d|元)", RegexOptions.CultureInvariant);
        static readonly Regex MonRx = new Regex(@"^(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\.?[\s\-/.,]*\d", RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

        /// true when Excel would (or might) turn the typed string into something else. NOT used to decide anything
        /// about the text (the op log says explicitly with "quote" which strings carry a quotePrefix): it only picks
        /// the cells whose number format is remembered before a write, because Excel changes the format of a General
        /// cell when it parses "1/2" into a date and a later rewrite as text would leave the date format behind.
        public static bool Risky(string s)
        {
            if (string.IsNullOrEmpty(s)) return false;
            char c0 = s[0];
            if (c0 == '=' || c0 == '＝' || c0 == '\'' || c0 == '+' || c0 == '-' || c0 == '@') return true;
            string n = s.Normalize(NormalizationForm.FormKC).Trim();
            if (n.Length == 0) return false;
            char f = n[0];
            if ((f >= '0' && f <= '9') || f == '.' || f == '+' || f == '-' || f == '(' || f == '$' || f == '¥' || f == '\\' || f == '€' || f == '£' || f == '−' || f == '=' || f == '@') return true;
            string u = n.ToUpperInvariant();
            if (u == "TRUE" || u == "FALSE") return true;
            if (f == '#' && ErrRx.IsMatch(n)) return true;
            if (EraRx.IsMatch(n) || MonRx.IsMatch(n)) return true;
            return false;
        }

        // unquoted external workbook reference: [Book1.xlsx]Sheet1!A1 (a structured reference such as Table1[Col] is never followed by "!")
        static readonly Regex ExtBareRx = new Regex(@"\[[^\[\]]+\][A-Za-z0-9_.\u0080-\uFFFF]*!", RegexOptions.CultureInvariant);
        // functions that reach outside the workbook when calculated: web requests, real-time-data COM servers, XLM call / exec
        static readonly Regex DangerFnRx = new Regex(@"(?<![A-Za-z0-9_])(?:_xlfn\.)?(?:WEBSERVICE|IMAGE|RTD|CALL|REGISTER(?:\.ID)?|EXEC)\s*\(", RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

        /// true when entering this formula text would make Excel reach outside the workbook by itself (round 5, CSV-injection style
        /// attacks on text that came from untrusted files):
        ///   - DDE: a "|" outside string literals and quoted names  =cmd|' /C calc'!A0  =MSEXCEL|'...'!A1
        ///   - an external workbook reference  [Book.xlsx]Sheet1!A1  '\\host\share\[Book.xlsx]Sheet1'!A1  'C:\dir\[Book.xlsx]S'!A1
        ///   - WEBSERVICE( IMAGE( RTD( CALL( REGISTER( REGISTER.ID( EXEC(   (web requests / COM servers / XLM call-outs)
        /// A formula the user typed that matches is written as plain text instead (and reported: UNSAFE_FORMULA); HYPERLINK is fine
        /// (it only calls out when the user clicks and Excel asks first). Quoting follows Excel's tokenizer: "..." strings and
        /// '...' sheet names are recognised only where a token starts, "" / '' are escapes.
        public static bool UnsafeFormula(string f)
        {
            if (string.IsNullOrEmpty(f)) return false;
            string s = f.Normalize(NormalizationForm.FormKC);   // the full-width forms (＝ ｜ ［) are read like the ASCII ones
            StringBuilder outside = new StringBuilder(s.Length);
            int n = s.Length, i = 0;
            while (i < n)
            {
                char ch = s[i];
                if (ch == '"')
                {
                    i++;
                    while (i < n)
                    {
                        if (s[i] == '"') { if (i + 1 < n && s[i + 1] == '"') { i += 2; continue; } break; }
                        i++;
                    }
                    i++;
                    outside.Append("\"\"");
                    continue;
                }
                if (ch == '\'')
                {
                    StringBuilder name = new StringBuilder();
                    int j = i + 1;
                    while (j < n)
                    {
                        if (s[j] == '\'') { if (j + 1 < n && s[j + 1] == '\'') { name.Append('\''); j += 2; continue; } break; }
                        name.Append(s[j]);
                        j++;
                    }
                    bool bang = j + 1 < n && s[j + 1] == '!';
                    // an internal sheet name never contains [ ] \ / (Excel refuses them): this is a path / UNC / URL / book reference
                    if (bang && name.ToString().IndexOfAny(new char[] { '[', ']', '\\', '/' }) >= 0) return true;
                    i = j + 1;
                    outside.Append("'_'");
                    continue;
                }
                if (ch == '|') return true;
                outside.Append(ch);
                i++;
            }
            string o = outside.ToString();
            return ExtBareRx.IsMatch(o) || DangerFnRx.IsMatch(o);
        }

        /// true when Excel would read this typed text as a formula (= + - @ first: Lotus-style entry is still accepted).
        static bool StartsLikeFormula(string s)
        {
            if (string.IsNullOrEmpty(s)) return false;
            char c = s.Normalize(NormalizationForm.FormKC)[0];
            return c == '=' || c == '+' || c == '-' || c == '@';
        }

        /// Flat indexes (r*Cols+c) of the cells of the block that were turned into plain text by UnsafeFormula.
        public static int[] NeutralizedCells(Block b)
        {
            return b.Neutralized.ToArray();
        }

        static readonly Regex GeneralRx = new Regex("(\"[^\"]*\")|\\bGeneral\\b", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

        /// Number format code for Range.NumberFormat: the English word "General" (CellWrite.numFmt 'General' resets
        /// the format) is replaced by the name this Excel knows (Application.International(xlGeneralFormatName), e.g.
        /// "G/標準" in Japanese Excel, which REJECTS the English word through COM: "Range クラスの NumberFormat
        /// プロパティを設定できません"). Text inside double quotes is left alone.
        public static string Localize(string fmt, string general)
        {
            if (string.IsNullOrEmpty(fmt) || string.IsNullOrEmpty(general)) return fmt;
            if (string.Equals(general, "General", StringComparison.OrdinalIgnoreCase)) return fmt;
            return GeneralRx.Replace(fmt, delegate(Match m) { return m.Groups[1].Success ? m.Value : general; });
        }

        static double ToDouble(object v) { return Convert.ToDouble(v, CultureInfo.InvariantCulture); }

        static bool Truthy(object x)
        {
            if (x == null) return false;
            if (x is bool) return (bool)x;
            string s = Convert.ToString(x, CultureInfo.InvariantCulture);
            return string.Equals(s, "true", StringComparison.OrdinalIgnoreCase) || s == "1";
        }

        /// Number formats requested by the cells of a setCells block ({"numFmt": "..."} on any object cell),
        /// as vertical runs per column (so the runner sets Range.NumberFormat with few calls: "@" before the values,
        /// every other format after them so Excel widens a too-narrow standard-width column like typing does).
        public static List<FormatRun> FormatRuns(object rowsObj)
        {
            List<FormatRun> runs = new List<FormatRun>();
            IList rows = rowsObj as IList;
            if (rows == null) return runs;
            int nR = rows.Count, nC = 0;
            for (int r = 0; r < nR; r++) { IList row = rows[r] as IList; if (row != null && row.Count > nC) nC = row.Count; }
            for (int c = 0; c < nC; c++)
            {
                FormatRun cur = null;
                for (int r = 0; r < nR; r++)
                {
                    IList row = rows[r] as IList;
                    object v = (row != null && c < row.Count) ? row[c] : null;
                    IDictionary<string, object> dict = v as IDictionary<string, object>;
                    string fmt = null;
                    object x;
                    if (dict != null && dict.TryGetValue("numFmt", out x) && x != null)
                    {
                        fmt = Convert.ToString(x, CultureInfo.InvariantCulture);
                        if (fmt.Length == 0 || fmt.Length > 255) fmt = null;
                    }
                    if (fmt != null && cur != null && cur.R2 == r - 1 && string.Equals(cur.Fmt, fmt, StringComparison.Ordinal)) { cur.R2 = r; continue; }
                    if (fmt != null) { cur = new FormatRun(); cur.R1 = r; cur.R2 = r; cur.C = c; cur.Fmt = fmt; runs.Add(cur); }
                    else cur = null;
                }
            }
            return runs;
        }

        static readonly int[] ErrCodes = { 2000, 2007, 2015, 2023, 2029, 2036, 2042, 2043, 2045, 2046, 2047, 2048, 2049, 2050, 2051 };
        static readonly string[] ErrTexts = { "#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#GETTING_DATA", "#SPILL!", "#CONNECT!", "#BLOCKED!", "#UNKNOWN!", "#FIELD!", "#CALC!", "#BUSY!" };

        /// Excel error text of a Value2 error (COM VT_ERROR arrives as Int32 0x800A07xx), null when not an error.
        public static string ErrorText(object v)
        {
            if (!(v is int)) return null;
            int code = ((int)v) & 0xFFFF;
            int i = Array.IndexOf(ErrCodes, code);
            return i >= 0 ? ErrTexts[i] : "#N/A";
        }

        /// Range.Value2 (scalar or 1-based object[,]) -> setCells rows (row lists of cells in the Build encoding):
        /// numbers, booleans, literal strings, {"e": "#N/A"} errors, null for blanks. Used to replace an array /
        /// spill range by its values.
        public static List<object> ValuesToRows(object value2)
        {
            List<object> rows = new List<object>();
            object[,] m = value2 as object[,];
            if (m == null) { List<object> one = new List<object>(); one.Add(ValueCell(value2)); rows.Add(one); return rows; }
            int r0 = m.GetLowerBound(0), c0 = m.GetLowerBound(1);
            for (int r = 0; r < m.GetLength(0); r++)
            {
                List<object> row = new List<object>();
                for (int c = 0; c < m.GetLength(1); c++) row.Add(ValueCell(m[r + r0, c + c0]));
                rows.Add(row);
            }
            return rows;
        }
        static object ValueCell(object v)
        {
            if (v == null || v is DBNull) return null;
            string e = ErrorText(v);
            if (e != null) { Dictionary<string, object> d = new Dictionary<string, object>(); d["e"] = e; return d; }
            if (v is double || v is bool || v is string) return v;
            if (v is float || v is decimal || v is long || v is short) return Convert.ToDouble(v, CultureInfo.InvariantCulture);
            if (v is DateTime) return ((DateTime)v).ToOADate();
            return Convert.ToString(v, CultureInfo.InvariantCulture);
        }

        /// rows: object[] of object[] as produced by JavaScriptSerializer.
        /// Cell encodings: null = clear the value (format kept); string = literal text; number; bool;
        /// {"f": "=..."} legacy formula (UI form, written with Range.Formula); {"f": "=...", "dyn": true}
        /// dynamic-array formula (Range.Formula2); {"e": "#N/A"} error literal; {"s": ..} / {"n": ..} / {"t", "v"};
        /// {"skip": true} = leave the cell untouched. Any object cell may carry "numFmt" (see FormatRuns).
        /// Text and the quote prefix (round 3, no guessing any more): {"t": "s", "v": "001", "quote": true} (also
        /// {"s": "001", "quote": true}) is written with a leading apostrophe = quotePrefix; every other string is written
        /// as it is (Range.Value2 = text) and verified afterwards: what Excel converted anyway (number, date, TRUE,
        /// formula) is rewritten with the apostrophe (the runner warns TEXT_COERCED). A string that itself starts with an
        /// apostrophe always gets one more (Excel eats the first).
        /// textCols[c] = target column is formatted "@" (unused since round 3; kept for the call signature).
        /// mode 0 = explicit (default), 1 = prefix every string, 2 = same as 0 (legacy name "raw").
        public static Block Build(object rowsObj, bool[] textCols, int mode)
        {
            IList rows = (IList)rowsObj;
            int nR = rows.Count, nC = 0;
            for (int r = 0; r < nR; r++) { IList row = rows[r] as IList; if (row != null && row.Count > nC) nC = row.Count; }
            Block b = new Block();
            b.Rows = nR; b.Cols = nC;
            b.Data = new object[nR, nC]; b.Kind = new byte[nR, nC]; b.Text = new string[nR, nC]; b.Prefix = new bool[nR, nC];
            for (int r = 0; r < nR; r++)
            {
                IList row = rows[r] as IList;
                for (int c = 0; c < nC; c++)
                {
                    object v = (row != null && c < row.Count) ? row[c] : null;
                    if (v == null) { b.Data[r, c] = null; b.Kind[r, c] = KEmpty; continue; }
                    if (v is string) { PutText(b, r, c, (string)v, false, mode); continue; }
                    if (v is bool) { b.Data[r, c] = v; b.Kind[r, c] = KBool; continue; }
                    if (v is int || v is long || v is decimal || v is double || v is float)
                    {
                        double d = ToDouble(v);
                        if (double.IsNaN(d) || double.IsInfinity(d)) throw new ArgumentException("bad number at row " + r + " col " + c);
                        b.Data[r, c] = d; b.Kind[r, c] = KNumber; continue;
                    }
                    IDictionary<string, object> dict = v as IDictionary<string, object>;
                    if (dict == null) throw new ArgumentException("unsupported cell value at row " + r + " col " + c + ": " + v.GetType().Name);
                    object x;
                    // "quote": the doc decided this text needs the quotePrefix (CellWrite.quote); Truthy: true / "true" / 1
                    object qv;
                    bool quote = dict.TryGetValue("quote", out qv) && Truthy(qv);
                    if (dict.TryGetValue("skip", out x)) { b.Data[r, c] = null; b.Kind[r, c] = KSkip; }
                    else if (dict.TryGetValue("f", out x) && x != null) {
                        string f = Convert.ToString(x, CultureInfo.InvariantCulture);
                        if (!f.StartsWith("=")) f = "=" + f;
                        object dyn;
                        if (UnsafeFormula(f)) {
                            // DDE / external workbook / web-request formulas are never entered: plain text instead (UNSAFE_FORMULA)
                            PutText(b, r, c, f, true, mode);
                            b.Neutralized.Add(r * nC + c);
                            continue;
                        }
                        b.Data[r, c] = f; b.Formulas++;
                        if (dict.TryGetValue("dyn", out dyn) && Truthy(dyn)) { b.Kind[r, c] = KFormula2; b.DynFormulas++; }
                        else { b.Kind[r, c] = KFormula; b.HasFormulas = true; }
                    } else if (dict.Count == 1 && dict.ContainsKey("numFmt")) {
                        b.Data[r, c] = null; b.Kind[r, c] = KSkip;      // number format only (value untouched)
                    } else if (dict.TryGetValue("s", out x)) {
                        PutText(b, r, c, Convert.ToString(x, CultureInfo.InvariantCulture), quote, mode);
                    } else if (dict.TryGetValue("n", out x)) {
                        b.Data[r, c] = ToDouble(x); b.Kind[r, c] = KNumber;
                    } else if (dict.TryGetValue("e", out x)) {
                        string e = Convert.ToString(x, CultureInfo.InvariantCulture);
                        if (!ErrRx.IsMatch(e)) throw new ArgumentException("bad error literal " + e);
                        b.Data[r, c] = e.ToUpperInvariant(); b.Kind[r, c] = KError;
                    } else if (dict.TryGetValue("t", out x)) {
                        // contracts.ts RunnerCell { t, v } (the formula case {f} is handled above)
                        string t = Convert.ToString(x, CultureInfo.InvariantCulture);
                        object val; dict.TryGetValue("v", out val);
                        if (val == null || t == "z") { b.Data[r, c] = null; b.Kind[r, c] = KEmpty; }
                        else if (t == "n" || t == "d") { double d = ToDouble(val); if (double.IsNaN(d) || double.IsInfinity(d)) throw new ArgumentException("bad number at row " + r + " col " + c); b.Data[r, c] = d; b.Kind[r, c] = KNumber; }
                        else if (t == "b") { b.Data[r, c] = (val is bool) ? val : (object)string.Equals(Convert.ToString(val, CultureInfo.InvariantCulture), "TRUE", StringComparison.OrdinalIgnoreCase); b.Kind[r, c] = KBool; }
                        else if (t == "e") { string e = Convert.ToString(val, CultureInfo.InvariantCulture); if (!ErrRx.IsMatch(e)) throw new ArgumentException("bad error literal " + e); b.Data[r, c] = e.ToUpperInvariant(); b.Kind[r, c] = KError; }
                        else PutText(b, r, c, Convert.ToString(val, CultureInfo.InvariantCulture), quote, mode);
                    } else throw new ArgumentException("unknown cell object keys at row " + r + " col " + c);
                }
            }
            return b;
        }

        /// quote === true -> quotePrefix (verified in Excel: Value2 = "'" + text stores the text with quotePrefix="1",
        /// also in a cell formatted "@"); anything else -> the text as it is, no guessing.
        static void PutText(Block b, int r, int c, string s, bool quote, int mode)
        {
            b.Kind[r, c] = KText; b.Text[r, c] = s;
            if (s.Length == 0) { b.Data[r, c] = null; b.Kind[r, c] = KEmpty; return; }
            bool prefix = quote || s[0] == '\'' || mode == 1;      // a leading apostrophe is always eaten by Excel
            if (!prefix && StartsLikeFormula(s) && UnsafeFormula(s)) { prefix = true; b.Neutralized.Add(r * b.Cols + c); }   // typed text that would become a DDE / external link
            if (prefix) { b.Data[r, c] = "'" + s; b.Prefix[r, c] = true; b.Prefixed++; } else b.Data[r, c] = s;
        }

        /// true when the text cell (flat index r*Cols+c) was written with the leading apostrophe up front.
        public static bool WasPrefixed(Block b, int idx)
        {
            return b.Prefix[idx / b.Cols, idx % b.Cols];
        }

        /// Flat indexes of the text cells written WITHOUT the apostrophe whose text Excel might convert (see Risky):
        /// the runner remembers the number format of those cells (when their column has mixed formats).
        public static int[] RiskyPlain(Block b)
        {
            List<int> list = new List<int>();
            for (int r = 0; r < b.Rows; r++)
                for (int c = 0; c < b.Cols; c++)
                    if (b.Kind[r, c] == KText && !b.Prefix[r, c] && Risky(b.Text[r, c])) list.Add(r * b.Cols + c);
            return list.ToArray();
        }

        /// true when the block contains cells that must be left untouched (so it cannot be written in one call).
        public static bool HasSkips(Block b)
        {
            for (int r = 0; r < b.Rows; r++) for (int c = 0; c < b.Cols; c++) if (b.Kind[r, c] == KSkip) return true;
            return false;
        }

        /// Flat indexes (r*Cols+c) of the formula cells (legacy and dynamic) of the block.
        public static int[] FormulaCells(Block b)
        {
            List<int> list = new List<int>();
            for (int r = 0; r < b.Rows; r++)
                for (int c = 0; c < b.Cols; c++)
                    if (b.Kind[r, c] == KFormula || b.Kind[r, c] == KFormula2) list.Add(r * b.Cols + c);
            return list.ToArray();
        }

        /// Removes the legacy formulas from the block (so the constants can be written in one call) and
        /// returns them as (flat index, formula text).
        public static List<KeyValuePair<int, string>> TakeFormulas(Block b)
        {
            return Take(b, KFormula);
        }

        /// Removes the dynamic-array formulas from the block (written one by one with Range.Formula2 after the
        /// block) and returns them as (flat index, formula text).
        public static List<KeyValuePair<int, string>> TakeDynamic(Block b)
        {
            return Take(b, KFormula2);
        }

        static List<KeyValuePair<int, string>> Take(Block b, byte kind)
        {
            List<KeyValuePair<int, string>> list = new List<KeyValuePair<int, string>>();
            for (int r = 0; r < b.Rows; r++)
                for (int c = 0; c < b.Cols; c++)
                    if (b.Kind[r, c] == kind && b.Data[r, c] is string)
                    {
                        list.Add(new KeyValuePair<int, string>(r * b.Cols + c, (string)b.Data[r, c]));
                        b.Data[r, c] = null;
                    }
            if (kind == KFormula) b.HasFormulas = false;
            return list;
        }

        static object At(object a, int r, int c)
        {
            // r, c are 0-based. Range.Value2 returns a scalar for one cell and a 1-based object[,] otherwise.
            object[,] m = a as object[,];
            if (m == null) return (r == 0 && c == 0) ? a : null;
            int r0 = m.GetLowerBound(0), c0 = m.GetLowerBound(1);
            if (r >= m.GetLength(0) || c >= m.GetLength(1)) return null;
            return m[r + r0, c + c0];
        }
        static int Len(object a, int dim)
        {
            object[,] m = a as object[,];
            if (m == null) return 1;
            return m.GetLength(dim);
        }

        /// Excel stores a line break as LF (verified: "a\r\nb" reads back as "a\nb"): compare without CR.
        static string NormNewlines(string s) { return s.IndexOf('\r') < 0 ? s : s.Replace("\r\n", "\n").Replace('\r', '\n'); }

        /// Compares Range.Value2 read back after the write with what was intended.
        /// Returns flat indexes (r*Cols+c) of literal-text cells that Excel converted.
        public static int[] Verify(object readBack, Block b)
        {
            List<int> bad = new List<int>();
            for (int r = 0; r < b.Rows; r++)
                for (int c = 0; c < b.Cols; c++)
                {
                    if (b.Kind[r, c] != KText) continue;
                    object got = At(readBack, r, c);
                    string gs = got as string;
                    if (gs == null || !string.Equals(NormNewlines(gs), NormNewlines(b.Text[r, c]), StringComparison.Ordinal)) bad.Add(r * b.Cols + c);
                }
            return bad.ToArray();
        }

        static bool Same(object x, object y)
        {
            if (x == null || x is DBNull) return y == null || y is DBNull;
            return object.Equals(x, y);
        }
        static bool Blank(object x) { return x == null || x is DBNull; }

        /// Move verification. `before` = Value2 window (probe x n along the moved axis) read before the op,
        /// `after` = the same window widened by `count` read after it. `alongRows` = the moved axis is the
        /// first dimension (row moves); otherwise the second (column moves). All indexes 1-based.
        /// Returns 0 = the block moved as expected, 1 = Excel inserted blank lines at `to` instead
        /// (degraded "Insert Cut Cells"), 2 = something else (unexpected state).
        /// The window must cover every line the move permutes (n >= max(start+count-1, to-1)); lines beyond the
        /// window are unknown and skipped.
        public static int CheckMove(object before, object after, bool alongRows, int start, int count, int to)
        {
            int dimA = alongRows ? 0 : 1, dimP = alongRows ? 1 : 0;
            int n = Len(before, dimA), probe = Len(before, dimP);
            int nAfter = Len(after, dimA);
            int s1 = start, s2 = start + count - 1;
            bool moved = true, degraded = true;
            for (int i = 1; i <= n + count && i <= nAfter; i++)
            {
                // source line (1-based) expected at position i: -1 = unknown (outside the window), 0 = blank
                int srcMove = i;
                if (to < s1) { if (i >= to && i < to + count) srcMove = s1 + (i - to); else if (i >= to + count && i <= s2) srcMove = i - count; }
                else { if (i >= to - count && i < to) srcMove = s1 + (i - (to - count)); else if (i >= s1 && i < to - count) srcMove = i + count; }
                if (srcMove > n) srcMove = -1;
                int srcIns = (i < to) ? i : (i < to + count) ? 0 : i - count;
                if (srcIns > n) srcIns = -1;
                for (int p = 0; p < probe; p++)
                {
                    object got = alongRows ? At(after, i - 1, p) : At(after, p, i - 1);
                    if (moved && srcMove >= 0)
                    {
                        object want = srcMove == 0 ? null : (alongRows ? At(before, srcMove - 1, p) : At(before, p, srcMove - 1));
                        if (!Same(want, got)) moved = false;
                    }
                    if (degraded && srcIns >= 0)
                    {
                        object want = srcIns == 0 ? null : (alongRows ? At(before, srcIns - 1, p) : At(before, p, srcIns - 1));
                        if (!Same(want, got)) degraded = false;
                    }
                    if (!moved && !degraded) return 2;
                }
            }
            if (moved) return 0;
            return degraded ? 1 : 2;
        }

        /// Diagnostic text for a failed CheckMove: the first position where the moved layout does not match.
        public static string DescribeMove(object before, object after, bool alongRows, int start, int count, int to)
        {
            int dimA = alongRows ? 0 : 1, dimP = alongRows ? 1 : 0;
            int n = Len(before, dimA), probe = Len(before, dimP), nAfter = Len(after, dimA);
            int s1 = start, s2 = start + count - 1;
            for (int i = 1; i <= n + count && i <= nAfter; i++)
            {
                int srcMove = i;
                if (to < s1) { if (i >= to && i < to + count) srcMove = s1 + (i - to); else if (i >= to + count && i <= s2) srcMove = i - count; }
                else { if (i >= to - count && i < to) srcMove = s1 + (i - (to - count)); else if (i >= s1 && i < to - count) srcMove = i + count; }
                if (srcMove > n) continue;
                for (int p = 0; p < probe; p++)
                {
                    object got = alongRows ? At(after, i - 1, p) : At(after, p, i - 1);
                    object want = srcMove == 0 ? null : (alongRows ? At(before, srcMove - 1, p) : At(before, p, srcMove - 1));
                    if (!Same(want, got))
                        return string.Format(CultureInfo.InvariantCulture, "line {0} probe {1}: want [{2}] {3} (from line {4}) got [{5}] {6}; window {7}x{8} -> {9}",
                            i, p + 1, want == null ? "null" : want.GetType().Name, want, srcMove, got == null ? "null" : got.GetType().Name, got, n, probe, nAfter);
                }
            }
            return "no mismatch";
        }

        /// true when both Value2 windows hold the same values.
        public static bool SameWindow(object a, object b)
        {
            int r = Math.Max(Len(a, 0), Len(b, 0)), c = Math.Max(Len(a, 1), Len(b, 1));
            for (int i = 0; i < r; i++) for (int j = 0; j < c; j++) if (!Same(At(a, i, j), At(b, i, j))) return false;
            return true;
        }

        /// true when every value of the window is blank.
        public static bool AllBlank(object a)
        {
            int r = Len(a, 0), c = Len(a, 1);
            for (int i = 0; i < r; i++) for (int j = 0; j < c; j++) if (!Blank(At(a, i, j))) return false;
            return true;
        }

        /// Minimal JSON writer. Unwraps PowerShell PSObject wrappers (New-Object / pipeline output are
        /// wrapped, and JavaScriptSerializer then walks the wrapper's members -> "circular reference").
        /// Output is pure ASCII (non-ASCII as \uXXXX) so stdout decoding can never garble it.
        public static string ToJson(object v)
        {
            StringBuilder sb = new StringBuilder(256);
            WriteJson(sb, v, 0);
            return sb.ToString();
        }
        static void WriteJson(StringBuilder sb, object v, int depth)
        {
            if (v != null && v.GetType().FullName == "System.Management.Automation.PSObject")
            {
                object bo = v.GetType().GetProperty("BaseObject").GetValue(v, null);
                if (bo != null && !object.ReferenceEquals(bo, v)) { WriteJson(sb, bo, depth); return; }
            }
            if (v == null || v is DBNull) { sb.Append("null"); return; }
            if (depth > 32) { WriteString(sb, v.ToString()); return; }
            if (v is string) { WriteString(sb, (string)v); return; }
            if (v is char) { WriteString(sb, v.ToString()); return; }
            if (v is bool) { sb.Append(((bool)v) ? "true" : "false"); return; }
            if (v is int || v is long || v is short || v is byte || v is uint || v is ulong || v is ushort || v is sbyte) { sb.Append(Convert.ToString(v, CultureInfo.InvariantCulture)); return; }
            if (v is double || v is float || v is decimal)
            {
                double d = Convert.ToDouble(v, CultureInfo.InvariantCulture);
                if (double.IsNaN(d) || double.IsInfinity(d)) sb.Append("null"); else sb.Append(d.ToString("R", CultureInfo.InvariantCulture));
                return;
            }
            if (v is DateTime) { WriteString(sb, ((DateTime)v).ToString("o", CultureInfo.InvariantCulture)); return; }
            IDictionary dict = v as IDictionary;
            if (dict != null)
            {
                sb.Append('{'); bool first = true;
                foreach (DictionaryEntry e in dict)
                {
                    if (!first) sb.Append(','); first = false;
                    WriteString(sb, Convert.ToString(e.Key, CultureInfo.InvariantCulture)); sb.Append(':');
                    WriteJson(sb, e.Value, depth + 1);
                }
                sb.Append('}'); return;
            }
            IEnumerable en = v as IEnumerable;
            if (en != null)
            {
                sb.Append('['); bool first = true;
                foreach (object x in en) { if (!first) sb.Append(','); first = false; WriteJson(sb, x, depth + 1); }
                sb.Append(']'); return;
            }
            WriteString(sb, v.ToString());
        }
        static void WriteString(StringBuilder sb, string s)
        {
            sb.Append('"');
            foreach (char ch in s)
            {
                switch (ch)
                {
                    case '"': sb.Append("\\\""); break;
                    case '\\': sb.Append("\\\\"); break;
                    case '\n': sb.Append("\\n"); break;
                    case '\r': sb.Append("\\r"); break;
                    case '\t': sb.Append("\\t"); break;
                    default:
                        if (ch < 0x20 || ch >= 0x7F) sb.Append("\\u").Append(((int)ch).ToString("x4"));
                        else sb.Append(ch);
                        break;
                }
            }
            sb.Append('"');
        }
    }
}
'@

function ConvertTo-Plain($v, [int]$depth = 0) {
  # JSON-safe copy: COM objects / PS wrappers become strings (used only before the helper is loaded).
  if ($null -eq $v) { return $null }
  if ($v -is [psobject]) { $v = $v.psobject.BaseObject }
  if ($v -is [string] -or $v -is [bool] -or $v -is [int] -or $v -is [long] -or $v -is [double] -or $v -is [decimal] -or $v -is [single]) { return $v }
  if ($depth -gt 8) { return [string]$v }
  if ($v -is [System.Collections.IDictionary]) {
    $d = New-Object 'System.Collections.Generic.Dictionary[string,object]'
    foreach ($k in $v.Keys) { $d[[string]$k] = ConvertTo-Plain $v[$k] ($depth + 1) }
    return $d
  }
  if ($v -is [System.Collections.IEnumerable]) {
    $l = New-Object 'System.Collections.Generic.List[object]'
    foreach ($x in $v) { $l.Add((ConvertTo-Plain $x ($depth + 1))) }
    return , $l
  }
  return [string]$v
}
function Emit([System.Collections.IDictionary]$o) {
  $o['t'] = [int]$script:Clock.ElapsedMilliseconds
  if ('XlRunner.Cells' -as [type]) { $s = [XlRunner.Cells]::ToJson($o) }   # PSObject-aware, ASCII-only
  else {
    try { $s = $script:Json.Serialize((ConvertTo-Plain $o)) } catch { $s = '{"type":"error","ok":false,"code":"INTERNAL","message":"emit failed"}' }
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $s.ToCharArray()) { if ([int]$ch -lt 128) { [void]$sb.Append($ch) } else { [void]$sb.Append(('\u{0:x4}' -f [int]$ch)) } }
    $s = $sb.ToString()
  }
  [Console]::Out.WriteLine($s)
  [Console]::Out.Flush()
}
function Warn([string]$code, [string]$msg, $extra) {
  $w = [ordered]@{ type = 'warn'; code = $code; message = $msg }
  if ($extra) { foreach ($k in $extra.Keys) { $w[$k] = $extra[$k] } }
  [void]$script:Warnings.Add(@{ code = $code; message = $msg })
  Emit $w
}
class RunnerError : System.Exception {
  [string]$Code; [int]$Exit
  RunnerError([string]$code, [int]$exit, [string]$msg) : base($msg) { $this.Code = $code; $this.Exit = $exit }
}
function Fail([string]$code, [int]$exit, [string]$msg) { throw [RunnerError]::new($code, $exit, $msg) }
function Find-RunnerError($ex) {
  $e = $ex
  while ($e) { if ($e -is [RunnerError]) { return $e }; $e = $e.InnerException }
  return $null
}
function HResult($err) {
  $e = $err.Exception
  while ($e) { if ($e -is [System.Runtime.InteropServices.COMException]) { return ('0x{0:X8}' -f $e.HResult) }; $e = $e.InnerException }
  if ($err.Exception) { return ('0x{0:X8}' -f $err.Exception.HResult) }
  return $null
}
function Test-Key($d, [string]$key) {
  # JavaScriptSerializer gives Dictionary[string,object] (ContainsKey only; IDictionary.Contains is explicit),
  # PowerShell literals give Hashtable / OrderedDictionary (Contains).
  if ($null -eq $d) { return $false }
  if ($d.PSObject.Methods['ContainsKey']) { return [bool]$d.ContainsKey($key) }
  if ($d -is [System.Collections.IDictionary]) { return [bool]$d.Contains($key) }
  return $false
}
function Get-Opt($d, [string]$key, $default) {
  if ((Test-Key $d $key) -and $null -ne $d[$key]) { return $d[$key] } else { return $default }
}
function Get-Int($d, [string]$key, $default, [int]$min, [int]$max) {
  $v = Get-Opt $d $key $default
  if ($null -eq $v) { Fail 'BAD_JOB' 2 "ジョブの $key がありません" }
  try { $n = [int]$v } catch { Fail 'BAD_JOB' 2 "ジョブの $key が数値ではありません: $v" }
  if ($n -lt $min -or $n -gt $max) { Fail 'BAD_JOB' 2 "ジョブの $key が範囲外です: $n" }
  return $n
}

# ---------------------------------------------------------------------------- helper types
function Import-Helper {
  if ('XlRunner.Cells' -as [type]) { return }
  $sw = [Diagnostics.Stopwatch]::StartNew()
  # Compiled DLL is cached per source hash (compiling takes ~1-2 s with csc of .NET Framework).
  $sha = [Security.Cryptography.SHA1]::Create()
  $hash = -join ($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($HelperSource)) | Select-Object -First 8 | ForEach-Object { $_.ToString('x2') })
  $cacheDir = $env:XLRUNNER_CACHE
  if (-not $cacheDir -and $PSScriptRoot) { $cacheDir = Join-Path $PSScriptRoot '.cache' }
  if (-not $cacheDir) { $cacheDir = Join-Path $env:LOCALAPPDATA 'ExcelMerger\cache' }
  $loaded = $false
  try {
    New-Item -ItemType Directory -Force -Path $cacheDir -ErrorAction Stop | Out-Null
    $dll = Join-Path $cacheDir "XlRunnerHelper-$hash.dll"
    if (Test-Path -LiteralPath $dll) { try { Add-Type -LiteralPath $dll; $loaded = $true } catch { $loaded = $false } }
    if (-not $loaded) {
      $tmp = "$dll.$PID.tmp.dll"
      Add-Type -TypeDefinition $HelperSource -OutputAssembly $tmp -OutputType Library
      try { Move-Item -LiteralPath $tmp -Destination $dll -Force } catch { $dll = $tmp }
      Add-Type -LiteralPath $dll
      $loaded = $true
    }
  } catch { $loaded = $false }
  if (-not $loaded) { Add-Type -TypeDefinition $HelperSource }   # in-memory fallback (cache folder not writable)
  $script:HelperMs = [int]$sw.ElapsedMilliseconds
}

# ---------------------------------------------------------------------------- clipboard guard
# Range.Cut(Destination) and Cut()+Insert go through the Windows clipboard (verified: they empty the user's
# clipboard). The user's clipboard is saved before the first such op and put back afterwards.
function Save-Clipboard {
  if ($null -ne $script:ClipSnapshot) { return }
  $snap = [XlRunner.Clip]::Save(64MB)
  if ($null -eq $snap) {
    $snap = New-Object XlRunner.Clip+Snapshot
    Warn 'CLIPBOARD_BUSY' '他のアプリがクリップボードを使用中のため、現在のクリップボードの内容を保存できませんでした' $null
  }
  $script:ClipSnapshot = $snap
}
function Put-ClipSnapshot([int]$tries) {
  # Restoring right after every Excel cut also takes the clipboard ownership away from the hidden Excel, so
  # clipboard managers (Win+V history) stop asking it to render big delayed formats (observed 2-40 s stalls).
  if (-not [XlRunner.Clip]::Restore($script:ClipSnapshot, $tries)) {
    if ($tries -ge 100) { Warn 'CLIPBOARD_RESTORE_FAILED' 'クリップボードを元に戻せませんでした（他のアプリが使用中）' $null }
    return $false
  }
  return $true
}
function Use-Clipboard { Save-Clipboard; $script:ClipTouched = $true; $script:ClipPending = $true }
function Set-ComProp($obj, [string]$name, $value) {
  # Late-bound IDispatch put. PowerShell binds through the Office PIA and rejects values that are not
  # members of the enum, e.g. CutCopyMode = False (enum XlCutCopyMode has only 1/2).
  [void][System.__ComObject].InvokeMember($name, [Reflection.BindingFlags]::SetProperty, $null, $obj, @($value))
}
function Clear-CutCopy { if ($script:xl) { try { Set-ComProp $script:xl 'CutCopyMode' $false } catch {} } }
function Done-Clipboard {
  Clear-CutCopy
  if ($script:ClipMode -eq 'restoreEachOp') { if (Put-ClipSnapshot 60) { $script:ClipPending = $false } }
}
function Restore-Clipboard {
  # Final safety net (also runs after a failed op).
  if (-not $script:ClipTouched -or $null -eq $script:ClipSnapshot) { return }
  Clear-CutCopy
  if (-not $script:ClipPending) { return }
  if (Put-ClipSnapshot 100) { $script:ClipPending = $false }
}

# ---------------------------------------------------------------------------- Excel lifetime
function Start-Excel {
  $sw = [Diagnostics.Stopwatch]::StartNew()
  try { $app = New-Object -ComObject Excel.Application }
  catch { Fail 'EXCEL_UNAVAILABLE' 3 ("Excel を起動できませんでした（インストールされていないか、起動がブロックされています）: " + $_.Exception.Message) }
  $script:xl = $app
  $p = [uint32]0
  [void][XlRunner.Native]::GetWindowThreadProcessId([IntPtr][int64]$app.Hwnd, [ref]$p)
  $script:XlPid = [int]$p
  $script:ExcelExited = $false
  # Report the PID first: the server kills exactly this PID on timeout / cancel / a stalled shutdown. `started` = the process
  # start time (Unix ms): Windows recycles PIDs, the server only kills a process that was started at that time.
  $startedMs = 0
  try { $startedMs = [long]([DateTimeOffset](Get-Process -Id $script:XlPid).StartTime).ToUnixTimeMilliseconds() } catch {}
  Emit ([ordered]@{ type = 'excel'; pid = $script:XlPid; started = $startedMs })
  $app.Visible = $false
  $app.DisplayAlerts = $false
  $app.ScreenUpdating = $false
  $app.EnableEvents = $false
  $app.AskToUpdateLinks = $false
  $app.AutomationSecurity = 3          # msoAutomationSecurityForceDisable: no macros ever run (set before Open)
  try { $app.Interactive = $false } catch {}
  try { $app.EnableAnimations = $false } catch {}
  try { $app.DisplayStatusBar = $false } catch {}
  $ver = $app.Version; $build = $app.Build
  Emit ([ordered]@{ type = 'progress'; phase = 'excel'; pid = $script:XlPid; version = "$ver"; build = "$build"; startMs = [int]$sw.ElapsedMilliseconds; helperMs = $script:HelperMs })
}

function Stop-Excel {
  $app = $script:xl
  if ($null -eq $app) { return }
  try { foreach ($wb in @($app.Workbooks)) { try { $wb.Close($false) } catch {} } } catch {}
  Clear-CutCopy
  try { Restore-Clipboard } catch {}
  try { $app.Interactive = $true } catch {}
  # Test hook (XLRUNNER_ALLOW_DEBUG=1 + XLRUNNER_HOLD_AFTER_RESULT_MS, see the end of this script): Excel "stuck in its shutdown" - no Quit, the reference is kept
  if ($env:XLRUNNER_ALLOW_DEBUG -eq '1' -and $env:XLRUNNER_HOLD_AFTER_RESULT_MS) { return }
  try { $app.Quit() } catch {}
  $script:Wb = $null
  foreach ($v in @(Get-Variable -Scope Script | Where-Object { $_.Value -is [__ComObject] })) { Set-Variable -Scope Script -Name $v.Name -Value $null }
  $script:xl = $null; $app = $null
  # (the GC / finalizer wait that lets go of Excel's COM objects is Release-Com, run AFTER the final line: it can block for seconds)
  # PowerShell's COM binder caches keep a reference until this process exits, so Excel lingers after Quit() and exits
  # ~0.5-1 s after powershell.exe ends (measured: never within 2 s while we wait). Its shutdown then talks to Microsoft's cloud
  # services and, when that stalls, takes ~60 s - and powershell.exe cannot exit before it. So: no waiting here by default
  # (QuitWaitMs = 0); the final result line goes out at once and the server kills ONLY this PID if Excel is still there
  # a few seconds later. Standalone use: -QuitWaitMs N [-KillOnExit].
  $script:ExcelExited = $true
  if ($script:XlPid -gt 0) {
    $proc = Get-Process -Id $script:XlPid -ErrorAction SilentlyContinue
    if ($proc -and $proc.ProcessName -eq 'EXCEL') {
      $script:ExcelExited = if ([int]$script:QuitWaitMs -gt 0) { $proc.WaitForExit([int]$script:QuitWaitMs) } else { $proc.HasExited }
      if (-not $script:ExcelExited -and $script:KillOnExit) {
        try { $proc.Kill(); [void]$proc.WaitForExit(5000); $script:ExcelExited = $true } catch {}
      }
    }
  }
}

function Release-Com {
  # Lets go of the COM objects. Run AFTER the final line went out: releasing an RCW of a quitting Excel blocks while that Excel winds down
  # (measured: 0.2 s normally, ~2.4 s in 2 of 10 jobs - it used to delay the result by exactly that; nobody waits for it any more).
  [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

# ---------------------------------------------------------------------------- workbook helpers
function Test-SamePath([string]$a, [string]$b) {
  if (-not $a -or -not $b) { return $false }
  try { return [string]::Equals([IO.Path]::GetFullPath($a).TrimEnd('\'), [IO.Path]::GetFullPath($b).TrimEnd('\'), [StringComparison]::OrdinalIgnoreCase) } catch { return $false }
}
function Test-FileLocked([string]$path) {
  if (-not (Test-Path -LiteralPath $path)) { return $false }
  try { $fs = [IO.File]::Open($path, 'Open', 'ReadWrite', 'None'); $fs.Close(); return $false } catch { return $true }
}
function Open-Workbook([string]$path) {
  if (-not (Test-Path -LiteralPath $path)) { Fail 'FILE_NOT_FOUND' 4 "ファイルが見つかりません: $path" }
  $M = $script:M
  try {
    # Open(Filename, UpdateLinks, ReadOnly, Format, Password, WriteResPassword, IgnoreReadOnlyRecommended,
    #      Origin, Delimiter, Editable, Notify, Converter, AddToMru)
    # A dummy Password turns the "enter password" prompt (which would hang a hidden Excel) into an error.
    # WriteResPassword longer than 15 chars makes Open fail with 0x800A03EC even for normal files.
    $wb = $script:xl.Workbooks.Open($path, 0, $false, $M, 'xlrNoPw_1', 'xlrNoPw_1', $true, $M, $M, $false, $false, $M, $false)
  } catch {
    $hr = HResult $_
    $msg = $_.Exception.Message
    if ($msg -match 'パスワード|password') { Fail 'PASSWORD_PROTECTED' 4 'パスワード付きのブックは Excel本体モードで開けません' }
    Fail 'OPEN_FAILED' 4 "Excel でブックを開けませんでした ($hr): $msg"
  }
  if (-not [string]::Equals([IO.Path]::GetFullPath([string]$wb.FullName), [IO.Path]::GetFullPath($path), [StringComparison]::OrdinalIgnoreCase)) {
    Fail 'OPEN_FAILED' 4 'Excel でブックを開けませんでした（同じ名前のブックが開かれています）'
  }
  return $wb
}
function Remove-RunnerTemp {
  # Disk hygiene: the working copy and the saved workbook are big and live on the system drive (%TEMP%): they are removed
  # after EVERY job, success or failure (the server removes its whole job folder as well; this covers the CLI and crashes).
  # Excel may still hold the files for a moment after Quit(): retry a few times.
  foreach ($f in @($script:TempFiles)) {
    for ($i = 0; $i -lt 10 -and (Test-Path -LiteralPath $f); $i++) {
      try { Remove-Item -LiteralPath $f -Force -ErrorAction Stop } catch { Start-Sleep -Milliseconds 200 }
    }
  }
  if ($script:OwnWork) {
    for ($i = 0; $i -lt 10 -and (Test-Path -LiteralPath $script:OwnWork); $i++) {
      try { Remove-Item -LiteralPath $script:OwnWork -Recurse -Force -ErrorAction Stop } catch { Start-Sleep -Milliseconds 200 }
    }
  }
}
function Get-Sheet($ref) {
  $wb = $script:Wb
  if ($null -eq $ref -or [string]$ref -eq '') { Fail 'BAD_JOB' 2 'シート名が指定されていません' }
  foreach ($s in @($wb.Worksheets)) { if ([string]::Equals([string]$s.Name, [string]$ref, [StringComparison]::OrdinalIgnoreCase)) { return $s } }
  Fail 'SHEET_NOT_FOUND' 5 "シート「$ref」が見つかりません"
}
function Col-Letter([int]$n) {
  $s = ''
  while ($n -gt 0) { $m = ($n - 1) % 26; $s = [string][char](65 + $m) + $s; $n = [int][Math]::Floor(($n - 1) / 26) }
  return $s
}
# NOTE: a Range is enumerable, and PowerShell UNROLLS enumerables returned from functions:
# `return $range` of 2 columns returns an array of 2 column ranges, and `.Cut()` on that array is then
# called on each element separately (silently wrong result). Always `return , $range`.
function Get-Lines($ws, [bool]$rows, [int]$start, [int]$count) {
  if ($rows) { return , $ws.Range($ws.Rows.Item($start), $ws.Rows.Item($start + $count - 1)) }
  return , $ws.Range($ws.Columns.Item($start), $ws.Columns.Item($start + $count - 1))
}
function Get-Used($ws) {
  $ur = $ws.UsedRange
  return @{ r1 = [int]$ur.Row; c1 = [int]$ur.Column; r2 = [int]$ur.Row + [int]$ur.Rows.Count - 1; c2 = [int]$ur.Column + [int]$ur.Columns.Count - 1 }
}
function Get-Window($ws, [int]$rows, [int]$cols) {
  if ($rows -lt 1 -or $cols -lt 1) { return $null }
  # unary comma: a returned object[,] would otherwise be unrolled into a flat Object[] (see the Range note above)
  return , $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item($rows, $cols)).Value2
}
function Read-MoveWindow($ws, [bool]$rows, [int]$lines, [int]$probe) {
  # window of `$lines` lines along the moved axis x `$probe` lines across it
  if ($rows) { $w = Get-Window $ws $lines $probe } else { $w = Get-Window $ws $probe $lines }
  return , $w
}
function Count-RefErrors($wb) {
  $n = 0; $samples = New-Object System.Collections.ArrayList
  foreach ($ws in @($wb.Worksheets)) {
    $ur = $ws.UsedRange
    $first = $ur.Find('#REF!', $script:M, -4123, 2)   # What, After, LookIn=xlFormulas, LookAt=xlPart
    if ($null -ne $first) {
      $addr0 = $first.Address($false, $false); $c = $first; $k = 0
      do {
        $n++; if ($samples.Count -lt 20) { [void]$samples.Add($ws.Name + '!' + $c.Address($false, $false)) }
        $c = $ur.FindNext($c); $k++
      } while ($null -ne $c -and $c.Address($false, $false) -ne $addr0 -and $k -lt 5000)
    }
  }
  $names = 0
  foreach ($nm in @($wb.Names)) { try { if ([string]$nm.RefersTo -like '*#REF!*') { $names++ } } catch {} }
  return @{ cells = $n; names = $names; samples = $samples }
}

# ---------------------------------------------------------------------------- pre-checks (refusals)
# Excel refuses moves that split merged cells / array formulas / move cells into or out of a table. The
# clipboard-free move (insert + Cut(Destination) + delete) gets around those checks and silently corrupts the
# book (verified: table columns deleted -> #REF!, foreign columns becoming table columns), so the runner
# checks first and refuses with Excel's reason.
$script:CheckBudget = 0
function Test-Flag($rng, [string]$prop) {
  $v = if ($prop -eq 'merge') { $rng.MergeCells } else { $rng.HasArray }
  if ($v -is [bool]) { return $v }
  return $null   # mixed
}
function Find-Crossing($ws, [bool]$rows, [int]$line, [int]$lo, [int]$hi, [string]$prop) {
  # true when a merged area / array range contains cells on both sides of the boundary just before `line`
  # (between line-1 and line). Bisection over the used part of that line keeps COM calls low.
  if ($script:CheckBudget -le 0) { return $false }
  $script:CheckBudget--
  # plain assignments: `$x = if (...) { $range }` would unroll the Range into its cells (pipeline output)
  if ($rows) { $rng = $ws.Range($ws.Cells.Item($line, $lo), $ws.Cells.Item($line, $hi)) } else { $rng = $ws.Range($ws.Cells.Item($lo, $line), $ws.Cells.Item($hi, $line)) }
  $f = Test-Flag $rng $prop
  if ($f -eq $false) { return $false }
  if ($lo -eq $hi) {
    if ($prop -eq 'merge') { $area = $rng.MergeArea } else { $area = $rng.CurrentArray }
    $first = if ($rows) { [int]$area.Row } else { [int]$area.Column }
    return ($first -lt $line)
  }
  $mid = [int][Math]::Floor(($lo + $hi) / 2)
  if (Find-Crossing $ws $rows $line $lo $mid $prop) { return $true }
  return (Find-Crossing $ws $rows $line ($mid + 1) $hi $prop)
}
function Test-Boundary($ws, [bool]$rows, [int]$line, $used, [string]$prop) {
  $max = if ($rows) { $script:MaxRows } else { $script:MaxCols }
  if ($line -le 1 -or $line -gt $max) { return $false }
  $lo = if ($rows) { $used.c1 } else { $used.r1 }
  $hi = if ($rows) { $used.c2 } else { $used.r2 }
  return (Find-Crossing $ws $rows $line $lo $hi $prop)
}
function Get-Tables($ws) {
  $list = New-Object System.Collections.ArrayList
  foreach ($lo in @($ws.ListObjects)) {
    $r = $lo.Range
    [void]$list.Add(@{ name = [string]$lo.Name; r1 = [int]$r.Row; c1 = [int]$r.Column; r2 = [int]$r.Row + [int]$r.Rows.Count - 1; c2 = [int]$r.Column + [int]$r.Columns.Count - 1 })
  }
  return , $list
}
function Assert-StructAllowed($ws, [string]$kind, [bool]$rows, [int]$start, [int]$count, [int]$to) {
  # kind: move | insert | delete. Returns 'inTable' for a move that stays inside one table, else 'plain'.
  $end = $start + $count - 1
  $mode = 'plain'
  $what = if ($rows) { '行' } else { '列' }
  foreach ($t in (Get-Tables $ws)) {
    $t1 = if ($rows) { $t.r1 } else { $t.c1 }
    $t2 = if ($rows) { $t.r2 } else { $t.c2 }
    if ($kind -ne 'move') { continue }
    $blockOut = ($end -lt $t1) -or ($start -gt $t2)
    $blockIn = ($start -ge $t1) -and ($end -le $t2)
    $toInside = ($to -gt $t1) -and ($to -le $t2)
    if ($rows -and -not $blockOut) { Fail 'REFUSED_TABLE' 5 "テーブル「$($t.name)」の行は Excel本体モードでは移動できません" }
    if ($blockOut) {
      if ($toInside) { Fail 'REFUSED_TABLE' 5 "テーブル「$($t.name)」の外の${what}をテーブルの中へ移動することはできません（Excel の制限）" }
      continue
    }
    if (-not $blockIn) { Fail 'REFUSED_TABLE' 5 "テーブル「$($t.name)」の一部だけにかかる${what}の範囲は移動できません（Excel の制限）" }
    if ($to -lt $t1 -or $to -gt $t2 + 1) { Fail 'REFUSED_TABLE' 5 "テーブル「$($t.name)」の${what}をテーブルの外へ移動することはできません（Excel の制限）" }
    $mode = 'inTable'
  }
  $used = Get-Used $ws
  $script:CheckBudget = 4000
  $noMerge = ($ws.UsedRange.MergeCells -eq $false)
  $noArray = ($ws.UsedRange.HasArray -eq $false)
  if ($kind -eq 'move') {
    if (-not $noMerge) {
      if ((Test-Boundary $ws $rows $start $used 'merge') -or (Test-Boundary $ws $rows ($end + 1) $used 'merge')) {
        Fail 'REFUSED_MERGE' 5 "結合セルの一部だけを移動することはできません（Excel の制限）。結合を解除してからやり直してください"
      }
      if (Test-Boundary $ws $rows $to $used 'merge') {
        Fail 'REFUSED_MERGE' 5 "結合セルの途中へは移動できません（Excel では結合が解除されてしまいます）"
      }
    }
    if (-not $noArray) {
      if ((Test-Boundary $ws $rows $start $used 'array') -or (Test-Boundary $ws $rows ($end + 1) $used 'array') -or (Test-Boundary $ws $rows $to $used 'array')) {
        Fail 'REFUSED_ARRAY' 5 "配列数式の一部を変更することはできません（Excel の制限）"
      }
    }
  } elseif ($kind -eq 'insert') {
    if (-not $noArray -and (Test-Boundary $ws $rows $start $used 'array')) { Fail 'REFUSED_ARRAY' 5 "配列数式の途中に${what}を挿入することはできません（Excel の制限）" }
  } elseif ($kind -eq 'delete') {
    if (-not $noArray -and ((Test-Boundary $ws $rows $start $used 'array') -or (Test-Boundary $ws $rows ($end + 1) $used 'array'))) {
      Fail 'REFUSED_ARRAY' 5 "配列数式の一部だけを削除することはできません（Excel の制限）"
    }
  }
  if ($script:CheckBudget -le 0) { Warn 'CHECK_LIMIT' '結合セル・配列数式の事前チェックを一部省略しました（件数が多いため）' $null }
  return $mode
}

# ---------------------------------------------------------------------------- operations
function Invoke-Excel([scriptblock]$call, [string]$what) {
  try { & $call }
  catch {
    if (Find-RunnerError $_.Exception) { throw }
    $hr = HResult $_
    Fail 'EXCEL_REFUSED' 5 "Excel が「$what」を実行できませんでした ($hr): $($_.Exception.Message)"
  }
}
function Get-ColumnWidths($ws, [int]$n) {
  # Range.ColumnWidth of columns 1..$n (0 = hidden)
  $w = New-Object 'double[]' $n
  for ($i = 1; $i -le $n; $i++) { $w[$i - 1] = [double]$ws.Columns.Item($i).ColumnWidth }
  return , $w
}
function Get-MovedIndex([int]$c, [int]$start, [int]$count, [int]$to) {
  # where line $c ends up after "move lines start..start+count-1 to before line $to" (excelRef mapIndex for a move)
  $end = $start + $count - 1
  if ($to -gt $end) {
    if ($c -ge $start -and $c -le $end) { return $c + ($to - $end - 1) }
    if ($c -gt $end -and $c -lt $to) { return $c - $count }
    return $c
  }
  if ($c -ge $start -and $c -le $end) { return $to + ($c - $start) }
  if ($c -ge $to -and $c -lt $start) { return $c + $count }
  return $c
}
function Repair-ColumnWidths($ws, $before, [int]$start, [int]$count, [int]$to) {
  # Excel's own Cut + Insert Cut Cells of whole columns INSIDE a table moves the cells of the table rows only but
  # shifts the column widths like an entire-column insert without closing the gap again (measured: widths of the columns
  # right of the insertion point drift by one slot, the vacated column is left at its default width). The app (and
  # replayLog) move a column WITH its width and leave every other column alone, so the widths are put there explicitly.
  $n = $before.Length
  $expected = New-Object 'double[]' $n
  for ($c = 1; $c -le $n; $c++) {
    $p = Get-MovedIndex $c $start $count $to
    if ($p -ge 1 -and $p -le $n) { $expected[$p - 1] = $before[$c - 1] }
  }
  $fixed = 0
  for ($p = 1; $p -le $n; $p++) {
    $col = $ws.Columns.Item($p)
    $now = [double]$col.ColumnWidth
    if ([Math]::Abs($now - $expected[$p - 1]) -gt 0.02) {
      try { $col.ColumnWidth = $expected[$p - 1]; $fixed++ } catch { Warn 'WIDTH_RESTORE' ("列幅を元に戻せませんでした: " + $col.Address($false, $false)) $null }
    }
  }
  return $fixed
}
# Every MOVE_VERIFY_FAILED ends with this sentence: the UI falls back to the built-in engine for this code (ApiError.code).
$script:UseBuiltIn = '「内蔵エンジン」で出力してください（画面の出力設定で選べます）。'
function Op-Move($op, [bool]$rows) {
  $ws = Get-Sheet $op['sheet']
  $max = if ($rows) { $script:MaxRows } else { $script:MaxCols }
  $start = Get-Int $op 'start' $null 1 $max
  $count = Get-Int $op 'count' 1 1 $max
  $to = Get-Int $op 'to' $null 1 $max
  $end = $start + $count - 1
  if ($end -gt $max) { Fail 'BAD_JOB' 2 'move の範囲がシートの外です' }
  if ($to -ge $start -and $to -le $end + 1) { return @{ noop = $true } }
  $what = if ($rows) { '行の移動' } else { '列の移動' }
  $mode = Assert-StructAllowed $ws 'move' $rows $start $count $to
  $method = [string](Get-Opt $op 'method' 'clipFree')
  if ($method -eq 'insertCut' -or $method -eq 'auto') { $method = 'clipFree' }
  if ($method -ne 'clipFree' -and $method -ne 'cutInsert') { Fail 'BAD_JOB' 2 "unknown move method $method" }
  # Inside a table the clipboard-free sequence would delete table columns (structured refs -> #REF!), so
  # there Excel's own "Insert Cut Cells" is used, verified and retried.
  if ($mode -eq 'inTable') { $method = 'cutInsert' }

  # Signature window: values of the first probe lines across every line the move permutes (plus a margin
  # of used lines beyond it, so a degraded blank insert shows up as shifted lines).
  $used = Get-Used $ws
  $span = [Math]::Max($end, $to - 1)
  $usedLast = if ($rows) { $used.r2 } else { $used.c2 }
  $n = [Math]::Max($span, [Math]::Min($usedLast, $span + 64))
  $probe = [Math]::Max(1, [Math]::Min(200, $(if ($rows) { $used.c2 } else { $used.r2 })))
  $nAfter = [Math]::Min($n + $count, $max)
  $before = Read-MoveWindow $ws $rows $n $probe
  $w0 = $null
  if ($mode -eq 'inTable' -and -not $rows) { $w0 = Get-ColumnWidths $ws ([int][Math]::Min(2048, [Math]::Max($used.c2, [Math]::Max($end, $to)) + $count + 2)) }

  $sw = [Diagnostics.Stopwatch]::StartNew()
  $attempts = 0
  $widthsFixed = 0
  if ($method -eq 'clipFree') {
    # 1) blank lines at the destination, 2) Range.Cut(Destination) = one call (no window for other apps
    # to break a pending cut), 3) delete the emptied source lines.
    $attempts = 1
    Use-Clipboard
    try {
      Invoke-Excel { [void](Get-Lines $ws $rows $to $count).Insert($(if ($rows) { -4121 } else { -4161 })) } $what
      $src = if ($to -lt $start) { $start + $count } else { $start }
      Invoke-Excel { [void](Get-Lines $ws $rows $src $count).Cut((Get-Lines $ws $rows $to $count)) } $what
      Invoke-Excel { [void](Get-Lines $ws $rows $src $count).Delete($(if ($rows) { -4162 } else { -4159 })) } $what
    } finally { Done-Clipboard }
    $afterWin = Read-MoveWindow $ws $rows $nAfter $probe
    $check = [XlRunner.Cells]::CheckMove($before, $afterWin, $rows, $start, $count, $to)
    if ($check -ne 0) { Emit ([ordered]@{ type = 'debug'; code = 'MOVE_VERIFY'; check = $check; detail = [XlRunner.Cells]::DescribeMove($before, $afterWin, $rows, $start, $count, $to) }) }
    if ($check -ne 0) { Fail 'MOVE_VERIFY_FAILED' 5 "${what}の結果を確認できなかったため中止しました（保存していません）。${script:UseBuiltIn}" }
  } else {
    # Excel's "Cut, then Insert Cut Cells". Another program that writes to the clipboard between the two
    # calls cancels the cut and Insert then adds BLANK lines (verified). Detect it, remove the blank lines
    # (insert + delete of the same lines leaves every reference as it was) and retry.
    # The clipboard is held open during Cut(): Excel then keeps the cut internal (verified: same result, the
    # clipboard does not change) and never becomes the clipboard owner. Otherwise clipboard managers (Ditto,
    # Win+V history) ask the hidden Excel to render its big delayed formats and keep the clipboard open meanwhile
    # (verified with Ditto: it stayed open for the rest of the job, every table move took ~6 s and the user's
    # clipboard could not be restored). Insert Cut Cells still empties the clipboard: Done-Clipboard restores it.
    $ok = $false
    while (-not $ok -and $attempts -lt 6) {
      if ($attempts -gt 0) { Start-Sleep -Milliseconds (Get-Random -Minimum 150 -Maximum 700) }   # back off: the other program is busy with the clipboard
      $attempts++
      Use-Clipboard
      try {
        $held = [XlRunner.Clip]::Hold(50)
        try { Invoke-Excel { [void](Get-Lines $ws $rows $start $count).Cut() } $what }
        finally { if ($held) { [void][XlRunner.Clip]::Release() } }
        $cm = 0; try { $cm = [int]$script:xl.CutCopyMode } catch {}
        if ($cm -eq 2) {
          Invoke-Excel { [void](Get-Lines $ws $rows $to $count).Insert($(if ($rows) { -4121 } else { -4161 })) } $what
          $inserted = $true
        } else { $inserted = $false }
      } finally { Done-Clipboard }
      if (-not $inserted) { continue }
      $afterWin = Read-MoveWindow $ws $rows $nAfter $probe
      $check = [XlRunner.Cells]::CheckMove($before, $afterWin, $rows, $start, $count, $to)
      if ($check -eq 0) { $ok = $true; break }
      if ($check -eq 1) {
        Invoke-Excel { [void](Get-Lines $ws $rows $to $count).Delete($(if ($rows) { -4162 } else { -4159 })) } $what
        $again = Read-MoveWindow $ws $rows $n $probe
        if (-not [XlRunner.Cells]::SameWindow($before, $again)) { Fail 'MOVE_VERIFY_FAILED' 5 "${what}に失敗し、元に戻せなかったため中止しました（保存していません）。${script:UseBuiltIn}" }
        Warn 'MOVE_RETRY' "他のアプリがクリップボードを使用したため${what}をやり直しました" $null
        continue
      }
      Emit ([ordered]@{ type = 'debug'; code = 'MOVE_VERIFY'; check = $check; detail = [XlRunner.Cells]::DescribeMove($before, $afterWin, $rows, $start, $count, $to) })
      if ($mode -eq 'inTable') { Fail 'MOVE_VERIFY_FAILED' 5 "テーブルの中の${what}を Excel本体では正しく行えなかったため中止しました（保存していません）。テーブルの上や下にタイトル・合計などのセルがあると、Excel 本体ではそれらのセルが列と一緒に動かないためです。${script:UseBuiltIn}" }
      Fail 'MOVE_VERIFY_FAILED' 5 "${what}の結果を確認できなかったため中止しました（保存していません）。${script:UseBuiltIn}"
    }
    if (-not $ok) { Fail 'MOVE_VERIFY_FAILED' 5 "他のアプリがクリップボードを使い続けているため${what}ができませんでした（保存していません）。しばらくしてからやり直すか、${script:UseBuiltIn}" }
    if ($null -ne $w0) { $widthsFixed = Repair-ColumnWidths $ws $w0 $start $count $to }
  }
  return @{ method = $method; attempts = $attempts; moveMs = [int]$sw.ElapsedMilliseconds; widthsRepaired = $widthsFixed }
}
function Op-Insert($op, [bool]$rows) {
  $ws = Get-Sheet $op['sheet']
  $max = if ($rows) { $script:MaxRows } else { $script:MaxCols }
  $at = Get-Int $op 'at' $null 1 $max
  $count = Get-Int $op 'count' 1 1 $max
  $what = if ($rows) { '行の挿入' } else { '列の挿入' }
  [void](Assert-StructAllowed $ws 'insert' $rows $at $count 0)
  $ff = [string](Get-Opt $op 'formatFrom' $(if ($rows) { 'above' } else { 'left' }))
  $origin = if ($ff -eq 'right' -or $ff -eq 'below') { 1 } else { 0 }   # xlFormatFromRightOrBelow / LeftOrAbove
  Invoke-Excel { [void](Get-Lines $ws $rows $at $count).Insert($(if ($rows) { -4121 } else { -4161 }), $origin) } $what
  if ($ff -eq 'none') { Invoke-Excel { [void](Get-Lines $ws $rows $at $count).ClearFormats() } $what }
  return @{}
}
function Op-Delete($op, [bool]$rows) {
  $ws = Get-Sheet $op['sheet']
  $max = if ($rows) { $script:MaxRows } else { $script:MaxCols }
  $at = Get-Int $op 'at' $null 1 $max
  $count = Get-Int $op 'count' 1 1 $max
  $what = if ($rows) { '行の削除' } else { '列の削除' }
  [void](Assert-StructAllowed $ws 'delete' $rows $at $count 0)
  Invoke-Excel { [void](Get-Lines $ws $rows $at $count).Delete($(if ($rows) { -4162 } else { -4159 })) } $what
  return @{}
}
function Op-MoveCols($op) { return (Op-Move $op $false) }
function Op-MoveRows($op) { return (Op-Move $op $true) }
function Op-InsertCols($op) { return (Op-Insert $op $false) }
function Op-InsertRows($op) { return (Op-Insert $op $true) }
function Op-DeleteCols($op) { return (Op-Delete $op $false) }
function Op-DeleteRows($op) { return (Op-Delete $op $true) }

function Write-Block($rng, $b) {
  # Legacy formulas and constants in one call. Range.Formula keeps the legacy (implicit intersection) semantics
  # a formula has in the file (toUiForm(<f>) == what Range.Formula shows for a normal cell); dynamic-array
  # formulas are written separately with Range.Formula2 (see Put-Formula).
  if ($b.HasFormulas) { $rng.Formula = $b.Data } else { $rng.Value2 = $b.Data }
}
$script:BadFormulas = 0
function Put-Formula($cell, [string]$f, [bool]$dyn) {
  # One formula; an unparsable one is kept as text and reported.
  try { if ($dyn) { $cell.Formula2 = $f } else { $cell.Formula = $f }; return $true }
  catch {
    $cell.Value2 = "'" + $f; $script:BadFormulas++
    if ($script:BadFormulas -le 20) { Warn 'BAD_FORMULA' ("Excel が解釈できない数式を文字列として書き込みました: " + $cell.Address($false, $false) + ' ' + $f) $null }
    return $false
  }
}
function Read-ColumnFormats($rng, [int]$nC) {
  # Per-column number format (1 call per column; $null when mixed). "@" (text) columns never coerce strings.
  $colFmt = New-Object 'string[]' $nC; $textCols = New-Object 'bool[]' $nC
  for ($c = 0; $c -lt $nC; $c++) {
    $f = $rng.Columns.Item($c + 1).NumberFormat
    if ($f -is [string]) { $colFmt[$c] = $f; $textCols[$c] = ($f -eq '@') }
  }
  return @{ fmt = $colFmt; text = $textCols }
}
$script:CoercedNotes = 0
$script:NeutralizedNotes = 0
function Note-Coerced($cell, [string]$text) {
  # A string that was NOT marked with "quote" but that Excel converted anyway (number, date, TRUE, formula ...) was
  # rewritten as text with the apostrophe. The doc marks such strings up front, so this is the safety net (user edits).
  $script:CoercedNotes++
  if ($script:CoercedNotes -le 20) {
    $t = if ($text.Length -gt 30) { $text.Substring(0, 30) + '...' } else { $text }
    Warn 'TEXT_COERCED' ("セル " + $cell.Address($false, $false) + " の「" + $t + "」は Excel が数値・日付・数式などに変換してしまうため、文字列として書き直しました（先頭に ' が付きます）") $null
  } elseif ($script:CoercedNotes -eq 21) {
    Warn 'TEXT_COERCED' '同様に文字列として書き直したセルが他にもあります' $null
  }
}
function Write-ConstantsSlow($rng, $b, [int]$nR, [int]$nC) {
  # Cell by cell (after the one-call block write was refused by Excel: a plain string such as "=(" is parsed as a
  # formula and cannot be entered). A text Excel refuses is written with the apostrophe (TEXT_COERCED).
  for ($r = 0; $r -lt $nR; $r++) {
    for ($c = 0; $c -lt $nC; $c++) {
      $cell = $rng.Cells.Item($r + 1, $c + 1)
      $v = $b.Data[$r, $c]
      try { $cell.Value2 = $v }
      catch {
        if ($b.Kind[$r, $c] -eq [XlRunner.Cells]::KText -and -not $b.Prefix[$r, $c]) {
          $cell.Value2 = "'" + [string]$b.Text[$r, $c]
          Note-Coerced $cell ([string]$b.Text[$r, $c])
        } else { Fail 'EXCEL_REFUSED' 5 "セルに値を書き込めませんでした ($(HResult $_)): $($_.Exception.Message)" }
      }
    }
  }
}
function Write-Cells($rng, $rows, [int]$mode) {
  # Writes a rectangular block of cells (setCells encoding, see XlRunner.Cells.Build) into $rng (same size).
  $nR = [int]$rng.Rows.Count; $nC = [int]$rng.Columns.Count
  $cf = Read-ColumnFormats $rng $nC
  try { $b = [XlRunner.Cells]::Build($rows, $cf.text, $mode) } catch { Fail 'BAD_JOB' 2 "setCells のデータが不正です: $($_.Exception.InnerException.Message)" }
  # Round 5: a formula / text that would call out (DDE =cmd|'/C calc'!A0, external workbook links, WEBSERVICE ...) was turned into plain text
  foreach ($idx in [XlRunner.Cells]::NeutralizedCells($b)) {
    $script:NeutralizedNotes++
    if ($script:NeutralizedNotes -le 20) {
      # (PowerShell variable names are case-insensitive: $nr / $nc would overwrite this function's $nR / $nC)
      $wr = [int][Math]::Floor($idx / $b.Cols); $wc = [int]($idx % $b.Cols)
      Warn 'UNSAFE_FORMULA' ("セル " + $rng.Cells.Item($wr + 1, $wc + 1).Address($false, $false) + " の内容は、外部のプログラム・ファイル・Web を呼び出す数式の形（DDE・外部ブックの参照・WEBSERVICE など）のため、数式にせず文字列として書き込みました") $null
    } elseif ($script:NeutralizedNotes -eq 21) {
      Warn 'UNSAFE_FORMULA' '同様に文字列として書き込んだセルが他にもあります' $null
    }
  }
  $bad0 = $script:BadFormulas
  # Entering a formula makes Excel copy the number format of a referenced cell into a General cell
  # (=SUM(M3:M5) over dates becomes a date; verified). The op log keeps the cell's style unless the write carries
  # numFmt (set just before by Op-SetCells), so the format each formula cell has now is put back afterwards.
  $fCells = [XlRunner.Cells]::FormulaCells($b)
  $fmtBefore = @{}
  foreach ($idx in $fCells) {
    $c = [int]($idx % $nC)
    if ($null -eq $cf.fmt[$c]) { $fmtBefore[$idx] = $rng.Cells.Item([int][Math]::Floor($idx / $nC) + 1, $c + 1).NumberFormat }
  }
  # Excel changes the number format of a General cell when it parses a string ("1/2" -> m"月"d"日"; verified) and a rewrite
  # as text would keep it: remember the format of the plain strings it might convert (columns with mixed formats only;
  # a column with one format is put back as a whole).
  $riskFmt = @{}
  $riskBudget = 2000   # one COM call per cell: a huge run of such strings in a mixed column is not worth it
  foreach ($idx in [XlRunner.Cells]::RiskyPlain($b)) {
    $c = [int]($idx % $nC)
    if ($null -eq $cf.fmt[$c]) {
      if ($riskBudget -le 0) { break }
      $riskBudget--
      $riskFmt[$idx] = $rng.Cells.Item([int][Math]::Floor($idx / $nC) + 1, $c + 1).NumberFormat
    }
  }
  if ([XlRunner.Cells]::HasSkips($b)) {
    # Cells to leave untouched: write cell by cell (rare; comJob sends dense blocks).
    for ($r = 0; $r -lt $nR; $r++) {
      for ($c = 0; $c -lt $nC; $c++) {
        $k = $b.Kind[$r, $c]
        if ($k -eq [XlRunner.Cells]::KSkip) { continue }
        $cell = $rng.Cells.Item($r + 1, $c + 1)
        if ($k -eq [XlRunner.Cells]::KFormula -or $k -eq [XlRunner.Cells]::KFormula2) { [void](Put-Formula $cell ([string]$b.Data[$r, $c]) ($k -eq [XlRunner.Cells]::KFormula2)) }
        else {
          try { $cell.Value2 = $b.Data[$r, $c] }
          catch {
            if ($k -eq [XlRunner.Cells]::KText -and -not $b.Prefix[$r, $c]) { $cell.Value2 = "'" + [string]$b.Text[$r, $c]; Note-Coerced $cell ([string]$b.Text[$r, $c]) }
            else { Fail 'EXCEL_REFUSED' 5 "セルに値を書き込めませんでした ($(HResult $_)): $($_.Exception.Message)" }
          }
        }
      }
    }
  } else {
    $dyn = [XlRunner.Cells]::TakeDynamic($b)
    try { Write-Block $rng $b }
    catch {
      # One unparsable formula makes the whole array assignment fail (0x800A03EC): write the constants first,
      # then the formulas one by one; the broken ones are kept as text and reported. A plain string that Excel
      # cannot enter ("=(") makes the constants fail too: then cell by cell, such a text gets the apostrophe.
      $fIdx = $null
      if ($b.HasFormulas) { $fIdx = [XlRunner.Cells]::TakeFormulas($b) }
      try { Write-Block $rng $b } catch { Write-ConstantsSlow $rng $b $nR $nC }
      if ($null -ne $fIdx) {
        foreach ($pair in $fIdx) {
          $r = [int][Math]::Floor($pair.Key / $nC); $c = [int]($pair.Key % $nC)
          [void](Put-Formula $rng.Cells.Item($r + 1, $c + 1) $pair.Value $false)
        }
      }
    }
    foreach ($pair in $dyn) {
      $r = [int][Math]::Floor($pair.Key / $nC); $c = [int]($pair.Key % $nC)
      [void](Put-Formula $rng.Cells.Item($r + 1, $c + 1) $pair.Value $true)
    }
  }
  $fmtKept = 0
  if ($fCells.Length -gt 0) {
    $doneCols = @{}
    foreach ($idx in $fCells) {
      $r = [int][Math]::Floor($idx / $nC); $c = [int]($idx % $nC)
      if ($null -ne $cf.fmt[$c]) {
        # the whole block column had one format before the write: one read, one write at most
        if ($doneCols.ContainsKey($c)) { continue }
        $doneCols[$c] = $true
        $part = $rng.Columns.Item($c + 1)
        $now = $part.NumberFormat   # $null when the column became mixed
        if ($now -isnot [string] -or -not [string]::Equals($now, [string]$cf.fmt[$c], [StringComparison]::Ordinal)) {
          try { $part.NumberFormat = $cf.fmt[$c]; $fmtKept++ } catch {}
        }
      } elseif ($fmtBefore.ContainsKey($idx) -and $null -ne $fmtBefore[$idx]) {
        $cell = $rng.Cells.Item($r + 1, $c + 1)
        if (-not [string]::Equals([string]$cell.NumberFormat, [string]$fmtBefore[$idx], [StringComparison]::Ordinal)) {
          try { $cell.NumberFormat = $fmtBefore[$idx]; $fmtKept++ } catch {}
        }
      }
    }
  }
  # Verify the strings: Excel parses Value2 = "001" / "1/2" / "TRUE" / "=abc" like typed input. Only the strings the op log
  # marked with "quote" carry the apostrophe up front (no guessing); what Excel converted anyway is rewritten as text.
  $bad = [XlRunner.Cells]::Verify($rng.Value2, $b)
  $repaired = 0
  foreach ($idx in $bad) {
    $r = [int][Math]::Floor($idx / $nC); $c = [int]($idx % $nC)
    $cell = $rng.Cells.Item($r + 1, $c + 1)
    if ($null -ne $cf.fmt[$c]) { try { $cell.NumberFormat = $cf.fmt[$c] } catch {} }
    elseif ($riskFmt.ContainsKey($idx) -and $null -ne $riskFmt[$idx]) { try { $cell.NumberFormat = $riskFmt[$idx] } catch {} }
    $cell.Value2 = "'" + [string]$b.Text[$r, $c]
    if (-not [XlRunner.Cells]::WasPrefixed($b, $idx)) { Note-Coerced $cell ([string]$b.Text[$r, $c]) }
    $repaired++
  }
  return @{ formulas = $b.Formulas; dyn = $b.DynFormulas; prefixed = $b.Prefixed; repaired = $repaired; badFormulas = ($script:BadFormulas - $bad0); formatsKept = $fmtKept }
}
function Test-AnyFlag($v) {
  # Range.HasArray / HasSpill of a multi-cell range: $true (all), $false (none), DBNull / $null (mixed)
  return -not ($v -is [bool] -and $v -eq $false)
}
function Find-FlaggedCell($rng, [string]$prop) {
  # First cell of $rng whose HasArray / HasSpill is true (null when none). Bisection over the rows / columns:
  # one COM call per probed piece instead of one per cell (a toValues run is often a whole column of 10,000 cells).
  $v = $null
  try { $v = $rng.$prop } catch { return $null }
  if ($v -is [bool] -and $v -eq $false) { return $null }
  $nR = [int]$rng.Rows.Count; $nC = [int]$rng.Columns.Count
  if ($nR -eq 1 -and $nC -eq 1) { if ($v -eq $true) { return , $rng } else { return $null } }
  if ($nR -ge $nC) {
    $h = [int][Math]::Floor($nR / 2)
    $a = $rng.Resize($h, $nC); $b = $rng.Offset($h, 0).Resize($nR - $h, $nC)
  } else {
    $h = [int][Math]::Floor($nC / 2)
    $a = $rng.Resize($nR, $h); $b = $rng.Offset(0, $h).Resize($nR, $nC - $h)
  }
  $x = Find-FlaggedCell $a $prop
  if ($null -ne $x) { return , $x }
  $y = Find-FlaggedCell $b $prop
  if ($null -ne $y) { return , $y }
  return $null
}
function Convert-ArraysToValues($ws, $rng) {
  # reason 'toValues' over an array formula / spill: Excel cannot change part of an array, and replacing only
  # the master of a spill would blank the spilled cells. Like「値の貼り付け」over the whole array: read its
  # values, clear it, write them back (then the op's own writes follow). Returns the number of arrays.
  $n = 0
  foreach ($prop in @('HasArray', 'HasSpill')) {   # HasSpill: Excel 365 / 2021 only (Find-FlaggedCell tolerates its absence)
    $seen = @{}
    for ($guard = 0; $guard -lt 100000; $guard++) {
      $cell = Find-FlaggedCell $rng $prop
      if ($null -eq $cell) { break }
      if ($prop -eq 'HasArray') { $anchor = $cell.CurrentArray; $area = $anchor }
      else { $anchor = $cell.SpillParent; $area = $anchor.SpillingToRange }
      $key = [string]$area.Address($false, $false)
      if ($seen.ContainsKey($key)) { break }   # conversion did not remove it: never loop
      $seen[$key] = $true
      $vals = [XlRunner.Cells]::ValuesToRows($area.Value2)
      Invoke-Excel { [void]$anchor.ClearContents() } '配列数式の値への変換'
      [void](Write-Cells $area $vals 0)
      $n++
    }
  }
  return $n
}
$script:GeneralName = $null
function Get-GeneralName {
  # Name of the General format in THIS Excel (Application.International(xlGeneralFormatName = 26)): "G/標準" in Japanese
  # Excel, which rejects the English word "General" through COM (measured: every NumberFormat = 'General' fails there).
  if ($null -ne $script:GeneralName) { return $script:GeneralName }
  $g = 'General'
  try { $x = [string]$script:xl.International(26); if ($x) { $g = $x } } catch {}
  $script:GeneralName = $g
  return $g
}
function Set-RunFormat($ws, $rng, $run) {
  # Range.NumberFormat over one vertical run of a setCells block; returns 1 when set, 0 when Excel refused the code.
  # CellWrite.numFmt 'General' (or a code containing the word) is translated to this Excel's name for it first.
  $part = $ws.Range($rng.Cells.Item($run.R1 + 1, $run.C + 1), $rng.Cells.Item($run.R2 + 1, $run.C + 1))
  $fmt = [XlRunner.Cells]::Localize([string]$run.Fmt, (Get-GeneralName))
  try { $part.NumberFormat = $fmt; return 1 }
  catch {
    if (-not [string]::Equals($fmt, [string]$run.Fmt, [StringComparison]::Ordinal)) { try { $part.NumberFormat = [string]$run.Fmt; return 1 } catch {} }
    Warn 'BAD_NUMFMT' ("表示形式を設定できませんでした: " + $part.Address($false, $false) + ' ' + $run.Fmt) $null
    return 0
  }
}
function Op-SetCells($op) {
  $ws = Get-Sheet $op['sheet']
  $rows = $op['rows']
  if ($null -eq $rows -or $rows.Count -eq 0) { return @{ cells = 0 } }
  $mode = switch ([string](Get-Opt $op 'textMode' 'auto')) { 'auto' { 0 } 'prefixAll' { 1 } 'raw' { 2 } default { 0 } }
  $nC = 0; foreach ($r in $rows) { if ($null -ne $r -and $r.Count -gt $nC) { $nC = $r.Count } }
  $nR = $rows.Count
  if ($nC -eq 0) { return @{ cells = 0 } }
  # only a plain cell address: Worksheet.Range also accepts names and external references ('\\host\share\[x.xlsx]S'!A1 would make Excel reach for the network)
  if ([string]$op['at'] -notmatch '^\$?[A-Za-z]{1,3}\$?[0-9]{1,7}$') { Fail 'BAD_JOB' 2 "setCells の位置が不正です: $($op['at'])" }
  try { $rng = $ws.Range([string]$op['at']).Resize($nR, $nC) } catch { Fail 'BAD_JOB' 2 "setCells の位置が不正です: $($op['at'])" }
  $sw = [Diagnostics.Stopwatch]::StartNew()
  $arrays = 0
  if ([bool](Get-Opt $op 'toValues' $false)) { $arrays = Convert-ArraysToValues $ws $rng }
  # Number formats (CellWrite.numFmt: a date / % / currency typed into a General cell changes its format too).
  #  - "@" (text) runs are set BEFORE the values, so strings stay literal text.
  #  - every other format is set AFTER the values: Excel widens a column of standard width for a date / 1,234,567 that
  #    no longer fits only when the format is applied to a cell that already holds the number (like typing it;
  #    tests/golden/RESULTS.md: "setting the format before the value does not widen"), and replayLog emulates exactly
  #    that widening. Set before the value the output would show "########" in a standard-width column.
  $fmts = 0
  $runs = @([XlRunner.Cells]::FormatRuns($rows))
  foreach ($run in $runs) { if ($run.Fmt -eq '@') { $fmts += (Set-RunFormat $ws $rng $run) } }
  $res = Write-Cells $rng $rows $mode
  foreach ($run in $runs) { if ($run.Fmt -ne '@') { $fmts += (Set-RunFormat $ws $rng $run) } }
  $res['cells'] = $nR * $nC; $res['numFmts'] = $fmts; $res['arraysToValues'] = $arrays; $res['writeMs'] = [int]$sw.ElapsedMilliseconds
  return $res
}
$script:WidthPad = $null
function Get-WidthPadding($ws, [int]$col) {
  # The <col width> stored in the file = ColumnWidth (characters) + padding, padding = (2 * ceil(mdw / 4) + 1) px
  # / mdw with mdw = max digit width of the workbook's Normal font in pixels (0.625 for 11pt 游ゴシック at 96 dpi,
  # 5/7 for 11pt Calibri). Measured in Excel itself (DPI-independent ratio): Width(points) = k * (chars * mdw + pad)
  # for integer character counts, so pad/mdw = (W10 - 10 * d) / d with d = (W20 - W10) / 10. Measured once per job
  # on the column that is about to get its final width anyway.
  if ($null -ne $script:WidthPad) { return $script:WidthPad }
  $pad = 0.625
  try {
    $line = Get-Lines $ws $false $col 1
    $line.ColumnWidth = 10; $w10 = [double]$line.Width
    $line.ColumnWidth = 20; $w20 = [double]$line.Width
    $d = ($w20 - $w10) / 10
    if ($d -gt 0) {
      $p = ($w10 - 10 * $d) / $d
      if ($p -gt 0.05 -and $p -lt 3) { $pad = [Math]::Round($p, 6) }
      else { Emit ([ordered]@{ type = 'debug'; code = 'WIDTH_PAD'; w10 = $w10; w20 = $w20 }) }
    }
  } catch { Emit ([ordered]@{ type = 'debug'; code = 'WIDTH_PAD'; message = [string]$_.Exception.Message }) }
  $script:WidthPad = $pad
  return $pad
}
function Op-SetColumnWidth($op) {
  $ws = Get-Sheet $op['sheet']
  $col = Get-Int $op 'col' $null 1 $script:MaxCols
  $count = Get-Int $op 'count' 1 1 $script:MaxCols
  $w = [double](Get-Opt $op 'width' 8.43)
  $unit = [string](Get-Opt $op 'unit' 'chars')
  if ($unit -ne 'chars' -and $unit -ne 'stored') { Fail 'BAD_JOB' 2 "列幅の単位が不正です: $unit" }
  if ($w -lt 0 -or $w -gt 256 -or ($unit -eq 'chars' -and $w -gt 255)) { Fail 'BAD_JOB' 2 "列幅が範囲外です: $w" }
  $pad = $null
  if ($unit -eq 'stored') {
    # LoggedOp colWidth = the value stored in <col width>; Range.ColumnWidth excludes the padding
    $pad = Get-WidthPadding $ws $col
    # double literals: [Math]::Max(0, 19.375) binds to Max(int, int) in PowerShell and returns 19
    $w = [Math]::Min([double]255, [Math]::Max([double]0, [double]($w - $pad)))
  }
  Invoke-Excel { (Get-Lines $ws $false $col $count).ColumnWidth = $w } '列幅の変更'
  return @{ columnWidth = $w; padding = $pad }
}
function Copy-Fill($src, $dst) {
  # Fill (Interior) of the top-left cell -> the other cells (one Range). Theme colours stay theme colours (+ tint), a
  # patterned fill keeps its pattern colour. No fill on the source: nothing to copy. Gradients are not copied.
  $si = $src.Interior
  $pat = [int]$si.Pattern
  if ($pat -eq -4142 -or $pat -eq -4135 -or $pat -eq -4136) { return $false }   # xlNone / xlLinearGradient / xlRectangularGradient
  $di = $dst.Interior
  $di.Pattern = $pat
  $theme = 0
  try { $theme = [int]$si.ThemeColor } catch {}
  if ($theme -gt 0) { $di.ThemeColor = $theme; $di.TintAndShade = [double]$si.TintAndShade } else { $di.Color = $si.Color }
  if ($pat -ne 1) {   # patterned: the pattern (foreground) colour as well
    $pth = 0
    try { $pth = [int]$si.PatternThemeColor } catch {}
    if ($pth -gt 0) { $di.PatternThemeColor = $pth; $di.PatternTintAndShade = [double]$si.PatternTintAndShade } else { $di.PatternColor = $si.PatternColor }
  }
  return $true
}
function Copy-Border($src, $dst) {
  # One border side (Border objects): line style, weight and colour (theme / automatic / RGB). A side the source does not
  # draw is not copied (an outline Excel itself authored on the other cells stays).
  $ls = [int]$src.LineStyle
  if ($ls -eq -4142) { return $false }   # xlNone
  $dst.LineStyle = $ls
  try { $dst.Weight = [int]$src.Weight } catch {}
  $th = $null
  try { $th = $src.ThemeColor } catch {}
  if ($th -is [int] -and $th -gt 0) { $dst.ThemeColor = $th; try { $dst.TintAndShade = [double]$src.TintAndShade } catch {} }
  elseif ([int]$src.ColorIndex -eq -4105) { $dst.ColorIndex = -4105 }   # xlColorIndexAutomatic
  else { $dst.Color = $src.Color }
  return $true
}
function Copy-MergeFormats($ws, $area) {
  # LoggedOp 'unmerge' (docs: src/core/types.ts, same rule as the JS patcher's unmerge): BEFORE Range.UnMerge every other cell of
  # the merged area receives the top-left cell's fill, and each border side the top-left cell HAS goes to the cells that lie on
  # that edge of the area (top: first row, bottom: last row, left: first column, right: last column; for a one-row band: top and
  # bottom on every cell, left on the first, right on the last cell). A side the top-left cell does not have leaves the other
  # cell's own side alone. A title band styled on the top-left cell only (openpyxl style: measured, Excel leaves the other cells
  # unstyled and UnMerge does not copy anything) would otherwise lose its fill and lines. Excel shares one border between two
  # cells: the neighbours outside the area get the same line on their side (no visual difference).
  $r1 = [int]$area.Row; $c1 = [int]$area.Column
  $r2 = $r1 + [int]$area.Rows.Count - 1; $c2 = $c1 + [int]$area.Columns.Count - 1
  if ($r1 -eq $r2 -and $c1 -eq $c2) { return 0 }
  $tl = $ws.Cells.Item($r1, $c1)
  $tlTop = $tl.Borders.Item(8); $tlBottom = $tl.Borders.Item(9); $tlLeft = $tl.Borders.Item(7); $tlRight = $tl.Borders.Item(10)
  $copied = 0
  for ($r = $r1; $r -le $r2; $r++) {
    $from = if ($r -eq $r1) { $c1 + 1 } else { $c1 }   # the top-left cell itself is the source
    if ($from -gt $c2) { continue }
    $row = $ws.Range($ws.Cells.Item($r, $from), $ws.Cells.Item($r, $c2))
    if (Copy-Fill $tl $row) { $copied++ }
    # xlEdgeTop / xlEdgeBottom of a one-row range = the top / bottom of every cell in it (measured)
    if ($r -eq $r1) { if (Copy-Border $tlTop $row.Borders.Item(8)) { $copied++ } }
    if ($r -eq $r2) { if (Copy-Border $tlBottom $row.Borders.Item(9)) { $copied++ } }
    if ($from -eq $c1) { if (Copy-Border $tlLeft $ws.Cells.Item($r, $c1).Borders.Item(7)) { $copied++ } }     # first column (not the top-left cell itself)
    if ($c2 -gt $c1 -or $r -ne $r1) { if (Copy-Border $tlRight $ws.Cells.Item($r, $c2).Borders.Item(10)) { $copied++ } }   # last column
  }
  return $copied
}
function Op-Unmerge($op) {
  # Unmerges merged areas (A1 refs as of the time of the op). Fill + borders of the top-left cell are copied to the other
  # cells first (Copy-MergeFormats); centerAcross: 「選択範囲内で中央」 (HorizontalAlignment = 7, xlCenterAcrossSelection)
  # over each former area so a title keeps its look.
  $ws = Get-Sheet $op['sheet']
  $refs = @(Get-Opt $op 'refs' @())
  $center = [bool](Get-Opt $op 'centerAcross' $false)
  foreach ($ref in $refs) {
    if ([string]$ref -notmatch '^\$?[A-Za-z]{1,3}\$?[0-9]{1,7}(:\$?[A-Za-z]{1,3}\$?[0-9]{1,7})?$') { Fail 'BAD_JOB' 2 "結合を解除する範囲が不正です: $ref" }
  }
  $n = 0; $copied = 0
  foreach ($ref in $refs) {
    $area = $ws.Range([string]$ref)
    if ($area.MergeCells -eq $true) {
      try { $copied += (Copy-MergeFormats $ws $area) }
      catch { Warn 'UNMERGE_FORMATS' ("結合を解除するセルの塗りつぶし・罫線をコピーできませんでした: " + $area.Address($false, $false) + ' ' + $_.Exception.Message) $null }
    }
    Invoke-Excel { [void]$area.UnMerge() } 'セル結合の解除'
    if ($center) { Invoke-Excel { $area.HorizontalAlignment = 7 } '選択範囲内で中央' }
    $n++
  }
  return @{ unmerged = $n; formatCopies = $copied }
}
function Op-DebugWait($op) {
  # Test hook (XLRUNNER_ALLOW_DEBUG=1 only): blocks inside Excel like a hidden modal dialog would.
  [void]$script:xl.Wait([DateTime]::Now.AddSeconds([double](Get-Opt $op 'seconds' 30)))
  return @{}
}

$script:Ops = @{
  moveCols = ${function:Op-MoveCols}; moveRows = ${function:Op-MoveRows}
  insertCols = ${function:Op-InsertCols}; insertRows = ${function:Op-InsertRows}
  deleteCols = ${function:Op-DeleteCols}; deleteRows = ${function:Op-DeleteRows}
  setCells = ${function:Op-SetCells}; setColumnWidth = ${function:Op-SetColumnWidth}
  unmerge = ${function:Op-Unmerge}
}
if ($env:XLRUNNER_ALLOW_DEBUG -eq '1') { $script:Ops['debugWait'] = ${function:Op-DebugWait} }

# ---------------------------------------------------------------------------- main
function Invoke-Job($job) {
  $ver = Get-Opt $job 'version' 1
  if ([int]$ver -ne 1) { Fail 'BAD_JOB' 2 "未対応のジョブ形式です (version $ver)" }
  $inPath = [string](Get-Opt $job 'input' '')
  $output = [string](Get-Opt $job 'output' '')
  if (-not $inPath) { Fail 'BAD_JOB' 2 'input がありません' }
  if (-not $output) { Fail 'BAD_JOB' 2 'output がありません' }
  $format = [string](Get-Opt $job 'format' 'xlsx')
  if ($format -ne 'xlsx' -and $format -ne 'xlsm') { Fail 'BAD_JOB' 2 "出力形式は xlsx か xlsm にしてください: $format" }
  $outExt = [IO.Path]::GetExtension($output).ToLowerInvariant()
  if ($outExt -ne ".$format") { Fail 'BAD_JOB' 2 "出力ファイルの拡張子が形式 ($format) と一致しません: $output" }
  $inExt = [IO.Path]::GetExtension($inPath).ToLowerInvariant()
  if ($inExt -ne '.xlsx' -and $inExt -ne '.xlsm') { Fail 'BAD_JOB' 2 "入力は .xlsx か .xlsm にしてください: $inPath" }
  $ops = @(Get-Opt $job 'ops' @())
  for ($i = 0; $i -lt $ops.Count; $i++) {
    if ($ops[$i] -isnot [System.Collections.IDictionary]) { Fail 'BAD_JOB' 2 "ops[$i] がオブジェクトではありません" }
    $name = [string]$ops[$i]['op']
    if (-not $script:Ops.ContainsKey($name)) { Fail 'BAD_JOB' 2 "未知の操作です: $name (ops[$i])" }
  }
  # files are checked only after the whole job is known to be valid
  if (-not (Test-Path -LiteralPath $inPath)) { Fail 'FILE_NOT_FOUND' 4 "入力ファイルが見つかりません: $inPath" }
  if (Test-FileLocked $output) { Fail 'OUTPUT_LOCKED' 7 "出力先のファイルが他のプログラム（Excel など）で開かれています。閉じてからやり直してください" }
  $opts = Get-Opt $job 'options' @{}
  $script:ClipMode = [string](Get-Opt $opts 'clipboard' 'restoreEachOp')
  $recalc = [bool](Get-Opt $opts 'recalc' $true)

  $work = [string](Get-Opt $job 'workDir' '')
  if (-not $work) { $work = Join-Path ([IO.Path]::GetTempPath()) ('xlrunner-' + [Guid]::NewGuid().ToString('N')); $script:OwnWork = $work }
  New-Item -ItemType Directory -Force -Path $work | Out-Null

  Start-Excel
  # Work on a private copy: the input is never modified.
  $sw = [Diagnostics.Stopwatch]::StartNew()
  $copy = Join-Path $work ('work' + $inExt)
  # Disk hygiene (round 5): the server stores the uploaded workbook AS <workDir>\work.xlsx - it is opened in place instead of being
  # copied again (one copy less of a big workbook on the system drive). Anything else is copied first: the input is never modified.
  $ownCopy = -not (Test-SamePath $inPath $copy)
  if ($ownCopy) {
    [void]$script:TempFiles.Add($copy)
    try { Copy-Item -LiteralPath $inPath -Destination $copy -Force; Set-ItemProperty -LiteralPath $copy -Name IsReadOnly -Value $false } catch { Fail 'OPEN_FAILED' 4 "入力ファイルをコピーできません: $($_.Exception.Message)" }
  } else {
    try { Set-ItemProperty -LiteralPath $copy -Name IsReadOnly -Value $false } catch {}
  }
  $wb = Open-Workbook $copy
  $script:Wb = $wb
  if ($format -eq 'xlsx') { try { if ($wb.HasVBProject) { Warn 'MACROS_DROPPED' 'マクロを含むブックですが .xlsx で保存するためマクロは含まれません（.xlsm を選ぶと保持されます）' $null } } catch {} }
  Emit ([ordered]@{ type = 'progress'; phase = 'open'; ms = [int]$sw.ElapsedMilliseconds })

  $calc0 = $script:xl.Calculation
  try { $script:xl.Calculation = -4135 } catch {}   # xlCalculationManual while replaying (restored before saving)
  $refBefore = Count-RefErrors $wb

  $n = $ops.Count
  for ($i = 0; $i -lt $n; $i++) {
    $op = $ops[$i]; $name = [string]$op['op']
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $seq0 = [XlRunner.Native]::GetClipboardSequenceNumber()
    try { $res = & $script:Ops[$name] $op }
    catch {
      $re = Find-RunnerError $_.Exception
      $script:FailedOp = $i
      if ($re) { throw [RunnerError]::new($re.Code, $re.Exit, $re.Message) }
      throw [RunnerError]::new('OP_FAILED', 5, "操作 $($i + 1)/$n ($name) に失敗しました ($(HResult $_)): $($_.Exception.Message)")
    }
    $ms = [int]$sw.ElapsedMilliseconds
    [void]$script:Timings.Add(@{ i = $i; op = $name; ms = $ms })
    $pl = [ordered]@{ type = 'progress'; phase = 'op'; i = $i; n = $n; op = $name; ms = $ms; clipSeqDelta = [int]([XlRunner.Native]::GetClipboardSequenceNumber() - $seq0) }
    if ($res) { foreach ($k in $res.Keys) { $pl[$k] = $res[$k] } }
    Emit $pl
  }
  Restore-Clipboard

  # Recalculate with the original mode (the saved file stores the calculation mode).
  $sw = [Diagnostics.Stopwatch]::StartNew()
  try { $script:xl.Calculation = $(if ($calc0 -eq -4135 -or $calc0 -eq 2 -or $calc0 -eq -4105) { $calc0 } else { -4105 }) } catch {}
  if ($recalc) { try { $script:xl.CalculateFull() } catch { Warn 'CALC_FAILED' '再計算に失敗しました（値は次回 Excel で開いたときに再計算されます）' $null } }
  Emit ([ordered]@{ type = 'progress'; phase = 'calc'; ms = [int]$sw.ElapsedMilliseconds })

  $refAfter = Count-RefErrors $wb
  if ($refAfter.cells -gt $refBefore.cells -or $refAfter.names -gt $refBefore.names) {
    Warn 'REF_ERRORS' ("削除などで参照先がなくなり、#REF! エラーが増えました（セル " + $refBefore.cells + ' → ' + $refAfter.cells + ' 個、名前 ' + $refBefore.names + ' → ' + $refAfter.names + ' 個）。出力したブックで確認してください') @{ samples = $refAfter.samples }
  }

  # Save into the work folder first, then copy to the destination (no half-written target).
  $sw = [Diagnostics.Stopwatch]::StartNew()
  $fmt = if ($format -eq 'xlsm') { 52 } else { 51 }
  # Normally: save into the work folder first, then copy to the destination (no half-written target). When the destination itself
  # lies in the work folder (the server's private job folder) Excel saves straight into it: one copy less on the system drive.
  $direct = Test-SamePath ([IO.Path]::GetDirectoryName($output)) $work
  $tmpOut = if ($direct) { $output } else { Join-Path $work ('saved.' + $format) }
  if (-not $direct) { [void]$script:TempFiles.Add($tmpOut) }
  if (Test-Path -LiteralPath $tmpOut) { Remove-Item -LiteralPath $tmpOut -Force }
  try { $wb.SaveAs($tmpOut, $fmt, $script:M, $script:M, $false, $false, 1, 2) }   # AccessMode xlNoChange, ConflictResolution xlLocalSessionChanges
  catch { Fail 'SAVE_FAILED' 6 "Excel で保存できませんでした ($(HResult $_)): $($_.Exception.Message)" }
  $wb.Close($false); $script:Wb = $null
  if (-not $direct) {
    $outDir = [IO.Path]::GetDirectoryName($output)
    if ($outDir -and -not (Test-Path -LiteralPath $outDir)) { New-Item -ItemType Directory -Force -Path $outDir | Out-Null }
    try { [IO.File]::Copy($tmpOut, $output, $true) }
    catch {
      if (Test-FileLocked $output) { Fail 'OUTPUT_LOCKED' 7 "出力先のファイルが他のプログラム（Excel など）で開かれています" }
      Fail 'SAVE_FAILED' 6 "出力先に書き込めません: $($_.Exception.Message)"
    }
    try { Remove-Item -LiteralPath $tmpOut -Force } catch {}
  }
  if ($ownCopy) { try { Remove-Item -LiteralPath $copy -Force } catch {} }
  Emit ([ordered]@{ type = 'progress'; phase = 'save'; ms = [int]$sw.ElapsedMilliseconds })
  return @{ output = $output; refErrors = $refAfter; clipboardTouched = $script:ClipTouched }
}

$script:FailedOp = $null
try {
  Import-Helper
  [void][XlRunner.MessageFilter]::Register()
  Emit ([ordered]@{ type = 'start'; runner = $script:RunnerVersion; ps = $PSVersionTable.PSVersion.ToString(); apartment = [string][Threading.Thread]::CurrentThread.ApartmentState; helperMs = $script:HelperMs; language = [string]$ExecutionContext.SessionState.LanguageMode })
  if ($Probe) {
    Start-Excel
    $script:Final = [ordered]@{ type = 'result'; ok = $true; probe = $true; pid = $script:XlPid; version = [string]$script:xl.Version; build = [string]$script:xl.Build }
  } else {
    if (-not $JobPath) { Fail 'BAD_JOB' 2 'ジョブファイルが指定されていません' }
    if (-not (Test-Path -LiteralPath $JobPath)) { Fail 'BAD_JOB' 2 "ジョブファイルがありません: $JobPath" }
    try { $job = $script:Json.DeserializeObject([IO.File]::ReadAllText($JobPath, [Text.Encoding]::UTF8)) }
    catch { Fail 'BAD_JOB' 2 "ジョブファイルの JSON が不正です: $($_.Exception.Message)" }
    if ($job -isnot [System.Collections.IDictionary]) { Fail 'BAD_JOB' 2 'ジョブは JSON オブジェクトにしてください' }
    $res = Invoke-Job $job
    $script:Final = [ordered]@{ type = 'result'; ok = $true; output = $res.output; refErrors = $res.refErrors; clipboardTouched = $res.clipboardTouched }
  }
} catch {
  $ex = $_.Exception
  $re = Find-RunnerError $ex
  if ($re) { $code = $re.Code; $script:ExitCode = $re.Exit; $msg = $re.Message }
  else { $code = 'INTERNAL'; $script:ExitCode = 9; $msg = "内部エラー: $($ex.Message)" }
  $script:Final = [ordered]@{ type = 'error'; ok = $false; code = $code; message = $msg; hresult = (HResult $_); opIndex = $script:FailedOp; where = [string]$_.InvocationInfo.PositionMessage }
} finally {
  try { Stop-Excel } catch {}
  try { Remove-RunnerTemp } catch {}
  try { [void][XlRunner.MessageFilter]::Revoke() } catch {}
  if ($null -eq $script:Final) { $script:Final = [ordered]@{ type = 'error'; ok = $false; code = 'INTERNAL'; message = '内部エラー（結果がありません）' }; $script:ExitCode = 9 }
  $script:Final['warnings'] = @($script:Warnings)
  $script:Final['timings'] = @($script:Timings)
  $script:Final['excelPid'] = $script:XlPid
  $script:Final['excelExited'] = [bool]$script:ExcelExited
  $script:Final['totalMs'] = [int]$script:Clock.ElapsedMilliseconds
  try { $script:Final['comRetries'] = [XlRunner.MessageFilter]::Retries } catch {}
  Emit $script:Final
  try { Release-Com } catch {}   # after the final line: the result is already on its way
}
# Test hook (XLRUNNER_ALLOW_DEBUG=1 only): keep this process alive after the final line, like a powershell.exe that cannot exit because its Excel
# needs ~60 s to shut down (round 5). The server must deliver the result at once and then end exactly this Excel (tests/com/hardening.test.ts).
if ($env:XLRUNNER_ALLOW_DEBUG -eq '1' -and $env:XLRUNNER_HOLD_AFTER_RESULT_MS) { try { Start-Sleep -Milliseconds ([int]$env:XLRUNNER_HOLD_AFTER_RESULT_MS) } catch {} }
exit $script:ExitCode
