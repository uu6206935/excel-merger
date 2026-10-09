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
