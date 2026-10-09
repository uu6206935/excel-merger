# Excel本体 engine (excel-runner.ps1)

Replays the editor's op log in a hidden Excel instance of its own, so Excel itself adjusts every reference
(formulas on all sheets, names, CF, DV, charts, pivots, tables, freeze panes, autofilter).

| file | role |
|---|---|
| `excel-runner.src.ps1` | runner source (PowerShell 5.1) with a `#__CSHARP__` placeholder |
| `XlRunnerHelper.cs` | C# 5 helpers: message filter (busy retries), clipboard snapshot/restore, cell block builder, move verification, ASCII JSON writer |
| `excel-runner.ps1` | **built** file = source + helper, UTF-8 **with BOM**, CRLF (`npm run build:runner` → `build-runner.mjs` → `merge-runner.mjs`) |
| `run-job.mjs` | Node side: registry detection (`detectExcel`), `runRunner` (ScriptBlock invocation, watchdog, cancel, `earlyResult`, kill ONLY the job's own Excel PID — verified by image name AND process start time) |
| `cli.mjs` | developer CLI: `node server/excel/cli.mjs --detect | --probe | job.json` |
| `../safe.mjs` | the server's pure input-hygiene helpers (file names, open paths, the op whitelist `validateRunnerOps` the job API applies before anything reaches PowerShell) |

## Why a result is delivered before the process has exited (round 5)

Measured on this machine (soak test, 40 jobs of the real runner, all with the runner's final `result` line out after 5-6 s): 6 jobs took
63-64 s. `powershell.exe` could not exit during those ~58 s and Excel was idle (flat CPU) with 8-9 open HTTPS connections to Microsoft endpoints:
a hidden Excel that has been told to `Quit()` lives on until PowerShell lets go of it (the COM binder keeps references, it never exited within
2 s), then it shuts down by talking to Office's cloud services (telemetry, identity), and when that stalls it needs ~60 s; `powershell.exe`
cannot exit before it (COM rundown). The server used to wait for the process (`close`) and then for the Excel PID.

Now: `runRunner({ earlyResult: true })` resolves when the final `result` / `error` line arrives (the server then verifies the output file: present,
not empty, complete ZIP directory), the process shutdown continues in the background (`closed` promise) and a watchdog kills ONLY our Excel PID
(`tasklist` image EXCEL.EXE **and** the start time the runner reported in `{"type":"excel","started":ms}`) and then `powershell.exe` after
`EXCEL_EXIT_GRACE_MS` (6 s; normally Excel is gone ~1 s after PowerShell). The runner itself no longer waits for Excel (`-QuitWaitMs` default 0;
`-QuitWaitMs N [-KillOnExit]` for standalone use) and the GC / finalizer wait that releases Excel's COM objects (`Release-Com`; it blocks while a
quitting Excel winds down: ~2.4 s in 2 of 10 jobs, it used to delay the result line by exactly that) runs AFTER the final line: `save` -> `result` is
now 80-300 ms in every job. Test hook: `XLRUNNER_ALLOW_DEBUG=1 XLRUNNER_HOLD_AFTER_RESULT_MS=<ms>` leaves Excel un-quit and the process alive after
the result (an Excel "stuck in its shutdown").

Numbers (this machine, `samples/01`, 6 ops): before, 40 consecutive jobs of the runner with the old `close` wait: 6 of 40 took 63.4-64.3 s (the result
line was out after 5.2-6.1 s in every one of them; first 10 jobs: 235.8 s in total). After, 40 consecutive jobs through the real server:
4.4-9.3 s each (mean 6.3 s, 0 stalls, 0 watchdog kills; first 10 jobs: 49.6 s in total); the stall itself did not recur after the fix, so it is simulated with
the real Excel by the hold hook (result after ~5 s, our Excel ended by the watchdog 1.5 s later, a second real Excel untouched). `tests/server/exit-stall.test.ts` (fake runner, fake EXCEL.EXE processes),
`tests/com/hardening.test.ts` (real Excel, 10-job soak).

## Invocation

The server runs `powershell.exe -NoLogo -NoProfile -NonInteractive -Command "& ([ScriptBlock]::Create([IO.File]::ReadAllText('<ps1>', [Text.Encoding]::UTF8))) -JobPath '<job.json>'"`.
A ScriptBlock is not subject to the execution policy (works under a GPO "Restricted" policy, no `-ExecutionPolicy Bypass`).
If PowerShell cannot even start it, `runRunner` retries once with `-ExecutionPolicy Bypass -File`.
Environment: `XLRUNNER_CACHE` (compiled helper DLL cache, the server uses `%LOCALAPPDATA%\ExcelMerger\cache`),
`XLRUNNER_ALLOW_DEBUG=1` (enables the `debugWait` test op).

## Job (UTF-8 JSON)

```jsonc
{
  "version": 1,
  "input": "C:\\...\\work.xlsx",         // filled in by the server; never modified by the runner's own logic (a copy is opened) - EXCEPT when it already is <workDir>\work.xlsx (the server's private upload): then it is opened in place (one copy less on the system drive)
  "output": "C:\\...\\output.xlsx",      // extension must match "format"; when it lies directly in workDir Excel saves straight into it (no saved.xlsx + copy)
  "workDir": "C:\\...\\job-xxxx",        // optional scratch folder (the server: its private job folder)
  "format": "xlsx",                      // or "xlsm" (keeps macros)
  "options": { "recalc": true, "clipboard": "restoreEachOp" },   // or "restoreAtEnd"
  "ops": [ /* in order, 1-based positions as of the time of the op */ ]
}
```

| op | fields | notes |
|---|---|---|
| `moveCols` / `moveRows` | `sheet, start, count, to, method: "clipFree"` | `to` = insert-before index in ORIGINAL coordinates (= excelRef `StructOp`). clipFree = insert blank lines at `to`, `Range.Cut(Destination)`, delete the source; the values of the moved window are verified afterwards (abort without saving on mismatch). Inside one table Excel's own Cut + Insert Cut Cells is used: verified, retried up to 6 times with a random back-off when another program touches the clipboard between the two calls (10 of 10 jobs exact when the clipboard is quiet), and the column widths are put where the app has them (Excel shifts the widths of the columns right of the insertion point). `to` = the gap right after the table's last column (`c2 + 1`) is an in-table reorder too (the column stays the last table column; logToRunnerJob sends it unchanged, the runner recognises the table by itself; the doc layer and replayLog accept it as well since round 3). Excel moves only the table rows there: with cells above / below the table in the same columns (titles, totals) a move to a destination inside the table is refused: error code **`MOVE_VERIFY_FAILED`**, nothing saved, and the message (Japanese) says to use the built-in engine (「内蔵エンジン」で出力してください); the server adds `fallbackToJs: true` to the job view and `src/api.ts` to the ApiError (`canFallBackToJs(e)`), and the UI retries with the built-in engine for every error except a cancel. |
| `insertCols` / `insertRows` | `sheet, at, count, formatFrom: "left"/"above" (default) · "right"/"below" · "none"` | |
| `deleteCols` / `deleteRows` | `sheet, at, count` | |
| `setCells` | `sheet, at: "C5", rows: [[cell, ...], ...], textMode?: "auto" (default: explicit flags only) \| "prefixAll" \| "raw" (= auto), toValues?: true` | rectangular block from `at`. `toValues` (LoggedOp reason `toValues`): an array formula (CSE) / spill under the block is replaced by its values as a whole before the writes (Excel cannot change part of an array) |
| `setColumnWidth` | `sheet, col, count?, width, unit?: "chars" \| "stored"` | `chars` (default) = Excel character units (`Range.ColumnWidth`), 0..255. `stored` = the value stored in `<col width>` (LoggedOp `colWidth`, includes the padding): the runner sets `ColumnWidth = width - padding`, padding measured in the workbook's own Excel (0.625 for 11pt 游ゴシック, 5/7 for 11pt Calibri) |
| `unmerge` | `sheet, refs: ["A1:F1", ...], centerAcross?: bool` | BEFORE `Range.UnMerge` the top-left cell's look is copied to the other cells of each area (LoggedOp contract, same rule as the JS patcher): its **fill** (`Interior`: pattern, colour or theme colour + tint, pattern colour) goes to every other cell, each **border side it has** goes to the cells on that edge of the area (top: first row, bottom: last row, left: first column, right: last column; a one-row band: top and bottom on every cell, left on the first, right on the last cell); a side it does not have leaves the other cell's own side alone. Excel leaves the other cells of an openpyxl-style band unstyled (measured) and `UnMerge` copies nothing, so the band would lose its fill and lines. Excel shares one line between two neighbours (A1.right is saved as B1.left): same picture, other records. Then `UnMerge`; `centerAcross`: `HorizontalAlignment = 7` (xlCenterAcrossSelection, 「選択範囲内で中央」) over each former area so a merged title keeps its look |

Cell encoding in `setCells`: `null` clears the value (format kept) · a bare string is **text written as it is**
(`Range.Value2 = text`, never guessed: `03-1234-5678` stays a plain text cell without quotePrefix) and verified afterwards ·
`{"t": "s", "v": "001", "quote": true}` is the same text **with the quote prefix** (`Value2 = "'" + text`,
quotePrefix="1" in the file, also in a cell whose format is `@`) — exactly when `CellWrite.quote === true`: the doc sets it
for strings Excel would re-parse (`001`, `1/2`, `=abc`, `TRUE`, `+81 …`), and **both engines honour it exactly**
(`quote === true` → quotePrefix, anything else → none). A text that begins with an apostrophe always gets one more (Excel eats
the first; there is no other way to store it). Safety net for strings the doc did not flag (user edits): the runner reads
`Value2` back, and a string Excel converted anyway (number, date, TRUE, formula, error) is rewritten as text WITH the
apostrophe, the number format the cell had is put back (Excel sets a date format for `1/2`; measured) and the warning
`TEXT_COERCED` names the cell (at most 20 are listed). A string Excel refuses to enter at all (`=(`) is written cell by cell
the same way. Known difference: Excel keeps an existing quotePrefix of a cell when a later plain write replaces the text
(measured: it stays through text / number / formula writes; only Clear / ClearFormats remove it, and no cheap call
replaces the cell format) — the JS patcher clears it, so a plain write over a quoted cell differs in that one flag ·
number · boolean · `{"f": "=SUM(A1:A3)"}` formula in UI form written with `Range.Formula` (legacy semantics, what a
normal cell holds in the file; an unparsable formula is written as text and reported) · `{"f": "=SORT(A1:A9)", "dyn": true}`
dynamic-array formula written with `Range.Formula2` · `{"e": "#N/A"}` error literal · `{"t","v"}` (contracts.ts
`RunnerCell`) · `{"skip": true}` leaves the cell untouched. Any object cell may carry `"numFmt"`: `Range.NumberFormat`
is applied to the cell (a vertical run of cells with the same code = one call). `"@"` is set BEFORE the values (strings
stay literal), every other code AFTER them: Excel widens a column of standard width for a date / `1,234,567` that no
longer fits only when the format is applied to a cell that already holds the number (like typing it; set before the
value it does not widen and the date would show `########`), and `replayLog` emulates exactly that widening. `General` resets the format: Japanese Excel **rejects the English word through COM** (every `NumberFormat = 'General'` fails there: "Range クラスの NumberFormat プロパティを設定できません"), so the runner translates it (also inside a longer code, outside quotes) to the name this Excel reports, `Application.International(xlGeneralFormatName)` = `G/標準`. A formula
cell without `numFmt` keeps the number format it had (Excel would copy the referenced cell's format into a General cell).

Safety (round 5): a formula - or a typed text that Excel would read as a formula (`=` `+` `-` `@` first) - that would make Excel reach outside
the workbook by itself is **never entered**; it is written as plain text with the quote prefix and reported as warning `UNSAFE_FORMULA`
(`XlRunner.Cells.UnsafeFormula`): DDE (`=cmd|' /C calc'!A0`, `=MSEXCEL|'...'!A1`: a `|` outside string literals and quoted names), external
workbook references (`[Book.xlsx]Sheet1!A1`, `'C:\dir\[Book.xlsx]S'!A1`, UNC / URL paths), `WEBSERVICE(` `IMAGE(` `RTD(` `CALL(` `REGISTER(` `REGISTER.ID(` `EXEC(`.
`HYPERLINK` and structured references are fine. The cell address of `setCells` must be a plain A1 address (`Worksheet.Range` also accepts names
and external references: `'\\host\share\[x.xlsx]S'!A1` would make Excel reach for the network). Verified with the real Excel (marker-file payload):
`tests/com/hardening.test.ts`, the classification in `tests/server/helper.test.ts`.

Refusals (checked before the op, Japanese messages, exit code 5, nothing saved): moving part of a merged area or into
the middle of one, moving columns into / out of / partly over a table, moving table rows, touching part of an array
formula.

## Output (stdout, one ASCII JSON object per line)

`{"type":"start"}` → `{"type":"excel","pid":N,"started":<process start, Unix ms>}` (the server kills exactly this PID on cancel / timeout / a stalled shutdown, and only when its start time still matches: Windows recycles PIDs) →
`{"type":"progress","phase":"excel"|"open"|"op"|"calc"|"save",...}` / `{"type":"warn","code","message"}` (codes: `TEXT_COERCED`, `BAD_FORMULA`, `BAD_NUMFMT`, `UNMERGE_FORMATS`, `MOVE_RETRY`, `CLIPBOARD_*`, `REF_ERRORS`, `MACROS_DROPPED`, `WIDTH_RESTORE`, `CHECK_LIMIT`, `CALC_FAILED`, `UNSAFE_FORMULA`) →
finally `{"type":"result","ok":true,"output","refErrors","warnings","timings",...}` or
`{"type":"error","ok":false,"code","message",...}`.
Exit codes: 0 ok, 2 bad job, 3 Excel unavailable, 4 open failed, 5 op failed / refused, 6 save failed, 7 output locked, 9 internal.

## Disk hygiene (the system drive is often nearly full)

- Standalone use: the runner opens a COPY (`work.xlsx`) and saves into `saved.xlsx` inside the work folder before copying to `output`. Both are removed after **every** job,
  success or failure (`Remove-RunnerTemp`, retried while a lingering Excel still holds them). A job without `workDir` gets `<TEMP>\xlrunner-<guid>`, which the
  runner removes itself at the end (also after a failure). Under the server (round 5) the upload is streamed straight to `<job folder>\work.xlsx` (nothing is
  held in memory), opened in place and saved straight into `<job folder>\output.xlsx`: **two** files on disk at the peak instead of four.
- The server keeps one folder per job, `%TEMP%\excel-merger\job-<8 hex>-<server pid>` (`%TEMP%` of the user, Unicode paths work; `EM_TEMP_DIR` overrides): a finished job keeps only
  `output.xlsx` for the download (the upload copy and `job.json` go at once), a failed / crashed / cancelled job loses the whole folder at once,
  `DELETE /api/jobs/:id` (sent by `src/api.ts` right after the download) removes the rest, finished jobs are dropped after 15 minutes anyway, and the folders are
  removed when the server shuts down.
- Excel's own TEMP: Excel started through COM inherits `TEMP` from the process that starts it and writes ~80 MB of Office diagnostics into
  `%TEMP%\Diagnostics\EXCEL` per start (never cleaned by Office). The server starts every runner with `TEMP` / `TMP` = `%TEMP%\excel-merger\xl-<8 hex>-<server pid>`
  and deletes that folder when that Excel has exited (also the diagnostics), so a job adds nothing to the user's real `%TEMP%`.
- Free space: `GET /api/status` has `diskLow` (< 2 GB free on the drive of the temp folder), `diskFreeMB`, `diskMessage` (Japanese) and `outputDiskLow`; the same
  warning is logged at start and when an Excel job starts. A job whose upload does not fit (free < 300 MB + 4 x the workbook size, `EM_MIN_FREE_MB`) is refused
  up front with errorCode `DISK_LOW` + `fallbackToJs` (the UI uses the built-in engine, which works in memory); a failed job on a nearly full drive is reported the same way.
- At server start stale folders are swept: `job-*-<pid>` / `xl-*-<pid>` of a server that is gone after 1 hour, **any** of them after 1 day (a pid can be reused, so "alive"
  proves nothing after a day). Nothing else in the temp root is touched.
- Tests: `tests/server/hygiene.test.ts` (fake runner), `tests/com/runner.test.ts` (scratch folder of a job without `workDir`), `tests/com/parity.test.ts`
  (work folder empty after a job).
- The tests that start Excel directly (`tests/com`) still write the diagnostics into the `TEMP` they run with: run them with `$env:TEMP='D:\...'` and delete that folder afterwards.
