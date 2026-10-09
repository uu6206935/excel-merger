# R2b: Ground truth - how real Excel adjusts references for structural column/row operations

Excel: Microsoft 365, Excel 16.0 (build «BUILD»), Japanese UI, driven through COM from Windows PowerShell 5.1.
Everything in this report was measured by a harness that you can run again (see "Re-running"). Nothing here is
from memory or documentation.

## TL;DR

1. **Fixtures**: `research/r2b/excel-ref-fixtures.json` holds «NCASES» cases (and «NCHANGED» of them changed) from
   «NOPS» structural operations on a sheet named `Data` (plus a Japanese copy named `データ 表`). They cover formulas on
   `Other`, `集計 シート` and on `Data` itself, defined names, CF, DV, merged cells, a ListObject with structured refs,
   AutoFilter, frozen panes and column widths. Each formula case has `before`/`after` (Range.Formula) **and**
   `fileBefore`/`fileAfter` (the `<f>` text in the saved .xlsx, which is what ExcelJS/SheetJS read).
2. **The core rule, verified on every case**: "Insert Cut Cells" of n whole columns `s..e` before column X is exactly
   **(1) insert n blank columns at X -> (2) cut/paste the source block onto them -> (3) delete the emptied source
   columns.** Each step uses simple interval rules. `excel-ref-model.mjs` (about 150 lines) implements them and
   reproduces **«MODELPASS» / «MODELTOTAL»** formula and name cases exactly (0 failures). The remaining exceptions are
   listed (legacy `@` quirk, and ranges next to a ListObject).
3. **Insert Cut Cells over COM is unreliable**: in «SILENTN» of «MOVEN» move operations it **silently inserted blank
   columns instead of moving** (no exception). The trigger depends on the content (combinations of CF, DV,
   AutoFilter, spills or legacy arrays). It also crashed Excel several times (RPC 0x800706BA/0x800706BE). The
   sequence *insert blank + `Range.Cut(Destination)` + delete* ("clipFree") gave **identical** results on
   «CLIPIDENT» of «CLIPTOTAL» comparable ops and worked every time. Neither method avoids the clipboard: **both wipe
   the user's Windows clipboard**.
4. JS engine: implement the interval model on the formula text. Do not model moves as "delete + insert" or as "pure
   endpoint following", because neither matches Excel (tables below). CF sqref needs cell-set semantics.
   Structured refs never change when you move or insert. Deleting a table column turns every structured ref that
   mentions that column into `#REF!`.

## Files (all in `research/r2b/`)

| file | what |
|---|---|
| `spec.mjs` -> `spec.json` | declarative catalogue: sheet layouts, 300+ formulas, names, CF, DV, merges, tables, the ops, and per-mode COM step lists |
| `harness.ps1` (UTF-8 BOM) | generic COM harness: `-Action build` creates the template; `-Action run -Mode primary\|clipFree\|complement\|clipFreeClean` replays ops and writes JSONL + saves every after-workbook |
| `run-all.mjs` | orchestrator: watchdog (kills **only** the harness's own Excel PID, taken from `Application.Hwnd`), crash retry, and saves/restores the user's text clipboard |
| `build-fixtures.mjs` + `xlsx-lite.mjs` | assembles `excel-ref-fixtures.json` and `excel-ref-modes-diff.json`. It parses the saved .xlsx (dependency-free zip+XML reader) to give file-level truth |
| `excel-ref-model.mjs` | **reference implementation** of the rules (tokenizer, `adjustFormula`, `mapCell`, `adjustSqref`, `adjustCfFormula`) |
| `verify-model.mjs` | checks the model against the fixtures (COM text **and** file text) |
| `make-examples.mjs` | prints the example tables used below |
| `probe3.ps1`, `probe4.ps1` | crash-isolated repro of the silent blank-insert, and the clipboard side-effect probe |
| `out/templates/*.xlsx`, `out/after/<suite>/<mode>/<op>.xlsx` | every template and every after-state workbook, for inspection |

### Re-running
```
cd research\r2b
node run-all.mjs --fresh                         # all suites, all modes (~45 min); resumes if interrupted
node run-all.mjs --suites main --modes primary   # subset
node build-fixtures.mjs && node verify-model.mjs --show 50
```
Run it only when no other agent is cutting/pasting in Excel. The clipboard is shared, and a foreign clipboard write
between `Cut()` and `Insert()` degrades the move into a blank insert. The harness detects this through row-1 markers
and retries.

## Fixture schema

```jsonc
{ "meta": {...}, "suites": { "main": { dataSheet, sheets, layout, baseline:{names,cf,dv,...} }, ... },
  "ops":   [ { suite, opId, primitive, params, comSteps, afterSource, primarySilentBlankInsert, error, markersAfter, clipboard, ... } ],
  "cases": [ {
    "id": "m02:o040", "suite": "main", "opId": "m02", "primitive": "moveCols", "sheet": "Data",
    "params": { "src": {"start":3,"end":4,"a1":"C:D"}, "count":2, "before": {"index":8,"a1":"H"} },
    "where": "Other",            // host sheet for formulas: Other | 集計 シート | Data | データ 表 ;  or name|cf|dv|merge:<sh>|autofilter:<sh>|freeze:<sh>|table:<name>|colwidths|markers
    "kind": "formula", "tag": "hrange",
    "cellBefore": "B40", "cellAfter": "B40",            // host cell (moves when the host is on Data)
    "before": "=SUM(Data!C5:F5)", "after": "=SUM(Data!C5:D5)",   // Range.Formula (en-US, A1)
    "before2"/"after2": "...",                         // Range.Formula2, only when different
    "fileBefore": "SUM(Data!C5:F5)", "fileAfter": "SUM(Data!C5:D5)",  // <f> text in the .xlsx ({f,t,ref} for array formulas)
    "changed": true, "source": "primary", "notes": "op error: ... | host cell deleted ..." } ] }
```
* `params` always carries 1-based numbers **and** A1 letters. `moveCols.before` is in **original** coordinates: the
  block ends up immediately to the left of the original column `before`.
* `after == null` with `cellAfter == null` means the host cell itself was deleted.
* An op that Excel refused has `op.error` = {message (Japanese UI text), hresult}. The cases of that op show the
  unchanged state.
* `source: "clipFree"` means the primary Insert Cut Cells silently failed, so the truth comes from the equivalent
  clipFree sequence.

## Operations run

«OPSTABLE»

## RULES (what Excel does)

Notation: an interval `[a,b]` is the column (or row) span of a reference. A cell has a = b but follows "single"
rules. `MAX` = 16384 (XFD) or 1048576.

### R1. Insert n columns before X (rows: identical with rows)
* Endpoint rule: `a' = a>=X ? a+n : a`, `b' = b>=X ? b+n : b`. A range that **spans** X grows. A range that **starts
  at X** shifts. A range that **ends at X-1 does not grow** (`C5:E5` with insert before F stays `C5:E5`).
* Whole-row refs (`5:5`) ignore column ops, and whole-column refs ignore row ops.
* Edge of sheet: a **single** cell pushed past XFD becomes `#REF!` (`Data!XFD5` becomes `Data!#REF!`). A **range
  ending at XFD stays anchored at XFD** (`K5:XFD5` becomes `M5:XFD5`, `XFA5:XFD5` becomes `XFC5:XFD5`, and
  `A5:XFD5` does not change).
* CF / formats: inserted columns inherit the left neighbour's format. So a CF area that **ends at X-1 also grows**.
  This differs from formula refs.

### R2. Delete columns s..e
* Range fully inside the block becomes `#REF!`. Otherwise it shrinks: `a' = a<s ? a : a>e ? a-n : s`,
  `b' = b<s ? b : b>e ? b-n : s-1`.
* A single cell inside the block becomes `#REF!`. `Data!XFD5` shifts to `Data!XFC5`. A range ending at XFD stays at XFD.
* A range that shrinks to one cell **keeps its area form** (`C5:C5`, not `C5`).

### R3. Move (Insert Cut Cells) block s..e before X  ==  insert, then cut/paste, then delete
Let n = e-s+1 and p = (X>e ? s : s+n) (the source position after step 1).
1. `insert n before X` (R1). This is why refs to the **last n columns of the sheet become `#REF!` on every move**,
   even the no-op move `C:D before E`.
2. cut `[p, p+n-1]` and paste onto `[X, X+n-1]`, using the move rule:
   * range entirely inside the source: translated by (X-p)
   * a single cell outside the source: unchanged
   * range that **contains the whole source AND the whole destination**: unchanged (the block moved *within* the range)
   * otherwise every endpoint that lies in the source is translated. If the result is inverted (start > end), **the
     range is left unchanged**. The next step then shrinks it.
3. `delete [p, p+n-1]` (R2).
* Target inside the block (`C:D before D`, `C:D before C`): refused with 0x800A03EC
  「この選択は適切ではありません。コピー領域と貼り付け領域が…重ならないようにしてください。」
* The target right after the block (`C:D before E`) is accepted as a no-op, but step 1 still turns the refs to the
  last 2 columns into `#REF!`.

Consequences, which a naive implementation gets wrong (m02 = C:D before H, giving A B E F G **C D** H):

| before | after | why |
|---|---|---|
| `B5:C5` | `B5:F5` | end endpoint was in the block, so it follows the block and the range now covers E,F,G too |
| `B5:D5` | `B5:G5` | same |
| `C5:E5` | `C5:C5` | start follows and the range inverts, so it stays unchanged. The delete step then shrinks it to the old E only |
| `C5:H5` | `C5:H5` | the range contains both the source and the destination, so it is unchanged |
| `D5:H5` | `G5:H5` | start D follows to G. The range now covers only old D and old H |
| `E5:H5` | `C5:H5` | shifted start. The range now also covers the moved C,D |
| `Data!XFD5` | `Data!#REF!` | the insert step pushes it off the sheet |

### R4. Things Excel never adjusts
* `INDIRECT("Data!C5")`, text literals, `OFFSET`/`INDEX` numeric arguments, and the **VLOOKUP/HLOOKUP column-index
  argument** (`VLOOKUP(x,Data!C2:F20,2,FALSE)` becomes `...Data!C2:D20,2...`, so it now returns a different column).
* **3D references** (`Data:Other!C5`, `'Data:集計 シート'!C5`, name `=Data:Other!$C$5`). Single-sheet ops do not touch them.
* Refs to other sheets. Unqualified refs on other sheets (`=C5` on Other).
* **Defined names with relative refs** (`nm_rel =Data!C5`, created with A1 active): not adjusted at all.

### R5. Spelling of `#REF!`
* A qualified ref becomes `Data!#REF!`, and a quoted one becomes `'データ 表'!#REF!`. An unqualified ref becomes
  `#REF!`. The rest of the formula stays (`SUM(Data!#REF!)`).
* A spill ref keeps its `#`: `SUM(Data!D25#)` becomes `SUM(Data!#REF!#)` (file text: `_xlfn.ANCHORARRAY(Data!#REF!)`).
* A legacy implicit-intersection formula (`=Data!C5:E5` entered with `.Formula`) becomes `=@Data!#REF!`. The file
  stores `_xlfn.SINGLE(Data!#REF!)`.
* In names: `=Data!#REF!`. In DV: `=#REF!` or `=M2>#REF!`. In CF: the same way.

### R6. Every ref kind follows the same interval rules
Absolute, relative and mixed refs behave identically for moves, inserts and deletes; `$` is preserved. The same holds
for union `(C5,E5:F5,H5)` (each area on its own), intersection `C1:E5 D1:D9`, whole-column `C:E`, spill `D25#`,
`SUMIF(range, , sum_range)` (each argument on its own), sheet-qualified refs on the same sheet (`Data!C5` on Data),
names, and the absolute parts of DV/CF formulas.

«EXAMPLES»

### R7. Defined names
Absolute names follow R1-R3 (for example, a whole-column name `=Data!$C:$E` becomes `=Data!$C:$C` after m02).
Sheet-scoped names (`Data!loc_E`, `'集計 シート'!loc_jp`) behave the same. The hidden `Data!_FilterDatabase`
follows the AutoFilter range. Constants and names that are not references do not change. Relative names: see R4.

### R8. Conditional formatting (read from the file; the identity comes from a unique dxf colour, because Excel **renumbers priorities** on moves)
* sqref: insert and delete follow R1/R2 with the CF extension (an area ending at X-1 also grows). **Move uses
  cell-set semantics**: each cell's membership moves with the cell, so areas can split (`B2:F20` becomes
  `B2:D20 F2:G20`).
* **Excel quirk ("leak")**: when a block is inserted before X, every CF area that ends exactly at column X-1 is
  extended over the moved block. So the moved columns pick up their new left neighbour's rules (m02 gave old C an
  extra 2 rules from old G). Plain cell formats (fill, bold, number format, widths) do **not** leak: they travel with
  their column.
* Formula: stored relative to the **top-left of the sqref bounding box**. Relative parts are offsets and are
  re-expressed from the new top-left (`D2>0` on `D2:D20 G2:G20` becomes `E2>0` on `E2:G20`). Absolute parts
  follow R1-R3 (`$E2` becomes `$G2`). A cross-sheet CF (`Data!$C$5>0` on Other) lives in `x14` extLst.
* The model reproduces «CFPASS»/«CFTOTAL» CF cases (with `emulateLeak`).

### R9. Data validation
The DV range behaves like a cell range (R1/R2, cells travel on a move). `formula1` absolute refs follow R1-R3.
List sources on the same sheet are stored without the sheet name (`=$C$2:$C$6`). Cross-sheet DV sources
(`=Data!$C$2:$C$6` on Other) live in `x14:dataValidations`. Deleting the source column gives `=#REF!`. A DV whose
cells were all deleted disappears.

### R10. AutoFilter, frozen panes, column widths
* AutoFilter range: R1-R3 as a range (m11 A before U changes `A1:T20` to `A1:S20`, and m12 T before A changes it to
  `B1:T20`).
* Frozen panes (`xSplit` = number of frozen columns k): only the insert and delete steps apply.
  `k' = k+n` if X<=k. A delete removes the frozen columns it hits. **The move itself does not count.** For example,
  m02 (C:D before H) turns frozen A:C into frozen A:B, and m06 (F:G before C) turns it into A:E.
* Column widths, column-level formats and cell formats travel with their columns in every move mode.

### R11. Tables (ListObject `tblMain` on Data!C3:H10)
* Moving whole sheet columns **inside** the table: allowed. The table columns are reordered, and structured refs do
  not change (`tblMain[点数]` still means 点数).
* Moving a column **out of** the table, or **into** it from outside, or a block that partly overlaps it: refused.
  - 「テーブル内の他のセルに予想外の影響を与える可能性があるため、テーブル内のセルの配置をこの方法で変更することはできません。」
  - 「この操作を実行するとワークシート上のテーブル内でセルが移動されてしまうため、この操作は行われません。」
  - The **clipFree sequence gets around this check and corrupts the table**: the table column is deleted, so refs
    become `#REF!`, or a foreign column becomes a table column named `1003`. The COM engine must pre-check
    `ListObjects`.
* Inserting a column inside the table adds a table column `列1`. Deleting a table column turns **every structured
  ref that mentions it** into `#REF!`, even column ranges whose other end survives (`tblMain[[点数]:[合計]]` becomes
  `#REF!`). Calculated columns become `=#REF!*2`. Deleting the whole table turns `ROWS(tblMain)` into `ROWS(#REF!)`.
* Deleting the header row is refused (「Range クラスの Delete プロパティを取得できません。」).
* Deviations from the plain-range model (the full list comes from `verify-model.mjs`):
  «TABLEQUIRKS»

### R12. Merged cells, legacy array formulas
* A merged area adjusts like a range (insert inside grows it, delete shrinks it, a move of the whole area moves it).
  A merge whose cells are all deleted is removed.
* Moving a column that splits a merge is refused: 「この操作は結合したセルには行えません。」. **Exception**: moving
  column B *into the middle* of C3:E3 (`B before D`) succeeded, and **Excel silently un-merged C3:E3.**
* A legacy CSE array (N2:P2) whose columns are partly inserted, deleted or moved is refused: 「配列の一部を変更することはできません。」
  Moving the whole array with Insert Cut Cells **silently inserted blank columns** instead (see COM notes). The
  clipFree sequence moved it correctly.

## COM engine notes (for the "Excel本体" export engine)

1. **Use insert-blank + `Range.Cut(Destination)` + delete, and verify the result.** Insert Cut Cells
   (`Cut()` + `Insert`) silently degrades into a blank insert (no exception, CutCopyMode was xlCut) in «SILENTLIST».
   A crash-isolated repro (`probe3.ps1`) showed:

   «PROBE3»

   Always check a signature (header row values, or a temporary marker row) after each move, and abort or undo by
   closing without saving.
2. **Excel crashes**: during about «NMOVESRUN» cut/insert ops, the harness's private Excel died «NCRASH» times
   (0x800706BA "RPC server is unavailable" / 0x800706BE). Work on a copy, one Excel process per job, add a watchdog
   and a retry, and never touch the user's Excel.
3. **Clipboard (`probe4.ps1`)**: plain Insert/Delete leave the clipboard alone (sequence +0, sentinel text kept).
   `Cut()+Insert` gives sequence +35 and **empties the user's clipboard**. `Cut(Destination)` does the same (+35,
   emptied). `Copy(Destination)` also gives +34. `Cut()` alone gives +0 at first, but +1141 and emptied once the
   workbook closes (lazy clipboard ownership). **There is no clipboard-free move in the Excel OM.** Save and restore
   the clipboard (text, HTML and image formats through .NET `Clipboard.GetDataObject`, with retries because it is
   often locked), or tell the user.
4. `$xl.CutCopyMode = 0` / `= $false` **throws in PowerShell** (the PIA enum has no 0), so cut mode silently stays
   on. Use `$xl.CutCopyMode = [Enum]::ToObject([Microsoft.Office.Interop.Excel.XlCutCopyMode], 0)`. A plain
   `Insert` done while cut mode is active performs *Insert Cut Cells*.
5. Refusals come back as COMException 0x800A03EC with the Japanese UI text. No modal dialog blocked a call with
   `DisplayAlerts=false`. The watchdog never fired.
6. `FormatConditions.Add` / `Validation.Add` relative formulas: Excel interpreted them relative to the selected cell.
   The harness selects the top-left cell first, and the file XML shows the correct anchor.
7. Freeze panes, AutoFilter and names are adjusted by Excel for free. The CF "leak" (R8) also happens with clipFree.
   The `clipFreeClean` variant (ClearFormats on the inserted blank columns before the paste) «CLEANRESULT».

## Mode comparison (`excel-ref-modes-diff.json`)
«MODEDIFF»

## Pitfalls checklist for implementers
* Do not implement move as "delete block + insert block" (that gives `B5:B5` for `B5:D5`, but Excel gives `B5:G5`).
  Do not implement it as "map both endpoints and sort" either (that gives `F5:H5` for `C5:H5`, but Excel keeps
  `C5:H5`). Use R3.
* Keep `C5:C5` as an area. Keep `$`. Keep the `#` of spills even on `#REF!`.
* Apply the rules to **file text** too: `_xlfn.ANCHORARRAY(Data!D25)`, `_xlfn.XLOOKUP`, `_xlfn.SINGLE`, array
  formulas `t="array" ref=".."` (the `ref` attribute must be moved too), and shared formulas `t="shared"` (see the
  `cfrel` suite).
* Moving columns changes which column a VLOOKUP index points at. Warn the user, as Excel does not help here.
* The Japanese sheet name `データ 表` needs quotes everywhere (`'データ 表'!#REF!`, `'集計 シート'!loc_jp`).
* Structured refs: skip text inside `[...]` when tokenizing. Moving or inserting never rewrites them.
* CF priorities are renumbered on moves, so do not use priority as an identity.
