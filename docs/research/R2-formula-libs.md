# R2: Excel formula tokenizing / reference rewriting / evaluation libraries

Scratch folder: `C:\Users\81802\AppData\Local\Temp\claude\D--20260705-----------\2d64adba-cbb0-49f4-bf69-6262ff0972bb\scratchpad\research\r2`
(everything below is reproducible from there; `r2\proto` = prototype to start from).

---

## 0. TL;DR (recommendation)

1. **No existing JS library rewrites references with Excel semantics.** HyperFormula and IronCalc both have
   `moveColumns/addColumns/removeColumns`, but compared to real Excel 365 they differ on 16-20% of formulas.
   They also silently skip formulas they cannot parse, which is worse. Details are in §2.
2. **Use our own small tokenizer plus an Excel-semantics adjuster.** `r2\proto\excelRef.ts` has 724 lines,
   0 dependencies and runs in both browser and Node. It is lossless: only reference tokens are rewritten, so
   unknown functions, LET/LAMBDA and strings are copied as-is. The move/insert/delete rules were derived from
   **Excel itself via COM** and verified:
   * 1,393 / 1,393 checks against Excel for cell formulas, defined names, CF formulas, data validation (DV)
     and chart series. There are 20 documented, deliberate deviations; all are sheet-edge quirks, see §3.5.
   * **4,608 / 4,608** in an exhaustive test: every range inside A..H × every block move (128 moves).
   * Conditional formatting (CF) and DV rules with relative refs: 62 / 62 per-cell checks. There are 2 known
     Excel artifacts that we deliberately do not copy.
   * Tokenizer: 39,800 assertions. These include 13,168 unique grammar-fuzzed formulas with known ground
     truth. **0 disagreements** with `@borgar/fx` on every formula fx can tokenize.
   * Speed: 50,000 realistic formulas adjusted in about 0.8 s (≈16 µs/formula). A 11,675-char formula with
     1,200 refs takes 2-3 ms.
3. **`@borgar/fx` (MIT, 0 deps) is the best existing tokenizer.** It scored 83/84 on the torture list with
   exact round trip. Use it only as a **dev-time differential-testing oracle**, not in production. The
   differential test found 2 fx bugs:
   * a structured ref followed by a space or intersection becomes `unknown` tokens;
   * a quoted external path (`'C:\p\[B.xlsx]S'!A1`) fails.
4. **HyperFormula (GPL-3.0):**
   * (a) **Not suitable** as the reference-rewriting engine.
   * (b) Usable only as a **preview evaluator**, with caveats.
   * **IronCalc (`@ironcalc/wasm`, MIT/Apache-2.0) is the better preview evaluator.** It passed 14/14 of our
     evaluation tests (Japanese sheet names, XLOOKUP, FILTER, LET, TEXTJOIN, DATEDIF…); HyperFormula passed
     8/14. The catch is a 659 KB gz wasm file.
   * The simplest option is no evaluator at all: set `fullCalcOnLoad` and let Excel recalculate.
5. **Excel-COM engine: never use `Range.Cut()` + `Range.Insert()` (clipboard).** In this environment 31 of 128
   clipboard moves silently inserted blank columns instead of moving. Use the clipboard-free **3-step move**
   instead. It gave the same result as Cut+Insert on all 3,492 comparisons. See §7.

---

## 1. Method

* Environment: Windows 11, Node v24.16.0, Excel 16 (Microsoft 365) via COM. Every Excel run used its **own
  hidden instance** (`New-Object -ComObject Excel.Application`). Only that instance's PID was killed, found
  via `GetWindowThreadProcessId(Hwnd)`.
* **Torture list:** `r2\torture.mjs` has 84 cases:
  * Japanese sheet names, quoted and unquoted; `'O''Brien'`
  * whole columns and rows; `$` mixes; 3D refs; external refs; `Sheet1!A1:Sheet1!B2`
  * names that look like cells (`TAX2025`, `LOG10`, `XFE1`, `A1048577`, `A1B`, `R1C1`)
  * structured refs, strings containing refs, array constants, errors
  * intersection and union operators; `_xlfn.` / `_xlws.` / `_xlpm.`; LET and LAMBDA; INDIRECT
  * spill `#`, `@`, `%`, unary minus, full-width characters, a 11,675-char formula
* **Ground truth from Excel.** These scripts live in `r2\oracle\` and are plain ASCII with a UTF-8 BOM:
  * `excel-oracle.ps1` → `excel-results.json`: 76 formulas on sheet F referencing sheet Data, plus
    same-sheet formulas, names, 2 CF rules, 1 DV, 1 chart series. It runs 16 ops: insert/delete columns,
    moves (Cut+Insert and 3-step variants), insert/delete/move rows.
  * `excel-exhaustive.ps1` → `exhaustive-results.json`: all 36 ranges in A..H × all 128 column-block moves
    (count 1-3), done by clipboard **and** by 3-step, with marker-cell verification.
  * `excel-cf.ps1` → `cf-results.json`: 7 CF formulas with relative refs + 1 custom DV × 8 ops.
  * `excel-storage.ps1` → `storage\storage-test.xlsx` + `storage-results.json`: which sheet names Excel
    quotes, and how modern syntax is stored in xlsx XML.
  * `excel-shared.ps1` → `storage\shared-test.xlsx`: shared formulas, used for the ExcelJS test.
* **Library harnesses** (in `r2\`):

  | script | what it does |
  |---|---|
  | `compare.mjs` | torture matrix for every library |
  | `hf-vs-excel.mjs` | HyperFormula ops vs Excel |
  | `ironcalc-vs-excel.mjs` | IronCalc ops vs Excel |
  | `eval-matrix.mjs` | evaluators + bundle sizes |
  | `exceljs-shared.mjs` | ExcelJS shared formulas / `spliceColumns` |

---

## 2. Library evaluation

### 2.1 Tokenizers / parsers: torture list (84 cases)

| library (version, license) | torture OK | lossless round-trip | names vs cells | notes |
|---|---|---|---|---|
| **own `excelRef.ts`** (proto) | **84/84** + 10 UI-form extras | yes, always | yes (bounds-checked; `LOG10(` vs `LOG10`, `LOG10!`) | also handles `'C:\p\[B.xlsx]S'!A1` |
| **@borgar/fx** 5.1.3, MIT, 0 deps, 21 KB gz | 83/84 | yes (`tokens.map(t=>t.value).join('')`) | yes | fails: quoted external path; **bug**: `Table1[[#This Row],[x]]` followed by whitespace becomes `unknown` tokens (71/13,168 fuzz formulas) |
| excel-formula-tokenizer 3.0.0 (Bachtal port), MIT | 54/84 | no (strings are unquoted) | **no** (everything is `operand/range`) | no structured-ref / spill / LET awareness |
| excel-formula (joshbtn) 1.5.0, MIT | 54/84 | no | no | same Bachtal core; pulls in jquery + bootstrap |
| excel-formula-ast / excel-formula-parser (psalaets) | 29/84 | no | no (`TAX_RATE` becomes a "cell") | throws on many inputs |
| @xlscalc/formula-engine 0.4.0, MIT | 51/84 parsed | `unparse` differs | yes | refuses 3D, external, structured, `_xlfn` |
| fast-formula-parser 1.0.19, MIT | 44/84 | n/a (evaluator) | partly | no external, 3D, structured, Japanese names (regex lexer, ASCII names) |
| hot-formula-parser 4.0.0, MIT | 17/84 | n/a | n/a | no sheet refs at all |
| xlsx-calc 0.9.2, MIT | not scored | n/a | n/a | regex `[A-Z]+[0-9]+` (upper case only); evaluator on SheetJS objects |
| HyperFormula 3.4.0, GPL-3.0 | 41/84 valid | re-renders (spacing, order) | yes | see §2.2 |

The `@borgar/fx` API used in the tests works and is a good dev-time oracle:

```js
import { tokenize, parseA1Ref, addA1RangeBounds } from '@borgar/fx/xlsx';   // xlsx variant: [1]Sheet!A1, sheetName/workbookName
const toks = tokenize(formula, { mergeRefs: true, withLocation: true });  // {type, value, loc:[s,e]}
// types: range, range_beam (A:A, 1:1), range_named, structured, func, string, number, bool, error, operator, whitespace, context, unknown
parseA1Ref("'集計 シート'!$A$2:$C$100"); // {sheetName:'集計 シート', range:{top,left,bottom,right,$top,...}}
```

Notes on fx:
* fx does **not** merge `Sheet1!A1:Sheet1!B2` into one token. That is correct: Excel adjusts the two
  endpoints independently (see §3.4).
* fx gives spill `A1#` as `range` + operator `#`.

### 2.2 HyperFormula 3.4.0: detailed findings

* **License:** `GPL-3.0-only` (commercial license otherwise). Pass `{ licenseKey: 'gpl-v3' }`.
* **Size:** `dist/hyperformula.full.min.js` is 923 KB raw / 192 KB gz. Dependencies: chevrotain 6, tiny-emitter.
* **API that worked:**

  ```js
  const { HyperFormula } = require('hyperformula');
  const hf = HyperFormula.buildFromSheets({ Data: [['x']], F: formulas.map(f => [f]) }, { licenseKey: 'gpl-v3' });
  const D = hf.getSheetId('Data');
  hf.addColumns(D, [2, 1]); hf.removeColumns(D, [2, 1]);   // 0-based [index, amount]
  hf.moveColumns(D, 4, 1, 1);                               // start, count, target (pre-move coords) — E before B
  hf.addRows / removeRows / moveRows
  hf.getCellFormula({ sheet, row, col });  hf.undo(); hf.redo();
  hf.calculateFormula('=SUM(A1:A3)', sheetId);  hf.validateFormula('=...'); hf.normalizeFormula('=...');
  ```

* **Structural ops compared to Excel** (`hf-vs-excel.mjs`):
  * **895/1,064** identical on the oracle set, and **4,086/4,608** on exhaustive moves.
  * Undo restored every formula correctly.
* **Blocking problems for rewriting:**
  1. **Parse errors are never adjusted.** HyperFormula keeps the text and does not rewrite it. Things it
     cannot parse:
     * unquoted Japanese sheet names (`集計!C1`)
     * the `_xlfn.` / `_xlws.` prefixes, which xlsx files always contain
     * 3D refs, structured refs, spill `A1#`
     * the union `(A1,B2)` and intersection `A1:C3 B2:D4` operators
     * `A1:INDEX(...)`
     * `Sheet1!#REF!`, `1.` (number), and `#GETTING_DATA`

     These formulas silently keep stale references.
  2. **`OFFSET(Data!C1,0,1)` is rewritten to `Data!D1`** on any structural op. It changes the user's formula.
  3. Default `maxRows = 40000`, `maxColumns = 18278`. `SUM(Data!C1:C1048576)` becomes `#NAME?` and is not
     adjusted. Inserting before XFD produces invalid refs such as `XFE1`.
  4. It uses a different move model: a moved endpoint shrinks the range. For example, B before A turns
     `B1:C1` into `C1:C1`, while Excel gives `A1:C1`.
  5. On delete it drops the sheet prefix (`#REF!` where Excel gives `Data!#REF!`).
  6. It normalizes text: spacing changes, and `Data!E1:Data!C1` is reordered to `Data!C1:Data!E1`.
* **As an evaluator:**
  * Bare `TRUE` / `FALSE` literals evaluate to **#NAME?** ("Named expression TRUE not recognized"). This
    breaks `VLOOKUP(...,FALSE)` and `TEXTJOIN(",",TRUE,...)`.
  * No `LET` and no `CONCAT`.
  * DATEDIF needs an upper-case unit.
  * It needs `_xlfn.` stripped and Japanese sheet names quoted.
  * Unknown functions return #NAME?, but their references *are* adjusted.
* **Verdict:** (a) **no**. (b) Possible for previews, but IronCalc does better.

### 2.3 IronCalc `@ironcalc/wasm` 0.8.4 (MIT/Apache-2.0)

* The engine is 1,928 KB raw / 659 KB gz wasm. It has no xlsx I/O in the wasm build.
* API:

  ```js
  import init, { Model } from '@ironcalc/wasm';
  await init({ module_or_path: fs.readFileSync('node_modules/@ironcalc/wasm/wasm_bg.wasm') }); // browser: URL or bytes
  const m = new Model('wb', 'en', 'UTC', 'en');
  m.renameSheet(0,'集計 シート'); m.newSheet();
  m.setUserInput(sheet, row1, col1, '=XLOOKUP(...)'); m.evaluate(); m.getFormattedCellValue(sheet,row,col); m.getCellContent(...)
  m.insertColumns(sheet, col, n); m.deleteColumns(...); m.moveColumns(sheet, col, count, delta); m.undo()
  ```

* **Structural ops compared to Excel:** 847/1,064 identical, and 3,784/4,608 on exhaustive moves.
  * Deleting a range endpoint gives `#REF!:C1`.
  * `$1:$1048576` is rendered as `Data!:` (a bug).
  * Moves use pure endpoint-follow.
  * Union and intersection are not adjusted.
  * **Not usable for rewriting.**
* **As an evaluator:** 14/14 in `eval-matrix.mjs`. This is the best of the four evaluators tested.

### 2.4 Evaluator matrix (`eval-matrix.mjs`; data on sheet `'集計 シート'`)

| test | HyperFormula | fast-formula-parser | IronCalc | @xlscalc |
|---|---|---|---|---|
| SUM, whole column, IFERROR | OK | OK | OK | OK |
| VLOOKUP(...,FALSE) | **#NAME? (bare FALSE)** | OK | OK | OK |
| XLOOKUP / INDEX+MATCH / SUMIFS / COUNTIFS | OK | throws (missing) | OK | OK |
| TEXTJOIN | #NAME? (bare TRUE) | missing | OK | OK |
| LET | missing | missing | OK | unsupported |
| TEXT(date,"yyyy/mm/dd") | OK | OK | OK | unsupported |
| DATEDIF(...,"d") | #NUM (lower-case unit) | OK | OK | OK |
| CONCAT | missing | OK | OK | OK |
| SUM(FILTER(...)) | OK | missing | OK | unsupported |
| size | 192 KB gz | 83 KB gz | 659 KB gz wasm | ~700 KB unpacked |

* `@formulajs/formulajs` (MIT) is only function implementations with no parser. Use it only inside a custom
  evaluator.
* `fast-formula-parser` supports 280 functions but has no XLOOKUP, FILTER, LET or MATCH.

### 2.5 ExcelJS (xlsx I/O layer). These problems affect how our engine must be used.

* **`ws.spliceColumns()` moves cells but does not change any formula text.** After inserting a column at A,
  the formula `IF(A2=…)` now sits in C2 but still says `A2`. Every structural op must therefore be followed
  by our `adjustFormula` on every formula, name, CF, DV and chart.
* **Shared-formula translation (`lib/utils/shared-formula.js` `slideFormula`) is regex-based and wrong:**
  * `slideFormula('IF(A2="A1",1,2)','A2','B3')` → `IF(B3="B2",1,2)`: it **shifts refs inside string
    literals**.
  * `SUM(B:B)` is **not** shifted. In a real Excel file, H2 of a shared group whose master G2 is
    `SUM(B:B)+COUNTIF(D2:D9,…)` was reported as `SUM(B:B)+…`; the correct value is `SUM(C:C)`.
  * Fix: expand shared formulas on load with our `translateFormula(master, dRow, dCol)`, which was verified
    on the same file.
* Excel wrote most filled-down formulas as individual formulas here, but `FillRight` produced
  `<f t="shared" ref="G2:H2" si="0">`. Expect both forms.

### 2.6 Others (reference-only)

* **Univer** (`@univerjs/engine-formula`, Apache-2.0): 40 MB unpacked, with its own lexer and ref-move
  logic. It is too heavy to embed, so it was not tested. It could serve as a design reference.
* SpreadJS and Syncfusion: commercial.
* XLParser (C#): a grammar reference (99.9% compatibility paper).
* `cellref`, `@segment/a1-notation`: trivial A1 helpers. Not needed.

---

## 3. Excel's reference-adjustment semantics (verified with Excel 365)

Notation is 1-based. An op acts on one sheet (the *target sheet*). A reference is adjusted when its
resolved sheet is the target sheet: a sheet prefix equal to the target (case-insensitive), or no prefix and
the formula lives on the target sheet. **External (`[1]Sheet!A1`) and 3D (`Data:Mid!C1`) refs are never
adjusted.** Excel confirmed this: `SUM(Data:Mid!C1)` was unchanged by every column op on Data.

### 3.1 Insert `count` columns before `at`

Each endpoint is handled independently: `v >= at ? v + count : v`.
* A range therefore **expands** when `lo < at <= hi`, and **shifts** when `at <= lo`. Inserting exactly at
  the left edge shifts the range; it does not expand it.
* A single cell pushed past XFD becomes `#REF!`. A range end is clamped to XFD and then normalized
  (`A1:XFC1` plus an insert gives `1:1`).
* Whole-row refs are unaffected by column ops; whole-column refs are unaffected by row ops.

### 3.2 Delete columns `[s1, s2]`

* A cell in the deleted block becomes `Sheet!#REF!`. The prefix is kept.
* A range entirely inside the deleted block becomes `Sheet!#REF!`.
* Otherwise:
  * `lo' = lo < s1 ? lo : lo > s2 ? lo-n : s1`
  * `hi' = hi > s2 ? hi-n : hi >= s1 ? s1-1 : hi`
  * The range shrinks and can become degenerate: `B1:D1` minus C:D gives `B1:B1`, which Excel keeps as
    `B1:B1`.
* **Sticky edge:** a range that reaches the last row or column stays anchored there. `$1:$1048576` minus
  row 3 is still `$1:$1048576`.

### 3.3 Move a block (Excel "Cut columns → Insert Cut Cells")

The op is `move(from, count, to)`, where `to` is the insert-before index in **original** coordinates. This
is exactly `Columns(from:from+count-1).Cut ; Columns(to).Insert`.

Let `p` be the index permutation. For a span `[lo, hi]`:

```
follow = [p(lo), p(hi)]                          (each endpoint follows its own cell)
modelA = delete the block, then insert `count` at the destination (post-delete coords)
1. span entirely inside the moved block                               -> follow
2. follow inverted (p(lo) > p(hi))                                    -> modelA
3. block entirely inside span AND an endpoint lying in the block
   moves INWARD (left end moving right / right end moving left)       -> modelA
4. otherwise                                                          -> follow
```

For spans that are disjoint from the block, `follow == modelA`, which is a plain shift or expansion.

This rule reproduces **all 4,608** exhaustive Excel results. It was also re-checked on the 16-op oracle.

Examples, all taken from Excel:

| moved block | range | result |
|---|---|---|
| E before B | `B1:E5` | `C1:E5` |
| E before B | `E1:F1` | `B1:F1` |
| E before B | `A1:E1` | unchanged |
| C:D before B | `A1:C1` | `A1:B1` |
| C:D before B | `B1:C1` | `D1:D1` |
| A:B before D | `B1:D1` | `C1:D1` |
| A:B before D | `B1:H1` | `C1:H1` |

**Multi-selection (non-contiguous) moves.** Excel cannot do these in one step, so `planMove(selected,
target)` decomposes them:
* Each maximal contiguous run becomes **one block move**.
* Left runs are moved right-to-left, each placed immediately before the previously placed run; then right
  runs are moved left-to-right.
* It was verified on 4,943 random selections.
* Example: `[B,E,F,I]` before D gives
  `[{from:2,count:1,to:4},{from:5,count:2,to:4},{from:9,count:1,to:6}]`.

The same op list must be replayed by the Excel-COM engine, so both engines agree.

### 3.4 Other observed behaviour (all from Excel)

* `Sheet!A1:Sheet!B2` (both halves qualified) is **two independent cell refs** joined by `:`. After E moves
  before B, `Data!C1:Data!E1` becomes `Data!D1:Data!B1`: it is not normalized and not treated as a range.
  **Never merge such tokens into one area.**
* `Data!B1:INDEX(Data!D:D,1)`: the left part is a cell and the right part is a function. Each ref is adjusted
  separately. Deleting B gives `Data!#REF!:INDEX(Data!C:C,1)`.
* Union `(Data!B1,Data!D1)` and intersection `Data!B1:D5 Data!C1:C10`: each ref is adjusted separately.
* Spill `Data!H20#` is adjusted like a cell.
* `INDIRECT("Data!C1")`, `"Data!C1"` string literals and `HYPERLINK("#Sheet2!A1")` are **not** adjusted,
  and Excel does not adjust them either.
* `VLOOKUP(Data!A1,Data!A:E,3,FALSE)`: the range is adjusted, but **the column index 3 is not**. The same
  applies to INDEX, OFFSET, MATCH and CHOOSE numeric arguments.
  * Moving columns inside a lookup table changes what VLOOKUP returns.
  * The app should warn. The prototype emits `warnings`.
* **Defined names** follow exactly the same rules as cell formulas. Workbook and sheet scope behave the same.
* **Chart `SERIES`:** the value refs adjust normally. Deleting the category column makes Excel drop the
  argument (`=SERIES(,,Data!$B$1:$B$5,1)`). Deleting the value column breaks the series (`Formula = ""`).
* **Entry normalization** happens only when a formula is typed or written through COM, not in stored files:
  * `C1:C1048576` becomes `C:C`
  * `A1:XFD1` becomes `1:1`
  * `A:XFD` becomes `$1:$1048576`

### 3.5 Deliberate deviations in the prototype (the 20 "knownDeviation" results)

* **18 cases:** Excel implements a move as insert + cut + delete. Any ref to column XFD, to a range ending at
  XFC/XFD, or to row 1048576 is therefore pushed off the sheet and becomes `#REF!` *during a move*. We keep
  such refs intact, because a move must never create `#REF!`.
* **2 cases:** chart SERIES dropping a deleted category argument. This is chart policy: the app should warn
  before deleting instead.

---

## 4. Conditional formatting / data validation with relative refs

In xlsx, CF and DV formulas are **relative to the top-left cell of the sqref**. Findings from `cf-results.json`:

* Excel treats each rule as **one formula at the anchor**.
* When an insert or delete splits an sqref area, Excel **splits the rule into fragments**. Each fragment is
  re-anchored at its **old** top-left cell and adjusted like a normal formula hosted there.
* When inserting strictly inside an area, the **first fragment absorbs the inserted columns**, so the new
  column inherits the rule.
* Example (delete C on `B2:E10`): `=C2>0` becomes two rules, `B2:B10 | =#REF!>0` and `C2:D10 | =D2>0`.
* On a **move**, Excel over-fragments. One move produced four single-column rules. Its results also contain
  `#REF!` artifacts from its internal insert+cut+delete. For example, after E moves before B, `=A2>0` gives
  `#REF!` for a cell whose left neighbour still exists.

Prototype (`adjustAnchoredRule(rule, op, sheet)`):
* **insert/delete:** Excel's fragment model. Per-cell effective formulas equal Excel's in every case tested.
* **move:** per-cell model (`adjustAnchoredRulePerCell`). Each cell's copy is adjusted, and cells are
  regrouped by R1C1 key. The cost is O(segments), so whole-column sqrefs are fine. It matches Excel except
  for the 2 artifacts above.
* The test compares **per-cell effective R1C1 formulas**, not rule text, because the fragmentation is
  arbitrary. Result: 62/62 (+2 known artifacts).

The COM `FormatCondition.Formula1` is expressed **relative to the ActiveCell**. Select the top-left cell of
`AppliesTo` before reading or writing it.

---

## 5. Token grammar (A1 mode; xlsx storage form and UI form)

```
formula    := ['='] token*                          (xlsx <f> has no '=')
ws         := [ \t\r\n]+                            (a single space between refs is the intersection operator — copy verbatim)
string     := '"' ( [^"] | '""' )* '"'
number     := ( D+ ['.' D*] | '.' D+ ) [ [eE] [+-]? D+ ]
bool       := TRUE | FALSE  (case-insensitive, NOT followed by '(')
error      := [prefix] ( #NULL! #DIV/0! #VALUE! #REF! #NAME? #NUM! #N/A #GETTING_DATA #SPILL! #CALC! #FIELD!
                         #BLOCKED! #CONNECT! #BUSY! #UNKNOWN! #PYTHON! #EXTERNAL! )   (case-insensitive)
func       := [prefix] ident   immediately followed by '('    (_xlfn._xlws.FILTER, _xludf.X, LAMBDA(...)(...))
ref        := [prefix] area    not followed by an ident char, '(', '!', '['
area       := cell [ ':' cell ] | col ':' col | row ':' row
cell       := ['$'] [A-Za-z]{1,3} ['$'] [0-9]{1,7}      with 1 <= col <= 16384 (XFD), 1 <= row <= 1048576
col        := ['$'] [A-Za-z]{1,3}   (<= XFD)          row := ['$'] [0-9]{1,7} (1..1048576)
sref       := [prefix] [ident] '[' ... ']'   balanced brackets; "'" escapes the next char  (Table1[[#This Row],[氏名]], [@列], Table1['#Col])
name       := [prefix] ident                (defined names, LET/LAMBDA params incl. _xlpm.x, table names)
prefix     := "'" ( [^'] | "''" )* "'" '!'   (content: Sheet | S1:S2 (3D) | [1]Sheet | [Book.xlsx]Sheet | C:\p\[Book.xlsx]Sheet)
            | '[' [^\[\]]+ ']' [sheetIdent] '!'          ([1]Sheet1!A1, [1]!Name)
            | sheetIdent [':' sheetIdent] '!'            (集計!A1, LOG10!B2, Sheet1:Sheet3!A1)
ident      := identStart identChar*
identStart := \p{L} | '_' | '\' | any non-ASCII char except full-width punctuation (、。（）etc.)
identChar  := \p{L} \p{N} \p{M} '_' '.' '\' '?' '$' '・' ①-⑳ … (non-ASCII, non-space, non-punct)
op         := + - * / ^ & = <> <= >= < > % : , ; ( ) { } @ # !
anything else -> 'unknown' (caller must refuse to rewrite / route to the Excel engine)
```

**Disambiguation order at an identifier start** (this order is what makes `LOG10`, `TAX2025`, `A1B` etc.
come out right):

1. The run is followed by `!` → sheet prefix. So `LOG10!B2` is a sheet `LOG10`.
   Excel writes that name **unquoted** even though it looks like a cell.
2. `run ':' run '!'` → 3D prefix.
3. The run is followed by `(` → function. So `LOG10(100)` is a function.
4. The text matches `area`, **and** the column is ≤ XFD, the row is ≤ 1048576, **and** the next char is not
   an identifier char, `(`, `!` or `[` → ref. So `LOG10`, `TAX2025`, `XFD1048576` and `R1` are cells.
   `XFE1`, `A1048577`, `A0`, `A1B` and `R1C1` are names.
5. `TRUE` / `FALSE` → bool.
6. The run is followed by `[` → structured reference.
7. Otherwise → name.
   * Note: `x:y` is the column range X:Y. Excel parses it the same way.

**At a digit:** `\$?\d+:\$?\d+` within bounds → row range; otherwise → number. Unquoted sheet names cannot
start with a digit (Excel quotes `'2024'` and `'１月'`).

**Sheet-name quoting observed in Excel 365:**

* **Unquoted:** `集計`, `売上・集計`, `表Ａ`, `あいう`, `アイウ`, `Ｓｈｅｅｔ`, `項目①`, `ab_c`, `_x`,
  `Sheet.1`, `LOG10`.
* **Quoted:**
  * names with a space, `-`, `（）` or `、`
  * names starting with a digit, including full-width digits: `'2024'`, `'１月'`
  * cell-like or R1C1-like names: `'A1'`, `'TAX2025'`, `'R'`, `'C'`, `'RC'`, `'R1C1'`, and the test sheets
    `'C7'` and `'S7'`
  * apostrophes (doubled)
* A 3D range is quoted as a whole: `'Sheet1:TAX2025'!A1`.
* **The xlsx writer and the UI differ.** `a.b.c` is quoted in the xlsx XML but not in `Formula2`.

→ **Never re-quote a prefix you did not change.** Copy prefixes verbatim. `quoteSheetName()` is
conservative: it quotes everything Excel quotes, as the test shows.

**How modern syntax is stored in xlsx** (`storage-test.xlsx`, `sheet29.xml`). The tokenizer handles all of these:

| typed in Excel | stored in `<f>` |
|---|---|
| `=XLOOKUP(...)` | `_xlfn.XLOOKUP(...)` |
| `=FILTER(...)` | `<f t="array" ref="E3:E4">_xlfn._xlws.FILTER(...)</f>` with `cm="1"` (dynamic array) |
| `=SUM(E3#)` | `SUM(_xlfn.ANCHORARRAY(E3))` (no `#` in files) |
| `=@A2:A3` | `A2:A3` (no `@`; `.Formula` also drops it) |
| `=LET(x,B2,y,B3,x+y)` | `_xlfn.LET(_xlpm.x,B2,_xlpm.y,B3,_xlpm.x+_xlpm.y)` |
| `=LAMBDA(a,a+1)(B2)` | `<f t="array" ref="E7">_xlfn.LAMBDA(_xlpm.a,_xlpm.a+1)(B2)</f>` |
| `=Table1[@金額]` / `=[@金額]` | `Table1[[#This Row],[金額]]` (also in `<calculatedColumnFormula>`) |
| `=TEXTJOIN(...)`, `=CONCAT(...)` | `_xlfn.TEXTJOIN`, `_xlfn.CONCAT` |
| `=INDIRECT(...)` in 365 | `<f t="array" aca="1" ref="E21" ca="1">` |

If the app lets users type modern functions, it must **add `_xlfn.` when writing**. Otherwise Excel shows
#NAME?. This needs a list of "future functions"; ExcelJS does not do it.

---

## 6. Prototype: `r2\proto\excelRef.ts`

Run `npm test` in `r2\proto` (Node 24 runs `.ts` via `--experimental-strip-types`). Current output:

```
tokenizer: pass=39800 fail=0; fuzz formulas=13168; differential mismatches vs fx=0 (fx produced 'unknown' tokens on 71 formulas)
perf: 11675-char formula (1200 refs): tokenize 2.08 ms, adjust 2.32 ms
oracle: pass=1393 fail=0 knownDeviation=20
exhaustive move: pass=4608 fail=0
cf/dv per-cell: pass=62 fail=0 knownExcelArtifacts=2; planMove random trials ok=4943
misc: pass=59 fail=0; 50,000 formulas adjusted in 795 ms (15.9 us/formula)
```

Tests:

| file | what it checks |
|---|---|
| `tests/tokenizer.test.ts` | torture list, UI-form extras, differential vs `@borgar/fx`, grammar fuzzer, translate, perf |
| `tests/oracle.test.ts` | 16 Excel ops × formulas / names / CF formulas / CF sqref / series |
| `tests/exhaustive.test.ts` | 128 moves × 36 ranges |
| `tests/cf.test.ts` | per-cell CF/DV + `planMove` |
| `tests/misc.test.ts` | quoting vs Excel, rename, #REF! detection, 50k-formula perf |

API and usage:

```ts
import { tokenize, adjustFormula, adjustFormulaSeq, planMove, adjustAnchoredRule, adjustSqref,
         translateFormula, renameSheetInFormula, quoteSheetName, listRefs, parseAreaText, renderArea } from './excelRef';

// user drags selected columns B,E,F,I in sheet '結合結果' and drops them before column D
const ops = planMove([2, 5, 6, 9], 4);            // StructOp[] (also replayed by the Excel-COM engine)

for (const op of ops) {
  for (const sh of workbook.sheets)                // every sheet: cross-sheet refs must follow too
    for (const cell of sh.formulaCells) {
      const r = adjustFormula(cell.formula, op, { hostSheet: sh.name, targetSheet: '結合結果' });
      if (r.unsafe) { /* unknown token: leave untouched + warn, or route to Excel engine */ }
      if (r.changed) { undoLog.push([cell.id, cell.formula]); cell.formula = r.formula; }
    }
  for (const dn of workbook.definedNames)          // names: hostSheet = null (refs are always prefixed)
    dn.formula = adjustFormula(dn.formula, op, { hostSheet: null, targetSheet: '結合結果' }).formula;
  // CF / DV on the target sheet: split/re-anchor; on other sheets: only adjustFormula(hostSheet=thatSheet)
  target.cf = target.cf.flatMap(rule => adjustAnchoredRule(rule, op, '結合結果'));
  // chart <c:f>, pivot source, print area (_xlnm.Print_Area): adjustFormula with hostSheet null
  // mergeCells / autoFilter / table ref / hyperlink ref / array-formula ref: adjustSqref or adjustArea
}

// before a DELETE: dry run, warn if anything would become #REF!
const wouldBreak = allFormulas.some(f => adjustFormula(f.text, delOp, ctx(f)).createdRefError);
```

Implementation notes:
* **Undo/redo:** adjustments are **not invertible** (shrinking a range and `#REF!` lose information). Store
  `(cellId, oldFormula, newFormula)` diffs per user action instead of inverse ops.
* **Caching:** `tokenize` is pure, so it can be cached per formula string if hundreds of ops are replayed.
  In practice it is already fast enough.
* **Shared formulas:** expand them on load with `translateFormula(master, dRow, dCol)` and write plain
  formulas back. Excel accepts non-shared formulas. Do **not** use ExcelJS `slideFormula` (§2.5).
* **Array / dynamic formulas:** adjust the `ref` attribute with `adjustSqref` and keep `cm`, `t`, `aca`
  and `ca` as they are.

**Inventory of reference-bearing places in an xlsx** (all need `adjustFormula` or `adjustSqref`).
Implementers should tick these off:

* **`sheetN.xml`**
  * `<c><f>`, including `t="shared" ref=` and `t="array" ref=`
  * `<mergeCell ref>`
  * `<conditionalFormatting sqref>` / `<cfRule><formula>`
  * `<dataValidation sqref>` / `<formula1/2>`
  * `<hyperlink ref location="Sheet!A1">` — `location` is formula-like text
  * `<autoFilter ref>`, `<dimension ref>`
  * `<selection sqref activeCell>` and `<pane topLeftCell>` (can simply be reset)
  * `<cols>` widths and styles (permute them)
  * `<rowBreaks>` / `<colBreaks>`, `<protectedRanges sqref>`, `<ignoredErrors sqref>`
  * **`<extLst>`**: `x14:conditionalFormatting` / `x14:dataValidation` with `<xm:f>` + `<xm:sqref>`, and
    `x14:sparklineGroups` (`<xm:f>`, `<xm:sqref>`). Excel stores DV lists that point to other sheets and
    data-bar/icon CF here. It is easy to forget.
* **`tables/tableN.xml`:** `ref`, `autoFilter ref`, the `tableColumns` order and count, and
  `calculatedColumnFormula`. Structured refs are name-based and need no rewrite.
* **`workbook.xml`:** `<definedName>`, including `_xlnm.Print_Area`, `_xlnm.Print_Titles` and
  `_xlnm._FilterDatabase`.
* **`charts/chartN.xml`:** `<c:f>` in `tx`, `cat`, `val`, `xVal`, `yVal` and `bubbleSize`. The caches can
  stay; Excel refreshes them.
* **Drawing anchors** (`xdr:from`/`xdr:to` col and row), and comments (`<comment ref>` + VML anchors).
* **`pivotCacheDefinition`:** `<worksheetSource ref sheet>`.
* **`calcChain.xml`:** delete it together with its rel and content type. Set `<calcPr fullCalcOnLoad="1"/>`.

---

## 7. Excel-COM engine notes (verified)

**Clipboard-free move (3-step).** It gave the same result as Cut+Insert on 3,492/3,492 formula comparisons,
and on names, CF, DV and series:

```powershell
# move columns [from .. from+n-1] so they end up before column `to` (original coords); $L = column letter fn
$blk = "$(L $to):$(L ($to+$n-1))"
$ws.Columns($blk).Insert() | Out-Null                               # 1) make room at the destination
if ($to -lt $from) { $src = "$(L ($from+$n)):$(L ($from+2*$n-1))" }  #    (source shifted right by n)
else               { $src = "$(L $from):$(L ($from+$n-1))" }
$ws.Columns($src).Cut($ws.Columns($blk)) | Out-Null                 # 2) Range.Cut(Destination): no clipboard
$ws.Columns($src).Delete() | Out-Null                               # 3) remove the emptied source columns
```

**`Columns(x).Cut()` + `Columns(y).Insert()` (clipboard) is unreliable.**
* In the verified run 31/128 moves inserted **blank columns** instead of the cut cells, with no error.
  An earlier unverified attempt also died with `RPC_E_CALL_REJECTED`.
* Any other process touching the clipboard between the two calls cancels cut mode. In production that could
  be the user's own apps.
* The only way to detect it is to verify with marker cells.

**Other COM pitfalls hit:**
* `$xl.CutCopyMode = $false` throws in PowerShell: it cannot convert to `XlCutCopyMode`. Use
  `InvokeMember('CutCopyMode', SetProperty, …, @($false))` or ignore it.
* `$xl.Calculation = -4135` throws `0x800A03EC` if no workbook is open.
* `RPC_E_CALL_REJECTED` (0x80010001) can occur mid-run. Wrap calls in a retry (250 ms × 20).
* After `Quit()` + `FinalReleaseComObject`, EXCEL.EXE often lingers for about 1 s. Kill only your own PID,
  obtained from `GetWindowThreadProcessId($xl.Hwnd)`.
* Excel re-quotes sheet names it considers cell-like (`'C7'!`, `'S7'!`). Text search-and-replace on formulas
  therefore breaks. Always use the tokenizer.
* Reading `.Formula` drops `@`. Use `.Formula2` for dynamic-array-aware text.

---

## 8. Pitfalls checklist

1. Do not use string or regex replace on formulas. Strings (`"A1"`), sheet names that look like cells
   (`LOG10!`), names that look like cells (`TAX2025` *is* a cell, `XFE1` is not) and structured refs all break
   regexes. ExcelJS's `slideFormula` shows the damage: it shifts inside strings and skips `B:B`.
2. Keep `Sheet!A1:Sheet!B2` as two refs. Excel adjusts the endpoints independently.
3. Keep verbatim everything you do not rewrite: prefix quoting, case, spacing, `_xlfn.`, `_xlpm.`.
4. 3D and external refs are never adjusted by single-sheet ops.
5. INDIRECT/ADDRESS text and VLOOKUP/HLOOKUP/INDEX/OFFSET/CHOOSE numeric positions are **not** adjusted,
   by us or by Excel. Warn when moving columns inside a lookup range.
6. A delete that hits a referenced cell, a whole referenced range or a chart series produces `#REF!` in
   Excel. Dry-run it (`createdRefError`) and block or warn to satisfy the no-`#REF!` requirement.
7. CF/DV relative refs are anchored at the sqref's top-left cell. Rules split into fragments on
   insert/delete, and a column inserted inside a CF range inherits the rule.
8. Moves of multiple non-contiguous columns must be decomposed into contiguous block moves, and both
   engines must replay the **same** list (`planMove`).
9. xlsx storage form ≠ UI form: `_xlfn.ANCHORARRAY(X)` is `X#`; `@` is dropped; `[@c]` is
   `Table1[[#This Row],[c]]`; `_xlpm.`. When the app writes a new modern function it must add `_xlfn.`.
10. HyperFormula as an engine: parse errors silently keep stale refs; it rewrites OFFSET; it has a 40,000
    maxRows default; bare TRUE/FALSE are #NAME?; the license is GPL.
11. Excel's own edge quirks: refs to XFD or row 1048576 become `#REF!` during moves; full-width ranges are
    normalized to `1:1` / `A:A`; a range reaching the last row stays anchored on delete. The prototype copies
    the last two and deliberately not the first.
12. The Excel-COM engine must be clipboard-free and must use its own instance.

---

## 9. Files

* Prototype: `r2\proto\excelRef.ts`, `r2\proto\tests\*.test.ts`, `r2\proto\README.md`, `r2\proto\package.json`
* Oracles:
  * scripts: `r2\oracle\excel-oracle.ps1`, `excel-exhaustive.ps1`, `excel-cf.ps1`, `excel-storage.ps1`,
    `excel-shared.ps1`
  * results: `excel-results.json`, `exhaustive-results.json`, `cf-results.json`, `storage\*.xlsx`,
    `storage\storage-results.json`
  * readable diff table: `table.txt`
* Library harnesses: `r2\compare.mjs` (→ `compare-results.json`), `r2\hf-vs-excel.mjs`,
  `r2\ironcalc-vs-excel.mjs`, `r2\eval-matrix.mjs`, `r2\exceljs-shared.mjs`, `r2\torture.mjs`
