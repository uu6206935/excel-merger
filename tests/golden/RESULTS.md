# Excel golden harness — results (patcher P)

Date: 2026-09-30. Excel: Microsoft 365 (16.0, Japanese UI) driven through COM from Windows PowerShell 5.1, own hidden
instance per script (PID recorded, only that PID is ever killed). Node 24, vitest 3.

## What is compared

`replayLog(fixture, ops)` against the workbook **Excel itself saved** after the same ops. Both are reduced to an XML
snapshot (`tests/golden/snapshot.ts`) that does not depend on how a writer numbers records:

- cell formulas (shared formulas expanded; sheet-prefix quoting normalised), non-formula values (shared strings
  resolved), dynamic-array spill values excluded (recalculation results);
- cell styles as resolved signatures (number-format code + font + fill + border + alignment), never xf indexes;
  a style equal to xf 0 counts as "no style";
- CF / DV as **per-cell effective rules** (type, dxf, formula re-expressed relative to the cell, R1C1-like), so
  different fragmentations of the same rules compare equal;
- defined names, tables (ref, autoFilter, columns, calculated columns), chart `<c:f>`, pivot cache sources, merges,
  column widths / hidden / outline, custom row heights, frozen panes, autoFilter (+ filter columns), hyperlinks,
  notes, threaded comments, drawing anchors (from/to cells), sparklines, dimension;
- `calcChain` removed and `fullCalcOnLoad="1"` set.

`tests/golden/golden.test.ts` runs every case without Excel (fixtures + Excel's saved results are in
`tests/fixtures/golden`). Documented deviations (`known` in `tests/golden/cases.ts`) are asserted to still differ, so
they cannot go stale silently.

## Fixtures

| file | built by | contents |
|---|---|---|
| `torture.xlsx` | R1 `build-torture.ps1` | Data: merge A1:F1, frozen C4, 5 CF types incl. x14 dataBar, DV whole + x14 cross-sheet list, shared formulas, CSE + dynamic arrays, names (whole column, sheet scope, print area/titles, _FilterDatabase), chart, picture, shape, note, threaded comment, hyperlinks, hidden + grouped columns, autofilter, pivot; sheets 集計 シート / Tbl (ListObject + totals) / Pivot / Lists |
| `g-table.xlsx` | `com/make-fixtures.ps1` | 名簿: ListObject tblMain A3:G13 (calculated columns), 合計 row 15 below it, CF (relative expression, cellIs), DV (x14 list + whole), names (absolute, whole-column, sheet-scoped, relative), frozen B4, hidden H, grouped J:K, merge I1:K1, note, hyperlink; Summary sheet with formulas / structured refs / a chart |
| `pm-base-plain.xlsx`, `pm-base-table.xlsx`, `pm-src.xlsx` | `com/make-fixtures.ps1` | premerge base (AutoFilter or ListObject, 合計 row, CF, name, chart on another sheet) and source (styled columns, custom number formats, formulas) |
| `probe/*.xlsx` | `com/probe-*.ps1` | before / after files of the behaviour probes below |
| `r2-arrays.xlsx`, `r2-rich.xlsx`, `r2-table.xlsx` | `com/make-fixtures-r2.ps1` | round 2: legacy CSE array + SEQUENCE / string spills (+ a sheet referring to them); rich text runs + furigana source; ListObject `Tbl` with calculated column, totals row (custom `totalsRowFormula` + SUBTOTAL), a defined name and another sheet using structured refs |

## Golden cases (`tests/golden/cases.ts`) — 66 cases (+ the known-deviation check), all passing

Round 1 cases below; round 2 cases in the next section.

| case | ops (Excel: clipboard-free moves) | result |
|---|---|---|
| t-ins-C, t-ins-A2, t-ins-K3 | insert columns (inside merge / at A / after the bordered table) | identical |
| t-del-C, t-del-EF, t-del-B | delete columns (x14 DV source, #REF! everywhere, note + threaded comment column) | identical |
| t-mov-HI-G, t-mov-G-K, t-mov-L-H, t-plan-GIJ-H | block moves / planMove of a non-contiguous selection | identical (incl. CF leak per rule type, anchors, freeze) |
| t-ins-row1, t-ins-row10, t-del-row5, t-del-rows20, t-del-rows3, t-mov-rows | row inserts / deletes / move | identical (spill range collapsed, see below) |
| t-seq | insert, delete, move, insert row, column width | identical |
| t-cells | cell writes before / after an insert | identical except `AUTO_NUMFMT` (below) |
| g-ins-D, g-ins-G, g-ins-A | insert inside / right after / before a table | identical (列1 table column + header cell) |
| g-del-E, g-del-HJ, g-del-B-row5 | delete table column (structured refs -> #REF!), delete across a merge, note cell | identical |
| g-mov-D-B, g-plan-ACE-F, g-mov-L-A | moves inside a table / left of it | identical except `TABLE_CLIPFREE`, `CF_INSIDE_MOVE` |
| g-ins-row8, g-ins-row13, g-ins-row14, g-del-rows6 | rows inside / at the end of / after a table | identical (calculated columns filled) |
| g-ins-row4 | row at the FIRST data row of a table | `TABLE_FIRST_ROW` deviation |
| g-mov-I-A, g-mov-L-C, g-mov-B-H | move splitting a merge / into / out of a table | refused with a Japanese message (Excel refuses too) |
| p-borders, p-cf-ins, p-cf-mov, p-cf-inside | probe files | identical except `CF_INSIDE_MOVE` |

### Documented deviations

- **CF_INSIDE_MOVE** — a block moved *inside* the range of one CF rule: Excel (insert + `Range.Cut(Destination)` +
  delete) shrinks the rule to the columns between destination and source (A2:F10 with D moved before B -> B2:D10,
  `probe-cf-move.ps1`). replayLog keeps cell-set semantics (every cell keeps its rule; excelRef per-cell model).
- **TABLE_CLIPFREE** — moving columns inside a ListObject with the clipboard-free sequence deletes and re-creates the
  table column in Excel, so structured references become `#REF!`. Native Insert Cut Cells reorders the table
  columns and keeps the references (R2b R11); replayLog does that (and the Excel runner uses native Cut+Insert
  for moves inside a table).
- **TABLE_FIRST_ROW** — inserting a sheet row at the first data row of a table: Excel grows ranges that exactly
  cover the table's data rows (and whose columns are inside the table) instead of shifting them, and formats the
  new row from the data rows (`probe-table-row.ps1`: `SUM(C4:C10)` -> `SUM(C4:C11)` while `SUM(C4:C8)`,
  `SUM(H4:H10)`, `C4` shift). replayLog shifts (plain rule). The app never inserts rows there.
- **AUTO_NUMFMT** — a formula written into a General cell gets the number format of its first precedent in Excel
  (UI auto-format). `CellWrite` never changes styles (types.ts), so replayLog keeps the cell's format.
- **SPILL_DIMENSION** (round 2) — the `<dimension>` of a sheet with dynamic arrays covers the spill ranges Excel
  computes when it recalculates. replayLog does not evaluate formulas: a new dynamic array is stored with its master
  cell only, and a spill that depends on an array converted to values keeps its cached extent. Excel re-spills on
  load (fullCalcOnLoad) — checked below with `opentest.ps1 -Cells`.

## Round 2 (2026-09-30)

| case | ops | result |
|---|---|---|
| r-dyn | dynamic-array formulas written by the user (Excel: `Range.Formula2`), then a column insert | identical |
| r-dyn-new | a dynamic array in a book without `xl/metadata.xml` (part + relationship + content type created, exactly the XML Excel writes: XLDAPR type, `dynamicArrayProperties fDynamic=1 fCollapsed=0`, `cellMetadata` `<rc t="1" v="0"/>`) | identical except `SPILL_DIMENSION` |
| r-tovalues | `toValues` over a CSE master, a SEQUENCE spill master and a string spill master (Excel: `rg.Value2 = rg.Value2` on the whole array) | identical except `SPILL_DIMENSION` ([B] refers to the spill) |
| r-edit-spill | a value typed over a spill master: spilled values disappear | identical |
| r-numfmt | `CellWrite.numFmt` (value, then `Range.NumberFormat`): built-in ids for ja-JP built-in codes (yyyy/m/d = 14, 0% = 9, h:mm = 20), custom ids otherwise, identical xfs reused; the standard-width column is widened for the date | identical |
| r-numfmt-w | 13 number formats in standard-width columns (dates, date-time, #,##0, $, %, 11 digits, 年月日, negative, yyyy-mm-dd), an auto-fitted column that must grow again, a user-width column that must not | identical |
| u-torture, u-plain, u-gtable | `unmerge` of a title (`Range.UnMerge` + `HorizontalAlignment = xlCenterAcrossSelection`), then a move the merge blocked | identical (every cell of the area has `horizontal="centerContinuous"`, other alignment kept, empty cells created with the style) |
| h-rename, h-dup, h-clear, h-special, h-space, h-number, h-case, h-then-ops, h-after-del | typing over a header cell of `tblMain` (g-table): rename, duplicate name (-> 点数12), cleared (-> 列2), `a-b#c`, space + parentheses, a number, case-only change, rename then insert / delete, delete then reuse of the deleted name | identical (table column, header text, every structured ref of both sheets, calculated columns) |
| h2-rename, h2-dup, h2-special, h2-then-move | the same on `Tbl` of r2-table: totals row (`totalsRowFormula`, SUBTOTAL), defined name `ScoreCol = Tbl[Score]`, `[[#Headers],[Score]]`, column ranges `[[Score]:[Twice]]`, `x'y[1]#` (escapes `'` `[` `]` `#`), rename + insert + rename | identical |
| h2-totals, h2-totals-clear, h2-totals-fn, h2-totals-values, h2-totals-ops | typing into the totals row of `Tbl`: a label / a number / a formula, clearing, `SUBTOTAL(n, Tbl[own column])` (built-in function), a constant over the totals formulas (`toValues`), a totals formula written before / after a column and row insert | identical (`totalsRowLabel` / `totalsRowFunction` / `totalsRowFormula` of the table part, cell text) |

Facts measured for round 2 (`com/probe-typed.ps1`, `com/probe-numfmt-width.ps1`):

- **Column widening by number formats**: `Range.NumberFormat` applied to a number that no longer fits widens a
  column of standard width (or one with `bestFit`) to fit: 2026/1/1 -> 8.63, 2026/9/30 -> 9.63, 2026/10/10 -> 10.75,
  2026/9/30 9:30 -> 14.13, 1,234,567 -> 8.88, 123,456,789 -> 11, 2026年9月30日 -> 14 (stored
  `<col width bestFit="1" customWidth="1">`); a user width stays (the date shows `####`), a hidden column stays.
  Through COM, setting the format *before* the value, or entering text with `Range.Formula`, does **not** widen.
  replayLog emulates it with a glyph-advance table of 游ゴシック 11 pt (`patch/autowidth.ts`, every probe value
  within 1 px).
- **Typing over a table header**: a number becomes the text header `2024`, a cleared header becomes the first free
  `列N`, a duplicate (case-insensitive) gets `2`, `3` …; structured references everywhere follow the new name.
- **Typing into a table's totals row** (`com/probe-totals.ps1`): a constant (text, number, TRUE, a date text) becomes the
  column's `totalsRowLabel` and the cell holds that text (a number `5` is the string "5"; `Value2 = Value2` over a
  totals formula does the same); a formula becomes `totalsRowFunction="custom"` + `<totalsRowFormula>`, except exactly
  `SUBTOTAL(n,Tbl[this column])` with n = 101 average / 102 countNums / 103 count / 104 max / 105 min / 107 stdDev /
  109 sum / 110 var, which is the built-in function (other numbers, other columns and spaces stay custom); clearing
  the cell removes label / function / formula. replayLog does the same (`Engine.writeTotalsCell`).
- Excel stores a number format code set through `Range.NumberFormat` with escaped literals (`yyyy\-mm\-dd`); the
  snapshot compares codes with literal escapes removed (they display the same).

Verified in Excel after opening our outputs (`opentest.ps1 -Cells`, full recalculation):

- r-dyn-new: `Report!C1 =SEQUENCE(2,2)` spills 1 2 / 3 4 into C1:D2; r-dyn: `E8 =SEQUENCE(3)` spills 1..3,
  `G8 =E8#*2` spills to G10 = 6.
- r-numfmt: `A!I2` shows `2026/9/30` (not `####`), I3 `50%`, I4 `¥1,000`, I6 `9:00`; r-numfmt-w: every value shows
  in full except the user-width column W (`####`, same as Excel).
- r-tovalues: `B!A2 =A!G2#` -> `#REF!` and its old spill cell B!A3 is empty after load (as in Excel's own result),
  `A!G3 = 2x` is a constant.
- Premerge with a rich-text source (`pm-rich`, `com/probe-runs.ps1`): the imported `Bold normal` shows run 1-4 bold
  red / 5-11 regular exactly like the source cell, `山田` keeps its reading `ヤマダ`, and the appended row's key cell
  (base format kept) shows the same runs (the run that relied on the source cell font gets it as `<rPr>`).
- KeyFillEntry `sourceCols` / `joiner` / `splitPart` (`pm-keyfill`): `S009-Z`, `学生Z　Z`, `学生Z`.

## Round 3 (2026-09-30)

Fixture `r3-misc.xlsx` (hand-written by `tools/make-r3-fixture.ts`, then opened and re-saved by Excel through
`run-ops.ps1 -NormalizeRaw`, so both sides start from Excel's own XML): **Band** (merged areas styled on the top-left cell
only: box + fill A1:F1, 3-row area A3:C5 with cells of its own, theme-filled band E7:G7, area without look A9:C9, own border
on C11 — `tests/doc/invariant.test.ts` reads this sheet, keep its layout), **Q** (quote-prefixed and text-format cells),
**Tb** (table `TblQ` A1:E4 = ID Q1 Q2 Q3 合計), **Outline** (left / top / bottom outline A1:F1, left / top outline A3:C5,
theme band with a medium red underline A7:C7).

| case | ops | result |
|---|---|---|
| u-torture, u-plain, u-gtable, u-theme, u-nolook, u-outline, u-underline | `unmerge` (Excel side: `Range.UnMerge`, then Interior + border sides of the top-left cell copied onto the other cells of the area by `run-ops.ps1`) of title bands / areas: theme-colour fill, outlines stored like Excel stores them, an underline only, a top-left cell without look | identical |
| u-band, u-band-center, u-area, u-ownborder | the same for a top-left cell with a box on all four sides | identical except `EDGE_MIRROR` (below) |
| q-text, q-numfmt | `CellWrite.quote`: `'00123`, `'=text` (quote true), `03-1234-5678` (quote false = plain text), a styled cell (xf + prefix), text format `@` with / without quote | identical |
| q-overwrite | a quote-prefixed cell overwritten by text / a number, cleared, a text-format prefixed cell overwritten | identical except `QUOTE_COM` (below) for the three overwritten cells; the cleared cell keeps its style in both |
| h2-totals-q | a constant typed into a totals cell with / without the quote flag | identical (label text, the cell keeps the prefix only when quote is set) |
| n-q1, n-q3, n-id2, n-q1x3, n-h1, n-h1x2 | columns inserted into a table (`Q1` -> `Q12` when `Q2` exists, `Q3` -> `Q4`, after `ID` -> `列1` `列2`, 3 columns right of `Q1` -> `Q12` `Q13` `Q14`, `点数1` -> `点数12` / `点数12` `点数13`) | identical (table part, header cells) |
| w-after-auto | a width set by the user after a number format widened the column, then a wider date | identical (Excel drops `bestFit`: the column stays 9) |
| g-mov-B-I | moving a column of a table one column beyond its last column | refused (was g-mov-B-H: H = last column + 1 is now an accepted in-table reorder, see below) |

Facts measured in Excel for round 3:

- **Unmerge**: Excel's own `Range.UnMerge` leaves every other cell of the area empty and the top-left cell as it was
  (`probe/probe-r3-native-unmerge.xlsx`): nothing spreads natively. Excel draws the borders of a merged area from the
  top-left cell (top, left) and from the bottom-right cell (bottom, right) (`Range.MergeArea.Borders`, `run-ops.ps1
  -MergeProbe`), and an Excel-made outline around a merged area is stored per cell exactly as the spread produces it
  (top / bottom on every cell of a one-row band, left on the first, right on the last; `probe/probe-r3-merged.xlsx`). A
  side set on the top-left cell only (an underline) is therefore invisible while the cells are merged and becomes
  visible when the look is spread — that is the purpose of the op. Formatting the other cells of a merged area before
  `UnMerge` is not possible (the format applies to the whole area): the Excel runner reads the top-left cell and writes the
  others after `UnMerge`.
- **`EDGE_MIRROR`**: writing a border through COM to a cell next to the top-left cell reconciles the shared edge: B3 / B11
  get the top-left cell's existing right / bottom side as their own left / top side (`probe/probe-r3-which.xlsx`), and in
  the golden the top-left cell's interior right / bottom side ends up only on the neighbour. The same lines are drawn;
  replayLog leaves the top-left cell unchanged. Borders set on single cells are not mirrored into the XML of the
  neighbour otherwise (`probe-r3-merged.xlsx` rows 30-32).
- **Quote prefix through COM**: `Formula = "'text"` sets quotePrefix, also over a text format `@` (the apostrophe is not part of
  the text). A prefixed cell keeps the flag when a value is written over it (`Value2`, `Formula`, `ClearContents`;
  `PrefixCharacter` reads empty for a number / an empty cell, but the flag stays in the xf and acts again when text is
  written); only `ClearFormats` / `Clear` remove it, with every other format; `Value2 = Value2` / `TextToColumns` turn
  number-like text into a number. `NumberFormat` cannot be set on a prefixed cell (COM error). Typing in the UI resets the
  flag, and so does the contract (`CellWrite.cell.quote`: anything but `true` leaves no prefix) — `QUOTE_COM` documents
  the gap for the Excel runner. A cleared cell keeps its style (prefix included) in both engines.
- **Column names of inserted table columns** (round 3 goldens `n-*`): left neighbour ending in a number `h1` -> `h2` when free,
  else `h12`, `h13` ... (measured for several columns at once: `Q12 Q13 Q14`, `点数12 点数13`); no number on the left -> the
  first free `列N`. Identical to `doc/tables.ts` `excelNewColumnNames` (a unit test keeps the two rules identical).
- **In-table reorder**: a block from inside a table may land right after the last table column (Excel's native Insert Cut Cells
  keeps it in the table as its last column; `tests/fixtures/formula/table-move-excel-results.json` — `tests/xlsx/patch-round3.test.ts`
  replays all nine probe scenarios and compares table range, header order and row-1 markers with Excel's). The
  clipboard-free golden sequence cannot reproduce it (it deletes and re-creates the table column, `TABLE_CLIPFREE`), so this
  is covered by the unit test, plus Excel opens `r3-mov-table-end*.xlsx` / `r3-mixed.xlsx` without repair.
- `Worksheet.ExportAsFixedFormat` (PDF) hangs in a hidden Excel instance on this PC (the instance had to be killed by PID): do not
  use it in tests.

Excel opens every output without repair: `run-ops.ps1 -OpenDir` on 106 files (every non-refused golden / probe case, 6 premerge
variants each after typical and round-2 edits, `pm-inside-table`, `pm-rich`, `pm-keyfill`, `r3-mov-table-end`, `r3-mov-table-end2`,
`r3-mixed`): **106 / 106 open normally** (one file reported `repair` once after a rejected COM call; the same file opened normally in
a second instance, `opentest.ps1 -Filter "t-mov-*.xlsx"`, log empty).

## Round 4 (2026-09-30)

Golden cases added (all identical to Excel): `q-wrap`, `r-spill-write`, `h2-del-totals-row`, `h2-del-last-data-and-totals`,
`h2-del-totals-ref`, `h2-del-totals-then-write`; refused-by-us cases (no Excel file): `g-mov-row-to-header`, `g-mov-header-row`.
The suite now has 93 cases (+ the known-deviation check); the known deviations are only the semantic ones (CF_INSIDE_MOVE,
AUTO_NUMFMT, TABLE_FIRST_ROW, TABLE_CLIPFREE, QUOTE_COM for text cells).

Facts measured in Excel for round 4:

- **WrapText by itself**: a text constant with a line feed written through `Value2` sets `wrapText="1"` on the cell's xf (LF and CR LF, also
  over a styled cell, a text-format `@` cell, a quote-prefixed cell, together with `numFmt: '@'`); a **lone CR does not**, a formula whose
  result has a line break does not, a cell that already wraps keeps its xf. The JS engine does the same (`Engine.setCell`; the task text
  said "\n or \r": the golden says LF only).
- **Deleting the totals row of a table** (sheet row delete through the last row of the table): ref shrinks, `totalsRowCount` goes
  (`totalsRowShown="0"`), the table's autoFilter covers all remaining rows, every column loses `totalsRowFunction` / `totalsRowLabel` /
  `<totalsRowFormula>`; a structured reference `Tbl[[#Totals],[Score]]` elsewhere keeps its text (no `#REF!`); the former totals row is an
  ordinary data row. Before this the JS engine only shrank the ref and kept `totalsRowCount="1"`: the LAST DATA row became the totals row and
  **Excel stopped responding** on the file (found by the random op-log fuzz below; unit tests `patch-round4-tables.test.ts`).
- **Writing into the spill range of a dynamic array** (not at its master): the master's spill is blocked (#SPILL!); the saved file has the
  master's ref collapsed to its own cell, the old spilled values gone and the written cell as a normal cell. The JS engine did not collapse the
  range: a value / formula inside `ref="P4:P8"` made Excel repair the sheet ("cell information removed"); now it collapses like a structural
  cut (`patch-round4-spill.test.ts`).
- A data row moved to / above the header row of a table, or the header / totals row moved, leaves a table whose header cells differ from its
  column names (Excel repairs it): refused with a Japanese PatchError (Excel's own Insert Cut Cells would refuse or shift the table; the UI
  never moves rows inside tables).
- **Excel's formula limits**: 8,192 characters AND 16,384 bytes of compiled formula (a reference costs about 8 bytes): a chain of 2,600
  references (7,799 characters) makes Excel itself repair / drop the formula although it is shorter than 8,192 characters. The patcher only
  checks the character limit (`MAX_FORMULA_LENGTH`); the doc layer is the place to reject formulas Excel cannot hold.

What the golden comparison normalises now (representation noise, no longer listed as known deviations):

- a border line shared by two neighbouring cells (A1.right = B1.left, H1.bottom = H2.top): both sides are compared as the union of the facing
  sides (`EDGE_MIRROR` is gone; `u-band`, `u-band-center`, `u-area`, `u-ownborder` are identical; `snapshot-edges.test.ts` keeps the rule
  honest: a line in another weight, one cell further away or missing still differs);
- a `quotePrefix` counts only on a text constant (invisible on numbers / booleans / formulas / empty cells): the `B2` entry of `q-overwrite` is
  no deviation any more (`B1` / `E1`, text cells, stay);
- the snapshot reads packages written with a namespace prefix, single-quoted attributes, pretty printing (white space between elements), a BOM
  or UTF-16 parts, Windows separators in relationship targets; table descriptions carry `hdr=` / `tot=` (header / totals row counts).

### Robustness sweep (package level)

Files: `tests/xlsx/patch-fuzz.test.ts`, `patch-robust.test.ts`, `patch-round4*.test.ts`; `tests/golden/fuzz.ts`, `mini-scenario.ts`,
`package-tools.ts`; tools `fuzz-cli.ts`, `make-fuzz-outputs.ts`, `make-ops-outputs.ts`.
Contract: every outcome is a correct package or a **PatchError with a Japanese message**; never a TypeError, never a package Excel repairs.

- **Explicit cases** (`patch-robust.test.ts`, 40 tests): namespace-prefixed SpreadsheetML is read through the prefix and written back in the
  default-namespace form (a prefix that cannot be removed safely is refused); cells / columns beyond XFD and rows beyond 1,048,576, duplicate or
  misplaced cells (refused when the sheet has to be written, with the real cause named); a full row of 16,384 styled cells; `<col>` clamps; rows with
  wrong `spans`; missing or wrong `<dimension>`; inline strings (also rich); shared strings with rich runs / phonetics / `xml:space`; books without
  `styles.xml` / `sharedStrings.xml` (the parts, relationships and content types are created when a write needs them); xlsm (vbaProject.bin,
  content type and relationship carried through); formulas of 8,000 and 150,000 characters; a typed formula longer than Excel allows or with control
  characters (refused); Japanese sheet names with quotes / ampersands / cell-reference look-alikes (`A1`, `R1C1`); damaged XML (cut off, raw `&`,
  unbalanced tags, undeclared prefixes), non-xlsx input, hostile op logs; every XML part the patcher writes is checked for well-formedness.
- **New in this round**: attributes in single quotes are normalised on read (before, a shared formula written with `t='shared'` silently lost its
  formula); `<si />` (a .NET writer's empty item) counts as a shared string (it used to shift every later index); `<sst count='…'>` no longer
  produces a duplicate attribute; the encoding declaration of a UTF-16 part is rewritten whatever its quotes; a zip entry that does not inflate to its
  declared size (passed through untouched, the output would have been damaged) is refused; a sheet with cells beyond XFD is refused by the op that has
  to write it with the real cause (セルが Excel の範囲（XFD 列）外) instead of a misleading "data pushed off the sheet".
- **Metamorphic** (`patch-fuzz.test.ts`; `FUZZ_ALL=1` runs all 81 golden cases): the same package spelled with a namespace prefix, single quotes +
  line breaks between elements, a BOM, or UTF-16 replays to the same result as the ordinary spelling; `buildPremerged` on the small hand-written
  scenario (`mini-scenario.ts`) in 8 unusual forms (prefixes, no styles.xml / sharedStrings.xml in base or source, inline strings, UTF-16, spans + no
  dimension + reversed order, a 16,384-cell row, quoted sheet names, xlsm) gives the ordinary result.
- **Seeded fuzz** (`tests/golden/fuzz.ts`): 4 fuzzers — damaged packages + golden op logs (`replay`), damaged base / source of the premerge scenarios
  (`premerge`), random op logs on intact packages (`ops`, half of the positions at the fixtures' tables / totals rows / spills / merges), random
  premerge plans on intact packages followed by random op logs (`plans`: column subsets / orders, matches, 0..3 appended rows, options). Mutations:
  truncate / delete / duplicate / swap / drop elements, junk attribute values and numbers, junk characters, renamed closing tags, deleted or emptied parts,
  BOM / CRLF / single quotes, bit flips / cuts / garbled bytes in the zip. Checked per mutant: Japanese PatchError (not internal) or an output whose written
  XML parts are well-formed, that the reader reads when it could read the mutant, and that replays once more.
  `npx tsx tests/golden/tools/fuzz-cli.ts replay|premerge|ops|plans <seed> <count>`: with the final code **replay 4,000 + premerge 2,000 + ops 6,000 + plans 1,000 mutants: 0 violations, 0 internal errors** (the suite runs ~360, `FUZZ_N=2000` more). During the round the runs found: 21 damaged zip entries that were passed through untouched (now refused) — everything else they found was found by Excel, below.
- **Excel as the oracle** (`make-ops-outputs.ts` + `opentest.ps1`, one hidden Excel instance per batch, own PID): the outputs of random op logs on the
  Excel-authored fixtures and on premerged packages. XML checks cannot see what Excel objects to; Excel found three real defects that all the structural
  checks had passed: the spill range written into (repair), the totals row deleted (Excel stopped responding), a data row moved above a table header
  (repair); all fixed above. After the fixes **900 random-op outputs (batches of 201 / 299 / 400 files on the golden fixtures and the premerged packages) and 469 random-plan outputs (250 premerged packages, 219 of them after random op logs) all open normally, no repair, no hang**; the 8 new golden cases were produced by Excel in the same runs. The earlier hand-picked outputs (`make-fuzz-outputs.ts`: respelled fixtures with their golden
  op logs, the premerge variants, hostile-but-valid books, accepted spelling mutants): **48 / 48** open normally (after replacing an 8,000-character
  reference chain that Excel itself cannot hold by one it can).

## Behaviour probes (facts measured in Excel, now implemented)

1. **Inserted cells** (`probe-insert-format.ps1`, case p-borders): a new cell takes the xf of its left (column
   insert) / upper (row insert) neighbour, but each border side survives only when that neighbour and the cell on
   the other side of the insertion point have the identical side (all-around thin + nothing on the right -> no
   borders; T+T -> T; T+B -> none; LRTB+LTB -> LTB). Fill, font, number format always come from the left/upper cell.
   Inserted rows copy the upper row's height. No left cell (insert at A / row 1) -> no format.
2. **CF / DV growth** (`probe-cf-leak.ps1`): every CF type and data validation grows over a column inserted right
   after its area (format inheritance). On a **move**, the block that lands right after an area picks up cellIs,
   expression, containsText, containsBlanks rules and validations, but NOT colorScale, dataBar, iconSet, top10,
   uniqueValues, aboveAverage. (R2b R9 said DV does not grow; the goldens g-ins-D / g-ins-row14 / t-mov-L-H show
   it does.)
3. **Frozen panes**: a move = insert at the target + delete of the source for the frozen count; a delete that
   removes ALL frozen columns keeps the split (g-plan-ACE-F: one frozen column survives moving column A away).
4. **Drawings** follow Excel's move sequence: insert at the target, only objects lying entirely in the moved block
   travel with it, then the source is deleted (clips objects). `editAs="oneCell"` objects are not resized by
   inserts but are clipped by deletes.
5. **Tables**: rows inserted inside the data body get the calculated-column formulas; inserting inside the table
   adds `列1..` columns with header cells; deleting every column removes the table (part, rel, content type).
6. **Dynamic arrays**: an op that cuts through a spill range clears the spilled values and collapses the master's
   ref to its own cell (Excel re-spills on load); otherwise stale spilled values would become constants.
7. **Column widths**: `<col width>` is the stored value; `Range.ColumnWidth` excludes the 5 px padding
   (0.625 chars with an 11 pt Yu Gothic default font). LoggedOp `colWidth.width` is the stored value.

## Excel opens every output without repair

`com/opentest.ps1` on 50 files produced by `cli.ts excelcheck` (all 36 non-refused golden / probe cases, 6 premerge
variants and each of them after typical edits, the region-narrower-than-table premerge, and the 10,000 x 40
performance output): **50 / 50 open normally**, no repair log.
Round 2: the same on 84 files (all 63 non-refused golden / probe cases incl. every round-2 case, the 18 premerge
outputs (6 variants, each after typical edits and after round-2 edits: header rename, typed date with number format,
dynamic array), `pm-inside-table`, `pm-rich`, `pm-keyfill`): **84 / 84 open normally**, no repair log.
After full recalculation the totals cases show `T!C6 = SUM(C2:C5) = 60` (formula written before a column + row insert
follows them), `T!B5 = SUBTOTAL(104,[Score]) = 30`, `T!C5 = SUBTOTAL(103,[Twice]) = 3`.

Premerge outputs opened in Excel and recalculated (`opentest.ps1 -Cells`): `学生!D9 = SUM(D2:D7) = 350` (the 合計
row includes the appended row), `COUNTA(B2:B7) = 6`, remapped imported formula `I2 = H2*2+F2 = 185.25`,
`Report!A2 = SUM(Scores) = 350`, the line-break formula text, appended key `A7 = S009` — for both the AutoFilter
and the ListObject base.

## Performance (Node 24, this PC)

| scenario | time |
|---|---|
| replayLog, 20 ops (moves, inserts, deletes, row ops, width, cell writes) on a 10,000 x 40 sheet with 50,000 formulas (16 MB sheet XML) | 2.9–3.9 s |
| buildPremerged, 10,000 base rows x 40 imported columns (400,000 cells, styles merged) | ~5 s |

## Re-running

```
# 1) fixtures (only when they must change; commits new bytes)
powershell -NoProfile -ExecutionPolicy Bypass -File tests\golden\com\make-fixtures.ps1 -OutDir tests\fixtures\golden
powershell -NoProfile -ExecutionPolicy Bypass -File tests\golden\com\make-fixtures-r2.ps1 -OutDir tests\fixtures\golden [-Only arrays|rich|table]
# 2) Excel's results of every case
npx tsx tests/golden/tools/cli.ts cases %TEMP%\cases.json
powershell -NoProfile -ExecutionPolicy Bypass -File tests\golden\com\run-ops.ps1 -Cases %TEMP%\cases.json -FixtureDir tests\fixtures\golden -OutDir tests\fixtures\golden\out [-Only "a,b"]
# 3) compare (same as the vitest suite, prints the differences)
npx tsx tests/golden/tools/cli.ts diffall - 8
npx vitest run tests/golden tests/xlsx/patch
# 4) open our outputs in Excel (+ recalculated cells, rich text runs)
npx tsx tests/golden/tools/cli.ts excelcheck <dir>
powershell -NoProfile -ExecutionPolicy Bypass -File tests\golden\com\opentest.ps1 -Dir <dir> [-Filter r-dyn-new.xlsx -Cells "Report!C1,Report!D2"]
powershell -NoProfile -ExecutionPolicy Bypass -File tests\golden\com\probe-runs.ps1 -File <dir>\pm-rich.xlsx -Cells "学生!E2,学生!E3,学生!B7"
# 5) round 3: the hand-written fixture is normalized by Excel, the goldens are produced and our outputs are opened in ONE instance
npx tsx tests/golden/tools/make-r3-fixture.ts %TEMP%\r3-misc-raw.xlsx
npx tsx tests/golden/tools/cli.ts cases %TEMP%\cases.json
npx tsx tests/golden/tools/cli.ts excelcheck <dir>
powershell -NoProfile -ExecutionPolicy Bypass -File tests\golden\com\run-ops.ps1 -Cases %TEMP%\cases.json -FixtureDir tests\fixtures\golden -OutDir tests\fixtures\golden\out -Only "u-band,q-text,..." -NormalizeRaw %TEMP%\r3-misc-raw.xlsx -NormalizedName r3-misc.xlsx -OpenDir <dir> [-MergeProbe -ProbeDir tests\fixtures\golden\probe]
powershell -NoProfile -ExecutionPolicy Bypass -File tests\golden\com\probe-r3.ps1 -Raw ... -FixtureOut tests\fixtures\golden\r3-misc.xlsx -ProbeDir tests\fixtures\golden\probe [-QuoteOnly]   # the probes behind the round 3 facts
# 6) round 4: robustness sweep. The vitest file runs ~300 seeded mutants; FUZZ_N=2000 / FUZZ_ALL=1 (all golden cases respelled) run more
npx vitest run tests/xlsx/patch-fuzz.test.ts
npx tsx tests/golden/tools/fuzz-cli.ts replay|premerge|ops|plans <firstSeed> <count> [--all]      # prints the seeds of internal errors / violations
# Excel as the oracle: write outputs of random op logs / random premerge plans, open them in ONE hidden Excel, read the JSON lines
npx tsx tests/golden/tools/make-ops-outputs.ts %TEMP%\ops-out <count> <firstSeed> [<plans>]
powershell -NoProfile -ExecutionPolicy Bypass -File tests\golden\com\opentest.ps1 -Dir %TEMP%\ops-out        # "open":"normal" for every file
npx tsx tests/golden/tools/make-fuzz-outputs.ts %TEMP%\fuzz-out                                              # the hand-picked hostile-but-valid set
# (a file that makes Excel itself stop responding turns every later line into "failed": put the suspects last and bisect)
```

Scripts are UTF-8 with BOM, print ASCII JSON lines, create their own hidden Excel and kill only its PID. Moves use
insert + `Range.Cut(Destination)` + delete (never `Cut()` + `Insert()`); refusals that Excel's UI would show for
table / merge / array moves are checked by the tests on our side (Excel's clipboard-free sequence does not refuse).

## Round 8 — imported formulas (task F1; full report: `tests/xlsx/formula-import-findings.md`)

`buildPremerged` now keeps the formulas of the 2nd and later files live (`src/core/formula/remap.ts`: row mapping through the join plan, ranges over the
whole table, key columns through `keyPairs`, defined names, Excel Tables, arrays). Verified with real Excel by `tests/oracle/import-values/run.ps1`
(ONE hidden instance, 579 jobs): 480 premerged / column-moved packages (48 corpus files × 5 configurations: results stored or not, default / all
columns, join left / full, 3 row orders; the app's sample base with and without single-sheet output) all open **without repair**; 37,851 live formulas equal
the value of the same formula in the original file (recalculated by Excel) for the same person (10,355 of them ranges widened to the merged table, compared
against the source with the unplaced rows deleted by Excel); 15,341 stored results are what Excel computes now (no stale value); no new `#REF!` / `#NAME?` /
`#VALUE!`; a base column moved across the imported block by `replayLog` keeps all 131,974 cell values at the moved positions. 1,907 comparisons are skipped on purpose
(position / row-count results, spills, renamed header cells: they differ from the source by design) and are counted in the report.
```
powershell -NoProfile -ExecutionPolicy Bypass -File tests\oracle\import-values\run.ps1 [-Regenerate] [-Only f04_]   # corpus: gen_corpus.py + author.ps1
EXCEL_TESTS=1 npx vitest run tests/oracle
```
