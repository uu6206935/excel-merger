# excelRef vs real Excel — documented deviations

`src/core/formula/excelRef.ts` is checked against two independent Excel 365 ground truths (plus targeted probes of
our own for trim ranges, formula forms and the dynamic-array flag, generated with `fixtures/formula/com/*`):

| suite | fixture | result |
|---|---|---|
| R2 oracle (`oracle.test.ts`, `exhaustive.test.ts`, `cf.test.ts`) | `fixtures/formula/excel-results.json`, `exhaustive-results.json`, `cf-results.json` | 1,393 + 4,608 + 61 checks, 20 + 3 documented deviations (sections 1, 2, 3, 12) |
| CF / DV of every rule type (`cfSpread.test.ts`) | `fixtures/formula/cf-spread-excel-results.json` (`com/excel-cf-spread.ps1`) | 93 rules (16 CF types, 3 validation types, R2b replicas, formulas pointing across the op) x 16 column ops, compared per cell; 3 documented quirks (section 11) |
| CF after one column move, small families (`cfProbe.test.ts`) | `fixtures/formula/cf-probe-excel-results.json` (`com/make-cf-probe-spec.mjs`, `excel-cf-probe.ps1`) | 129 scenarios (reference offsets x moved column x destination, absolute / mixed references, rule blocks); 6 documented quirks (section 11) |
| Random CF rules x column ops (`cfRandom.test.ts`) | `fixtures/formula/cf-random-excel-results.json` (`com/make-cf-random-spec.mjs`, `excel-cf-probe.ps1`) | 899 scenarios (single rules / two-rule blocks on one or two areas, relative and absolute-column references, 1 to 3 successive ops): 893 equal Excel per cell, 6 documented quirks (section 11) |
| Column moves around a table (`tableMove.test.ts`) | `fixtures/formula/table-move-excel-results.json` (`com/excel-table-move.ps1`, native Insert Cut Cells) | 9 scenarios + the refusals of the R2b table suite (section 13) |
| R2b ground truth (`r2b.test.ts`) | `fixtures/formula/r2b-fixtures.json` (trimmed copy of `research/r2b/excel-ref-fixtures.json`, built by `build-r2b-fixture.mjs`) | 20,755 checked cases: formula 16,659 + 68 dev., defined names 2,322 + 7, CF 532 + 1, DV 714, AutoFilter 92 + 2, tables (xml) 20, tables (COM: range, header order, calculated column) 20, column order (row-1 markers) 131, column widths 131 + 4, merges 12 + 1, merge-tracking names 39. 1,096 cases skipped (below) |
| Trim ranges (`tro.test.ts`) | `fixtures/formula/forms-tro-excel-results.json`, `tro-ops-excel-results.json` | 16 typed formulas + 10 formulas × 9 ops, file and UI form; 1 display quirk (section 9) |
| Formula forms (`forms.test.ts`) | `fixtures/formula/forms-*-excel-results.json` + R2b pairs | section 8 |
| Dynamic-array flag (`dynamic.test.ts`) | `fixtures/formula/forms-*-excel-results.json` (cell `cm`) | 481 typed formulas, 0 missed, 2 harmless extras (section 8) |

Every mismatch must be listed here. `r2b.test.ts` fails on an unlisted mismatch, on a listed case that no longer
deviates, and on a listed case whose result does not match its category below. The exact ids are in
`tests/formula/r2bDeviations.ts`; they are repeated below.

Skipped R2b cases (no expectation to compare with):
* **1,074 cases of 13 ops that Excel refused** (`m05`, `m17`, `cf_m05`, `cf_m17`: destination inside the moved block;
  `t03`, `t06`, `t07`, `t21`: move into / out of / across the table; `t18`: delete the table header row; `g01`, `g14`:
  move that splits a merged area; `g11`, `g12`: change part of an array formula). The doc refuses the same ops
  (ARCHITECTURE decision 5) and `planMove` never produces a destination inside the block.
* **18 cases of op `g09`**: Excel crashed / no ground truth.
* 4 column-order cases of row ops that insert / delete row 1 (the marker row itself).
* Names that did not exist before the op (`before == null`; none in the current fixture).

## 1. Sheet-edge references during a move (`sheet-edge-move`) — intentional

Excel performs a move ("Insert Cut Cells") as *insert n blank columns at the destination → cut/paste → delete the
emptied source columns*. Step 1 pushes whatever sits in the last n columns (XFD, XFC… or row 1048576) off the sheet,
so after **any** move a reference to those columns becomes `#REF!` (`Data!XFD5` → `Data!#REF!`) and ranges near the
edge change (`XFB5:XFC5` → `XFB5:XFD5`). excelRef keeps such references intact: **a move must never create
`#REF!`** (ARCHITECTURE decision 4). Nothing in a real workbook lives in XFA:XFD, so the difference is theoretical.
Check: our result contains no `#REF!`, Excel's contains `#REF!` or an `XF?` column.

* R2 oracle: 18 cases (`oracle.test.ts` `KNOWN`: moves × refs to `XFD1`, `XFC1`, `C1048576`).
* R2b (50): `m01:o018 m01:o089 m02:o017 m02:o018 m02:o089 m03:o017 m03:o018 m03:o089 m04:o017 m04:o018 m04:o089
  m06:o017 m06:o018 m06:o089 m07:o018 m07:o089 m08:o017 m08:o018 m08:o089 m09:o017 m09:o018 m09:o089 m10:o017
  m10:o018 m10:o089 m11:o018 m11:o089 m12:o018 m12:o089 m13:o018 m13:o089 m14:o017 m14:o018 m14:o088 m14:o089
  m15:o017 m15:o018 m15:o088 m15:o089 m16:o018 m16:o089 jp_m02:o017 jp_m02:o018 jp_m02:o089 jp_m03:o017
  jp_m03:o018 jp_m03:o089 jp_m06:o017 jp_m06:o018 jp_m06:o089` (m07 comes from the clipFree sequence — same behaviour).

Consequence for the Excel本体 engine: the COM runner will produce `#REF!` for refs to the last columns of a sheet
after a move; the JS engine will not. Accepted (no real data there).

## 2. Chart `SERIES` losing a deleted category argument — intentional (R2 oracle, 2 cases)

Deleting the category column makes Excel drop the argument (`=SERIES(,,Data!$B$1:$B$5,1)`); excelRef writes
`Data!#REF!` into it. Deletes that break references are blocked / converted to values before they happen
(`analyzeDelete`), so this never reaches a file.

## 3. CF move artifacts — Excel's result not copied (R2 `cf-results.json`, 2 cases; the third is section 12)

`mv_E_to_B|=A2>0` and `mv_C_to_B|=A2>0`: after a move inside a CF range, Excel's fragments for a rule that refers to
the column LEFT of the range contain `#REF!` (and shifted offsets) caused by its internal insert+cut+delete. excelRef
uses the per-cell model for rules whose cells move (every cell keeps pointing at the cell it pointed at). All other
CF/DV cases of both fixtures match Excel exactly, including (R2b) the "leak" of a CF rule onto a moved block that
lands inside or right after it, the growth of a CF area that ends right before an insertion, and one-formula
(anchor-hosted) adjustment of rules whose cells do not move.

## 4. Legacy implicit intersection that becomes `#REF!` (`legacy-single`) — cosmetic

A legacy (pre-dynamic-array) formula that is a bare range (`=Data!C5:E5` entered with `.Formula`) is stored as
`Data!C5:E5`. When the whole range is deleted Excel writes `_xlfn.SINGLE(Data!#REF!)` (UI `=@Data!#REF!`); excelRef
writes `Data!#REF!`. Both evaluate to `#REF!`, and such deletes are blocked before they happen. Check: Excel's text is
ours wrapped in `_xlfn.SINGLE(...)` / `=@`.

* R2b (7): `d03:o194 d07:o194 x01:o194 x02:o194 x03:o194 jp_d03:o194 jp_x03:o194`.

## 5. Deleting the AutoFilter header row (`autofilter-removed`) — not a formula-engine concern

Deleting row 1 of an AutoFilter range `A1:T20` removes the AutoFilter and its hidden `_xlnm._FilterDatabase` name;
`adjustArea` gives `A1:T19` (the range rule). The patcher / doc must drop the AutoFilter and the name when the op
deletes the filter's header row.

* R2b (4): `x06:name:Data!_FilterDatabase x06:autofilter:Data cf_x06:name:Data!_FilterDatabase cf_x06:autofilter:Data`.

## 6. ListObject-adjacent quirks (`table-quirk`) — Excel quirks, not copied

Table `tblMain` on `Data!C3:H10`. excelRef applies the plain-range rules; Excel does something else:

* **Moving columns inside the table** (`t01` E→D, `t02` D→G, `t19` D:E→H, `t22` G→C): the table columns are reordered
  correctly (and `adjustTableRef` keeps `C3:H10`, verified), but Excel also **widens ranges that span the table's
  columns outside the table rows by the block width** (`SUM(Data!C11:H11)` → `C11:I11`, name `=Data!$C:$H` →
  `$C:$I`) and a formula cell below the table moved from `K20` to `L20` — i.e. Excel's internal "insert blank" step
  is not undone outside the table rows. excelRef does a clean whole-column move (row-1 markers confirm the columns
  themselves moved normally).
  The same leftover shows in the **column widths** (`t01 t02 t19 t22`): a default-width (8.38) column appears right
  after the moved block and every width to its right is shifted by the block width, although the row-1 markers and
  the table's own columns moved cleanly. excelRef / the patcher move widths with their columns (`mapIndex`).
  Ids: `t01:t028 t01:tD02 t01:name:nm_tcols t02:t028 t02:tD02 t02:name:nm_tcols t19:t028 t19:tD02
  t19:name:nm_tcols t22:t028 t22:tD02 t22:name:nm_tcols t01:colwidths t02:colwidths t19:colwidths t22:colwidths`.
  → The Excel本体 engine can therefore differ from the JS engine for refs next to a table when the user reorders
  columns inside a table.
* **Inserting a row at the first data row of the table** (`t25`, insert before row 4): references to the table body
  (`SUM(Data!E4:E10)`, `COUNTA(Data!C4:H10)`, name `=Data!$E$4:$E$10`) and the CF on the body **grow** (`E4:E11`)
  instead of shifting (`E5:E11`): Excel treats it as "insert a table row". Row ops are not offered by the UI.
  Ids: `t25:t017 t25:t018 t25:tD03 t25:name:nm_tA1 t25:cf:Data:FFFFFF00`.

## 7. Column moved into the middle of a merged area (`merge-unmerged`)

`g03` (B before D with merge `C3:E3`): Excel silently **un-merges** `C3:E3`; the range rule gives `B3:E3`. The doc must
refuse a move whose destination lies strictly inside a merged area (same message as splitting a merge), so the
difference never reaches a file.

* R2b (1): `g03:merge:Data`.

## 8. Formula forms (`forms.ts`, `forms.test.ts`) — harmless

Checked on 374 formulas typed into Excel through `Range.Formula2` and read back from the saved XML
(`fixtures/formula/forms-excel-results.json`: every candidate future function + LET / LAMBDA / spill / `@` / tables;
`fixtures/formula/forms-extra-excel-results.json`: unknown functions, eta-reduced lambdas `_xleta.`, names with `#` /
`@`, table specials with a totals row, case normalisation; `fixtures/formula/forms-new-excel-results.json`: 2024-2026
functions incl. IMPORTTEXT / IMPORTCSV). Generators: `fixtures/formula/com/*`. Also checked on the
34,900 (Formula2, saved `<f>`) pairs of the R2b fixture (every formula before and after every op): all equal in both
directions except 99 legacy-`@` pairs (first bullet below; e.g. `=@Data!C5:E5` stored as `Data!C5:E5`).

* **Legacy implicit intersection `@`** (`S!A12 S!A13 S!A27 S!A28 X!A5 X!A11 X!A16`): Excel drops a typed `@` where
  the legacy (non-dynamic) evaluation intersects anyway — `=@Z1:Z3` is stored as `Z1:Z3`, `=IF(@B1:B3>1,1,0)` as
  `IF(B1:B3>1,1,0)`, `=@nmRange` as `nmRange`, `=-@B1:B3` as `-B1:B3`, `=ROWS(Z1#)+@Z1#` as
  `ROWS(_xlfn.ANCHORARRAY(Z1))+_xlfn.ANCHORARRAY(Z1)` — and shows it again with `@`. It keeps `_xlfn.SINGLE` where
  legacy would not intersect (`@B1`, `@TODAY()`, `@INDEX(B1:B3,2)`, `SUM(@x)`). The rule depends on every function's
  parameter classes (value vs. reference) and on whether the rest of the formula is dynamic, which forms.ts does not
  model. `toFileForm` writes `_xlfn.SINGLE(Z1:Z3)` (what Excel itself writes
  inside other functions, e.g. `SUM(_xlfn.SINGLE(Z1:Z3))`), which Excel loads as the same `=@Z1:Z3`. Conversely
  `toUiForm('Z1:Z3')` gives `=Z1:Z3` where Excel displays `=@Z1:Z3`: writing that back through `Range.Formula2` would
  make it a spilling formula, so the COM runner should write non-dynamic cells with `Range.Formula`.
* **Functions this Excel build does not know** (COPILOT, QUERYSTRING, LONGTEXT, PYTHON_STR/TYPE/TYPENAME): the build
  stores them unprefixed (as unknown names); `toFileForm` follows MS-XLSX 2.2.2 and writes `_xlfn.`. `toUiForm`
  strips any `_xlfn.` / `_xlws.` / `_xlpm.` prefix, so both spellings read back fine.
* Structured references: an unqualified `[@col]` / `[col]` needs `FormOptions.hostTable` (the file always names the
  table), a bare table name `Tbl` needs `FormOptions.tables` to become `Tbl[]`.
* **Case of table COLUMN names** (`X!A49`): Excel writes `tblsales[A]` as `TblSales[a]` (header case). `toFileForm`
  fixes the table name (`FormOptions.tables`) and the `[#This Row]`-style specials but keeps the typed column-name
  case; Excel matches column names case-insensitively and re-renders them on load. Likewise defined names and sheet
  prefixes get their real case only when listed in `FormOptions.names` / `FormOptions.sheets`.
* **Eta-reduced lambdas** (`_xleta.`; `X!A41 X!A50..A56 X!A64..A69` all match Excel): recognised in the LAMBDA
  positions of BYROW / BYCOL / MAP / REDUCE / SCAN / MAKEARRAY / GROUPBY / PIVOTBY (also inside a HSTACK / VSTACK
  there) and as a LET value that is later called (`LET(f,SUM,f(x))`). A bare function name elsewhere stays a plain
  name, so a workbook name such as `PRICE` never becomes `_xleta.PRICE`.
* **Dynamic-array flag** (not a forms.ts output): Excel marks some formulas typed through `Formula2` as dynamic
  (`cm="1"`, `t="array"`) — `IFS(...)`, `SUMPRODUCT((B1:B3>1)*B1:B3)`, `FOOBAR(1)`, `X!B1:B3`, `{1,2,3}`, spilling
  functions — and others not (`XLOOKUP(...)`, `LET(x,1,x)`, `SUM(Z1#)`); the flags are in the fixtures (`file.cm`).
  `toFileForm` only produces the `<f>` text; whoever writes a typed formula into a file decides `cm` —
  `dynamic.ts isDynamicArrayFormula` predicts it conservatively (`dynamic.test.ts`: 481 typed formulas incl.
  `fixtures/formula/forms-dyn-excel-results.json`, no dynamic formula missed; 2 harmless extras `INDEX(Z1#,2)`,
  `INDEX(B1:C4,2)` that Excel does not flag).
* **Writing back through COM**: `toUiForm(<f>)` equals `Range.Formula` for every non-array cell and `Range.Formula2`
  for every array / dynamic cell (411 forms-fixture cells + 34,911 R2b formulas checked). So a legacy cell must be
  written with `Range.Formula` (writing `=Data!C5:E5` or `=IF(B1:B3>1,1,0)` through `Formula2` turns it into a
  spilling formula with another result); a dynamic one (`cm`) and a formula the user typed in the app with
  `Range.Formula2`.

## 9. Trim ranges that became `#REF!` without a sheet prefix (`tro.test.ts`) — Excel display quirk

Trim refs (`B1:.B10`, `B1.:B10`, `B1.:.B10`; file form `_xlfn._TRO_TRAILING(B1:B10)` / `_TRO_LEADING` / `_TRO_ALL`)
are checked on 16 typed formulas and on 10 formulas × 9 ops (`fixtures/formula/forms-tro-excel-results.json`,
`tro-ops-excel-results.json`): file form, UI form and both conversions equal Excel, including the `#REF!` marker
(`Data!#REF!:.` ⇄ `_xlfn._TRO_TRAILING(Data!#REF!)`, `.:` leading, `.:.` all). One exception: when an UNPREFIXED
trimmed range on the formula's own sheet is deleted, Excel stores `SUM(_xlfn._TRO_TRAILING(#REF!))` but displays
`=#REF! SUM()`; toUiForm gives `=SUM(#REF!:.)` (the exact inverse of the stored text). Cases `delBF Data!K11 K12`,
`delRows1to10 Data!K11`. Deletes that create `#REF!` are converted to values before they happen anyway.

## 10. Conditional formats / validations: growth and spread (round 2) — implemented to match Excel

Measured on rules of every type with the clipboard-free move sequence (the one the Excel runner executes and
`replayLog` implements): `cfSpread.test.ts`, `cfProbe.test.ts`, the patcher goldens.

* **Insert right after an area**: every CF type and every data validation grows over the inserted cells (format
  inheritance). Data validations behave exactly like conditional formats. (R2b said "validations do not grow": its
  validation areas were read through defined names that tracked them — plain range semantics — so it could not see the
  growth; `r2b.test.ts` compares them with `inheritFormat: false` for that reason.)
* **Move**: the moved block lands at X; when X-1 is in a rule it joins the rule ("spread") for cellIs, expression,
  containsText, containsBlanks, containsErrors, beginsWith, timePeriod ... and every validation type, but NOT for
  colorScale, dataBar, iconSet, top10, uniqueValues, duplicateValues, aboveAverage (`NO_SPREAD_CF_TYPES`).
* The spread is cancelled when a rule of the `<conditionalFormatting>` block has, in the range object that covers
  the destination, a formula that refers to the moved cells or to the destination, or when that range object contains
  the moved block. Excel keeps ONE range object per block: pass the other rules of the block as `siblings`.
* The part of a rule right of a moved column is one piece (the formula of its first cell); a relative reference to the
  moved column from that part dies in Excel's delete step (`#REF!`), an absolute one follows the column. The cells
  left of the moved column and the moved cells keep their own formulas.

## 11. CF quirks of Excel that are not copied (`cfSpread.test.ts`, `cfProbe.test.ts` KNOWN)

* **Absolute reference to the moved column** (`=$G209>0` on C:E, G moved before D; `=$D2>0` on C:G, D moved to H):
  Excel takes the cells left of the destination out of the rule and lets it cover the columns between destination and
  source (the CF_INSIDE_MOVE family of tests/golden/RESULTS.md). We keep "every cell keeps its rule". Ids:
  `movGtoD|cf:expression|C209:E210`, `movGtoD|cf:cellIs|C260:E261`, `p3_absD_sD_xH`, `p3_absE_sE_xH`.
* **Destination piece of formulas that mix a relative reference to the moved column with an absolute one before /
  between source and destination** (`=D2>$F2`, `=$B2>D2` on C:G, D moved to H): Excel takes the destination cell out
  of the rule again, we keep it (one more cell is formatted). Ids: `movDtoH|cf:expression|C296:G297`, `p4_07`, `p4_09`,
  `p4_16`, `p4_26`; the random run has five more: `p7_042 p7_071 p7_197 p8_078 p8_276` (two-reference formulas with an
  absolute column, most of them in two-rule blocks; 6 of 899 random scenarios). Formulas with one reference, only
  relative references or an absolute reference right of the destination match.
* **A moved piece that is a whole area of a two-area rule** whose cell refers to another cell of the piece (after three
  ops, `p8_168`): Excel writes `#REF!` for that cell, we keep the reference following the piece.

## 12. CF rule block of eight rules: per-column range objects (`cf.test.ts`)

`mv_B_to_F|=A2>0` (R2): in the block of 8 rules on B2:E10 Excel fragments the range into one object PER COLUMN, so
every cell right of the moved column keeps its own formula (`D` -> `C2>0`) instead of sharing the first cell's
(`E2>0`-relative). The block's fragmentation depends on all its rules; `siblings` reproduces the small blocks
(cfProbe p5, 2 rules) but not this one.

## 13. Column moves around a table (native Insert Cut Cells; `tableMove.test.ts`)

Table on C3:H10, verified in Excel 365:

* a block from INSIDE the table landing anywhere in [c1, c2 + 1] is accepted and only reorders the table columns, the
  range stays: destination `c2 + 1` (right after the last column) included (`in_to_I`, `in2_to_I`, `first_to_I`);
  further away (R2b t03) Excel refuses. `tableMoveKind` / `adjustTableRef` implement exactly this. **The doc
  (`doc/struct.ts` moveRefusal: `to > b.c2`) and the patcher (`patch/engine.ts`: `op.to <= hi`) still refuse / treat
  `to = c2 + 1` as leaving the table** — they should use `tableMoveKind`.
* a block covering the table moves it (`whole_to_K`, `cover_to_L`); a block from outside landing right after the
  table (`out_to_I`) or at its first column (`outL_to_C`) does not extend / enter it; landing strictly inside
  it, partly over it: refused (R2b t06, t07, t21).

## Not deviations (implemented to match Excel)

* Column permutation (`mapIndex` / `movePerm`): the row-1 markers of every column op (moves incl. the clip-free
  sequence, inserts, deletes) and the column widths match Excel. An inserted column takes the width of the column on
  its left ("format same as left"; default width when inserted at column A) — the patcher's rule for user inserts
  (the premerge insert is a clean insert by design).
* ListObjects (COM view): range, header order after moves / deletes (a deleted column disappears, an inserted one gets
  a new name `列1`), and the calculated-column formula `=[@点数]*2` (→ `=#REF!*2` when 点数 is deleted) match Excel.

* Defined names: relative parts are offsets and are **not** adjusted (`AdjustContext.isName`, `adjustNameArea`).
* Structured references to a deleted table column / deleted table become `#REF!` (`AdjustContext.tables`); moves and
  inserts never change them. UI-form `[@col]` is resolved through `AdjustContext.hostCell`.
* A range spanning the whole axis (`1:1048576`, `A5:XFD5`) never changes; a range reaching the last row/column stays
  anchored there on delete; a single cell pushed past XFD becomes `#REF!` on insert.
* CF / DV sqref: inserts grow an area that ends right before the insertion point; moves spread a rule onto a block
  landing inside / right after one of its areas (validations too; `AnchoredRuleOptions.type`, section 10).
