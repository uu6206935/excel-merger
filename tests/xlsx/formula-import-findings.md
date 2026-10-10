# Imported formulas — findings and verification (round 8, task F1)

User report (translated): *"For the 1st Excel file the formulas are read in properly, but from the 2nd file on the formulas are not loaded and become
plain/blank cells. Also, when columns are swapped, or rows are rearranged because rows are matched by name, the formulas must stay correct."*

Everything below is reproducible: `tests/oracle/import-values/` (corpus generator, diagnosis tool, real-Excel oracle) and the fixtures in
`tests/fixtures/formulaimport/`. Numbers are formula cells of the imported columns (matched + appended rows), summed over the three row orders.

## 1. Why the 1st file works and the 2nd does not

* The 1st file is the **base**: its package is the container of the merged workbook. Its cells, formulas, names, tables … stay in place and the
  patcher only *adjusts* them (insert columns / rows). Nothing is re-created, so every formula survives.
* Every other file is **imported cell by cell** into the base sheet (`buildPremerged`). A formula of that file reads cells of *its own* sheet at
  *its own* positions; in the merged sheet the same data sits in other columns and — because rows are matched by key (ID / 氏名), not by
  position — in other rows. The old rule kept a formula live only when **every** reference was a same-row reference to an imported column;
  anything else was replaced by the cached value stored in the source file.
* A source file written by a program that does not calculate (openpyxl, pandas, ClosedXML, POI, many exporters) stores **no result** next to its
  formulas. A formula that cannot be kept live and has no stored result leaves **nothing to import: an empty cell**. That is exactly the report
  ("function not loaded, just a blank cell"). The same formula in a file saved by Excel (results stored) was "only" turned into a value.

## 2. Diagnosis with evidence (old behaviour = frozen copy `tests/oracle/import-values/legacy/premerge-legacy.ts`)

Corpus (`gen_corpus.py` + `author.ps1`): 15 source workbooks × 3 row orders (same / reversed / shuffled — 27 of 31 people match the base,
4 only exist in the source, 3 base people are missing in the source) = 45 files, plus one more workbook × 3 orders built for the app's own
sample base `samples/01_学生名簿.xlsx` (merged title, header row 3, names, charts, pivot, footer; 57 of its 60 students + 3 strangers) = 48 files, each in two variants: **no results** (openpyxl / XlsxWriter,
as written) and **cached** (the same workbook opened, `CalculateFull`-ed and saved by real Excel; shared formulas created by `Range.Formula`
on a range). Pipeline = the app's own (`analyzeSources` → `defaultPlan` → `joinSources` → `buildPremerged`), default imports (key columns
excluded) and all-columns imports. `diagnose.ts` classifies every formula cell of the imported columns of the *output* as
live (still a formula) / value (cached result written) / EMPTY (nothing written).

### 2.1 Before (old premerge) — live / value / EMPTY

| item (formulas in the source) | cached, default imports | cached, all columns | no results, default imports | no results, all columns |
|---|---|---|---|---|
| f01 same-row arithmetic / IF / AVERAGE / TEXT / dates / "" / #N/A | 648 / 0 / 0 | 648 / 0 / 0 | 648 / 0 / 0 | 648 / 0 / 0 |
| f02 shared formulas (Excel `FillDown`) and dependents | 243 / 0 / 0 | 243 / 0 / 0 | — (no formulas without Excel) | — |
| f03 formulas reading the key columns (`=A2&"様"`, `COUNTIF($A$2:$A$32,A2)`, `VLOOKUP(A2,$A$2:$E$32,…)`) | 0 / 486 / 81 | 324 / 162 / 81 | 0 / 0 / 567 | 324 / 0 / 243 |
| f04 lookups (INDEX-MATCH, SUMIF, COUNTIF, RANK, LARGE, XLOOKUP, SUMPRODUCT, MATCH, VLOOKUP) | 0 / 810 / 0 | 0 / 810 / 0 | 0 / 0 / 810 | 0 / 0 / 810 |
| f05 ranges (`C$2:C$32`, `D2:D32`, `D1:D32`, `E:E`, `D2:D1000`, 1 row short) | 2 / 808 / 0 | 2 / 808 / 0 | 2 / 0 / 808 | 2 / 0 / 808 |
| f06 running total, previous-row difference, share of total, rate in a side cell, `ROW()`, `ROWS`, running COUNTA | 4 / 725 / 0 | 6 / 723 / 0 | 4 / 0 / 725 | 6 / 0 / 723 |
| f07 other sheet of the source (`Master!…`, `Sheet2!…`), self-qualified `'評価'!C2` | 81 / 405 / 0 | 81 / 405 / 0 | 81 / 0 / 405 | 81 / 0 / 405 |
| f08 defined names (constant, range, side cell, other sheet, OFFSET, text, sheet scope) | 0 / 729 / 0 | 0 / 729 / 0 | 0 / 0 / 729 | 0 / 0 / 729 |
| f09 Excel Table + structured references | 0 / 729 / 0 | 0 / 729 / 0 | 0 / 0 / 729 | 0 / 0 / 729 |
| f11 text / date / time / boolean / #DIV/0! / #N/A / "" results | 567 / 0 / 0 | 567 / 0 / 0 | 567 / 0 / 0 | 567 / 0 / 0 |
| f12 footer totals row below the table, formulas reading it | 81 / 243 / 0 | 81 / 243 / 0 | 81 / 0 / 243 | 81 / 0 / 243 |
| f13 HYPERLINK / CONCAT | 0 / 243 / 0 | 162 / 81 / 0 | 0 / 0 / 243 | 162 / 0 / 81 |
| f16 external workbook, 3D, INDIRECT, OFFSET, ADDRESS, full-width `＋` | 162 / 405 / 0 | 162 / 405 / 0 | 81 / 0 / 486 | 81 / 0 / 486 |
| f14 two-row header | 81 / 162 / 0 | 81 / 162 / 0 | 81 / 0 / 162 | 81 / 0 / 162 |
| f10 CSE / dynamic arrays / spill (XlsxWriter: stores a placeholder 0) | 0 / 166 / 0 | 0 / 166 / 0 | 0 / 166 / 0 | 0 / 166 / 0 |
| g01 the app's sample base: same-row sum, `B2&"様"`, deviation / rank / share over the table, previous row, running total, INDEX-MATCH by ID, name `税率`, COUNTIF duplicate check | 175 / 1,364 / 171 | 346 / 1,193 / 171 | 175 / 0 / 1,535 | 346 / 0 / 1,364 |
| **total** (9,571 formula cells in the cached columns, 9,328 in the no-results columns: f02 has no formulas without Excel) | **2,044 / 7,275 / 252** | **2,703 / 6,616 / 252** | **1,720 / 166 / 7,442** | **2,379 / 166 / 6,783** |

(The 252 EMPTY cells in the cached columns are a second, smaller defect found on the way: a formula whose result is the empty string `""` (the 重複 check `IF(COUNTIF(…)>1,"重複","")`)
is stored by Excel as `<c t="str"><f>…</f><v/></c>`; the old reader looked for `<v>…</v>`, did not find it and treated the result as *missing*. Fixed in `premerge.ts`
(`<v/>` with `t="str"` is an empty text); the frozen copy keeps the defect, so the "before" columns still show it.
With join "full" and the shuffled order, default imports, cached: 778 live / 2,740 value / 91 EMPTY of 3,609.)

Reasons of the old rule (the first token that failed), "no results / default imports" column — every one of these cells is EMPTY except the arrays:

| reason (old rule) | cells |
|---|---|
| reference to another row, or a range (even the whole table `SUM(C$2:C$32)`) | 3,199 |
| column that is not imported (the key columns) | 983 |
| defined name | 819 |
| structured reference `Table[…]` | 729 |
| cell outside the table (side cell, footer, title) | 569 |
| `ROW` / `COLUMN` / `OFFSET` / `INDIRECT` / `CELL` / `SHEET` | 405 |
| whole-column reference (`E:E`) | 333 |
| array / dynamic array / spill (a value: XlsxWriter's placeholder) | 166 |
| other sheet | 162 |
| external workbook / 3D / unsafe / unknown token | 81 + 81 + 81 |

### 2.2 Hypotheses

| hypothesis | verdict | evidence |
|---|---|---|
| no cached value + cannot-remap → EMPTY cell | **TRUE — the cause of the report** | the same formulas (9,571 with results / 9,328 without): 252 EMPTY with results stored (the `<v/>` defect above), **7,442 EMPTY without** (`diagnose.ts`, table above); 112 of 465 formulas EMPTY in the repository's foreign-producer corpus (§2.3) |
| array flag (`t="array"`, `cm`) | true, never live; value when a result is stored, EMPTY otherwise | f10: 166 cells per column; XlsxWriter stores a placeholder 0, so they were never EMPTY there (see §6.4) |
| `isUnsafeFormula` (external workbook, DDE, WEBSERVICE …) | true, by design (still a value after the fix) | f16 `'C:\data\[other.xlsx]Sheet1'!A1` |
| `rewriteRefs` returns null for names / structured refs / positional functions | **TRUE** | 819 + 729 + 405 cells ("no results / default imports") |
| refs to non-imported columns | **TRUE** (default imports exclude exactly the key columns formulas like to read) | f03: `=A2&"様"`, `COUNTIF($A$2:$A$32,A2)`, `VLOOKUP(A2,$A$2:$E$32,3,FALSE)` 0 live of 567 (f03, default imports) |
| ranges | **TRUE** (any range, even `SUM(C$2:C$32)`, was a value) | f05 / f04: 2 live of 810 |
| row relations after reordering | true, but the old rule never tried: *any* reference to another row was a value, in the same order as well | f06 `=C3-C2`, `SUM(C$2:C3)` in "same" order: value |
| shared-formula dependents become `arr` / empty | **FALSE** | the patcher unshares on open (`SheetModel.unshare`): f02 cached 243 / 243 live; a dependent whose master is missing becomes a plain value (unit test) |
| `.xlsm` source | no effect | `f01_samerow__same.xlsm`: same 216 live |
| formulas on rows that match nothing | not imported at all (only placed rows are written; join "full" appends them) | corpus: 4 source-only people |

### 2.3 Foreign producers (repository corpus `tests/fixtures/foreign/corpus`, the OTHER file of `partner.xlsx`; only files with formulas)

| file | formulas | before: live / value / EMPTY | after: live / value / EMPTY | after: not kept, why |
|---|---|---|---|---|
| hc-05-mac-excel-2011-1904 | 28 | 28 / 0 / 0 | 28 / 0 / 0 | |
| hc-28-table-awkward-names | 8 | 0 / 8 / 0 | 8 / 0 / 0 | |
| hc-34-formulas-without-cache | 84 | 0 / 0 / **84** | 56 / 0 / 28 | otherSheet: 28 (`INDEX(マスタ!…, MOD(ROW(),8)+1)`) |
| js-04-exceljs-basic | 28 | 28 / 0 / 0 | 28 / 0 / 0 | |
| py-01 … py-02, py-05 … py-06, py-10 … py-12, py-27 (openpyxl) | 8 × 28 | 28 / 0 / 0 each | 28 / 0 / 0 each | |
| py-04-openpyxl-table | 28 | 0 / 0 / **28** | 28 / 0 / 0 | |
| py-13-openpyxl-blank-rows-stray | 8 | 8 / 0 / 0 | 8 / 0 / 0 | |
| py-14-openpyxl-formats-errors | 1 | 1 / 0 / 0 | 1 / 0 / 0 | |
| py-15-xlsxwriter-basic | 28 | 28 / 0 / 0 | 28 / 0 / 0 | |
| py-19-xlsxwriter-table | 28 | 0 / 28 (a placeholder 0) / 0 | 28 / 0 / 0 | |
| **total** | **465** | **317 / 36 / 112** | **437 / 0 / 28** | |

## 3. What was changed

`src/core/formula/remap.ts` (new, token based on `excelRef.tokenize` / `parsePrefix` / `renderArea`; `$` flags and everything that is not a
reference are copied verbatim) is used by `premerge.ts` for every imported formula:

* **a. Row mapping T(r)**: a reference to one source row is rewritten to the merged row of that source row (matched base row or appended
  row); the formula's own row is T(sr) = target; the header row maps to the base header row; a row without T → `rowRelation` /
  `externalCell`. Source row placed twice (duplicate key in the base): only the own row is known.
* **b. Whole-table ranges** (`C2:C32`, `C1:C32`, `C:C`, `C2:C1000` … covering first…last data row) become the whole merged data table
  (`first data row … last data row + appended rows`, whole columns stay whole columns). Why: the formula now lives in the merged table and is
  recalculated there (Excel, `fullCalcOnLoad`); rows of the source that matched nothing are not in it, appended rows are. Partial ranges only
  when T(a)…T(b) are consecutive and ascending. Reported as info `formulaRangesWidened` with the count.
* **c. Columns**: imported columns by `colMap` (multi-column areas need consecutive merged columns in the same order, so a `VLOOKUP` column
  index stays right). A non-imported source column that is the source side of a `keyPairs` pair reads the base partner column when
  **every** matched row (and every appended row, by what `keyFill` writes) has the same type, value and displayed text on both sides (a
  column-level decision); otherwise `notImported`.
* **d. Defined names** (workbook scope or the data sheet's scope; `SourceBook.names`, or `workbook.xml` for a bare book) whose `refersTo` is
  one reference to the data sheet, a structured reference, or a constant are inlined; everything else `unresolvedName` (`otherSheet` /
  `externalCell` when the reference shows that). **Table structured references** (`[#This Row]`, `[col]`, `[#Data]`, `[#Headers]`, column
  ranges, the whole table inside a function, XlsxWriter's table-less `[[#This Row],col]`) are resolved with the source's table info; totals-row
  references → `unresolvedTable`.
* **e. Functions**: `ROW COLUMN OFFSET INDIRECT ADDRESS CELL FORMULATEXT SHEET SHEETS INFO ISFORMULA GETPIVOTDATA` → `positional` (value);
  `ROWS`/`COLUMNS` with an explicit reference are fine; spill references (`F2#`, `ANCHORARRAY`) → `array`; `_xludf.` / `[1]!Fn(` / named LAMBDA → `unsupported` / `unresolvedName`.
* **f. Shared formulas / arrays**: dependents are expanded by the patcher before the import; a single-cell CSE array stays `t="array" ref=<cell>`,
  a single-cell dynamic array stays `cm` + metadata; multi-cell arrays / spills become values (`array`).
* **g. Results**: the stored result stays next to a kept formula only when every reference reads the very same values (`exact`); otherwise the
  formula is written **without** a result and Excel computes it on load (`fullCalcOnLoad` is always set; checked on every package). A formula
  that cannot be kept and has no result is EMPTY and gets the warning `formulasNoCache` (「元のファイルに計算結果が保存されていない数式です（Excelで開いて保存し直してください）」).
  Results that are only placeholders — XlsxWriter writes `0` for every formula and sets `fullCalcOnLoad` — count as "no result" (warning `formulasPlaceholderCache`):
  the oracle caught stale `0`s next to formulas Excel computes as 78 / 212.
* **h. Report**: `MergeReport.formulaStats` per source `{kept, toValues, empty, reasons}`; items `formulasKept` (per source), `formulasToValues`
  (one per reason, plain words: 「他のシートを参照している数式 N 件は値にしました」「表の外のセル（例: 税率）を参照している数式 N 件は値にしました」…),
  `formulasNoCache`, `formulaRangesWidened`, `formulasRecalc` (kept without a result: the grid shows them empty until Excel opens the file), `formulasPlaceholderCache`.
* `buildPremerged(input, onProgress?, trace?)`: the optional third argument receives one `FormulaTrace` per imported formula (tools / tests).
* **k. Unchanged**: appended rows, imported styles, table / filter extension, header conflicts, and `formulaPolicy: 'values'` (every formula a value; only placeholder results — see g — are no longer copied).

## 4. After — live / value / EMPTY

| item | cached, default imports | cached, all columns | no results, default imports | no results, all columns |
|---|---|---|---|---|
| f01 same-row | 648 / 0 / 0 | 648 / 0 / 0 | 648 / 0 / 0 | 648 / 0 / 0 |
| f02 shared | 243 / 0 / 0 | 243 / 0 / 0 | — | — |
| f03 key columns | 486 / 81 / 0 | 567 / 0 / 0 | 486 / 0 / 81 | 567 / 0 / 0 |
| f04 lookups | **810 / 0 / 0** | 810 / 0 / 0 | **810 / 0 / 0** | 810 / 0 / 0 |
| f05 ranges | 621 / 189 / 0 | 621 / 189 / 0 | 621 / 0 / 189 | 621 / 0 / 189 |
| f06 running / previous row / share / side cell / ROW | 453 / 276 / 0 | 453 / 276 / 0 | 453 / 0 / 276 | 453 / 0 / 276 |
| f07 other sheet | 162 / 324 / 0 | 162 / 324 / 0 | 162 / 0 / 324 | 162 / 0 / 324 |
| f08 defined names | 486 / 243 / 0 | 486 / 243 / 0 | 486 / 0 / 243 | 486 / 0 / 243 |
| f09 tables | 648 / 81 / 0 | 729 / 0 / 0 | 648 / 0 / 81 | 729 / 0 / 0 |
| f11 values | 567 / 0 / 0 | 567 / 0 / 0 | 567 / 0 / 0 | 567 / 0 / 0 |
| f12 footer | 162 / 162 / 0 | 162 / 162 / 0 | 162 / 0 / 162 | 162 / 0 / 162 |
| f13 HYPERLINK | 162 / 81 / 0 | 162 / 81 / 0 | 162 / 0 / 81 | 162 / 0 / 81 |
| f16 unsafe | 162 / 405 / 0 | 162 / 405 / 0 | 81 / 0 / 486 | 81 / 0 / 486 |
| f14 two-row header | 243 / 0 / 0 | 243 / 0 / 0 | 243 / 0 / 0 | 243 / 0 / 0 |
| f10 arrays | 162 / 4 / 0 | 162 / 4 / 0 | 166 / 0 / 0 | 166 / 0 / 0 |
| g01 the app's sample base | 1,595 / 115 / 0 | 1,595 / 115 / 0 | 1,595 / 0 / 115 | 1,595 / 0 / 115 |
| **total** | **7,610 / 1,961 / 0** | **7,772 / 1,799 / 0** | **7,290 / 0 / 2,038** | **7,452 / 0 / 1,876** |

**EMPTY cells without stored results: 7,442 → 2,038; live formulas: 1,720 → 7,290** (with stored results: 2,044 → 7,610 live, 252 → 0 EMPTY; join "full", shuffled,
default imports, cached: 778 → 2,839 live of 3,609). What is still not live, and why ("no results / default imports" column, 2,038 cells): other sheet 486
(a lookup into a master sheet of the source: that sheet is not part of the merge) · `ROW()` / `OFFSET` / `INDIRECT` 405 · cells outside the table (a rate in a side
cell, a footer total) 326 · partial ranges whose rows are no longer consecutive after the reorder (a running total, `C3-C2` after the rows were put in key order)
416 · a non-imported column that cannot be substituted 162 · an `OFFSET` name 81 · external workbook 81 · a full-width `＋` 81 (+ 4 multi-cell arrays with
results). Importing "all columns" (key columns included) brings the key-column formulas back (f03, f09). On the app's own sample base (g01) 1,595 of 1,710 formulas
stay live (the 115 others are the running totals of the two reordered files).

## 5. Real Excel — value-invariance oracle (`tests/oracle/import-values/run.ps1`)

One hidden Excel of our own opens **every premerged package** (a repair prompt would be a failure), `CalculateFull`s it and dumps the merged
sheet. Configurations per corpus file (48 files): *cached / default imports / left join*, *cached / all columns / left*, *no results / default /
left*, *no results / default / **full** join (appended rows)*, *cached / all columns / full*; for the app's sample base (g01: merged title, charts, pivot,
names, footer below the table) *cached / left*, *no results / full* and *cached / all columns*, each with and without single-sheet output
(`dropOtherSheets`). For every formula that stayed **live**:

* the value must equal the value of the same formula in the original source (the Excel-recalculated twin) **for the same person**; ranges
  widened to the merged table are compared with the source after Excel *deleted the rows that were not placed* (Excel shrinks the ranges the
  same way) — position / row-count results (`MATCH` on its own, `ROWS`, `COUNTBLANK`, `COUNTA` over a widened key column), spills (`UNIQUE`,
  `FILTER`, `SORT`) in a changed layout, and formulas that read a header cell whose text changed are skipped and counted;
* the **result stored next to the formula** in the package must be what Excel computes now (no stale result);
* no new `#REF!` / `#NAME?` / `#VALUE!`;
* **column move**: the premerged package gets the app's own `replayLog` move of the base table's last column behind the last imported column (the engine the user's
  swaps run through); the moved package is opened and recalculated by Excel and every cell must hold the same value at the moved position.

Result of the final run (48 files × 5 configurations = 240 premerged packages + 240 column-moved twins = **480 packages, all opened without
repair**; 99 "rows deleted" source copies for the widened ranges; 579 Excel jobs in one hidden instance, a failed job is retried twice):

| item | packages (opened without repair) | live formulas | compared with the source | of these: widened ranges | skipped | **different** | new errors | stored results checked / stale | cells after a column move / different |
|---|---|---|---|---|---|---|---|---|---|
| g01 the app's sample base | 30 (30) | 8,153 | 8,153 | 2,571 | 0 | **0** | 0 | 2,214 / 0 | 25,596 / 0 |
| f01 same-row | 30 (30) | 3,432 | 3,432 | 0 | 0 | **0** | 0 | 2,040 / 0 | 7,731 / 0 |
| f02 shared formulas | 30 (30) | 765 | 765 | 0 | 0 | **0** | 0 | 765 / 0 | 5,286 / 0 |
| f03 key columns | 30 (30) | 2,748 | 2,748 | 567 | 0 | **0** | 0 | 1,299 / 0 | 7,242 / 0 |
| f04 lookups | 30 (30) | 4,290 | 3,942 | 2,268 | 348 | **0** | 0 | 837 / 0 | 8,709 / 0 |
| f05 ranges | 30 (30) | 3,351 | 3,351 | 1,701 | 0 | **0** | 0 | 852 / 0 | 8,709 / 0 |
| f06 running / previous row / side cell | 30 (30) | 2,423 | 1,987 | 494 | 436 | **0** | 0 | 851 / 0 | 8,220 / 0 |
| f07 other sheet | 30 (30) | 858 | 858 | 243 | 0 | **0** | 0 | 348 / 0 | 6,753 / 0 |
| f08 defined names | 30 (30) | 2,574 | 2,574 | 972 | 0 | **0** | 0 | 882 / 0 | 8,220 / 0 |
| f09 tables | 30 (30) | 3,606 | 3,177 | 810 | 429 | **0** | 0 | 1,392 / 0 | 8,220 / 0 |
| f11 result types | 30 (30) | 3,003 | 3,003 | 0 | 0 | **0** | 0 | 1,785 / 0 | 7,242 / 0 |
| f12 footer | 30 (30) | 858 | 858 | 243 | 0 | **0** | 0 | 348 / 0 | 5,775 / 0 |
| f13 HYPERLINK | 30 (30) | 858 | 858 | 0 | 0 | **0** | 0 | 510 / 0 | 5,286 / 0 |
| f16 unsafe | 30 (30) | 684 | 684 | 0 | 0 | **0** | 0 | 510 / 0 | 7,242 / 0 |
| f14 two-row header | 30 (30) | 1,287 | 1,032 | 243 | 255 | **0** | 0 | 522 / 0 | 5,286 / 0 |
| f10 arrays | 30 (30) | 868 | 429 | 243 | 439 | **0** | 0 | 186 / 0 | 6,457 / 0 |
| **total** | **480 (480)** | **39,758** | **37,851** | **10,355** | **1,907** | **0** | **0** | **15,341 / 0** | **131,974 / 0** |

Skipped (1,907, never counted as passed): `MATCH` on its own in a table whose layout changed 348 · `ROWS` / `COUNTBLANK` over a widened table 603 · `COUNTA` over a
widened key column 262 · a formula that reads a header cell whose text was changed (two-row header joined, suffix added) 255 · `UNIQUE` / `FILTER` / `SORT` / `TAKE` over a widened table 439.
All of these are results that *should* differ from the source ("recomputed over the merged table"); their stored result is still checked (no stale result).

What the oracle found while it was being built (fixed, then re-run): (1) XlsxWriter's placeholder `0`s stored next to formulas Excel computes
as 78 / 212 (stale results → now "no result", §3 g); (2) the corpus itself had tied scores, so a lookup by score depends on the row order (corpus
fixed: unique scores); (3) the empty-string results Excel stores as `<v/>` were read as "no result" (§2.1). The first full run before those fixes reported
1,146 differences; the final run reports 0.

## 6. Notes, gaps, limits

1. **The grid shows kept formulas without a stored result as empty** until Excel opens the exported file (the app has no formula evaluator).
   They are counted in `formulaStats` / item `formulasRecalc`; the UI should show a placeholder such as 「計算待ち」 for `cell.f && t==='z'`.
2. Whole-table formulas are **recomputed over the merged table** (rule b) — an average / total / rank / `COUNTIF` over "the table" no longer
   sees source rows that matched nothing and does see rows of the base that the source lacks (blank imported cells). That is the intent and
   is reported (`formulaRangesWidened`), but numbers can differ from the source file.
3. A lookup by a **non-unique** value (two people with the same score) returns "the first match" — which depends on the row order. The
   corpus therefore uses unique scores (with ties a reordered merge cannot equal the source).
4. XlsxWriter stores `0` as the result of every formula and sets `fullCalcOnLoad`; all-zero results in such a book are treated as "no result".
   A real workbook whose formulas all evaluate to exactly 0 *and* asks for a recalculation is treated the same way (harmless: Excel recomputes).
5. Not covered: lookups into other sheets of the source (the other sheet is not imported), `INDIRECT`/`OFFSET` (positional), formulas that read
   cells outside the table (kept as values), `.xls`/`.csv`/`.ods` sources (values only by construction).
6. `tests/xlsx/foreign-producers.test.ts` accepted-difference for py-19 (XlsxWriter table): the premerge no longer copies its placeholder `0`s.

## 7. Reproduce

```
powershell -NoProfile -ExecutionPolicy Bypass -File tests\oracle\import-values\run.ps1            # corpus (regenerated when missing) + real Excel
powershell ... run.ps1 -Regenerate                                                                 # rebuild the corpus first (Python venv + Excel)
npx tsx tests/oracle/import-values/diagnose.ts --dir <corpus> --engine both                        # before / after tables
npx tsx tests/oracle/import-values/diagnose.ts --foreign --dir tests/fixtures/foreign/corpus --engine both
npx tsx tests/oracle/import-values/make-expected.ts                                                # rewrite tests/fixtures/formulaimport/expected.json
EXCEL_TESTS=1 npx vitest run tests/oracle                                                          # the oracle under vitest (fixtures)
powershell ... run.ps1 -Sources                                                                   # includeSourceSheets: the same comparison with the copies (round 9)
powershell ... run.ps1 -Random 14 -Seed 7                                                          # the property test: random workbooks, Excel computes the source (round 9)
npx tsx tests/oracle/import-values/make-expected.ts --copies                                       # rewrite tests/fixtures/formulaimport/expected-copies.json
```

## 8. Round 9 (task F2): the formulas fall back to the copies of the source's sheets (includeSourceSheets)

With `PremergeOptions.includeSourceSheets` the ORIGINAL sheets of every other file are in the output (named `<file>_<sheet>`, "(1)" for the one added later,
31 characters), so what the merged table cannot keep can read those copies and stay a formula. `FormulaRemapper` (`RemapSpec.sheetCopies`) works in three
steps, and the first one is exactly round 8:

1. **Merged table** (rules a-e): a formula that round 8 keeps is never touched (same text, same stored result; `viaCopy` unset). Checked on the whole corpus:
   every formula round 8 kept is kept with the same text.
2. **Hybrid**: every reference that rules a-e cannot keep reads the copy of its sheet at the ORIGINAL coordinates (`'<copy>'!C5`, the source row of the
   person — the copy keeps the original rows). A reference to another sheet of the source always reads its copy; defined names are inlined (a name of a side
   cell / another sheet, or an expression over absolute single cells, in parentheses), structured references are resolved to A1 on the copy (totals rows too).
3. **Copy mode** (every reference reads the copies = the source formula with renamed sheets): when a hybrid formula would be wrong — a range of the merged
   table and a range of a copy in one formula are not row-aligned (SUMIF / SUMPRODUCT / INDEX-MATCH / `SUM(a)/SUM(b)`), `INDEX(merged range, position read
   from a copy)`, reference operators (`:` / intersection / union) between operands of two sheets — and for `ROW(ref)`, `COLUMN(ref)`, `OFFSET(ref, …)`.

Stays a value (reason of round 8): `ROW()` / `COLUMN()` without an argument, `INDIRECT`, `ADDRESS`, `CELL`, `FORMULATEXT`, `SHEET(S)`, `INFO`, 3-D references,
external workbooks, spills, dynamic names (`OFFSET` in a name), and a **multi-cell range that now lives on another sheet in a place where Excel intersects it
implicitly** (a bare `Master!A2:A10*2`: the source took the cell of ITS row, a different row on the merged sheet) — kept only as a bare argument of a function
that takes a range, inside SUMPRODUCT / LOOKUP, or in an array formula.

Stored results: the copy reads the very same cells, so a formula that reads only copies (and same-row merged cells) keeps the source's result.
`MergeReport.formulaStats[].reasons.viaCopy` (informational, counted among `kept`; NOT part of `toValues + empty`) and the info item `formulasViaCopy`
(「N 件の数式は元のシート（コピー）を参照して数式のまま残しました」) report them.

| corpus (tests/fixtures/formulaimport, 19 files x 4 configurations, 16,032 formulas) | kept | values | empty cells | via a copy |
|---|---|---|---|---|
| round 8 (`expected.json`) | 12,746 | 1,536 | 1,750 | – |
| round 9 (`expected-copies.json`) | **15,074** | 436 | 522 | 2,328 |

What is left: positional (`ROW()-1`, `HYPERLINK(…ROW())`, `INDIRECT`) 560, dynamic names 112, 3-D references 112, external workbooks 112, odd syntax 58, spills 4.

Real Excel (hidden own instance; `run.ps1 -Sources` / `-Random`, results in the tables below; `EXCEL_TESTS=1 npx vitest run tests/oracle` runs a smaller version):

| run | packages | kept (via copy) | compared | MISMATCH | stale results | moved-column cells |
|---|---|---|---|---|---|---|
| generated corpus, 48 files x 5 configurations, `-Sources` | 240 | 46,654 (6,896) | 44,747 (10,355 widened) | **0** | 0 / 19,309 | 131,974 / 0 differ |
| random workbooks `-Random 14 -Seed 7` (3 row orders x 5 configurations) | 210 | 18,648 (11,639) | 18,147 | **0** | – | 44,352 / 0 |
| random workbooks `-Random 16 -Seed 11` (with the name variants) | 240 | 20,829 (12,992) | 20,055 | **0** | – | 50,688 / 0 |
| random workbooks `-Random 16 -Seed 23` | 240 | 20,379 (13,726) | 20,025 | **0** | – | 50,688 / 0 |

The random property test (tests/oracle/import-values/random.ts): a source workbook with the data sheet 評価 (ID, 氏名, three distinct-valued score columns, 8 formula
columns drawn from 40 templates: same-row arithmetic, side cell, footer, lookups into Master / レート, previous rows, running totals, whole-table ranges, half-imported
SUMIF / SUMPRODUCT / INDEX-MATCH, `ROW(ref)` / `OFFSET`, `ROW()`), 3 base people missing in the source, 3 people only the source has, in three row orders, imported
in 5 configurations (default imports with join left / full; the formula columns plus only C, plus only D and E, or every other formula column — so that formulas read
columns that are NOT imported). Names vary: plain, with spaces / parentheses / an apostrophe / a leading digit,
and long enough that the two long sheet names are cut to 31 characters and collide ("(1)"). Excel computes the source; every kept formula of the merged workbook is
compared with the source value of the same person (formulas over a whole table are compared with the source after Excel deleted the rows that were not placed).
Mutation check: with the alignment rule removed from `remap.ts` the same run reports 365 mismatches in 4 workbooks (all `SUMIF($E…,$C…)` and `INDEX($C…,MATCH(E…,$E…))`
over a half-imported table); with it, 0.

Hazards the design had to rule out (unit tests in tests/formula/remap-copies.test.ts; the mutation check above proves the first one): a range of the merged table
next to a range of a copy (the rows are not aligned); `INDEX(range of the merged table, position read from a cell of the source)`; a bare multi-cell range on another
sheet (Excel intersects it with the row of the FORMULA = the merged row, not the source row); `:` / space / union between operands of two sheets.

### Known gaps (round 9)

- `ROW()` / `COLUMN()` without an argument stay values (they are the position of the formula cell, which changes).
- `INDEX(range, cell)` where BOTH the range and the cell are imported and the merged order differs from the source order: kept live by round 8 (the cell is read as a
  position into the merged table, the source meant its own row order). Repro: source rows P3 P1 P5 …, `=INDEX($C$2:$C$11,E2)` with C, E imported and a base in another order.
- A copy cell that A1 could not copy as a formula (structured reference / 3-D / external reference without a stored result) is empty on the copy: a formula that reads it
  reads an empty cell.
- `HYPERLINK("#'評価'!A1", …)` text constants are not renamed (they never were).
