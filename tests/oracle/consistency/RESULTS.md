# Consistency oracle in real Excel — results (task V1)

> User requirement: *"when columns are swapped, or rows get rearranged because rows are matched by name, the formulas must change
> consistently (stay correct)"*. This oracle proves it in real Excel, in BOTH export engines, and says exactly where Excel itself changes a
> value (and that the doc warned about it).

**Verdict (two independent runs, 554 random scenarios, 548 of them also through the app's Excel engine, 538,469 cells): no unexplained value change, no new
error, no disagreement between the two engines.** The oracle found and I fixed two real defects (inserted column next to a hidden column, silent
VLOOKUP key-column swap), found one premerge defect that is NOT fixed (reported below), and one product limit (convert-to-values of formulas
that have no stored result). Everything else it reports is a documented Excel behaviour, counted.

Re-run: `powershell -NoProfile -ExecutionPolicy Bypass -File tests\oracle\consistency\run.ps1 -Count 340` (about 75 minutes: one Excel start
per scenario for the app's Excel engine + one hidden Excel per 40 workbooks for the recalculation; everything else is Node). Options: `-Out`,
`-SeedBase`, `-Only <regex of set ids>`, `-NoRunner` (no app Excel engine), `-NoExcel` (Node half only), `-Reuse` (re-judge an earlier output
folder). vitest: `tests/oracle/consistency/oracle.test.ts` (Node half always, the whole oracle with `EXCEL_TESTS=1`, `EXCEL_ORACLE_COUNT`),
`regress.test.ts` (one test per defect found). Scratch tools: `show.mjs <out> <scenario>` (ops, runner answer, findings), `inspect.mjs <dump>
<sheet> <range>`, `report.mjs <out>` (the tables below), `mutate.mjs` (plants bugs in a COPY of the repository — see "Does the oracle see bugs?").

## 1. Method

```
scenario = set (base + other files) x variant (join left/full, single-sheet output on/off, import default/all columns) x seed

  app pipeline in Node      readBook -> keys (analyzeSources, defaultPlan, joinSources + keyPairs) -> buildPremerged -> readBook -> createEditor
  random op log             doc API only: moveColumns (1-4 columns, contiguous or not, anywhere; title merges unmerged), insertColumns,
                            deleteColumns (convertDependents when something would break), setCellInput (constants and formulas)
  BASELINE                  the premerged package + the typed values (written at the columns' ORIGINAL positions) + the values the doc wrote for
                            converted formulas — NO move / insert / delete at all.  Excel opens it and recalculates fully (CalculateFull)
  built-in engine           replayLog(premerged, log)                          -> <id>.js.xlsx
  Excel engine              server/excel/excel-runner.ps1 through server/excel/run-job.mjs (the app's own path, one Excel per job) -> <id>.xl.xlsx
  judge                     ONE hidden Excel per batch of 40 workbooks opens baseline, js, xl (a file that needs repair is a failure),
                            CalculateFull, dumps value + Formula of EVERY cell of EVERY sheet; cells are found again by (sheet, row, COLUMN ID)
```

What is checked per scenario (class names in the tables):

| check | class when it fails |
|---|---|
| every cell has the value it had in the baseline, found again by column id (typed values/formulas, appended rows, all sheets) | `value-changed`, `new-error`, `new-ref-error`, `lookup-silent` |
| the two engines' outputs: same value AND same Excel `Formula` text in every cell of every sheet | `engine-value`, `engine-formula` |
| the two outputs as packages: column widths / hidden / outline, names, merges, tables, autofilter, CF and DV by the cells they cover (`tests/com/lib.ts packageDiff`, `tools/lib/package-diff.mjs`) | `engine-package-*`, `cf-quirk` |
| a reference to ONE cell / ONE column follows its column id (in the doc, and in both outputs' formula text) | `doc-vs-replay` (`ref-identity(doc)`), `ref-identity-js/xl` |
| the header text found at the column's new position; inserted columns are blank | `layout-header`, `inserted-not-blank` |
| the doc's own view of the result == the package read back from replayLog (`tests/doc/compare.ts`) | `doc-vs-replay` |
| `deleteColumns(convertDependents)` writes the value Excel computes; a blank is announced | `convert-stale`, `convert-blank`, `convert-blank-silent` |
| a name / CF / DV rule that ends with `#REF!` was announced by the doc | `ref-error-silent` |
| join 'full': formulas below the table grow over the appended rows, every other reference keeps its target | `append-range`, `append-range-running` |
| **row rearrangement**: every imported formula that stayed a formula has the value the same formula has in the SOURCE file (recalculated by Excel) for the SAME PERSON (whole-table ranges: against the source with the unplaced rows deleted); the person of every placed row agrees on a key pair | `src-value`, `row-person` |

A value may differ from the baseline only for a reason that is **proved, not assumed**, and every such change is counted:
*range-span* (a column was moved into / out of / deleted from the middle of a multi-column range, so Excel itself changes the range — the final
formula's column ids differ from the baseline's), *lookup-warned* (a VLOOKUP/INDEX/CHOOSE by column number: the doc must have emitted the
lookup warning, else `lookup-silent`), *positional* (COLUMN/OFFSET/INDIRECT …), *cascade* (the cell reads an explained cell). A scenario that is
"strict" (about 60 %) forbids ops that would change any range's span, so there EVERY value must be unchanged — except lookups by column number, which only the doc's warning explains (and a range-span explanation in a strict scenario is itself a failure, `oracle-inconsistent`). Also never accepted, explained or not:
`#REF!` / `#NAME?` (except the documented VLOOKUP out-of-range value), a circular reference (no defined value in Excel: the output is not judged and
the generator refuses ops that make a formula read its own cell).

Scenario sets (`sets.ts`, `gen.ts`): the app's samples (`01_学生名簿` with title merge, hidden helper column, CF/DV, 集計 sheet with whole-column,
cross-sheet and name references, chart, pivot; base 02 with 平均/判定; base 05; the 5-file demo), a generated **rich base** (relative / absolute /
mixed / whole-column / cross-sheet / defined-name / structured-reference / VLOOKUP-by-number formulas, total row, report sheet, CF + DV with formulas;
variants plain, Excel Table, merged title, quoted sheet names) merged with the 15 formula sources of F1's corpus (`tests/fixtures/formulaimport`,
results stored or not) and with generated second files whose rows are **reversed / shuffled / have extra and missing people / duplicates / all of it**
(`row-*`, plain and Table sources).

## 2. Final run (335 scenarios, seeds 1…, `run.ps1 -Count 340`)

| family | scenarios | ops in the logs | cells compared | formula cells | changed vs baseline (all explained) | explained by |
|---|---|---|---|---|---|---|
| samples (smp-*) | 16 | 105 | 26,799 | 3,344 | 226 | range-span 174, cascade 26 |
| rich base + one F1 formula source (rich-*) | 177 | 1,019 | 157,138 | 72,241 | 2,005 | range-span 1,317, lookup-warned 575, cascade 54 |
| rich base + 3 sources | 6 | 52 | 10,513 | 6,095 | 230 | range-span 170, cascade 60 |
| rich Table / merged-title / quoted-name base | 35 | 190 | 33,551 | 19,001 | 712 | range-span 583, lookup-warned 91, cascade 38 |
| rows rearranged (row-*) | 101 | 591 | 95,646 | 53,541 | 1,373 | range-span 1,137, lookup-warned 120, cascade 116 |
| **total** | **335** | **1,957** | **323,647** | **154,222** | **4,546** | range-span 3,381 · lookup-warned 786 · cascade 294 |

* ops accepted by the doc: 679 moves (+14 with title-merge unmerge), 266 inserts, 126 deletes (75 scenarios), 411 typed values/formulas; 4 refused by
  the doc (Excel rules), 397 not taken by the generator (they would change a range's span in a strict scenario, or make a formula read itself).
  226 scenarios `inv` (moves/inserts/typed cells), 109 `del`; 159 join 'full', 168 single-sheet output, 197 strict; 117 produced the lookup warning.
* Excel engine: 330 ran in real Excel (median 5.5 s per job); 5 answered `MOVE_VERIFY_FAILED` = in-table column moves of the Excel Table bases,
  the documented case where the app falls back to the built-in engine (class `runner-fallback`, not a failure). **330 of 330: values and Formula text
  identical to the built-in engine in every cell of every sheet; no package difference except the CF quirk below.**
* **No repair**: all 1,501 workbooks of the run (baselines, both engines' outputs, sources) opened normally.
* Explained changes that END in an error: 938 `#N/A` (VLOOKUP whose key column was swapped — warned by the doc), 700 `#REF!` as a VALUE (VLOOKUP with
  column index 4 over a range that lost a column — warned; the formula text holds no `#REF!`), 1 `#CALC!`, 1 `#VALUE!` (a SUM range that now covers a
  column with an error value).
* **Row-rearrangement angle** (`srcCheck.ts`): 274 (set, variant) pairs, **51,302 imported formula cells that stayed formulas compared with the source file's own
  Excel-recalculated value for the same person — 0 differences** (10,061 of them whole-table ranges compared with the source after Excel deleted the
  unplaced rows), 3,087 skipped by rule (position results, `ROWS`/`COUNTBLANK` over a widened table … see tests/xlsx/formula-import-findings.md §5),
  7,265 placed rows checked for the person by key text — 0 differences:

| source order | pairs | imported formulas compared | placed rows | of them whole-table ranges | merged row != source row in |
|---|---|---|---|---|---|
| same | 7 | 2,310 | 210 | 0 | 0 pairs |
| reversed | 12 | 3,624 | 360 | 0 | 12 pairs |
| shuffled | 12 | 3,624 | 360 | 0 | 12 pairs |
| 4 extra people (join full appends them) | 12 | 3,852 | 384 | 1,080 | 12 pairs |
| 5 people missing | 12 | 3,024 | 300 | 1,812 | 12 pairs |
| 2 people twice | 12 | 3,624 | 360 | 2,172 | 12 pairs |
| all of it (mix) | 12 | 3,246 | 324 | 1,944 | 12 pairs |
| Table source (rowtbl-*) | 3 | 844 | 81 | 308 | 2 pairs |

* Node half only (no Excel), 1,965 more scenarios (seeds 20000…: samples 93, rich 1,203, rows 669) with the final code: **0 `doc-vs-replay` findings** (the doc's view == the package read back
  from replayLog; every one-cell/one-column reference of the doc still points at the same column id) — 3,943 moves, 773 deletes, 1,562 inserts, 2,350 typed cells.

### 2b. Second, independent run (new seeds 1001…, `run.ps1 -Count 220 -SeedBase 1001`, made after the two fixes below)

219 scenarios (samples 11, rich 139, rows rearranged 69; 52 sets), 1,182 ops (406 moves + 8 with unmerge, 164 inserts, 117 deletes in 65 scenarios, 245 typed cells), **214,822 cells compared
(101,027 formula cells)**; 2,470 changed vs baseline, all explained (range-span 1,763, lookup-warned 335, cascade 318; they end in 391 `#N/A`, 520 `#REF!` as a lookup value, 64 `#VALUE!` propagated from typed
cells); 218 ran in real Excel (1 `MOVE_VERIFY_FAILED` in-table move) and equal the built-in engine in value and Formula text; **44,058 imported formulas equal their source for the same person**
(9,617 against a reduced source; 6,385 placed rows checked); 1,396 converted formulas got Excel's value, 521 were blanked (announced). Classes: only `append-range-running` (open, reported), `cf-quirk` 2,
`convert-blank` 6 scenarios, `ref-error-announced` 3, `runner-fallback` 1 — no failure.

**Together: 554 random scenarios, 548 of them through the app's Excel engine as well, 538,469 cells (255,249 formula cells), 95,360 imported formulas compared with their source for the same person.**

## 3. What the oracle found

| # | class | who was wrong | status |
|---|---|---|---|
| 1 | `engine-package-cols` — a column inserted next to a HIDDEN column | **doc + replayLog**: both copied `hidden` from the left column into the new column; Excel inserts a VISIBLE column (measured: format from left or right, hidden column, hidden + grouped, inside a collapsed group; width and outline level ARE copied) | **fixed**: `src/core/doc/struct.ts makeInsertedColumns`, `src/core/xlsx/patch/cols.ts inheritAttrs` (one line each) + `regress.test.ts` (unit + real-Excel). A known deviation noted in `tests/xlsx/dropsheets-engines.test.ts` ("replayLog copies hidden …"); that comment is stale now |
| 2 | `new-error` / `lookup-silent` — swapping the KEY column of a VLOOKUP range (`VLOOKUP(A2,$A$2:$D$31,4,FALSE)`, A and B swapped) | **doc / excelRef.lookupWarnings**: the warning only looked at the RETURNED column (index 4 stays D), not at the SEARCH column (first column of the range). Every lookup became `#N/A`, silently (Excel does not adjust it either) | **fixed**: `excelRef.lookupWarnings` also warns when the first column of a VLOOKUP range changes, with its own sentence (「先頭の列（検索する列）で探します …」) + `regress.test.ts`. No existing test changed; the whole suite passes (npx vitest run: 3,038 passed, 48 skipped, 0 failed; tsc clean) |
| 3 | `append-range-running` — join 'full' | **premerge / `patch/formula.ts extendArea`** (not mine): the range-extension rule that makes a total row below the table include appended rows also extends a RUNNING TOTAL in the last base row: `=SUM($D$2:D31)` (row 31) becomes `=SUM($D$2:D35)` and sums the appended people | **fixed in round 9 (task A1)**: `patch/formula.ts extendArea` no longer extends an area whose end is the formula's own row with a relative end (a formula INSIDE the data rows); a total row BELOW the table still grows. `regress.test.ts` holds the minimal repro as a plain `it` (it was `it.fails`), `tests/xlsx/addsheets-runningtotal.test.ts` the rule; `check.ts KNOWN_OPEN` is empty |
| 4 | `convert-blank` — deleteColumns(convertDependents) | **product limit** (no formula evaluator): a formula without a stored result (every imported formula premerge keeps live without result — F1's gap #1) is converted to an EMPTY cell; 974 of 2,378 converted formulas in the run (1,404 got the exact value). The doc DOES announce it afterwards (「計算結果が保存されていなかったため空欄になりました」: asserted, 0 silent) | **open, reported**. Idea: let the Excel engine (or a hidden Excel at merge time) compute the cached values once, so the grid and the convert path have them; at least warn BEFORE the user confirms (DeleteImpact has no field for it: contractIssues) |

Not defects (documented behaviour, classes `info` in the summary):

* `cf-quirk` (1 scenario): Excel's documented conditional-format quirk (`tests/formula/cfProbe.test.ts` KNOWN `abs-ref-to-moved-column`, golden RESULTS CF_INSIDE_MOVE):
  a rule whose formula holds an ABSOLUTE column reference (`$C2`) loses the cells left of the moved column when that column moves inside its range.
  Reduced by the oracle (greedy op-log reduction against the two engines) to two ops: move column A to the right of G, then swap B and C. CF coverage only, never a value.
* `pivot-*`: PivotTables of the kept sheets. After a typed value in the pivot's source the built-in engine's pivot refreshes on load (`refreshOnLoad`), the Excel
  engine's saved pivot is stale: both fine, so pivot sheets are compared only in scenarios without typed/converted values, and where BOTH engines agree
  but differ from the baseline Excel itself changed the pivot (a moved source column left the pivot's source range): `pivot-value-changed`, info. Excel
  autofits the columns of a refreshed pivot sheet: `pivot-engine-package-cols/widths`, info.
* `circular-*` (3 scenarios): a column moved INTO the range of a formula that lives in it (`=AVERAGE(O2:T2)` in column R): Excel's value is undefined; the
  generator now refuses the ops that cause it directly, the rare indirect ones are reported and not judged.
* `ref-error-announced` (4 scenarios): a defined name / CF / DV rule whose only column was deleted ends with `#REF!` (cannot be converted to values); the doc says so.
* Filtered as benign (`pkgdiff.ts`): Excel writes `xl/richData` parts when it saves a dynamic-array workbook and infers `calculatedColumnFormula` for Table
  columns; neither changes a value.

Defects of the ORACLE itself that were found and fixed while building it (so that nobody re-introduces them): a one-column area (`SUM(O2:O2)`) is a span
too; typed formulas must read constants only (a typed `=SUM(C48,Q48)` into a cell that `Q48 = AVERAGE(N48:P48)` reads is a circular reference); Excel
reports a circular reference per sheet (`Worksheet.CircularReference`) — outputs with one are not judged; an inserted column in a Table gets its header from Excel;
the person check is only meaningful where both files write IDs the same way (the sample files do not on purpose); the pivot differences above.

## 4. Facts measured in Excel during this task (own hidden instance)

* Insert a column next to a hidden column: the new column is visible (`Hidden = False`) for `xlFormatFromLeftOrAbove` and `…RightOrBelow`; with the left column hidden AND
  grouped, and inside a collapsed group too. The stored `<col>` of the new column has the left column's `width` / `customWidth` / `outlineLevel`, never `hidden` or `collapsed`.
* Moving a column into the middle of a multi-column range makes Excel extend the range (`=AVERAGE(D2:F2)` + column C moved between D and E → `D2:G2`): a value changes
  legitimately. Moving it out shrinks it; deleting a column inside shrinks it (no `#REF!`).
* `VLOOKUP(x,A:D,4,0)` after a range column was deleted: `#REF!` as a VALUE (index beyond the range) — the formula text keeps no `#REF!`.
* A formula that reads a cell depending on itself is a circular reference: the value is undefined (two files with the same circular `AVERAGE(O2:T2)` in column R came out as 89 and "").
* Excel adds `calculatedColumnFormula` to Table columns and `xl/richData/*` parts to dynamic-array workbooks when it saves.

## 5. Does the oracle see bugs? (`mutate.mjs`, a COPY of the repository gets one planted defect; 40 scenarios each)

| id | planted defect | result |
|---|---|---|
| M0 | none (control) | **clean**: only the reported open class, runner fallbacks and convert-blank |
| M1 | `excelRef.moveSpan`: a block inside a range whose end it moves inward keeps the "follow" model | **detected** `engine-formula` (built-in `$J$2:$P$35` vs Excel `$A$2:$P$35`) — only real Excel can see it (doc, replay and excelRef agree with each other) |
| M2 | `excelRef.deleteSpan`: a range whose end is in the deleted block keeps one column too many | **detected** `engine-formula` (`$A$2:$D$31` vs `$A$2:$H$31`) |
| M3 | replay (`patch/formula.ts`): sheet-qualified references are not rewritten | **detected** `value-changed-js` 7, `ref-identity-js` 36, `doc-vs-replay` 16 (16 scenarios) |
| M4 | doc: defined names not adjusted | **detected** `doc-vs-replay` (8) |
| M5 | doc: formulas of other sheets not adjusted | **detected** `doc-vs-replay` (4) |
| M6 | premerge row mapping (`formula/remap.ts`): every 7th source row one row too far down | **detected** `src-value` (48) |

## 6. Limits — what this does NOT prove

* Excel's own behaviour is the reference: a defect that Excel's UI has and the runner's clipboard-free sequence does not (or vice versa) is out of reach (the CF quirk is the known one).
* Typed formulas read constants only and are typed once (no chains); a circular reference through other cells is reported, not judged.
* Pivot tables are compared only where no typed value changes their data; charts are not compared (package view only).
* Sources of `.xls` / `.csv` are values only by construction; the generated `row-*` sources are `.xlsx` with formulas.
* The controller does not hand `JoinResult.keyPairs` to buildPremerged yet (src/state/controller.ts): the oracle passes it (as the lead will); without it key-column formulas
  of imported files become values (F1's gap) — the invariance proved here is independent of that.
* Time and disk: one run needs ~75 minutes and leaves ~70 MB (dumps, packages); Office writes ~80 MB of logs per Excel start into a private TEMP that is deleted afterwards (the run stops when less than 1.5 GB is free).
