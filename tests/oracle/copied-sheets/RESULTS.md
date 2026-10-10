# Copied-sheets oracle in real Excel — results (task V9, round 9 "plan A")

> User requirement (translated): *"the merged file must ALSO contain every other sheet — the original sheets of the 2nd, 3rd ... file as well —
> without any error; the merged sheet is first; names are `<file>_<sheet>`, and `(1)`, `(2)` on the sheet that was added later when a name is still taken."*
> This oracle proves it in real Excel, in BOTH export engines, after random edits of the merged sheet.

**Verdict (final run, 72 combinations, 78 random-edit scenarios, 355 copied sheets, ~120,000 cells, 337 workbooks opened and recalculated by Excel):
no repair prompt, no new #REF! / #NAME?, every sheet name equals `planSheetNames` in the planned order, every copy equals its source, the edits never
change a copy, imported formulas that read a copy give the source's value for the same person, and the two engines write the same thing
(0 formula differences).** The oracle found ONE real product defect (frozen panes after the whole frozen block is moved or deleted, fixed with
regression tests, below). Everything else it reports is a documented limit of the copier or a difference in how Excel writes things, counted as INFO.
27 of 27 planted faults are found by the same judge (negative controls).

Re-run (needs Excel; every Excel is our own hidden instance, with its own TEMP below `-Out`, deleted afterwards):

```
powershell -NoProfile -ExecutionPolicy Bypass -File tests\oracle\copied-sheets\run.ps1 -CheckOutput 10 -Negative      # everything, about 1 hour
powershell -NoProfile -ExecutionPolicy Bypass -File tests\oracle\copied-sheets\run.ps1 -Only '^(smp-01-all-keep|ob-all-keep)$' -CheckOutput 4 -Negative
powershell ... run.ps1 -Out <dir> -Reuse                         # re-judge an earlier output folder (a dump that is bad / too short is read again)
node tests\oracle\copied-sheets\run.mjs --out <dir> [--only <regex of combination ids>] [--limit N] [--check-output N] [--negative] [--seed-base 0] [--no-runner] [--no-excel] [--reuse]
EXCEL_TESTS=1 npx vitest run tests/oracle/copied-sheets          # the same from vitest (EXCEL_ORACLE_ONLY narrows it); without EXCEL_TESTS only the Node half runs
```

Options of `run.ps1`: `-Out`, `-Only`, `-Limit`, `-CheckOutput N` (the QA tools over N scenarios per group), `-SeedBase`, `-Negative`, `-NoRunner` (no app Excel
engine), `-NoExcel` (Node half only), `-Reuse`. Output folder: `results.md` / `results.json` (the tables), `findings.json`, `scenarios.json`, `dump\<id>.book.json`
(what Excel read), `pkg\` (premerged, base, `.js.xlsx`, `.xl.xlsx`), `negative.json`, `tools.json`.

## 1. Method

```
combination = base + other files (card order) x options (join type, dropOtherSheets true / false, includeSourceSheets = true)   72 of them
scenario    = combination x random edit log (seed)                                                                              78 of them

  app pipeline in Node      readBook -> keys -> buildPremerged (includeSourceSheets) -> readBook -> createEditor        (pipeline.ts: mergeWithCopies)
  random op log             tests/oracle/consistency/ops.ts: moveColumns (1-4 columns, non-contiguous too, title merge unmerged on request),
                            insertColumns, deleteColumns (convert), typed cells (constants and formulas) — accepted: edit 97, insert 65, move 110
                            (39 with unmerge), delete 23
  BASELINE                  the premerged package + the typed values at the columns' ORIGINAL positions, no structural op. Excel recalculates fully
                            (CalculateFull); this is what every later output must still say
  built-in engine           replayLog(premerged, log)                                  -> <id>.js.xlsx
  Excel engine              the app's own path (server/excel/excel-runner.ps1 through server/excel/run-job.mjs, one Excel per job) -> <id>.xl.xlsx
  Excel dumps               ONE hidden Excel per batch of 24 workbooks (excel-book.ps1): open (a repair prompt = failure), CalculateFull, every worksheet's values
                            (Value2) and formulas from A1, sheet type / visibility, defined names, and for the copies a "deep" dump: sampled cell looks (number
                            format, font, fill, alignment, borders), widths, hidden rows / columns, conditional formats, validations, links, shapes, charts
                            with series, tables, notes, pivot tables
```

What is judged (`judge.ts` + the consistency oracle's `check.ts` / `srcCheck.ts`, `pkgdiff.ts`, `tests/xlsx/addsheets-helpers.ts`):

| check | class (FAIL) |
|---|---|
| no repair / failed open, no lint finding in the package (relationships, content types, names, ids) | `open`, `lint-*` |
| sheet names exactly equal `planSheetNames`, order = merged sheet, base's other sheets (when kept), then every other file's sheets in card order | `names`, `names-order` |
| every copy equals its source cell by cell in the premerged package, and each engine's output equals the premerged package (values, formulas modulo the renamed prefix, merges, hidden, widths, CF, DV, tables, charts, pictures, shapes, notes, hyperlinks, pivots) | `copy-value`, `copy-formula`, `copy-format`, `copy-visibility`, `copy-object`, `copy-table`, `copy-pivot`, `copy-names*` |
| no workbook-level name leaks out of a source (a source's names become sheet-scoped names of the copy that reads them) | `copy-names-leak` |
| the edits never change a copy (copy of the output = copy of the baseline) | `copy-changed` |
| the merged sheet stays value-invariant under the ops (the consistency judge: every moved / inserted / deleted column, typed cells, converted formulas) | `value-changed`, `ref-identity`, `parity`, ... |
| imported formulas that read a copy give the source's value for the same person | `import-*` (srcchecks) |
| no new #REF! / #NAME? in the merged sheet | `merged-new-error` |
| both engines' outputs are identical (values, formulas, the package view of tools/compare.mjs) | `engine-value`, `engine-formula`, `engine-package-*` |
| the QA tools agree: `tools/check-output.mjs --single-sheet / --with-copies --source ...` has no FAIL for a shape reason, and `compare` of the two outputs has 0 formula differences (`-CheckOutput N`) | `tool-check-output-*`, `tool-compare-formulas` |
| the oracle sees planted bugs (`-Negative`, `negative.ts`): value, formula, sheet name, visibility, conditional formats, column width, data validation, hidden column, chart series, picture, leaked name | `negative-*` |

How a copy is compared: the copies are compared with their sources **in the premerged package as Excel renders it** (so a font that Excel itself completes
or a default the copier cannot see is not a difference); each engine's output is compared with the premerged package. The source fidelity depends on the
file type: `.xlsx` / `.xlsm` = everything, `.xls` = values only (the app reads BIFF for values), `.csv` = the app's own reading only (Excel parses a csv by its
own rules: a leading zero, dates), so for csv the sheet name and the values of the app's reader are judged, not Excel's.

## 2. Coverage of the final run

| | |
|---|---|
| combinations / scenarios | 72 / 78 (smp 16 · cor 37 · ob 12 · nam 4 · rich 9; 0 skipped) |
| output shapes | dropOtherSheets true 31 · base's other sheets kept 47 · join `full` 5 · `.xlsm` output 1 |
| files | samples 01..05 (xlsx, csv, xls) as base and in several card orders; 37 corpus scenarios (Google Sheets, LibreOffice, Mac Excel 2011 / 1904, Strict, SheetJS / ExcelJS / openpyxl / XlsxWriter / pandas, hidden / very hidden first sheet, many sheets, tables, external links, macro-enabled, protected ...); A2's Excel-authored object fixtures (charts, shapes, notes, tables, pivots, slicers / controls, Strict); three files with the same name; a long file name; forbidden characters; tricky sheet names; the generated rich base with F1's formula sources |
| copied sheets | 355 (345 compared with their source cell by cell: ~120,289 cells, 12,066 sampled cells rendered by Excel) |
| imported formulas compared with the source (same person) | 8,104 |
| workbooks opened + recalculated by Excel | 337 (15 batches of up to 24) |
| app Excel engine | 77 ok · 1 fell back (REFUSED_TABLE, a documented Excel limit: a table column cannot be moved out of its table) |
| QA tools | 20 verdicts (both outputs of 10 scenarios) + 10 engine comparisons: 0 findings |
| negative controls | 27 / 27 planted faults found (11 kinds) |

## 3. What it found

### 3.1 Product defect, fixed: frozen panes after the whole frozen block is moved away or deleted (class `engine-package-view`)

Both engines must write the same sheet view. After a structural op that removes **all** frozen columns (or rows) — deleting them or moving them behind the
freeze — the built-in engine (`replayLog`) and the editor's document kept `xSplit = k`; **Excel leaves exactly ONE frozen column / row** (measured for
k = 1..4, columns and rows, in a hidden Excel of our own and in the app's Excel engine).

Minimal repro: sample `01_学生名簿.xlsx` (the roster freezes C4 = 2 columns, 3 rows); move `学生ID` + `氏名` to the right (any slot behind column C).
Excel engine: `<pane xSplit="1" ySplit="3" topLeftCell="B4">`; built-in engine and doc: `xSplit="2"`.

Fixed in two places that must agree (both replay the same op log): `src/core/xlsx/patch/views.ts` (`frozenAfter`, the delete step of an op: a delete that covers
`1..k` leaves one) and `src/core/doc/struct.ts` (`freezeOp.cut`: `at <= 1 && at + cnt - 1 >= k ? 1 : ...`). `views.ts` belongs to the patcher (P): the change is two lines
with a comment, no other behaviour moved. Regression tests: `tests/oracle/copied-sheets/regress.test.ts` (`frozenAfter` for k = 1..4 and the cases that must stay
as they were, `adjustSheetViews` rebuilding `<pane>`, the doc / replay agreement on the sample, and — with `EXCEL_TESTS=1` — real Excel through the runner).
The existing `tests/doc` fuzz tests failed until both sides were fixed consistently, which is how the doc/replay agreement is kept honest.

### 3.2 Defects of the QA tools, found by running them over the oracle's outputs, fixed

| what | fix |
|---|---|
| `check-output --with-copies`: a csv source's copy was named after the temporary file Excel opened (`03_連絡先_8600e48f87e9`) | the only sheet of a text file is named after the file, like the app's reader does (`copies.mjs: textSheetName`) |
| a pivot table renamed `Name_2` by Excel (names are unique per workbook) was a "different" pivot table | pivot names are compared modulo `_N` (`copies.mjs`) |
| charts of a Strict source copied into a normal workbook (a documented limit the app warns about) were a FAIL | `copy-leftout` WARN when exactly the Strict / normal gap applies (`tools/lib/strict.mjs` detects Strict from the workbook part); any other difference of that copy still fails |
| package view: `t="d"` cells (ISO dates of a Strict source, copied as they are; Excel reads them as the date serial) were "different" from the Excel engine's number | `package-diff.mjs: isoDateSerial` |
| `dump-check`: the text cells a pivot table draws are re-rendered by Excel when the source changed, which looked like an incomplete Excel dump (`inspect-mismatch`) | the cells inside a pivot table's location are not compared (`dump-check.mjs`) |
| engine comparison noise (the package view reported how Excel writes things, not what it holds): cell-style name `Normal` vs `標準`, `General` case, numeric pane splits, true/1 view booleans, default `sheetProtection`, default `baseColWidth` / `zeroHeight`, an empty `sheetFormatPr`, a custom default row height, `$` in chart series refs | normalized in `package-diff.mjs`; each has a unit test in `tests/qa/package-diff.test.ts` |

### 3.3 Oracle artifacts (not product defects) found and handled

* A formula typed into a **Text-formatted** cell stays text in Excel; the consistency baseline assumed a formula. The scenario now records the typed text as
  the baseline value (`origText`), and the dump treats a string cell whose Value2 equals its Formula text as a constant. (Belongs to the consistency oracle V1's baseline.)
* The package linter of A1 reported `REL_DANGLING` for a shape hyperlink target `#'Sheet'!A1` — Excel writes hyperlink targets of shapes that way. Filtered in the oracle.
* **A hidden Excel answers short about once in 50 dumps**: `UsedRange` one row short while the file is fine (7 of 337 dumps in the final run; three scenarios showed
  an "engine-value / value-changed-js" FAIL for a cell in the last row that the file does contain). A re-dump of the same file gave the right answer. `excel.ts`
  now compares every dump with the package (`packageLastCells` / `isShortDump`: the dump's used range must reach the last cell that holds a value or formula)
  and reads a short dump again in a fresh Excel. An Excel instance that dies in the middle of a batch is detected and replaced by another of our own (`excel-book.ps1:
  Alive / Restart-Excel`), the jobs that failed are redone.
* Excel's chart-sheet `Type` is the chart type, not the sheet type: the sheet type comes from the Worksheets / Charts / DialogSheets collections.
* Workbook-level names of a source become sheet-scoped names on the copy that reads them (measured); the names judge is workbook-wide with that rule.
* Excel leaves one frozen column when all are removed (3.1); blank cells take the workbook's default style, so blank cells are compared for fill, border and merge only.
* A 1904-date source: dates in the copy differ by 1462 days in the 1900 workbook (a documented property); judged through the intersection of the runs with and without the shift.

### 3.4 Not fixed — known gaps (all counted as INFO by the oracle, each with its class)

| class | what | minimal repro / where |
|---|---|---|
| `copy-valign` | a cell without an explicit alignment keeps Excel's default vertical alignment of the *source* workbook's Normal style; in the copy the destination's default applies (bottom vs center) | `tests/fixtures/foreign/corpus/js-10-write-excel-file.xlsx`: R2C1 `-4108 -> -4107` |
| `copy-font-name` | fonts that the source gets from its theme (scheme fonts, e.g. ＭＳ Ｐゴシック written by an old Excel, or the theme's minor font) are written with the destination's theme font name | `tests/fixtures/foreign/corpus/js-02-sheetjs-date1904.xlsx` — A1 documented this |
| `copy-formula-lenient` / `copy-value-lenient` | external-workbook, 3-D and structured-reference formulas cannot be kept: the copy holds the stored value (the value is still the source's) | `tests/fixtures/foreign/corpus/hc-37-external-link.xlsx` (documented in README) |
| `copy-object-leftout` | what `sheetCopyLimits` names: slicers, timelines, SmartArt, ink, form controls / ActiveX / OLE, header/footer pictures; threaded comments become plain notes (the note counts differ by the replies) | `ob-comments`, `ob-slicer-control`, `ob-strict` (charts / shapes between Strict and normal) |
| `runner-fallback` | the app's Excel engine refuses to move a table column outside its table (REFUSED_TABLE, an Excel limit); the app falls back to the built-in engine | `ob-table-as-base-keep__s1` (1 of 78) |
| (not an oracle class) | **pivot tables whose source columns were edited**: the built-in engine keeps the pivot's stored rendering, Excel re-renders it when the source moved; the verdict says `pivot-source` for both. Existing behaviour of the pivot patcher (pivots refresh when the user refreshes), not part of this feature | `smp-01-full-keep__s2` (`tools\compare.mjs` on the two outputs: `values` of the `ピボット` sheet) |
| (not measured) | row heights and the geometry of drawings (anchor offsets, picture sizes) are not read by the dump | — |
| (not run) | `tests/fixtures/foreign/corpus/hc-31` (xlsm that Excel itself cannot open) and the user's very large files: unit-tested only; the speed of a 10 sheets x 20,000 rows x 30 columns copy was measured earlier (17-21 s, ~2 GB) | — |
| (observation) | a Strict source's date cells (`t="d"`) are copied into a normal workbook as written; Excel 2010+ reads them as dates (measured), older readers may not | `tests/fixtures/foreign/corpus/hc-08-strict-ooxml.xlsx`, sheet `hc-08-strict-ooxml_名簿` F2 |

### 3.5 INFO classes explained

`copy-font-name`, `copy-valign` — see above. `copy-object-leftout` — the sheet lost something `sheetCopyLimits` names. `info-engine-note-linebreak` — Excel stores a note's
line break as CRLF, the patcher as LF. `info-engine-default-width` — the same default column width written as `defaultColWidth` or as `baseColWidth` +
`defaultColWidth`. `info-engine-dv-ext` — a data validation list on another sheet is written by Excel in the x14 extension, by the patcher in the main part.
`info-engine-font-complete` — Excel completes an incomplete font. `circular-js/xl` — a circular reference that is already in the baseline (the values of that sheet are not judged).
`convert-blank` — converting a formula that has no stored result to values yields a blank (documented in V1).

## 4. Does the oracle see bugs? (negative controls)

`negative.ts` plants ONE fault at a time into a premerged package of four scenarios and requires the same judge to find it (27 / 27, 11 kinds):

| planted fault | found as |
|---|---|
| a number of a copy changed by 7 | `copy-value` |
| a formula of a copy got +1 | `open` / `copy-formula` |
| a copy renamed (formulas elsewhere still point at the old name) | `names` |
| a copy hidden | `copy-visibility` |
| conditional formats / data validation removed, a column 12 characters wider, a column hidden | `copy-format` |
| a chart series reads one row less / a picture removed (scenarios with objects) | `copy-object` |
| a workbook-level name that no base has | `copy-names-leak` |

(The oracle also fails when the judge itself regresses: `judge.test.ts` has 15 tests with planted faults on hand-made dumps.)

## 5. Files

`run.mjs`, `run.ps1`, `oracle.ts` (driver, static checks, summary), `combos.ts` (the 72 combinations), `pipeline.ts` (the app pipeline with copies), `excel.ts` +
`excel-book.ps1` (Excel dumps, the runner), `judge.ts` (the judge), `negative.ts`, `toolcheck.ts` (QA tools over the outputs), `oracle.test.ts`, `judge.test.ts`,
`regress.test.ts`. Reused from the consistency oracle: `ops.ts`, `check.ts`, `srcCheck.ts`, `pkgdiff.ts`, `sets.ts`. Scratch space: the folder given by `-Out`
(about 55 MB) and a per-Excel TEMP below it that is removed after each Excel exits.
