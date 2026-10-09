# tools/ — samples and QA tools (owner: T)

All Excel automation runs in **our own hidden Excel instance** (`tools/lib/xlcom.ps1`): its PID is announced
on stdout and only that PID is ever killed (by the Node watchdog in `tools/lib/ps.mjs` when a script hangs).
Excel is never killed by name. Every input is copied to an ASCII temp path first. PowerShell scripts are UTF-8
with BOM and print ASCII-only JSON lines.

Every Excel start writes ~80 MB of Office diagnostics into `%TEMP%\Diagnostics`: on a small system drive run the tools, vitest and
playwright with `TEMP` / `TMP` pointing to another drive. `QA_NO_EXCEL=1` makes `excelAvailable()` false, so the E2E specs and
`tests/qa/excel.test.ts` skip every step that needs Excel (the UI parts still run). The E2E helpers inspect the base and all outputs
of a flow in ONE Excel instance (`verifyAndProbe` in `e2e/helpers.ts`).

| Command | What it does |
|---|---|
| `npm run samples` | Regenerates `samples/` (see `samples/README.md`): data from `lib/sample-data.mjs` (deterministic), workbooks built by `make-samples.ps1` in Excel, CSV (Shift_JIS) and `truth.json` in Node. `--no-excel` = CSV + truth only. |
| `node tools/inspect.mjs <file...> [--json <dir>] [--recalc]` | What Excel sees: open mode (`normal` / `repair` / `extract` / `failed`, repair log), names, and per sheet every formula (`Formula2`) + value + error, CF, DV, tables, charts (series formulas), pivots (source, fields), merges, widths, hidden columns, freeze panes, autofilter, style probes. |
| `node tools/check-output.mjs <out.xlsx> --base <base.xlsx> [--json] [--self-check]` | Verdict for an exported workbook (exit 0 = PASS, 1 = FAIL, 2 = could not run). `--self-check` reads both files a second time in another Excel instance and requires the two readings to be identical (exit 2 otherwise). |
| `node tools/compare.mjs <a> <b> [--excel-only\|--package-only] [--no-values] [--no-recalc] [--pivot-layout] [--all] [--json]` | "Same op log, two engines": diff of two workbooks (or two `inspect --json` dumps) that reports only REAL differences — see below. Exit 0 = identical, 1 = differences, 2 = could not run. |

## compare: two views, real differences only

- **Excel view** (needs Excel): both files opened and recalculated in our own hidden Excel (`inspect.ps1`) —
  formulas (`Formula2`), values, arrays/spills, merges, CF, DV, tables, charts, pivots, widths, freeze, filter,
  style probes, objects.
- **Package view** (pure JS, `lib/package-diff.mjs`): both packages parsed into a normalized model — formulas
  (shared formulas expanded, array / dynamic-array kind), constants (strings as text), **resolved** cell styles
  (number-format code, font, fill, border, alignment, protection, quotePrefix, cell style name), columns (per column:
  width, hidden, outline, style), rows, merges, pane, autofilter, CF (rules in priority order with resolved DXF), DV,
  sparklines, hyperlinks, tables, drawings/charts (anchor cell + series formulas), comments, pivots (source, fields),
  protection, page breaks, defined names, sheets, special parts (VBA, external links, slicers …).

The first line says `N differences — formula differences: F`; `--json` gives `{identical, formulaDiffs, excel, package}`.
Each view prints a table of the differing rows only (none when identical), then **Filtered (known-benign, not
counted)** with a count and examples per category:

| key | filtered because |
|---|---|
| `calcChain` | `xl/calcChain.xml` present / different (Excel rebuilds it; the JS engine drops it) |
| `docProps` | `docProps/core.xml` / `app.xml` (author, dates, application) |
| `xfIndex` | cell `s=` differs but the resolved style is the same |
| `sstIndex` | sharedStrings index / order differs, same text |
| `sharedFormula` | shared formula vs individual formulas with the same expanded text |
| `cachedValues` | cached `<v>` of formula cells (recalculated values are compared in the Excel view) |
| `spill` | dynamic-array spill extent and the spilled values (decided by Excel's recalc; compared as `spills` in the Excel view) |
| `emptyCells` | an empty `<c/>` present in one file only, same effective style |
| `dimension` | `<dimension ref>` |
| `rowAttrs` | automatic row heights (no customHeight), spans, dyDescent |
| `colSplit` | how `<col>` ranges are split/merged, customWidth/bestFit flags, default-width columns |
| `selection` | selection, active cell, scroll position, tabSelected |
| `cfPriority` | CF priority numbers / dxfId / element order when the ordered rules and their formats are equal |
| `workbookMisc` | calcPr (fullCalcOnLoad …), fileVersion, bookViews, revision pointers, pivot cacheId |
| `filterDatabase` | the `_xlnm._FilterDatabase` name (compared through the autofilter) |
| `pivotLayout` | pivot table location (Excel re-lays it out on refresh; `--pivot-layout` compares it) |
| `anchorOffsets` | colOff/rowOff of drawing anchors (the anchor cells are compared) |
| `xmlNoise` (always) | attribute order, namespace declarations/prefixes, mc:Ignorable, xr:uid revision ids, whitespace |
| `styleNoise` (always) | font family/charset/scheme, xf applyXxx flags, unused styles |
| `otherParts` (always) | part names / rIds, theme, printer settings, VML, pivot cache records, chart value caches |
| Excel view: volatile values | values of NOW / TODAY / RAND / RANDBETWEEN / RANDARRAY / CELL / INFO and their spills |
| Excel view: pivot layout | widths / used range / probes of sheets holding a pivot table (`--pivot-layout` compares them) |

Everything else is a real difference. Unit tests: `tests/qa/package-diff.test.ts`, `tests/qa/compare-core.test.ts`.

## check-output verdict

Both files are opened (recalculated) in our Excel. **FAIL** when:

- the output needs repair or cannot be opened;
- there are more `#REF!` references (formula text, names, CF, DV, chart series) or more `#REF!` / `#NAME?` values than in the base;
- a sheet, defined name, chart, pivot table or table is gone;
- a formula (other sheets by address, the merged sheet by following its header) now references a column with a
  **different header text** than in the base (single-column references; multi-column ranges only warn);
- a CF / DV rule of the merged sheet no longer covers the column with its original header (same kind of rule, same
  formula); a pivot source lost a header column.
  A pivot source / CF / DV that only lost columns the user DELETED (header gone from the merged sheet) is INFO
  (`pivot-source-deleted`, `cf-deleted`, `dv-deleted`, `cf-shrunk`, `dv-shrunk`) — Excel does the same when a column is deleted.
  A CF / DV range that GREW over a neighbouring column (insert right of a column carrying the rule extends it: verified
  Excel behaviour, both engines do it; also columns moved into the range) is INFO (`cf-grown`, `dv-grown`,
  `stats.ruleGrowth = {cf, dv}`) — it only fails when the rule no longer covers the original header column.

WARN: other error values increased, sheet order, freeze / autofilter / hidden columns lost, values of formulas on
other sheets changed (normal after appending rows), bounded ranges over the table that did NOT grow over appended
rows (`appended-ranges-excluded`). INFO: formulas converted to values, added / removed columns, rows appended at the
end of the table (join 'full': `appended-rows`) and which formulas / names / pivot sources / CF / DV / autofilter
ranges grew over them (`appended-ranges`, `stats.appended = {rows, ranges, extended, notExtended}`).

The header row of each sheet is guessed from the base (row with the most text cells among the first 40 rows);
`--sheet` / `--header-row` override it for the merged sheet.

## Why a verdict can trust Excel's reading (the round-5 flake)

Once a built-in-engine output was flagged FAIL (defined name `大学一覧` missing, all 190 checked formulas "pointing at other
headers", chart series / CF / DV not on their headers) although the package was fine and the same file passed alone and in the
next batch. The cause is on the Excel side of the inspection: Excel answers COM calls while it is still loading / calculating /
competing with another instance, and `inspect.ps1` swallowed the rejected calls (`try … catch {}` → "nothing there"): no names,
no cell values (the header row is read from them), no CF / DV. Three layers now remove it:

1. **`inspect.ps1` waits and retries.** It waits for `Application.Ready` + `CalculationState = xlDone` (twice in a row) before
   opening, after opening and after `CalculateFull` (which returns while multi-threaded calculation may still run); opens every
   file once, read-only; reads the name table until the number of names equals `Names.Count`; retries `RPC_E_CALL_REJECTED` /
   `SERVERCALL_RETRYLATER` / `VBA_E_IGNORE` (also when wrapped in PowerShell's exceptions) for every read, a whole sheet included;
   and COUNTS what never got through (`transientFailures`, `notReady`, `wrongWorkbook`, `nameCountMismatch` in the dump).
2. **`lib/dump-check.mjs` cross-checks every dump with its own package** (xlsx / xlsm): defined names, every formula cell and the
   ranges its references point to, the string constants (header rows), CF / DV / chart / pivot / table presence, plus the counters
   above. `inspectFiles` inspects a disagreeing file again in a FRESH Excel instance (3 attempts in all; `attempts`, and
   `suspicious` when the last one still disagreed). The verdict turns a dump that stays inconsistent into `inspect-mismatch`
   (FAIL, listed in `MUST_NOT_FAIL` of the E2E helpers) instead of a bogus file problem.
3. **Self-check.** `check-output --self-check` (two separate readings must be identical); `tests/qa/dump-check.test.ts` (pure:
   dumps built from the real sample packages and damaged the way the flake looked; `QA_EXCEL=1`: real dumps are consistent on the
   first attempt and identical across instances, and a fault injected into the first attempt — `EM_INSPECT_FAULT_ONCE=<marker file>`
   makes the first dump come back without names / values — is healed by the retry: `attempts === 2`).

## Library

- `lib/sample-data.mjs` — sample data + ground truth (`buildSpec`, `buildTruth`).
- `lib/inspect.mjs` — `inspectFiles(files, {recalc, probeCells, maxAttempts, noRetry, noCrossCheck})` → dumps (`lib/inspect.d.mts`
  has the JSON shape); `probeCells: ['名簿!A1', …]` adds `cellProbes` (number format, displayed text, value, Formula2, alignment,
  bold, size, font colour, fill, merge area, spill range, prefix) — used by the E2E specs.
- `lib/dump-check.mjs` — `packageFacts(file)`, `dumpIssues(facts, dump)`, `dumpDifferences(a, b)` (see above).
- `lib/package-diff.mjs` — `comparePackages(a, b)` → the package view of compare.mjs (`FILTER_LABELS`).
- `lib/compare-core.mjs` — `extractRefs`, header helpers, `compareInspections`.
- `lib/verdict.mjs` — `checkOutput(baseDump, outDump)`.
- `lib/ps.mjs` — `runPs(script, args, {timeoutMs})` with the PID-scoped watchdog.

Tests: `npx vitest run tests/qa` (pure + samples structure); `QA_EXCEL=1 npx vitest run tests/qa/excel.test.ts tests/qa/dump-check.test.ts`
drives real Excel (variants built by `tests/qa/fixtures/make-variants.ps1`: PASS / #REF! / #NAME? / wrong column /
repair / unreadable; the dump-check self-checks).

## E2E (`e2e/`, `playwright.config.ts`)

`E2E_BUILD=1 npx playwright test` (server mode: builds the UI into `test-results/e2e-dist`, starts `server/server.mjs` on :5199 and a
second server on :5176 whose "Excel" is `e2e/fixtures/fake-excel-runner.mjs`) or, standalone, `npx vite --config e2e/vite.e2e.config.ts`
+ `E2E_BASE_URL=http://127.0.0.1:5176 npx playwright test`. Set `TEMP`/`TMP` to a big drive (every Excel start writes ~80 MB) and
`QA_NO_EXCEL=1` to skip the steps that need Excel (check-output, cell probes; the server's Excel engine is unaffected).
Specs wait for STATE (`waitMergeReady`, `waitIdle`, `expect.poll`), never for time. Files: `01`–`05` smoke / keys / edit+export / grid /
round 2; `06` round 3 (non-contiguous moves, base = 02, **結合対象外 + the manual key picker**); `07` hostile inputs (cut-off, password-protected,
empty, not-a-workbook, .txt / lock files, duplicates); `08` three files in three orders checked row by row against `samples/truth.json`,
export twice, undo all the way back = the untouched merge, the 警告 section, Ctrl+Z / Ctrl+Y / Ctrl+Shift+Z; `09` the Excel engine fails
(every runner error code, a crash, no output, 中止) → the built-in engine takes over; `10` clumsy users (double clicks, odd export names, removing
the base file); `11` drops on the window (synthesised DragEvents) and Ctrl+S / Ctrl+O / F6. Every run writes into its own `test-results/e2e-run-<id>`
(and builds into `e2e-dist-<id>`): a second `playwright test` in the same repo clears ITS outputDir at startup, which used to delete the trace files
of a test running here. A dump that needed a second Excel start is noted in the test annotations and in `test-results/e2e-inspect-retries.log`. Helpers: `helpers.ts` (contract, export, verification),
`round5.ts` (hostile-file builders, `alignmentProblems`, `verifyCases` = many outputs in ONE Excel start, `setFakeExcel`).

## Consistency oracle (`tests/oracle/consistency/`, task V1)

"Formulas stay correct when columns are moved / inserted / deleted, in BOTH export engines, on top of merges whose rows were rearranged by key
matching" — proven in real Excel, re-runnable. `powershell -NoProfile -ExecutionPolicy Bypass -File tests\oracle\consistency\run.ps1 [-Count 340]
[-Out <dir>] [-Only <regex>] [-NoRunner] [-NoExcel] [-Reuse]`; vitest: `tests/oracle/consistency/oracle.test.ts` (Node half always, the whole oracle
with `EXCEL_TESTS=1`), `regress.test.ts` (regressions found by it). It uses `tools/lib/package-diff.mjs` (`comparePackages`) for the package view of the
two engines' outputs; results and failure classes: `tests/oracle/consistency/RESULTS.md`. Helpers: `inspect.mjs` (cells of a dump), `show.mjs`
(one scenario: ops, runner answer, findings), `mutate.mjs` (does the oracle see a deliberately planted bug? — run on a copy of the repository).
Every Excel start gets its own `TEMP` under the output folder and removes it afterwards (~80 MB of Office logs per start).
