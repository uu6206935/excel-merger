# Excel 統合ツール — Architecture (contract for all implementers)

Goal: a LOCAL Windows app for non-technical Japanese office users. Drop several Excel files → keys (ID / 氏名 …)
are detected automatically → merge → reorder / insert / delete columns in a grid (Ctrl/Shift multi-select,
drag & drop) with Excel-exact reference adjustment and undo/redo → export .xlsx that keeps formulas, formats,
borders, widths, multiple sheets, charts, pivots … Flow for the user: **ファイルを投げる → 結合 → 列を並べる → 出力**.

Look & feel: copy the reference app `D:\ダウンロード\ExcelMerger\interview-scheduler` (read its
`src/styles.css`, `src/App.tsx`, `src/components/*`, `docs/screenshot*.png`): white cards on #f5f6f8, 1px #d6d6d6
borders, radius 6, accent #1f6feb, header bar with title + status dot + segmented toggle on the right, a row of file
cards on top, a settings bar, the main table on the left and a 280px side panel on the right with a big blue
primary button, 元に戻す/やり直す, Excel出力, a summary box, lists and a collapsible 警告 section.

Research reports (READ the one for your module): `docs/research/R1..R6*.md`. Prototypes they reference live in
`C:\Users\81802\AppData\Local\Temp\claude\D--20260705-----------\2d64adba-cbb0-49f4-bf69-6262ff0972bb\scratchpad\research\r1..r6`.

## 1. Stack

- React 19 + TypeScript + Vite 6 + vite-plugin-singlefile → `dist/index.html` (works from file:// = standalone mode).
- Node server `server/server.mjs` (built-in modules only) serves dist, runs the Excel本体 engine
  (`server/excel/excel-runner.ps1` via PowerShell COM), saves/opens output files, auto-shuts down with the window.
- Launch: `start.bat` (ASCII only!) → `ExcelMerger.exe` (Node SEA) → else `node server\server.mjs` → else open
  `dist\index.html` in Edge `--app` (standalone). UI opens in Edge `--app=http://localhost:PORT/` (localhost, not
  127.0.0.1 — see R5 MOTW finding).
- Libraries: jszip (package I/O), xlsx/SheetJS CE (xls/csv/ods reading + SSF number formatting),
  encoding-japanese (CSV Shift_JIS), fastest-levenshtein (fuzzy), @tanstack/react-virtual (grid),
  esbuild + postject + resedit (exe). Own code: `excelRef.ts` formula engine (R2, verified vs Excel), keys (R6),
  grid (R3), COM runner (R4).
- NOT used (research verdicts): ExcelJS round-trip (damages real workbooks), HyperFormula/IronCalc for rewriting,
  Handsontable/AG Grid/Glide (no non-contiguous multi-column drag, IME issues, size/licence).

## 2. Data flow

```
File bytes ──readBook──▶ SourceBook (sheets, cells, XStyle[], names, features)
                         │ pickDataSheet / detectRegion  (user may change sheet / header row)
                         ▼
             KeySource{book, region} × N ──analyzeSources──▶ KeyAnalysis (pairs base↔other, chips ID/氏名)
                         │ defaultPlan / user edits (詳細画面)            │
                         ▼                                                ▼
                     JoinPlan ─────────────joinSources──────────────▶ JoinResult (rows, reports, checks)
                                                                          │ + import column selection
                                                                          ▼
      base package bytes ──buildPremerged (JS patcher)──▶ premerged .xlsx bytes + PremergeLayout + MergeReport
                                                                          │ readBook (read back = single source of truth)
                                                                          ▼
                                                     createEditor ──▶ EditorSnapshot {doc, log:[]}
                                    grid ops (move/insert/delete/edit/width) ─▶ new snapshots (history)
                                                                          │ log = user ops
                     ┌────────────────────────────────────────────────────┴───────────────────────┐
     JS engine: replayLog(premerged, log)                              Excel本体 engine: server + excel-runner.ps1
     (package patcher, Excel semantics)                                opens premerged in its own hidden Excel,
                                                                       replays the same log, saves
                     └──────────────────────────────▶ output .xlsx ◀───────────────────────────────┘
```

Key decisions:
1. **Premerge is always done by the JS patcher** (both engines start from the same premerged package). It inserts
   the imported columns right after the base table's last column (a *clean* insert: no format inheritance, no CF
   leak), writes imported cells with their SOURCE styles (raw style records merged into the base `styles.xml`,
   deduplicated), header cells, column widths; appended rows (join 'full') are inserted after the last data row with
   Excel insert-row semantics PLUS a range-extension rule (refs whose rows end exactly at the old last data row and
   start at/above the first data row grow over the new rows — so 合計 formulas include them). If the base region is
   an Excel Table (ListObject) and/or has an AutoFilter, it is extended over the imported columns and appended rows
   (unique table column names required; Excel refuses moves into/out of tables otherwise).
2. **The user's edits are an op log** (`LoggedOp[]`): struct ops (excelRef `StructOp` on the merged sheet), cell
   writes, column widths. The doc applies each op in memory for display (Excel semantics via excelRef); the JS engine
   replays it on the package XML; the Excel engine replays it in Excel. **Same log, same semantics.**
3. **Non-contiguous moves** are decomposed with `excelRef.planMove` (contiguous block moves). Only that function
   decides the decomposition (the grid's `computeMove` is for preview only).
4. **Never produce #REF!**: before deleting, `analyzeDelete` dry-runs the op (`adjustFormula(...).createdRefError`)
   over every formula/name/rule; the UI asks 「参照している数式を値に変換して削除 / キャンセル」. Moves never
   create #REF! (excelRef keeps sheet-edge refs intact; documented deviation).
5. **Refuse what Excel refuses** (doc ops return `{ok:false, reason}` in Japanese): moving/deleting part of a merged
   area, moving columns into/out of/partly over a table, changing part of an array formula. (The COM clip-free move
   would otherwise corrupt tables — R2b R11.)
6. Imported formulas (premerge): kept live when every reference is a same-row reference to a column of the same
   source sheet that is also imported (rewritten to the new column/row) or an absolute reference that can be
   rewritten the same way; otherwise replaced by the cached value (counted in MergeReport.formulasToValues).
   Unmatched rows of an imported column get the column's most common data style (borders/number formats stay
   consistent) and no value.
7. Header conflicts: an imported header equal to an existing header becomes `見出し（ファイル名）` (option).
8. Styles for display: `XStyle` (resolved colours). Styles for writing: raw XML records (patcher).

## 3. Modules and owners (one owner per file; do not edit files you do not own — ask via your report instead)

| Module | Files | Owner |
|---|---|---|
| Contract | `src/core/types.ts`, `src/core/contracts.ts`, facades `src/core/*/index.ts`, this doc | lead (read-only for everyone else) |
| Formula engine | `src/core/formula/excelRef.ts` (from R2, keep its API), `src/core/formula/forms.ts` (UI form ⇄ file form, _xlfn future-function list, ANCHORARRAY ⇄ #, [@c] ⇄ Table[[#This Row],[c]]), tests `tests/formula/*` + fixtures | F |
| Reader | `src/core/xlsx/{read,styles,theme,cellText,legacy,zip,xml}.ts`, tests `tests/xlsx/read*.test.ts` | R |
| Patcher | `src/core/xlsx/{premerge,replay}.ts`, `src/core/xlsx/patch/*`, `src/core/export/js.ts`, tests `tests/xlsx/patch*.test.ts`, golden fixtures `tests/fixtures/golden/*` | P |
| Keys | `src/core/keys/*` (except index.ts), tests `tests/keys/*` | K |
| Doc / editor | `src/core/doc/*` (except index.ts), tests `tests/doc/*` | D |
| Grid + display | `src/components/grid/*`, `src/core/display/*`, tests `tests/grid/*` | G |
| App / UI | `src/main.tsx`, `src/App.tsx`, `src/state/*`, `src/components/*` (not grid), `src/styles.css` | U |
| Server / COM / launch / exe | `server/**`, `src/api.ts`, `src/core/export/comJob.ts`, `scripts/*`, `start.bat`, `build.bat`, tests `tests/server/*`, `tests/com/*` | S |
| Samples / E2E / QA tools | `samples/*`, `tools/*`, `e2e/*`, `playwright.config.ts` | T |

## 4. Facade signatures

See the doc comments in `src/core/xlsx/index.ts`, `src/core/keys/index.ts`, `src/core/doc/index.ts`,
`src/core/export/index.ts` and the types in `src/core/contracts.ts`. `src/api.ts` (owner S):

```ts
export interface ServerStatus { app: 'excel-merger'; version: string; excel: { installed: boolean; version?: string }; outputDir: string }
export async function getServerStatus(): Promise<ServerStatus | null>      // null = standalone (file:// or no server)
export async function runExcelJob(premerged: Uint8Array, job: RunnerJob, onProgress?: ProgressFn, signal?: AbortSignal): Promise<ExcelJobResult>  // {bytes, warnings}
export async function saveToOutput(bytes: Uint8Array, fileName: string): Promise<{ path: string; name: string }>
export async function openFile(path: string): Promise<void>                 // opens an output file in Excel
export async function openOutputFolder(): Promise<void>
export function startHeartbeat(): void                                       // SSE /api/events + pagehide beacon /api/bye
```

Server endpoints (all JSON unless noted; Host/Origin guards as in the reference server):
`GET /api/ping` → `{app:'excel-merger', version}` · `GET /api/status` · `POST /api/jobs` (JSON RunnerJob) → `{id}` ·
`PUT /api/jobs/:id/input` (xlsx bytes) → starts · `GET /api/jobs/:id` → `{state:'queued'|'running'|'done'|'error', progress, message, error?}` ·
`GET /api/jobs/:id/result` (bytes) · `DELETE /api/jobs/:id` (cancel; kills only that job's Excel PID) ·
`POST /api/save` (bytes, header `x-file-name`) → `{path,name}` · `POST /api/open` `{path}` (only inside output dir) ·
`POST /api/open-folder` · `GET /api/events` (SSE) · `POST /api/bye`.

## 5. UI specification (owner U, grid owner G)

Layout (min width 1280, like the reference):
- **Header**: 「Excel 統合ツール」 · status (● ローカル実行中 / ● スタンドアロン) · 「Excel本体: 利用可能/なし」 ·
  step indicator on the right as a segmented control: ① ファイル ② 結合キー ③ 列の編集 ④ 出力 (current step filled).
- **File cards row** (`upload-row`, horizontally scrollable): one card per file + a final dashed 「＋ ここにExcelファイルをドロップ
  （複数可）/ ファイルを選択」 card. The whole window accepts drops. Card: file name + row-count badge, ★基準 toggle
  (base file; default = first dropped), クリア; シート selector; 見出し行 selector; tags (形式 xlsx/xls/csv, 値のみ
  for values fidelity); detected keys line 「ID: 学生ID / 氏名: 氏名」; a 5-row mini preview table (like the
  reference's マッチング card); card colour strip = the source colour used in the grid header band.
  Empty state (no files): big dropzone in the main area with a 3-step explanation.
- **Settings bar**: 結合キー chips `[✓ ID] [✓ 氏名]` (always shown first; toggling a chip enables/disables that
  family in every pair), extra chips for other selected keys, 「＋ 他の候補 N件」, 「詳細設定…」 button → modal;
  結合方法 select (基準ファイルの行に合わせる / 他ファイルにしかない行も末尾に追加); live summary text
  「基準 30 名 / 一致 28 / 未一致 2 / 重複 1」; toast area (as in the reference).
- **Key detail modal (詳細画面)**: tabs per other file. Table of candidates: 使う (checkbox) · 優先 (↑↓) · キー ·
  基準の列 · このファイルの列 · 一致率 (bar + %) · 重複 · 空欄 · 一意性 · 判定 (良好/一部一致/…, coloured).
  Composite (氏名＋大学 …) and value-discovered candidates are listed too. Options: 異体字を同一視, 先頭ゼロを無視,
  敬称を除去, 旧姓の括弧を除去, 大学名の表記ゆれを統一; 重複があるとき (先頭の行 / 最後の行 / 結合しない).
  Section 「取り込む列」: checkbox list per file (default all non-key columns). Section 「もしかして（候補）」:
  suggestions with 「この組み合わせで結合」 buttons (→ manualLinks). Footer: OK.
- **Main area before merge**: 結合プレビュー = per-file check table (ファイル · 行数 · キー · 一致 · 未一致 · 重複 · 空欄)
  + list of the columns that will be added. **After merge**: formula bar (address box + fx + text; edits commit on
  Enter) above the grid. Grid: sticky header with column letter, header text, source colour band, キー badge;
  frozen first column(s) as in the file; column select click / Ctrl / Shift; drag selected columns (insertion
  indicator); right-click menu: 左に列を挿入 / 右に列を挿入 / 列を削除 / 列幅を元に戻す; Delete key = delete selected
  columns (confirm dialog if formulas depend on them); F2/double-click/typing edits (IME safe); Ctrl+Z / Ctrl+Y /
  Ctrl+Shift+Z anywhere; hint bar above the grid 「列見出しをクリックで選択（Ctrl・Shiftで複数）→ ドラッグで移動」.
- **Side panel (280px)**: primary big button 「結合する」 (before merge) / 「Excel出力」 (after merge);
  元に戻す / やり直す row; 「結合をやり直す」 small button (after merge, confirms that edits are lost);
  summary box; tabs or sections: 「チェック結果」 (未一致 / 重複 / 要確認 / 数式→値 lists with counts, clicking a row
  scrolls the grid to it) and 「列の一覧」 (vertical list of current columns with source colour; drag to reorder
  = same moveColumns op); collapsible 警告 with fidelity notes (e.g. 「グラフ 2 個: Excel本体モードで確実に保持」).
- **Export dialog**: file name (default `<基準ファイル名>_統合_<yyyyMMdd>.xlsx`), engine radio: 「Excel本体で出力（推奨・最も
  忠実）」 enabled when the server says Excel is installed, else 「内蔵エンジンで出力」; progress; result:
  「保存しました: <path>」 with 「Excelで開く」「フォルダを開く」 (server mode) or a browser download /
  showSaveFilePicker (standalone). If the Excel engine fails, offer the JS engine automatically.
- Japanese UI text everywhere; code comments in English. Keyboard focus rings as in the reference.

### Round 2–4 behaviour (implemented; part of the spec)

- **Files without a usable key** do not block the merge: such a file is 結合対象外 (card badge 「結合対象外（結合キーが見つかりません）」
  + button 「キーを選ぶ…」 that opens the key detail modal on its tab; also listed in the summary, merge preview and 警告). It rejoins
  automatically once a key is selected. If NO file has a key, 「結合する」 is disabled with an explanation.
- **Key detail modal**: per-file tab with the candidate table, a match-rate sentence under it, and the manual picker
  「列を選んで結合キーに追加」 (1–3 columns per side → keys.addManualCandidate; manual candidates carry a 手動 tag and a × button).
- **Title merges** (merged areas entirely above the header row) that block a column move: confirm dialog
  「結合を解除して移動」 → LoggedOp 'unmerge' (centerContinuous + the top-left cell's fill / perimeter borders spread over the
  former area) + the move, in one undo step.
- **Typed input** follows Excel-ja: dates / % / currency set the cell's number format (CellWrite.numFmt), formulas are flagged
  dynamic-array when Excel would spill them (CellWrite.dyn), text that Excel would re-parse gets the quote prefix
  (CellWrite.quote), a text constant containing a line feed turns WrapText on (LF only, not a lone CR), a constant typed into a
  table's totals row becomes the column's label text.
- **Table rules** (Excel-measured): a block inside a table may be reordered anywhere in [c1, c2+1]; a block covering the table
  moves it; moving into/out of/partly over a table is refused; a column inserted into a table gets Excel's new-name rule
  (h1 → h2 when free, else h12, h13 …); typing a table header renames the column and every structured reference.
- **Excel本体 engine fallback**: error codes MOVE_VERIFY_FAILED / EXCEL_UNAVAILABLE / PASSWORD_PROTECTED / OPEN_FAILED /
  RUNNER_MISSING → the UI retries with the built-in engine automatically (api.canFallBackToJs).

- **Round 5 additions**: empty state = collapsible 4-step 「使い方」 panel (data-testids `help-button`, `help-panel`, `help-dialog`,
  `help-close`, `header-add-files`, `file-note-<i>`, `file-numeric-header-<i>`, `export-slow`); shortcuts Ctrl+O (add files), Ctrl+S
  (export), F6 …; two-row headers are joined into the column name and carried through ImportColumnSpec.header; GET /api/status
  also returns `diskLow`, `diskFreeMB`, `diskMessage`, `outputDiskLow` (< 2 GB free on the temp / output drive) and the job upload
  answers HTTP 507 / code DISK_LOW (→ built-in engine fallback); an Excel job is 'done' when the runner's final line arrived and the
  output file was verified, NOT when powershell/Excel have exited (they are shut down in the background; only our own Excel PID,
  identified by PID + start time, is ever killed); typed formulas that would reach outside the workbook (DDE `|`, `[book.xlsx]` refs,
  WEBSERVICE/IMAGE/RTD/CALL/REGISTER/EXEC) are written as TEXT by the Excel engine (warning UNSAFE_FORMULA; `isUnsafeFormula` in
  src/core/export/comJob.ts); new key candidates 'ID（英字の接頭辞を除く）', 'ID（桁数違い・末尾一致）', 'メール', '電話番号',
  '氏名ローマ字'; new CheckItem codes dataBelow, tooManyDuplicates, nameNotSplit.

Round 6: header chip `disk-low-chip` and export-dialog notice (`export-disk-notice`, `export-disk-temp`, `export-disk-advice`, `export-disk-output`) when /api/status says diskLow / outputDiskLow (default engine then = built-in); `scripts/make-release.mjs` (`npm run release`) builds release\ExcelMerger\ + ExcelMerger_<yyyyMMdd>.zip (exe, start.bat, README, はじめにお読みください.txt, samples); the reader folds sheetFormatPr baseColWidth into SheetGrid.defaultColWidth (Mac Excel).

- **Round 7 (user request)**: (1) *Single-sheet output*: PremergeOptions.dropOtherSheets (UI default ON, checkbox 「他のシートも出力に含める」
  in the settings bar turns it off) → the output workbook holds ONLY the merged sheet; see the contract comment in contracts.ts for the
  exact rules (formulas reading removed sheets → cached values, names/DV/CF/hyperlinks/charts cleaned, no #REF!).
  (2) *File cards*: drag a card to reorder files (before merge: reorders the sources = the order of the imported column groups; after
  the merge: moves that file's column group in the grid); click a card (not its buttons) → the grid scrolls so the file's columns
  are centred and ALL columns of that file are selected. (3) *Source band in the grid header*: a run of columns of the same file is
  one band; click = select the whole run; drag = move the whole run (same insertion indicator / auto-scroll as column drag).
  Logic is based on column origin (DocColumn.origin), not on colours.

- **Round 8 (user request): imported formulas must stay formulas.** Today only same-row refs to imported columns survive; everything else
  (and every file whose formulas have no cached result) ends up as values / EMPTY cells. Required behaviour of buildPremerged for the
  formulas of imported columns (2nd and later files): (1) shared-formula dependents are expanded and kept; (2) row mapping: a ref to
  source row r is rewritten to the merged row T(r) of that source row (join plan: matched rows + rows appended by join 'full'), so
  reordering by key never breaks a formula; ranges covering the whole data table (D2:D31, D:D, D$1:D$31) become the whole merged data
  table of the mapped column; partial ranges only when their rows are contiguous and in order in the merged sheet; (3) refs to
  non-imported key columns read the base's partner column (JoinResult.keyPairs) when the texts agree on all matched rows; (4)
  defined names and table structured refs of the source are resolved (via SourceBook.names / SheetFeatures.tables) and inlined as A1
  refs; (5) everything that cannot be kept is imported as the cached value and counted per reason in MergeReport.formulaStats; a
  formula without cached result that cannot be kept is the only way to get an empty cell and is reported (code formulasNoCache).
  After the premerge the merged sheet is an ordinary workbook: column moves / inserts / deletes adjust these formulas with Excel
  semantics (doc + both engines); verified by a real-Excel value-invariance oracle (tests/oracle).
- **Round 9 (user request, "plan A")**: PremergeOptions.includeSourceSheets copies EVERY sheet of every other (2nd, 3rd …) file into the
  output after the merged sheet and the base's own other sheets (UI: second checkbox 「他のファイルのシートも含める」 under 「他のシートも
  出力に含める」, default off). Names: planSheetNames (src/core/xlsx/sheetnames.ts) = "<file name without extension>_<sheet name>", at most
  31 characters (the file part is shortened first), forbidden characters → "_", and "(1)", "(2)" … appended to the sheet added LATER when a
  name is taken (case-insensitive, full/half width folded); the merged sheet and the base's sheets are never renamed. Copies are faithful
  (cells, formulas with renamed sheet refs, styles / dxfs / shared strings re-indexed, names scoped to the copy, CF / DV / merges / cols /
  hyperlinks, then drawings / charts / images / tables / comments / pivots); what cannot be copied is left out with a 'sheetCopyLimits'
  warning naming the sheet. Imported formulas (round 8) that read other sheets of their source, or cells the merged table cannot
  represent, read the COPIES instead and therefore stay formulas. PremergeLayout.addedSheets / MergeReport.addedSheets list the result.

Additional data-testids: `file-excluded-<i>`, `excluded-notice`, `excluded-pick-key-<sourceId>`, `preview-excluded-<sourceId>`,
`side-excluded`, `key-detail-tab-<sourceId>`, `detail-excluded`, `key-candidates`, `key-row-<candidateId>`,
`key-remove-<candidateId>`, `match-rate`, `match-rate-text`, `match-rate-empty`, `manual-key-picker`, `manual-base-select-<n>`,
`manual-other-select-<n>`, `manual-base-add-column`, `manual-other-add-column`, `manual-base-remove-<n>`,
`manual-other-remove-<n>`, `manual-key-hint`, `manual-key-add`, `manual-key-message`. Source file ids are s1, s2 … in drop order.

### data-testid contract (UI owner U and grid owner G implement; E2E owner T uses)

`file-input` (hidden `<input type=file multiple>`), `drop-zone`, `file-card-<i>`, `base-toggle-<i>`,
`sheet-select-<i>`, `header-row-select-<i>`, `file-clear-<i>`, `key-chip-id`, `key-chip-name`, `key-detail-button`,
`key-detail-modal`, `key-detail-ok`, `join-type-select`, `merge-button`, `remerge-button`, `export-button`,
`undo-button`, `redo-button`, `summary-line`, `check-panel`, `check-unmatched-count`, `check-duplicate-count`,
`column-list`, `grid`, `grid-header-<i>` (0-based current position), `grid-cell-<row>-<i>` (row = Excel row),
`formula-bar-input`, `context-menu`, `export-dialog`, `export-engine-excel`, `export-engine-js`, `export-filename`,
`export-confirm`, `export-result`, `delete-confirm-dialog`, `delete-convert-button`, `delete-cancel-button`, `toast`,
`server-status`.

## 6. Invariants / tests everyone can rely on

- `tokenize(f).map(t => t.text).join('') === f` for any f; only reference tokens are ever rewritten.
- For every op sequence: `readBook(replayLog(premerged, log))` equals `doc` after the same ops for all cell
  values/formulas/styles of the merged sheet and every formula on other sheets (tests/doc/invariant.test.ts).
- For golden cases: `replayLog` output opened in Excel equals Excel's own result of the same ops (tools/inspect).
- The Excel runner and `replayLog` produce the same formulas for the same log (tests/com, needs Excel).
- Output files open in Excel **without repair** and contain no new `#REF!`.

## 7. Coding rules

- TypeScript strict; no `any` leaks across module boundaries. Pure functions in `src/core/**` (no React, no DOM,
  no Node APIs) so they run in vitest (environment node) and in the browser.
- Write code files with the Write/Edit tools, not shell heredocs (backslashes get eaten — R3/R6 pitfall);
  re-check regexes (`\d`, `\s`) after automated edits.
- PowerShell scripts: UTF-8 **with BOM**, ASCII-only stdout (JSON lines). .bat files: ASCII only, CRLF.
- Excel COM safety: own instance only; never kill EXCEL.EXE by name; kill only your own PID; Visible=false,
  DisplayAlerts=false; save/restore the user's clipboard around Cut operations.
