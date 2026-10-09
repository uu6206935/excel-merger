# R3 — Grid / spreadsheet UI for the merged-table editor

Scratch root: `C:\Users\81802\AppData\Local\Temp\claude\D--20260705-----------\2d64adba-cbb0-49f4-bf69-6262ff0972bb\scratchpad\research\r3`
(abbreviated `r3\` below). Date of research: 2026-09-30. Machine: Win11, Node 24.16, Chrome 153 / Edge 153, Excel 16 (M365).

---

## 0. TL;DR

**Build our own grid on `@tanstack/react-virtual` (3.14.x, MIT). Do not use an off-the-shelf grid.**
A working prototype exists in `r3\proto` (React 19 + Vite 6 + vite-plugin-singlefile). It passes all 42 headless-browser checks, including every hard requirement: 10,000 rows × 60 columns, non-contiguous multi-column drag, insertion indicator, auto-scroll, IME, frozen panes and undo/redo. The single-file `dist/index.html` is **318 KB (101 KB gzip)**, and React alone accounts for 218 KB of that.

Why the libraries lose (each point was checked hands-on; details in §2):

* **Dragging non-contiguous selected columns is not supported by any candidate.** Handsontable 18 turns the drag into a range re-selection (verified). Glide moves only the column under the pointer (verified). AG Grid, RevoGrid and react-data-grid only reorder one column at a time.
* **Starting a Japanese IME on a selected cell** (first keystroke, not yet editing) fails in most grids. In our CDP-simulated test the composition was lost in AG Grid, RevoGrid, Glide and default Handsontable. Only three passed both variants: our "always-focused sink textarea" design, Handsontable with `imeFastEdit: true`, and a Process-keydown-driven editor in react-data-grid / jspreadsheet / FortuneSheet.
* **Size, licence and React 19 support also rule some out.** Univer's single file is 11 MB. Handsontable is 1.2–1.9 MB and its licence is non-commercial only. AG Grid is 1.0–1.4 MB. Glide supports React 19 only in an alpha, where the first type-to-edit drops keys (issue #1202).
* We already own the data model: stable column ids, an op log for the Excel engine and our own formula adjustment. A grid that owns its own model (Handsontable or Univer) would duplicate it and fight it.

Use **SheetJS SSF** for display formatting via `import { SSF } from 'xlsx'`; the same package is already in the reference app. Wrap it with the ja-JP shim `numfmt.ts`, which passes 33 cases including 和暦, 曜日, ¥ and the ja-JP built-in ids.

**Important finding for the Excel本体 engine (§7):**
* Replaying column moves in Excel with `Cut()` + `Insert` goes through the Windows clipboard. In our runs it was non-deterministic (another process was using Excel/clipboard concurrently) and once killed Excel ("RPC server is unavailable").
* The clipboard-free replay is 100% deterministic (45/45 over 3 rounds) and produced exactly the same order and formulas. It has three steps: insert blank columns at the target → `src.Cut(Destination)` → delete the emptied source columns.
* Both engines must replay the *same* sequence of contiguous block moves (`planBlockMoves`), because range endpoints depend on that sequence.

---

## 1. Requirements → how the recommended design covers them

| Requirement | Covered by (prototype file) | Verified by |
|---|---|---|
| React 19 + single-file build, no workers/wasm | Vite 6 + `vite-plugin-singlefile` 2.3.3, same config as the reference app (`r3\proto\vite.config.ts`) | `npm run build` → one 318 KB `dist/index.html`; e2e runs it from `file://` |
| 10k rows × 100 cols, virtualised | two `useVirtualizer` instances (rows + horizontal cols) on one scroll element (`src/grid/Grid.tsx`) | about 720–1,070 cell nodes in the DOM at any time; jumping to row 10001 takes about 60 ms |
| Header click / Ctrl (toggle, non-contiguous) / Shift (range) / Ctrl+Shift (add range) | `src/grid/selection.ts` `clickSelect()` | e2e checks 3a–3d |
| Drag selected (non-contiguous) columns anywhere, insertion indicator, auto-scroll near edges | pointer events + pointer capture on the header container, rAF auto-scroll, `gapAt()` binary search, `computeMove()` | e2e: exact final order, indicator x within 2 px, auto-scroll 26→1846 px, drag survives the source header being virtualised away, Esc cancels |
| Correct final order for non-contiguous moves + Excel replay plan | `src/grid/colops.ts` `computeMove`, `planBlockMoves`, `applyBlockMove` | 200,000 randomised cases (`tests/move.test.ts`); replayed in real Excel (§7) |
| Insert/delete columns (context menu, Delete, Ctrl+-) | `Grid.tsx` menu + keydown; `sheet.ts` ops | e2e checks 8 and 9 (a delete that would break a formula asks first: "#REF! になります") |
| Excel styles per cell (fill, font colour/bold/italic/size/underline, borders, h/v align, wrap, indent) | `src/grid/styles.ts` `styleToCss()` cached per style id; border overlay `.g-bd` | screenshots `r3\proto\shots\*.png` |
| Number formats (dates, 和暦, currency, %, colours) | `src/grid/numfmt.ts` (SSF + ja-JP shim) + a 200k-entry cache in `Grid.tsx` | `tests/numfmt.test.ts` 33/33 |
| Inline edit (F2 / double-click / type-to-edit), formula bar, keyboard nav | always-focused sink `<textarea>` (IME-safe), `App.tsx` formula bar | e2e check 11 |
| Japanese IME-safe editing | the sink textarea sits over the active cell and becomes the editor in place, so the first key's composition never changes element; Enter/Tab/arrows are ignored while `isComposing` or keyCode 229 | e2e check 12 (CDP `Input.imeSetComposition` → 日本語 stored, first character not lost) |
| Frozen first columns + sticky header | CSS `position: sticky` (row absolutely positioned; its children sticky) | e2e: the ID header stays at the same pixel after scrolling 2,500 px × 50,000 px |
| Colour strip showing the source file per column | header band runs (contiguous columns from the same file), sticky label with file name and column count | screenshot 1 |
| Ctrl+Z / Ctrl+Y | pure reducer with an op history (`sheet.ts` `reducer`) | e2e (move, delete, resize, cell edit) |
| Accessibility basics | `role=grid/row/columnheader/gridcell`, `aria-rowindex/colindex/selected/multiselectable`, labelled sink, `aria-live` announcement of the active cell | manual review |
| Licence OK | everything MIT / Apache-2.0 | — |

---

## 2. Candidate evaluation (hands-on unless marked "docs")

Single-file sizes come from the same Vite 6 + singlefile config with a minimal 1000×20 example (`r3\eval-sizes`). The React-only baseline is **218 KB**. "IME" is a CDP simulation on a *selected* cell, first keystroke, two variants:
* (a) composition events only;
* (b) preceded by keydown `Process`/229, as real Windows IME does.

✓ means 日本語 was stored. It is a simulation, so treat ✗ as "at risk", not "proven broken".

| Library (version) | License | React 19 | 1-file size (gzip) | Non-contiguous multi-col drag | IME (a)/(b) | Per-cell Excel styles | Verdict |
|---|---|---|---|---|---|---|---|
| **Custom on @tanstack/react-virtual 3.14.13** | MIT | yes (peer ^19) | **318 KB / 101 KB** (whole prototype incl. SSF) | **yes** (implemented) | **✓ / ✓** | yes (DOM + CSS) | **USE** |
| Handsontable 18.1.1 + @handsontable/react-wrapper 18.1.1 | proprietary; free only for non-commercial/evaluation, and must not be used to build "competitive" software | yes | 1,888 KB / 440 KB full; 1,264 KB / 284 KB with minimal modules | **no**: a UI drag on a Ctrl-selection re-selects a contiguous range (screenshot `r3\eval-hot\hot-drag.png`); contiguous drag works; API `moveColumns([1,3,5], 10)` works | ✗/✗ by default; **✓/✓ with `imeFastEdit: true`** | yes (renderers/classes) | reject: licence for a distributed tool, size, no non-contiguous drag. Its formulas plugin (HyperFormula) must not be used: a different engine; docs state IndexMapper moves of formula cells are unsupported |
| AG Grid Community 36.2.0 + ag-grid-react | MIT | yes | 1,359 KB / 387 KB (AllCommunityModule); 1,010 KB / 292 KB minimal | **no**: one header or group at a time; no header multi-select; cell range selection is Enterprise; `moveColumns(keys, toIndex)` is API-only (docs) | ✗/✗ | via cellStyle/cellRenderer | reject |
| Glide Data Grid 6.0.4-alpha24 (6.0.3 latest stable) | MIT | **6.0.3 peer = React ≤18**; alpha24 allows 19 | 571 KB / 186 KB | **no**: multi column selection works (`columnSelect="multi"`), but dragging moves only the grabbed column (`onColumnMoved(6,12)`, verified) | ✗/✗ (composition went to an a11y `<td>`; nothing stored) | canvas: fill/text via theme overrides; borders need custom `drawCell` | reject. Issue #1202 (2026-09-25): on React 19 the first type-to-edit drops keystrokes (React.lazy overlay). Last push 2026-01 |
| RevoGrid 4.28.2 + @revolist/react-datagrid 4.28.0 | MIT core | yes | 582 KB / 174 KB | no evidence: docs show single-column drag via `canMoveColumns` only | ✗/✗ | via cell templates | reject |
| react-data-grid 7.0.0-beta.61 | MIT | yes (peer ^19.2) | 273 KB / 86 KB | no: `onColumnsReorder(sourceKey, targetKey)`, single | ✗ / **✓** | via `renderCell` | reference only: good small codebase for a11y and keyboard patterns |
| TanStack Table 9.2.4 | MIT | yes | not measured (react-virtual alone adds ≈26 KB over the React baseline) | headless (column-order state only) | n/a | n/a | not needed: no sorting, filtering or grouping required; the order array is trivial |
| dnd-kit (@dnd-kit/core 6.3.1 / @dnd-kit/react 0.5.0) | MIT | yes | — | no built-in multi-item drag; sortable items must stay mounted (breaks with virtualisation) | n/a | n/a | reject for the header; optional for small lists elsewhere, e.g. file cards |
| Univer 1.0.3 (@univerjs/presets + preset-sheets-core) | Apache-2.0 (xlsx import/export is **Pro/paid**) | yes | **11,278 KB / 2,829 KB**; about 5 s to first render from `file://` | not verified; found no evidence of non-contiguous column *move* (a 2024-12 PR title shows non-adjacent *delete*) | IME editor receives the composition; storing was not confirmed (canvas) | full Excel style model | reference only: too heavy, duplicates our model, needs its own xlsx↔IWorkbookData converter |
| FortuneSheet 1.0.4 (@fortune-sheet/react) | MIT | yes | 2,970 KB / 659 KB | no column drag-move found in docs or issues (not verified hands-on; Luckysheet lineage) | ✗ / ✓ | yes (canvas) | reject: size, activity (last push 2025-12) |
| jspreadsheet-ce 5.0.4 | MIT | not React-based (wrapper 5.0.2) | 546 KB / 143 KB | single `columnDrag` | ✗ / ✓ | yes | reject: renders every row (22,066 DOM nodes for 1000×20) |
| x-data-spreadsheet 1.1.9 | MIT | not React | — | no | — | yes (canvas) | reject: npm 2022, repo idle since 2024-08 |
| react-spreadsheet 0.10.1 | MIT | yes | — | no | — | no | reject: no virtualisation |

Evidence files:
* `r3\eval-hot\t.mjs`, `hot-drag.png` (Handsontable UI and API move test)
* `r3\eval-glide\t.mjs`, `ime.mjs`, `glide-*.png`
* `r3\eval-sizes\` (all builds): `check.mjs` (every candidate renders from `file://` with React 19 and no console errors), `ime-check.mjs` (IME matrix; `PROCESS_KEY=1` for variant b), `shot-*.png`, `ime-*.png`

---

## 3. The prototype (`r3\proto`)

Run it:

```bat
cd r3\proto
npm install            :: already done
npm run dev            :: http://127.0.0.1:5199 (StrictMode dev build; e2e also passes against it)
npm run build          :: dist\index.html, open directly via file:// works
node tests\move.test.ts      :: 200k randomised move/plan checks (Node 24 strips TS types natively)
node tests\numfmt.test.ts    :: 33 number-format cases + micro-benchmark
node tests\e2e.mjs [url]     :: 42 checks in headless Chrome via playwright-core (channel 'chrome', falls back to 'msedge'; no browser download)
node tests\perf.mjs          :: scroll timing breakdown via CDP Performance.getMetrics
```

Files:

```
src/grid/types.ts      XStyle (flattened ExcelJS style, colours resolved), GridColumn {id,title,width,source,keyRole}, CellView {value, formula?, styleId?}
src/grid/colops.ts     computeMove / isNoopMove / planBlockMoves / applyBlockMove / colLetter / rangeBetween / toRuns   (pure, unit-tested)
src/grid/selection.ts  clickSelect (click / Ctrl / Shift / Ctrl+Shift), orderedSelection
src/grid/styles.ts     styleToCss (cached per style id), resolveExcelColor (argb / theme+tint / indexed), OFFICE_2013_THEME, OFFICE_2023_THEME, applyTint
src/grid/numfmt.ts     formatValue(value, numFmt) -> {text, color, kind}; ja-JP built-in table; 和暦/曜日/¥/[DBNum]/[$-F800] shims on top of SSF
src/grid/Grid.tsx      the component (≈780 lines): virtualisers, sticky layout, header bands, selection, drag/resize/auto-scroll, context menu, IME sink editor, keyboard
src/sheet.ts           demo model: 60 generated columns from 3 "files", style table, formula templates with column-id refs, Op + pure history reducer
src/App.tsx            toolbar (undo/redo/insert/delete/freeze), formula bar, legend, status bar (shows the Excel replay step count)
tests/                 move.test.ts, numfmt.test.ts, e2e.mjs, perf.mjs
shots/                 1-drag-indicator.png, 2-autoscroll.png, 3-inserted.png, 4-ime-composing.png, 5-scrolled-frozen.png
```

### 3.1 Grid contract (what the real app passes in)

```ts
interface GridProps {
  columns: GridColumn[];            // display order; ids are stable across moves
  sources: SourceInfo[];            // {name, color} per input file -> header band colour
  appSourceColor: string;           // band colour for columns added in the app
  rowCount: number;
  frozenCols: number;               // positional, Excel-like (moving a column into the pane freezes it)
  getCell(row, col): CellView | undefined;       // must be cheap; memoised per data version by the caller
  cssOf(styleId): CellCss;                         // makeStyleCache(styles)
  selection: {cols: ReadonlySet<string>, anchor: string | null}; onSelectionChange(sel)
  active: {row, col} | null;        onActiveChange(pos)
  onMoveColumns(ids: string[], gap: number)      // gap = insertion slot 0..n in the CURRENT order
  onResizeColumn(id, width, commit: {from} | null)  // null = live preview (no history), then one commit
  onCommitEdit(pos, text)                          // raw text; the caller parses (keep text for text columns!)
  onDeleteColumns(ids), onInsertColumn(index), onUndo(), onRedo()
  dataVersion: number                              // bump on cell edits so memoised rows and cells refresh
}
```

### 3.2 Layout: one scroller, native scrollbars, CSS sticky

```
.g-scroller (overflow:auto)                          <- the ONLY scroll element (both virtualisers use it)
  .g-canvas (width = 52 + Σwidths, height = header + Σrows, position:relative)
    .g-header  sticky top:0 z4 (+ band runs, header cells absolute at left = item.start)
      .g-pin   sticky left:0 z3 (corner + frozen header cells + their resize handles)
    .g-row × visible (absolute, transform: translateY(item.start))
      cells of visible scrolling columns (absolute, left = item.start)
      .g-pin   sticky left:0 z2 (row number + frozen cells)
    .g-active (frame) + .g-sink (textarea editor)
.g-overlay (sibling of the scroller, pointer-events:none): drop indicator + ghost label
```

Virtualiser setup that works (paddingStart and scrollPaddingStart make `scrollToIndex` respect the sticky header and the frozen pane):

```ts
const rowV = useVirtualizer({ count: rowCount, getScrollElement, estimateSize: () => ROW_H, overscan: 10,
  paddingStart: HEADER_H, scrollPaddingStart: HEADER_H });
const colKey = useCallback((i: number) => `${columns[i]?.id}:${columns[i]?.width}`, [columns]);
const colV = useVirtualizer({ horizontal: true, count: n, getScrollElement,
  estimateSize: (i) => columns[i].width,
  getItemKey: colKey,                       // <- see pitfall P1
  overscan: 3, paddingStart: ROWHDR_W, scrollPaddingStart: ROWHDR_W + frozenW });
const vCols = useMemo(() => colV.getVirtualItems().filter((c) => c.index >= frozen), [...]); // frozen ones are rendered in .g-pin
```

### 3.3 Column drag (pointer based, multi and non-contiguous)

* **pointerdown on the header container** (event delegation via `closest('[data-col]')` / `closest('[data-resize]')`):
  * An unselected column gets `clickSelect` immediately.
  * Pressing on an *already selected* column keeps the multi-selection, so it can be dragged. A plain click without movement collapses the selection to that column on pointerup (Explorer semantics).
  * `headerRef.current.setPointerCapture(e.pointerId)`, on the container, never on the cell (P2).
* **pointermove:** above a 4 px threshold the state becomes `move`, `body.g-dragging` is set (cursor + `user-select:none`) and a rAF loop starts.
* **rAF tick:**
  * Auto-scroll when the pointer is within 56 px of the right edge, or within 56 px right of the frozen-pane edge (speed proportional, max 32 px/frame). The pointer may be outside the window thanks to capture.
  * Recompute `gap = gapAt(lastClientX)`.
  * `setDragUi({gap, x, y, ids})`.
* **gapAt:** binary search over the prefix sums for the first column whose centre is right of x. The frozen pane uses unscrolled x (`gap ∈ [0, frozen]`), the scrolling area uses `x + scrollLeft` (`gap ∈ [frozen, n]`).
* **Indicator x:**
  * `gap <= frozen`: `52 + starts[gap]`
  * otherwise: `max(52 + frozenW, 52 + starts[gap] - scrollLeft)`
  * It is drawn in the overlay, full height, with a cap triangle.
* **Ghost:** "3列を移動: 氏名カナ、学籍番号、区分". Dragged columns are dimmed (opacity .4).
* **pointerup:** `onMoveColumns(ids, gap)`.
  * The parent calls `isNoopMove` first, so dropping inside the selection span does not create an undo entry.
  * pointercancel and Esc cancel the drag; lostpointercapture ends it.
* **Final order** (non-contiguous, relative order kept):

```ts
export function computeMove<T>(order: readonly T[], moving: ReadonlySet<T>, gap: number): T[] {
  const g = Math.max(0, Math.min(order.length, gap));
  const moved: T[] = [], rest: T[] = []; let insertAt = 0;
  for (let i = 0; i < order.length; i++) {
    const id = order[i];
    if (moving.has(id)) moved.push(id); else { rest.push(id); if (i < g) insertAt++; }
  }
  return [...rest.slice(0, insertAt), ...moved, ...rest.slice(insertAt)];
}
// e.g. ABCDEFGHIJ, move {B,D,G} to gap 9 (before J) -> ACEFHIBDGJ ; {B,D,G} gap 4 -> ACBDGEFHIJ
```

* **Excel replay plan** (Excel cannot Cut a multi-area selection): `planBlockMoves(order, moving, gap) -> {final, steps: BlockMove[]}`, where `BlockMove = {from, count, to}` is 0-based and `to` is measured before removal (Excel "Insert Cut Cells" semantics).
  * Every moved run is inserted right before an anchor: the first non-moving column at or after the gap.
  * Runs are processed left to right. A trailing run already in place becomes the anchor, and no-op steps are skipped.
  * Guarantees: `steps.length <= number of runs`, and 0 steps for a no-op. It reproduces `computeMove` exactly (asserted internally and over 200k random cases; average 1.48 steps).
  * The op log stores `{ids, gap, before, after, excel: steps}`.

### 3.4 IME-safe editing ("sink" textarea)

* One uncontrolled `<textarea class="g-sink">` is always mounted inside the canvas at the active cell. It stays focused (`focus({preventScroll:true})`) and invisible (`opacity:0; caret-color:transparent; pointer-events:none`) while not editing.
* **Typing a printable key:** the browser inserts it into the sink, `onInput` switches to `edit={mode:'enter'}` and the sink becomes visible in place.
* **IME:** `onCompositionStart` does the same. Because no element is swapped, the composition that started on the first key continues. The IME candidate window also appears next to the cell, since the caret is there.
* **F2 / double-click:** mode `edit` with the raw value (formula text for formulas); arrows move the caret.
* **Mode `enter`:** arrows commit and move (Excel behaviour).
* **Keydown guard:** `if (e.nativeEvent.isComposing || e.keyCode === 229) return;`. This keeps the Enter that confirms a conversion from committing the cell.
* **Other keys:**
  * Alt+Enter inserts a newline (`setRangeText('\n', …)`).
  * Esc cancels; Enter/Tab (with Shift) commits and moves.
  * While editing, Ctrl+Z is the native textarea undo; otherwise Ctrl+Z/Ctrl+Y/Ctrl+Shift+Z drive the grid history.
* **Keep it uncontrolled:** never set a React `value` on the sink during composition (P6).
* **Formula bar:** a normal controlled `<input>` with the same `isComposing` guard on Enter.

### 3.5 Keyboard map implemented

| Keys | Action |
|---|---|
| Arrows / Tab / Shift+Tab / Enter / Shift+Enter | move the active cell |
| Ctrl+Arrows | jump to the edge |
| Home / End / Ctrl+Home / Ctrl+End / PageUp / PageDown | usual navigation |
| F2 | edit the cell (toggles enter/edit mode while editing) |
| Delete | with a column selection: delete the columns (as specified); otherwise clear the cell |
| Backspace | clear the cell and start editing |
| Ctrl+- | delete the selected columns |
| Esc | clear the column selection, or cancel an edit or drag |
| Ctrl+Z / Ctrl+Y / Ctrl+Shift+Z | undo / redo |

UX note: in Excel, Delete on selected columns only *clears* them. We chose "delete the columns" because the spec asks for it. Consider showing a toast with 「元に戻す」 after a Delete-key column deletion.

### 3.6 Test results (latest build)

`node tests\e2e.mjs` → **42 passed, 0 failed** (headless Chrome 153, `file://`). Excerpt:

```
PASS  virtualised: rendered cells << 600,000  -- 722 cells in DOM
PASS  Ctrl+click non-contiguous selection  -- ["氏名カナ","学籍番号","区分"]
PASS  drop indicator visible at the 科目3|科目4 boundary  -- ind=1117 k4=1119
PASS  non-contiguous move -> exact final order  -- ID,氏名,大学,生年月日,入学日,科目1,科目2,科目3,氏名カナ,学籍番号,区分,科目4,...
PASS  Excel replay plan generated  -- Excel再生: 3手順
PASS  formula endpoint followed the moved column  -- =AVERAGE(F2:AF2) -> =AVERAGE(G2:AF2)
PASS  auto-scroll to the right while dragging  -- 26 -> 1846
PASS  drag survives source header being virtualised away  -- srcRendered=0 headers=25
PASS  delete of referenced column asks for confirmation
PASS  column resize by handle  -- 150 -> 210
PASS  text column keeps leading zeros  -- 0012345
PASS  IME composition switches to edit mode in place  -- editing=1 value=にほんご
PASS  IME commit keeps the whole word (no lost first char)  -- 日本語
PASS  frozen column + sticky header stay in place while scrolled
PASS  drop into frozen pane (gap 1)
```

Against the StrictMode dev server (`node tests/e2e.mjs http://127.0.0.1:5197/`) 41 of 42 passed on the first run. The only failure was that auto-scroll in the slower dev build moved the column to index 27 instead of more than 30 within 1.2 s. The threshold is now more than 20; that check was not re-run against dev.

**Scroll performance** (`tests\perf.mjs`, headless Chrome, *software* rasterisation, so pessimistic; the idle rAF cadence is 16.7 ms):

| Scenario | Median frame | p95 | Script / layout / style per frame |
|---|---|---|---|
| vertical 1 row/frame | 16.7 ms | 17–20 ms | 2–3 / 1 / 0.7 ms |
| vertical 8 rows/frame (fast fling) | 20–23 ms | 26–37 ms | 5–6.5 / 3.5 / 1.5 ms |
| horizontal 120 px/frame | 20–22 ms | 47–69 ms | 3.5–5 / 2.5–3 / 0.5 ms |
| diagonal fling | 30–36 ms | 55–74 ms | 9.5–11.5 / 6 / 2.2 ms |

The remainder of each frame is paint/raster, which a real GPU desktop will do far faster. JS cost is kept low by:
* `memo` Row, plus `memo` Cell (horizontal scroll only renders newly visible cells);
* a formatted-text cache;
* one DOM node per cell (no inner span), plus one only when the cell has borders;
* `contain: size layout style` on cells.

---

## 4. Pitfalls hit (and the fixes in the prototype)

* **P1 — virtualizer ignores width changes.** In virtual-core 3.17.x (installed by react-virtual 3.14.13), `estimateSize` is not part of the measurement memo deps, but `getItemKey` is. After a reorder or resize the old widths stay. Fix: `getItemKey: useCallback(i => id + ':' + width, [columns])`, or call `colV.measure()`.
* **P2 — pointer capture lost during auto-scroll.** Capturing the pointer on the header cell breaks as soon as auto-scroll virtualises that cell away: it unmounts and fires `lostpointercapture`. Capture on the stable header container instead and delegate with `closest()`. The e2e test checks the drag still works after the source header disappears.
* **P3 — text selection and focus theft during drag.** `preventDefault` on `pointerdown` does not stop them in Chromium. Fix:
  * `onMouseDown={e => e.preventDefault()}` on the header;
  * on the body too, except on the scrollbar (`e.target === scroller`) and the sink;
  * `user-select:none` on the grid;
  * `body.g-dragging * { cursor: grabbing; user-select: none }` while dragging.
* **P4 — sticky + virtualisation.** An absolutely positioned element cannot be sticky, but its children can. So rows are `absolute + translateY` and their `.g-pin` child is `sticky; left:0`. The header is `sticky; top:0`, and its own `.g-pin` is `sticky; left:0` with a higher z-index. Sticky inside transformed rows works in Chromium 153.
* **P5 — `overflow:hidden` breaks sticky labels.** It makes an element a scroll container, so a sticky label inside it sticks to that element instead of the scroller: the band label was drawn about 180 px off. Use `overflow: clip`, which clips without creating a scroll container.
* **P6 — React-controlled inputs vs IME.** Never swap the focused element or rewrite `value` while `isComposing`. Leave the sink uncontrolled. Guard Enter/Tab/arrows with `isComposing || keyCode === 229`.
* **P7 — resize handle only ~4 px wide.** Inside a cell with `overflow:hidden`, the part beyond the border was clipped and the next cell painted over the rest. Handles are now separate 8 px elements centred on each boundary, rendered after the header cells (`data-resize=<colId>`).
* **P8 — StrictMode double-invokes reducers and state updaters.** My first undo implementation called `setState` inside a `setPast` updater and would double-apply. Keep history in one pure `useReducer` and copy the edits `Map` on write.
* **P9 — per-frame re-render during drag.** `new Set(dragUi.ids)` memoised on `dragUi` (a new object every frame) re-rendered every visible cell 60× per second. Memoise on the ids array reference instead.
* **P10 — keep cell focus off the grid cells.** Clicking cells must not move DOM focus off the sink. Call `sink.focus({preventScroll:true})` after every pointer interaction; `preventScroll` avoids jumps when the active cell is scrolled away.
* **P11 — frozen overlays need scroll-aware x.** The active-cell frame and editor for a frozen column are positioned at `scrollLeft + x` and use a higher z-index (3) than the pin. The frame re-renders on scroll anyway, since the virtualizer triggers a re-render.
* **P12 — number parsing on commit.** Typing into a text column (学籍番号 "0012345") must stay text. Excel would convert it, and so would `Number()`. `parseInput` keeps text when the previous value was a string. Real app: also respect `@` format and the column type.
* **P13 — `useFlushSync`.** react-virtual defaults to `useFlushSync: true`, which re-renders synchronously on scroll. It was smooth and produced no React 19 warnings in this setup. There is also a new `directDomUpdates` option (3.14) that was not needed here.
* **P14 — Bash heredocs on this machine ate backslashes in TS source.** `'\\'` came out as `'\'`. Use the Write/Edit tools for code containing backslashes.

---

## 5. Number formatting (SSF) — findings

`ssf@0.11.2` on npm (2022) and `XLSX.SSF` from `xlsx-0.20.3.tgz` (cdn.sheetjs.com, the reference app's dependency) behave identically in all 42 raw probes. Use `import { SSF } from 'xlsx'`, which tree-shakes in Vite.

Raw SSF gets right: General (11 significant digits, `0.1+0.2 → 0.3`), `#,##0`, `"¥"#,##0`, `[$¥-411]#,##0`, `\¥#,##0;[Red]\-\¥#,##0`, `#,##0_);[Red](#,##0)`, `0.0%`, `yyyy/m/d`, `yyyy"年"m"月"d"日"`, `m"月"d"日"`, `h:mm:ss`, `[h]:mm`, `0.00E+00`, `# ?/?`, `000-0000`, `"〒"000-0000`, `#,##0"円"`, `@`.

Raw SSF gets wrong (fixed by `numfmt.ts`):

| Problem | Raw SSF output | Fix in wrapper |
|---|---|---|
| 和暦 `[$-411]ggge"年"m"月"d"日"` | `ggg2024年10月1日` | substitute g/gg/ggg (令和/令/R …) and e/ee (era year) per value as quoted literals |
| `ge.m.d` (built-in 27/36/50/57), `yyyy.m.d` | throws | quote `.` in date codes (keep `ss.0`) |
| weekday `aaa` / `aaaa` | `aaa` | 火 / 火曜日 (uses `parse_date_code().q`) |
| bare `¥#,##0` | throws "unrecognized character ¥" | quote bare ¥/￥ outside quotes and brackets |
| `[DBNum1][$-411]General` | `$123` | strip `[DBNumN]` and pure locale tags; `[$¥-411]` → `"¥"` |
| built-in id 14 / 22 | `10/1/24` (en-US) | ja-JP table: 14 = `yyyy/m/d`, 22 = `yyyy/m/d h:mm`, 27–36, 50–58 (identical to ExcelJS's `ja-jp` table) |
| ExcelJS reports id 14 as `mm-dd-yy`, id 22 as `m/d/yy "h":mm` | US style | map back to the ja-JP codes |
| `[$-F800]` (system long date) | — | → `yyyy"年"m"月"d"日"` |
| section colour `[Red]`, `[Blue]`, `[Color10]`, `[赤]` | dropped | returned as `color` for the chosen section |

API: `formatValue(value, numFmt) → { text, color?, kind: 'number'|'text'|'bool'|'error'|'empty' }`.

`kind` drives Excel "General" alignment: numbers right, booleans and errors centre, text left. Cost is 2–10 µs per call, so cache it (the grid keeps a 200k-entry Map keyed by format + type + value).

Not handled, and acceptable for display: conditional sections `[<=9999999]` are passed through to SSF without colour; 元年 (`[$-ja-JP-x-gannen]`); `*` fill characters; `_)` padding width is SSF's.

---

## 6. Styles and colours for the grid (bridge from ExcelJS)

* **Flatten styles into a table.** Convert ExcelJS `cell.style` into `XStyle`: `{fill, color, bold, italic, underline, strike, size(pt), font, hAlign, vAlign, wrap, indent, border{top,right,bottom,left:{style,color}}, numFmt}`. Keep one table indexed by style id, e.g. by hashing the JSON, so `cssOf` stays cached.
* **Colours:** ExcelJS returns `{argb}`, `{theme, tint}` or `{indexed}` unresolved. `resolveExcelColor(c, themePalette)` handles all three, using Excel's HLS tint algorithm.
  * Theme index order is 0 = lt1/bg1, 1 = dk1/tx1, 2 = lt2, 3 = dk2, 4–9 = accent1–6, 10 = hlink, 11 = folHlink.
  * Parse `xl/theme/theme1.xml` for the real palette. The defaults `OFFICE_2013_THEME` and `OFFICE_2023_THEME` are included; newer M365 files use the 2023 theme.
* **Borders:** thin/hair/dotted/dashed → 1 px, medium* → 2 px, thick → 3 px, double → 3 px double. They are drawn by `.g-bd`, absolutely positioned at `left:-1px; top:-1px` so they cover the shared gridline. That is why cells cannot use `contain: paint`.
* **Text layout:** `font-size` in `pt`, so 11pt matches Excel at 100% zoom. Wrap uses `white-space: pre-wrap; word-break: break-all` because Japanese has no spaces. Excel's default vertical alignment is bottom (`align-items: flex-end`).
* **Column width** (Excel characters → px): `px = trunc(((256*w + trunc(128/mdw)) / 256) * mdw)`. mdw is the max digit width of the workbook's default font: 7 px for Calibri 11, and it should be measured with canvas `measureText('0')` for 游ゴシック. Keep the original `w` for export unless the user resized the column.
* **Row heights** come from Excel `ht` in pt × 96/72. Use `estimateSize: i => rowHeights[i] ?? default` and call `rowV.measure()` when they change. Wrapped cells need taller rows or they are clipped; the prototype uses a fixed 22 px.
* **Not rendered (preserved on export only):** conditional formatting, data bars, images, charts, comments. Merged cells are not rendered either. For merges, add header merge spans at least. Also block or warn when a move or delete would split a merge: Excel refuses ("結合セルの一部を変更することはできません").
* **Hidden columns** (Excel `hidden`): render as width 0 plus a small marker. Moves must keep them attached to their neighbours. The prototype does not handle this yet.

---

## 7. Excel本体 engine: replaying grid column ops (verified in real Excel)

Scripts are in `r3\excel-com\`: `make-cases.mts` generates 15 cases from `planBlockMoves`; `verify-blockmoves2.ps1 -Mode nocb|cutinsert -Rounds N`. The COM safety rules were followed: own `New-Object -ComObject Excel.Application`, invisible, no alerts, nothing saved, PID from `GetWindowThreadProcessId`, only our PID is ever stopped.

Results:

1. **`planBlockMoves` steps replayed in Excel reproduce `computeMove` exactly** in all 15 cases (orders like `ACEFHIBDGJ`, `HJABCDEFGI`, `BDFHJACEGI`).
2. **Clipboard-based replay is unreliable.** Each step was `src.Cut(); ws.Columns(to+1).Insert(-4161)`.
   * Round 1 passed. In round 2 some Inserts produced blank columns instead of moving ("ANOMALY … L3hasFormula=False"), results became non-deterministic, and then Excel died (`The RPC server is unavailable. 0x800706BA`).
   * Other Excel automation was running on the machine at the same time, which fits clipboard contention. The user's own copy/paste would cause the same problem.
3. **Clipboard-free replay is deterministic** (45/45 over 3 rounds, 0 anomalies). It gives **identical order and formulas** to the good cut-insert round:

```powershell
# BlockMove {from,count,to}: 0-based, 'to' measured before removal
[void]$ws.Range($ws.Columns.Item($to+1), $ws.Columns.Item($to+$count)).Insert(-4161)      # 1) room at target
$srcStart = if ($to -le $from) { $from + $count } else { $from }                          #    source shifted if room was made left of it
$src = $ws.Range($ws.Columns.Item($srcStart+1), $ws.Columns.Item($srcStart+$count))
$dst = $ws.Range($ws.Columns.Item($to+1),      $ws.Columns.Item($to+$count))
[void]$src.Cut($dst)                                                                       # 2) move, NO clipboard; refs follow
[void]$ws.Range($ws.Columns.Item($srcStart+1), $ws.Columns.Item($srcStart+$count)).Delete(-4159)  # 3) drop emptied source
```

4. **Formula semantics observed** (sentinels `L3 = B3*100+D3*10+J3`, `M3 = SUM(A3:J3)`, `N3 = SUM(C3:F3)` on values 1..10).
   * Single-cell references always follow the moved column: `L3` still evaluates to 250 in every case.
   * **Range endpoints follow their own cells.** Cells moved outside the endpoints leave the range. Examples:
     * moving {B,D,G} to the end turns `SUM(A3:J3)` into `SUM(A3:G3)` = 42;
     * moving {H,J} to the front gives `SUM(C3:J3)` = 37;
     * moving {A,C,E,G,I} to the end gives `SUM(A3:E3)` = 30.
   * So what a non-contiguous move does to ranges depends on how it is decomposed. **The in-browser (ExcelJS) engine must apply the same `BlockMove` sequence step by step**, each with Excel's single-block semantics, to match the Excel engine byte for byte. Persist `excel: BlockMove[]` in the op log (the prototype does).
5. **PowerShell/COM gotchas:**
   * `$xl.CutCopyMode = $false` and `= 0` both throw ("cannot convert … XlCutCopyMode"). Simply do not set it.
   * Wrap *every* cleanup call in `try{}catch{}`. After a crash `$wb.Close()` throws, and an unguarded finally skips the PID kill (this is how one instance was orphaned).
   * Never pipe a COM script's output into `Select-Object -First N`. It stops the child PowerShell before its `finally` runs, which orphaned our instance (pid 2580, stopped by PID afterwards).

---

## 8. Integration advice for the real app

* **Model:** columns = `{id, sourceFile, sourceSheet, sourceCol, title, width, keyRole}`, with the order as an id array. Key cell storage by column id so moves never copy data.
* **Formulas:** tokenise into stable references (column id + row) for display. The prototype renders templates such as `=AVERAGE({c8}{r}:{c31}{r})` with the current letters, and range endpoints follow their columns (same as Excel §7.4). Formulas the user types (`=A2*2`) must be parsed into id references too, otherwise they will not follow later moves. Share this tokeniser with the formula-adjustment topic.
* **Before delete:** compute which formulas reference the columns to be deleted and confirm (the prototype does this). Excel itself would produce `#REF!`, and the requirement says never produce it.
* **Op log** = the history `past[]`. Each op keeps its Excel replay data (`move.excel`, insert index, deleted indexes in descending order). Undo pops ops; export replays `past[]` in order.
* **Scope of `getCell`:** it must be O(1) and allocation-light. Pre-resolve style ids and formula display strings per data version.
* **Key-column badges** (`keyRole`) render "キー" in the header. Use the band colour legend for files; the prototype shows 学生名簿 / 成績一覧 / 連絡先 / アプリで追加.
* **file:// mode:** everything in the grid works without the server (no workers, no fetch). The Excel本体 engine is only available when `server.mjs` is running, so gate that button.
* **Selection for merge-key editing** (詳細画面): the same `clickSelect` works for any list of ids.
* **Not implemented yet (easy on this design):**
  * cell range selection with Shift+Arrows or mouse drag;
  * copy/paste (use a `paste` event on the sink);
  * column autofit on double-clicking the handle (measure visible texts with canvas);
  * row height and wrap;
  * hidden columns;
  * merged-cell rendering;
  * a scrollbar-aware drop-indicator clamp at the far right.

---

## 9. UX ideas from similar open-source tools

* **Mito** (mito-ds/mito, `mitosheet/src/mito/components/taskpanes/Merge/MergeTaskpane.tsx`, AGPL, so ideas only):
  * The merge dialog offers a *join type dropdown with a plain-language subtext per option*: lookup = keep every row of sheet 1 and take the *first* match from sheet 2; left; right; inner; outer; "unique in left"; "unique in right".
  * It suggests key pairs automatically by matching identical header names (`getFirstSuggestedMergeKeys`) and supports multiple key pairs.
  * Copy this for our 詳細画面: default to "lookup" (like VLOOKUP, first match), with Japanese one-line explanations.
  * "unique in left/right" maps directly to our 未一致一覧 check.
* **VisiData** (saulpw/visidata, GPL-3, TUI):
  * Join sheets on the key columns marked in each sheet, with explicit join types (inner/outer/full/diff/append/extend/merge).
  * The `merge` type keeps rows from both sides and highlights differing values. That is a good model for showing conflicts when two files disagree on a non-key column (e.g. different 電話番号 for the same ID).
* **Power Query "Merge queries"** (not OSS, but what Excel users know): a live counter under the key selection, roughly "The selection matches 950 of 1000 rows from the first table" (wording from memory, not re-checked). Show the same kind of match-rate sentence (in Japanese) as soon as keys change.
* **OpenRefine** (BSD-3): column-header drop-down menus that hold every column action, and facets that count unmatched or duplicate values with one-click filtering. This is useful for jumping from a check result (重複ID 3件) straight to those rows in the grid.
* **Small Excel/CSV join tools** for scope reference: cmyh100/ExcelMergeTool (Node, MIT, 2019), StinkCat/VlookupTool, Sensorfactdev/csv-joiner.
  * They follow the same "pick file A, pick file B, pick key, output" wizard, which matches our 「投げる→結合→並べる→出力」 flow.
  * None of them preserves formulas or formatting, which is our differentiator.

---

## 10. Artifacts

* Prototype: `r3\proto\` (src listed in §3; `dist\index.html`; tests; `shots\1-drag-indicator.png`, `2-autoscroll.png`, `3-inserted.png`, `4-ime-composing.png`, `5-scrolled-frozen.png`)
* Library evaluations:
  * `r3\eval-hot\` (Handsontable UI and API column-move test)
  * `r3\eval-glide\` (Glide multi-select, drag and IME test)
  * `r3\eval-sizes\` (single-file size builds for 11 variants, render check, IME matrix)
* SSF probes: `r3\ssf-test\t.mjs`, `t2.mjs`
* Excel COM verification: `r3\excel-com\make-cases.mts`, `cases.json`, `verify-blockmoves.ps1`, `verify-blockmoves2.ps1`, `trace.ps1`
