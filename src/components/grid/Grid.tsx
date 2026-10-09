/**
 * Excel-like virtualised grid for the merged sheet (owner: grid G). Built on @tanstack/react-virtual (see
 * docs/research/R3-grid-ui.md). The grid is CONTROLLED: it never changes data itself, it reports user intent
 * through callbacks and redraws from props.
 *
 * ── Props (GridProps) ──────────────────────────────────────────────────────────────────────────────────────
 *  columns        GridColumn[] in display order: {id (stable), title, width px, source (index into `sources`,
 *                 -1 = added in the app), keyRole ('ID' / '氏名' → key badge), hidden}. Column i is Excel column
 *                 letter colLetter(i) ("A" for 0).
 *  sources        SourceInfo[] {name, color}: header colour band + tooltip per input file.
 *  appSourceColor band colour of columns added in the app (default #9ca3af); appSourceName its label.
 *  rows           Excel row numbers shown, top to bottom (normally firstDataRow..lastDataRow; hidden rows may be
 *                 left out). The sheet's header row is NOT a grid row: its text is the column title. MEMOISE it.
 *  headerRow      Excel row number of the header row (shown in the top-left corner), default 1.
 *  frozenCols     number of frozen (always visible) columns on the left, positional like Excel's freeze panes.
 *  getCell        (row = EXCEL row, col) → CellView | undefined. Must be O(1); build CellView with cellViewOf().
 *  cssOf          styleId → CellCss (core/display/styles makeStyleCache(doc.styles)).
 *  dataVersion    bump whenever cell contents change (memoised rows/cells redraw on it).
 *  rowHeight      optional (row) → px (e.g. ptToPx(rowProps.height)); default 22 px. MEMOISE it.
 *  selection      {cols: Set<id>, anchor} column (header) selection; onSelectionChange(sel).
 *  active         {row: Excel row, col: id} | null; onActiveChange(pos).
 *  highlightRows  Set of Excel rows tinted amber (e.g. 未一致 rows); highlightLabel = tooltip on their row number.
 *  readOnly       no in-cell editing / Delete clearing (column operations still work).
 *
 * ── Callbacks (user intent; the caller runs the doc ops and passes new props) ─────────────────────────────
 *  onMoveColumns(ids, gap)      ids in current order (non-contiguous ok); gap = insertion slot 0..n in the
 *                               CURRENT order (pre-move coordinates, = doc.moveColumns(s, ids, gap)). Never
 *                               called for no-op drops.
 *  onResizeColumn(id, px, commit)  commit=null: live preview while dragging (no history). Every resize gesture
 *                               that sent previews ends with exactly ONE call with commit={from: startPx};
 *                               px === from means cancelled (drop the preview, no history). Double-click on
 *                               the handle = autofit: a single call with commit={from}.
 *  onCommitEdit(pos, text)      raw text typed in the cell / formula bar ('' = clear the cell). The caller parses
 *                               it (doc.setCellInput). Not called when an F2 edit ends unchanged.
 *  onDeleteColumns(ids)         Delete key / Ctrl+- / menu 「列を削除」 with the selected columns (current order).
 *                               The caller confirms (analyzeDelete) and deletes.
 *  onInsertColumns(index, count) 「左に列を挿入」/「右に列を挿入」/ Ctrl++: insert `count` empty columns before
 *                               0-based `index` (Excel: as many as the contiguous selected columns).
 *  onResetColumnWidths(ids)     「列幅を元に戻す」 (menu item disabled when omitted).
 *  onUndo() / onRedo()          Ctrl+Z / Ctrl+Y / Ctrl+Shift+Z while the grid has focus and no cell is being
 *                               edited (inside an edit they are the text box's own undo).
 *  onEditingChange(state|null)  in-cell editing text, for the formula bar mirror.
 *  onNotice(text)               a short remark for the user (e.g. a paste of several cells only pastes the first one).
 *
 * ── Imperative handle (ref: GridHandle) ─────────────────────────────────────────────────────────────────────
 *  scrollToRow(row, {select}) · scrollToColumn(id, {select}) · scrollToCell(pos) · scrollToColumnsCentered(ids, {select})
 *  · focus() · isEditing() · commitEdit()
 *
 * ── Keyboard (focus lives in an invisible, always-focused <textarea> "sink": IME safe, R3 §3.4) ─────────────
 *  arrows / Ctrl+arrows / Tab / Enter / Home / End / Ctrl+Home / Ctrl+End / PageUp / PageDown move the active cell;
 *  typing or IME composition starts editing in place; F2 / double-click edit the existing content; Esc cancels;
 *  Enter / Tab / (in enter-mode) arrows commit and move; Alt+Enter = new line; Ctrl+Enter = commit in place;
 *  Delete = delete the selected columns, or clear the active cell; Backspace = clear and edit;
 *  Ctrl+Space = select the active column; Shift+←/→ extend a column selection; Ctrl+- delete / Ctrl++ insert
 *  columns; Ctrl+C copies the active cell's displayed text; Shift+F10 / menu key = column menu.
 *
 * ── Mouse ───────────────────────────────────────────────────────────────────────────────────────────────────
 *  header click / Ctrl+click (toggle, non-contiguous) / Shift+click (range) / Ctrl+Shift+click (add range);
 *  drag selected columns (insertion indicator, auto horizontal scroll near the edges, Esc cancels);
 *  drag the header boundary to resize, double-click it to autofit; right-click = column menu.
 *  Source band (the coloured strip above the headers; one band = a RUN of contiguous columns of the same source, counted
 *  over ALL columns, not only the visible ones): click = select all columns of the run (anchor = its first column);
 *  Ctrl+click toggles the run, Shift+click extends like a header Shift+click; drag = move the whole run with exactly
 *  the same gesture as a header drag (ghost, indicator, auto-scroll, Esc, drop inside the run = no-op, pointer captured
 *  on the header container); right-click = column menu of the run. data-testid grid-band-<run index>.
 *
 * ── data-testid ─────────────────────────────────────────────────────────────────────────────────────────────
 *  grid (the scroll container, role=grid), grid-header-<i>, grid-cell-<ExcelRow>-<i>, grid-rowhdr-<ExcelRow>,
 *  grid-editor (the sink), grid-drop-indicator, context-menu (+ context-menu-item-insert-left / insert-right /
 *  delete / reset-width).
 *
 * Layout (one scroll container, native scrollbars, CSS sticky for the frozen parts):
 *   .g-scroller (overflow:auto)                   <- the ONLY scroll element (both virtualisers use it)
 *     .g-canvas (width = gutter + Σwidths (+ endSlack), height = header + Σrows)
 *       .g-header  sticky top            <- source bands + header cells (absolute) ; .g-pin sticky left (corner + frozen)
 *       .g-row × visible (absolute, translateY) <- cells of visible columns ; .g-pin sticky left (row number + frozen cells)
 *       .g-active + .g-sink (textarea)   <- active-cell frame + IME-safe editor
 *   .g-overlay (sibling, pointer-events:none)     <- drop indicator + drag ghost
 */
import {
  forwardRef,
  memo,
  useCallback,
  useEffect,
  useImperativeHandle,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
  type CSSProperties,
} from 'react';
import { createPortal } from 'react-dom';
import { useVirtualizer, type VirtualItem } from '@tanstack/react-virtual';
import type { CellCss } from '../../core/display/styles';
import { autoFitWidth } from './autofit';
import { cellCopyText, cellEditText, isBlankCell, renderCell } from './cellview';
import { colLetter, isNoopMove, toRuns } from './colops';
import { ContextMenu, type MenuEntry } from './ContextMenu';
import { centerScrollLeft, gapAt, insertTarget, isNavKey, makeRowIndex, navigate, prefixSums } from './geometry';
import { bandPieces, computeRuns, runFullySelected, type BandPiece } from './runs';
import { clickSelect, extendTo, orderedSelection, selectOnly, selectRun, type ColSelection } from './selection';
import type { CellPos, CellView, EditingState, GridColumn, GridHandle, SourceInfo } from './types';
import './grid.css';

export type { ColSelection };

export const ROW_H = 22;
export const ROWHDR_W = 56;
const BAND_H = 18;
const TITLE_H = 38;
export const HEADER_H = BAND_H + TITLE_H;
const DRAG_THRESHOLD = 4;
const EDGE = 56; // auto-scroll zone (px from the visible edge)
const MAX_SPEED = 32; // px per animation frame
export const MIN_COL_W = 12;
const MAX_COL_W = 1200;
const DBLCLICK_MS = 450;
const DEFAULT_APP_COLOR = '#9ca3af';
const HEADER_FONT = "700 12px 'Yu Gothic UI', 'Meiryo UI', sans-serif";

export interface GridProps {
  columns: readonly GridColumn[];
  sources: readonly SourceInfo[];
  appSourceColor?: string;
  appSourceName?: string;
  rows: readonly number[];
  headerRow?: number;
  frozenCols?: number;
  getCell: (row: number, col: GridColumn) => CellView | undefined;
  cssOf: (styleId: number | undefined) => CellCss;
  dataVersion: number;
  rowHeight?: (row: number) => number | undefined;
  selection: ColSelection;
  onSelectionChange: (sel: ColSelection) => void;
  active: CellPos | null;
  onActiveChange: (pos: CellPos) => void;
  onMoveColumns: (ids: string[], gap: number) => void;
  onResizeColumn: (id: string, width: number, commit: { from: number } | null) => void;
  onCommitEdit: (pos: CellPos, text: string) => void;
  onDeleteColumns: (ids: string[]) => void;
  onInsertColumns: (index: number, count: number) => void;
  onResetColumnWidths?: (ids: string[]) => void;
  onUndo: () => void;
  onRedo: () => void;
  onEditingChange?: (state: EditingState | null) => void;
  onNotice?: (text: string) => void;
  highlightRows?: ReadonlySet<number>;
  highlightLabel?: string;
  readOnly?: boolean;
  ariaLabel?: string;
  emptyText?: string;
  className?: string;
}

type DragState =
  | { kind: 'pending' | 'move'; pointerId: number; startX: number; lastX: number; lastY: number; ids: string[]; collapseTo: string | null; label?: string }
  | { kind: 'resize'; pointerId: number; colId: string; startX: number; startW: number; lastW: number; moved: boolean }
  | null;

interface DragUi {
  gap: number;
  x: number;
  y: number;
  ids: string[];
  /** ghost text of a band drag (「02_面接評価（7列）を移動」); a header drag builds its own from the titles */
  label?: string;
  /** the pointer is well outside the grid (e.g. over the side panel): releasing there cancels the move */
  outside: boolean;
}

/** how far (px) past the grid's edge a drop still counts as inside (overshooting while auto-scrolling is normal) */
const DROP_MARGIN = 56;

interface EditState {
  mode: 'enter' | 'edit';
  pos: CellPos;
  initial: string;
}

interface MenuState {
  x: number;
  y: number;
  colId: string;
  ids: string[];
}

const EMPTY_SET: ReadonlySet<number> = new Set<number>();

function safeGet(getCell: GridProps['getCell'], row: number, col: GridColumn): CellView | undefined {
  try {
    return getCell(row, col);
  } catch {
    return undefined;
  }
}

// ---------------------------------------------------------------------------------------------

export const Grid = forwardRef<GridHandle, GridProps>(function Grid(props, ref) {
  const { columns, rows, selection, active, readOnly } = props;
  const wrapRef = useRef<HTMLDivElement>(null);
  const scrollRef = useRef<HTMLDivElement>(null);
  const sinkRef = useRef<HTMLTextAreaElement>(null);
  const headerRef = useRef<HTMLDivElement>(null);
  const n = columns.length;
  const nRows = rows.length;
  const frozen = Math.max(0, Math.min(props.frozenCols ?? 0, n));
  const appColor = props.appSourceColor ?? DEFAULT_APP_COLOR;
  const appName = props.appSourceName ?? 'アプリで追加';
  const highlight = props.highlightRows ?? EMPTY_SET;

  const order = useMemo(() => columns.map((c) => c.id), [columns]);
  const indexOf = useMemo(() => new Map(order.map((id, i) => [id, i])), [order]);
  const rowIndex = useMemo(() => makeRowIndex(rows), [rows]);

  // Source bands: maximal runs of contiguous columns of the same source, counted over ALL columns (not only the
  // virtualised ones) and NOT split by the frozen pane (only their drawing is, see bandPieces). The click / drag target.
  const runs = useMemo(() => computeRuns(columns.map((c) => c.source)), [columns]);

  // prefix sums of widths: starts[i] = x of column i inside the data area; starts[n] = total
  const starts = useMemo(() => prefixSums(columns.map((c) => Math.max(0, c.width || 0))), [columns]);
  const frozenW = starts[frozen];
  const totalW = starts[n];
  // Blank scroll room behind the last column. Only scrollToColumnsCentered asks for it (so that the LAST file's columns can be
  // centred like any other; the browser cannot scroll past the content) and it stays until the component unmounts.
  const [endSlack, setEndSlack] = useState(0);
  const pendingScrollRef = useRef<ScrollToOptions | null>(null);

  // row heights (uniform unless rowHeight is given)
  const rowHeightFn = props.rowHeight;
  const heights = useMemo(() => {
    if (!rowHeightFn) return null;
    const h = new Float64Array(nRows);
    for (let i = 0; i < nRows; i++) {
      const v = rowHeightFn(rows[i]);
      h[i] = v !== undefined && v > 0 ? Math.max(8, Math.min(600, v)) : ROW_H;
    }
    return h;
  }, [rowHeightFn, rows, nRows]);
  const rowStarts = useMemo(() => (heights ? prefixSums(heights) : null), [heights]);
  const rowTop = useCallback((idx: number) => (rowStarts ? rowStarts[idx] : idx * ROW_H), [rowStarts]);
  const rowH = useCallback((idx: number) => (heights ? heights[idx] : ROW_H), [heights]);

  const getScrollElement = useCallback(() => scrollRef.current, []);
  // estimateSize is NOT a memo dependency inside virtual-core, getItemKey is (R3 pitfall P1): a new key
  // function whenever rows/heights (or column order/widths) change makes the virtualiser re-read the sizes.
  const rowKey = useCallback((i: number) => i, [rows, heights]); // eslint-disable-line react-hooks/exhaustive-deps
  const rowV = useVirtualizer({
    count: nRows,
    getScrollElement,
    estimateSize: (i) => (heights ? heights[i] : ROW_H),
    getItemKey: rowKey,
    overscan: 8,
    paddingStart: HEADER_H,
    scrollPaddingStart: HEADER_H,
  });
  const colKey = useCallback((i: number) => `${columns[i]?.id}:${columns[i]?.width}`, [columns]);
  const colV = useVirtualizer({
    horizontal: true,
    count: n,
    getScrollElement,
    estimateSize: (i) => Math.max(0, columns[i]?.width || 0),
    getItemKey: colKey,
    overscan: 3,
    paddingStart: ROWHDR_W,
    scrollPaddingStart: ROWHDR_W + frozenW,
  });

  const vRows = rowV.getVirtualItems();
  const vColsAll = colV.getVirtualItems();
  const vCols = useMemo(() => vColsAll.filter((c) => c.index >= frozen), [vColsAll, frozen]);

  // The virtualisers re-render only when the visible RANGE changes, not on every scrolled pixel. Things drawn
  // over the frozen pane in canvas coordinates (active frame / editor of a frozen column) therefore follow the
  // horizontal scroll through a CSS variable instead of React state.
  useEffect(() => {
    const el = scrollRef.current;
    if (!el) return;
    const on = () => el.style.setProperty('--g-sl', `${el.scrollLeft}px`);
    on();
    el.addEventListener('scroll', on, { passive: true });
    return () => el.removeEventListener('scroll', on);
  }, []);

  // latest props for handlers registered once
  const propsRef = useRef(props);
  propsRef.current = props;

  const focusSink = useCallback(() => sinkRef.current?.focus({ preventScroll: true }), []);

  // ------------------------------------------------------------------ active cell
  const activeIdx = active ? (indexOf.get(active.col) ?? -1) : -1;
  const activeRowIdx = active ? rowIndex(active.row) : -1;

  const ensureVisible = useCallback(
    (rowIdx: number, colIdx: number) => {
      if (rowIdx >= 0 && rowIdx < nRows) rowV.scrollToIndex(rowIdx, { align: 'auto' });
      if (colIdx >= frozen && colIdx < n) colV.scrollToIndex(colIdx, { align: 'auto' });
    },
    [rowV, colV, frozen, n, nRows],
  );

  const clearColSelection = () => {
    if (selection.cols.size) props.onSelectionChange({ cols: new Set(), anchor: null });
  };

  const setActiveAt = (rowIdx: number, colIdx: number, opts?: { keepSelection?: boolean }) => {
    if (rowIdx < 0 || colIdx < 0 || rowIdx >= nRows || colIdx >= n) return;
    props.onActiveChange({ row: rows[rowIdx], col: order[colIdx] });
    if (!opts?.keepSelection) clearColSelection();
    ensureVisible(rowIdx, colIdx);
  };

  // ------------------------------------------------------------------ editing (IME-safe sink)
  const [edit, setEdit] = useState<EditState | null>(null);
  const editRef = useRef<EditState | null>(null);
  editRef.current = edit;

  const cellAt = (pos: CellPos): CellView | undefined => {
    const ci = indexOf.get(pos.col);
    return ci === undefined ? undefined : safeGet(props.getCell, pos.row, columns[ci]);
  };

  const beginEdit = (mode: 'enter' | 'edit', initial: string) => {
    if (readOnly || !active || activeIdx < 0 || activeRowIdx < 0) return;
    const t = sinkRef.current;
    if (!t) return;
    t.value = initial;
    const st: EditState = { mode, pos: active, initial };
    setEdit(st);
    props.onEditingChange?.({ pos: active, text: initial });
    ensureVisible(activeRowIdx, activeIdx);
    requestAnimationFrame(() => {
      t.focus({ preventScroll: true });
      t.setSelectionRange(t.value.length, t.value.length);
    });
  };

  const finishEdit = (commitIt: boolean, move: 'down' | 'up' | 'left' | 'right' | null) => {
    const st = editRef.current;
    const t = sinkRef.current;
    if (!st || !t) return;
    const text = t.value;
    t.value = '';
    editRef.current = null;
    setEdit(null);
    propsRef.current.onEditingChange?.(null);
    if (commitIt && !(st.mode === 'edit' && text === st.initial)) propsRef.current.onCommitEdit(st.pos, text);
    if (move) {
      const ri = rowIndex(st.pos.row);
      const ci = indexOf.get(st.pos.col) ?? -1;
      if (ri >= 0 && ci >= 0) {
        const dr = move === 'down' ? 1 : move === 'up' ? -1 : 0;
        const dc = move === 'right' ? 1 : move === 'left' ? -1 : 0;
        setActiveAt(Math.max(0, Math.min(nRows - 1, ri + dr)), Math.max(0, Math.min(n - 1, ci + dc)));
      }
    }
  };

  // text typed while NOT editing lands in the (invisible) sink -> enter mode in place. No element swap
  // happens, so an IME composition that started on the first key continues (R3 §3.4 / P6).
  const onSinkInput = () => {
    const t = sinkRef.current!;
    if (!edit) {
      if (t.value === '') return;
      if (readOnly || !active || activeIdx < 0 || activeRowIdx < 0) {
        t.value = '';
        return;
      }
      const st: EditState = { mode: 'enter', pos: active, initial: '' };
      editRef.current = st;
      setEdit(st);
      ensureVisible(activeRowIdx, activeIdx);
      props.onEditingChange?.({ pos: active, text: t.value });
    } else {
      props.onEditingChange?.({ pos: edit.pos, text: t.value });
    }
  };
  const onCompositionStart = () => {
    if (edit || readOnly || !active || activeIdx < 0 || activeRowIdx < 0) return;
    const st: EditState = { mode: 'enter', pos: active, initial: '' };
    editRef.current = st;
    setEdit(st);
    ensureVisible(activeRowIdx, activeIdx);
    props.onEditingChange?.({ pos: active, text: sinkRef.current?.value ?? '' });
  };

  // ------------------------------------------------------------------ context menu
  const [menu, setMenu] = useState<MenuState | null>(null);
  const openMenu = (x: number, y: number, colId: string) => {
    if (editRef.current) finishEdit(true, null);
    const ids = selection.cols.has(colId) ? orderedSelection(order, selection.cols) : [colId];
    setMenu({ x, y, colId, ids });
  };
  const closeMenu = useCallback(() => {
    setMenu(null);
    // give the keyboard back to the grid, unless a menu action moved the focus elsewhere (e.g. the App's
    // delete-confirm dialog): after the menu unmounts, focus is on <body> only when nobody claimed it
    requestAnimationFrame(() => {
      const ae = document.activeElement;
      if (!ae || ae === document.body || (wrapRef.current && wrapRef.current.contains(ae))) sinkRef.current?.focus({ preventScroll: true });
    });
  }, []);

  const menuItems = (m: MenuState): MenuEntry[] => {
    const left = insertTarget(order, m.ids, m.colId, 'left');
    const right = insertTarget(order, m.ids, m.colId, 'right');
    const cnt = m.ids.length;
    return [
      { key: 'insert-left', label: '左に列を挿入', hint: left.count > 1 ? `${left.count}列` : undefined, onSelect: () => props.onInsertColumns(left.index, left.count) },
      { key: 'insert-right', label: '右に列を挿入', hint: right.count > 1 ? `${right.count}列` : undefined, onSelect: () => props.onInsertColumns(right.index, right.count) },
      'separator',
      { key: 'delete', label: '列を削除', danger: true, hint: cnt > 1 ? `${cnt}列` : 'Delete', onSelect: () => props.onDeleteColumns(m.ids) },
      'separator',
      { key: 'reset-width', label: '列幅を元に戻す', disabled: !props.onResetColumnWidths, onSelect: () => props.onResetColumnWidths?.(m.ids) },
    ];
  };

  // ------------------------------------------------------------------ drag & resize
  const dragRef = useRef<DragState>(null);
  const [dragUi, setDragUi] = useState<DragUi | null>(null);
  const rafRef = useRef(0);
  const lastHandleDownRef = useRef<{ id: string; t: number } | null>(null);
  // depend on the ids ARRAY (same reference for the whole drag), not on dragUi (new object every
  // animation frame) - otherwise every visible cell re-renders 60x per second while dragging (P9)
  const dragIdsArr = dragUi?.ids;
  const dragIds = useMemo(() => (dragIdsArr ? new Set(dragIdsArr) : null), [dragIdsArr]);

  /** insertion gap (0..n) under a viewport x coordinate */
  const gapAtClient = useCallback(
    (clientX: number): number => {
      const el = scrollRef.current;
      if (!el) return 0;
      const rect = el.getBoundingClientRect();
      const vx = clientX - rect.left - ROWHDR_W;
      if (frozen > 0 && vx < frozenW) return gapAt(vx, starts, 0, frozen);
      return gapAt(vx + el.scrollLeft, starts, frozen, n);
    },
    [frozen, frozenW, n, starts],
  );
  const gapAtRef = useRef(gapAtClient);
  gapAtRef.current = gapAtClient;

  /** pointer far outside the grid (side panel, file cards, another window): a drop there is a cancel, not a move to the end */
  const isOutsideGrid = useCallback((clientX: number, clientY: number): boolean => {
    const el = scrollRef.current;
    if (!el) return false;
    const r = el.getBoundingClientRect();
    return clientX < r.left - DROP_MARGIN || clientX > r.right + DROP_MARGIN || clientY < r.top - DROP_MARGIN || clientY > r.bottom + DROP_MARGIN;
  }, []);

  const stopAutoScroll = () => {
    cancelAnimationFrame(rafRef.current);
    rafRef.current = 0;
  };

  const tick = useCallback(() => {
    const d = dragRef.current;
    const el = scrollRef.current;
    if (!d || d.kind !== 'move' || !el) {
      rafRef.current = 0;
      return;
    }
    const rect = el.getBoundingClientRect();
    const leftEdge = rect.left + ROWHDR_W + frozenW;
    const rightEdge = rect.left + el.clientWidth;
    let v = 0;
    if (d.lastX > rightEdge - EDGE) v = Math.min(MAX_SPEED, Math.ceil(((d.lastX - (rightEdge - EDGE)) / EDGE) * MAX_SPEED));
    else if (d.lastX < leftEdge + EDGE && (frozen === 0 || d.lastX > leftEdge - 8)) {
      // with a frozen pane, moving further left (into the pane) stops scrolling so the pane can be targeted
      v = -Math.min(MAX_SPEED, Math.ceil(((leftEdge + EDGE - d.lastX) / EDGE) * MAX_SPEED));
    }
    if (v !== 0) el.scrollLeft += v; // scroll event -> the virtualiser re-renders the new range
    setDragUi({ gap: gapAtRef.current(d.lastX), x: d.lastX, y: d.lastY, ids: d.ids, label: d.label, outside: isOutsideGrid(d.lastX, d.lastY) });
    rafRef.current = requestAnimationFrame(tick);
  }, [frozenW, frozen, isOutsideGrid]);

  useEffect(() => () => stopAutoScroll(), []);

  const releaseCapture = (pointerId: number) => {
    const h = headerRef.current;
    try {
      if (h?.hasPointerCapture(pointerId)) h.releasePointerCapture(pointerId);
    } catch {
      /* ignore */
    }
  };

  /** cancel a running drag / resize (Esc, pointercancel, lost capture) */
  const cancelGesture = useCallback(() => {
    const d = dragRef.current;
    if (!d) return;
    dragRef.current = null;
    stopAutoScroll();
    document.body.classList.remove('g-dragging', 'g-resizing');
    setDragUi(null);
    releaseCapture(d.pointerId);
    if (d.kind === 'resize' && d.moved) propsRef.current.onResizeColumn(d.colId, d.startW, { from: d.startW });
  }, []);

  // Esc cancels a running drag / resize (capture phase: before the sink sees it)
  useEffect(() => {
    const onKey = (ev: KeyboardEvent) => {
      if (ev.key === 'Escape' && dragRef.current) {
        ev.preventDefault();
        ev.stopPropagation();
        cancelGesture();
      }
    };
    window.addEventListener('keydown', onKey, true);
    return () => window.removeEventListener('keydown', onKey, true);
  }, [cancelGesture]);

  const autoFit = (ids: string[]) => {
    const p = propsRef.current;
    for (const id of ids) {
      const ci = indexOf.get(id);
      if (ci === undefined) continue;
      const col = columns[ci];
      const w = autoFitWidth(
        nRows,
        (i) => {
          const cell = safeGet(p.getCell, rows[i], col);
          if (!cell) return null;
          const css = p.cssOf(cell.styleId);
          const r = renderCell(cell, css);
          return r.text ? { text: r.text, font: css.font } : null;
        },
        { text: col.title, font: HEADER_FONT, extra: col.keyRole ? 16 : 14 },
        { min: 24, max: 600, padding: 10 },
      );
      if (w !== col.width) p.onResizeColumn(id, w, { from: col.width });
    }
  };

  /** pointer down on a source band: select the run (header-click semantics) and arm a drag of the whole run */
  const onBandPointerDown = (e: React.PointerEvent<HTMLDivElement>, bandEl: HTMLElement) => {
    const run = runs[Number(bandEl.dataset.band)];
    if (!run) return;
    const ids = order.slice(run.s, run.e + 1);
    if (editRef.current) finishEdit(true, null);
    if (menu) setMenu(null);
    if (e.button === 2) {
      if (!ids.every((id) => selection.cols.has(id))) props.onSelectionChange(selectRun(order, selection, ids, { ctrl: false, shift: false }));
      return;
    }
    if (e.button !== 0) return;
    // Capture on the header CONTAINER, which is never unmounted (R3 P2): the band itself is virtualised away while the drag scrolls
    try {
      headerRef.current?.setPointerCapture(e.pointerId);
    } catch {
      /* ignore (synthetic events) */
    }
    lastHandleDownRef.current = null;
    const next = selectRun(order, selection, ids, { ctrl: e.ctrlKey || e.metaKey, shift: e.shiftKey });
    props.onSelectionChange(next);
    const first = ids[0];
    if (next.cols.has(first) && nRows > 0) {
      const row = active && rowIndex(active.row) >= 0 ? active.row : rows[0];
      props.onActiveChange({ row, col: first });
    }
    const name = run.source >= 0 ? (props.sources[run.source]?.name ?? appName) : appName;
    dragRef.current = next.cols.has(first)
      ? { kind: 'pending', pointerId: e.pointerId, startX: e.clientX, lastX: e.clientX, lastY: e.clientY, ids, collapseTo: null, label: `${name}（${ids.length}列）を移動` }
      : null;
    focusSink();
  };

  const onHeaderPointerDown = (e: React.PointerEvent<HTMLDivElement>) => {
    const target = e.target as HTMLElement;
    // resize handles are separate elements ON the column boundary (R3 P7)
    const handle = target.closest<HTMLElement>('[data-resize]');
    const cell = target.closest<HTMLElement>('[data-col]');
    if (!cell && !handle) {
      const bandEl = target.closest<HTMLElement>('[data-band]');
      if (bandEl) onBandPointerDown(e, bandEl);
      return;
    }
    const id = handle ? handle.dataset.resize! : cell!.dataset.col!;
    const ci = indexOf.get(id);
    if (ci === undefined) return;
    if (editRef.current) finishEdit(true, null);
    if (menu) setMenu(null);
    if (e.button === 2) {
      if (!selection.cols.has(id)) props.onSelectionChange(selectOnly(id));
      return;
    }
    if (e.button !== 0) return;
    // Capture on the header CONTAINER, which is never unmounted (R3 P2)
    try {
      headerRef.current?.setPointerCapture(e.pointerId);
    } catch {
      /* ignore (synthetic events) */
    }
    if (handle) {
      const now = performance.now();
      const last = lastHandleDownRef.current;
      if (last && last.id === id && now - last.t < DBLCLICK_MS) {
        lastHandleDownRef.current = null;
        releaseCapture(e.pointerId);
        const ids = selection.cols.has(id) && selection.cols.size > 1 ? orderedSelection(order, selection.cols) : [id];
        autoFit(ids);
        return;
      }
      lastHandleDownRef.current = { id, t: now };
      const w = columns[ci].width;
      dragRef.current = { kind: 'resize', pointerId: e.pointerId, colId: id, startX: e.clientX, startW: w, lastW: w, moved: false };
      return;
    }
    lastHandleDownRef.current = null;
    const ctrl = e.ctrlKey || e.metaKey;
    const shift = e.shiftKey;
    let next: ColSelection = selection;
    let collapseTo: string | null = null;
    if (!ctrl && !shift && selection.cols.has(id)) {
      // pressing on an already-selected column keeps the multi-selection so it can be dragged;
      // a plain click without drag collapses to this column on pointerup (Explorer semantics)
      collapseTo = selection.cols.size > 1 ? id : null;
    } else {
      next = clickSelect(order, selection, id, { ctrl, shift });
      props.onSelectionChange(next);
    }
    if (next.cols.has(id) && nRows > 0) {
      const row = active && rowIndex(active.row) >= 0 ? active.row : rows[0];
      props.onActiveChange({ row, col: id });
    }
    dragRef.current = next.cols.has(id)
      ? { kind: 'pending', pointerId: e.pointerId, startX: e.clientX, lastX: e.clientX, lastY: e.clientY, ids: orderedSelection(order, next.cols), collapseTo }
      : null;
    focusSink();
  };

  const onHeaderPointerMove = (e: React.PointerEvent<HTMLDivElement>) => {
    const d = dragRef.current;
    if (!d || d.pointerId !== e.pointerId) return;
    if (d.kind === 'resize') {
      const w = Math.max(MIN_COL_W, Math.min(MAX_COL_W, Math.round(d.startW + e.clientX - d.startX)));
      if (!d.moved && Math.abs(e.clientX - d.startX) < 2) return;
      if (!d.moved) {
        d.moved = true;
        document.body.classList.add('g-resizing');
      }
      if (w !== d.lastW) {
        d.lastW = w;
        props.onResizeColumn(d.colId, w, null);
      }
      return;
    }
    d.lastX = e.clientX;
    d.lastY = e.clientY;
    if (d.kind === 'pending' && Math.abs(e.clientX - d.startX) >= DRAG_THRESHOLD) {
      d.kind = 'move';
      document.body.classList.add('g-dragging');
    }
    if (d.kind === 'move' && !rafRef.current) rafRef.current = requestAnimationFrame(tick);
  };

  const endDrag = (e: React.PointerEvent<HTMLDivElement>, cancelled: boolean) => {
    const d = dragRef.current;
    if (!d || d.pointerId !== e.pointerId) return;
    if (cancelled) {
      cancelGesture();
      return;
    }
    dragRef.current = null;
    stopAutoScroll();
    document.body.classList.remove('g-dragging', 'g-resizing');
    setDragUi(null);
    releaseCapture(e.pointerId);
    if (d.kind === 'resize') {
      if (d.moved) {
        const w = Math.max(MIN_COL_W, Math.min(MAX_COL_W, Math.round(d.startW + e.clientX - d.startX)));
        props.onResizeColumn(d.colId, w, { from: d.startW });
      }
    } else if (d.kind === 'move') {
      const gap = gapAtClient(e.clientX);
      // released far outside the grid = the user let go of it (over the side panel …): nothing moves
      if (!isOutsideGrid(e.clientX, e.clientY) && !isNoopMove(order, new Set(d.ids), gap)) props.onMoveColumns(d.ids, gap);
    } else if (d.kind === 'pending' && d.collapseTo) {
      props.onSelectionChange(selectOnly(d.collapseTo));
    }
    focusSink();
  };

  // ------------------------------------------------------------------ keyboard
  const pageRows = () => Math.max(1, Math.floor(((scrollRef.current?.clientHeight ?? 400) - HEADER_H) / ROW_H) - 1);

  const onSinkKeyDown = (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
    // Never act on keys that belong to an IME composition (the Enter that confirms a conversion etc.)
    if (e.nativeEvent.isComposing || e.keyCode === 229) return;
    const ctrl = e.ctrlKey || e.metaKey;
    const key = e.key;
    if (edit) {
      const t = e.currentTarget;
      if (key === 'Enter' && e.altKey) {
        e.preventDefault();
        t.setRangeText('\n', t.selectionStart, t.selectionEnd, 'end');
        props.onEditingChange?.({ pos: edit.pos, text: t.value });
      } else if (key === 'Enter') {
        e.preventDefault();
        finishEdit(true, ctrl ? null : e.shiftKey ? 'up' : 'down');
      } else if (key === 'Tab') {
        e.preventDefault();
        finishEdit(true, e.shiftKey ? 'left' : 'right');
      } else if (key === 'Escape') {
        e.preventDefault();
        finishEdit(false, null);
      } else if (key === 'F2') {
        e.preventDefault();
        setEdit({ ...edit, mode: edit.mode === 'enter' ? 'edit' : 'enter' });
      } else if (edit.mode === 'enter' && key.startsWith('Arrow') && !e.shiftKey && !ctrl) {
        e.preventDefault();
        finishEdit(true, key === 'ArrowDown' ? 'down' : key === 'ArrowUp' ? 'up' : key === 'ArrowLeft' ? 'left' : 'right');
      }
      // Ctrl+S / Ctrl+O are the app's shortcuts (export / file chooser): let them bubble up to the window
      if (!(ctrl && (key === 's' || key === 'S' || key === 'o' || key === 'O'))) e.stopPropagation();
      return; // everything else (incl. Ctrl+Z inside the text) is native textarea behaviour
    }
    const handled = () => {
      e.preventDefault();
      e.stopPropagation();
    };
    const lower = key.length === 1 ? key.toLowerCase() : key;
    if (ctrl && !e.altKey) {
      if (lower === 'z') {
        handled();
        if (e.shiftKey) props.onRedo();
        else props.onUndo();
        return;
      }
      if (lower === 'y') {
        handled();
        props.onRedo();
        return;
      }
      if (key === '-' && selection.cols.size) {
        handled();
        props.onDeleteColumns(orderedSelection(order, selection.cols));
        return;
      }
      if ((key === '+' || (key === '=' && e.shiftKey) || key === ';') && selection.cols.size && active) {
        handled();
        const ids = orderedSelection(order, selection.cols);
        const t = insertTarget(order, ids, selection.cols.has(active.col) ? active.col : ids[0], 'left');
        props.onInsertColumns(t.index, t.count);
        return;
      }
      if (key === ' ' || key === 'Spacebar') {
        handled();
        if (active && activeIdx >= 0) props.onSelectionChange(selectOnly(active.col));
        return;
      }
      if (lower === 'c' || lower === 'x' || lower === 'v' || lower === 'a') return; // copy event / native
      if (lower === 's' || lower === 'o') return; // the app's export / file-chooser shortcuts: bubble up to the window
    }
    if ((key === 'F10' && e.shiftKey) || key === 'ContextMenu') {
      handled();
      if (active && activeIdx >= 0) {
        const r = scrollRef.current?.querySelector<HTMLElement>(`[data-testid="grid-header-${activeIdx}"]`)?.getBoundingClientRect();
        const box = scrollRef.current!.getBoundingClientRect();
        openMenu(r ? r.left + 8 : box.left + ROWHDR_W + 8, r ? r.bottom : box.top + HEADER_H, active.col);
      }
      return;
    }
    // Shift+Left/Right with a column selection extends it (Excel)
    if (e.shiftKey && !ctrl && selection.cols.size && (key === 'ArrowLeft' || key === 'ArrowRight') && activeIdx >= 0) {
      handled();
      const ci = Math.max(0, Math.min(n - 1, activeIdx + (key === 'ArrowRight' ? 1 : -1)));
      const id = order[ci];
      props.onSelectionChange(extendTo(order, selection, id));
      if (active) props.onActiveChange({ row: active.row, col: id });
      ensureVisible(-1, ci);
      return;
    }
    if (key === 'Tab' && !ctrl && !e.altKey) {
      // Tab moves the active cell (Excel) but must not trap keyboard users: at the last / first column the focus leaves the grid
      const here = { rowIdx: Math.max(0, activeRowIdx), colIdx: Math.max(0, activeIdx) };
      const there = nRows === 0 || n === 0 ? null : navigate('Tab', { ctrl: false, shift: e.shiftKey }, here, { rows: nRows, cols: n, page: 1 });
      if (!there || (there.rowIdx === here.rowIdx && there.colIdx === here.colIdx)) return;
    }
    if (isNavKey(key)) {
      handled();
      if (nRows === 0 || n === 0) return;
      const at = { rowIdx: Math.max(0, activeRowIdx), colIdx: Math.max(0, activeIdx) };
      const next = navigate(key, { ctrl, shift: e.shiftKey }, at, { rows: nRows, cols: n, page: pageRows() });
      if (next) setActiveAt(next.rowIdx, next.colIdx);
      return;
    }
    switch (key) {
      case 'F2':
        handled();
        if (active) {
          const cell = cellAt(active);
          beginEdit('edit', cellEditText(cell, props.cssOf(cell?.styleId)));
        }
        return;
      case 'Delete':
        handled();
        if (selection.cols.size) props.onDeleteColumns(orderedSelection(order, selection.cols));
        else if (active && !readOnly && !isBlankCell(cellAt(active))) props.onCommitEdit(active, '');
        return;
      case 'Backspace':
        handled();
        if (active) beginEdit('enter', '');
        return;
      case 'Escape':
        if (selection.cols.size) {
          handled();
          clearColSelection();
        }
        return;
      default:
        // printable keys: the default action inserts them into the sink -> onSinkInput (enter mode)
        if (ctrl || e.altKey) e.stopPropagation();
        return;
    }
  };

  const onSinkCopy = (e: React.ClipboardEvent<HTMLTextAreaElement>) => {
    if (edit) return; // native copy of the selected text while editing
    if (!active) return;
    const cell = cellAt(active);
    e.preventDefault();
    e.clipboardData.setData('text/plain', cellCopyText(cell, props.cssOf(cell?.styleId)));
  };

  const onSinkPaste = (e: React.ClipboardEvent<HTMLTextAreaElement>) => {
    if (edit || readOnly) return;
    // paste into the active cell (first line / first cell of a copied range only)
    const text = e.clipboardData.getData('text/plain');
    e.preventDefault();
    if (!active || activeIdx < 0) return;
    const lines = text.replace(/\r\n?/g, '\n').replace(/\n$/, '').split('\n');
    const first = lines[0]?.split('\t')[0] ?? '';
    if (lines.length > 1 || (lines[0]?.includes('\t') ?? false)) props.onNotice?.('複数のセルの貼り付けには対応していません。先頭のセルだけを貼り付けました');
    props.onCommitEdit(active, first);
  };

  // keep focus in the sink: on mount and whenever editing stops, unless the user moved focus elsewhere
  useLayoutEffect(() => {
    if (edit) return;
    const ae = document.activeElement;
    if (!ae || ae === document.body || (wrapRef.current && wrapRef.current.contains(ae))) focusSink();
  }, [edit, focusSink]);

  // an edit whose cell disappeared (column deleted / rows changed) is dropped
  useEffect(() => {
    if (!edit) return;
    if (!indexOf.has(edit.pos.col) || rowIndex(edit.pos.row) < 0) {
      if (sinkRef.current) sinkRef.current.value = '';
      editRef.current = null;
      setEdit(null);
      propsRef.current.onEditingChange?.(null);
    }
  }, [edit, indexOf, rowIndex]);

  // ------------------------------------------------------------------ body pointer
  const onBodyPointerDown = (e: React.PointerEvent<HTMLDivElement>) => {
    if (e.target === sinkRef.current) return;
    const cell = (e.target as HTMLElement).closest<HTMLElement>('[data-r]');
    if (!cell) return;
    if (e.button !== 0 && e.button !== 2) return;
    const pos = { row: Number(cell.dataset.r), col: cell.dataset.c! };
    if (editRef.current) {
      const st = editRef.current;
      if (st.pos.row === pos.row && st.pos.col === pos.col) return; // click inside the cell being edited
      finishEdit(true, null);
    }
    if (menu) setMenu(null);
    if (e.button === 2 && selection.cols.has(pos.col)) {
      props.onActiveChange(pos); // keep the column selection for the menu
    } else {
      props.onActiveChange(pos);
      clearColSelection();
    }
    focusSink();
  };
  const onBodyMouseDown = (e: React.MouseEvent) => {
    // keep focus in the sink and stop native text selection; not on the scrollbar / the editor itself
    if (e.target !== scrollRef.current && e.target !== sinkRef.current) e.preventDefault();
  };
  const onBodyDoubleClick = (e: React.MouseEvent) => {
    const cellEl = (e.target as HTMLElement).closest<HTMLElement>('[data-r]');
    if (!cellEl || readOnly) return;
    const pos = { row: Number(cellEl.dataset.r), col: cellEl.dataset.c! };
    const ci = indexOf.get(pos.col);
    const ri = rowIndex(pos.row);
    if (ci === undefined || ri < 0) return;
    const cell = safeGet(props.getCell, pos.row, columns[ci]);
    const t = sinkRef.current;
    if (!t) return;
    props.onActiveChange(pos);
    const initial = cellEditText(cell, props.cssOf(cell?.styleId));
    t.value = initial;
    const st: EditState = { mode: 'edit', pos, initial };
    editRef.current = st;
    setEdit(st);
    props.onEditingChange?.({ pos, text: initial });
    requestAnimationFrame(() => {
      t.focus({ preventScroll: true });
      t.setSelectionRange(t.value.length, t.value.length);
    });
  };
  const onBodyContextMenu = (e: React.MouseEvent) => {
    const cellEl = (e.target as HTMLElement).closest<HTMLElement>('[data-r]');
    if (!cellEl || e.target === sinkRef.current) return;
    e.preventDefault();
    openMenu(e.clientX, e.clientY, cellEl.dataset.c!);
  };

  // ------------------------------------------------------------------ imperative handle
  useImperativeHandle(
    ref,
    (): GridHandle => ({
      scrollToRow(row, opts) {
        const ri = rowIndex(row);
        if (ri < 0) return false;
        rowV.scrollToIndex(ri, { align: 'center' });
        const p = propsRef.current;
        if (opts?.select && order.length) {
          const col = p.active && indexOf.has(p.active.col) ? p.active.col : order[Math.min(frozen, order.length - 1)];
          p.onActiveChange({ row, col });
        }
        return true;
      },
      scrollToColumn(id, opts) {
        const ci = indexOf.get(id);
        if (ci === undefined) return false;
        if (ci >= frozen) colV.scrollToIndex(ci, { align: 'auto' });
        const p = propsRef.current;
        if (opts?.select) {
          p.onSelectionChange(selectOnly(id));
          if (rows.length) {
            const row = p.active && rowIndex(p.active.row) >= 0 ? p.active.row : rows[0];
            p.onActiveChange({ row, col: id });
          }
        }
        return true;
      },
      scrollToCell(pos) {
        const ci = indexOf.get(pos.col);
        const ri = rowIndex(pos.row);
        if (ci === undefined || ri < 0) return false;
        rowV.scrollToIndex(ri, { align: 'center' });
        if (ci >= frozen) colV.scrollToIndex(ci, { align: 'auto' });
        propsRef.current.onActiveChange(pos);
        return true;
      },
      scrollToColumnsCentered(ids, opts) {
        const el = scrollRef.current;
        const idx = ids.map((id) => indexOf.get(id)).filter((i): i is number => i !== undefined);
        if (!el || idx.length === 0) return false;
        // the first contiguous run of the ids (a file whose columns are split over several places: centre the first one)
        const [s, e] = toRuns(idx)[0];
        let left = el.scrollLeft;
        let slack = 0;
        if (e >= frozen) {
          const s2 = Math.max(s, frozen); // the part inside the frozen pane is always visible
          left = centerScrollLeft(ROWHDR_W + starts[s2], ROWHDR_W + starts[e + 1], ROWHDR_W + frozenW, el.clientWidth, Number.POSITIVE_INFINITY);
          // the group may lie so far right that the content ends before it can reach the middle: add blank room behind it
          // (+ 24: the scroller's real maximum is a little below scrollWidth - clientWidth because of its scrollbar gutter)
          const need = left + el.clientWidth - (ROWHDR_W + totalW);
          slack = need > -24 ? Math.min(el.clientWidth, need + 24) : 0;
        }
        const reduce = typeof window.matchMedia === 'function' && window.matchMedia('(prefers-reduced-motion: reduce)').matches;
        const behavior = opts?.behavior ?? (reduce ? 'instant' : 'smooth');
        const scroll: ScrollToOptions = { left, top: 0, behavior };
        if (slack > endSlack) {
          pendingScrollRef.current = scroll; // applied once the wider canvas is laid out
          setEndSlack(Math.ceil(slack));
        } else {
          el.scrollTo(scroll);
        }
        if (opts?.select) {
          const p = propsRef.current;
          const known = ids.filter((id) => indexOf.has(id));
          p.onSelectionChange({ cols: new Set(known), anchor: known[0] ?? null });
          if (rows.length) p.onActiveChange({ row: rows[0], col: known[0] });
        }
        return true;
      },
      focus() {
        focusSink();
      },
      isEditing() {
        return editRef.current !== null;
      },
      commitEdit() {
        if (editRef.current) finishEdit(true, null);
      },
    }),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [rowIndex, indexOf, order, rows, frozen, frozenW, totalW, endSlack, starts, rowV, colV, focusSink],
  );

  // scrollToColumnsCentered waited for the extra room behind the last column
  useLayoutEffect(() => {
    const pending = pendingScrollRef.current;
    if (!pending) return;
    pendingScrollRef.current = null;
    scrollRef.current?.scrollTo(pending);
  }, [endSlack]);

  // ------------------------------------------------------------------ geometry
  /** x of column i inside the canvas, taking the frozen pane into account (P11; --g-sl = scrollLeft) */
  const colX = (i: number): number | string =>
    i < frozen ? `calc(var(--g-sl, 0px) + ${ROWHDR_W + starts[i]}px)` : ROWHDR_W + starts[i];

  const activeRect =
    activeIdx >= 0 && activeRowIdx >= 0
      ? { left: colX(activeIdx), top: HEADER_H + rowTop(activeRowIdx), width: columns[activeIdx].width, height: rowH(activeRowIdx) }
      : null;

  // The bands are DRAWN in pieces: a run that crosses the frozen boundary is cut there (the pinned part is sticky), but
  // every piece belongs to its logical run: same label count (the run's true size), same click / drag target.
  const pieces = useMemo(() => bandPieces(runs, frozen), [runs, frozen]);

  const firstVis = vCols.length ? vCols[0].index : frozen;
  const lastVis = vCols.length ? vCols[vCols.length - 1].index : frozen - 1;

  const sourceColor = (s: number) => (s >= 0 ? props.sources[s]?.color : undefined) ?? appColor;
  const sourceName = (s: number) => (s >= 0 ? props.sources[s]?.name : undefined) ?? appName;

  const renderBand = (piece: BandPiece) => {
    const run = runs[piece.run];
    const w = starts[piece.e + 1] - starts[piece.s];
    const name = sourceName(run.source);
    const selected = runFullySelected(order, run, selection.cols);
    return (
      <div
        key={piece.pinned && piece.split ? `band-p-${run.s}` : `band-${piece.s}`}
        className={`g-band${selected ? ' is-sel' : ''}`}
        data-band={piece.run}
        data-testid={piece.pinned && piece.split ? `grid-band-${piece.run}-pin` : `grid-band-${piece.run}`}
        role="button"
        aria-pressed={selected}
        aria-label={`${name} ${run.count}列（クリックで選択、ドラッグで移動）`}
        style={{ left: ROWHDR_W + starts[piece.s], width: w, background: sourceColor(run.source) }}
        title={`${name}（${run.count}列）\nドラッグで移動／クリックで選択`}
      >
        <span className="g-band-label" style={piece.pinned ? undefined : { left: ROWHDR_W + frozenW + 4 }}>
          {name}
          <span className="g-band-count">{run.count}列</span>
        </span>
      </div>
    );
  };

  const renderHeaderCell = (i: number) => {
    const c = columns[i];
    const sel = selection.cols.has(c.id);
    const dragging = dragIds?.has(c.id);
    const letter = colLetter(i);
    const src = sourceName(c.source ?? -1);
    return (
      <div
        key={c.id}
        role="columnheader"
        aria-colindex={i + 1}
        aria-selected={sel}
        data-col={c.id}
        data-testid={`grid-header-${i}`}
        className={`g-hcell${sel ? ' is-sel' : ''}${dragging ? ' is-dragging' : ''}${activeIdx === i ? ' is-activecol' : ''}${c.hidden ? ' is-hiddencol' : ''}${i === frozen - 1 ? ' is-frozen-edge' : ''}`}
        style={{ left: ROWHDR_W + starts[i], width: c.width, top: BAND_H, height: TITLE_H }}
        title={`${letter}列「${c.title}」\nファイル: ${src}${c.keyRole ? `\n結合キー: ${c.keyRole}` : ''}${c.hidden ? '\n（非表示の列）' : ''}`}
      >
        <div className="g-hcell-top">
          <span className="g-letter">{letter}</span>
          {c.keyRole && (
            <span className="g-keybadge" aria-label={`結合キー ${c.keyRole}`}>
              キー:{c.keyRole}
            </span>
          )}
        </div>
        <div className="g-htitle">{c.title}</div>
      </div>
    );
  };

  const renderResizeHandle = (i: number) => (
    <div
      key={'rz-' + columns[i].id}
      className="g-resize"
      data-resize={columns[i].id}
      title="ドラッグで列幅を変更（ダブルクリックで自動調整）"
      style={{ left: ROWHDR_W + starts[i + 1] - 4, top: BAND_H, height: TITLE_H }}
    />
  );

  // drop indicator x in viewport coordinates (relative to the scroller box)
  const scEl = scrollRef.current;
  let indicatorX: number | null = null;
  if (dragUi) {
    const g = dragUi.gap;
    const sl = scEl?.scrollLeft ?? 0; // read live: the drag re-renders every animation frame
    indicatorX = g <= frozen ? ROWHDR_W + starts[g] : Math.max(ROWHDR_W + frozenW, ROWHDR_W + starts[g] - sl);
    if (scEl) indicatorX = Math.min(indicatorX, scEl.clientWidth - 2);
  }
  const wrapRect = dragUi ? wrapRef.current?.getBoundingClientRect() : undefined;

  const activeCol = activeIdx >= 0 ? columns[activeIdx] : null;
  const activeCell = active && activeCol ? safeGet(props.getCell, active.row, activeCol) : undefined;
  const activeCss = props.cssOf(activeCell?.styleId);
  const liveText =
    active && activeCol
      ? `${colLetter(activeIdx)}${active.row} ${activeCol.title}: ${renderCell(activeCell, activeCss).text || '空白'}`
      : '';

  const sinkStyle: CSSProperties = activeRect
    ? {
        left: activeRect.left,
        top: activeRect.top,
        minWidth: activeRect.width,
        width: edit ? undefined : activeRect.width,
        minHeight: activeRect.height,
        height: edit ? undefined : activeRect.height,
        // the editor uses the cell's font (longhands only: React warns when shorthand and longhand mix)
        fontFamily: activeCss.box.fontFamily as string | undefined,
        fontSize: activeCss.box.fontSize as string | undefined,
        fontWeight: activeCss.box.fontWeight as number | undefined,
        fontStyle: activeCss.box.fontStyle as string | undefined,
      }
    : { left: 0, top: HEADER_H, width: 1, height: 1 };

  const empty = n === 0 || nRows === 0;

  return (
    <div ref={wrapRef} className={`g-wrap${props.className ? ' ' + props.className : ''}`}>
      <div
        ref={scrollRef}
        className="g-scroller"
        role="grid"
        data-testid="grid"
        aria-rowcount={nRows + 1}
        aria-colcount={n}
        aria-multiselectable
        aria-label={props.ariaLabel ?? '結合結果の表'}
        onPointerDown={onBodyPointerDown}
        onMouseDown={onBodyMouseDown}
        onDoubleClick={onBodyDoubleClick}
        onContextMenu={onBodyContextMenu}
      >
        <div className="g-canvas" style={{ width: ROWHDR_W + totalW + endSlack, height: Math.max(rowV.getTotalSize(), HEADER_H) }}>
          {/* ---------------- header ---------------- */}
          <div
            ref={headerRef}
            className="g-header"
            role="row"
            aria-rowindex={1}
            style={{ width: ROWHDR_W + totalW, height: HEADER_H }}
            onPointerDown={(e) => {
              e.stopPropagation();
              onHeaderPointerDown(e);
            }}
            onPointerMove={onHeaderPointerMove}
            onPointerUp={(e) => endDrag(e, false)}
            onPointerCancel={(e) => endDrag(e, true)}
            onLostPointerCapture={(e) => endDrag(e, true)}
            onMouseDown={(e) => {
              e.stopPropagation();
              e.preventDefault();
            }}
            onDoubleClick={(e) => e.stopPropagation()}
            onContextMenu={(e) => {
              e.preventDefault();
              e.stopPropagation();
              const t = e.target as HTMLElement;
              const cell = t.closest<HTMLElement>('[data-col]');
              if (cell) {
                openMenu(e.clientX, e.clientY, cell.dataset.col!);
                return;
              }
              // a band: the column menu of the run (its columns were selected by the right button press)
              const run = runs[Number(t.closest<HTMLElement>('[data-band]')?.dataset.band)];
              if (run) openMenu(e.clientX, e.clientY, order[run.s]);
            }}
          >
            {pieces.filter((p) => !p.pinned && p.e >= firstVis && p.s <= lastVis).map((p) => renderBand(p))}
            {vCols.map((vc) => renderHeaderCell(vc.index))}
            {vCols.map((vc) => renderResizeHandle(vc.index))}
            <div className="g-pin" style={{ width: ROWHDR_W + frozenW, height: HEADER_H }}>
              <div className="g-corner" style={{ width: ROWHDR_W, height: HEADER_H }} title="見出し行">
                <span className="g-corner-band" style={{ height: BAND_H }} />
                <span className="g-corner-row">{props.headerRow ?? 1}</span>
              </div>
              {pieces.filter((p) => p.pinned).map((p) => renderBand(p))}
              {Array.from({ length: frozen }, (_, i) => renderHeaderCell(i))}
              {Array.from({ length: frozen }, (_, i) => renderResizeHandle(i))}
            </div>
          </div>

          {/* ---------------- rows ---------------- */}
          {vRows.map((vr) => {
            const excelRow = rows[vr.index];
            return (
              <Row
                key={vr.index}
                idx={vr.index}
                row={excelRow}
                top={vr.start}
                height={vr.size}
                vCols={vCols}
                columns={columns}
                starts={starts}
                frozen={frozen}
                frozenW={frozenW}
                totalW={totalW}
                selected={selection.cols}
                dragIds={dragIds}
                getCell={props.getCell}
                cssOf={props.cssOf}
                dataVersion={props.dataVersion}
                highlighted={highlight.has(excelRow)}
                highlightLabel={props.highlightLabel}
                isActiveRow={activeRowIdx === vr.index}
              />
            );
          })}

          {/* ---------------- active cell + IME-safe editor ---------------- */}
          {activeRect && <div className={`g-active${activeIdx < frozen ? ' is-frozen' : ''}`} style={activeRect} aria-hidden />}
          <textarea
            ref={sinkRef}
            className={`g-sink${edit ? ' is-editing' : ''}${activeIdx >= 0 && activeIdx < frozen ? ' is-frozen' : ''}`}
            data-testid="grid-editor"
            style={sinkStyle}
            aria-label={active && activeIdx >= 0 ? `セル ${colLetter(activeIdx)}${active.row} に入力` : 'セル入力'}
            spellCheck={false}
            autoComplete="off"
            autoCorrect="off"
            autoCapitalize="off"
            readOnly={readOnly}
            rows={1}
            onKeyDown={onSinkKeyDown}
            onInput={onSinkInput}
            onCompositionStart={onCompositionStart}
            onCopy={onSinkCopy}
            onPaste={onSinkPaste}
            onBlur={() => {
              // window deactivation (Alt+Tab) keeps the edit, like Excel
              if (editRef.current && document.hasFocus()) finishEdit(true, null);
            }}
          />
          {empty && (
            <div className="g-empty" style={{ top: HEADER_H + 24, left: ROWHDR_W + 16 }}>
              {props.emptyText ?? (n === 0 ? '表示する列がありません' : '表示する行がありません')}
            </div>
          )}
        </div>
      </div>

      {/* ---------------- overlay: drop indicator + ghost ---------------- */}
      <div className="g-overlay" aria-hidden>
        {dragUi && !dragUi.outside && indicatorX !== null && (
          <div className="g-drop" data-testid="grid-drop-indicator" style={{ left: indicatorX - 2, height: scEl?.clientHeight ?? 0 }}>
            <span className="g-drop-cap" />
          </div>
        )}
        {dragUi && wrapRect && (
          <div className={`g-ghost${dragUi.outside ? ' g-ghost-cancel' : ''}`} style={{ left: dragUi.x - wrapRect.left + 14, top: dragUi.y - wrapRect.top + 10 }}>
            {dragUi.outside ? (
              <>表の外です（ここで離すと移動を取り消します）</>
            ) : dragUi.label ? (
              <>{dragUi.label}</>
            ) : (
              <>
                {dragUi.ids.length}列を移動: {dragUi.ids.slice(0, 3).map((id) => columns[indexOf.get(id) ?? -1]?.title ?? '').join('、')}
                {dragUi.ids.length > 3 ? ' …' : ''}
              </>
            )}
          </div>
        )}
      </div>

      {menu &&
        createPortal(
          <ContextMenu x={menu.x} y={menu.y} items={menuItems(menu)} onClose={closeMenu} label="列のメニュー" />,
          document.body,
        )}
      <div className="g-vh" aria-live="polite">
        {liveText}
      </div>
    </div>
  );
});

// ---------------------------------------------------------------------------------------------

interface RowProps {
  idx: number;
  row: number;
  top: number;
  height: number;
  vCols: VirtualItem[];
  columns: readonly GridColumn[];
  starts: Float64Array;
  frozen: number;
  frozenW: number;
  totalW: number;
  selected: ReadonlySet<string>;
  dragIds: ReadonlySet<string> | null;
  getCell: GridProps['getCell'];
  cssOf: GridProps['cssOf'];
  dataVersion: number;
  highlighted: boolean;
  highlightLabel?: string;
  isActiveRow: boolean;
}

interface CellProps {
  row: number;
  i: number;
  col: GridColumn;
  left: number;
  sel: boolean;
  dragging: boolean;
  getCell: GridProps['getCell'];
  cssOf: GridProps['cssOf'];
  dataVersion: number;
}

const JUSTIFY: Record<string, string> = { left: 'flex-start', right: 'flex-end', center: 'center', justify: 'flex-start' };

/** memo per cell: horizontal scrolling / selection changes only re-render the cells that changed */
const GridCell = memo(function GridCell(p: CellProps) {
  const cell = safeGet(p.getCell, p.row, p.col);
  let css: CellCss;
  try {
    css = p.cssOf(cell?.styleId);
  } catch {
    css = { box: {}, vAlign: 'bottom', wrap: false, font: HEADER_FONT };
  }
  const r = renderCell(cell, css);
  const cls =
    'g-cell' +
    (p.sel ? ' is-sel' : '') +
    (p.dragging ? ' is-dragging' : '') +
    (r.formulaText ? ' is-ftext' : '') +
    (p.col.hidden ? ' is-hiddencol' : '');
  const style: CSSProperties = {
    left: p.left,
    width: p.col.width,
    justifyContent: JUSTIFY[r.align],
    textAlign: r.align,
    ...(css.box as CSSProperties),
  };
  if (r.color && !r.formulaText) style.color = r.color;
  return (
    <div role="gridcell" aria-colindex={p.i + 1} className={cls} data-r={p.row} data-c={p.col.id} data-testid={`grid-cell-${p.row}-${p.i}`} style={style}>
      {css.border && <div className="g-bd" style={css.border as CSSProperties} />}
      {r.text}
    </div>
  );
});

const Row = memo(function Row(p: RowProps) {
  const renderOne = (i: number, left: number) => {
    const col = p.columns[i];
    return (
      <GridCell
        key={col.id}
        row={p.row}
        i={i}
        col={col}
        left={left}
        sel={p.selected.has(col.id)}
        dragging={!!p.dragIds?.has(col.id)}
        getCell={p.getCell}
        cssOf={p.cssOf}
        dataVersion={p.dataVersion}
      />
    );
  };
  return (
    <div
      className={`g-row${p.highlighted ? ' is-hl' : ''}`}
      role="row"
      aria-rowindex={p.idx + 2}
      style={{ transform: `translateY(${p.top}px)`, width: ROWHDR_W + p.totalW, height: p.height }}
    >
      {p.vCols.map((vc) => renderOne(vc.index, vc.start))}
      <div className="g-pin" style={{ width: ROWHDR_W + p.frozenW, height: p.height }}>
        <div
          className={`g-rowhdr${p.isActiveRow ? ' is-activerow' : ''}`}
          style={{ width: ROWHDR_W }}
          data-testid={`grid-rowhdr-${p.row}`}
          title={p.highlighted && p.highlightLabel ? `${p.row}行目: ${p.highlightLabel}` : undefined}
        >
          {p.row}
        </div>
        {Array.from({ length: p.frozen }, (_, i) => renderOne(i, ROWHDR_W + p.starts[i]))}
      </div>
    </div>
  );
});
