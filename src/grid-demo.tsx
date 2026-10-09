/**
 * DEV-ONLY demo page for the grid (owner: grid G): http://127.0.0.1:5175/grid-demo.html
 * (`npx vite --port 5175 --strictPort`). Not part of the product build (vite builds index.html only).
 *
 * A generated sheet (10,000 data rows × 60 columns from 3 "files", Excel rows 2..10001, header row 1) with an
 * in-memory op history, so every grid callback can be exercised by tests/grid/e2e (Playwright).
 * URL params: ?rows=N&cols=M&frozen=K&readonly=1
 * window.__gridDemo exposes the state for tests (column order, widths, callback log, cell values).
 */
import { StrictMode, useCallback, useEffect, useMemo, useReducer, useRef, useState } from 'react';
import { createRoot } from 'react-dom/client';
import { Grid } from './components/grid/Grid';
import { FormulaBar } from './components/grid/FormulaBar';
import { cellEditText } from './components/grid/cellview';
import { colLetter, computeMove, isNoopMove } from './components/grid/colops';
import { EMPTY_SEL, type ColSelection } from './components/grid/selection';
import type { CellPos, CellView, EditingState, GridColumn, GridHandle, SourceInfo, XStyle } from './components/grid/types';
import { makeStyleCache } from './core/display/styles';
import { JA_BUILTIN_NUMFMT } from './core/display/numfmt';

// ---------------------------------------------------------------------------------------------
// demo data

const params = new URLSearchParams(location.search);
const ROW_COUNT = Math.max(0, Math.min(200000, Number(params.get('rows') ?? 10000) || 0));
const COL_COUNT = Math.max(20, Math.min(400, Number(params.get('cols') ?? 60) || 60));
const INITIAL_FROZEN = Math.max(0, Math.min(5, Number(params.get('frozen') ?? 2)));
const READ_ONLY = params.get('readonly') === '1';
/** ?gaps=1: rows 10, 20, 30 … are hidden (left out of `rows`) */
const GAPS = params.get('gaps') === '1';
/** ?tall=1: every 5th row is 40 px high (file row heights) */
const TALL = params.get('tall') === '1';
const FIRST_ROW = 2;
const tallRowHeight = (row: number) => (row % 5 === 2 ? 40 : undefined);

const SOURCES: SourceInfo[] = [
  { name: '学生名簿.xlsx（基準）', color: '#2563eb' },
  { name: '成績一覧.xlsx', color: '#16a34a' },
  { name: '連絡先.xlsx', color: '#d97706' },
];
const APP_SOURCE_COLOR = '#9ca3af';

const STYLES: XStyle[] = [
  /* 0 */ {},
  /* 1 ID      */ { bold: true, hAlign: 'center', border: { right: { style: 'thin', color: '#9ca3af' } } },
  /* 2 和暦     */ { numFmt: '[$-411]ggge"年"m"月"d"日"' },
  /* 3 date    */ { numFmt: JA_BUILTIN_NUMFMT[14] },
  /* 4 score   */ { numFmt: '0', hAlign: 'center' },
  /* 5 low     */ { numFmt: '0', hAlign: 'center', fill: '#FDE2E2', color: '#B91C1C', bold: true },
  /* 6 avg     */ { numFmt: '0.0', fill: '#EEF6EE', border: { left: { style: 'medium', color: '#16A34A' }, right: { style: 'medium', color: '#16A34A' } } },
  /* 7 yen     */ { numFmt: '"¥"#,##0;[Red]-"¥"#,##0' },
  /* 8 pct     */ { numFmt: '0.0%', italic: true },
  /* 9 mail    */ { color: '#1D4ED8', underline: true },
  /* 10 zip    */ { numFmt: '"〒"000-0000' },
  /* 11 memo   */ { wrap: true, vAlign: 'top', size: 9, color: '#555555' },
  /* 12 name   */ { bold: true },
  /* 13 box    */ { border: { top: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'double', color: '#DC2626' }, left: { style: 'thin' } }, hAlign: 'center' },
  /* 14 phone  */ { hAlign: 'left', font: 'Consolas' },
  /* 15 text @ */ { numFmt: '@' },
];
const cssOf = makeStyleCache(STYLES);

const FAMILY = ['佐藤', '鈴木', '高橋', '田中', '伊藤', '渡辺', '山本', '中村', '小林', '加藤', '吉田', '山田', '佐々木', '山口', '松本'];
const GIVEN = ['翔太', '陽菜', '蓮', '結衣', '大翔', 'さくら', '悠真', '美咲', '颯太', '葵', '湊', '凛', '樹', '芽依', '陸'];
const KANA_F = ['サトウ', 'スズキ', 'タカハシ', 'タナカ', 'イトウ', 'ワタナベ', 'ヤマモト', 'ナカムラ', 'コバヤシ', 'カトウ', 'ヨシダ', 'ヤマダ', 'ササキ', 'ヤマグチ', 'マツモト'];
const KANA_G = ['ショウタ', 'ヒナ', 'レン', 'ユイ', 'ヒロト', 'サクラ', 'ユウマ', 'ミサキ', 'ソウタ', 'アオイ', 'ミナト', 'リン', 'イツキ', 'メイ', 'リク'];
const UNIV = ['慶應義塾大学', '早稲田大学', '東京大学', '京都大学', '大阪大学', '一橋大学', '東北大学', '九州大学'];

function h(n: number): number {
  let x = (n ^ 0x9e3779b9) >>> 0;
  x = Math.imul(x ^ (x >>> 16), 0x85ebca6b) >>> 0;
  x = Math.imul(x ^ (x >>> 13), 0xc2b2ae35) >>> 0;
  return (x ^ (x >>> 16)) >>> 0;
}

/** r = 0-based data index; formulas use {colId} templates and {r} = Excel row */
type Gen = (r: number) => CellView;
interface ColDef extends GridColumn {
  gen: Gen;
}
const SUBJECTS = 24;

function buildColumns(): ColDef[] {
  const cols: ColDef[] = [];
  const add = (c: Omit<ColDef, 'id'> & { id?: string }) => cols.push({ ...c, id: c.id ?? `c${cols.length}` } as ColDef);
  add({ id: 'id', title: 'ID', width: 70, source: 0, keyRole: 'ID', gen: (r) => ({ value: 100001 + r, styleId: 1 }) });
  add({ id: 'name', title: '氏名', width: 110, source: 0, keyRole: '氏名', gen: (r) => ({ value: `${FAMILY[h(r) % 15]} ${GIVEN[h(r + 7) % 15]}`, styleId: 12, t: 's' }) });
  add({ title: '氏名カナ', width: 150, source: 0, gen: (r) => ({ value: `${KANA_F[h(r) % 15]} ${KANA_G[h(r + 7) % 15]}`, t: 's' }) });
  add({ id: 'univ', title: '大学', width: 120, source: 0, gen: (r) => ({ value: UNIV[h(r + 3) % UNIV.length], t: 's' }) });
  add({ id: 'stno', title: '学籍番号', width: 100, source: 0, gen: (r) => ({ value: `2${String(h(r + 11) % 10000000).padStart(7, '0')}`, styleId: 15, t: 's' }) });
  add({ title: '生年月日', width: 150, source: 0, gen: (r) => ({ value: 36526 + (h(r + 5) % 2200), styleId: 2, t: 'd' }) });
  add({ title: '入学日', width: 90, source: 0, gen: (r) => ({ value: 44652 + (h(r + 9) % 3) * 365, styleId: 3, t: 'd' }) });
  add({ title: '区分', width: 70, source: 0, gen: (r) => ({ value: ['一般', '推薦', 'AO'][h(r + 2) % 3], styleId: 13, t: 's' }) });
  const firstSubject = cols.length;
  for (let i = 0; i < SUBJECTS; i++) {
    add({
      title: `科目${i + 1}`,
      width: 64,
      source: 1,
      gen: (r) => {
        const v = 40 + (h(r * 131 + i * 7919) % 61);
        return { value: v, styleId: v < 50 ? 5 : 4 };
      },
    });
  }
  const firstId = cols[firstSubject].id;
  const lastId = cols[cols.length - 1].id;
  add({
    id: 'avg',
    title: '平均点',
    width: 80,
    source: 1,
    gen: (r) => {
      let s = 0;
      for (let i = 0; i < SUBJECTS; i++) s += 40 + (h(r * 131 + i * 7919) % 61);
      return { value: s / SUBJECTS, formula: `=AVERAGE({${firstId}}{r}:{${lastId}}{r})`, styleId: 6 };
    },
  });
  add({ title: '奨学金', width: 100, source: 1, gen: (r) => ({ value: (h(r + 21) % 7 === 0 ? -1 : 1) * (h(r + 13) % 40) * 5000, styleId: 7 }) });
  add({ id: 'ratio', title: '出席率', width: 80, source: 1, gen: (r) => ({ value: 0.6 + (h(r + 17) % 400) / 1000, styleId: 8 }) });
  add({ id: 'uncalc', title: '順位（未計算）', width: 150, source: 1, gen: () => ({ value: null, formula: '=RANK({avg}{r},{avg}$2:{avg}$10001)' }) });
  add({ title: 'メール', width: 210, source: 2, gen: (r) => ({ value: `student${100001 + r}@example.ac.jp`, styleId: 9, t: 's' }) });
  add({ title: '電話番号', width: 120, source: 2, gen: (r) => ({ value: `090-${String(h(r + 31) % 10000).padStart(4, '0')}-${String(h(r + 37) % 10000).padStart(4, '0')}`, styleId: 14, t: 's' }) });
  add({ title: '郵便番号', width: 90, source: 2, gen: (r) => ({ value: 1000000 + (h(r + 41) % 8999999), styleId: 10 }) });
  add({ title: '内部メモ', width: 8, source: 2, hidden: true, gen: () => ({ value: '非表示', t: 's' }) });
  add({ title: '住所', width: 200, source: 2, gen: (r) => ({ value: `東京都港区三田${(h(r + 43) % 5) + 1}-${(h(r + 47) % 40) + 1}-${(h(r + 53) % 20) + 1}`, t: 's' }) });
  add({ title: '備考', width: 180, source: 2, gen: (r) => ({ value: r % 5 === 0 ? '面談希望あり。第一志望は総合商社。' : null, styleId: 11, t: 's' }) });
  let k = 1;
  while (cols.length < COL_COUNT) {
    const n = k++;
    add({ title: `追加項目${n}`, width: 90, source: 2, gen: (r) => ({ value: (h(r * 17 + n * 101) % 1000) / 10 }) });
  }
  return cols;
}

const DEFS = buildColumns();
const DEF_BY_ID = new Map(DEFS.map((d) => [d.id, d]));
const ORIGINAL_WIDTH = new Map(DEFS.map((d) => [d.id, d.width]));

// ---------------------------------------------------------------------------------------------
// state + history

interface SheetState {
  columns: GridColumn[];
  /** user edits: key = `${colId}\u0000${ExcelRow}`; undefined value = cleared */
  edits: Map<string, CellView | null>;
  version: number;
}

type Op =
  | { t: 'move'; before: string[]; after: string[] }
  | { t: 'insert'; index: number; cols: GridColumn[] }
  | { t: 'delete'; removed: Array<{ index: number; col: GridColumn }> }
  | { t: 'edit'; key: string; before: CellView | null | undefined; after: CellView | null }
  | { t: 'width'; changes: Array<{ id: string; before: number; after: number }> };

interface Store {
  sheet: SheetState;
  past: Op[];
  future: Op[];
  live: Record<string, number>;
}

type Action = { type: 'run'; op: Op } | { type: 'undo' } | { type: 'redo' } | { type: 'live'; id: string; px: number | null };

function apply(s: SheetState, op: Op, dir: 1 | -1): SheetState {
  switch (op.t) {
    case 'move': {
      const byId = new Map(s.columns.map((c) => [c.id, c]));
      return { ...s, columns: (dir === 1 ? op.after : op.before).map((id) => byId.get(id)!), version: s.version + 1 };
    }
    case 'insert': {
      const cols = s.columns.slice();
      if (dir === 1) cols.splice(op.index, 0, ...op.cols);
      else cols.splice(op.index, op.cols.length);
      return { ...s, columns: cols, version: s.version + 1 };
    }
    case 'delete': {
      const cols = s.columns.slice();
      if (dir === 1) for (const r of [...op.removed].sort((a, b) => b.index - a.index)) cols.splice(r.index, 1);
      else for (const r of [...op.removed].sort((a, b) => a.index - b.index)) cols.splice(r.index, 0, r.col);
      return { ...s, columns: cols, version: s.version + 1 };
    }
    case 'edit': {
      const edits = new Map(s.edits);
      const v = dir === 1 ? op.after : op.before;
      if (v === undefined) edits.delete(op.key);
      else edits.set(op.key, v);
      return { ...s, edits, version: s.version + 1 };
    }
    case 'width': {
      const m = new Map(op.changes.map((c) => [c.id, dir === 1 ? c.after : c.before]));
      return { ...s, columns: s.columns.map((c) => (m.has(c.id) ? { ...c, width: m.get(c.id)! } : c)), version: s.version + 1 };
    }
  }
}

function reducer(st: Store, a: Action): Store {
  switch (a.type) {
    case 'run':
      return { ...st, sheet: apply(st.sheet, a.op, 1), past: [...st.past, a.op], future: [] };
    case 'undo': {
      const op = st.past[st.past.length - 1];
      if (!op) return st;
      return { ...st, sheet: apply(st.sheet, op, -1), past: st.past.slice(0, -1), future: [op, ...st.future] };
    }
    case 'redo': {
      const op = st.future[0];
      if (!op) return st;
      return { ...st, sheet: apply(st.sheet, op, 1), past: [...st.past, op], future: st.future.slice(1) };
    }
    case 'live': {
      const live = { ...st.live };
      if (a.px === null) delete live[a.id];
      else live[a.id] = a.px;
      return { ...st, live };
    }
  }
}

const ek = (colId: string, row: number) => `${colId}\u0000${row}`;

function renderFormula(tpl: string, row: number, letters: Map<string, string>): string {
  return tpl.replace(/\{([^}]+)\}/g, (_m, id: string) => (id === 'r' ? String(row) : (letters.get(id) ?? '#REF!')));
}

function parseInput(text: string, prev: CellView | undefined): CellView | null {
  if (text === '') return null;
  const styleId = prev?.styleId;
  if (text.startsWith('=')) return { value: null, formula: text, styleId };
  const prevIsText = prev?.t === 's' || typeof prev?.value === 'string';
  const num = Number(text.replace(/,/g, ''));
  if (!prevIsText && text.trim() !== '' && Number.isFinite(num)) return { value: num, styleId };
  return { value: text, styleId, t: 's' };
}

let newSeq = 1;

// ---------------------------------------------------------------------------------------------

interface CallLog {
  name: string;
  args: unknown[];
}

declare global {
  interface Window {
    __gridDemo?: {
      calls: CallLog[];
      order: () => string[];
      titles: () => string[];
      width: (id: string) => number | undefined;
      idOf: (title: string) => string | undefined;
      cellText: (row: number, id: string) => string;
      history: () => { past: number; future: number };
      grid: () => GridHandle | null;
    };
  }
}

function Demo() {
  const [store, dispatch] = useReducer(reducer, undefined, () => ({
    sheet: { columns: DEFS.map(({ gen: _g, ...c }) => c), edits: new Map(), version: 0 },
    past: [],
    future: [],
    live: {},
  }));
  const { sheet } = store;
  const [selection, setSelection] = useState<ColSelection>(EMPTY_SEL);
  const [active, setActive] = useState<CellPos | null>(ROW_COUNT ? { row: FIRST_ROW, col: DEFS[0].id } : null);
  const [frozenCols, setFrozenCols] = useState(INITIAL_FROZEN);
  const [editing, setEditing] = useState<EditingState | null>(null);
  const [jump, setJump] = useState('');
  const gridRef = useRef<GridHandle>(null);
  const callsRef = useRef<CallLog[]>([]);
  const storeRef = useRef(store);
  storeRef.current = store;

  const log = useCallback((name: string, ...args: unknown[]) => {
    callsRef.current.push({ name, args: JSON.parse(JSON.stringify(args)) });
  }, []);

  const rows = useMemo(() => Array.from({ length: ROW_COUNT }, (_, i) => i + FIRST_ROW).filter((r) => !GAPS || r % 10 !== 0), []);
  const highlightRows = useMemo(() => new Set(rows.filter((r) => r % 97 === 0)), [rows]);

  const columns = useMemo(
    () => (Object.keys(store.live).length ? sheet.columns.map((c) => (store.live[c.id] !== undefined ? { ...c, width: store.live[c.id] } : c)) : sheet.columns),
    [sheet.columns, store.live],
  );
  const letters = useMemo(() => new Map(sheet.columns.map((c, i) => [c.id, colLetter(i)])), [sheet.columns]);

  const getCell = useCallback(
    (row: number, col: GridColumn): CellView | undefined => {
      const e = sheet.edits.get(ek(col.id, row));
      if (e !== undefined) return e ?? undefined;
      const d = DEF_BY_ID.get(col.id);
      if (!d) return undefined;
      const v = d.gen(row - FIRST_ROW);
      return v.formula ? { ...v, formula: renderFormula(v.formula, row, letters) } : v;
    },
    [sheet.edits, letters],
  );

  const onMoveColumns = useCallback(
    (ids: string[], gap: number) => {
      log('onMoveColumns', ids, gap);
      const before = storeRef.current.sheet.columns.map((c) => c.id);
      if (isNoopMove(before, new Set(ids), gap)) return;
      dispatch({ type: 'run', op: { t: 'move', before, after: computeMove(before, new Set(ids), gap) } });
    },
    [log],
  );
  const onInsertColumns = useCallback(
    (index: number, count: number) => {
      log('onInsertColumns', index, count);
      const cols: GridColumn[] = Array.from({ length: count }, () => {
        const n = newSeq++;
        return { id: `new${n}`, title: `新しい列${n}`, width: 100, source: -1 };
      });
      dispatch({ type: 'run', op: { t: 'insert', index, cols } });
    },
    [log],
  );
  const onDeleteColumns = useCallback(
    (ids: string[]) => {
      log('onDeleteColumns', ids);
      const set = new Set(ids);
      const removed = storeRef.current.sheet.columns.map((col, index) => ({ col, index })).filter((x) => set.has(x.col.id));
      if (!removed.length) return;
      dispatch({ type: 'run', op: { t: 'delete', removed } });
      setSelection(EMPTY_SEL);
    },
    [log],
  );
  const onResizeColumn = useCallback(
    (id: string, width: number, commit: { from: number } | null) => {
      log('onResizeColumn', id, width, commit);
      if (!commit) {
        dispatch({ type: 'live', id, px: width });
        return;
      }
      dispatch({ type: 'live', id, px: null });
      if (commit.from !== width) dispatch({ type: 'run', op: { t: 'width', changes: [{ id, before: commit.from, after: width }] } });
    },
    [log],
  );
  const onResetColumnWidths = useCallback(
    (ids: string[]) => {
      log('onResetColumnWidths', ids);
      const cur = storeRef.current.sheet.columns;
      const changes = ids
        .map((id) => ({ id, before: cur.find((c) => c.id === id)?.width ?? 0, after: ORIGINAL_WIDTH.get(id) ?? 100 }))
        .filter((c) => c.before !== c.after);
      if (changes.length) dispatch({ type: 'run', op: { t: 'width', changes } });
    },
    [log],
  );
  const onCommitEdit = useCallback(
    (pos: CellPos, text: string) => {
      log('onCommitEdit', pos, text);
      const st = storeRef.current.sheet;
      const col = st.columns.find((c) => c.id === pos.col);
      if (!col) return;
      const key = ek(pos.col, pos.row);
      const before = st.edits.get(key);
      const d = DEF_BY_ID.get(pos.col);
      const prev = before !== undefined ? (before ?? undefined) : d?.gen(pos.row - FIRST_ROW);
      dispatch({ type: 'run', op: { t: 'edit', key, before, after: parseInput(text, prev) } });
    },
    [log],
  );
  const onUndo = useCallback(() => {
    log('onUndo');
    dispatch({ type: 'undo' });
  }, [log]);
  const onRedo = useCallback(() => {
    log('onRedo');
    dispatch({ type: 'redo' });
  }, [log]);
  const onSelectionChange = useCallback((sel: ColSelection) => setSelection(sel), []);
  const onActiveChange = useCallback((pos: CellPos) => setActive(pos), []);

  // test hooks
  useEffect(() => {
    window.__gridDemo = {
      calls: callsRef.current,
      order: () => storeRef.current.sheet.columns.map((c) => c.id),
      titles: () => storeRef.current.sheet.columns.map((c) => c.title),
      width: (id) => storeRef.current.sheet.columns.find((c) => c.id === id)?.width,
      idOf: (title) => storeRef.current.sheet.columns.find((c) => c.title === title)?.id,
      cellText: (row, id) => {
        const e = storeRef.current.sheet.edits.get(ek(id, row));
        if (e !== undefined) return e === null ? '' : String(e.formula ?? e.value ?? '');
        const v = DEF_BY_ID.get(id)?.gen(row - FIRST_ROW);
        return v ? String(v.value ?? '') : '';
      },
      history: () => ({ past: storeRef.current.past.length, future: storeRef.current.future.length }),
      grid: () => gridRef.current,
    };
  }, []);

  // global Ctrl+Z / Ctrl+Y when focus is outside the grid and outside text inputs (the App does the same)
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.defaultPrevented) return;
      const t = e.target as HTMLElement;
      if (t.closest('.g-wrap') || t.tagName === 'INPUT' || t.tagName === 'TEXTAREA') return;
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'z') {
        e.preventDefault();
        if (e.shiftKey) onRedo();
        else onUndo();
      } else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'y') {
        e.preventDefault();
        onRedo();
      }
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [onUndo, onRedo]);

  const activeIdx = active ? sheet.columns.findIndex((c) => c.id === active.col) : -1;
  const activeCell = active && activeIdx >= 0 ? getCell(active.row, sheet.columns[activeIdx]) : undefined;
  const fbValue = cellEditText(activeCell, cssOf(activeCell?.styleId));
  const address = active && activeIdx >= 0 ? `${colLetter(activeIdx)}${active.row}` : '';

  return (
    <div style={{ display: 'flex', flexDirection: 'column', height: '100%', gap: 8, padding: '10px 16px 12px', boxSizing: 'border-box' }}>
      <header style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
        <h1 style={{ margin: 0, fontSize: 16 }}>グリッド デモ</h1>
        <span style={{ color: '#6b7280', fontSize: 12 }}>
          {ROW_COUNT.toLocaleString()}行 × {sheet.columns.length}列（開発用）
        </span>
        <button data-testid="demo-undo" onClick={onUndo} disabled={!store.past.length}>
          元に戻す
        </button>
        <button data-testid="demo-redo" onClick={onRedo} disabled={!store.future.length}>
          やり直す
        </button>
        <label style={{ display: 'inline-flex', gap: 4, alignItems: 'center' }}>
          固定列
          <select data-testid="demo-frozen" value={frozenCols} onChange={(e) => setFrozenCols(Number(e.target.value))}>
            {[0, 1, 2, 3].map((v) => (
              <option key={v} value={v}>
                {v}
              </option>
            ))}
          </select>
        </label>
        <form
          onSubmit={(e) => {
            e.preventDefault();
            gridRef.current?.scrollToRow(Number(jump), { select: true });
            gridRef.current?.focus();
          }}
          style={{ display: 'inline-flex', gap: 4 }}
        >
          <input data-testid="demo-jump" value={jump} onChange={(e) => setJump(e.target.value)} placeholder="行番号" style={{ width: 70 }} />
          <button type="submit">行へ移動</button>
        </form>
        <span data-testid="demo-status" style={{ marginLeft: 'auto', color: '#6b7280', fontSize: 12 }}>
          履歴 {store.past.length} / やり直し {store.future.length} / 選択 {selection.cols.size}列
        </span>
      </header>
      <FormulaBar
        pos={active}
        address={address}
        value={fbValue}
        mirror={editing?.text ?? null}
        onCommitEdit={onCommitEdit}
        onDone={() => gridRef.current?.focus()}
        readOnly={READ_ONLY}
      />
      <div style={{ color: '#1d3f7a', background: '#eef4ff', border: '1px solid #c9dafc', borderRadius: 6, padding: '4px 10px', fontSize: 12 }}>
        列見出しをクリックで選択（Ctrl・Shiftで複数）→ ドラッグで移動。右クリックで挿入・削除。
      </div>
      <main style={{ flex: 1, minHeight: 0, border: '1px solid #d6d6d6', borderRadius: 6, overflow: 'hidden', background: '#fff' }}>
        <Grid
          ref={gridRef}
          columns={columns}
          sources={SOURCES}
          appSourceColor={APP_SOURCE_COLOR}
          rows={rows}
          headerRow={1}
          frozenCols={frozenCols}
          getCell={getCell}
          cssOf={cssOf}
          dataVersion={sheet.version}
          selection={selection}
          onSelectionChange={onSelectionChange}
          active={active}
          onActiveChange={onActiveChange}
          onMoveColumns={onMoveColumns}
          onResizeColumn={onResizeColumn}
          onCommitEdit={onCommitEdit}
          onDeleteColumns={onDeleteColumns}
          onInsertColumns={onInsertColumns}
          onResetColumnWidths={onResetColumnWidths}
          onUndo={onUndo}
          onRedo={onRedo}
          onEditingChange={setEditing}
          highlightRows={highlightRows}
          highlightLabel="未一致（他のファイルに該当なし）"
          rowHeight={TALL ? tallRowHeight : undefined}
          readOnly={READ_ONLY}
        />
      </main>
    </div>
  );
}

createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <Demo />
  </StrictMode>,
);
