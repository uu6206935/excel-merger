/**
 * Small shared helpers of the doc module (coordinates, lookups, formula-cache access).
 */
import { MAX_COL, MAX_ROW, numToCol, parseAreaText } from '../formula/excelRef';
import type { Area } from '../formula/excelRef';
import type { Cell, CellType, DocColumn, ExternalFormulaCell, MergedDoc, RangeBox, XStyle } from '../types';
import { compile, compileWithHint } from './compiled';
import type { Compiled } from './compiled';

export const lc = (s: string): string => s.toLowerCase();

/** "C3" */
export const addr = (col: number, row: number): string => `${numToCol(col)}${row}`;


export function boxToArea(b: RangeBox): Area {
  const single = b.r1 === b.r2 && b.c1 === b.c2;
  return { kind: single ? 'cell' : 'area', c1: b.c1, r1: b.r1, c2: b.c2, r2: b.r2, ac1: false, ar1: false, ac2: false, ar2: false };
}

/** Area -> normalized RangeBox (whole columns -> rows 1..MAX_ROW, whole rows -> cols 1..MAX_COL). */
export function areaToBox(a: Area): RangeBox {
  if (a.kind === 'cols') return { r1: 1, r2: MAX_ROW, c1: Math.min(a.c1, a.c2), c2: Math.max(a.c1, a.c2) };
  if (a.kind === 'rows') return { r1: Math.min(a.r1, a.r2), r2: Math.max(a.r1, a.r2), c1: 1, c2: MAX_COL };
  return { r1: Math.min(a.r1, a.r2), r2: Math.max(a.r1, a.r2), c1: Math.min(a.c1, a.c2), c2: Math.max(a.c1, a.c2) };
}

export function boxText(b: RangeBox): string {
  const s = addr(b.c1, b.r1);
  return b.r1 === b.r2 && b.c1 === b.c2 ? s : `${s}:${addr(b.c2, b.r2)}`;
}

/** Parse a plain A1 range ("C3:E3", "$A$1:$T$20", "C:C") as a normalized box; null when not a range. */
export function parseBox(text: string): RangeBox | null {
  const a = parseAreaText(text.replace(/^.*!/, '').replace(/\$/g, '').trim());
  return a ? areaToBox(a) : null;
}

// ---------------------------------------------------------------------------------------------
// column lookups (cached per immutable column array)
// ---------------------------------------------------------------------------------------------
const idIndexCache = new WeakMap<DocColumn[], Map<string, number>>();

/** 0-based index of the column with this id in the CURRENT order, -1 when unknown. */
export function colIndexOf(doc: MergedDoc, id: string): number {
  let m = idIndexCache.get(doc.columns);
  if (!m) {
    m = new Map();
    doc.columns.forEach((c, i) => m!.set(c.id, i));
    idIndexCache.set(doc.columns, m);
  }
  return m.get(id) ?? -1;
}

// ---------------------------------------------------------------------------------------------
// compiled-formula cache.
// Keyed by the immutable ARRAY that holds formulas (a column's `cells`, or doc.externalFormulas) and aligned by
// index; every entry remembers the formula text it was built from and is only used while that text is unchanged.
// (A per-cell WeakMap was measured: with 50,000 formulas its insertions + GC dominated a column move.)
// ---------------------------------------------------------------------------------------------
export interface CompiledCol {
  f: Array<string | undefined>;
  c: Array<Compiled | undefined>;
}
const colCache = new WeakMap<object, CompiledCol>();

/** The compiled-form array of a formula-holding array (created empty on first use). */
export function compiledCol(holder: object): CompiledCol {
  let cc = colCache.get(holder);
  if (!cc) {
    cc = { f: [], c: [] };
    colCache.set(holder, cc);
  }
  return cc;
}

/** Compiled form of formula `f` at index `i` (fill-down `hint` = compiled form `dr` rows above). */
export function compiledAt(cc: CompiledCol, i: number, f: string, hint: Compiled | null = null, dr = 0): Compiled {
  if (cc.f[i] === f) return cc.c[i]!;
  const c = hint && dr !== 0 ? compileWithHint(f, hint, dr, 0) : compile(f);
  cc.f[i] = f;
  cc.c[i] = c;
  return c;
}

/** Attach compiled forms to a (new) holder array. */
export function setCompiledCol(holder: object, cc: CompiledCol): void {
  colCache.set(holder, cc);
}

/** A new holder array derived from `from` (same indices) starts with a copy of its compiled forms. */
export function carryCompiled(from: object, to: object): CompiledCol {
  const cc = colCache.get(from);
  const out: CompiledCol = cc ? { f: cc.f.slice(), c: cc.c.slice() } : { f: [], c: [] };
  colCache.set(to, out);
  return out;
}

/** Compile every formula of the doc (merged sheet column-major with fill-down hints, then other sheets). */
export function precompile(doc: MergedDoc): void {
  for (const col of doc.columns) {
    let prev: Compiled | null = null;
    let prevRow = 0;
    const cells = col.cells;
    const cc = compiledCol(cells);
    for (let r = 0; r < cells.length; r++) {
      const cell = cells[r];
      if (!cell || cell.f === undefined) continue;
      prev = compiledAt(cc, r, cell.f, prev, r - prevRow);
      prevRow = r;
    }
  }
  const ec = compiledCol(doc.externalFormulas);
  let prev: Compiled | null = null;
  let pk = '';
  let prow = 0;
  doc.externalFormulas.forEach((e, i) => {
    const k = `${e.sheet}\u0000${e.col}`;
    prev = compiledAt(ec, i, e.f, k === pk ? prev : null, e.row - prow);
    pk = k;
    prow = e.row;
  });
}

// ---------------------------------------------------------------------------------------------
// cached results of formulas on the other sheets (for "convert to values"). Kept outside the contract objects
// (ExternalFormulaCell is exactly {sheet,row,col,f}) and keyed by the immutable ExternalFormulaCell object.
// ---------------------------------------------------------------------------------------------
export interface CachedValue {
  t: CellType;
  v: string | number | boolean | null;
  /** the formula is an array / spill master (range) — cannot be converted to a value cell by cell */
  arr?: string;
}
const extCache = new WeakMap<ExternalFormulaCell, CachedValue>();

export function cachedOf(e: ExternalFormulaCell): CachedValue | undefined {
  return extCache.get(e);
}

export function setCached(e: ExternalFormulaCell, v: CachedValue | undefined): void {
  if (v) extCache.set(e, v);
}

// ---------------------------------------------------------------------------------------------
// styles
// ---------------------------------------------------------------------------------------------
export function styleOf(doc: MergedDoc, xf: number | undefined): XStyle | undefined {
  return doc.styles[xf ?? 0];
}

/** The effective xf of a (possibly missing) cell: its own, else the column default. */
export function effectiveXf(col: DocColumn, cell: Cell | undefined): number | undefined {
  if (cell && cell.s !== undefined) return cell.s;
  return col.props.s;
}

/**
 * xf a NEW cell in `row` gets when something is typed into it: the row format (a row with
 * customFormat — the reader only reports RowProps.s then), else the column format. undefined = the column default.
 */
export function newCellXf(doc: MergedDoc, row: number): number | undefined {
  const s = doc.rowProps[row - 1]?.s;
  return s !== undefined && s > 0 ? s : undefined;
}

// ---------------------------------------------------------------------------------------------
// spill ranges
// ---------------------------------------------------------------------------------------------
/**
 * The cached values spilled by the dynamic-array master at (row, col) — the formula-less cells of its range `box`
 * other than the master — are removed (styles kept), exactly like the patcher's clearSpill. Columns are copied on
 * write; the input array is untouched.
 */
export function clearSpillValues(columns: DocColumn[], box: RangeBox, row: number, col: number): DocColumn[] {
  let out = columns;
  for (let c = box.c1; c <= box.c2; c++) {
    const dc = out[c - 1];
    if (!dc) continue;
    let cells = dc.cells;
    for (let r = box.r1; r <= box.r2; r++) {
      if (r === row && c === col) continue;
      const cell = cells[r - 1];
      if (!cell || cell.f !== undefined || cell.v === null) continue;
      if (cells === dc.cells) {
        cells = cells.slice();
        carryCompiled(dc.cells, cells);
      }
      cells[r - 1] = cell.s !== undefined ? { t: 'z', v: null, s: cell.s } : undefined;
    }
    if (cells !== dc.cells) {
      if (out === columns) out = columns.slice();
      out[c - 1] = { ...dc, cells };
    }
  }
  return out;
}
