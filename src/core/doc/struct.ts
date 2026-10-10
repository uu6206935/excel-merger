/**
 * Structural column operations on a MergedDoc with Excel semantics (excelRef is the single source of truth for
 * reference math). One call applies a SEQUENCE of StructOps (e.g. the block moves of excelRef.planMove) and
 * returns a new doc that shares every unchanged column / cell / formula object with the input.
 *
 * What follows an op (verified semantics, see docs/research/R2*.md):
 *  - merged-sheet formulas (host = merged sheet) and array ranges of array masters,
 *  - formulas on other sheets (host = their sheet), CF/DV rules of other sheets (formulas only),
 *  - defined names (isName: relative parts are offsets and never adjusted),
 *  - merged-sheet features: merges, tables (ListObjects), array ranges, autoFilter, frozen columns, hyperlinks,
 *    CF/DV rules (excelRef.adjustAnchoredRule),
 *  - column props / cell styles travel with their columns; inserted columns take "format from left"
 *    (cell xf per row + width; nothing at column A — Excel 365 via COM).
 *  - dynamic-array (spill) ranges cut through by an op (insert inside, delete / move part of it) collapse to the master
 *    cell and drop the spilled cached values (Excel re-spills on load; the package patcher does the same).
 * Refusals (Japanese reasons) follow Excel: part of a merged area / legacy array formula, moving columns into /
 * out of / partly over a table (checked per block op on the state before that op).
 */
import { adjustAnchoredRule, adjustArea, adjustFormula, adjustSqref, movePerm, parseAreaText, renderArea, tableMoveKind, MAX_COL } from '../formula/excelRef';
import type { AnchoredRuleOptions, Area, StructOp, TableRef } from '../formula/excelRef';
import type { Cell, DefinedName, DocColumn, ExternalFormulaCell, FeatureRule, MergedDoc, RangeBox, SheetFeatures, TableInfo } from '../types';
import { adjustCompiled } from './compiled';
import type { AdjustMemo, Compiled } from './compiled';
import { InsertStyles } from './insertStyle';
import { boxRef, excelNewColumnNames, tableBox, tableRefs } from './tables';
import { addr, areaToBox, boxText, boxToArea, cachedOf, carryCompiled, clearSpillValues, compiledAt, compiledCol, lc, setCached, setCompiledCol } from './util';
import type { CompiledCol } from './util';

/** collapsed: spill ranges (A1 text, as before the op) that an op cut through (see collapseSpills) */
export type StructOutcome = { ok: true; doc: MergedDoc; refErrors: number; collapsed: string[] } | { ok: false; reason: string };

// ---------------------------------------------------------------------------------------------
// refusals
// ---------------------------------------------------------------------------------------------
function isDynamicArray(columns: DocColumn[], a: RangeBox): boolean {
  const master = columns[a.c1 - 1]?.cells[a.r1 - 1];
  return !!master?.dyn;
}

/** Why Excel refuses a column MOVE because of the merged area `m` (splitting it / landing inside it), or null. */
export function mergeRefusal(m: RangeBox, op: StructOp): string | null {
  if (op.type !== 'move' || op.axis !== 'col' || m.c1 === m.c2) return null;
  const s1 = op.from;
  const s2 = op.from + op.count - 1;
  const overlap = s1 <= m.c2 && s2 >= m.c1;
  const covers = s1 <= m.c1 && m.c2 <= s2;
  if (overlap && !covers) return `結合されたセル（${boxText(m)}）の一部の列だけを移動することはできません。結合を解除するか、結合範囲の列をまとめて移動してください。`;
  if (!overlap && op.to > m.c1 && op.to <= m.c2) return `結合されたセル（${boxText(m)}）の途中には列を移動できません。`;
  return null;
}

/**
 * Merged areas of `doc` (original objects) that block the op sequence (each checked at its position before each
 * op, like refusalFor).
 */
export function blockingMerges(doc: MergedDoc, ops: readonly StructOp[], candidates: readonly RangeBox[] = doc.features.merges): RangeBox[] {
  const out: RangeBox[] = [];
  for (const m of candidates) {
    let b: RangeBox | null = m;
    for (const op of ops) {
      if (!b) break;
      if (mergeRefusal(b, op)) {
        out.push(m);
        break;
      }
      b = adjustBox(b, op);
    }
  }
  return out;
}

/** Why Excel would refuse `op` on the current state, or null. */
export function refusalOf(columns: DocColumn[], f: SheetFeatures, op: StructOp): string | null {
  const n = columns.length;
  if (op.axis !== 'col') return '行の挿入・削除・移動には対応していません。';
  if (op.type === 'insert') {
    if (op.at < 1 || op.at > n + 1 || op.count < 1) return '挿入する位置が正しくありません。';
    if (n + op.count > MAX_COL) return `列が多すぎるため挿入できません（Excel の上限は ${MAX_COL.toLocaleString()} 列です）。`;
    for (const a of f.arrays) {
      if (isDynamicArray(columns, a)) continue;
      if (op.at > a.c1 && op.at <= a.c2) return `配列数式（${boxText(a)}）の途中には列を挿入できません。（Excel: 配列の一部を変更することはできません）`;
    }
    return null;
  }
  if (op.type === 'delete') {
    const s1 = op.at;
    const s2 = op.at + op.count - 1;
    if (s1 < 1 || s2 > n || op.count < 1) return '削除する列が正しくありません。';
    for (const a of f.arrays) {
      if (isDynamicArray(columns, a)) continue;
      const overlap = s1 <= a.c2 && s2 >= a.c1;
      const covers = s1 <= a.c1 && a.c2 <= s2;
      if (overlap && !covers) return `配列数式（${boxText(a)}）の一部の列だけを削除することはできません。（Excel: 配列の一部を変更することはできません）`;
    }
    return null;
  }
  // move
  const s1 = op.from;
  const s2 = op.from + op.count - 1;
  const to = op.to;
  if (s1 < 1 || s2 > n || op.count < 1 || to < 1 || to > n + 1) return '移動する列または移動先が正しくありません。';
  for (const m of f.merges) {
    const why = mergeRefusal(m, op);
    if (why) return why;
  }
  for (const a of f.arrays) {
    if (a.c1 === a.c2 || isDynamicArray(columns, a)) continue;
    const overlap = s1 <= a.c2 && s2 >= a.c1;
    const covers = s1 <= a.c1 && a.c2 <= s2;
    if (overlap && !covers) return `配列数式（${boxText(a)}）の一部の列だけを移動することはできません。（Excel: 配列の一部を変更することはできません）`;
    if (!overlap && to > a.c1 && to <= a.c2) return `配列数式（${boxText(a)}）の途中には列を移動できません。（Excel: 配列の一部を変更することはできません）`;
  }
  // Excel's rule is excelRef.tableMoveKind (verified, tests/formula/tableMove.test.ts): a block inside the table (and not
  // covering it) may land anywhere in [first column, last column + 1] and only reorders the table columns — the gap
  // right after the last column included; a block from outside must not land strictly inside; a block partly over the
  // table is refused; a block covering the table moves it.
  for (const t of f.tables) {
    const b = tableBox(t);
    if (!b) continue;
    if (tableMoveKind(b.c1, b.c2, op) !== 'refused') continue;
    const overlap = s1 <= b.c2 && s2 >= b.c1;
    const inside = b.c1 <= s1 && s2 <= b.c2;
    if (inside) return `テーブル「${t.name}」の列をテーブルの外へ移動することはできません。（Excel と同じ制限です）`;
    if (overlap) return `テーブル「${t.name}」の一部の列とテーブル外の列を一緒に移動することはできません。（Excel と同じ制限です）`;
    return `テーブル「${t.name}」の中へ、テーブル外の列を移動することはできません。（Excel と同じ制限です）`;
  }
  return null;
}

// ---------------------------------------------------------------------------------------------
// spill ranges
// ---------------------------------------------------------------------------------------------
/** The op cuts through the multi-cell range `a` (insert inside it, delete / move part of it, a move landing inside). */
function cutsRange(op: StructOp, a: RangeBox): boolean {
  if (op.axis !== 'col' || (a.c1 === a.c2 && a.r1 === a.r2)) return false;
  const lo = a.c1;
  const hi = a.c2;
  if (op.type === 'insert') return op.at > lo && op.at <= hi;
  if (op.type === 'delete') return !(op.at + op.count - 1 < lo || op.at > hi);
  const s = op.from;
  const e = op.from + op.count - 1;
  const inter = !(e < lo || s > hi);
  const covers = s <= lo && e >= hi;
  return (inter && !covers) || (!inter && op.to > lo && op.to <= hi);
}

interface Collapse {
  columns: DocColumn[];
  features: SheetFeatures;
  /** the masters whose spill was collapsed (column id + row) */
  masters: Array<{ id: string; row: number }>;
  texts: string[];
}

/**
 * Dynamic-array (spill) ranges: the cells below / beside the master only cache values that Excel recomputes on
 * load. When `op` cuts through such a range those values must not travel as constants (they would be stale
 * constants in the output): they are cleared and the range collapses to the master cell — exactly what the package
 * patcher does on replay (Engine.collapseSpills; Excel re-spills by itself). Legacy (CSE) arrays are refused instead.
 */
function collapseSpills(columns: DocColumn[], f: SheetFeatures, op: StructOp): Collapse | null {
  let cols = columns;
  let arrays = f.arrays;
  const masters: Collapse['masters'] = [];
  const texts: string[] = [];
  f.arrays.forEach((a, i) => {
    if (!isDynamicArray(columns, a) || !cutsRange(op, a)) return;
    cols = clearSpillValues(cols, a, a.r1, a.c1);
    if (arrays === f.arrays) arrays = arrays.slice();
    arrays[i] = { r1: a.r1, c1: a.c1, r2: a.r1, c2: a.c1 };
    masters.push({ id: columns[a.c1 - 1].id, row: a.r1 });
    texts.push(boxText(a));
  });
  return masters.length ? { columns: cols, features: { ...f, arrays }, masters, texts } : null;
}

/** Warning for the spill ranges an op cut through (OpResult.warnings). */
export function collapseWarnings(collapsed: readonly string[]): string[] {
  if (!collapsed.length) return [];
  const list = collapsed.slice(0, 3).join('、') + (collapsed.length > 3 ? ` ほか ${collapsed.length - 3} 個` : '');
  return [`スピル範囲（${list}）の途中を変更したため、スピルされていた値はいったん空になります。Excel で開くと数式から再計算されます。`];
}

/** The masters' own `arr` text follows the collapse: just their address (adjustColumns moved the old range along). */
function pinMasterRanges(columns: DocColumn[], masters: Collapse['masters']): DocColumn[] {
  let out = columns;
  const index = new Map(columns.map((c, i) => [c.id, i]));
  for (const { id, row } of masters) {
    const i = index.get(id);
    if (i === undefined) continue;
    const col = out[i];
    const cell = col.cells[row - 1];
    const want = addr(i + 1, row);
    if (!cell || cell.arr === want) continue;
    const cells = col.cells.slice();
    carryCompiled(col.cells, cells);
    cells[row - 1] = { ...cell, arr: want };
    if (out === columns) out = columns.slice();
    out[i] = { ...col, cells };
  }
  return out;
}

// ---------------------------------------------------------------------------------------------
// columns
// ---------------------------------------------------------------------------------------------
function permuteColumns(cols: DocColumn[], op: Extract<StructOp, { type: 'move' }>): DocColumn[] {
  const p = movePerm(op.from, op.count, op.to);
  const out = new Array<DocColumn>(cols.length);
  for (let i = 1; i <= cols.length; i++) out[p(i) - 1] = cols[i - 1];
  return out;
}

/**
 * New columns inserted between `left` (undefined at column A) and `right` (the column pushed right): Excel
 * "format same as left" — per row the left cell's xf with only the border sides the right cell shares
 * (insertStyle.ts), and the left column's properties (width, column style, outline level). NOT `hidden`: Excel inserts a
 * VISIBLE column next to a hidden one (measured in Excel 365: format from left or from right, left column hidden, hidden and
 * grouped, inside a collapsed group: the new column is visible every time; tests/oracle/consistency/regress.test.ts).
 */
export function makeInsertedColumns(left: DocColumn | undefined, right: DocColumn | undefined, ids: string[], rowCount: number, styles: InsertStyles): DocColumn[] {
  const cells = new Array<Cell | undefined>(rowCount).fill(undefined);
  const props: DocColumn['props'] = {};
  if (left) {
    for (let r = 0; r < rowCount; r++) {
      const c = left.cells[r];
      if (!c || c.s === undefined || c.s === 0) continue;
      const s = styles.inserted(c.s, right?.cells[r]?.s ?? 0);
      if (s) cells[r] = { t: 'z', v: null, s };
    }
    const lp = left.props;
    if (lp.width !== undefined) props.width = lp.width;
    if (lp.customWidth !== undefined) props.customWidth = lp.customWidth;
    if (lp.s !== undefined) props.s = lp.s;
    if (lp.outlineLevel !== undefined) props.outlineLevel = lp.outlineLevel;
  }
  // the new columns share one (immutable) cells array
  return ids.map((id) => ({ id, origin: { kind: 'new' }, props: { ...props }, cells }));
}

// ---------------------------------------------------------------------------------------------
// features
// ---------------------------------------------------------------------------------------------
function mapChanged<T>(arr: T[], fn: (x: T) => T | T[] | null): T[] {
  let changed = false;
  const out: T[] = [];
  for (const x of arr) {
    const y = fn(x);
    if (y === x) {
      out.push(x);
      continue;
    }
    changed = true;
    if (y === null) continue;
    if (Array.isArray(y)) out.push(...y);
    else out.push(y);
  }
  return changed ? out : arr;
}

function adjustBox(b: RangeBox, op: StructOp): RangeBox | null {
  const a = boxToArea(b);
  const na = adjustArea(a, op);
  if (!na) return null;
  if (na === a) return b;
  const nb = areaToBox(na);
  return nb.r1 === b.r1 && nb.r2 === b.r2 && nb.c1 === b.c1 && nb.c2 === b.c2 ? b : nb;
}

function adjustRangeText(text: string, op: StructOp): string | null {
  const a = parseAreaText(text.replace(/\$/g, ''));
  if (!a) return text;
  const na = adjustArea(a, op);
  if (!na) return null;
  return na === a ? text : renderArea(na);
}

function freezeOp(fr: SheetFeatures['freeze'], op: StructOp): SheetFeatures['freeze'] {
  if (!fr || !fr.cols) return fr;
  // a delete removes the frozen columns it hits — except when it removes ALL of them: then ONE frozen column stays, whatever k was
  // (Excel golden g-plan-ACE-F: k = 1 stays 1; measured in Excel by the copied-sheets oracle, round 9: k = 2, 3, 4 -> 1; the doc used to keep k)
  const cut = (k: number, at: number, cnt: number) => (at <= 1 && at + cnt - 1 >= k ? 1 : k - Math.max(0, Math.min(k, at + cnt - 1) - at + 1));
  let k = fr.cols;
  if (op.type === 'insert') {
    if (op.at <= k) k += op.count;
  } else if (op.type === 'delete') {
    k = cut(k, op.at, op.count);
  } else {
    // R2b R10: only the insert and delete steps of "insert, cut/paste, delete" count
    const e = op.from + op.count - 1;
    if (op.to <= k) k += op.count;
    const p = op.to > e ? op.from : op.from + op.count;
    k = cut(k, p, op.count);
  }
  return k === fr.cols ? fr : { ...fr, cols: k };
}

/** sqref text -> plain cell areas (null when some part is not understood). */
export function sqrefAreas(sqref: string): Area[] | null {
  const areas: Area[] = [];
  for (const part of sqref.split(/\s+/)) {
    if (!part) continue;
    const a = parseAreaText(part.replace(/\$/g, ''));
    if (!a) return null;
    const b = areaToBox(a);
    areas.push({ kind: b.r1 === b.r2 && b.c1 === b.c2 ? 'cell' : 'area', r1: b.r1, r2: b.r2, c1: b.c1, c2: b.c2, ac1: false, ar1: false, ac2: false, ar2: false });
  }
  return areas;
}

/**
 * excelRef options for a rule of the merged sheet: its kind and xlsx type. excelRef decides (measured with Excel):
 * every CF type and every validation grows over a column inserted right after its area; on a move, cellIs /
 * expression / ... and validations spread onto a block landing inside / right after an area, the statistical CF
 * types (excelRef.NO_SPREAD_CF_TYPES: colorScale, dataBar, iconSet, top10, uniqueValues, duplicateValues,
 * aboveAverage) do not. A rule without a type spreads.
 */
export function ruleOptions(rule: FeatureRule): AnchoredRuleOptions {
  return rule.type !== undefined ? { kind: rule.kind, type: rule.type } : { kind: rule.kind };
}

/**
 * The CF rules that share a <conditionalFormatting> block with each rule: same kind 'cf', same storage (classic / x14
 * ext) and the same sqref. Excel keeps ONE range object per block (every rule of a block is fragmented at the same
 * places; on a move the destination cells that joined the block leave it again when ANY rule of the block refers to
 * the moved cells / the destination), so excelRef.adjustAnchoredRule needs the other rules' formulas (`siblings`,
 * same slot convention as FeatureRule.formulas). Only rules that are not alone in their block have an entry;
 * data validations never share a block. The reader does not keep block identity: equal sqref is the proxy.
 */
export function ruleSiblings(rules: readonly FeatureRule[]): Map<FeatureRule, string[][]> {
  const out = new Map<FeatureRule, string[][]>();
  if (rules.length < 2) return out;
  const groups = new Map<CfBlock, FeatureRule[]>();
  const proxies = new Map<string, CfBlock>();
  for (const r of rules) {
    if (r.kind !== 'cf') continue;
    const b = blockOfRule(r, proxies);
    const g = groups.get(b);
    if (g) g.push(r);
    else groups.set(b, [r]);
  }
  for (const g of groups.values()) {
    if (g.length < 2) continue;
    for (const r of g) out.set(r, g.filter((x) => x !== r).map((x) => x.formulas));
  }
  return out;
}

/**
 * Identity of the CF block (one <conditionalFormatting> = one Excel range object) a rule lives in. The reader does not
 * report blocks, so the rules of the file are grouped by storage + area when the doc first needs them; from then on the
 * identity FOLLOWS the ops exactly like the package patcher's regrouping (patch/rules.ts adjustCfBlock): the fragments of
 * one block that end up with the same area form the new block, while rules that merely coincide in area later (fragments
 * of two different blocks that land on the same cells) stay separate blocks — so they do not veto each other's spread.
 * `sqref` = the block's area, canonical (all rules of a block share it). A rule object is immutable, so the WeakMap
 * entry is the rule's identity in every snapshot that holds it (undo / redo).
 */
interface CfBlock {
  sqref: string;
}
const cfBlocks = new WeakMap<FeatureRule, CfBlock>();

/** Canonical text of an sqref (areas in order, `$` and spacing ignored); the trimmed text when it is not understood. */
function canonSqref(sqref: string): string {
  const areas = sqrefAreas(sqref);
  return areas ? areas.map((a) => `${a.r1}:${a.c1}:${a.r2}:${a.c2}`).join(' ') : sqref.replace(/\$/g, '').trim().split(/\s+/).join(' ');
}

function blockOfRule(r: FeatureRule, proxies?: Map<string, CfBlock>): CfBlock {
  let b = cfBlocks.get(r);
  if (b) return b;
  const sqref = canonSqref(r.sqref);
  const key = `${r.ext ? 'x' : 'c'}|${sqref}`;
  b = proxies?.get(key);
  if (!b) {
    b = { sqref };
    proxies?.set(key, b);
  }
  cfBlocks.set(r, b);
  return b;
}

/**
 * The fragments `res` that `adjustRule` made of `rule` (`rule` itself = unchanged): each joins the block (old block,
 * fragment area) — the old block itself when the area did not change. `memo` holds the blocks made during one op.
 */
function trackBlocks(rule: FeatureRule, res: FeatureRule[] | FeatureRule, memo: Map<CfBlock, Map<string, CfBlock>>): void {
  if (rule.kind !== 'cf' || res === rule) return;
  const old = blockOfRule(rule);
  for (const fr of Array.isArray(res) ? res : [res]) {
    if (fr === rule) continue;
    const sqref = canonSqref(fr.sqref);
    if (sqref === old.sqref) {
      cfBlocks.set(fr, old);
      continue;
    }
    let m = memo.get(old);
    if (!m) memo.set(old, (m = new Map()));
    let nb = m.get(sqref);
    if (!nb) m.set(sqref, (nb = { sqref }));
    cfBlocks.set(fr, nb);
  }
}

/**
 * CF / DV rule on the merged sheet under one op. FeatureRule formulas and excelRef.AnchoredRule formulas use the
 * same convention (relative to the top-left of the sqref BOUNDING BOX), so no re-anchoring is needed here.
 * `siblings`: the formulas of the other rules of the same CF block (ruleSiblings).
 * Returns the rule object itself when nothing changed, else the resulting fragments ([] = every cell deleted).
 */
export function adjustRule(rule: FeatureRule, op: StructOp, sheet: string, siblings?: string[][]): FeatureRule[] | FeatureRule {
  const areas = sqrefAreas(rule.sqref);
  if (!areas || !areas.length) return rule; // not understood: leave untouched
  const opts = ruleOptions(rule);
  if (siblings?.length) opts.siblings = siblings;
  const res = adjustAnchoredRule({ sqref: areas, formulas: rule.formulas }, op, sheet, opts);
  const out: FeatureRule[] = res.map((r) => ({ ...rule, sqref: r.sqref.map((a) => renderArea(a)).join(' '), formulas: r.formulas }));
  if (out.length === 1 && out[0].sqref === normalizedSqref(areas) && out[0].formulas.every((f, i) => f === rule.formulas[i])) return rule;
  return out;
}

const normalizedSqref = (areas: Area[]): string => areas.map((a) => renderArea(a)).join(' ');

interface HeaderWrite {
  id: string;
  row: number;
  text: string;
}

/** Tables after `op`. `onHeader(pos, row, name)` receives the header cells Excel writes for new table columns. */
function tablesOp(tables: TableInfo[], op: StructOp, onHeader?: (pos: number, row: number, name: string) => void): TableInfo[] {
  return mapChanged(tables, (t) => {
    const b = tableBox(t);
    if (!b) return t;
    const { c1: t1, c2: t2 } = b;
    if (op.type === 'move') {
      const p = movePerm(op.from, op.count, op.to);
      const s1 = op.from;
      const s2 = op.from + op.count - 1;
      const inside = t1 <= s1 && s2 <= t2 && !(s1 <= t1 && t2 <= s2);
      if (inside) {
        // columns reordered inside the table; the table range itself stays
        const cols = new Array<string>(t.columns.length);
        for (let c = t1; c <= t2; c++) cols[p(c) - t1] = t.columns[c - t1];
        return cols.every((x, i) => x === t.columns[i]) ? t : { ...t, columns: cols };
      }
      const n1 = p(t1);
      const n2 = p(t2);
      if (n1 === t1 && n2 === t2) return t;
      return { ...t, ref: boxRef({ ...b, c1: Math.min(n1, n2), c2: Math.max(n1, n2) }) };
    }
    if (op.type === 'insert') {
      if (op.at <= t1) return { ...t, ref: boxRef({ ...b, c1: t1 + op.count, c2: t2 + op.count }) };
      if (op.at > t2) return t;
      const idx = op.at - t1;
      const names = excelNewColumnNames(t.columns, t.columns[idx - 1] ?? null, op.count);
      const columns = [...t.columns.slice(0, idx), ...names, ...t.columns.slice(idx)];
      if (t.headerRow && onHeader) names.forEach((name, k) => onHeader(op.at + k, b.r1, name));
      return { ...t, ref: boxRef({ ...b, c2: t2 + op.count }), columns };
    }
    // delete
    const s1 = op.at;
    const s2 = op.at + op.count - 1;
    if (s2 < t1) return { ...t, ref: boxRef({ ...b, c1: t1 - op.count, c2: t2 - op.count }) };
    if (s1 > t2) return t;
    if (s1 <= t1 && t2 <= s2) return null; // whole table deleted
    const columns = t.columns.filter((_, i) => t1 + i < s1 || t1 + i > s2);
    const nc1 = t1 < s1 ? t1 : s1;
    const nc2 = nc1 + columns.length - 1;
    return { ...t, ref: boxRef({ ...b, c1: nc1, c2: nc2 }), columns };
  });
}

function featuresOp(f: SheetFeatures, op: StructOp, sheet: string, columnsAfter: DocColumn[], writes: HeaderWrite[]): SheetFeatures {
  const merges = mapChanged(f.merges, (m) => {
    const nb = adjustBox(m, op);
    if (!nb) return null;
    if (nb.r1 === nb.r2 && nb.c1 === nb.c2) return null; // a merge of one cell is no merge
    return nb;
  });
  const arrays = mapChanged(f.arrays, (a) => adjustBox(a, op));
  const tables = tablesOp(f.tables, op, (pos, row, text) => writes.push({ id: columnsAfter[pos - 1]?.id ?? '', row, text }));
  const sib = ruleSiblings(f.rules);
  const blockMemo = new Map<CfBlock, Map<string, CfBlock>>();
  const rules = mapChanged(f.rules, (r) => {
    const res = adjustRule(r, op, sheet, sib.get(r));
    trackBlocks(r, res, blockMemo);
    return res;
  });
  const hyperlinks = mapChanged(f.hyperlinks, (h) => {
    const ref = adjustSqref(h.ref, op);
    if (!ref) return null;
    return ref === h.ref ? h : { ...h, ref };
  });
  let autoFilter = f.autoFilter;
  if (autoFilter) autoFilter = adjustRangeText(autoFilter, op) ?? undefined;
  const freeze = freezeOp(f.freeze, op);
  if (
    merges === f.merges && arrays === f.arrays && tables === f.tables && rules === f.rules &&
    hyperlinks === f.hyperlinks && autoFilter === f.autoFilter && freeze === f.freeze
  ) return f;
  const out: SheetFeatures = { ...f, merges, arrays, tables, rules, hyperlinks };
  if (autoFilter === undefined) delete out.autoFilter;
  else out.autoFilter = autoFilter;
  if (freeze === undefined) delete out.freeze;
  else out.freeze = freeze;
  return out;
}

// ---------------------------------------------------------------------------------------------
// formulas
// ---------------------------------------------------------------------------------------------
export interface FormulaCtx {
  ops: readonly StructOp[];
  sheet: string;
  sheetLc: string;
  /** per op: excelRef TableRef[] of the merged sheet BEFORE that op */
  tableRefsAt: TableRef[][];
  /** lower-cased names of tables hit by delete ops (structured refs may break) */
  hitTables: string[];
  /** per host sheet (lower-cased): memo of reference adjustments shared by filled-down formulas */
  memos: Map<string | null, AdjustMemo>;
}

interface AdjustedText {
  text: string;
  compiled: Compiled | null;
  refError: boolean;
}

/** Adjust one formula (null = unchanged). Structured references to deleted table columns go through adjustFormula. */
function adjustText(ctx: FormulaCtx, text: string, c: Compiled, hostLc: string | null, host: string | null): AdjustedText | null {
  let slow = false;
  if (ctx.hitTables.length) {
    if (c.sref) slow = true;
    else {
      const t = text.toLowerCase();
      slow = ctx.hitTables.some((n) => t.includes(n));
    }
  }
  if (!slow) {
    let memo = ctx.memos.get(hostLc);
    if (!memo) ctx.memos.set(hostLc, (memo = new Map()));
    const out = adjustCompiled(c, ctx.ops, hostLc, ctx.sheetLc, false, memo);
    if (!out || !out.changed) return null;
    return { text: out.text, compiled: out.compiled, refError: out.createdRefError };
  }
  let cur = text;
  let refError = false;
  ctx.ops.forEach((op, i) => {
    const r = adjustFormula(cur, op, { hostSheet: host, targetSheet: ctx.sheet, tables: ctx.tableRefsAt[i] });
    if (r.createdRefError) refError = true;
    cur = r.formula;
  });
  return cur === text ? null : { text: cur, compiled: null, refError };
}

/** Dry run for one formula: would the ops create #REF!? */
export function wouldBreak(ctx: FormulaCtx, text: string, c: Compiled, host: string | null): boolean {
  const r = adjustText(ctx, text, c, host === null ? null : lc(host), host);
  return !!r && r.refError;
}

function adjustArrText(arr: string, ops: readonly StructOp[]): string | undefined {
  const a0 = parseAreaText(arr.replace(/\$/g, ''));
  if (!a0) return arr;
  let a: Area | null = a0;
  for (const op of ops) {
    a = adjustArea(a, op);
    if (!a) return undefined;
  }
  return a === a0 ? arr : renderArea(a);
}

function adjustColumns(columns: DocColumn[], ctx: FormulaCtx, counter: { refErrors: number }): DocColumn[] {
  let changedAny = false;
  const out = columns.map((col) => {
    let cells = col.cells;
    let copied = false;
    const cc = compiledCol(col.cells);
    let ncc: CompiledCol | null = null;
    let prev: Compiled | null = null;
    let prevRow = 0;
    for (let r = 0; r < cells.length; r++) {
      const cell = cells[r];
      if (!cell || cell.f === undefined) continue;
      const f = cell.f;
      const c = compiledAt(cc, r, f, prev, r - prevRow);
      prev = c;
      prevRow = r;
      const adj = adjustText(ctx, f, c, ctx.sheetLc, ctx.sheet);
      const narr = cell.arr !== undefined ? adjustArrText(cell.arr, ctx.ops) : undefined;
      if (!adj && narr === cell.arr) continue;
      const ncell: Cell = { ...cell };
      if (adj) {
        ncell.f = adj.text;
        if (adj.refError) {
          ncell.t = 'e';
          ncell.v = '#REF!';
          counter.refErrors++;
        }
      }
      if (narr !== cell.arr) {
        if (narr === undefined) delete ncell.arr;
        else ncell.arr = narr;
      }
      if (!copied) {
        cells = cells.slice();
        copied = true;
        ncc = { f: cc.f.slice(), c: cc.c.slice() };
      }
      cells[r] = ncell;
      if (adj) {
        ncc!.f[r] = adj.compiled ? adj.text : undefined;
        ncc!.c[r] = adj.compiled ?? undefined;
      }
    }
    if (!copied) return col;
    changedAny = true;
    setCompiledCol(cells, ncc!);
    return { ...col, cells };
  });
  return changedAny ? out : columns;
}

function adjustExternal(list: ExternalFormulaCell[], ctx: FormulaCtx, counter: { refErrors: number }): ExternalFormulaCell[] {
  let out: ExternalFormulaCell[] | null = null;
  const cc = compiledCol(list);
  let ncc: CompiledCol | null = null;
  let prev: Compiled | null = null;
  let pk = '';
  let prow = 0;
  for (let i = 0; i < list.length; i++) {
    const e = list[i];
    const k = `${e.sheet}\u0000${e.col}`;
    const c = compiledAt(cc, i, e.f, k === pk ? prev : null, e.row - prow);
    prev = c;
    pk = k;
    prow = e.row;
    const adj = adjustText(ctx, e.f, c, lc(e.sheet), e.sheet);
    if (!adj) continue;
    if (!out) {
      out = list.slice();
      ncc = { f: cc.f.slice(), c: cc.c.slice() };
    }
    const ne: ExternalFormulaCell = { sheet: e.sheet, row: e.row, col: e.col, f: adj.text };
    ncc!.f[i] = adj.compiled ? adj.text : undefined;
    ncc!.c[i] = adj.compiled ?? undefined;
    if (adj.refError) {
      counter.refErrors++;
      // what Excel will compute; a later "convert to values" must not resurrect the stale value
      const old = cachedOf(e);
      setCached(ne, old?.arr !== undefined ? { t: 'e', v: '#REF!', arr: old.arr } : { t: 'e', v: '#REF!' });
    } else setCached(ne, cachedOf(e));
    out[i] = ne;
  }
  if (out) setCompiledCol(out, ncc!);
  return out ?? list;
}

function adjustNames(names: DefinedName[], ctx: FormulaCtx): DefinedName[] {
  return mapChanged(names, (n) => {
    let cur = n.refersTo;
    ctx.ops.forEach((op, i) => {
      cur = adjustFormula(cur, op, { hostSheet: n.scope, targetSheet: ctx.sheet, isName: true, tables: ctx.tableRefsAt[i] }).formula;
    });
    return cur === n.refersTo ? n : { ...n, refersTo: cur };
  });
}

function adjustExternalRules(list: MergedDoc['externalRules'], ctx: FormulaCtx): MergedDoc['externalRules'] {
  return mapChanged(list, (er) => {
    let changed = false;
    const formulas = er.rule.formulas.map((f) => {
      let cur = f;
      ctx.ops.forEach((op, i) => {
        cur = adjustFormula(cur, op, { hostSheet: er.sheet, targetSheet: ctx.sheet, tables: ctx.tableRefsAt[i] }).formula;
      });
      if (cur !== f) changed = true;
      return cur;
    });
    return changed ? { sheet: er.sheet, rule: { ...er.rule, formulas } } : er;
  });
}

function hitTablesOf(tables: TableInfo[], op: StructOp): string[] {
  if (op.type !== 'delete') return [];
  const s1 = op.at;
  const s2 = op.at + op.count - 1;
  const out: string[] = [];
  for (const t of tables) {
    const b = tableBox(t);
    if (b && s1 <= b.c2 && s2 >= b.c1) out.push(t.name.toLowerCase());
  }
  return out;
}

/** Build the formula context for `ops` starting from `features` (tables are tracked op by op). */
export function formulaCtxFor(doc: MergedDoc, ops: readonly StructOp[]): FormulaCtx {
  let tables = doc.features.tables;
  const tableRefsAt: TableRef[][] = [];
  const hit = new Set<string>();
  for (const op of ops) {
    tableRefsAt.push(tableRefs(tables, doc.sheetName));
    for (const n of hitTablesOf(tables, op)) hit.add(n);
    tables = tablesOp(tables, op);
  }
  return { ops, sheet: doc.sheetName, sheetLc: lc(doc.sheetName), tableRefsAt, hitTables: [...hit], memos: new Map() };
}

// ---------------------------------------------------------------------------------------------
// entry point
// ---------------------------------------------------------------------------------------------
/** Why Excel would refuse the op sequence (each op checked on the state before it), or null. No formula work. */
export function refusalFor(doc: MergedDoc, ops: readonly StructOp[]): string | null {
  let columns = doc.columns;
  // only merges, arrays and tables can refuse: rules / links need not follow the ops here
  const f0 = doc.features;
  let features: SheetFeatures = { merges: f0.merges, arrays: f0.arrays, tables: f0.tables, rules: [], hyperlinks: [] };
  for (const op of ops) {
    const why = refusalOf(columns, features, op);
    if (why) return why;
    if (op.type === 'move') columns = permuteColumns(columns, op);
    else if (op.type === 'insert') columns = [...columns.slice(0, op.at - 1), ...new Array<DocColumn>(op.count).fill(columns[0]), ...columns.slice(op.at - 1)];
    else columns = [...columns.slice(0, op.at - 1), ...columns.slice(op.at - 1 + op.count)];
    features = featuresOp(features, op, doc.sheetName, columns, []);
  }
  return null;
}

/**
 * Apply `ops` (all on the merged sheet, column axis). `newIds` supplies the ids of inserted columns in order.
 */
export function applyStructOps(doc: MergedDoc, ops: readonly StructOp[], newIds: readonly string[] = []): StructOutcome {
  if (!ops.length) return { ok: true, doc, refErrors: 0, collapsed: [] };
  let columns = doc.columns;
  let features = doc.features;
  const tableRefsAt: TableRef[][] = [];
  const hit = new Set<string>();
  const writes: HeaderWrite[] = [];
  const insertStyles = new InsertStyles(doc.styles);
  const collapsed: Collapse['masters'] = [];
  const collapsedTexts: string[] = [];
  let idPos = 0;
  for (const op of ops) {
    const why = refusalOf(columns, features, op);
    if (why) return { ok: false, reason: why };
    const cut = collapseSpills(columns, features, op);
    if (cut) {
      columns = cut.columns;
      features = cut.features;
      collapsed.push(...cut.masters);
      collapsedTexts.push(...cut.texts);
    }
    tableRefsAt.push(tableRefs(features.tables, doc.sheetName));
    for (const n of hitTablesOf(features.tables, op)) hit.add(n);
    if (op.type === 'move') columns = permuteColumns(columns, op);
    else if (op.type === 'insert') {
      const ids = newIds.slice(idPos, idPos + op.count);
      idPos += op.count;
      while (ids.length < op.count) ids.push(`n?${idPos - op.count + ids.length}`);
      const inserted = makeInsertedColumns(columns[op.at - 2], columns[op.at - 1], ids, doc.rowCount, insertStyles);
      columns = [...columns.slice(0, op.at - 1), ...inserted, ...columns.slice(op.at - 1)];
    } else columns = [...columns.slice(0, op.at - 1), ...columns.slice(op.at - 1 + op.count)];
    features = featuresOp(features, op, doc.sheetName, columns, writes);
  }
  const ctx: FormulaCtx = { ops, sheet: doc.sheetName, sheetLc: lc(doc.sheetName), tableRefsAt, hitTables: [...hit], memos: new Map() };
  const counter = { refErrors: 0 };
  columns = adjustColumns(columns, ctx, counter);
  if (collapsed.length) columns = pinMasterRanges(columns, collapsed);
  // header cells of table columns created by inserts (Excel writes the generated names)
  if (writes.length) {
    const byId = new Map(columns.map((c, i) => [c.id, i]));
    columns = columns.slice();
    for (const w of writes) {
      const i = byId.get(w.id);
      if (i === undefined || w.row < 1 || w.row > doc.rowCount) continue;
      const col = columns[i];
      const cells = col.cells.slice();
      const old = cells[w.row - 1];
      cells[w.row - 1] = old && old.s !== undefined ? { t: 's', v: w.text, s: old.s } : { t: 's', v: w.text };
      columns[i] = { ...col, cells };
    }
  }
  const externalFormulas = adjustExternal(doc.externalFormulas, ctx, counter);
  const names = adjustNames(doc.names, ctx);
  const externalRules = adjustExternalRules(doc.externalRules, ctx);
  const next: MergedDoc = {
    ...doc,
    columns,
    features,
    styles: insertStyles.styles,
    externalFormulas,
    names,
    externalRules,
    version: doc.version + 1,
  };
  return { ok: true, doc: next, refErrors: counter.refErrors, collapsed: collapsedTexts };
}

