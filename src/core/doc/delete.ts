/**
 * Column deletion: impact analysis (what would become #REF!) and the delete itself, optionally converting every
 * formula that would break into its cached value first (logged as {kind:'cells', reason:'toValues'} BEFORE the
 * struct ops, so both export engines replay exactly the same thing).
 *
 * Deleted columns are grouped into contiguous runs and deleted right to left (positions of the runs to the left
 * stay valid), one excelRef StructOp per run.
 */
import type { StructOp } from '../formula/excelRef';
import type { Cell, CellWrite, DocColumn, EditorSnapshot, ExternalFormulaCell, LoggedOp, MergedDoc } from '../types';
import type { DeleteImpact, OpResult } from '../contracts';
import { breakCount, describeBreaks, findBreaks, lookupIssues, lookupWarningTexts, nameLabel, spillReads } from './deps';
import type { BreakList } from './deps';
import { InsertStyles, textWrapsOnWrite } from './insertStyle';
import { applyStructOps, collapseWarnings, refusalFor } from './struct';
import { cachedOf, carryCompiled, colIndexOf } from './util';

/** 0-based indices of `ids` (deduplicated); a string reason when an id is unknown. */
export function indicesOf(doc: MergedDoc, ids: readonly string[]): number[] | string {
  const out = new Set<number>();
  for (const id of ids) {
    const i = colIndexOf(doc, id);
    if (i < 0) return '指定された列が見つかりません。画面を更新してからやり直してください。';
    out.add(i);
  }
  return [...out].sort((a, b) => a - b);
}

/** Delete ops for 0-based indices: contiguous runs, right to left. */
export function deleteOpsFor(indices: readonly number[]): StructOp[] {
  const sorted = [...new Set(indices)].sort((a, b) => a - b);
  const runs: Array<[number, number]> = [];
  for (const i of sorted) {
    const last = runs[runs.length - 1];
    if (last && i === last[1] + 1) last[1] = i;
    else runs.push([i, i]);
  }
  return runs.reverse().map(([a, b]) => ({ type: 'delete', axis: 'col', at: a + 1, count: b - a + 1 }));
}

export function analyzeDeleteImpl(s: EditorSnapshot, ids: string[]): DeleteImpact {
  const doc = s.doc;
  const idx = indicesOf(doc, ids);
  if (typeof idx === 'string' || !idx.length) return { breaks: [], breakCount: 0, lookupWarnings: [] };
  const ops = deleteOpsFor(idx);
  const gone = new Set(idx);
  const b = findBreaks(doc, ops, gone);
  const lookups = lookupIssues(doc, ops, (i) => gone.has(i));
  return { breaks: describeBreaks(doc, b), breakCount: breakCount(b), lookupWarnings: lookupWarningTexts(lookups, '削除') };
}

// ---------------------------------------------------------------------------------------------
// convert-to-values
// ---------------------------------------------------------------------------------------------
interface Conversion {
  doc: MergedDoc;
  log: LoggedOp[];
  converted: number;
  /** formulas without a cached result (typed in this session / never calculated): cleared */
  blanks: number;
  /** array / spill formulas converted as a whole */
  arrays: number;
}

/** one formula to convert: its logged write plus what ordering needs to know */
interface Item {
  sheet: string;
  write: CellWrite;
  /** the cell's own key when it is an array / spill master */
  master: string | null;
  /** keys of the spill anchors the formula reads (`key#`) */
  reads: string[];
  ext: boolean;
}

const cellKey = (sheet: string, col: number, row: number) => `${sheet.toLowerCase()}!${col},${row}`;

/**
 * Cached values over every formula of `b`. An array / spill MASTER is written like any other formula: a 'toValues'
 * write over a master replaces the whole array by its values (the patcher / Excel runner convert the range; the
 * other cells of the range already hold their cached values), so the doc drops the array range and keeps the values.
 *
 * Order: a formula that reads a spill (`master#`) is converted BEFORE that master — with automatic calculation, once the
 * master is a plain value `master#` is #REF! in Excel and the reader's own spilled values are gone. (The Excel runner
 * replays with manual calculation, so it accepts either order — tests/doc/excel-engine.test.ts — but the order stays
 * valid for any engine.) Entries are emitted by rank (readers first); without spill readers everything has rank 0: the
 * merged sheet first, then one entry per other sheet.
 */
function convertBreaks(doc: MergedDoc, b: BreakList): Conversion {
  let converted = 0;
  let blanks = 0;
  let arrays = 0;
  const masters = new Set<string>();
  const items: Item[] = [];
  let derived: InsertStyles | null = null;
  // merged sheet: column by column, rows ascending (contiguous blocks for the COM runner)
  const cells = [...b.cells].sort((x, y) => x.index - y.index || x.row - y.row);
  let columns: DocColumn[] = doc.columns;
  if (cells.length) {
    columns = doc.columns.slice();
    const copied = new Map<number, Array<Cell | undefined>>();
    for (const { index, row } of cells) {
      let arr = copied.get(index);
      if (!arr) {
        arr = columns[index].cells.slice();
        carryCompiled(columns[index].cells, arr);
        copied.set(index, arr);
      }
      const cell = arr[row - 1];
      if (!cell || cell.f === undefined) continue;
      converted++;
      const isMaster = cell.arr !== undefined || !!cell.dyn;
      if (isMaster) {
        arrays++;
        masters.add(`${row},${index + 1}`);
      }
      const item: Item = {
        sheet: doc.sheetName,
        write: { row, col: index + 1, cell: null },
        master: isMaster ? cellKey(doc.sheetName, index + 1, row) : null,
        reads: spillReads(cell.f).map((r) => cellKey(r.sheet ?? doc.sheetName, r.col, r.row)),
        ext: false,
      };
      items.push(item);
      if (cell.t === 'z' || cell.v === null) {
        blanks++;
        arr[row - 1] = cell.s !== undefined ? { t: 'z', v: null, s: cell.s } : undefined;
      } else {
        item.write.cell = { t: cell.t, v: cell.v };
        const nc: Cell = { t: cell.t, v: cell.v };
        if (cell.s !== undefined) nc.s = cell.s;
        // a text result with a line feed written as a constant turns the cell's wrap text on (both engines; InsertStyles.withWrap)
        if (textWrapsOnWrite(item.write.cell)) {
          derived ??= new InsertStyles(doc.styles);
          const cur = nc.s ?? 0;
          const ns = derived.withWrap(cur);
          if (ns !== cur) nc.s = ns;
        }
        arr[row - 1] = nc;
      }
    }
    for (const [index, arr] of copied) columns[index] = { ...columns[index], cells: arr };
  }
  // other sheets
  let externalFormulas: ExternalFormulaCell[] = doc.externalFormulas;
  if (b.external.length) {
    const drop = new Set(b.external);
    for (const i of [...b.external].sort((x, y) => x - y)) {
      const e = doc.externalFormulas[i];
      const cv = cachedOf(e);
      converted++;
      const isMaster = cv?.arr !== undefined;
      if (isMaster) arrays++;
      const item: Item = {
        sheet: e.sheet,
        write: { row: e.row, col: e.col, cell: null },
        master: isMaster ? cellKey(e.sheet, e.col, e.row) : null,
        reads: spillReads(e.f).map((r) => cellKey(r.sheet ?? e.sheet, r.col, r.row)),
        ext: true,
      };
      items.push(item);
      if (!cv || cv.t === 'z' || cv.v === null) blanks++;
      else item.write.cell = { t: cv.t, v: cv.v };
    }
    externalFormulas = doc.externalFormulas.filter((_, i) => !drop.has(i));
  }
  const log = emitConversions(doc.sheetName, items);
  // array ranges whose master became a value are plain values now
  const features = masters.size
    ? { ...doc.features, arrays: doc.features.arrays.filter((a) => !masters.has(`${a.r1},${a.c1}`)) }
    : doc.features;
  const next: MergedDoc = { ...doc, columns, externalFormulas, features, styles: derived ? derived.styles : doc.styles, version: doc.version + 1 };
  return { doc: next, log, converted, blanks, arrays };
}

/** 'toValues' log entries: rank 0 = nothing converted reads it; a master ranks above every converted reader of it. */
function emitConversions(mergedSheet: string, items: Item[]): LoggedOp[] {
  const byMaster = new Map<string, Item>();
  for (const it of items) if (it.master) byMaster.set(it.master, it);
  const readers = new Map<string, Item[]>();
  for (const it of items) {
    for (const k of it.reads) {
      const m = byMaster.get(k);
      if (m && m !== it) {
        const list = readers.get(k);
        if (list) list.push(it);
        else readers.set(k, [it]);
      }
    }
  }
  const rank = new Map<Item, number>();
  const visiting = new Set<Item>();
  const rankOf = (it: Item): number => {
    const hit = rank.get(it);
    if (hit !== undefined) return hit;
    if (visiting.has(it)) return 0; // circular spill references: no order exists
    visiting.add(it);
    let r = 0;
    for (const x of it.master ? readers.get(it.master) ?? [] : []) r = Math.max(r, rankOf(x) + 1);
    visiting.delete(it);
    rank.set(it, r);
    return r;
  };
  let top = 0;
  for (const it of items) top = Math.max(top, rankOf(it));
  const log: LoggedOp[] = [];
  for (let r = 0; r <= top; r++) {
    const level = items.filter((it) => rank.get(it) === r);
    const merged = level.filter((it) => !it.ext).map((it) => it.write);
    if (merged.length) log.push({ kind: 'cells', sheet: mergedSheet, writes: merged, reason: 'toValues' });
    const bySheet = new Map<string, CellWrite[]>();
    for (const it of level) {
      if (!it.ext) continue;
      let list = bySheet.get(it.sheet);
      if (!list) bySheet.set(it.sheet, (list = []));
      list.push(it.write);
    }
    for (const [sheet, list] of bySheet) {
      list.sort((x, y) => x.col - y.col || x.row - y.row);
      log.push({ kind: 'cells', sheet, writes: list, reason: 'toValues' });
    }
  }
  return log;
}

// ---------------------------------------------------------------------------------------------
// delete
// ---------------------------------------------------------------------------------------------
export function deleteColumnsImpl(s: EditorSnapshot, ids: string[], opts: { convertDependents: boolean }): OpResult {
  const doc = s.doc;
  const idx = indicesOf(doc, ids);
  if (typeof idx === 'string') return { ok: false, reason: idx };
  if (!idx.length) return { ok: true, snapshot: s, warnings: [] };
  if (idx.length >= doc.columns.length) return { ok: false, reason: 'すべての列を削除することはできません。少なくとも 1 列は残してください。' };
  const ops = deleteOpsFor(idx);
  const gone = new Set(idx);
  const warnings: string[] = [];
  const log: LoggedOp[] = [];
  let base = doc;
  const convert = opts.convertDependents;
  // refusals first (independent of formulas): nothing is logged when Excel would refuse
  if (!convert) {
    const why = refusalFor(doc, ops);
    if (why) return { ok: false, reason: why };
  }
  const b = findBreaks(doc, ops, gone);
  const lookups = lookupIssues(doc, ops, (i) => gone.has(i));
  if (convert && (b.cells.length || b.external.length)) {
    const conv = convertBreaks(doc, b);
    // an array formula converted to values no longer stands in the way of deleting part of its range
    const why = refusalFor(conv.doc, ops);
    if (why) return { ok: false, reason: why };
    base = conv.doc;
    log.push(...conv.log);
    warnings.push(`削除した列を参照していた数式 ${conv.converted.toLocaleString()} 個を値に変換しました。`);
    if (conv.arrays) warnings.push(`そのうち配列数式 ${conv.arrays.toLocaleString()} 個は、範囲全体を値に変換しました。`);
    if (conv.blanks) warnings.push(`そのうち ${conv.blanks.toLocaleString()} 個は計算結果が保存されていなかったため空欄になりました。`);
  } else if (convert) {
    const why = refusalFor(doc, ops);
    if (why) return { ok: false, reason: why };
  }
  const result = applyStructOps(base, ops);
  if (!result.ok) return { ok: false, reason: result.reason };
  if (result.refErrors > 0) warnings.push(`${result.refErrors.toLocaleString()} 個の数式が参照先を失い #REF! になりました。`);
  if (b.names.length) {
    const list = b.names.slice(0, 3).map(nameLabel).join('、');
    warnings.push(`${list}${b.names.length > 3 ? ` ほか ${b.names.length - 3} 個` : ''}は削除した列だけを参照しているため #REF! になります（名前は値に変換できません）。`);
  }
  const ruleCount = b.rules.length + b.externalRules.length;
  if (ruleCount) warnings.push(`条件付き書式・入力規則 ${ruleCount} 個が削除した列を参照しているため、その部分は #REF! になります。`);
  warnings.push(...lookupWarningTexts(lookups, '削除'));
  warnings.push(...collapseWarnings(result.collapsed));
  for (const op of ops) log.push({ kind: 'struct', sheet: doc.sheetName, op });
  const label = idx.length === 1 ? '列を削除' : `${idx.length} 列を削除`;
  return { ok: true, snapshot: { doc: result.doc, log: [...s.log, ...log], label }, warnings };
}
