/**
 * Random operation logs of the consistency oracle (task V1), built through the doc API only (moveColumns incl. non-contiguous selections
 * and title-merge unmerge, insertColumns, deleteColumns with convertDependents, setCellInput), with column-id bookkeeping so that the
 * expected result of a log can be written down WITHOUT replaying the layout operations:
 *  - the BASELINE of a scenario is the premerged package plus the typed values / formulas (rendered at the columns' ORIGINAL positions)
 *    plus the values the doc wrote for converted formulas — no move / insert / delete at all. Excel recalculates it; every cell must have
 *    the same value after the layout operations (found again by column id), unless Excel itself changes it (a range whose column span changes,
 *    a lookup by column number: both are explained and counted, never silently accepted).
 */
import { analyzeDelete, deleteColumns, insertColumns, moveColumns, setCellInput } from '../../../src/core/doc';
import { tokenize } from '../../../src/core/formula/excelRef';
import type { CellWrite, EditorSnapshot, LoggedOp, MergedDoc } from '../../../src/core/types';
import { colLetter } from './lib';

export type Rnd = () => number;

export interface EditRec {
  /** stable id of the column the text was typed into */
  colId: string;
  row: number;
  /** what was typed (current positions) */
  text: string;
  /** the same input with every reference rendered at the columns' original positions (the baseline) */
  origText: string;
}

export interface Trail {
  desc: string;
  ok: boolean;
  refused?: string;
  warnings: string[];
}

export interface Converted {
  /** merged-sheet cell (column id) or other sheet (id undefined) */
  sheet: string;
  id?: string;
  row: number;
  /** column of other-sheet cells */
  col?: number;
  write: NonNullable<CellWrite['cell']>;
}

export interface Generated {
  final: EditorSnapshot;
  trail: Trail[];
  edits: EditRec[];
  deleted: string[];
  converted: Converted[];
  /** an accepted op produced the lookup-by-column-number warning */
  lookupWarned: boolean;
  /** an accepted op changed the column span of a multi-column range of some formula / name (Excel changes values then) */
  affecting: boolean;
  /** an accepted op changed the column order of a Table (structured references stay by name) */
  opsKinds: Record<string, number>;
}

export interface GenOpts {
  steps: number;
  /** 'inv': moves / inserts / typed cells.  'del': moves / inserts / deletes (convertDependents), no typed cells. */
  mode: 'inv' | 'del';
  /** false: an op that changes the span of a multi-column range is not taken (the scenario then has to keep EVERY value) */
  allowAffecting: boolean;
}

export const LOOKUP_WARN = /列番号・位置で列を指定する関数/;

// ---------------------------------------------------------------------------------------------------------------------
// range membership (does Excel have a reason to change a value?)
// ---------------------------------------------------------------------------------------------------------------------

type Sig = string;

/** ids covered by a column span, inserted columns left out */
function idsOf(doc: MergedDoc, c1: number, c2: number): string {
  const ids: string[] = [];
  for (let c = c1; c <= Math.min(c2, doc.columns.length); c++) {
    const id = doc.columns[c - 1]?.id;
    if (id && !/^n\d+$/.test(id)) ids.push(id);
  }
  return ids.sort().join(',');
}

function tableNames(doc: MergedDoc, table: string, a: string, b: string): string {
  const t = doc.features.tables.find((x) => x.name.toLowerCase() === table.toLowerCase());
  if (!t) return `?${table}`;
  const i = t.columns.indexOf(a);
  const j = t.columns.indexOf(b);
  if (i < 0 || j < 0) return `?${a}:${b}`;
  return t.columns.slice(Math.min(i, j), Math.max(i, j) + 1).filter((x) => !/^列\d+$|^Column\d+$/.test(x)).sort().join(',');
}

/**
 * signatures of the areas (A1:B2, C:D, also a one-column area: a range can shrink to one column) of one formula, in token order: the set of
 * column ids the area covers right now
 */
export function areaSigs(doc: MergedDoc, formula: string, host: string): Sig[] {
  const out: Sig[] = [];
  for (const t of tokenize(formula)) {
    if (t.type === 'ref' && t.area && (t.area.kind === 'area' || t.area.kind === 'cols')) {
      const sheet = t.prefix ? t.prefix.sheet : host;
      if (t.prefix?.book || (t.prefix && t.prefix.sheet === null)) continue;
      if (sheet === null || sheet.toLowerCase() !== doc.sheetName.toLowerCase()) continue;
      out.push(`A:${idsOf(doc, t.area.c1, t.area.c2)}`);
    } else if (t.type === 'sref') {
      const m = /^([^[\]]*)\[.*\[([^[\]]+)\]:\[([^[\]]+)\]/.exec(t.text);
      if (m) out.push(`T:${tableNames(doc, m[1], m[2], m[3])}`);
    }
  }
  return out;
}

/** every formula / name of the doc -> its multi-column area signatures (keys are stable across moves: column ids, other-sheet addresses) */
export function membership(doc: MergedDoc): Map<string, Sig[]> {
  const out = new Map<string, Sig[]>();
  doc.columns.forEach((col) => {
    col.cells.forEach((cell, r) => {
      if (!cell || cell.f === undefined) return;
      out.set(`c:${col.id}:${r + 1}`, areaSigs(doc, cell.f, doc.sheetName));
    });
  });
  for (const e of doc.externalFormulas) {
    out.set(`x:${e.sheet}:${e.row}:${e.col}`, areaSigs(doc, e.f, e.sheet));
  }
  for (const n of doc.names) {
    out.set(`n:${n.scope ?? ''}:${n.name}`, areaSigs(doc, n.refersTo, n.scope ?? doc.sheetName));
  }
  return out;
}

/** true when some formula keeps existing but covers other columns than before */
export function membershipChanged(before: Map<string, Sig[]>, after: Map<string, Sig[]>): boolean {
  for (const [k, a] of before) {
    const b = after.get(k);
    if (!b) continue; // the cell became a value / was deleted: not a span change
    if (a.length !== b.length || a.some((x, i) => x !== b[i])) return true;
  }
  return false;
}

// ---------------------------------------------------------------------------------------------------------------------
// the generator
// ---------------------------------------------------------------------------------------------------------------------

const CONSTS = ['12', '2.5', '-7', 'メモ', 'TRUE', '0', '48', '100'];
/** a typed formula: asks `r` for every cell it reads (so the generator can check them) and renders the text */
type Tpl = (r: (id: string, row: number) => string, ids: string[], row: number) => string;
const TEMPLATES: Tpl[] = [
  (r, [a], row) => `=${r(a, row)}*2`,
  (r, [a], row) => `=${r(a, row)}&"x"`,
  (r, [a, b], row) => `=SUM(${r(a, row)},${r(b, row)})`,
  (r, [a, b], row) => `=${r(a, row)}+${r(b, row)}`,
  (r, [a], row) => `=IF(${r(a, row)}="","",1)`,
  (r, [a], row) => `=COUNT(${r(a, row)})`,
  (r, [a], row) => `=${r(a, Math.max(1, row - 1))}`,
  (r, [a], row) => `=SUM(${r(a, row)}:${r(a, row + 1)})`,
];

export function generateOps(s0: EditorSnapshot, rnd: Rnd, o: GenOpts): Generated {
  let s = s0;
  const trail: Trail[] = [];
  const edits: EditRec[] = [];
  const deleted: string[] = [];
  const converted: Converted[] = [];
  const opsKinds: Record<string, number> = {};
  let lookupWarned = false;
  let affecting = false;
  const origIds = s0.doc.columns.map((c) => c.id);
  const origPos = new Map(origIds.map((id, i) => [id, i + 1]));
  const region = { first: s0.doc.firstDataRow, last: s0.doc.lastDataRow };
  const idAtOrig = (i: number) => origIds[i];
  const present = () => s.doc.columns.map((c) => c.id);
  // typed formulas read ONLY cells that hold plain constants in the premerged sheet and were never given a typed formula, and never the
  // cell they are typed into: so a typed formula can never close a loop (a circular reference has no defined value in Excel)
  const formulaTargets = new Set<string>();
  const isConstant = (id: string, row: number): boolean => {
    const i = origPos.get(id);
    if (!i) return false;
    const c = s0.doc.columns[i - 1].cells[row - 1];
    return !(c && c.f !== undefined) && !formulaTargets.has(`${id}:${row}`);
  };

  const refAt = (doc: { columns: Array<{ id: string }> }, id: string, row: number): string => `${colLetter(doc.columns.findIndex((c) => c.id === id) + 1)}${row}`;
  const refOrig = (id: string, row: number): string => `${colLetter(origPos.get(id) ?? 1)}${row}`;

  const accept = (desc: string, res: ReturnType<typeof moveColumns>, kind: string): boolean => {
    if (!res.ok) {
      trail.push({ desc, ok: false, refused: res.reason, warnings: [] });
      return false;
    }
    const loop = selfReference(res.snapshot.doc);
    if (loop && !selfReference(s.doc)) {
      trail.push({ desc, ok: false, refused: `(not taken: a formula would read its own cell: ${loop.slice(0, 80)})`, warnings: res.warnings });
      return false;
    }
    const before = membership(s.doc);
    const after = membership(res.snapshot.doc);
    const aff = membershipChanged(before, after);
    if (aff && !o.allowAffecting) {
      trail.push({ desc, ok: false, refused: '(not taken: changes the column span of a range)', warnings: res.warnings });
      return false;
    }
    if (aff) affecting = true;
    if (res.warnings.some((w) => LOOKUP_WARN.test(w))) lookupWarned = true;
    // the doc's own dependency analysis of a deletion names the cells it converted: collect them from the log entries that were added
    for (const e of res.snapshot.log.slice(s.log.length)) {
      if (e.kind === 'cells' && e.reason === 'toValues') {
        for (const w of e.writes) {
          if (!w.cell) {
            converted.push({ sheet: e.sheet, id: e.sheet === s.doc.sheetName ? s.doc.columns[w.col - 1]?.id : undefined, row: w.row, col: w.col, write: { t: 'z', v: null } });
          } else {
            converted.push({ sheet: e.sheet, id: e.sheet === s.doc.sheetName ? s.doc.columns[w.col - 1]?.id : undefined, row: w.row, col: w.col, write: w.cell });
          }
        }
      }
    }
    trail.push({ desc, ok: true, warnings: res.warnings });
    opsKinds[kind] = (opsKinds[kind] ?? 0) + 1;
    s = res.snapshot;
    return true;
  };

  for (let step = 0; step < o.steps; step++) {
    const n = s.doc.columns.length;
    const x = rnd();
    if (x < 0.5) {
      // move: 1-4 columns, contiguous or not, anywhere
      const k = 1 + Math.floor(rnd() * 4);
      const idx = new Set<number>();
      if (rnd() < 0.35) {
        const a = Math.floor(rnd() * n);
        for (let j = 0; j < k && a + j < n; j++) idx.add(a + j);
      } else {
        for (let j = 0; j < k; j++) idx.add(Math.floor(rnd() * n));
      }
      const picked = [...idx].sort((a, b) => a - b);
      const gap = Math.floor(rnd() * (n + 1));
      const ids = picked.map((i) => s.doc.columns[i].id);
      let res = moveColumns(s, ids, gap);
      let unmerged = false;
      if (!res.ok && res.fix?.kind === 'unmergeTitles' && rnd() < 0.85) {
        res = moveColumns(s, ids, gap, { unmergeTitles: true });
        unmerged = true;
      }
      accept(`move [${picked.map((i) => `${colLetter(i + 1)}(${s.doc.columns[i].id})`).join(' ')}] -> slot ${gap}${unmerged ? ' +unmerge' : ''}`, res, unmerged ? 'move+unmerge' : 'move');
    } else if (x < 0.63) {
      const at = Math.floor(rnd() * (n + 1));
      const count = 1 + Math.floor(rnd() * 2);
      accept(`insert ${count} @${at}`, insertColumns(s, at, count), 'insert');
    } else if (o.mode === 'del' && x < 0.85) {
      const k = 1 + Math.floor(rnd() * 2);
      const idx = new Set<number>();
      for (let j = 0; j < k; j++) idx.add(Math.floor(rnd() * n));
      const ids = [...idx].sort((a, b) => a - b).map((i) => s.doc.columns[i].id);
      const impact = analyzeDelete(s, ids);
      const convert = impact.breakCount > 0 || rnd() < 0.5;
      const ok = accept(`delete [${ids.join(' ')}]${convert ? ' (convert)' : ''}`, deleteColumns(s, ids, { convertDependents: convert }), 'delete');
      if (ok) deleted.push(...ids);
    } else if (o.mode === 'inv') {
      // a typed value or formula in a data cell of a column that existed at the start
      const cands = present().filter((id) => origPos.has(id));
      if (!cands.length) continue;
      const colId = cands[Math.floor(rnd() * cands.length)];
      const row = region.first + Math.floor(rnd() * Math.max(1, region.last - region.first + 1));
      let text = '';
      let origText = '';
      let formula = false;
      if (rnd() < 0.4) {
        text = origText = CONSTS[Math.floor(rnd() * CONSTS.length)];
      } else {
        for (let attempt = 0; attempt < 8 && !text; attempt++) {
          const ids = [cands[Math.floor(rnd() * cands.length)], cands[Math.floor(rnd() * cands.length)]];
          const tpl = TEMPLATES[Math.floor(rnd() * TEMPLATES.length)];
          const reads: Array<[string, number]> = [];
          tpl((id, rr) => (reads.push([id, rr]), 'A1'), ids, row);
          if (reads.some(([id, rr]) => !isConstant(id, rr) || (id === colId && rr === row))) continue;
          text = tpl((id, rr) => refAt(s.doc, id, rr), ids, row);
          origText = tpl((id, rr) => refOrig(id, rr), ids, row);
          formula = true;
        }
        if (!text) continue;
      }
      const res = setCellInput(s, colId, row, text);
      if (accept(`type ${JSON.stringify(text)} @${colId} row ${row}`, res, 'edit')) {
        edits.push({ colId, row, text, origText });
        if (formula) formulaTargets.add(`${colId}:${row}`);
      }
    }
  }
  void idAtOrig;
  return { final: s, trail, edits, deleted, converted, lookupWarned, affecting, opsKinds };
}

/**
 * The log that produces the BASELINE from the premerged package: the typed texts at the original positions (set with the doc API on the
 * initial snapshot, so the cells are written exactly as the user's input was) and the values the doc wrote for converted formulas.
 * null = the doc refused one of the edits in the original layout (the scenario cannot be judged).
 */
export function baselineLog(s0: EditorSnapshot, g: Generated): { log: LoggedOp[] } | null {
  let s = s0;
  for (const e of g.edits) {
    const res = setCellInput(s, e.colId, e.row, e.origText);
    if (!res.ok) return null;
    s = res.snapshot;
  }
  const log: LoggedOp[] = [...s.log];
  const origIds = s0.doc.columns.map((c) => c.id);
  const byEntry = new Map<string, CellWrite[]>();
  for (const c of g.converted) {
    const col = c.id !== undefined ? origIds.indexOf(c.id) + 1 : c.col!;
    if (c.id !== undefined && col < 1) continue;
    const key = c.sheet;
    const list = byEntry.get(key) ?? [];
    list.push({ row: c.row, col, cell: c.write.t === 'z' && c.write.v === null ? null : c.write });
    byEntry.set(key, list);
  }
  for (const [sheet, writes] of byEntry) log.push({ kind: 'cells', sheet, writes, reason: 'toValues' });
  return { log };
}

/**
 * The doc alone (no Excel): every reference to ONE cell or ONE column of the merged sheet — in the sheet's own formulas, in the formulas of
 * the other sheets and in the defined names — must still point at the same column id after the operations (ranges follow Excel's own
 * span rules and are judged by value). A reference to a deleted column must not survive as a column reference.
 */
export function docIdentity(initial: MergedDoc, fin: MergedDoc, deleted: ReadonlySet<string>, typed: ReadonlySet<string> = new Set()): string[] {
  const out: string[] = [];
  const idsI = initial.columns.map((c) => c.id);
  const idsF = fin.columns.map((c) => c.id);
  const sheetLc = initial.sheetName.toLowerCase();
  const compare = (where: string, before: string, after: string, host: string | null) => {
    const tb = tokenize(before).filter((t) => t.type === 'ref' && t.area && !t.prefix?.book);
    const ta = tokenize(after).filter((t) => t.type === 'ref' && t.area && !t.prefix?.book);
    if (tb.length !== ta.length) {
      out.push(`${where}: ${before} -> ${after} (the number of references changed)`);
      return;
    }
    tb.forEach((x, i) => {
      const y = ta[i];
      const a = x.area!;
      const b = y.area!;
      const target = (x.prefix ? x.prefix.sheet : host)?.toLowerCase();
      if (target !== sheetLc) return;
      const one = (ar: typeof a) => ar.kind === 'cell' || ((ar.kind === 'area' || ar.kind === 'cols') && ar.c1 === ar.c2);
      if (!one(a)) return;
      if (!one(b)) {
        out.push(`${where}: ${x.text} -> ${y.text} (a one-column reference became a range)`);
        return;
      }
      const idB = idsI[a.c1 - 1];
      const idA = idsF[b.c1 - 1];
      if (idB !== idA) out.push(`${where}: ${x.text} (${idB}) -> ${y.text} (${idA}); ${before} -> ${after}`);
      if (a.kind === 'cell' && b.kind === 'cell' && a.r1 !== b.r1) out.push(`${where}: ${x.text} -> ${y.text} (the row changed)`);
    });
  };
  const finIndex = new Map(fin.columns.map((c, i) => [c.id, i]));
  initial.columns.forEach((col) => {
    if (deleted.has(col.id)) return;
    const j = finIndex.get(col.id);
    if (j === undefined) return;
    col.cells.forEach((cell, r) => {
      if (!cell || cell.f === undefined) return;
      const nc = fin.columns[j].cells[r];
      if (!nc || nc.f === undefined || typed.has(`${col.id}:${r + 1}`)) return; // converted to a value / typed over
      if (nc.f === cell.f && j === initial.columns.indexOf(col)) return;
      compare(`${initial.sheetName}!${colLetter(initial.columns.indexOf(col) + 1)}${r + 1} (${col.id})`, cell.f, nc.f, initial.sheetName);
    });
  });
  const ext0 = new Map(initial.externalFormulas.map((e) => [`${e.sheet}!${e.col},${e.row}`, e.f]));
  for (const e of fin.externalFormulas) {
    const f0 = ext0.get(`${e.sheet}!${e.col},${e.row}`);
    if (f0 !== undefined && f0 !== e.f) compare(`${e.sheet}!${colLetter(e.col)}${e.row}`, f0, e.f, e.sheet);
  }
  const names0 = new Map(initial.names.map((n) => [`${n.scope ?? ''}!${n.name}`, n]));
  for (const n of fin.names) {
    const n0 = names0.get(`${n.scope ?? ''}!${n.name}`);
    if (n0 && n0.refersTo !== n.refersTo && !/#REF!/.test(n.refersTo)) compare(`name ${n.name}`, n0.refersTo, n.refersTo, n.scope);
  }
  return out;
}

/**
 * A formula of the merged sheet whose own reference range contains the formula's own cell (the column was moved INTO the range of a
 * formula that lives in it): a circular reference, which has no defined value in Excel. First such cell as text, or null.
 */
export function selfReference(doc: MergedDoc): string | null {
  const sheetLc = doc.sheetName.toLowerCase();
  for (let ci = 0; ci < doc.columns.length; ci++) {
    const cells = doc.columns[ci].cells;
    for (let r = 0; r < cells.length; r++) {
      const f = cells[r]?.f;
      if (f === undefined || !/[:A-Za-z]/.test(f)) continue;
      for (const t of tokenize(f)) {
        if (t.type !== 'ref' || !t.area || t.prefix?.book) continue;
        if (t.prefix && (t.prefix.sheet ?? '').toLowerCase() !== sheetLc) continue;
        const a = t.area;
        const inCols = a.kind === 'rows' ? true : ci + 1 >= Math.min(a.c1, a.c2) && ci + 1 <= Math.max(a.c1, a.c2);
        const inRows = a.kind === 'cols' ? true : r + 1 >= Math.min(a.r1, a.r2) && r + 1 <= Math.max(a.r1, a.r2);
        if (inCols && inRows) return `${colLetter(ci + 1)}${r + 1} (${doc.columns[ci].id}): ${f}`;
      }
    }
  }
  return null;
}
