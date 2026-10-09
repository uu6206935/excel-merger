// Second, independent Excel ground truth (R2b): ~21,500 cases of real Excel 365 results for column/row
// insert / delete / move ("Insert Cut Cells") on a sheet, read back both through COM (Range.Formula/Formula2)
// and from the saved .xlsx (<f> text). Fixture: tests/fixtures/formula/r2b-fixtures.json (trimmed copy, see
// build-r2b-fixture.mjs). Every case is run through excelRef (adjustFormula / adjustSqref / adjustArea /
// adjustAnchoredRule / mapIndex) and every mismatch must be listed in R2B_DEVIATIONS (documented in
// tests/formula/DEVIATIONS.md). Set R2B_REPORT=1 to print grouped mismatches.
import { describe, expect, it } from 'vitest';
import {
  adjustFormula, adjustArea, adjustSqref, adjustAnchoredRule, adjustTableRef, parseAreaText, renderArea, mapIndex, relKey, movePerm, numToCol,
  type StructOp, type Area, type AdjustContext, type TableRef, type AnchoredRule, type AnchoredRuleOptions,
} from '../../src/core/formula/excelRef';
import { readFixture } from './helpers';
import { R2B_DEVIATIONS, type DeviationKind } from './r2bDeviations';

interface R2bOp {
  opId: string; suite: string; primitive: string; sheet: string;
  params: { src?: { start: number; end: number }; before?: { index: number }; count?: number };
  afterSource: string; error: string | null; noTruth: boolean;
}
type FileF = string | { f: string | null; t?: string; ref?: string; si?: string; aca?: string };
interface R2bCase {
  id: string; opId: string; where: string; kind: string; tag?: string;
  cellBefore?: string; cellAfter?: string | null;
  before?: any; after?: any; before2?: string; after2?: string;
  fileBefore?: FileF; fileAfter?: FileF; arrayBefore?: string; arrayAfter?: string;
  source: string; name?: string; scope?: string; dvSheet?: string; cfSheet?: string; ext?: boolean; cfType?: string;
}
interface R2bFixture {
  suites: Record<string, { dataSheet: string; sheets: string[]; tables: TableRef[] }>;
  ops: R2bOp[];
  cases: R2bCase[];
}

const FX = readFixture<R2bFixture>('r2b-fixtures.json');
const OPS = new Map(FX.ops.map((o) => [o.opId, o]));

function toOp(o: R2bOp): StructOp {
  const P = o.params;
  switch (o.primitive) {
    case 'moveCols': return { type: 'move', axis: 'col', from: P.src!.start, count: P.src!.end - P.src!.start + 1, to: P.before!.index };
    case 'insertCols': return { type: 'insert', axis: 'col', at: P.before!.index, count: P.count! };
    case 'deleteCols': return { type: 'delete', axis: 'col', at: P.src!.start, count: P.src!.end - P.src!.start + 1 };
    case 'insertRows': return { type: 'insert', axis: 'row', at: P.before!.index, count: P.count! };
    case 'deleteRows': return { type: 'delete', axis: 'row', at: P.src!.start, count: P.src!.end - P.src!.start + 1 };
  }
  throw new Error(`unknown primitive ${o.primitive}`);
}

const fileText = (f: FileF | undefined): string | null => (f == null ? null : typeof f === 'string' ? f : f.f);
const strip = (f: string) => (f.startsWith('=') ? f.slice(1) : f);

/** host cell A1 ("B40") -> new A1 through the op (null = deleted) */
function mapCell(a1: string, op: StructOp): string | null {
  const a = parseAreaText(a1);
  if (!a) throw new Error(`bad cell ${a1}`);
  if (a.kind !== 'cell') { const r = adjustArea(a, op); return r ? renderArea(r) : null; }
  const v = mapIndex(op.axis === 'col' ? a.c1 : a.r1, op);
  if (v == null) return null;
  return renderArea(op.axis === 'col' ? { ...a, c1: v, c2: v } : { ...a, r1: v, r2: v });
}

// ------------------------------------------------------------------------------------------------ per-cell rules
interface Rect { c1: number; c2: number; r1: number; r2: number; key: string }
const areasOf = (sqref: string): Area[] => sqref.split(/[\s,]+/).filter(Boolean).map((s) => {
  const a = parseAreaText(s.replace(/^\$?([A-Z]+)\$?(\d+)$/, '$1$2'));
  if (!a) throw new Error(`bad sqref part ${s}`);
  return a;
});
const lo = (x: number, y: number) => Math.min(x, y) || 1;
function rectsOf(rules: Array<{ sqref: Area[]; formulas: string[] }>): Rect[] {
  const out: Rect[] = [];
  for (const r of rules) {
    const tr = Math.min(...r.sqref.map((a) => lo(a.r1, a.r2))), tc = Math.min(...r.sqref.map((a) => lo(a.c1, a.c2)));
    const key = r.formulas.map((f) => relKey(f, tr, tc)).join(' | ');
    for (const a of r.sqref) {
      const full = (v1: number, v2: number, max: number) => (v1 === 0 && v2 === 0 ? [1, max] : [Math.min(v1, v2), Math.max(v1, v2)]);
      const [c1, c2] = full(a.c1, a.c2, 16384), [r1, r2] = full(a.r1, a.r2, 1048576);
      out.push({ c1, c2, r1, r2, key });
    }
  }
  return out;
}
/** Compare two per-cell coverings (coordinate compression); returns a description of the first difference or null. */
function diffRects(a: Rect[], b: Rect[]): string | null {
  const xs = [...new Set([...a, ...b].flatMap((r) => [r.c1, r.c2 + 1]))].sort((x, y) => x - y);
  const ys = [...new Set([...a, ...b].flatMap((r) => [r.r1, r.r2 + 1]))].sort((x, y) => x - y);
  const at = (rs: Rect[], c: number, r: number) => rs.filter((q) => q.c1 <= c && c <= q.c2 && q.r1 <= r && r <= q.r2).map((q) => q.key).sort().join(' || ');
  for (let i = 0; i < xs.length - 1; i++) for (let j = 0; j < ys.length - 1; j++) {
    const ka = at(a, xs[i], ys[j]), kb = at(b, xs[i], ys[j]);
    if (ka !== kb) return `cell ${renderArea({ kind: 'cell', c1: xs[i], r1: ys[j], c2: xs[i], r2: ys[j], ac1: false, ar1: false, ac2: false, ar2: false })}: excel [${ka}] ours [${kb}]`;
  }
  return null;
}
type Rule = { sqref: string; formulas: string[] };
const asRules = (x: any): Rule[] => (x == null ? [] : Array.isArray(x) ? x : [x]);

// ------------------------------------------------------------------------------------------------ checkers
interface Outcome { ok: boolean; skip?: string; detail?: string; ours?: string; excel?: string }

function checkFormula(c: R2bCase, o: R2bOp, op: StructOp, tables: TableRef[]): Outcome {
  const hostOnTarget = c.where.toLowerCase() === o.sheet.toLowerCase();
  const probs: string[] = [];
  // host cell mapping (formulas on the op sheet travel with their cell)
  if (hostOnTarget && c.cellBefore) {
    const got = mapCell(c.cellBefore, op);
    if (got !== (c.cellAfter ?? null)) probs.push(`host ${c.cellBefore} -> excel ${c.cellAfter} ours ${got}`);
  }
  if (c.after == null) return probs.length ? { ok: false, detail: probs.join('; ') } : { ok: true }; // host deleted
  const hc = hostOnTarget && c.cellBefore ? parseAreaText(c.cellBefore.split(':')[0]) : null;
  const ctx: AdjustContext = { hostSheet: c.where, targetSheet: o.sheet, tables, hostCell: hc ? { row: hc.r1, col: hc.c1 } : undefined };
  let ours: string | undefined, excel: string | undefined;
  // FILE form (what the reader/patcher work on)
  const fb = fileText(c.fileBefore), fa = fileText(c.fileAfter);
  if (fb != null && fa != null) {
    const got = adjustFormula(fb, op, ctx).formula;
    if (got !== fa) { probs.push(`FILE ${fb} -> excel ${fa} ours ${got}`); ours = got; excel = fa; }
  }
  // array formula range (legacy CSE arrays on the op sheet move with their cells)
  if (c.arrayBefore && c.arrayAfter && hostOnTarget) {
    const got = adjustSqref(c.arrayBefore, op);
    if (got !== c.arrayAfter) probs.push(`array ref ${c.arrayBefore} -> excel ${c.arrayAfter} ours ${got}`);
  }
  // UI form (Range.Formula / Formula2)
  for (const [b, a] of [[c.before, c.after], [c.before2, c.after2]] as const) {
    if (typeof b !== 'string' || typeof a !== 'string') continue;
    const got = '=' + adjustFormula(strip(b), op, ctx).formula;
    if (got !== a) { probs.push(`UI ${b} -> excel ${a} ours ${got}`); ours ??= got; excel ??= a; }
  }
  return probs.length ? { ok: false, detail: probs.join('; '), ours, excel } : { ok: true };
}

function checkName(c: R2bCase, o: R2bOp, op: StructOp, tables: TableRef[]): Outcome {
  if (c.before == null) return { ok: true, skip: 'name did not exist' };
  const got = '=' + adjustFormula(strip(c.before), op, { hostSheet: null, targetSheet: o.sheet, isName: true, tables }).formula;
  return got === c.after ? { ok: true } : { ok: false, detail: `${c.name}: ${c.before} -> excel ${c.after} ours ${got}`, ours: got, excel: c.after };
}

function oursRules(before: Rule, ruleSheet: string, o: R2bOp, op: StructOp, opts: AnchoredRuleOptions): Rule[] {
  const rule: AnchoredRule = { sqref: areasOf(before.sqref), formulas: before.formulas.map(strip) };
  if (ruleSheet.toLowerCase() !== o.sheet.toLowerCase()) {
    // rule on another sheet: the range stays, only references to the op sheet move
    return [{ sqref: before.sqref, formulas: rule.formulas.map((f) => adjustFormula(f, op, { hostSheet: ruleSheet, targetSheet: o.sheet }).formula) }];
  }
  return adjustAnchoredRule(rule, op, o.sheet, opts).map((r) => ({ sqref: r.sqref.map(renderArea).join(' '), formulas: r.formulas }));
}
const rectsOfRules = (rs: Rule[]) => rectsOf(rs.map((r) => ({ sqref: areasOf(r.sqref), formulas: r.formulas.map(strip) })));

function checkCf(c: R2bCase, o: R2bOp, op: StructOp): Outcome {
  const ours = oursRules(c.before as Rule, c.cfSheet!, o, op, { kind: 'cf', cfType: c.cfType });
  const d = diffRects(rectsOfRules(asRules(c.after)), rectsOfRules(ours));
  return d ? { ok: false, detail: `${JSON.stringify(c.before)} -> excel ${JSON.stringify(c.after)} ours ${JSON.stringify(ours)}: ${d}` } : { ok: true };
}

/**
 * R2b did not read the validated area from the validation itself: it tracked it through a defined name
 * (V_<id> = the area, see research/r2b/harness.ps1), i.e. plain range semantics. That can never show the growth of a
 * validation over inserted / moved cells, so this check compares with inheritFormat: false. The growth and the spread
 * of validations were measured on the validations themselves: tests/formula/cfSpread.test.ts (and the patcher goldens
 * g-ins-D, g-ins-row14, t-mov-L-H).
 */
function checkDv(c: R2bCase, o: R2bOp, op: StructOp): Outcome {
  const b = { sqref: c.before.range as string, formulas: [c.before.formula1 as string] };
  const ours = oursRules(b, c.dvSheet!, o, op, { kind: 'dv', type: 'list', inheritFormat: false });
  const excel = c.after?.range ? [{ sqref: c.after.range as string, formulas: [c.after.formula1 as string] }] : [];
  const d = diffRects(rectsOfRules(excel), rectsOfRules(ours));
  return d ? { ok: false, detail: `${JSON.stringify(c.before)} -> excel ${JSON.stringify(c.after)} ours ${JSON.stringify(ours)}: ${d}` } : { ok: true };
}

function checkRange(before: string, after: string | null, op: StructOp, what: string): Outcome {
  const a = parseAreaText(before);
  const r = a ? adjustArea(a, op) : null;
  const got = r ? renderArea(r) : null;
  return got === after ? { ok: true } : { ok: false, detail: `${what} ${before} -> excel ${after} ours ${got}`, ours: String(got), excel: String(after) };
}

function checkMerges(c: R2bCase, op: StructOp): Outcome {
  const got = (c.before as string[]).map((m) => { const r = adjustArea(parseAreaText(m)!, op); return r ? renderArea(r) : null; }).filter((x): x is string => x != null);
  const want = c.after as string[];
  const same = got.length === want.length && [...got].sort().join() === [...want].sort().join();
  return same ? { ok: true } : { ok: false, detail: `merges ${c.before} -> excel ${want} ours ${got}` };
}

function checkMergeTracked(c: R2bCase, o: R2bOp, op: StructOp): Outcome {
  const probs: string[] = [];
  const got = '=' + adjustFormula(strip(c.before.refersTo), op, { hostSheet: null, targetSheet: o.sheet, isName: true }).formula;
  if (got !== c.after.refersTo) probs.push(`refersTo ${c.before.refersTo} -> excel ${c.after.refersTo} ours ${got}`);
  if (c.after.mergeCells && c.before.mergeCells) {
    const r = checkRange(c.before.mergeArea, c.after.mergeArea, op, 'mergeArea');
    if (!r.ok) probs.push(r.detail!);
  }
  return probs.length ? { ok: false, detail: probs.join('; ') } : { ok: true };
}

function checkTable(c: R2bCase, op: StructOp): Outcome {
  const probs: string[] = [];
  const got = adjustTableRef(c.before.ref, op);
  if (got !== (c.after?.ref ?? null)) probs.push(`table ref ${c.before.ref} -> excel ${c.after?.ref} ours ${got}`);
  if (op.type === 'move' && c.after && got) {
    // table columns are reordered exactly like the sheet columns they sit in
    const a = parseAreaText(c.before.ref)!, na = parseAreaText(got)!;
    const p = movePerm(op.from, op.count, op.to);
    const names: string[] = [];
    (c.before.columns as Array<{ name: string }>).forEach((col, i) => { names[p(a.c1 + i) - na.c1] = col.name; });
    const want = (c.after.columns as Array<{ name: string }>).map((x) => x.name);
    if (names.join('|') !== want.join('|')) probs.push(`table columns excel ${want} ours ${names}`);
  }
  return probs.length ? { ok: false, detail: probs.join('; ') } : { ok: true };
}

/**
 * Source column of every column 1..n after the op (null = a column inserted by the op). Through mapIndex, i.e. the
 * same permutation the patcher uses for cells, widths and anchors.
 */
function sourceColumns(op: StructOp, n: number): Array<number | null> {
  const src: Array<number | null> = new Array(n).fill(null);
  for (let i = 1; i <= n + 64; i++) {
    const j = op.axis === 'col' ? mapIndex(i, op) : i;
    if (j != null && j >= 1 && j <= n) src[j - 1] = i;
  }
  return src;
}
/** row ops that touch row 1 change the marker row itself (not a column permutation) */
const touchesRow1 = (op: StructOp) => op.axis === 'row' && (op.type === 'move' ? op.from === 1 || op.to === 1 : op.at === 1);

/** row-1 markers "h<col letters>" in columns A..AD (1..30): the column permutation of the op */
const MARKER_COLS = 30;
function checkColumnOrder(c: R2bCase, op: StructOp): Outcome {
  if (touchesRow1(op)) return { ok: true, skip: 'op changes the marker row itself' };
  const before = c.before as Array<string | null>, after = c.after as Array<string | null>;
  const marker = (i: number) => (i <= before.length ? before[i - 1] : i <= MARKER_COLS ? 'h' + numToCol(i) : null);
  const want = sourceColumns(op, after.length).map((i) => (i == null ? null : marker(i)));
  return JSON.stringify(want) === JSON.stringify(after) ? { ok: true } : { ok: false, detail: `markers excel ${after} ours ${want}` };
}

/**
 * Column widths (COM ColumnWidth of columns 1..30, initially 6+col; columns past 30 have the default 8.38): widths
 * travel with moved columns, deleted columns take theirs away, and an inserted column takes the width of the column on
 * its LEFT (Excel's default "format same as left"); inserted at column A it gets the default width.
 */
const DEFAULT_WIDTH = 8.38;
function checkColumnWidths(c: R2bCase, op: StructOp): Outcome {
  const before = c.before as number[], after = c.after as number[];
  const width = (i: number) => (i <= before.length ? before[i - 1] : DEFAULT_WIDTH);
  const src = sourceColumns(op, after.length);
  const want = src.map((i) => {
    if (i != null) return width(i);
    if (op.type !== 'insert') return NaN;
    return op.at > 1 ? width(op.at - 1) : DEFAULT_WIDTH; // inserted: format of the left neighbour (before the op)
  });
  const same = want.length === after.length && want.every((w, k) => Math.abs(w - after[k]) < 0.01);
  return same ? { ok: true } : { ok: false, detail: `widths excel ${after} ours ${want} (k=${src.findIndex((_, k) => Math.abs(want[k] - after[k]) >= 0.01)})` };
}

interface ComTable { sheet: string; name: string; range: string; headers: string[]; columns: Array<{ name: string; formula: string }>; dataBody?: string }
/**
 * ListObjects read back through COM: range, header order (moved / deleted / inserted table columns) and the
 * calculated-column formula of the first data row (UI form "=[@点数]*2"; a deleted source column makes it #REF!).
 */
function checkTableCom(c: R2bCase, o: R2bOp, op: StructOp, tables: TableRef[]): Outcome {
  const probs: string[] = [];
  const afterList = (c.after ?? []) as ComTable[];
  for (const b of c.before as ComTable[]) {
    const a = afterList.find((x) => x.name === b.name) ?? null;
    const got = b.sheet === o.sheet ? adjustTableRef(b.range, op) : b.range;
    if (got !== (a?.range ?? null)) { probs.push(`${b.name} range ${b.range} -> excel ${a?.range} ours ${got}`); continue; }
    if (!a || !got) continue;
    const ba = parseAreaText(b.range)!, na = parseAreaText(got)!;
    if (a.dataBody !== renderArea({ ...na, r1: na.r1 + 1 })) probs.push(`${b.name} dataBody excel ${a.dataBody}`);
    // header order: every surviving column lands where the permutation puts it; the other slots are new columns
    const width = na.c2 - na.c1 + 1;
    const want: Array<string | null> = new Array(width).fill(null);
    b.headers.forEach((h, k) => {
      const col = ba.c1 + k;
      const nc = b.sheet !== o.sheet || op.axis === 'row' ? col : mapIndex(col, op);
      if (nc != null && nc >= na.c1 && nc <= na.c2) want[nc - na.c1] = h;
    });
    const bad = a.headers.length !== width || want.some((h, k) => (h == null ? b.headers.includes(a.headers[k]) : h !== a.headers[k]));
    if (bad) probs.push(`${b.name} headers excel ${a.headers} ours ${want.map((h) => h ?? '(new)')}`);
    // calculated column formula (first data row), where the column survives
    b.columns.forEach((col, k) => {
      if (!col.formula.startsWith('=')) return;
      const ak = a.columns.findIndex((x) => x.name === col.name);
      if (ak < 0) return;
      const hostCell = { row: ba.r1 + 1, col: ba.c1 + k };
      const f = '=' + adjustFormula(strip(col.formula), op, { hostSheet: b.sheet, targetSheet: o.sheet, tables, hostCell }).formula;
      if (f !== a.columns[ak].formula) probs.push(`${b.name}[${col.name}] ${col.formula} -> excel ${a.columns[ak].formula} ours ${f}`);
    });
  }
  return probs.length ? { ok: false, detail: probs.join('; ') } : { ok: true };
}

// ------------------------------------------------------------------------------------------------ run
interface Mismatch { id: string; kind: string; group: string; detail: string; ours?: string; excel?: string }
function runAll() {
  const counts: Record<string, { pass: number; fail: number; deviation: number; skipped: number }> = {};
  const mismatches: Mismatch[] = [];
  const skippedReasons: Record<string, number> = {};
  const bump = (k: string, f: 'pass' | 'fail' | 'deviation' | 'skipped') => { (counts[k] ??= { pass: 0, fail: 0, deviation: 0, skipped: 0 })[f]++; };
  for (const c of FX.cases) {
    const o = OPS.get(c.opId)!;
    if (o.noTruth) { bump(c.kind, 'skipped'); skippedReasons['no ground truth (op crashed)'] = (skippedReasons['no ground truth (op crashed)'] ?? 0) + 1; continue; }
    if (o.error) { bump(c.kind, 'skipped'); skippedReasons['Excel refused the op'] = (skippedReasons['Excel refused the op'] ?? 0) + 1; continue; }
    const op = toOp(o);
    const tables = FX.suites[o.suite]?.tables ?? [];
    let r: Outcome;
    switch (c.kind) {
      case 'formula': r = checkFormula(c, o, op, tables); break;
      case 'definedName': r = checkName(c, o, op, tables); break;
      case 'conditionalFormat': r = checkCf(c, o, op); break;
      case 'dataValidation': r = checkDv(c, o, op); break;
      case 'autoFilter': r = checkRange(c.before, c.after ?? null, op, 'autoFilter'); break;
      case 'mergedCells': r = checkMerges(c, op); break;
      case 'mergeTracked': r = checkMergeTracked(c, o, op); break;
      case 'table': r = checkTable(c, op); break;
      case 'tableCom': r = checkTableCom(c, o, op, tables); break;
      case 'columnOrder': r = checkColumnOrder(c, op); break;
      case 'columnWidths': r = checkColumnWidths(c, op); break;
      default: continue;
    }
    if (r.skip) { bump(c.kind, 'skipped'); skippedReasons[r.skip] = (skippedReasons[r.skip] ?? 0) + 1; continue; }
    if (r.ok) { bump(c.kind, 'pass'); continue; }
    const dev = R2B_DEVIATIONS[c.id];
    bump(c.kind, dev ? 'deviation' : 'fail');
    mismatches.push({ id: c.id, kind: c.kind, group: `${c.kind}|${o.suite}|${c.tag ?? ''}|${o.primitive}|${c.source}`, detail: r.detail!, ours: r.ours, excel: r.excel });
  }
  return { counts, mismatches, skippedReasons };
}

const RESULT = runAll();

if (process.env.R2B_REPORT) {
  const groups = new Map<string, Mismatch[]>();
  for (const m of RESULT.mismatches) { const g = groups.get(m.group) ?? []; g.push(m); groups.set(m.group, g); }
  console.log(JSON.stringify(RESULT.counts), RESULT.skippedReasons);
  for (const [g, ms] of groups) {
    console.log(`\n### ${g} (${ms.length})`);
    for (const m of ms.slice(0, Number(process.env.R2B_REPORT) || 6)) console.log(`  ${m.id}${R2B_DEVIATIONS[m.id] ? ` [${R2B_DEVIATIONS[m.id]}]` : ''}: ${m.detail}`);
  }
  console.log('\nALL IDS\n' + JSON.stringify(Object.fromEntries(RESULT.mismatches.map((m) => [m.id, R2B_DEVIATIONS[m.id] ?? '?']))));
}

/** Category-specific sanity checks of a documented deviation: our result must be the documented "intent". */
function deviationHolds(kind: DeviationKind, m: Mismatch): boolean {
  switch (kind) {
    // Excel's move = insert + cut + delete pushes refs to the last columns off the sheet; we keep them intact
    case 'sheet-edge-move': return !!m.ours && !m.ours.includes('#REF!') && !!m.excel && (m.excel.includes('#REF!') || /XF[A-D]/.test(m.excel));
    // Excel wraps a legacy implicit-intersection range that became #REF! in _xlfn.SINGLE / @; same value (#REF!)
    case 'legacy-single': return !!m.ours && !!m.excel && (m.excel === `_xlfn.SINGLE(${m.ours})` || m.excel === '=@' + m.ours.slice(1));
    // deleting the AutoFilter header row removes the filter (and its _FilterDatabase name): patcher's job
    case 'autofilter-removed': return true;
    // ListObject-adjacent reference quirks (table suite): we apply the plain-range model
    case 'table-quirk': return true;
    // Excel silently un-merges a merged area when a column is moved into its middle; the doc refuses such moves
    case 'merge-unmerged': return true;
  }
}

describe('R2b: excelRef vs a second, independent Excel ground truth', () => {
  it('uses the whole fixture', () => {
    expect(FX.cases.length).toBeGreaterThan(20000);
    const total = Object.values(RESULT.counts).reduce((s, x) => s + x.pass + x.fail + x.deviation, 0);
    expect(total).toBeGreaterThan(19000);
  });
  for (const kind of ['formula', 'definedName', 'conditionalFormat', 'dataValidation', 'autoFilter', 'mergedCells', 'mergeTracked', 'table',
    'tableCom', 'columnOrder', 'columnWidths']) {
    it(`${kind}: every case equals Excel except documented deviations`, () => {
      const fails = RESULT.mismatches.filter((m) => m.kind === kind && !R2B_DEVIATIONS[m.id]).map((m) => `${m.id}: ${m.detail}`);
      expect(fails.slice(0, 40)).toEqual([]);
      expect(RESULT.counts[kind]?.pass ?? 0).toBeGreaterThan(0);
    });
  }
  it('every documented deviation still deviates, for the documented reason', () => {
    const seen = new Map(RESULT.mismatches.map((m) => [m.id, m]));
    const stale = Object.keys(R2B_DEVIATIONS).filter((id) => !seen.has(id));
    expect(stale).toEqual([]);
    const wrong = [...seen.values()].filter((m) => R2B_DEVIATIONS[m.id] && !deviationHolds(R2B_DEVIATIONS[m.id], m)).map((m) => `${m.id} [${R2B_DEVIATIONS[m.id]}]: ${m.detail}`);
    expect(wrong).toEqual([]);
  });
});
