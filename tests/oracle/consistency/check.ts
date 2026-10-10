/**
 * The judge of the consistency oracle (task V1): compares the Excel-recalculated dumps of the baseline and of the two engines' outputs.
 *
 *  - baseline  : the premerged package (+ typed texts at the original positions + converted values), no layout operation at all
 *  - js / xl   : the output of replayLog (built-in engine) / of server/excel/excel-runner.ps1 (real Excel) for the same log
 * Cells are found again by (sheet, row, COLUMN ID): the id of a column survives moves, the position does not.
 *
 * Findings (class = the first word):
 *   open-*            a package that needs repair / cannot be opened
 *   value-changed     a cell has another value than in the baseline and nothing explains it (the user's complaint)
 *   new-error         an error value that the baseline does not have
 *   lookup-silent     a lookup-by-column-number cell changed without the doc having warned
 *   engine-value      the two engines' outputs differ in a value (same position)
 *   engine-formula    ... in a formula (Excel's Formula text of the cell)
 *   engine-package    ... in names / CF / DV / tables / merges (tools/lib/package-diff)
 *   layout-header     the header text at the final position of a column id is not the one the id had
 *   inserted-not-blank, ref-identity, doc-vs-replay, runner-*  (see below)
 * Explained changes (not failures, counted): range-span (a multi-column range now covers other columns: Excel changes the value),
 * lookup-warned, positional-fn, cascade (reads an explained cell), convert-blank (the doc converted a formula with no stored result).
 */
import { tokenize } from '../../../src/core/formula/excelRef';
import type { DefinedName, TableInfo } from '../../../src/core/types';
import { areaSigs } from './ops';
import type { MergedDoc } from '../../../src/core/types';
import { norm, type BookDump, type SheetDump, type XVal } from './excel';
import { colLetter } from './lib';

export interface DocInfo {
  names: DefinedName[];
  tables: TableInfo[];
  /** column ids in order */
  ids: string[];
}

export interface ScenarioRec {
  id: string;
  set: string;
  variant: string;
  seed: number;
  mode: 'inv' | 'del';
  allowAffecting: boolean;
  mergedSheet: string;
  headerRow: number;
  firstDataRow: number;
  lastDataRow: number;
  initial: DocInfo;
  final: DocInfo;
  deleted: string[];
  converted: Array<{ sheet: string; id?: string; row: number; col?: number; value: string | number | boolean | null; blank: boolean }>;
  trail: Array<{ desc: string; ok: boolean; refused?: string; warnings: string[] }>;
  lookupWarned: boolean;
  affecting: boolean;
  opsKinds: Record<string, number>;
  logLength: number;
  /** the baseline file is the premerged package itself (no edits, no conversions) */
  baselineIsPre: boolean;
  files: { pre: string; base: string; js: string; xl: string };
  /** sheets that hold a PivotTable: their cells are rendered by Excel (reported as their own class, never as a formula problem) */
  pivotSheets: string[];
  /**
   * the pivot sheets can be compared with the baseline: no typed value / converted formula changed the pivot's data (the built-in
   * engine marks a changed pivot source refreshOnLoad, so a typed value shows up in the pivot, which the Excel engine's stale pivot does not)
   */
  pivotComparable: boolean;
  /** CF / DV rules (merged sheet and other sheets) whose formulas contain #REF! before / after the operations */
  refRules?: { initial: number; final: number };
  docVsReplay: string[];
  /** diffs of the two outputs read by the project's package diff (names / CF / DV / tables / merges ...), filled by the judge */
  skipped?: string;
}

export interface Finding {
  scenario: string;
  cls: string;
  detail: string;
}

export interface Stats {
  cells: number;
  formulaCells: number;
  mismatches: number;
  explained: Record<string, number>;
  /** first examples of every explanation (for the report) */
  examples?: Record<string, string[]>;
}

const num = (v: unknown): v is number => typeof v === 'number';
export function sameValue(a: unknown, b: unknown): boolean {
  if ((a === null || a === '') && (b === null || b === '')) return true;
  if (num(a) && num(b)) return Math.abs(a - b) <= 1e-9 * Math.max(1, Math.abs(a));
  // Excel keeps the new #CALC! error as #VALUE! in a file
  if ((a === '#CALC!' && b === '#VALUE!') || (a === '#VALUE!' && b === '#CALC!')) return true;
  return a === b;
}

const isErr = (v: unknown): boolean => typeof v === 'string' && /^#(NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|GETTING_DATA|SPILL!|CALC!|FIELD!)$/.test(v);

const VOLATILE = /\b(TODAY|NOW|RAND|RANDBETWEEN|RANDARRAY)\s*\(/i;
const POSITIONAL = /\b(COLUMN|COLUMNS|OFFSET|INDIRECT|ADDRESS|CELL|INFO|FORMULATEXT|SHEET|SHEETS)\s*\(/i;
const LOOKUPFN = /\b(VLOOKUP|HLOOKUP|INDEX|CHOOSE|CHOOSECOLS|XLOOKUP|MATCH|XMATCH)\s*\(/i;

function sheetOf(d: BookDump, name: string): SheetDump | undefined {
  return d.sheets.find((s) => s.name.toLowerCase() === name.toLowerCase());
}
const at = <T,>(rows: T[][] | undefined, r: number, c: number): T | undefined => rows?.[r - 1]?.[c - 1];

interface Cellref {
  sheet: string;
  row: number;
  /** column in the BASELINE */
  col: number;
  id?: string;
}

interface Mis extends Cellref {
  bv: unknown;
  /** value in the output under judgement */
  av: unknown;
  bf: string;
  af: string;
}

/** the cells of a (possibly many-cell) reference token that lie in `explained` */
function dependsOn(bf: string, host: Cellref, explained: Map<string, Set<number>>, info: DocInfo, mergedSheet: string, depth = 0): boolean {
  if (!bf) return false;
  const hit = (sheet: string, r1: number, c1: number, r2: number, c2: number): boolean => {
    const set = explained.get(sheet.toLowerCase());
    if (!set) return false;
    for (const k of set) {
      const r = Math.floor(k / 20000);
      const c = k % 20000;
      if (r >= r1 && r <= r2 && c >= c1 && c <= c2) return true;
    }
    return false;
  };
  for (const t of tokenize(bf.replace(/^=/, ''))) {
    if (t.type === 'ref' && t.area) {
      if (t.prefix?.book) continue;
      const sheet = t.prefix?.sheet ?? host.sheet;
      const a = t.area;
      const r1 = a.kind === 'cols' ? 1 : a.r1;
      const r2 = a.kind === 'cols' ? 1048576 : a.r2;
      const c1 = a.kind === 'rows' ? 1 : a.c1;
      const c2 = a.kind === 'rows' ? 16384 : a.c2;
      if (hit(sheet, r1, c1, r2, c2)) return true;
    } else if (t.type === 'name' && depth < 3) {
      const bare = t.prefix ? t.text.slice(t.prefix.text.length) : t.text;
      const n = info.names.find((x) => x.name.toLowerCase() === bare.toLowerCase() && x.scope !== null && x.scope.toLowerCase() === host.sheet.toLowerCase()) ?? info.names.find((x) => x.name.toLowerCase() === bare.toLowerCase() && x.scope === null);
      if (n && dependsOn(n.refersTo, { ...host, sheet: n.scope ?? mergedSheet }, explained, info, mergedSheet, depth + 1)) return true;
    } else if (t.type === 'sref') {
      const m = /^([^[\]]*)\[/.exec(t.text);
      const tb = info.tables.find((x) => x.name.toLowerCase() === (m?.[1] ?? '').toLowerCase()) ?? (m?.[1] ? undefined : info.tables[0]);
      if (!tb) continue;
      const box = /^\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)$/.exec(tb.ref);
      if (!box) continue;
      const cn = (s: string) => [...s].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
      const rowOnly = /#This Row|^\[@|\[@/i.test(t.text);
      if (hit(mergedSheet, rowOnly ? host.row : Number(box[2]), cn(box[1]), rowOnly ? host.row : Number(box[4]), cn(box[3]))) return true;
    }
  }
  return false;
}

/** a sheet with a circular reference: Excel gives its cells no defined value (order of evaluation) */
export const circular = (d: BookDump): string => d.sheets.filter((x) => x.circ).map((x) => `${x.name}!${x.circ}`).join(', ');

export interface JudgeInput {
  rec: ScenarioRec;
  base: BookDump;
  /** the output under judgement */
  out: BookDump;
  engine: 'js' | 'xl';
  /** value dump of the premerged package (when the baseline differs from it): used to look at converted cells */
  pre?: BookDump;
}

export interface JudgeResult {
  findings: Finding[];
  stats: Stats;
}

/** Judges one engine's output against the baseline. */
export function judgeOutput(inp: JudgeInput): JudgeResult {
  const { rec, base, out, engine } = inp;
  const findings: Finding[] = [];
  const stats: Stats = { cells: 0, formulaCells: 0, mismatches: 0, explained: {} };
  const add = (cls: string, detail: string) => findings.push({ scenario: rec.id, cls: `${cls}-${engine}`, detail });
  stats.examples = {};
  let curMis: Mis | null = null;
  const bump = (k: string) => {
    stats.explained[k] = (stats.explained[k] ?? 0) + 1;
    const ex = (stats.examples![k] ??= []);
    if (ex.length < 3 && curMis) ex.push(`${curMis.sheet}!${colLetter(curMis.col)}${curMis.row}${curMis.id ? ` (${curMis.id})` : ''}: ${JSON.stringify(curMis.bv)} [${curMis.bf}] -> ${JSON.stringify(curMis.av)} [${curMis.af}]`);
  };
  if (out.open !== 'normal') {
    add('open', `${out.open}: ${out.log}`);
    return { findings, stats };
  }
  // Excel moved a column INTO the range of a formula that lives in (or depends on) that column: a circular reference, which has no value
  const loops = circular(out) || circular(base);
  if (loops) {
    findings.push({ scenario: rec.id, cls: `circular-${engine}`, detail: `circular reference at ${loops}${circular(base) ? ' (already in the baseline)' : ''}: the values of this output are not judged` });
    return { findings, stats };
  }
  const merged = rec.mergedSheet;
  const initialPos = new Map(rec.initial.ids.map((id, i) => [id, i + 1]));
  const finalPos = new Map(rec.final.ids.map((id, i) => [id, i + 1]));
  const survivors = rec.initial.ids.filter((id) => finalPos.has(id));
  const mis: Mis[] = [];
  const mapCol = (sheet: string, col: number): number | null => {
    if (sheet.toLowerCase() !== merged.toLowerCase()) return col;
    const id = rec.initial.ids[col - 1];
    return id === undefined ? null : finalPos.get(id) ?? null;
  };
  const dumpSheets = new Set(base.sheets.map((s) => s.name.toLowerCase()));
  for (const bs of base.sheets) {
    const os = sheetOf(out, bs.name);
    if (!os) {
      add('open', `sheet ${bs.name} missing in the output`);
      continue;
    }
    const isMerged = bs.name.toLowerCase() === merged.toLowerCase();
    const rows = Math.max(bs.rows, os.rows);
    const colsBase = isMerged ? rec.initial.ids.length : bs.cols;
    for (let c = 1; c <= Math.max(colsBase, isMerged ? 0 : os.cols); c++) {
      const oc = mapCol(bs.name, c);
      if (oc === null) continue; // a deleted column
      for (let r = 1; r <= rows; r++) {
        const bv = norm(at(bs.v, r, c));
        const av = norm(at(os.v, r, oc));
        const bf = at(bs.f, r, c) ?? '';
        const af = at(os.f, r, oc) ?? '';
        stats.cells++;
        if (bf) stats.formulaCells++;
        if (VOLATILE.test(bf)) continue;
        if (!sameValue(bv, av)) {
          stats.mismatches++;
          mis.push({ sheet: bs.name, row: r, col: c, id: isMerged ? rec.initial.ids[c - 1] : undefined, bv, av, bf, af });
        }
      }
    }
  }
  void dumpSheets;
  // ---- explain: roots
  const explained = new Map<string, Set<number>>();
  const mark = (m: Mis) => {
    // an explained change that ends in an error value is counted apart: Excel itself shrank a lookup range or the like, but it is worth a look
    if (isErr(m.av) && !isErr(m.bv)) {
      const key = `(ends in ${m.av})`;
      stats.explained[key] = (stats.explained[key] ?? 0) + 1;
      const ex = (stats.examples![key] ??= []);
      if (ex.length < 2) ex.push(`${m.sheet}!${colLetter(m.col)}${m.row}: ${JSON.stringify(m.bv)} [${m.bf}] -> ${m.av} [${m.af}]`);
    }
    const k = m.sheet.toLowerCase();
    let s = explained.get(k);
    if (!s) explained.set(k, (s = new Set()));
    s.add(m.row * 20000 + m.col);
  };
  const rest: Mis[] = [];
  // #REF! and #NAME? are never an acceptable outcome of a column operation (the doc converts what would break), explained or not — except the
  // documented lookup case: VLOOKUP(x,A:D,4,..) whose range lost a column returns #REF! as a VALUE (the formula text has no #REF!), and the doc warned
  const lookupCase = (m: Mis) => rec.lookupWarned && LOOKUPFN.test(m.bf) && !/#REF!/i.test(m.af);
  // ... and an error that is only PROPAGATED (AVERAGE over a range that contains the VLOOKUP's #REF!) belongs to the cell it comes from
  const readsError = (m: Mis): boolean => {
    for (const t of tokenize(m.af.replace(/^=/, ''))) {
      if (t.type !== 'ref' || !t.area || t.prefix?.book) continue;
      const sh = sheetOf(out, t.prefix?.sheet ?? m.sheet);
      if (!sh) continue;
      const a = t.area;
      const r1 = a.kind === 'cols' ? 1 : a.r1;
      const r2 = a.kind === 'cols' ? sh.rows : Math.min(a.r2, sh.rows);
      const c1 = a.kind === 'rows' ? 1 : a.c1;
      const c2 = a.kind === 'rows' ? sh.cols : Math.min(a.c2, sh.cols);
      for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) if (isErr(norm(at(sh.v, r, c))) && !(sh.name === m.sheet && r === m.row && c === (isMergedSheet(m.sheet) ? finalPos.get(rec.initial.ids[m.col - 1]!) ?? -1 : m.col))) return true;
    }
    return false;
  };
  const isMergedSheet = (n: string) => n.toLowerCase() === merged.toLowerCase();
  const refErrors = mis.filter((m) => (m.av === '#REF!' || m.av === '#NAME?') && m.bv !== m.av && !lookupCase(m) && !readsError(m));
  for (const m of refErrors.slice(0, 5)) add('new-ref-error', `${m.sheet}!${colLetter(m.col)}${m.row}${m.id ? ` (${m.id})` : ''}: baseline ${JSON.stringify(m.bv)} [${m.bf || 'constant'}] -> ${m.av} [${m.af || 'constant'}]`);
  if (refErrors.length > 5) add('new-ref-error', `... ${refErrors.length - 5} more`);
  // the doc's view of the spans (before / after) decides whether a range changed its columns
  const finalDoc = { sheetName: merged, columns: rec.final.ids.map((id) => ({ id })), features: { tables: rec.final.tables } } as unknown as MergedDoc;
  const initialDoc = { sheetName: merged, columns: rec.initial.ids.map((id) => ({ id })), features: { tables: rec.initial.tables } } as unknown as MergedDoc;
  const spanChanged = (m: Mis): boolean => {
    if (!m.bf || !m.af) return false;
    const host = m.sheet;
    const a = areaSigs(initialDoc, m.bf.replace(/^=/, ''), host);
    const b = areaSigs(finalDoc, m.af.replace(/^=/, ''), host);
    if (!a.length && !b.length) return false;
    if (a.length !== b.length) return true;
    return a.some((x, i) => x !== b[i]);
  };
  const usesChangedName = (m: Mis): boolean => {
    for (const t of tokenize(m.bf.replace(/^=/, ''))) {
      if (t.type !== 'name') continue;
      const bare = t.prefix ? t.text.slice(t.prefix.text.length) : t.text;
      const n0 = rec.initial.names.find((x) => x.name.toLowerCase() === bare.toLowerCase());
      const n1 = rec.final.names.find((x) => x.name.toLowerCase() === bare.toLowerCase());
      if (!n0 || !n1) continue;
      const a = areaSigs(initialDoc, n0.refersTo, n0.scope ?? merged);
      const b = areaSigs(finalDoc, n1.refersTo, n1.scope ?? merged);
      if (a.length !== b.length || a.some((x, i) => x !== b[i])) return true;
    }
    return false;
  };
  const isPivot = (sheet: string) => rec.pivotSheets.some((x) => x.toLowerCase() === sheet.toLowerCase());
  const pivotMis = mis.filter((m) => isPivot(m.sheet));
  if (pivotMis.length && rec.pivotComparable) add('pivot-value-changed', `${pivotMis.length} cells of the pivot sheet(s) ${[...new Set(pivotMis.map((m) => m.sheet))].join(', ')} differ from the baseline, e.g. ${pivotMis[0].sheet}!${colLetter(pivotMis[0].col)}${pivotMis[0].row}: ${JSON.stringify(pivotMis[0].bv)} -> ${JSON.stringify(pivotMis[0].av)}`);
  for (const m of mis) {
    curMis = m;
    if (isPivot(m.sheet)) continue;
    if (m.bf && POSITIONAL.test(m.bf)) { bump('positional-fn'); mark(m); continue; }
    if (m.bf && spanChanged(m)) { bump('range-span'); mark(m); continue; }
    if (m.bf && usesChangedName(m)) { bump('range-span'); mark(m); continue; }
    if (m.bf && LOOKUPFN.test(m.bf) && rec.lookupWarned) { bump('lookup-warned'); mark(m); continue; }
    rest.push(m);
  }
  // ---- cascade: a cell that reads an explained cell
  let pending = rest.slice();
  for (let grew = true; grew; ) {
    grew = false;
    const next: Mis[] = [];
    for (const m of pending) {
      curMis = m;
      if (m.bf && dependsOn(m.bf, m, explained, rec.initial, merged)) {
        bump('cascade');
        mark(m);
        grew = true;
      } else next.push(m);
    }
    pending = next;
  }
  // ---- what is left is a failure
  for (const m of pending) {
    const where = `${m.sheet}!${colLetter(m.col)}${m.row}${m.id ? ` (${m.id})` : ''}`;
    const detail = `${where}: baseline ${JSON.stringify(m.bv)} [${m.bf || 'constant'}] -> ${JSON.stringify(m.av)} [${m.af || 'constant'}]`;
    if (isErr(m.av) && !isErr(m.bv)) add('new-error', detail);
    else if (m.bf && LOOKUPFN.test(m.bf) && !rec.lookupWarned && /VLOOKUP|HLOOKUP|CHOOSE/i.test(m.bf)) add('lookup-silent', detail);
    else add('value-changed', detail);
  }
  // ---- the layout bookkeeping: a column id finds its header again; inserted columns are blank
  const ms = sheetOf(out, merged);
  const mb = sheetOf(base, merged);
  if (ms && mb) {
    for (const id of survivors) {
      const h0 = norm(at(mb.v, rec.headerRow, initialPos.get(id)!));
      const h1 = norm(at(ms.v, rec.headerRow, finalPos.get(id)!));
      if (!sameValue(h0, h1)) add('layout-header', `column ${id}: header ${JSON.stringify(h0)} at ${colLetter(initialPos.get(id)!)} -> ${JSON.stringify(h1)} at ${colLetter(finalPos.get(id)!)}`);
    }
    for (const id of rec.final.ids.filter((x) => /^n\d+$/.test(x))) {
      const c = finalPos.get(id)!;
      // (an inserted column inside a Table gets its header cell from Excel: 列1, 列2 …)
      for (let r = rec.final.tables.length ? rec.headerRow + 1 : 1; r <= ms.rows; r++) {
        const v = norm(at(ms.v, r, c));
        if (v !== null && v !== '') {
          add('inserted-not-blank', `inserted column ${id} at ${colLetter(c)}${r} holds ${JSON.stringify(v)}`);
          break;
        }
      }
    }
  }
  return { findings, stats };
}

/** The two engines against each other: every cell of every sheet at the same position (value and Excel's formula text). */
export function judgeParity(rec: ScenarioRec, js: BookDump, xl: BookDump): Finding[] {
  const findings: Finding[] = [];
  const add = (cls: string, detail: string) => findings.push({ scenario: rec.id, cls, detail });
  if (js.open !== 'normal' || xl.open !== 'normal') return findings; // reported by judgeOutput
  if (circular(js) || circular(xl)) return findings; // no defined values (reported by judgeOutput)
  let nv = 0;
  let nf = 0;
  let np = 0;
  for (const a of js.sheets) {
    const b = sheetOf(xl, a.name);
    if (!b) {
      add('engine-value', `sheet ${a.name} exists in the built-in output only`);
      continue;
    }
    const rows = Math.max(a.rows, b.rows);
    const cols = Math.max(a.cols, b.cols);
    for (let r = 1; r <= rows; r++) {
      for (let c = 1; c <= cols; c++) {
        const va = norm(at(a.v, r, c));
        const vb = norm(at(b.v, r, c));
        const fa = at(a.f, r, c) ?? '';
        const fb = at(b.f, r, c) ?? '';
        if (!sameValue(va, vb) && !VOLATILE.test(fa) && rec.pivotSheets.some((x) => x.toLowerCase() === a.name.toLowerCase())) {
          if (rec.pivotComparable && np++ < 3) add('pivot-engine-value', `${a.name}!${colLetter(c)}${r}: built-in ${JSON.stringify(va)} | Excel ${JSON.stringify(vb)}`);
        } else if (!sameValue(va, vb) && !VOLATILE.test(fa)) {
          if (nv++ < 6) add('engine-value', `${a.name}!${colLetter(c)}${r}: built-in ${JSON.stringify(va)} [${fa}] | Excel ${JSON.stringify(vb)} [${fb}]`);
        }
        if (fa !== fb) {
          if (nf++ < 6) add('engine-formula', `${a.name}!${colLetter(c)}${r}: built-in [${fa}] | Excel [${fb}]`);
        }
      }
    }
  }
  if (nv > 6) add('engine-value', `... ${nv - 6} more value differences`);
  if (nf > 6) add('engine-formula', `... ${nf - 6} more formula differences`);
  return findings;
}

/**
 * Single-cell / single-column references must follow their column ids: a baseline formula's reference to column X at position p must,
 * in the output, point at the position of X. (Multi-column ranges follow Excel's own span rules and are judged by value.)
 */
export function judgeRefIdentity(rec: ScenarioRec, base: BookDump, out: BookDump, engine: 'js' | 'xl'): Finding[] {
  const findings: Finding[] = [];
  if (out.open !== 'normal') return findings;
  const merged = rec.mergedSheet;
  const finalPos = new Map(rec.final.ids.map((id, i) => [id, i + 1]));
  let n = 0;
  const mismatch = (where: string, msg: string) => {
    if (n++ < 8) findings.push({ scenario: rec.id, cls: `ref-identity-${engine}`, detail: `${where}: ${msg}` });
  };
  for (const bs of base.sheets) {
    const os = sheetOf(out, bs.name);
    if (!os) continue;
    const isMerged = bs.name.toLowerCase() === merged.toLowerCase();
    for (let c = 1; c <= (isMerged ? rec.initial.ids.length : bs.cols); c++) {
      let oc = c;
      if (isMerged) {
        const p = finalPos.get(rec.initial.ids[c - 1]);
        if (p === undefined) continue;
        oc = p;
      }
      for (let r = 1; r <= bs.rows; r++) {
        const bf = at(bs.f, r, c);
        const of = at(os.f, r, oc);
        if (!bf || !of || !/^=/.test(bf)) continue;
        const tb = tokenize(bf.slice(1)).filter((t) => t.type === 'ref' && t.area && !t.prefix?.book);
        const to = tokenize(of.slice(1)).filter((t) => t.type === 'ref' && t.area && !t.prefix?.book);
        if (tb.length !== to.length) continue; // a reference that became #REF! etc. is reported by the value check
        tb.forEach((x, i) => {
          const y = to[i];
          const a = x.area!;
          const b = y.area!;
          const sheetB = (x.prefix?.sheet ?? bs.name).toLowerCase();
          if (sheetB !== merged.toLowerCase()) return;
          if (!(a.kind === 'cell' || (a.kind === 'cols' && a.c1 === a.c2) || (a.kind === 'area' && a.c1 === a.c2))) return;
          if (!(b.kind === 'cell' || (b.kind === 'cols' && b.c1 === b.c2) || (b.kind === 'area' && b.c1 === b.c2))) {
            mismatch(`${bs.name}!${colLetter(c)}${r}`, `single-column reference ${x.text} became ${y.text}`);
            return;
          }
          const idB = rec.initial.ids[a.c1 - 1];
          const idO = rec.final.ids[b.c1 - 1];
          if (idB !== idO && !(idB === undefined)) mismatch(`${bs.name}!${colLetter(c)}${r}`, `${x.text} (${idB}) -> ${y.text} (${idO}); formula ${bf} -> ${of}`);
          if (a.kind !== 'cols' && a.r1 !== b.r1) mismatch(`${bs.name}!${colLetter(c)}${r}`, `row changed: ${x.text} -> ${y.text}`);
        });
      }
    }
  }
  return findings;
}

export function dumpFormulaCount(d: BookDump): number {
  let n = 0;
  for (const s of d.sheets) for (const row of s.f) for (const f of row) if (f) n++;
  return n;
}

export type { XVal };

/**
 * deleteColumns(convertDependents) writes the doc's CACHED value over every formula that would become #REF!. The value must be what Excel
 * computes for that cell (the premerged package, recalculated): a different number is a stale value (class convert-stale); a blank written
 * where Excel has a value is the documented case of a formula that has no stored result (the doc warns: 「計算結果が保存されていなかった
 * ため空欄になりました」) and is counted as convert-blank — it still loses the value.
 */
export function judgeConverted(rec: ScenarioRec, pre: BookDump): { findings: Finding[]; blank: number; stale: number; ok: number } {
  const findings: Finding[] = [];
  let blank = 0;
  let stale = 0;
  let ok = 0;
  const initialPos = new Map(rec.initial.ids.map((id, i) => [id, i + 1]));
  for (const c of rec.converted) {
    const sheet = sheetOf(pre, c.sheet);
    if (!sheet) continue;
    const col = c.id !== undefined ? initialPos.get(c.id) : c.col;
    if (col === undefined) continue;
    const expected = norm(at(sheet.v, c.row, col));
    const written = c.blank ? null : c.value;
    if (sameValue(expected, written)) { ok++; continue; }
    const where = `${c.sheet}!${colLetter(col)}${c.row}${c.id ? ` (${c.id})` : ''}: Excel computes ${JSON.stringify(expected)}, the doc wrote ${c.blank ? 'a blank' : JSON.stringify(written)}`;
    if (c.blank && expected !== null && expected !== '') {
      blank++;
      if (blank <= 3) findings.push({ scenario: rec.id, cls: 'convert-blank', detail: where });
    } else {
      stale++;
      if (stale <= 3) findings.push({ scenario: rec.id, cls: 'convert-stale', detail: where });
    }
  }
  return { findings, blank, stale, ok };
}

/**
 * Classes that are REPORTED (RESULTS.md counts them) but are not failures of the consistency requirement:
 *   skipped / circular      the scenario could not be judged (a doc refusal at set-up; Excel gives circular references no defined value)
 *   runner-fallback         the Excel engine refused the log with a code the app answers by switching to the built-in engine
 *   convert-blank           deleteColumns(convertDependents) blanked a formula that has no stored result (the doc warns, documented)
 *   cf-quirk                Excel's documented conditional-format quirk (absolute column reference + a move inside the range)
 *   ref-error-announced     a defined name / CF / DV rule ends with #REF! because its only column was deleted (they cannot be converted to
 *                           values); the doc said so. Without the warning it is the failure class ref-error-silent
 *   pivot-value-changed     the PivotTable of the kept 集計/ピボット sheets shows other numbers than the baseline although BOTH engines
 *                           agree: Excel itself changed it (a moved source column left the pivot's source range)
 *   pivot-engine-package-*  Excel autofits the columns of a refreshed pivot sheet
 */
export const INFO_CLASSES = /^(skipped|runner-fallback|convert-blank|cf-quirk|circular|ref-error-announced|pivot-value-changed|pivot-engine-package-(cols|widths))/;

/**
 * Defects the oracle FOUND that are not fixed because the code belongs to someone else (reported in RESULTS.md with a minimal repro). They are
 * shown as "open" in the summary and do not fail the run; delete the entry when the owner has fixed it (the run then shows none of it).
 *   (none)  append-range-running was FIXED in round 9 (patch/formula.ts extendArea: a formula inside the data rows keeps a relative range
 *           that ends at its own row; regress.test.ts)
 */
export const KNOWN_OPEN = /^(?!)/;
