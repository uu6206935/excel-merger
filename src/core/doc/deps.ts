/**
 * Dependency checks before / after structural ops:
 *  - findBreaks: dry run of delete ops (excelRef semantics via the compiled fast path) over every formula of the
 *    merged sheet and the other sheets, every defined name and every CF / DV rule: what would become #REF!.
 *  - lookupIssues: formulas that address a range's columns BY NUMBER (VLOOKUP column index, INDEX / MATCH /
 *    OFFSET positions, CHOOSECOLS ...) whose range changes its column order / content under the ops. Excel does not
 *    adjust those numbers, so the formula silently starts returning another column.
 */
import { adjustFormula, lookupWarnings, tokenize } from '../formula/excelRef';
import type { LookupWarning, StructOp } from '../formula/excelRef';
import { toUiForm } from '../formula/forms';
import type { DefinedName, FeatureRule, MergedDoc } from '../types';
import type { Compiled } from './compiled';
import { adjustRule, formulaCtxFor, wouldBreak } from './struct';
import type { FormulaCtx } from './struct';
import { addr, cachedOf, compiledAt, compiledCol, lc } from './util';

export interface BreakList {
  /** merged-sheet formula cells: 0-based column index (current order) + 1-based row */
  cells: Array<{ index: number; row: number }>;
  /** indices into doc.externalFormulas */
  external: number[];
  names: DefinedName[];
  rules: FeatureRule[];
  externalRules: Array<{ sheet: string; rule: FeatureRule }>;
}

export const breakCount = (b: BreakList): number => b.cells.length + b.external.length + b.names.length + b.rules.length + b.externalRules.length;

/** "集計!B3" for messages (sheet names are shown as they are, without formula quoting). */
export const placeText = (sheet: string, col: number, row: number): string => `${sheet}!${addr(col, row)}`;

const hasRefError = (s: string): boolean => /#REF!/i.test(s);

/** Everything that the ops would turn into #REF! (cells in columns listed in `gone` are ignored: they disappear). */
export function findBreaks(doc: MergedDoc, ops: readonly StructOp[], gone: ReadonlySet<number>): BreakList {
  const ctx = formulaCtxFor(doc, ops);
  const out: BreakList = { cells: [], external: [], names: [], rules: [], externalRules: [] };
  // merged sheet cells
  doc.columns.forEach((col, index) => {
    if (gone.has(index)) return;
    const cells = col.cells;
    const cc = compiledCol(cells);
    let prev: Compiled | null = null;
    let prevRow = 0;
    for (let r = 0; r < cells.length; r++) {
      const cell = cells[r];
      if (!cell || cell.f === undefined) continue;
      const c = compiledAt(cc, r, cell.f, prev, r - prevRow);
      prev = c;
      prevRow = r;
      if (wouldBreak(ctx, cell.f, c, doc.sheetName)) out.cells.push({ index, row: r + 1 });
    }
  });
  // other sheets
  const ec = compiledCol(doc.externalFormulas);
  doc.externalFormulas.forEach((e, i) => {
    const c = compiledAt(ec, i, e.f);
    if (wouldBreak(ctx, e.f, c, e.sheet)) out.external.push(i);
  });
  // names (relative parts are offsets: isName)
  for (const n of doc.names) {
    let cur = n.refersTo;
    let broke = false;
    ops.forEach((op, i) => {
      const r = adjustFormula(cur, op, { hostSheet: n.scope, targetSheet: doc.sheetName, isName: true, tables: ctx.tableRefsAt[i] });
      if (r.createdRefError) broke = true;
      cur = r.formula;
    });
    if (broke && !isInternalName(n)) out.names.push(n);
  }
  // CF / DV on the merged sheet (fragments; a fragment whose cells are all deleted simply disappears)
  for (const rule of doc.features.rules) {
    if (rule.formulas.some(hasRefError)) continue;
    let cur: FeatureRule[] = [rule];
    for (const op of ops) cur = cur.flatMap((r) => { const x = adjustRule(r, op, doc.sheetName); return Array.isArray(x) ? x : [x]; });
    if (cur.some((r) => r.formulas.some(hasRefError))) out.rules.push(rule);
  }
  // CF / DV of other sheets
  for (const er of doc.externalRules) {
    let broke = false;
    for (const f of er.rule.formulas) {
      let cur = f;
      ops.forEach((op, i) => {
        const r = adjustFormula(cur, op, { hostSheet: er.sheet, targetSheet: doc.sheetName, tables: ctx.tableRefsAt[i] });
        if (r.createdRefError) broke = true;
        cur = r.formula;
      });
    }
    if (broke) out.externalRules.push(er);
  }
  // formulas that use a name which becomes #REF! evaluate to #REF! too (they are converted like direct references);
  // readers of a broken array / spill master (`master#`) break with it. The two feed each other, so iterate.
  for (let round = 0; round < 4; round++) {
    const before = breakCount(out);
    if (out.names.length) addNameUsers(doc, out, gone);
    addSpillReaders(doc, out, gone);
    if (breakCount(out) === before) break;
  }
  return out;
}

// ---------------------------------------------------------------------------------------------
// spill readers
// ---------------------------------------------------------------------------------------------
const ANCHOR = /anchorarray/i;

/**
 * Cells a formula reads through the spill operator (`A1#` = _xlfn.ANCHORARRAY(A1)): sheet as written (null = the
 * host's own sheet), 1-based row / column of the anchor cell.
 */
export function spillReads(f: string): Array<{ sheet: string | null; row: number; col: number }> {
  if (!ANCHOR.test(f)) return [];
  const toks = tokenize(f).filter((t) => t.type !== 'ws');
  const out: Array<{ sheet: string | null; row: number; col: number }> = [];
  for (let i = 0; i + 2 < toks.length; i++) {
    const t = toks[i];
    if (t.type !== 'func' || !/(^|\.)anchorarray$/i.test(t.text) || toks[i + 1].text !== '(') continue;
    const r = toks[i + 2];
    if (r.type !== 'ref' || !r.area || r.prefix?.book || r.prefix?.sheet2 !== undefined) continue;
    out.push({ sheet: r.prefix?.sheet ?? null, row: r.area.r1, col: r.area.c1 });
  }
  return out;
}

/**
 * An array / spill master that breaks (and is converted to values, or becomes #REF!) stops being a spill anchor:
 * every formula reading it with `master#` evaluates to #REF! (Excel: A1# of a plain cell), so those formulas break
 * too — and, being converted themselves, may be masters whose readers break in turn.
 */
function addSpillReaders(doc: MergedDoc, out: BreakList, gone: ReadonlySet<number>): void {
  const key = (sheet: string, col: number, row: number) => `${lc(sheet)}!${col},${row}`;
  const masters = new Set<string>();
  for (const { index, row } of out.cells) {
    const c = doc.columns[index]?.cells[row - 1];
    if (c && (c.arr !== undefined || c.dyn)) masters.add(key(doc.sheetName, index + 1, row));
  }
  for (const i of out.external) {
    const e = doc.externalFormulas[i];
    if (cachedOf(e)?.arr !== undefined) masters.add(key(e.sheet, e.col, e.row));
  }
  if (!masters.size) return;
  // candidates: formulas that read a spill at all (one scan), as [where, anchors read]
  type Cand = { kind: 'cell'; index: number; row: number; reads: string[]; master: string | null } | { kind: 'ext'; i: number; reads: string[]; master: string | null } | { kind: 'name'; n: DefinedName; reads: string[] };
  const cands: Cand[] = [];
  const readKeys = (f: string, host: string) => spillReads(f).map((r) => key(r.sheet ?? host, r.col, r.row));
  const seenCells = new Set(out.cells.map((c) => `${c.index}:${c.row}`));
  const seenExt = new Set(out.external);
  const seenNames = new Set(out.names);
  doc.columns.forEach((col, index) => {
    if (gone.has(index)) return;
    col.cells.forEach((cell, r) => {
      if (!cell || cell.f === undefined || seenCells.has(`${index}:${r + 1}`)) return;
      const reads = readKeys(cell.f, doc.sheetName);
      if (reads.length) cands.push({ kind: 'cell', index, row: r + 1, reads, master: cell.arr !== undefined || cell.dyn ? key(doc.sheetName, index + 1, r + 1) : null });
    });
  });
  doc.externalFormulas.forEach((e, i) => {
    if (seenExt.has(i)) return;
    const reads = readKeys(e.f, e.sheet);
    if (reads.length) cands.push({ kind: 'ext', i, reads, master: cachedOf(e)?.arr !== undefined ? key(e.sheet, e.col, e.row) : null });
  });
  for (const n of doc.names) {
    if (seenNames.has(n) || isInternalName(n)) continue;
    const reads = readKeys(n.refersTo, n.scope ?? doc.sheetName);
    if (reads.length) cands.push({ kind: 'name', n, reads });
  }
  for (let grew = true; grew; ) {
    grew = false;
    for (let k = 0; k < cands.length; k++) {
      const c = cands[k];
      if (!c.reads.some((x) => masters.has(x))) continue;
      cands.splice(k--, 1);
      grew = true;
      if (c.kind === 'cell') out.cells.push({ index: c.index, row: c.row });
      else if (c.kind === 'ext') out.external.push(c.i);
      else out.names.push(c.n);
      if (c.kind !== 'name' && c.master) masters.add(c.master);
    }
  }
}

/** The defined name a 'name' token resolves to (sheet scope of the host / prefix first, then workbook scope). */
function resolveName(names: DefinedName[], text: string, prefixSheet: string | null, host: string | null): DefinedName | undefined {
  const n = text.toLowerCase();
  const scope = (prefixSheet ?? host)?.toLowerCase() ?? null;
  let wb: DefinedName | undefined;
  for (const d of names) {
    if (d.name.toLowerCase() !== n) continue;
    if (d.scope === null) wb = d;
    else if (scope !== null && d.scope.toLowerCase() === scope) return d;
  }
  return wb;
}

function usesName(f: string, host: string | null, names: DefinedName[], broken: Set<DefinedName>, needles: string[]): boolean {
  const low = f.toLowerCase();
  if (!needles.some((s) => low.includes(s))) return false;
  for (const t of tokenize(f)) {
    if (t.type !== 'name' || t.prefix?.book) continue;
    const bare = t.prefix ? t.text.slice(t.prefix.text.length) : t.text;
    const d = resolveName(names, bare, t.prefix ? t.prefix.sheet : null, host);
    if (d && broken.has(d)) return true;
  }
  return false;
}

function addNameUsers(doc: MergedDoc, out: BreakList, gone: ReadonlySet<number>): void {
  const broken = new Set(out.names);
  const needles = [...new Set(out.names.map((n) => n.name.toLowerCase()))];
  const seenCells = new Set(out.cells.map((c) => `${c.index}:${c.row}`));
  doc.columns.forEach((col, index) => {
    if (gone.has(index)) return;
    const cells = col.cells;
    for (let r = 0; r < cells.length; r++) {
      const cell = cells[r];
      if (!cell || cell.f === undefined || seenCells.has(`${index}:${r + 1}`)) continue;
      if (usesName(cell.f, doc.sheetName, doc.names, broken, needles)) out.cells.push({ index, row: r + 1 });
    }
  });
  const seenExt = new Set(out.external);
  doc.externalFormulas.forEach((e, i) => {
    if (seenExt.has(i)) return;
    if (usesName(e.f, e.sheet, doc.names, broken, needles)) out.external.push(i);
  });
}

/** Excel-maintained hidden names (the filter database follows the AutoFilter by itself). */
const isInternalName = (n: DefinedName): boolean => /^_xlnm\._FilterDatabase$/i.test(n.name);

const NAME_LABEL: Record<string, string> = {
  '_xlnm.print_area': '印刷範囲',
  '_xlnm.print_titles': '印刷タイトル',
};
export const nameLabel = (n: DefinedName): string => {
  const base = NAME_LABEL[n.name.toLowerCase()] ?? n.name;
  return n.scope ? `名前「${base}」（${n.scope}）` : `名前「${base}」`;
};
const ruleLabel = (sheet: string, r: FeatureRule): string => `${r.kind === 'dv' ? '入力規則' : '条件付き書式'}（${sheet}!${r.sqref.split(/\s+/).join(',')}）`;

/** UI list of the breaks (examples, UI form) + the total count. */
export function describeBreaks(doc: MergedDoc, b: BreakList, limit = 50): Array<{ where: string; text: string }> {
  const out: Array<{ where: string; text: string }> = [];
  const push = (where: string, text: string) => {
    if (out.length < limit) out.push({ where, text });
  };
  // what cannot be converted to values comes first (it stays #REF! even with "convert")
  const stays = '（値に変換できないため #REF! になります）';
  for (const n of b.names) push(nameLabel(n), toUiForm(n.refersTo) + stays);
  for (const r of b.rules) push(ruleLabel(doc.sheetName, r), r.formulas.map((f) => toUiForm(f)).join(' / ') + stays);
  for (const er of b.externalRules) push(ruleLabel(er.sheet, er.rule), er.rule.formulas.map((f) => toUiForm(f)).join(' / ') + stays);
  for (const i of b.external) {
    const e = doc.externalFormulas[i];
    push(placeText(e.sheet, e.col, e.row), toUiForm(e.f));
  }
  for (const { index, row } of b.cells) push(placeText(doc.sheetName, index + 1, row), toUiForm(doc.columns[index].cells[row - 1]!.f!));
  return out;
}

// ---------------------------------------------------------------------------------------------
// lookup (column-number) issues
// ---------------------------------------------------------------------------------------------
export interface LookupIssue {
  where: string;
  /** the formula (UI form) */
  text: string;
  /** excelRef.lookupWarnings message for this formula */
  message?: string;
}

/**
 * Formulas whose lookup-style functions (VLOOKUP / HLOOKUP / INDEX / MATCH / XMATCH / XLOOKUP / OFFSET / CHOOSE)
 * would silently point at another column after the ops: excelRef.lookupWarnings (exact when the index is a literal
 * number; ranges, tables and defined names), op by op. `skip(index)` excludes merged-sheet columns (e.g. the ones
 * being deleted). Filled-down copies share one check: the result does not depend on row numbers, so formulas with
 * the same literal text and the same column spans (plus "single row" flags) are checked once.
 */
/** columns (cells arrays) known to hold no lookup-style formula */
const lookupFree = new WeakMap<object, boolean>();

export function lookupIssues(doc: MergedDoc, ops: readonly StructOp[], skip: (index: number) => boolean = () => false): { count: number; examples: LookupIssue[] } {
  if (!ops.length) return { count: 0, examples: [] };
  const sheet = doc.sheetName;
  const target = lc(sheet);
  const tablesAt = formulaCtxFor(doc, ops).tableRefsAt;
  const resolveName = (name: string): string | null => {
    const n = name.toLowerCase();
    const hit = doc.names.find((x) => x.scope === null && x.name.toLowerCase() === n) ?? doc.names.find((x) => x.name.toLowerCase() === n);
    return hit ? hit.refersTo : null;
  };
  const exact = (f: string, host: string): LookupWarning[] => {
    let cur = f;
    for (let i = 0; i < ops.length; i++) {
      const actx = { hostSheet: host, targetSheet: sheet, tables: tablesAt[i] };
      let w: LookupWarning[];
      try {
        w = lookupWarnings(cur, ops[i], actx, { resolveName });
      } catch {
        w = [];
      }
      if (w.length) return w;
      if (i + 1 < ops.length) cur = adjustFormula(cur, ops[i], actx).formula;
    }
    return [];
  };
  // filled-down copies share their `parts` array (compileWithHint): key = parts identity + column spans of the refs
  const cache = new Map<string[], Map<string, boolean>>();
  const sigOf = (c: Compiled, hostLc: string): string => {
    let k = hostLc;
    for (const r of c.refs) k += '|' + r.prefix + r.area.kind + r.area.c1 + ',' + r.area.c2 + (r.area.r1 === r.area.r2 ? '=' : '');
    return k;
  };
  let count = 0;
  const examples: LookupIssue[] = [];
  const consider = (host: string, hostLc: string, f: string, c: Compiled, where: () => string) => {
    if (!c.lookup) return;
    let byParts = cache.get(c.parts);
    if (!byParts) cache.set(c.parts, (byParts = new Map()));
    const k = sigOf(c, hostLc);
    let hit = byParts.get(k);
    if (hit === undefined) {
      hit = exact(f, host).length > 0;
      byParts.set(k, hit);
    }
    if (!hit) return;
    count++;
    if (examples.length < 10) examples.push({ where: where(), text: toUiForm(f), message: exact(f, host)[0]?.message });
  };
  doc.columns.forEach((col, index) => {
    if (skip(index)) return;
    const cells = col.cells;
    if (lookupFree.get(cells)) return; // no lookup-style formula in this (immutable) column
    const cc = compiledCol(cells);
    let any = false;
    for (let r = 0; r < cells.length; r++) {
      const cell = cells[r];
      if (!cell || cell.f === undefined) continue;
      const c = compiledAt(cc, r, cell.f);
      if (!c.lookup) continue;
      any = true;
      consider(sheet, target, cell.f, c, () => placeText(sheet, index + 1, r + 1));
    }
    if (!any) lookupFree.set(cells, true);
  });
  const ec = compiledCol(doc.externalFormulas);
  doc.externalFormulas.forEach((e, i) => consider(e.sheet, lc(e.sheet), e.f, compiledAt(ec, i, e.f), () => placeText(e.sheet, e.col, e.row)));
  return { count, examples };
}

/** Warning sentences for OpResult.warnings / DeleteImpact.lookupWarnings. */
export function lookupWarningTexts(issues: { count: number; examples: LookupIssue[] }, verb: '移動' | '削除' | '挿入', max = 5): string[] {
  if (!issues.count) return [];
  const out = [
    `列番号・位置で列を指定する関数（VLOOKUP・INDEX・MATCH など）を使った数式が ${issues.count.toLocaleString()} 個あります。列の${verb}で、これらの数式が別の列の値を返す可能性があります（Excel でも自動では直りません）。`,
  ];
  const first = issues.examples.find((e) => e.message);
  if (first?.message) out.push(first.message);
  for (const ex of issues.examples.slice(0, max)) out.push(`${ex.where}: ${ex.text}`);
  if (issues.count > max) out.push(`ほか ${(issues.count - max).toLocaleString()} 個`);
  return out;
}

export type { FormulaCtx };
