// Pipeline harness for the foreign-producer corpus (round 5, module H1): one workbook that Excel did not write goes through
// readBook -> pickDataSheet / detectRegion -> analyzeSources with a second file -> joinSources -> buildPremerged -> readBook ->
// createEditor -> a random op log (moves incl. non-contiguous, inserts, deletes with convert, edits, widths) -> replayLog, and
// the invariants of docs/ARCHITECTURE.md are checked:
//   - no exception except a clean refusal (PatchError / ReadBookError with a Japanese message, not `internal`);
//   - the premerged sheet keeps the base sheet's cells (display text + resolved style) and holds the imported cells of the source
//     (display text; 1904 <-> 1900 book pairs included);
//   - doc == readBook(replayLog(premerged, log)) for the merged sheet and every other sheet's formulas (tests/doc/compare.ts);
//   - every output package is re-readable.
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { analyzeSources, defaultPlan, detectRegion, joinSources, pickDataSheet } from '../../../src/core/keys';
import { buildPremerged, cellText, readBook, replayLog } from '../../../src/core/xlsx';
import { createEditor } from '../../../src/core/doc';
import { PatchError } from '../../../src/core/xlsx/patch/pkg';
import { ReadBookError } from '../../../src/core/xlsx/zip';
import type { KeySource, PremergeInput, PremergeLayout } from '../../../src/core/contracts';
import type { Cell, EditorSnapshot, JoinType, LoggedOp, MergedDoc, SourceBook } from '../../../src/core/types';
import { styleKey } from '../../../src/core/doc/insertStyle';
import { compare } from '../../doc/compare';
import { parseAreaText, numToCol } from '../../../src/core/formula/excelRef';
import { Package, loadZip } from '../../../src/core/xlsx/zip';
import { rng, step, TEXTS } from '../../doc/fuzzcore';
import { formatProblems, lintPackage, newProblems, type LintProblem } from './ooxml-lint';

const here = path.dirname(fileURLToPath(import.meta.url));
export const CORPUS_DIR = path.join(here, 'corpus');
export const PARTNER = 'partner.xlsx';

const JP = /[぀-ヿ㐀-鿿＀-￯]/;
export const isJapanese = (s: string) => JP.test(s);

export interface Loaded {
  name: string;
  bytes: Uint8Array;
  book: SourceBook;
  ks: KeySource;
}

export async function load(dir: string, name: string, id: string): Promise<Loaded> {
  return loadBytes(name, new Uint8Array(fs.readFileSync(path.join(dir, name))), id);
}

export async function loadBytes(name: string, bytes: Uint8Array, id: string): Promise<Loaded> {
  const book = await readBook(bytes, name, { id });
  const sheet = pickDataSheet(book);
  const region = detectRegion(book, sheet);
  return { name, bytes, book, ks: { id, book, region, label: name.replace(/\.[^.]+$/, '') } };
}

const lintCache = new WeakMap<Uint8Array, Promise<LintProblem[]>>();
/** structural lint of a package (tests/fixtures/foreign/ooxml-lint.ts), cached per byte array */
export function lintOf(bytes: Uint8Array): Promise<LintProblem[]> {
  let p = lintCache.get(bytes);
  if (!p) lintCache.set(bytes, (p = lintPackage(bytes)));
  return p;
}

export type Verdict = 'ok' | 'refused' | 'bug';

export interface RunResult {
  file: string;
  role: 'base' | 'other';
  join: JoinType;
  seed: number;
  /** what happened: ok / a clean refusal (message) / a bug (problems) */
  verdict: Verdict;
  stage: string;
  refusal?: string;
  problems: string[];
  info: { region?: string; rows?: number; imports?: number; appended?: number; ops?: number; trail?: string[] };
  /** the op log of the run (for replaying it on another package) */
  log?: LoggedOp[];
  premerged?: Uint8Array;
  output?: Uint8Array;
}

/** a thrown error classified: a clean refusal carries a Japanese message and is not `internal` */
export function classify(e: unknown): { clean: boolean; message: string } {
  const message = e instanceof Error ? e.message : String(e);
  if (e instanceof PatchError) return { clean: !e.internal && isJapanese(message), message };
  if (e instanceof ReadBookError) return { clean: isJapanese(message), message };
  return { clean: false, message: `${(e as Error)?.name ?? 'Error'}: ${message}${e instanceof Error && e.stack ? '\n' + e.stack.split('\n').slice(1, 4).join('\n') : ''}` };
}

/** text / number (dates are numbers) / boolean / error: a text '85' that becomes the number 85 shows the same but is another cell */
const typeClass = (c: Cell | undefined): string => (!c || c.v === null || c.v === '' ? 'empty' : c.t === 'd' ? 'n' : c.t);

/** display text + style of every cell of the base data sheet must be unchanged in the premerged sheet (columns left of the insert) */
function basePreserved(orig: SourceBook, region: KeySource['region'], merged: SourceBook, layout: PremergeLayout, lastRowShift: number): string[] {
  const problems: string[] = [];
  const a = orig.sheets.find((s) => s.name === region.sheet);
  const b = merged.sheets.find((s) => s.name === layout.sheet);
  if (!a || !b) return [`merged sheet ${layout.sheet} missing`];
  const maxC = Math.min(region.lastCol, a.maxCol);
  let n = 0;
  for (let r = 1; r <= a.maxRow; r++) {
    // rows below the appended block moved down by the number of appended rows
    const r2 = layout.appendAt !== null && r >= layout.appendAt ? r + lastRowShift : r;
    for (let c = 1; c <= maxC; c++) {
      const ca = a.rows[r - 1]?.[c - 1];
      const cb = b.rows[r2 - 1]?.[c - 1];
      // formulas may be rewritten (range extension over appended rows), their cached results must stay: the premerge calculates nothing
      if (ca?.f !== undefined || cb?.f !== undefined) {
        if ((ca?.f === undefined) !== (cb?.f === undefined)) problems.push(`base ${c},${r}: formula ${ca?.f} -> ${cb?.f}`);
        else if (cellText(orig, ca) !== cellText(merged, cb)) problems.push(`base ${c},${r}: cached "${cellText(orig, ca)}" -> "${cellText(merged, cb)}" (${ca?.f})`);
        continue;
      }
      const ta = cellText(orig, ca);
      const tb = cellText(merged, cb);
      if (ta !== tb && !(orig.date1904 !== merged.date1904)) problems.push(`base ${c},${r}: "${ta}" -> "${tb}"`);
      else if (ta !== tb) problems.push(`base ${c},${r}: "${ta}" -> "${tb}" (date system differs)`);
      if (ta !== '' && typeClass(ca) !== typeClass(cb)) problems.push(`base type ${c},${r}: ${typeClass(ca)} -> ${typeClass(cb)} ("${ta}")`);
      if (ca && cb && ta !== '' && styleKey(orig.styles[ca.s ?? 0]) !== styleKey(merged.styles[cb.s ?? 0])) problems.push(`base style ${c},${r}: ${styleKey(orig.styles[ca.s ?? 0])} -> ${styleKey(merged.styles[cb.s ?? 0])}`);
      if (++n > 200000 || problems.length >= 12) return problems;
    }
  }
  return problems;
}

/** matched source cells arrive with the same display text (constants; a formula whose cached value is kept may be a live formula) */
function importsArrived(input: PremergeInput, merged: SourceBook, layout: PremergeLayout, baseHeaderRow: number): string[] {
  const problems: string[] = [];
  const m = merged.sheets.find((s) => s.name === layout.sheet);
  if (!m) return problems;
  const sameStyle = (sb: SourceBook, a: Cell | undefined, mb: SourceBook, b: Cell | undefined) => styleKey(sb.styles[a?.s ?? 0]) === styleKey(mb.styles[b?.s ?? 0]);
  const srcSheet = (id: string) => {
    const s = input.sources.find((x) => x.id === id)!;
    return { ks: s, grid: s.book.sheets.find((g) => g.name === s.region.sheet)! };
  };
  for (const im of layout.imported) {
    const { ks, grid } = srcSheet(im.sourceId);
    if (!grid) continue;
    const styled = ks.book.fidelity === 'full';
    if (styled) {
      // the header cell keeps its look, the column its width
      const hs = grid.rows[ks.region.headerRow - 1]?.[im.sourceCol - 1];
      const hm = m.rows[baseHeaderRow - 1]?.[im.col - 1];
      if (hs && hm && cellText(ks.book, hs) !== '' && !sameStyle(ks.book, hs, merged, hm)) problems.push(`import ${im.sourceId} col ${im.sourceCol}: header style ${styleKey(ks.book.styles[hs.s ?? 0])} -> ${styleKey(merged.styles[hm.s ?? 0])}`);
      const ws = grid.colProps[im.sourceCol - 1]?.width;
      const wm = m.colProps[im.col - 1]?.width;
      if (ws !== undefined && ws > 0 && !grid.colProps[im.sourceCol - 1]?.hidden && (wm === undefined || Math.abs(wm - ws) > 1e-6)) problems.push(`import ${im.sourceId} col ${im.sourceCol}: width ${ws} -> ${wm}`);
    }
    let checked = 0;
    let ai = 0;
    input.rows.forEach((plan) => {
      const appended = plan.baseRow == null ? layout.appended[ai++] : undefined;
      if (checked > 40 || problems.length >= 12) return;
      const target = plan.baseRow ?? appended?.row;
      const match = plan.matches[im.sourceId];
      const sr = match ? match.row : plan.appendedFrom?.sourceId === im.sourceId ? plan.appendedFrom.row : null;
      if (sr == null || target == null) return;
      const sc = grid.rows[sr - 1]?.[im.sourceCol - 1];
      const mc = m.rows[target - 1]?.[im.col - 1];
      checked++;
      const ta = cellText(ks.book, sc);
      const tb = cellText(merged, mc);
      if (ta !== tb) problems.push(`import ${im.sourceId} col ${im.sourceCol} -> ${im.col} row ${sr}->${target}: "${ta}" -> "${tb}"`);
      else if (ta !== '' && typeClass(sc) !== typeClass(mc)) problems.push(`import ${im.sourceId} col ${im.sourceCol} -> ${im.col} row ${sr}->${target}: type ${typeClass(sc)} -> ${typeClass(mc)} ("${ta}")`);
      else if (styled && ta !== '' && mc && !sameStyle(ks.book, sc, merged, mc)) problems.push(`import ${im.sourceId} col ${im.sourceCol} -> ${im.col} row ${sr}->${target}: style ${styleKey(ks.book.styles[sc?.s ?? 0])} -> ${styleKey(merged.styles[mc.s ?? 0])}`);
    });
  }
  return problems;
}

const xmlText = (s: string) => s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');

/** formulas of every chart part of a package, in part / document order (prefix agnostic: <c:f>, <f> of a default-namespace chart, c15:f ...) */
async function chartFormulas(bytes: Uint8Array): Promise<string[]> {
  const pkg = new Package(await loadZip(bytes));
  const out: string[] = [];
  for (const name of pkg.partNames.filter((n) => /(^|\/)charts\/chart[^/]*\.xml$/i.test(n)).sort()) {
    const x = (await pkg.text(name)) ?? '';
    for (const m of x.matchAll(/<(?:[\w.-]+:)?f(?:\s[^>]*)?>([^<]*)<\/(?:[\w.-]+:)?f>/g)) out.push(xmlText(m[1]));
  }
  return out;
}

/** `'Sheet'!$E$2:$E$31` of one sheet -> first column, or null (other sheet, several areas, errors) */
function firstColumnOf(formula: string, sheet: string): number | null {
  const m = /^=?(?:'((?:[^']|'')+)'|([^'!:,()\s]+))!(\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)$/.exec(formula.trim());
  if (!m) return null;
  const name = (m[1] ?? m[2]).replace(/''/g, "'");
  if (name.toLowerCase() !== sheet.toLowerCase()) return null;
  return parseAreaText(m[3].replace(/\$/g, ''))?.c1 ?? null;
}

/**
 * Charts follow the columns they plot: every chart formula that referred to a column of the base sheet refers to the column the doc now
 * has for it (its id b<n>; a deleted column is skipped). Comments (legacy notes) stay where their VML shape is: the comment refs of the
 * output equal the Row / Column positions of its VML note shapes.
 */
export async function packageChecks(origBytes: Uint8Array, outBytes: Uint8Array, sheet: string, colIds: string[]): Promise<string[]> {
  const problems: string[] = [];
  const before = await chartFormulas(origBytes);
  const after = await chartFormulas(outBytes);
  if (before.length !== after.length) problems.push(`chart formulas: ${before.length} before, ${after.length} after`);
  else before.forEach((f, i) => {
    const c1 = firstColumnOf(f, sheet);
    if (c1 === null) return;
    if (/#REF!/.test(after[i])) return;
    // the doc knows where every original base column is now (column id 'b<n>'); a deleted column has no id any more
    const want = colIds.indexOf('b' + c1) + 1;
    if (want === 0) return;
    const c2 = firstColumnOf(after[i], sheet);
    if (c2 !== want) problems.push(`chart ${f} -> ${after[i]}: the column is now ${numToCol(want)}`);
  });
  // legacy notes: comment refs vs VML positions
  const pkg = new Package(await loadZip(outBytes));
  for (const name of pkg.partNames.filter((n) => /(^|\/)comments?\d*\.xml$/i.test(n))) {
    const refs = [...((await pkg.text(name)) ?? '').matchAll(/<comment\b[^>]*\sref="([A-Z]+\d+)"/g)].map((m) => m[1]).sort();
    const vml = new Set<string>();
    for (const r of pkg.partNames.filter((n) => /\.vml$/i.test(n))) {
      const x = (await pkg.text(r)) ?? '';
      if (!x.includes(name.replace(/^.*\//, '')) && !/commentsDrawing|legacy|vmlDrawing/i.test(r) && !/ObjectType="Note"/.test(x)) continue;
      for (const s of x.matchAll(/<(?:[\w.-]+:)?ClientData\b[^>]*ObjectType="Note"[\s\S]*?<\/(?:[\w.-]+:)?ClientData>/g)) {
        const row = /<(?:[\w.-]+:)?Row>\s*(\d+)/.exec(s[0])?.[1];
        const col = /<(?:[\w.-]+:)?Column>\s*(\d+)/.exec(s[0])?.[1];
        if (row !== undefined && col !== undefined) vml.add(`${numToCol(Number(col) + 1)}${Number(row) + 1}`);
      }
    }
    const shapes = [...vml].sort();
    if (vml.size && refs.join(' ') !== shapes.join(' ')) problems.push(`comments ${name}: refs [${refs.join(' ')}] but VML notes at [${shapes.join(' ')}]`);
  }
  return problems;
}

/**
 * Known doc-module gap (reported, not a patcher issue): a plain number typed into a DATE-formatted cell of a 1904 book. The doc keeps
 * the number as is (a 1900-system serial, "1900/1/12"), the file keeps it raw too and Excel / the reader show it as a 1904-system
 * date (raw 12 -> "1904/1/13" = 1474 in 1900 terms), i.e. the file differs from the doc by exactly 1462 days.
 */
function isTypedNumberIn1904DateCell(diff: string): boolean {
  const m = /value doc n:(-?[\d.]+) file n:(-?[\d.]+)$/.exec(diff);
  return !!m && Math.abs(Number(m[2]) - Number(m[1]) - 1462) < 1e-9;
}

export interface RunOptions {
  role: 'base' | 'other';
  join: JoinType;
  seed: number;
  steps: number;
  /** keep bytes in the result (for writing the files out) */
  keep?: boolean;
  texts?: string[];
}

/** The default plan of `base` + `other` (ids s1 / s2) through buildPremerged: the premerged package, its layout, the report and the input. */
export async function premergeFiles(base: Loaded, other: Loaded, join: JoinType = 'left') {
  const baseKs: KeySource = { ...base.ks, id: 's1' };
  const otherKs: KeySource = { ...other.ks, id: 's2' };
  const analysis = analyzeSources(baseKs, [otherKs]);
  const plan = joinSources(analysis, { ...defaultPlan(analysis), joinType: join });
  const input: PremergeInput = {
    base: baseKs, sources: [otherKs], rows: plan.rows, imports: plan.defaultImports, keyFill: plan.keyFill,
    options: { formulaPolicy: 'keep-if-possible', extendTable: true, headerConflict: 'suffix' },
  };
  const pm = await buildPremerged(input);
  return { ...pm, input, baseKs, otherKs, book: await readBook(pm.bytes, base.name, { id: 'merged' }) };
}

export async function runPipeline(file: Loaded, partner: Loaded, o: RunOptions): Promise<RunResult> {
  const res: RunResult = { file: file.name, role: o.role, join: o.join, seed: o.seed, verdict: 'ok', stage: 'start', problems: [], info: {} };
  const fail = (stage: string, e: unknown) => {
    const c = classify(e);
    res.stage = stage;
    if (c.clean) {
      res.verdict = 'refused';
      res.refusal = c.message;
    } else {
      res.verdict = 'bug';
      res.problems.push(`${stage}: ${c.message}`);
    }
  };
  const base = o.role === 'base' ? file : partner;
  const other = o.role === 'base' ? partner : file;
  const baseKs: KeySource = { ...base.ks, id: 's1' };
  const otherKs: KeySource = { ...other.ks, id: 's2' };
  res.info.region = `${file.ks.region.sheet}!hdr${file.ks.region.headerRow} data${file.ks.region.firstDataRow}-${file.ks.region.lastDataRow} cols${file.ks.region.firstCol}-${file.ks.region.lastCol}`;
  let premerged: { bytes: Uint8Array; layout: PremergeLayout; input: PremergeInput } | undefined;
  try {
    res.stage = 'keys';
    const analysis = analyzeSources(baseKs, [otherKs]);
    const plan = { ...defaultPlan(analysis), joinType: o.join };
    const join = joinSources(analysis, plan);
    const input: PremergeInput = {
      base: baseKs,
      sources: [otherKs],
      rows: join.rows,
      imports: join.defaultImports,
      keyFill: join.keyFill,
      options: { formulaPolicy: 'keep-if-possible', extendTable: true, headerConflict: 'suffix' },
    };
    res.info.rows = join.rows.length;
    res.info.imports = join.defaultImports.length;
    res.stage = 'premerge';
    const pm = await buildPremerged(input);
    res.info.appended = pm.report.appendedRows;
    premerged = { bytes: pm.bytes, layout: pm.layout, input };
    if (o.keep) res.premerged = pm.bytes;
  } catch (e) {
    fail(res.stage, e);
    return res;
  }
  const { bytes, layout, input } = premerged;
  let mergedBook: SourceBook;
  try {
    res.stage = 'read premerged';
    mergedBook = await readBook(bytes, base.name, { id: 'merged' });
  } catch (e) {
    res.verdict = 'bug';
    res.problems.push(`premerged output is not readable: ${classify(e).message}`);
    return res;
  }
  // what Excel would repair: new structural problems that the base package did not have
  const baseLint = await lintOf(base.bytes);
  const preLint = newProblems(baseLint, await lintOf(bytes));
  if (preLint.length) res.problems.push(`lint(premerged): ${formatProblems(preLint)}`);
  res.problems.push(...basePreserved(base.book, baseKs.region, mergedBook, layout, layout.appended.length));
  res.problems.push(...importsArrived(input, mergedBook, layout, baseKs.region.headerRow));
  // editor + random ops + replay
  try {
    res.stage = 'editor';
    const colors = ['#2563eb', '#15803d'];
    const sourceInfo: MergedDoc['sourceInfo'] = { s1: { label: baseKs.label, color: colors[0] }, s2: { label: otherKs.label, color: colors[1] } };
    let last = baseKs.region.lastDataRow;
    for (const a of layout.appended) if (a.row > last) last = a.row;
    const region = { headerRow: baseKs.region.headerRow, firstDataRow: baseKs.region.firstDataRow, lastDataRow: last };
    let s: EditorSnapshot = createEditor(mergedBook, layout, sourceInfo, region);
    const r = rng(o.seed * 7919 + file.name.length * 31 + (o.role === 'base' ? 1 : 2));
    const trail: string[] = [];
    res.stage = 'ops';
    for (let i = 0; i < o.steps; i++) {
      const { res: rr, what } = step(s, r, false, s.doc.lastDataRow + 2, o.texts ?? TEXTS);
      trail.push(rr.ok ? what : `${what} [refused: ${rr.reason}]`);
      if (rr.ok) s = rr.snapshot;
      else if (!isJapanese(rr.reason)) res.problems.push(`refusal without a Japanese message: ${rr.reason}`);
    }
    res.info.trail = trail;
    res.info.ops = s.log.length;
    res.log = s.log;
    res.stage = 'replay';
    const out = await replayLog(bytes, s.log);
    if (o.keep) res.output = out;
    res.stage = 'read output';
    const book2 = await readBook(out, base.name, { id: 'out' });
    res.stage = 'compare';
    const diffs = compare(s.doc, book2).filter((d) => !(book2.date1904 && isTypedNumberIn1904DateCell(d)));
    if (diffs.length) res.problems.push(...diffs.slice(0, 8).map((d) => `doc != file: ${d}`), ...(diffs.length > 8 ? [`... ${diffs.length} differences`] : []));
    res.stage = 'lint';
    const outLint = newProblems([...baseLint, ...(await lintOf(bytes))], await lintOf(out));
    if (outLint.length) res.problems.push(`lint(output): ${formatProblems(outLint)}`);
    res.stage = 'package checks';
    res.problems.push(...(await packageChecks(base.bytes, out, baseKs.region.sheet, s.doc.columns.map((c) => c.id))));
  } catch (e) {
    fail(res.stage, e);
  }
  if (res.problems.length && res.verdict === 'ok') res.verdict = 'bug';
  return res;
}

// ---------------------------------------------------------------------------------------------------------------------------------
// Two workbooks compared the way a user sees them (shared by the mutation and the no-op tests)
/**
 * Formula text up to the spelling of sheet names: 'Name'! and Name! are the same reference, and a reference qualified with the sheet the
 * formula lives on is the same as an unqualified one (the patcher keeps the text as written, only references it moves are re-rendered).
 */
const PLAIN = /^[^\s'!&<>"()+*/,:;=-]+$/;
const unquote = (f: string) => f.replace(/'([^'!]+)'!/g, (m0, n: string) => (PLAIN.test(n) ? `${n}!` : m0));
const normF = (f: string | undefined, host: string | null) => {
  if (f === undefined) return '';
  const u = unquote(f);
  return host && PLAIN.test(host) ? u.split(`${host}!`).join('') : u;
};

/** what makes two cells "the same" for the user: typed value, formula, array flag, resolved style */
function cellKey(b: SourceBook, c: Cell | undefined, host: string): string {
  if (!c || (c.v === null && c.f === undefined && !c.s)) return '';
  const v = c.v === null || c.v === '' ? '' : `${c.t === 'd' ? 'n' : c.t}:${String(c.v)}`;
  return `${v}|${normF(c.f, host)}|${c.arr ?? ''}|${c.dyn ? 'dyn' : ''}|${styleKey(b.styles[c.s ?? 0])}`;
}

export function bookDiff(a: SourceBook, b: SourceBook): string[] {
  const out: string[] = [];
  if (a.sheets.map((s) => s.name).join('|') !== b.sheets.map((s) => s.name).join('|')) out.push(`sheets ${a.sheets.map((s) => s.name)} vs ${b.sheets.map((s) => s.name)}`);
  for (const sa of a.sheets) {
    const sb = b.sheets.find((s) => s.name === sa.name);
    if (!sb) continue;
    const rows = Math.max(sa.maxRow, sb.maxRow);
    const cols = Math.max(sa.maxCol, sb.maxCol);
    for (let r = 1; r <= rows && out.length < 12; r++) {
      for (let c = 1; c <= cols; c++) {
        const ka = cellKey(a, sa.rows[r - 1]?.[c - 1], sa.name);
        const kb = cellKey(b, sb.rows[r - 1]?.[c - 1], sb.name);
        if (ka !== kb) { out.push(`${sa.name}!${c},${r}: ${ka} vs ${kb}`); break; }
      }
    }
    const cp = (g: typeof sa, c: number) => `${g.colProps[c]?.width ?? ''}/${g.colProps[c]?.hidden ? 'h' : ''}/${g.colProps[c]?.customWidth ? 'c' : ''}/${g.colProps[c]?.outlineLevel ?? ''}`;
    for (let c = 0; c < cols; c++) if (cp(sa, c) !== cp(sb, c)) { out.push(`${sa.name} column ${c + 1}: ${cp(sa, c)} vs ${cp(sb, c)}`); break; }
    for (let r = 0; r < rows; r++) {
      const pa = JSON.stringify(sa.rowProps[r] ?? {});
      const pb = JSON.stringify(sb.rowProps[r] ?? {});
      if (pa !== pb) { out.push(`${sa.name} row ${r + 1}: ${pa} vs ${pb}`); break; }
    }
    const fa = sa.features;
    const fb = sb.features;
    const box = (l: Array<{ r1: number; c1: number; r2: number; c2: number }>) => l.map((x) => `${x.r1},${x.c1},${x.r2},${x.c2}`).sort().join(' ');
    if (box(fa.merges) !== box(fb.merges)) out.push(`${sa.name} merges`);
    if (fa.autoFilter !== fb.autoFilter) out.push(`${sa.name} autoFilter ${fa.autoFilter} vs ${fb.autoFilter}`);
    if (JSON.stringify(fa.freeze ?? null) !== JSON.stringify(fb.freeze ?? null)) out.push(`${sa.name} freeze`);
    if (JSON.stringify(fa.tables) !== JSON.stringify(fb.tables)) out.push(`${sa.name} tables ${JSON.stringify(fa.tables)} vs ${JSON.stringify(fb.tables)}`);
    const rule = (x: (typeof fa.rules)[number]) => `${x.kind}|${x.ext ? 'x14' : ''}|${x.type}|${x.sqref}|${x.formulas.map((f) => normF(f, sa.name)).join(';')}`;
    if (fa.rules.map(rule).sort().join('\n') !== fb.rules.map(rule).sort().join('\n')) out.push(`${sa.name} rules`);
    if (JSON.stringify(fa.hyperlinks) !== JSON.stringify(fb.hyperlinks)) out.push(`${sa.name} hyperlinks`);
  }
  const nm = (x: SourceBook) => x.names.map((n) => `${n.scope ?? ''}!${n.name}=${unquote(n.refersTo)}`).sort().join('\n');
  if (nm(a) !== nm(b)) out.push(`names ${nm(a)} vs ${nm(b)}`);
  if (a.date1904 !== b.date1904) out.push('date1904');
  return out;
}

