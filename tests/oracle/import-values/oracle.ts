/**
 * Real-Excel value-invariance oracle for imported formulas (task F1, step 3). Re-runnable:
 *
 *   powershell -NoProfile -ExecutionPolicy Bypass -File tests\oracle\import-values\run.ps1 [-Corpus <dir>] [-Only <regex>] [-Configs all|quick]
 *
 * For every corpus file (tests/oracle/import-values/gen_corpus.py + author.ps1) and a few configurations (default imports / all columns,
 * join left / full, with and without stored results) the app's pipeline (keys module -> buildPremerged) builds the merged package; ONE hidden
 * Excel (run.ps1 -> oracle-excel.ps1) opens every package (a repair prompt is a failure), recalculates it fully and dumps the merged sheet.
 * Every formula that stayed LIVE is then compared with the value the same formula has in the ORIGINAL source file recalculated by Excel
 * (the Excel-saved twin in <corpus>\cached) for the same person:
 *   - ranges widened to the merged table (rule b) are compared with the source after Excel deleted the rows that were not placed
 *     (Excel shrinks the ranges exactly like the widening does) — functions whose result is a POSITION (MATCH on its own, ROWS …) only for
 *     the unchanged order;
 *   - no new #REF! / #NAME? / #VALUE! / #DIV/0! … may appear.
 * Results go to <out>\results.json and <out>\results.md.
 *
 * Round 9 (includeSourceSheets, task F2): `--sources` builds every package with PremergeOptions.includeSourceSheets (the original sheets of the
 * source are copied; the formulas the merged table cannot keep read the copies) and compares exactly the same way — a formula read through a
 * copy must have the very value it has in the source. `--random <n> [--seed <s>]` replaces the corpus by n random source workbooks (random.ts:
 * 3 row orders each, rows missing on both sides, join left / full, imported and non-imported columns) whose values are computed by Excel
 * itself (no cached twin): the property test "the remapped formula evaluated by Excel equals the source value".
 */
import { execFileSync } from 'node:child_process';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { tokenize } from '../../../src/core/formula/excelRef';
import type { SourceBook } from '../../../src/core/types';
import { readBook, replayLog } from '../../../src/core/xlsx';
import { buildPremerged, type FormulaTrace } from '../../../src/core/xlsx/premerge';
import { buildPipeline, loadKeySource, loadManifest, placedRows, sheetOf, toPremergeInput, type CorpusEntry } from './lib';
import { writeRandomCorpus } from './random';

interface Config { name: string; results: 'cached' | 'nocache'; imports: 'default' | 'all' | number[]; join: 'left' | 'full'; /** PremergeOptions.dropOtherSheets (the app's default) */ drop?: boolean; /** PremergeOptions.includeSourceSheets */ sources?: boolean }
const CONFIGS: Config[] = [
  { name: 'cached-default-left', results: 'cached', imports: 'default', join: 'left' },
  { name: 'cached-all-left', results: 'cached', imports: 'all', join: 'left' },
  { name: 'nocache-default-left', results: 'nocache', imports: 'default', join: 'left' },
  { name: 'nocache-default-full', results: 'nocache', imports: 'default', join: 'full' },
  { name: 'cached-all-full', results: 'cached', imports: 'all', join: 'full' },
];
/** the app's own sample workbook as the base (merged title, charts, pivot, names …): with and without single-sheet output */
const SAMPLE_CONFIGS: Config[] = [
  { name: 'cached-default-left', results: 'cached', imports: 'default', join: 'left' },
  { name: 'cached-default-left-onesheet', results: 'cached', imports: 'default', join: 'left', drop: true },
  { name: 'nocache-default-full', results: 'nocache', imports: 'default', join: 'full' },
  { name: 'nocache-default-full-onesheet', results: 'nocache', imports: 'default', join: 'full', drop: true },
  { name: 'cached-all-left-onesheet', results: 'cached', imports: 'all', join: 'left', drop: true },
];

/**
 * The property test: every configuration imports the formula columns (G-N) of the random source but a different part of the input columns,
 * so that formulas read columns that are NOT imported (C-E, the keys) — that is what the copies are for.
 */
const FORMULA_COLS = [7, 8, 9, 10, 11, 12, 13, 14];
const RANDOM_CONFIGS: Config[] = [
  { name: 'rnd-default-left', results: 'nocache', imports: 'default', join: 'left' },
  { name: 'rnd-default-full', results: 'nocache', imports: 'default', join: 'full' },
  { name: 'rnd-c-full', results: 'nocache', imports: [3, ...FORMULA_COLS], join: 'full' },
  { name: 'rnd-de-left', results: 'nocache', imports: [4, 5, ...FORMULA_COLS], join: 'left' },
  { name: 'rnd-alt-left', results: 'nocache', imports: [3, 4, 5, 8, 10, 12, 14], join: 'left' },
];

interface Job {
  id: string;
  /** the Excel dump (<out>\<srcDump>.values.json) of the source workbook: the expected values when there is no cached twin */
  srcDump?: string;
  entry: CorpusEntry;
  cfg: Config;
  kept: FormulaTrace[];
  counts: { kept: number; value: number; empty: number; viaCopy: number };
  mergedSheet: string;
  /** reduced source (widened cells): the cached twin with the unplaced data rows deleted */
  reduced: { file: string; deleteRows: number[] } | null;
  /** placed source row -> merged row */
  placed: Array<[number, number]>;
  /** the Excel-saved twin of the source (null: the random corpus, see srcDump) */
  twin: SourceBook | null;
  order: string;
  /** last merged column (the base's last column is moved behind it by the move check) */
  lastCol: number;
  /** the column the move check moved (the base region's last column) */
  moveFrom: number;
}

type XVal = null | boolean | number | string | { e: number };
interface Dump { open: 'normal' | 'repair' | 'failed'; log?: string; rows: XVal[][] }

const ERR: Record<number, string> = {
  [-2146826281]: '#DIV/0!', [-2146826246]: '#N/A', [-2146826259]: '#NAME?', [-2146826288]: '#NULL!', [-2146826252]: '#NUM!', [-2146826265]: '#REF!', [-2146826273]: '#VALUE!', [-2146826233]: '#GETTING_DATA', [-2146826238]: '#CALC!', [-2146826239]: '#SPILL!', [-2146826240]: '#FIELD!',
};
const POSITION_ONLY = /\bMATCH\(/;

function argv(args: string[], k: string, d = ''): string {
  const i = args.indexOf(`--${k}`);
  return i >= 0 && args[i + 1] ? args[i + 1] : d;
}

export async function main(args: string[]): Promise<void> {
  const out = argv(args, 'out');
  let corpus = argv(args, 'corpus');
  const only = argv(args, 'only') ? new RegExp(argv(args, 'only')) : null;
  const quick = argv(args, 'configs', 'all') === 'quick';
  const noExcel = args.includes('--no-excel');
  /** the random corpus (random.ts) instead of the corpus directory: the property test of the copies */
  const random = Number(argv(args, 'random', '0')) || 0;
  /** every package is built with PremergeOptions.includeSourceSheets (the random corpus always is) */
  const sources = args.includes('--sources') || random > 0;
  const seed = Number(argv(args, 'seed', '1')) || 1;
  fs.mkdirSync(out, { recursive: true });
  if (random) {
    corpus = path.join(out, 'random-corpus');
    fs.mkdirSync(corpus, { recursive: true });
    await writeRandomCorpus(corpus, random, seed);
  }
  /** compare against the dumps of an earlier run (same corpus / packages): no Excel */
  const reuse = args.includes('--reuse');
  fs.mkdirSync(path.join(out, 'pkg'), { recursive: true });
  const bases = new Map<string, Awaited<ReturnType<typeof loadKeySource>>>();
  const baseOf = async (file: string) => { let b = bases.get(file); if (!b) bases.set(file, (b = await loadKeySource(path.join(corpus, file), 'b'))); return b; };
  const jobs: Job[] = [];
  const excelJobs: Array<{ id: string; file: string; kind: 'output' | 'reduced'; sheet: string; deleteRows?: number[] }> = [];
  const entries = loadManifest(corpus).filter((e) => !only || only.test(e.file));
  for (const e of entries) {
    // the Excel-saved twin of the source (cached results); the random corpus has none: Excel computes the source itself
    const twinPath = random ? path.join(corpus, 'nocache', e.file) : path.join(corpus, 'cached', e.file);
    const twin = random ? null : await readBook(new Uint8Array(fs.readFileSync(twinPath)), e.file, { id: 'twin' });
    const srcDump = random ? `${e.name}__${e.order}.src` : undefined;
    if (srcDump) excelJobs.push({ id: srcDump, file: twinPath, kind: 'output', sheet: e.sheet });
    const base = await baseOf(e.base ?? 'base.xlsx');
    const configs = random ? RANDOM_CONFIGS : e.base ? SAMPLE_CONFIGS : CONFIGS;
    for (const cfg0 of quick ? configs.slice(0, 1) : configs) {
      const cfg: Config = sources ? { ...cfg0, name: `${cfg0.name}+src`, sources: true } : cfg0;
      const id = `${e.name}__${e.order}__${cfg.name}`;
      const src = await loadKeySource(path.join(corpus, cfg.results === 'cached' ? 'cached' : 'nocache', e.file), 's1');
      const p = buildPipeline(base, src, { imports: cfg.imports, joinType: cfg.join });
      const trace: FormulaTrace[] = [];
      const res = await buildPremerged(toPremergeInput(p, { ...(cfg.drop ? { dropOtherSheets: true } : {}), ...(cfg.sources ? { includeSourceSheets: true } : {}) }), undefined, (t) => trace.push(t));
      const file = path.join(out, 'pkg', `${id}.xlsx`);
      fs.writeFileSync(file, res.bytes);
      const placed = placedRows(p.join, 's1', p.base.region.lastDataRow);
      const placedSet = new Set(placed.map((x) => x[0]));
      const r = src.region;
      const del: number[] = [];
      for (let row = r.firstDataRow; row <= r.lastDataRow; row++) if (!placedSet.has(row) && sheetOf(src.book, r.sheet)?.rows[row - 1]?.some((c) => c && (c.v != null || c.f))) del.push(row);
      const kept = trace.filter((t) => t.outcome === 'kept');
      const job: Job = {
        id, entry: e, cfg, kept, srcDump,
        counts: { kept: kept.length, value: trace.filter((t) => t.outcome === 'value').length, empty: trace.filter((t) => t.outcome === 'empty').length, viaCopy: kept.filter((t) => t.viaCopy).length },
        mergedSheet: res.layout.sheet,
        reduced: kept.some((t) => t.widened) && del.length ? { file: twinPath, deleteRows: del.sort((a, b) => b - a) } : null,
        placed, twin, order: e.order,
        lastCol: Math.max(...res.layout.imported.map((x) => x.col), p.base.region.lastCol),
        moveFrom: p.base.region.lastCol,
      };
      jobs.push(job);
      excelJobs.push({ id, file, kind: 'output', sheet: res.layout.sheet });
      // the user's other complaint: move a column across the imported block with the app's own engine (replayLog) — every value must stay where
      // it belongs (the base's last column goes behind the last column, everything right of it shifts one to the left)
      const moved = await replayLog(res.bytes, [{ kind: 'struct', sheet: res.layout.sheet, op: { type: 'move', axis: 'col', from: p.base.region.lastCol, count: 1, to: job.lastCol + 1 } }]);
      const movedFile = path.join(out, 'pkg', `${id}.moved.xlsx`);
      fs.writeFileSync(movedFile, moved);
      excelJobs.push({ id: `${id}.mv`, file: movedFile, kind: 'output', sheet: res.layout.sheet });
      if (job.reduced) excelJobs.push({ id: `${id}.red`, file: twinPath, kind: 'reduced', sheet: r.sheet, deleteRows: job.reduced.deleteRows });
    }
  }
  fs.writeFileSync(path.join(out, 'jobs.json'), JSON.stringify(excelJobs, null, 1));
  if (noExcel) { console.log(`wrote ${jobs.length} packages to ${out}\\pkg (Excel step skipped)`); return; }
  // ---- Excel: one hidden instance for all jobs
  const ps1 = path.join(path.dirname(fileURLToPath(import.meta.url)), 'oracle-excel.ps1');
  if (!reuse) execFileSync('powershell.exe', ['-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', ps1, '-JobsFile', path.join(out, 'jobs.json'), '-OutDir', out], { stdio: ['ignore', 'inherit', 'inherit'], timeout: 40 * 60_000 });
  // ---- compare
  const results: Array<Record<string, unknown>> = [];
  for (const j of jobs) {
    const dump = JSON.parse(fs.readFileSync(path.join(out, `${j.id}.values.json`), 'utf8')) as Dump;
    const red = j.reduced ? (JSON.parse(fs.readFileSync(path.join(out, `${j.id}.red.values.json`), 'utf8')) as Dump) : null;
    const stored = await readBook(new Uint8Array(fs.readFileSync(path.join(out, 'pkg', `${j.id}.xlsx`))), `${j.id}.xlsx`, { id: 'pkg' });
    const src = j.srcDump ? (JSON.parse(fs.readFileSync(path.join(out, `${j.srcDump}.values.json`), 'utf8')) as Dump) : null;
    const r = compareJob(j, dump, red, stored, src);
    const mv = JSON.parse(fs.readFileSync(path.join(out, `${j.id}.mv.values.json`), 'utf8')) as Dump;
    Object.assign(r, compareMoved(j, dump, mv));
    results.push(r);
  }
  fs.writeFileSync(path.join(out, 'results.json'), JSON.stringify(results, null, 1));
  const md = markdown(results);
  fs.writeFileSync(path.join(out, 'results.md'), md);
  console.log(md);
  const bad = results.filter((r) => r.open !== 'normal' || (r.mismatches as number) > 0 || r.moveOpen !== 'normal' || (r.moveMismatch as number) > 0);
  if (bad.length) process.exitCode = 1;
}

/** the moved package must hold the same values at the moved positions (formulas adjusted by the app's engine, recalculated by Excel) */
function compareMoved(j: Job, before: Dump, after: Dump): Record<string, unknown> {
  const res = { moveOpen: after.open, moveCompared: 0, moveMismatch: 0, moveExamples: [] as unknown[] };
  if (after.open === 'failed') return res;
  const to = (c: number) => (c === j.moveFrom ? j.lastCol : c > j.moveFrom && c <= j.lastCol ? c - 1 : c);
  const norm = (v: XVal) => (v && typeof v === 'object' ? ERR[v.e] ?? `#E${v.e}` : v);
  for (let r = 0; r < before.rows.length; r++) {
    for (let c = 1; c <= j.lastCol; c++) {
      const a = norm(before.rows[r]?.[c - 1] ?? null), b = norm(after.rows[r]?.[to(c) - 1] ?? null);
      res.moveCompared++;
      if (!same(a, b)) {
        res.moveMismatch++;
        if (res.moveExamples.length < 6) res.moveExamples.push({ row: r + 1, col: c, before: a, after: b });
      }
    }
  }
  return res;
}

function expectedOf(twin: SourceBook, sheet: string, row: number, col: number): XVal {
  const c = sheetOf(twin, sheet)?.rows[row - 1]?.[col - 1];
  if (!c || c.v == null || c.t === 'z') return null;
  if (c.t === 'e') return String(c.v);
  return c.v;
}

/** does the formula read the merged sheet itself (a reference without a sheet prefix)? */
function readsMerged(formula: string): boolean {
  return tokenize(formula).some((t) => (t.type === 'ref' || t.type === 'name' || t.type === 'sref') && !t.prefix);
}

function same(exp: XVal, act: XVal): boolean {
  const a = act && typeof act === 'object' ? ERR[act.e] ?? `#E${act.e}` : act;
  const e = exp;
  if ((e === null || e === '') && (a === null || a === '')) return true;
  // Excel stores the new #CALC! error as #VALUE! in the file
  if (e === '#VALUE!' && a === '#CALC!') return true;
  if (typeof e === 'number' && typeof a === 'number') return Math.abs(e - a) <= 1e-9 * Math.max(1, Math.abs(e));
  return e === a;
}

function compareJob(j: Job, dump: Dump, red: Dump | null, stored: SourceBook, src: Dump | null): Record<string, unknown> {
  const skipNotes: Record<string, number> = {};
  const bySource: Record<string, number> = {};
  const res: Record<string, unknown> = { id: j.id, item: j.entry.name, order: j.entry.order, config: j.cfg.name, open: dump.open, log: dump.log ?? '', ...j.counts, compared: 0, nonEmpty: 0, errorValues: 0, widenedCompared: 0, skipped: 0, skipNotes, mismatches: 0, newErrors: 0, staleCache: 0, cachedChecked: 0, bySource, examples: [] as unknown[] };
  if (dump.open === 'failed') return res;
  const sheet = j.entry.sheet;
  const examples = res.examples as unknown[];
  const delBelow = (row: number) => (j.reduced ? row - j.reduced.deleteRows.filter((d) => d < row).length : row);
  /** every placed source row sits on the same row number in the merged sheet: positions inside the table did not change */
  const identical = j.placed.every(([s, t]) => s === t);
  const skip = (note: string) => { (res.skipped as number)++; skipNotes[note] = (skipNotes[note] ?? 0) + 1; };
  for (const t of j.kept) {
    const act = dump.rows[t.row - 1]?.[t.col - 1] ?? null;
    // the result stored next to a kept formula must be what Excel computes now (never a stale value next to a changed formula)
    if (t.cached) {
      const sc = sheetOf(stored, j.mergedSheet)?.rows[t.row - 1]?.[t.col - 1];
      const sv: XVal = !sc || sc.v == null || sc.t === 'z' ? null : sc.t === 'e' ? String(sc.v) : sc.v;
      const av: XVal = act && typeof act === 'object' ? ERR[act.e] ?? `#E${act.e}` : act;
      (res.cachedChecked as number)++;
      if (!same(sv, av)) {
        (res.staleCache as number)++;
        (res.mismatches as number)++;
        bySource[`STALE ${t.source}`] = (bySource[`STALE ${t.source}`] ?? 0) + 1;
        if (examples.length < 8) examples.push({ stale: true, row: t.row, col: t.col, source: t.source, formula: t.formula, stored: sv, actual: av });
      }
    }
    let exp: XVal;
    // a formula that reads a header cell whose text changed (two-row header joined, suffix added): its result differs on purpose
    if (t.exact === false && !t.widened) { skip('reads a header cell whose text changed'); continue; }
    if (t.widened && /\b(ROWS|COUNTBLANK)\(/.test(t.source)) { skip('counts the rows of a widened table'); continue; }
    if (t.widened && t.substituted && /\bCOUNTA\(/.test(t.source)) { skip('counts non-blank key cells of a widened table'); continue; }
    if (t.widened && /\b(UNIQUE|FILTER|SORT|SORTBY|TAKE)\(/.test(t.source)) { skip('array function over a widened table (row order / spill)'); continue; }
    // (a formula that reads ONLY copies of the source's sheets sees the source's own layout: its MATCH position is the source's)
    // (ISNUMBER(MATCH(..)) / ISNA(MATCH(..)) only ask WHETHER something is found: value-based)
    const positionOnly = POSITION_ONLY.test(t.source.replace(/\b(ISNUMBER|ISNA|ISERROR)\(MATCH\(/g, '('));
    if (positionOnly && !/\bINDEX\(/.test(t.source) && !identical && !(t.viaCopy && !readsMerged(t.formula ?? ''))) { skip('MATCH position in a table whose layout changed'); continue; }
    if (t.widened && red) {
      exp = red.rows[delBelow(t.srcRow) - 1]?.[t.srcCol - 1] ?? null;
      (res.widenedCompared as number)++;
    } else {
      // the Excel-saved twin, or (random corpus) the values Excel computed for the source workbook itself
      exp = j.twin ? expectedOf(j.twin, sheet, t.srcRow, t.srcCol) : src?.rows[t.srcRow - 1]?.[t.srcCol - 1] ?? null;
    }
    if (exp && typeof exp === 'object') exp = ERR[(exp as { e: number }).e] ?? `#E${(exp as { e: number }).e}`;
    (res.compared as number)++;
    // how many comparisons are about a real value (not an empty cell) / an error value: a property test over empty cells proves nothing
    if (exp !== null && exp !== '') { (res.nonEmpty as number)++; if (typeof exp === 'string' && /^#/.test(exp)) (res.errorValues as number)++; }
    if (!same(exp, act)) {
      const isErr = act && typeof act === 'object';
      if (isErr) (res.newErrors as number)++;
      (res.mismatches as number)++;
      bySource[t.source] = (bySource[t.source] ?? 0) + 1;
      if (examples.length < 8) examples.push({ srcRow: t.srcRow, srcCol: t.srcCol, row: t.row, col: t.col, source: t.source, formula: t.formula, widened: !!t.widened, expected: exp, actual: act });
    }
  }
  return res;
}

function markdown(rs: Array<Record<string, unknown>>): string {
  const out: string[] = ['| job | open | kept (via copy) | value | empty | compared (widened) | skipped | MISMATCH | new errors | stale cache | move: open / cells / diff |', '|---|---|---|---|---|---|---|---|---|---|---|'];
  for (const r of rs) out.push(`| ${r.id} | ${r.open} | ${r.kept} (${r.viaCopy}) | ${r.value} | ${r.empty} | ${r.compared} (${r.widenedCompared}) | ${r.skipped} | ${r.mismatches} | ${r.newErrors} | ${r.staleCache}/${r.cachedChecked} | ${r.moveOpen} / ${r.moveCompared} / ${r.moveMismatch} |`);
  const sum = (k: string) => rs.reduce((a, r) => a + (r[k] as number), 0);
  out.push(`| **total** (non-empty ${sum('nonEmpty')}, error values ${sum('errorValues')}) | ${rs.every((r) => r.open === 'normal') ? 'all normal' : 'NOT ALL NORMAL'} | ${sum('kept')} (${sum('viaCopy')}) | ${sum('value')} | ${sum('empty')} | ${sum('compared')} (${sum('widenedCompared')}) | ${sum('skipped')} | ${sum('mismatches')} | ${sum('newErrors')} | ${sum('staleCache')}/${sum('cachedChecked')} | ${rs.every((r) => r.moveOpen === 'normal') ? 'all normal' : 'NOT ALL NORMAL'} / ${sum('moveCompared')} / ${sum('moveMismatch')} |`);
  for (const r of rs.filter((x) => (x.mismatches as number) > 0)) {
    out.push(`\nMISMATCH ${r.id}:`);
    for (const e of r.examples as unknown[]) out.push('  ' + JSON.stringify(e));
  }
  for (const r of rs.filter((x) => (x.moveMismatch as number) > 0 || x.moveOpen !== 'normal')) {
    out.push(`\nMOVE ${r.moveOpen !== 'normal' ? 'NOT OPENED NORMALLY' : 'MISMATCH'} ${r.id}:`);
    for (const e of r.moveExamples as unknown[]) out.push('  ' + JSON.stringify(e));
  }
  return out.join('\n') + '\n';
}
