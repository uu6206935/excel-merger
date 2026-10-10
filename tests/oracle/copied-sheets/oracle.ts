/**
 * Copied-sheets oracle in real Excel (task V9): the output of "plan A" — the merged sheet first, then the original sheets of every file as
 * <file>_<sheet> — must be exactly what the user asked for, in BOTH export engines, after random edits of the merged sheet.
 * Re-runnable:
 *
 *   powershell -NoProfile -ExecutionPolicy Bypass -File tests\oracle\copied-sheets\run.ps1 [-Out <dir>] [-Only <regex>] [-Limit N] [-CheckOutput N] [-NoRunner] [-NoExcel] [-Reuse]
 *   node tests/oracle/copied-sheets/run.mjs --out <dir> [--only <regex>] [--limit N] [--check-output N] [--negative] [--seed-base 0] [--no-runner] [--no-excel] [--reuse]
 *
 * Method (see RESULTS.md): for every combination (combos.ts: samples in several card orders, foreign corpus files, Excel-authored object fixtures,
 * names that collide / need shortening, the rich formula sets) the app's pipeline builds the premerged package WITH copies (pipeline.ts); a random op
 * log is made through the doc API (consistency/ops.ts: moves incl. non-contiguous + title unmerge, inserts, deletes with convert, typed cells); the log is
 * exported by replayLog (built-in engine) and by the real Excel through excel-runner.ps1; ONE hidden Excel per batch opens the premerged package,
 * both outputs and every source file, recalculates fully (CalculateFull) and dumps values, formulas and how Excel renders the sheets (excel-book.ps1);
 * judge.ts and the consistency judge compare in Node. Failure classes: see judge.ts / RESULTS.md.
 */
import crypto from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import { compare } from '../../doc/compare';
import { rng } from '../../doc/fuzzcore';
import { lintPackage, lintText } from '../../xlsx/addsheets-helpers';
import { readBook, replayLog } from '../../../src/core/xlsx';
import type { KeySource } from '../../../src/core/contracts';
import type { MergedDoc, SourceBook } from '../../../src/core/types';
import { planSheetNames, sheetNameKey, SHEET_NAME_MAX } from '../../../src/core/xlsx/sheetnames';
import { judgeConverted, judgeOutput, judgeParity, judgeRefIdentity, KNOWN_OPEN, type DocInfo, type Finding, type Stats } from '../consistency/check';
import { baselineLog, docIdentity, generateOps } from '../consistency/ops';
import { pivotSheetNames, REPO } from '../consistency/lib';
import { packageFindings } from '../consistency/pkgdiff';
import { buildSrcChecks, judgeSrc, type SrcCheck } from '../consistency/srcCheck';
import { buildCombos, writeGenerated, type Combo, type Participant } from './combos';
import { ensureDisk, readBook as readDump, runBooks, runRunnerJob, type BookDump, type BookJob } from './excel';
import { expectedSheets, INFO_CLASSES, judgeCopies, judgeCopiesByReader, judgeCopyNames, judgeMergedErrors, judgeNames, type CopyInfo, type CopyRec, type SourceRec, type Which } from './judge';
import { DEFAULT_PARAMS, loadSpec, mergeWithCopies, type MergedWithCopies, type PipelineParams } from './pipeline';
import { runNegative, MUTATIONS } from './negative';
import { toolChecks } from './toolcheck';

function argv(args: string[], k: string, d = ''): string {
  const i = args.indexOf(`--${k}`);
  return i >= 0 && args[i + 1] ? args[i + 1] : d;
}

const hash = (s: string): number => [...s].reduce((h, ch) => (Math.imul(h, 31) + ch.charCodeAt(0)) >>> 0, 7);
const safe = (s: string) => s.replace(/[^A-Za-z0-9_.-]/g, '_');
const sha = (b: Uint8Array): string => crypto.createHash('sha1').update(b).digest('hex').slice(0, 12);
const extOf = (fileName: string): string => (/\.([A-Za-z0-9]{2,5})$/.exec(fileName)?.[1] ?? 'xlsx').toLowerCase();
const jobIdOf = (file: string): string => safe(path.basename(file).replace(/\.(xlsx|xlsm)$/i, ''));
const docInfo = (d: MergedDoc): DocInfo => ({ names: d.names, tables: d.features.tables, ids: d.columns.map((c) => c.id) });
const refRuleCount = (d: MergedDoc): number => d.features.rules.filter((r) => r.formulas.some((f) => /#REF!/.test(f))).length + d.externalRules.filter((e) => e.rule.formulas.some((f) => /#REF!/.test(f))).length;

export interface RunnerRec {
  ok: boolean;
  code?: string;
  message?: string;
  warnings: Array<{ code: string; message: string }>;
  ms: number;
  refErrors?: { cells: number; names: number; samples: string[] };
}

export interface OracleResult {
  code: number;
  recs: CopyRec[];
  findings: Finding[];
  summary: string;
}

/** a source file as the oracle describes it (Excel is asked to open a copy of the original bytes under srcs/) */
function sourceRec(ks: KeySource, mode: 'join' | 'copy', srcDir: string): SourceRec {
  const bytes = ks.book.bytes;
  const ext = extOf(ks.book.fileName);
  const file = path.join(srcDir, `${sha(bytes)}.${ext}`);
  if (!fs.existsSync(file)) fs.writeFileSync(file, bytes);
  const fidelity: SourceRec['fidelity'] = /^(csv|tsv|txt)$/.test(ext) ? 'reader' : ext === 'xls' ? 'values' : 'full';
  return {
    id: ks.id, fileName: ks.book.fileName, path: file, mode, fidelity,
    sheets: ks.book.sheets.map((s) => s.name),
    date1904: !!ks.book.date1904,
    externalLinks: (ks.book.features.externalLinks ?? 0) > 0,
  };
}

interface Built {
  combo: Combo;
  merged: MergedWithCopies;
  base: KeySource;
  preFile: string;
  srcs: SourceRec[];
  baseSpec: SourceRec;
  pivots: string[];
  limits: string[];
  copies: CopyInfo[];
}

interface Dirs {
  genDir: string;
  pkgDir: string;
  srcDir: string;
}

async function build(combo: Combo, dirs: Dirs): Promise<Built> {
  let base: KeySource;
  let participants: Participant[];
  if (combo.build) {
    const b = await combo.build({ genDir: dirs.genDir });
    base = b.base;
    participants = b.others;
  } else {
    base = await loadSpec(combo.base!, 'b');
    participants = [];
    for (const [i, o] of (combo.others ?? []).entries()) participants.push({ ks: await loadSpec(o, `s${i + 2}`), mode: o.mode ?? 'join' });
  }
  const params: PipelineParams = { ...DEFAULT_PARAMS, ...(combo.params ?? {}) };
  const merged = await mergeWithCopies(base, participants, params);
  const ext = merged.format === 'xlsm' ? 'xlsm' : 'xlsx';
  const preFile = path.join(dirs.pkgDir, `${safe(combo.id)}.pre.${ext}`);
  fs.writeFileSync(preFile, merged.result.bytes);
  const modeOf = (id: string): 'join' | 'copy' => participants.find((p) => p.ks.id === id)?.mode ?? 'join';
  const srcs = merged.participants.map((ks) => sourceRec(ks, modeOf(ks.id), dirs.srcDir));
  const layout = merged.result.layout;
  const copies: CopyInfo[] = (layout.addedSheets ?? []).map((a) => ({ sourceId: a.sourceId, file: a.file, original: a.original, name: a.name, suffixed: a.suffixed, complete: a.complete }));
  const limits = merged.result.report.items.filter((i) => i.code === 'sheetCopyLimits').map((i) => i.message);
  return { combo, merged, base, preFile, srcs, baseSpec: sourceRec(base, 'join', dirs.srcDir), pivots: await pivotSheetNames(merged.result.bytes), limits, copies };
}

/** what needs no Excel: the contract of the layout / report (names legal, unique, suffixes only where a name was taken, report = layout) */
export function staticChecks(id: string, merged: MergedWithCopies, drop: boolean): Finding[] {
  const out: Finding[] = [];
  const add = (cls: string, detail: string) => out.push({ scenario: id, cls, detail });
  const layout = merged.result.layout;
  const added = layout.addedSheets ?? [];
  const reported = merged.result.report.addedSheets ?? [];
  if (JSON.stringify(reported) !== JSON.stringify(added.map((a) => a.name))) add('static-report-layout', `MergeReport.addedSheets ${JSON.stringify(reported)} != PremergeLayout.addedSheets ${JSON.stringify(added.map((a) => a.name))}`);
  const taken: string[] = [layout.sheet];
  if (!drop) taken.push(...merged.base.book.sheets.map((s) => s.name).filter((n) => n !== layout.sheet));
  const seen = new Set(taken.map(sheetNameKey));
  for (const a of added) {
    if (a.name.length > SHEET_NAME_MAX) add('static-name', `"${a.name}" is longer than 31 characters`);
    if (/[\\/?*[\]:]/.test(a.name) || /^'|'$/.test(a.name)) add('static-name', `"${a.name}" holds a character Excel does not allow`);
    const key = sheetNameKey(a.name);
    if (seen.has(key)) add('static-name', `"${a.name}" is already taken (case-insensitive) by an earlier sheet`);
    seen.add(key);
    const plain = a.name.replace(/\(\d+\)$/, '');
    if (a.suffixed && !taken.some((t) => sheetNameKey(t) === sheetNameKey(plain)) && !added.slice(0, added.indexOf(a)).some((x) => sheetNameKey(x.name) === sheetNameKey(plain))) add('static-suffix', `"${a.name}" got a suffix but "${plain}" was free`);
    if (!a.suffixed && /\(\d+\)$/.test(a.name) && !/\(\d+\)$/.test(a.original)) add('static-suffix', `"${a.name}" ends with (n) but is marked as not suffixed`);
    taken.push(a.name);
  }
  return out;
}

interface Made {
  rec: CopyRec;
  log: unknown;
  baseBytes: Uint8Array | null;
  jsBytes: Uint8Array | null;
}

async function makeScenario(b: Built, seed: number, pkgDir: string): Promise<Made | null> {
  const { merged, combo } = b;
  const ck = safe(combo.id);
  const r = rng(seed * 7919 + hash(ck));
  const mode: 'inv' | 'del' = r() < 0.62 ? 'inv' : 'del';
  const allowAffecting = r() < 0.4;
  const steps = 3 + Math.floor(r() * 7);
  const g = generateOps(merged.editor, r, { steps, mode, allowAffecting });
  if (!g.final.log.length) return null;
  // a formula typed into a cell with the Text format stays TEXT (Excel does the same; the doc models it): text is not adjusted by later moves, so the
  // baseline must hold the text as it was typed, not rendered back to the original positions like a formula (the consistency oracle assumes formulas)
  for (const e of g.edits) {
    const cell = g.final.doc.columns.find((c) => c.id === e.colId)?.cells[e.row - 1];
    if (cell && cell.f === undefined && typeof cell.v === 'string' && cell.v.startsWith('=')) e.origText = e.text;
  }
  const ext = merged.format === 'xlsm' ? 'xlsm' : 'xlsx';
  const id = `${ck}__s${seed}`;
  const files = { pre: path.basename(b.preFile), base: path.basename(b.preFile), js: `${id}.js.${ext}`, xl: `${id}.xl.${ext}` };
  const params = { ...DEFAULT_PARAMS, ...(combo.params ?? {}) };
  const rec: CopyRec = {
    id, set: combo.group, variant: combo.id, seed, mode, allowAffecting,
    mergedSheet: merged.editor.doc.sheetName,
    headerRow: merged.editor.doc.headerRow, firstDataRow: merged.editor.doc.firstDataRow, lastDataRow: merged.editor.doc.lastDataRow,
    initial: docInfo(merged.editor.doc), final: docInfo(g.final.doc), deleted: g.deleted,
    converted: g.converted.map((x) => ({ sheet: x.sheet, id: x.id, row: x.row, col: x.col, value: x.write.v, blank: x.write.t === 'z' && x.write.v === null })),
    trail: g.trail, lookupWarned: g.lookupWarned, affecting: g.affecting, opsKinds: g.opsKinds, logLength: g.final.log.length,
    baselineIsPre: true, refRules: { initial: refRuleCount(merged.editor.doc), final: refRuleCount(g.final.doc) },
    files, pivotSheets: b.pivots, pivotComparable: !g.edits.length && !g.converted.length && !g.deleted.length, docVsReplay: [],
    combo: combo.id, group: combo.group, drop: params.dropOtherSheets, ext,
    baseFile: b.baseSpec.path, baseFileName: b.baseSpec.fileName, sources: b.srcs, copies: b.copies, limits: b.limits,
  };
  // baseline = premerged + typed values + converted values, no layout operation
  const bl = baselineLog(merged.editor, g);
  let baseBytes: Uint8Array | null = null;
  if (bl === null) rec.skipped = 'the doc refused a typed text in the original layout';
  else if (bl.log.length) {
    baseBytes = await replayLog(merged.result.bytes, bl.log);
    rec.baselineIsPre = false;
    rec.files.base = `${id}.base.${ext}`;
  }
  let jsBytes: Uint8Array | null = null;
  try {
    jsBytes = await replayLog(merged.result.bytes, g.final.log);
    const book = await readBook(jsBytes, `js.${ext}`, { id: 'js' });
    rec.docVsReplay = compare(g.final.doc, book).slice(0, 12);
  } catch (e) {
    rec.docVsReplay = [`replayLog threw: ${(e as Error).message}`];
    rec.skipped = rec.skipped ?? 'replayLog threw';
  }
  const ident = docIdentity(merged.editor.doc, g.final.doc, new Set(g.deleted), new Set(g.edits.map((e) => `${e.colId}:${e.row}`)));
  if (ident.length) rec.docVsReplay.push(...ident.slice(0, 4).map((x) => `ref-identity(doc): ${x}`));
  void pkgDir;
  return { rec, log: g.final.log, baseBytes, jsBytes };
}

export async function runOracle(args: string[]): Promise<OracleResult> {
  const out = path.resolve(argv(args, 'out', path.join(REPO, '..', '.tmp', 'v9', 'oracle')));
  const seedBase = Number(argv(args, 'seed-base', '0'));
  const only = argv(args, 'only') ? new RegExp(argv(args, 'only')) : null;
  const limit = Number(argv(args, 'limit', '0'));
  // --check-output N: the QA tools (check-output --single-sheet / --with-copies, compare) over N scenarios of every group, both engines (one more Excel instance)
  const toolN = Number(argv(args, 'check-output', '0'));
  // --negative: planted faults (negative.ts) in the premerged packages of a few scenarios, judged by the same judge: every one must be found
  const negative = args.includes('--negative');
  const noExcel = args.includes('--no-excel');
  const noRunner = args.includes('--no-runner') || noExcel;
  const reuse = args.includes('--reuse');
  const log = (s: string) => console.log(`[${new Date().toISOString().slice(11, 19)}] ${s}`);
  const tempRoot = path.join(out, 'xltmp');
  const dumpDir = path.join(out, 'dump');
  const pkgDir = path.join(out, 'pkg');
  const srcDir = path.join(out, 'srcs');
  const genDir = path.join(out, 'gen');
  const twinDir = path.join(out, 'twin');
  for (const d of [out, dumpDir, pkgDir, srcDir, genDir, twinDir]) fs.mkdirSync(d, { recursive: true });
  if (!noExcel) ensureDisk(out);

  // ---- A. scenarios (the Node half)
  const recsFile = path.join(out, 'scenarios.json');
  const srcFile = path.join(out, 'srcchecks.json');
  const staticFile = path.join(out, 'static.json');
  let recs: CopyRec[] = [];
  let srcChecks: Array<SrcCheck & { skip: boolean }> = [];
  const staticFindings: Finding[] = [];
  if (reuse && fs.existsSync(recsFile)) {
    recs = JSON.parse(fs.readFileSync(recsFile, 'utf8')) as CopyRec[];
    srcChecks = JSON.parse(fs.readFileSync(srcFile, 'utf8')) as Array<SrcCheck & { skip: boolean }>;
    if (fs.existsSync(staticFile)) staticFindings.push(...(JSON.parse(fs.readFileSync(staticFile, 'utf8')) as Finding[]));
    if (only) {
      recs = recs.filter((r) => only.test(r.combo));
      srcChecks = srcChecks.filter((c) => only.test(c.combo));
    }
    log(`reusing ${recs.length} scenarios from ${recsFile}`);
  } else {
    await writeGenerated(genDir);
    let combos = buildCombos().filter((c) => !only || only.test(c.id));
    if (limit > 0) combos = combos.slice(0, limit);
    log(`${combos.length} combinations planned`);
    let skipped = 0;
    for (const combo of combos) {
      let b: Built;
      try {
        b = await build(combo, { genDir, pkgDir, srcDir });
      } catch (e) {
        log(`SETUP FAILED ${combo.id}: ${(e as Error).message}`);
        staticFindings.push({ scenario: combo.id, cls: 'setup-failed', detail: (e as Error).message.slice(0, 400) });
        continue;
      }
      staticFindings.push(...staticChecks(combo.id, b.merged, { ...DEFAULT_PARAMS, ...(combo.params ?? {}) }.dropOtherSheets));
      for (const c of buildSrcChecks(combo.id, b.merged, twinDir, false)) {
        const s = b.srcs.find((x) => x.id === c.sourceId);
        srcChecks.push({ ...c, skip: !s || s.fidelity !== 'full' || s.mode !== 'join' });
      }
      for (const seed of combo.seeds ?? [1]) {
        const made = await makeScenario(b, seed + seedBase, pkgDir);
        if (!made) { skipped++; continue; }
        recs.push(made.rec);
        fs.writeFileSync(path.join(pkgDir, `${made.rec.id}.log.json`), JSON.stringify(made.log));
        if (made.baseBytes) fs.writeFileSync(path.join(pkgDir, made.rec.files.base), made.baseBytes);
        if (made.jsBytes) fs.writeFileSync(path.join(pkgDir, made.rec.files.js), made.jsBytes);
      }
      log(`${combo.id}: ${b.copies.length} copies${b.limits.length ? `, ${b.limits.length} limit note(s)` : ''}`);
    }
    fs.writeFileSync(recsFile, JSON.stringify(recs, null, 1));
    fs.writeFileSync(srcFile, JSON.stringify(srcChecks, null, 1));
    fs.writeFileSync(staticFile, JSON.stringify(staticFindings, null, 1));
    log(`${recs.length} scenarios built, ${skipped} skipped (no accepted op)`);
  }

  // ---- B. the Excel engine (one Excel start per scenario: that is how the runner works)
  const runnerFile = path.join(out, 'runner.json');
  const runner: Record<string, RunnerRec> = fs.existsSync(runnerFile) ? (JSON.parse(fs.readFileSync(runnerFile, 'utf8')) as Record<string, RunnerRec>) : {};
  if (!noRunner) {
    const cacheDir = path.join(out, 'runner-cache');
    let done = 0;
    for (const rec of recs) {
      if (rec.skipped) continue;
      const xl = path.join(pkgDir, rec.files.xl);
      if (runner[rec.id] && (fs.existsSync(xl) || !runner[rec.id].ok)) continue;
      const logOps = JSON.parse(fs.readFileSync(path.join(pkgDir, `${rec.id}.log.json`), 'utf8'));
      const res = await runRunnerJob(path.join(pkgDir, rec.files.pre), logOps, xl, rec.mergedSheet, { tempRoot, cacheDir, format: rec.ext });
      runner[rec.id] = res;
      if (!res.ok) fs.rmSync(xl, { force: true });
      fs.writeFileSync(runnerFile, JSON.stringify(runner, null, 1));
      if (++done % 5 === 0) log(`runner: ${done} jobs done (last ${res.ms} ms${res.ok ? '' : ` FAILED ${res.code}`})`);
    }
    log(`runner: ${done} new jobs, ${Object.values(runner).filter((r) => !r.ok).length} failed in total`);
  }

  // ---- C. Excel: sources, premerged package, baseline and both outputs of every scenario, one hidden Excel per batch
  const jobs = new Map<string, BookJob>();
  const addJob = (file: string, deep: BookJob['deep']) => {
    const id = jobIdOf(file);
    const have = jobs.get(id);
    if (!fs.existsSync(file)) return;
    // a file that is both a source and a base is dumped once, deeply
    if (!have || (have.deep === 'none' && deep !== 'none')) jobs.set(id, { id, file, deep });
    else if (Array.isArray(have.deep) && Array.isArray(deep)) have.deep = [...new Set([...have.deep, ...deep])];
  };
  for (const rec of recs) {
    if (rec.skipped) continue;
    for (const s of rec.sources) addJob(s.path, 'all');
    addJob(rec.baseFile, 'none');
    const copyNames = [...new Set([...rec.copies.map((c) => c.name), ...plannedFromReader(rec)])];
    addJob(path.join(pkgDir, rec.files.pre), copyNames);
    addJob(path.join(pkgDir, rec.files.base), 'none');
    addJob(path.join(pkgDir, rec.files.js), copyNames);
    if (runner[rec.id]?.ok) addJob(path.join(pkgDir, rec.files.xl), copyNames);
  }
  if (!noExcel) {
    log(`excel: ${jobs.size} workbooks to open and recalculate`);
    await runBooks([...jobs.values()], dumpDir, { tempRoot, batchSize: 24, log });
  }

  // ---- D. judge
  const findings: Finding[] = [...staticFindings];
  const perScenario: Array<{ id: string; combo: string; copies: number; sheetsCompared: number; cells: number; findings: number; explained: Record<string, number> }> = [];
  const dump = (file: string): BookDump | null => readDump(dumpDir, jobIdOf(file));
  const bookCache = new Map<string, SourceBook>();
  const bookOf = async (file: string, fileName: string, id: string): Promise<SourceBook> => {
    const key = `${file}|${fileName}|${id}`;
    let b = bookCache.get(key);
    if (!b) {
      b = await readBook(new Uint8Array(fs.readFileSync(file)), fileName, { id });
      bookCache.set(key, b);
    }
    return b;
  };
  const totals = { sheets: 0, cells: 0, formatCells: 0, copies: 0, ops: {} as Record<string, number> };
  for (const rec of recs) {
    if (rec.skipped) { findings.push({ scenario: rec.id, cls: 'skipped', detail: rec.skipped }); continue; }
    const f0 = findings.length;
    for (const [k, v] of Object.entries(rec.opsKinds)) totals.ops[k] = (totals.ops[k] ?? 0) + v;
    totals.copies += rec.copies.length;
    if (rec.docVsReplay.length) findings.push({ scenario: rec.id, cls: 'doc-vs-replay', detail: rec.docVsReplay.slice(0, 4).join(' | ') });
    const xlRun = runner[rec.id];
    if (xlRun && !xlRun.ok) {
      // (the app answers these by exporting with the built-in engine instead: src/api.ts canFallBackToJs; REFUSED_TABLE = Excel will not move a column out of a table)
      const fallback = /MOVE_VERIFY_FAILED|EXCEL_UNAVAILABLE|PASSWORD_PROTECTED|OPEN_FAILED|RUNNER_MISSING|DISK_LOW|REFUSED_TABLE/.test(xlRun.code ?? '');
      findings.push({ scenario: rec.id, cls: fallback ? 'runner-fallback' : 'runner-failed', detail: `${xlRun.code}: ${xlRun.message}` });
    }
    const dPre = dump(path.join(pkgDir, rec.files.pre));
    const dBase = dump(path.join(pkgDir, rec.files.base));
    const dJs = dump(path.join(pkgDir, rec.files.js));
    const dXl = xlRun?.ok ? dump(path.join(pkgDir, rec.files.xl)) : null;
    const srcDumps = new Map<string, BookDump | null>(rec.sources.map((s) => [s.id, dump(s.path)]));
    const baseFileDump = dump(rec.baseFile);
    if (!dPre || !dJs || !dBase) { findings.push({ scenario: rec.id, cls: 'no-dump', detail: `pre ${dPre ? 'ok' : 'MISSING'}, base ${dBase ? 'ok' : 'MISSING'}, js ${dJs ? 'ok' : 'MISSING'}` }); continue; }
    const srcBooks = new Map<string, SourceBook>();
    for (const s of rec.sources) srcBooks.set(s.id, await bookOf(s.path, s.fileName, s.id));
    // 1. names / copies vs sources, in every output
    const exp = expectedSheets(rec, srcDumps, baseFileDump);
    const outs: Array<[Which, BookDump | null]> = [['pre', dPre], ['js', dJs], ['xl', dXl]];
    const layoutNames = rec.copies.map((c) => c.name);
    const expCopyNames = exp.copies.map((c) => c.name);
    if (JSON.stringify(layoutNames) !== JSON.stringify(expCopyNames)) findings.push({ scenario: rec.id, cls: 'layout-names', detail: `PremergeLayout.addedSheets ${JSON.stringify(layoutNames)} != planSheetNames over the files ${JSON.stringify(expCopyNames)}` });
    for (const [which, d] of outs) {
      if (!d) continue;
      findings.push(...judgeNames(rec, which, d, exp));
      const j = judgeCopies(rec, which, d, exp, srcDumps, srcBooks);
      findings.push(...j.findings);
      findings.push(...judgeCopyNames(rec, which, d, exp, srcDumps, baseFileDump));
      if (which === 'pre') { totals.sheets += j.stats.sheetsCompared; totals.cells += j.stats.cellsCompared; totals.formatCells += j.stats.formatCells; }
    }
    findings.push(...judgeMergedErrors(rec, dPre, baseFileDump, srcDumps));
    // 2. the reader's view of the package patcher's own writing + the package linter (pre, js)
    for (const [which, file] of [['pre', rec.files.pre], ['js', rec.files.js]] as Array<[Which, string]>) {
      const bytes = new Uint8Array(fs.readFileSync(path.join(pkgDir, file)));
      const book = await bookOf(path.join(pkgDir, file), `${which}.${rec.ext}`, which);
      findings.push(...judgeCopiesByReader(rec, which, book, exp, srcBooks));
      // (an internal hyperlink of a shape is a relationship whose Target is "#'Sheet'!A1" — Excel writes it that way; the linter reads it as a path)
      const lint = (await lintPackage(bytes)).filter((l) => !(l.code === 'REL_DANGLING' && /-> xl\/drawings\/#/.test(l.detail)));
      if (lint.length) findings.push({ scenario: rec.id, cls: `lint-${which}`, detail: lintText(lint).slice(0, 500) });
    }
    // 3. the merged sheet: the consistency oracle's judge (value invariance by column id, new errors, references following their columns)
    //    — it compares EVERY sheet of the baseline with the output, so a copy that changed under the edits is found here too
    const stats: Stats = { cells: 0, formulaCells: 0, mismatches: 0, explained: {} };
    const pre = rec.converted.length ? dPre : undefined;
    const rj = judgeOutput({ rec, base: dBase, out: dJs, engine: 'js', pre });
    findings.push(...rj.findings);
    Object.assign(stats, rj.stats);
    findings.push(...judgeRefIdentity(rec, dBase, dJs, 'js'));
    if (pre && rec.converted.length) {
      const jc = judgeConverted(rec, pre);
      findings.push(...jc.findings);
    }
    if (dXl) {
      const rx = judgeOutput({ rec, base: dBase, out: dXl, engine: 'xl', pre });
      findings.push(...rx.findings);
      findings.push(...judgeRefIdentity(rec, dBase, dXl, 'xl'));
      findings.push(...judgeParity(rec, dJs, dXl));
      try {
        findings.push(...refineEngineFindings(await packageFindings(rec, new Uint8Array(fs.readFileSync(path.join(pkgDir, rec.files.js))), new Uint8Array(fs.readFileSync(path.join(pkgDir, rec.files.xl))))));
      } catch (e) {
        findings.push({ scenario: rec.id, cls: 'engine-package-error', detail: (e as Error).message });
      }
    }
    perScenario.push({ id: rec.id, combo: rec.combo, copies: rec.copies.length, sheetsCompared: 0, cells: stats.cells, findings: findings.length - f0, explained: stats.explained });
  }
  // 4. imported formulas that stayed formulas (some read the copies): same value as in the source for the same person
  let srcCompared = 0;
  for (const c of srcChecks) {
    if (c.skip || (!c.cells.length && !c.placed.length)) continue;
    const rec = recs.find((r) => r.combo === c.combo && !r.skipped);
    if (!rec) continue;
    const pre = dump(path.join(pkgDir, rec.files.pre));
    const s = rec.sources.find((x) => x.id === c.sourceId);
    const twin = s ? dump(s.path) : null;
    if (!pre || !twin) { findings.push({ scenario: c.combo, cls: 'no-dump', detail: `source check: pre ${pre ? 'ok' : 'MISSING'}, source ${twin ? 'ok' : 'MISSING'}` }); continue; }
    // (no reduced twin: a whole-table range over a table with rows Excel had to delete is not compared — it is counted as skipped)
    const cells = c.reducedJob ? c.cells.filter((x) => !x.widened) : c.cells;
    const r = judgeSrc({ ...c, cells }, pre, twin, null);
    findings.push(...r.findings);
    srcCompared += r.stats.keptCompared;
  }
  let negativeTable: Array<{ combo: string; mutation: string; what: string; found: string[]; ok: boolean }> = [];
  if (negative && !noExcel) {
    const wanted = ['smp-01-all-keep', 'ob-all-keep', 'cor-hc-35-names-cf-dv', 'cor-py-19-xlsxwriter-table'];
    const picks = wanted.map((c) => recs.find((r) => r.combo === c && !r.skipped)).filter((r): r is CopyRec => !!r);
    const ctx = new Map<string, { srcDumps: Map<string, BookDump | null>; base: BookDump | null; srcBooks: Map<string, SourceBook> }>();
    for (const rec of picks) {
      const srcBooks = new Map<string, SourceBook>();
      for (const sr of rec.sources) srcBooks.set(sr.id, await bookOf(sr.path, sr.fileName, sr.id));
      ctx.set(rec.id, { srcDumps: new Map(rec.sources.map((sr) => [sr.id, dump(sr.path)])), base: dump(rec.baseFile), srcBooks });
    }
    const neg = await runNegative(picks, { pkgDir, dumpDir, tempRoot }, (rec) => ctx.get(rec.id)!, jobIdOf, log);
    findings.push(...neg.findings);
    negativeTable = neg.table;
    log(`negative controls: ${neg.table.filter((t) => t.ok).length} / ${neg.table.length} planted faults found (${MUTATIONS.length} kinds)`);
  }
  fs.writeFileSync(path.join(out, 'negative.json'), JSON.stringify(negativeTable, null, 1));
  let toolSummary: string[] = [];
  if (toolN > 0 && !noExcel) {
    const perGroup = new Map<string, CopyRec[]>();
    for (const rec of recs) if (!rec.skipped && fs.existsSync(path.join(pkgDir, rec.files.js))) perGroup.set(rec.group, [...(perGroup.get(rec.group) ?? []), rec]);
    const picks: Array<{ rec: CopyRec; js: string; xl: string | null }> = [];
    for (const list of perGroup.values()) {
      const step = Math.max(1, Math.floor(list.length / Math.max(1, Math.ceil(toolN / perGroup.size))));
      for (let i = 0; i < list.length && picks.length < toolN + perGroup.size; i += step) {
        const rec = list[i];
        picks.push({ rec, js: path.join(pkgDir, rec.files.js), xl: runner[rec.id]?.ok ? path.join(pkgDir, rec.files.xl) : null });
      }
    }
    const t = await toolChecks(picks.slice(0, toolN), log);
    findings.push(...t.findings);
    toolSummary = t.stats.summary;
    log(`tools: ${t.stats.verdicts} verdicts and ${t.stats.compared} engine comparisons, ${t.findings.length} finding(s)`);
  }
  fs.writeFileSync(path.join(out, 'tools.json'), JSON.stringify(toolSummary, null, 1));
  fs.writeFileSync(path.join(out, 'findings.json'), JSON.stringify(findings, null, 1));
  fs.writeFileSync(path.join(out, 'results.json'), JSON.stringify(perScenario, null, 1));
  let md = summary(recs, findings, runner, { ...totals, srcCompared });
  if (negativeTable.length) {
    md += `
Negative controls (one planted fault at a time, the same judge): ${negativeTable.filter((t) => t.ok).length} / ${negativeTable.length} found

| scenario | planted fault | found as |
|---|---|---|
${negativeTable.map((t) => `| ${t.combo} | ${t.mutation}: ${t.what} | ${t.ok ? t.found.join(', ') : '**NOT FOUND** ' + t.found.join(', ')} |`).join('\n')}
`;
  }
  if (toolSummary.length) md += `
QA tools over the outputs (check-output --single-sheet / --with-copies, compare):

${toolSummary.map((l) => `- ${l}`).join('\n')}
`;
  fs.writeFileSync(path.join(out, 'results.md'), md);
  console.log(md);
  const failing = findings.filter((f) => !INFO_CLASSES.test(f.cls) && !KNOWN_OPEN.test(f.cls));
  return { code: failing.length ? 1 : 0, recs, findings, summary: md };
}

/** the copies' names as the naming rule gives them over the READER's view of the files (a second opinion for which sheets to dump deeply) */
function plannedFromReader(rec: CopyRec): string[] {
  const base = rec.drop ? [] : [];
  return planSheetNames([rec.mergedSheet, ...base], rec.sources.map((s) => ({ file: s.fileName, sheets: s.sheets }))).map((p) => p.name);
}

/**
 * Differences between the two engines' packages that are only how Excel writes things: a note's line break (Excel stores CR LF, the patcher keeps the
 * LF of the source) is named apart (INFO); everything else stays as the package diff reports it.
 */
export function refineEngineFindings(list: Finding[]): Finding[] {
  const out = list.map((f) => {
    if (f.cls === 'engine-package-comments') {
      const m = /a のみ: ([\s\S]*?)\s+b のみ: ([\s\S]*)$/.exec(f.detail);
      if (m && m[1].replace(/\r\n/g, '\n').trim() === m[2].replace(/\r\n/g, '\n').trim()) return { ...f, cls: 'info-engine-note-linebreak' };
    }
    // the built-in cell style "Normal" is saved by a Japanese Excel as 「標準」 (builtinId 0): the same style
    if (f.cls === 'engine-package-cols') {
      const body = f.detail.replace(/^[^:]*: /, '');
      if (body.split(' ; ').every((p) => /列の書式 style: Normal -> 標準\s*$/.test(p))) return { ...f, cls: 'info-engine-style-name' };
    }
    // a font that names neither a typeface nor a size (ExcelJS writes <font><b/></font>) is completed by Excel with the default ones
    if (/^engine-package-(rows|cols)$/.test(f.cls) && /^[^:]*: (?:\d+ 件: )?(?:[^;]*?(?:行|列)の書式 font: (\S+) -> name=[^,;]+,sz=[\d.]+,\1\s*(?:;\s*)?)+$/.test(f.detail)) return { ...f, cls: 'info-engine-font-complete' };
    // the default column width is written as baseColWidth (Mac Excel) or as defaultColWidth: the same width within half a character
    if (f.cls === 'engine-package-sheetProps') {
      const m = /"tab=([^"/|]*)[/|]baseColWidth=([\d.]*),defaultColWidth=([\d.]*),zeroHeight=([^"]*)" -> "tab=([^"/|]*)[/|]baseColWidth=([\d.]*),defaultColWidth=([\d.]*),zeroHeight=([^"]*)"/.exec(f.detail);
      if (m && m[1] === m[5] && m[4] === m[8]) {
        const eff = (b: string, d: string) => (d ? Number(d) : b ? Number(b) + 5 / 7 : 8.43);
        if (Math.abs(eff(m[2], m[3]) - eff(m[6], m[7])) <= 0.5) return { ...f, cls: 'info-engine-default-width' };
      }
    }
    return f;
  });
  // Excel writes a data validation list that reads ANOTHER sheet into the x14 extension (extLst), the patcher keeps it in <dataValidations>: the same rule
  const only = (side: 'A' | 'B') => out.filter((f) => f.cls === 'engine-package-rules' && new RegExp(`rule only in ${side}: dv${side === 'B' ? '\\(ext\\)' : ''} list `).test(f.detail));
  const strip = (f: Finding) => f.detail.replace(/\(ext\)/, '');
  for (const a of only('A')) {
    const b = only('B').find((x) => x.scenario === a.scenario && strip(x).replace('only in B', 'only in A') === strip(a));
    if (b) {
      a.cls = 'info-engine-dv-ext';
      b.cls = 'info-engine-dv-ext';
    }
  }
  return out;
}

export async function main(args: string[]): Promise<number> {
  return (await runOracle(args)).code;
}

function summary(recs: CopyRec[], findings: Finding[], runner: Record<string, RunnerRec>, t: { sheets: number; cells: number; formatCells: number; copies: number; ops: Record<string, number>; srcCompared: number }): string {
  const lines: string[] = [];
  const groups: Record<string, number> = {};
  for (const r of recs) groups[r.group] = (groups[r.group] ?? 0) + 1;
  lines.push(`combinations: ${new Set(recs.map((r) => r.combo)).size}; scenarios (random edit logs): ${recs.length} (${recs.filter((r) => r.skipped).length} skipped); by group ${JSON.stringify(groups)}`);
  lines.push(`single-sheet base (dropOtherSheets) ${recs.filter((r) => r.drop).length} / base's other sheets kept ${recs.filter((r) => !r.drop).length}; join full ${recs.filter((r) => /full/.test(r.combo)).length}; .xlsm output ${recs.filter((r) => r.ext === 'xlsm').length}`);
  lines.push(`copied sheets: ${t.copies} (compared with their sources in the premerged package: ${t.sheets} sheets, ~${t.cells} cells, ${t.formatCells} sampled cells rendered by Excel); imported formulas compared with the source (same person): ${t.srcCompared}`);
  lines.push(`accepted ops: ${JSON.stringify(t.ops)}`);
  lines.push(`with the Excel engine: ${Object.values(runner).filter((r) => r.ok).length} ok / ${Object.values(runner).filter((r) => !r.ok).length} failed`);
  const byCls = new Map<string, Finding[]>();
  for (const f of findings) byCls.set(f.cls, [...(byCls.get(f.cls) ?? []), f]);
  lines.push('', '| class | kind | findings | scenarios | example |', '|---|---|---|---|---|');
  for (const [cls, fs_] of [...byCls].sort()) lines.push(`| ${cls} | ${INFO_CLASSES.test(cls) ? 'info' : KNOWN_OPEN.test(cls) ? 'open (reported)' : '**FAIL**'} | ${fs_.length} | ${new Set(fs_.map((x) => x.scenario)).size} | ${fs_[0].scenario}: ${fs_[0].detail.replace(/\|/g, '/').slice(0, 260)} |`);
  if (!byCls.size) lines.push('| (none) | | 0 | 0 | |');
  return lines.join('\n') + '\n';
}
