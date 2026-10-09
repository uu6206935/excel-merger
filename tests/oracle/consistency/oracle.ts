/**
 * Consistency oracle in real Excel (task V1): formulas must stay correct when columns are moved / inserted / deleted, in BOTH export
 * engines, on top of merges whose rows were rearranged by key matching. Re-runnable:
 *
 *   powershell -NoProfile -ExecutionPolicy Bypass -File tests\oracle\consistency\run.ps1 [-Count 320] [-Out <dir>] [-Only <regex>] [-NoRunner] [-NoExcel] [-Reuse]
 *   node tests/oracle/consistency/run.mjs --out <dir> --count 320 [--seed-base 1] [--only <regex>] [--no-runner] [--no-excel] [--reuse]
 *
 * Method (see RESULTS.md): for every scenario (set x variant x seed) the app's pipeline builds the premerged package (readBook -> keys ->
 * joinSources -> buildPremerged -> createEditor); a random op log is made through the doc API (ops.ts); the BASELINE is the premerged
 * package with the typed values / converted values but with NO layout operation; the log is exported by replayLog (built-in engine) and by
 * the real Excel through excel-runner.ps1; ONE hidden Excel per batch opens baseline + both outputs, recalculates fully (CalculateFull),
 * and dumps the values + formulas of every sheet; check.ts compares cell by cell, found again by (sheet, row, column id).
 */
import fs from 'node:fs';
import path from 'node:path';
import { compare } from '../../doc/compare';
import { rng } from '../../doc/fuzzcore';
import { readBook, replayLog } from '../../../src/core/xlsx';
import type { MergedDoc } from '../../../src/core/types';
import { INFO_CLASSES, KNOWN_OPEN, judgeConverted, judgeOutput, judgeParity, judgeRefIdentity, type DocInfo, type Finding, type ScenarioRec, type Stats } from './check';
import { ensureDisk, readDump, runBatch, runRunnerJob, type BatchJob } from './excel';
import { generatedWorkbooks, buildSets, type Built, type SetDef, type Variant } from './sets';
import { baselineLog, docIdentity, generateOps } from './ops';
import { checkAppendedRows } from './appended';
import { packageFindings } from './pkgdiff';
import { buildSrcChecks, judgeSrc, type SrcCheck, type SrcStats } from './srcCheck';
import { DEFAULT_PARAMS, REPO, mergeAll, pivotSheetNames, type Merged } from './lib';

function argv(args: string[], k: string, d = ''): string {
  const i = args.indexOf(`--${k}`);
  return i >= 0 && args[i + 1] ? args[i + 1] : d;
}

const hash = (s: string): number => [...s].reduce((h, ch) => (Math.imul(h, 31) + ch.charCodeAt(0)) >>> 0, 7);
const safe = (s: string) => s.replace(/[^A-Za-z0-9_.-]/g, '_');
const refRuleCount = (d: MergedDoc): number => d.features.rules.filter((r) => r.formulas.some((f) => /#REF!/.test(f))).length + d.externalRules.filter((e) => e.rule.formulas.some((f) => /#REF!/.test(f))).length;
const docInfo = (d: MergedDoc): DocInfo => ({ names: d.names, tables: d.features.tables, ids: d.columns.map((c) => c.id) });

interface Plan {
  set: SetDef;
  variant: Variant;
  seed: number;
}

/** weighted round robin over (set, variant) pairs until `count` scenarios are planned */
function plan(sets: SetDef[], count: number, seedBase: number, only: RegExp | null): Plan[] {
  const combos: Array<{ set: SetDef; variant: Variant; w: number }> = [];
  for (const s of sets) {
    if (only && !only.test(s.id)) continue;
    for (const v of s.variants) combos.push({ set: s, variant: v, w: s.weight ?? 1 });
  }
  // a stable pseudo-random order: any prefix of the plan samples all kinds of sets
  combos.sort((a, b) => hash(`${a.set.id}/${a.variant.name}`) - hash(`${b.set.id}/${b.variant.name}`));
  const out: Plan[] = [];
  const credit = new Map<number, number>();
  const used = new Map<number, number>();
  let guard = 0;
  while (out.length < count && combos.length && guard++ < count * 50) {
    for (let i = 0; i < combos.length && out.length < count; i++) {
      const c = (credit.get(i) ?? 0) + combos[i].w;
      if (c >= 1) {
        credit.set(i, c - 1);
        const k = used.get(i) ?? 0;
        used.set(i, k + 1);
        out.push({ set: combos[i].set, variant: combos[i].variant, seed: seedBase + k });
      } else credit.set(i, c);
    }
  }
  return out;
}

export interface RunnerRec {
  ok: boolean;
  code?: string;
  message?: string;
  warnings: Array<{ code: string; message: string }>;
  ms: number;
  refErrors?: { cells: number; names: number; samples: string[] };
}

export async function runOracle(args: string[]): Promise<{ code: number; recs: ScenarioRec[]; findings: Finding[]; summary: string }> {
  const out = path.resolve(argv(args, 'out', path.join(REPO, '..', '.tmp', 'v1', 'oracle')));
  const count = Number(argv(args, 'count', '320'));
  const seedBase = Number(argv(args, 'seed-base', '1'));
  const only = argv(args, 'only') ? new RegExp(argv(args, 'only')) : null;
  const noExcel = args.includes('--no-excel');
  const noRunner = args.includes('--no-runner') || noExcel;
  const reuse = args.includes('--reuse');
  const log = (s: string) => console.log(`[${new Date().toISOString().slice(11, 19)}] ${s}`);
  const tempRoot = path.join(out, 'xltmp');
  const dumpDir = path.join(out, 'dump');
  const pkgDir = path.join(out, 'pkg');
  const genDir = path.join(out, 'gen');
  for (const d of [out, dumpDir, pkgDir, path.join(genDir, 'raw'), path.join(genDir, 'cached')]) fs.mkdirSync(d, { recursive: true });
  if (!noExcel) ensureDisk(out);

  // ---- A. the generated workbooks (raw, then authored by Excel so that results are stored)
  const gen = await generatedWorkbooks();
  const authorJobs: BatchJob[] = [];
  for (const [name, bytes] of gen) {
    const raw = path.join(genDir, 'raw', `${name}.xlsx`);
    if (!fs.existsSync(raw) || fs.readFileSync(raw).length !== bytes.length) fs.writeFileSync(raw, bytes);
    const cached = path.join(genDir, 'cached', `${name}.xlsx`);
    if (!fs.existsSync(cached)) authorJobs.push({ id: `author-${name}`, file: raw, kind: 'resave', saveAs: cached });
  }
  if (authorJobs.length && !noExcel) {
    log(`authoring ${authorJobs.length} generated workbooks in Excel (results stored)`);
    // a stale marker of an earlier run must not skip the job
    for (const j of authorJobs) fs.rmSync(path.join(dumpDir, `${j.id}.values.json`), { force: true });
    await runBatch(authorJobs, dumpDir, { tempRoot, batchSize: 40, log });
  }

  // ---- B. scenarios
  const recsFile = path.join(out, 'scenarios.json');
  const srcFile = path.join(out, 'srcchecks.json');
  let recs: ScenarioRec[];
  let srcChecks: SrcCheck[] = [];
  /** findings that need no Excel (checked while the premerged sheets are built) */
  let staticFindings: Finding[] = [];
  const staticFile = path.join(out, 'static.json');
  if (reuse && fs.existsSync(recsFile)) {
    recs = JSON.parse(fs.readFileSync(recsFile, 'utf8')) as ScenarioRec[];
    for (const r of recs) r.pivotComparable ??= !r.opsKinds.edit && !r.converted.length && !r.deleted.length;
    if (fs.existsSync(staticFile)) staticFindings = JSON.parse(fs.readFileSync(staticFile, 'utf8')) as Finding[];
    if (fs.existsSync(srcFile)) srcChecks = JSON.parse(fs.readFileSync(srcFile, 'utf8')) as SrcCheck[];
    if (only) { recs = recs.filter((r) => only.test(r.set)); srcChecks = srcChecks.filter((c) => only.test(c.combo)); }
    log(`reusing ${recs.length} scenarios from ${recsFile}`);
  } else {
    recs = [];
    const sets = buildSets();
    const plans = plan(sets, count, seedBase, only);
    log(`${plans.length} scenarios planned over ${new Set(plans.map((p) => `${p.set.id}/${p.variant.name}`)).size} (set, variant) pairs`);
    const cache = new Map<string, { merged: Merged; preFile: string; built: Built } | null>();
    const pivots = new Map<string, string[]>();
    staticFindings = [];
    const comboKey = (p: Plan) => `${p.set.id}__${p.variant.name}`;
    let skipped = 0;
    for (const p of plans) {
      const ck = comboKey(p);
      let c = cache.get(ck);
      if (c === undefined) {
        try {
          const built = await p.set.build({ genDir });
          const merged = await mergeAll(built.base, built.others, { ...DEFAULT_PARAMS, ...p.variant.params });
          const preFile = path.join(pkgDir, `${safe(ck)}.pre.xlsx`);
          fs.writeFileSync(preFile, merged.result.bytes);
          pivots.set(ck, await pivotSheetNames(merged.result.bytes));
          srcChecks.push(...buildSrcChecks(ck, merged, pkgDir, p.set.strictKeys === true));
          staticFindings.push(...checkAppendedRows(ck, merged));
          c = { merged, preFile, built };
        } catch (e) {
          log(`SETUP FAILED ${ck}: ${(e as Error).message}`);
          c = null;
        }
        cache.set(ck, c);
      }
      if (!c) { skipped++; continue; }
      const { merged, preFile } = c;
      const r = rng(p.seed * 7919 + hash(ck));
      const mode: 'inv' | 'del' = r() < 0.62 ? 'inv' : 'del';
      const allowAffecting = r() < 0.4;
      const steps = 3 + Math.floor(r() * 7);
      const g = generateOps(merged.editor, r, { steps, mode, allowAffecting });
      if (!g.final.log.length) { skipped++; continue; }
      const id = `${safe(ck)}__s${p.seed}`;
      const files = { pre: path.basename(preFile), base: path.basename(preFile), js: `${id}.js.xlsx`, xl: `${id}.xl.xlsx` };
      const rec: ScenarioRec = {
        id, set: p.set.id, variant: p.variant.name, seed: p.seed, mode, allowAffecting,
        mergedSheet: merged.editor.doc.sheetName,
        headerRow: merged.editor.doc.headerRow, firstDataRow: merged.editor.doc.firstDataRow, lastDataRow: merged.editor.doc.lastDataRow,
        initial: docInfo(merged.editor.doc), final: docInfo(g.final.doc), deleted: g.deleted,
        converted: g.converted.map((x) => ({ sheet: x.sheet, id: x.id, row: x.row, col: x.col, value: x.write.v, blank: x.write.t === 'z' && x.write.v === null })),
        trail: g.trail, lookupWarned: g.lookupWarned, affecting: g.affecting, opsKinds: g.opsKinds, logLength: g.final.log.length,
        baselineIsPre: true, refRules: { initial: refRuleCount(merged.editor.doc), final: refRuleCount(g.final.doc) }, files, pivotSheets: pivots.get(ck) ?? [], pivotComparable: !g.edits.length && !g.converted.length && !g.deleted.length, docVsReplay: [],
      };
      // baseline = premerged + typed values + converted values, no layout operation
      const bl = baselineLog(merged.editor, g);
      let baseBytes: Uint8Array | null = null;
      if (bl === null) {
        rec.skipped = 'the doc refused a typed text in the original layout';
      } else if (bl.log.length) {
        baseBytes = await replayLog(merged.result.bytes, bl.log);
        rec.baselineIsPre = false;
        rec.files.base = `${id}.base.xlsx`;
        fs.writeFileSync(path.join(pkgDir, rec.files.base), baseBytes);
      }
      // the built-in engine
      let jsBytes: Uint8Array | null = null;
      try {
        jsBytes = await replayLog(merged.result.bytes, g.final.log);
        fs.writeFileSync(path.join(pkgDir, rec.files.js), jsBytes);
        const book = await readBook(jsBytes, 'js.xlsx', { id: 'js' });
        rec.docVsReplay = compare(g.final.doc, book).slice(0, 12);
      } catch (e) {
        rec.docVsReplay = [`replayLog threw: ${(e as Error).message}`];
        rec.skipped = rec.skipped ?? 'replayLog threw';
      }
      // the doc alone: references to one cell / one column follow their column ids
      const ident = docIdentity(merged.editor.doc, g.final.doc, new Set(g.deleted), new Set(g.edits.map((e) => `${e.colId}:${e.row}`)));
      if (ident.length) rec.docVsReplay.push(...ident.slice(0, 4).map((x) => `ref-identity(doc): ${x}`));
      fs.writeFileSync(path.join(pkgDir, `${id}.log.json`), JSON.stringify(g.final.log));
      recs.push(rec);
    }
    fs.writeFileSync(recsFile, JSON.stringify(recs, null, 1));
    fs.writeFileSync(srcFile, JSON.stringify(srcChecks, null, 1));
    fs.writeFileSync(staticFile, JSON.stringify(staticFindings, null, 1));
    log(`${recs.length} scenarios built, ${skipped} skipped (no accepted op / set-up failure)`);
  }

  // ---- C. the Excel engine (one Excel start per scenario: that is how the runner works)
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
      const res = await runRunnerJob(path.join(pkgDir, rec.files.pre), logOps, xl, rec.mergedSheet, { tempRoot, cacheDir });
      runner[rec.id] = res;
      if (!res.ok) fs.rmSync(xl, { force: true });
      fs.writeFileSync(runnerFile, JSON.stringify(runner, null, 1));
      if (++done % 10 === 0) log(`runner: ${done} jobs done (last ${res.ms} ms${res.ok ? '' : ` FAILED ${res.code}`})`);
    }
    log(`runner: ${done} new jobs, ${Object.values(runner).filter((r) => !r.ok).length} failed in total`);
  }

  // ---- D. Excel: baseline + both outputs of every scenario, one hidden Excel per batch
  const jobs = new Map<string, BatchJob>();
  const addJob = (file: string) => {
    const id = safe(path.basename(file).replace(/\.xlsx$/, ''));
    if (!jobs.has(id) && fs.existsSync(path.join(pkgDir, file))) jobs.set(id, { id, file: path.join(pkgDir, file), kind: 'dump' });
    return id;
  };
  for (const rec of recs) {
    if (rec.skipped) continue;
    addJob(rec.files.base);
    if (rec.converted.length) addJob(rec.files.pre);
    addJob(rec.files.js);
    if (runner[rec.id]?.ok) addJob(rec.files.xl);
  }
  for (const c of srcChecks) {
    if (!c.cells.length && !c.placed.length) continue;
    const preFile = `${safe(c.combo)}.pre.xlsx`;
    addJob(preFile);
    const twin = `${safe(c.combo)}.src-${c.sourceId}.xlsx`;
    if (fs.existsSync(path.join(pkgDir, twin))) jobs.set(c.twinJob, { id: c.twinJob, file: path.join(pkgDir, twin), kind: 'dump' });
    if (c.reducedJob) jobs.set(c.reducedJob, { id: c.reducedJob, file: path.join(pkgDir, twin), kind: 'reduced', sheet: c.sourceSheet, deleteRows: c.deleteRows });
  }
  if (!noExcel) {
    log(`excel: ${jobs.size} workbooks to open and recalculate`);
    await runBatch([...jobs.values()], dumpDir, { tempRoot, batchSize: 40, log });
  }

  // ---- E. judge
  const findings: Finding[] = [...staticFindings];
  const convertTotals = { blank: 0, stale: 0, ok: 0 };
  const perScenario: Array<{ id: string; set: string; variant: string; mode: string; ops: number; stats: Stats; findings: number; explained: Record<string, number> }> = [];
  const idOf = (file: string) => safe(file.replace(/\.xlsx$/, ''));
  for (const rec of recs) {
    if (rec.skipped) { findings.push({ scenario: rec.id, cls: 'skipped', detail: rec.skipped }); continue; }
    const f0 = findings.length;
    const base = readDump(dumpDir, idOf(rec.files.base));
    const js = readDump(dumpDir, idOf(rec.files.js));
    const xlRun = runner[rec.id];
    const xl = xlRun?.ok ? readDump(dumpDir, idOf(rec.files.xl)) : null;
    if (rec.docVsReplay.length) findings.push({ scenario: rec.id, cls: 'doc-vs-replay', detail: rec.docVsReplay.slice(0, 4).join(' | ') });
    if (xlRun && !xlRun.ok) {
      const fallback = /MOVE_VERIFY_FAILED|EXCEL_UNAVAILABLE|PASSWORD_PROTECTED|OPEN_FAILED|RUNNER_MISSING/.test(xlRun.code ?? '');
      findings.push({ scenario: rec.id, cls: fallback ? 'runner-fallback' : 'runner-failed', detail: `${xlRun.code}: ${xlRun.message}` });
    }
    // a name / CF / DV rule that ends up with #REF! (a deleted column nothing could be converted for) must have been announced by the doc
    const newRefNames = rec.final.names.filter((n) => /#REF!/.test(n.refersTo) && !rec.initial.names.some((m) => m.name === n.name && m.scope === n.scope && /#REF!/.test(m.refersTo)));
    const newRefRules = (rec.refRules?.final ?? 0) > (rec.refRules?.initial ?? 0);
    if (newRefNames.length || newRefRules) {
      const said = rec.trail.some((t) => t.warnings.some((w) => /#REF!/.test(w)));
      findings.push({ scenario: rec.id, cls: said ? 'ref-error-announced' : 'ref-error-silent', detail: `${newRefNames.map((n) => `name ${n.name} = ${n.refersTo}`).join('; ')}${newRefRules ? ` ${(rec.refRules?.final ?? 0) - (rec.refRules?.initial ?? 0)} CF/DV rules` : ''} now hold #REF!${said ? ' (the doc warned)' : ' and the doc said nothing'}` });
    }
    const stats: Stats = { cells: 0, formulaCells: 0, mismatches: 0, explained: {} };
    if (base && js) {
      const pre = rec.converted.length ? readDump(dumpDir, idOf(rec.files.pre)) ?? undefined : undefined;
      const rj = judgeOutput({ rec, base, out: js, engine: 'js', pre });
      findings.push(...rj.findings);
      Object.assign(stats, rj.stats);
      findings.push(...judgeRefIdentity(rec, base, js, 'js'));
      if (pre && rec.converted.length) {
        const jc = judgeConverted(rec, pre);
        findings.push(...jc.findings);
        if (jc.blank && !rec.trail.some((t) => t.warnings.some((w) => /計算結果が保存されていなかった/.test(w)))) findings.push({ scenario: rec.id, cls: 'convert-blank-silent', detail: `${jc.blank} converted formulas were blanked but the doc gave no warning` });
        convertTotals.blank += jc.blank;
        convertTotals.stale += jc.stale;
        convertTotals.ok += jc.ok;
      }
      if (xl) {
        const rx = judgeOutput({ rec, base, out: xl, engine: 'xl', pre });
        findings.push(...rx.findings);
        findings.push(...judgeRefIdentity(rec, base, xl, 'xl'));
        findings.push(...judgeParity(rec, js, xl));
        try {
          findings.push(...(await packageFindings(rec, new Uint8Array(fs.readFileSync(path.join(pkgDir, rec.files.js))), new Uint8Array(fs.readFileSync(path.join(pkgDir, rec.files.xl))))));
        } catch (e) {
          findings.push({ scenario: rec.id, cls: 'engine-package-error', detail: (e as Error).message });
        }
      }
    } else {
      findings.push({ scenario: rec.id, cls: 'no-dump', detail: `base ${base ? 'ok' : 'MISSING'}, js ${js ? 'ok' : 'MISSING'}` });
    }
    if (!rec.allowAffecting && (stats.explained['range-span'] ?? 0) > 0) findings.push({ scenario: rec.id, cls: 'oracle-inconsistent', detail: `a strict scenario (no op may change a range's column span) shows ${stats.explained['range-span']} range-span changes: ops.ts membership() and check.ts spanChanged() disagree` });
    perScenario.push({ id: rec.id, set: rec.set, variant: rec.variant, mode: rec.mode, ops: rec.logLength, stats, findings: findings.length - f0, explained: stats.explained });
    if (process.env.ORACLE_VERBOSE && stats.examples) for (const [k, v] of Object.entries(stats.examples)) for (const e of v) console.log(`  [${rec.id}] explained ${k}: ${e}`);
  }
  const srcStats: SrcStats & { checks: number } = { keptCompared: 0, widenedCompared: 0, skipped: 0, personRows: 0, checks: 0 };
  for (const c of srcChecks) {
    if (!c.cells.length && !c.placed.length) continue;
    const pre = readDump(dumpDir, idOf(`${safe(c.combo)}.pre.xlsx`));
    const twin = readDump(dumpDir, c.twinJob);
    const red = c.reducedJob ? readDump(dumpDir, c.reducedJob) : null;
    if (!pre || !twin) { findings.push({ scenario: c.combo, cls: 'no-dump', detail: `source check: pre ${pre ? 'ok' : 'MISSING'}, source ${twin ? 'ok' : 'MISSING'}` }); continue; }
    const r = judgeSrc(c, pre, twin, red);
    findings.push(...r.findings);
    srcStats.checks++;
    srcStats.keptCompared += r.stats.keptCompared;
    srcStats.widenedCompared += r.stats.widenedCompared;
    srcStats.skipped += r.stats.skipped;
    srcStats.personRows += r.stats.personRows;
  }
  fs.writeFileSync(path.join(out, 'srcstats.json'), JSON.stringify(srcStats));
  fs.writeFileSync(path.join(out, 'findings.json'), JSON.stringify(findings, null, 1));
  fs.writeFileSync(path.join(out, 'results.json'), JSON.stringify(perScenario, null, 1));
  const md = summary(recs, findings, perScenario, runner, srcStats, convertTotals);
  fs.writeFileSync(path.join(out, 'results.md'), md);
  console.log(md);
  const failing = findings.filter((f) => !INFO_CLASSES.test(f.cls) && !KNOWN_OPEN.test(f.cls));
  return { code: failing.length ? 1 : 0, recs, findings, summary: md };
}

export async function main(args: string[]): Promise<number> {
  return (await runOracle(args)).code;
}

function summary(recs: ScenarioRec[], findings: Finding[], per: Array<{ stats: Stats; explained: Record<string, number> }>, runner: Record<string, RunnerRec>, src: SrcStats & { checks: number }, conv: { blank: number; stale: number; ok: number }): string {
  const lines: string[] = [];
  const sum = (f: (x: { stats: Stats }) => number) => per.reduce((a, x) => a + f(x), 0);
  const ops: Record<string, number> = {};
  for (const r of recs) for (const [k, v] of Object.entries(r.opsKinds)) ops[k] = (ops[k] ?? 0) + v;
  const fam = (re: RegExp) => recs.filter((r) => re.test(r.set)).length;
  lines.push(`families: samples ${fam(/^smp-/)} · rich base + F1 sources ${fam(/^rich-/)} · rows rearranged ${fam(/^row/)}; modes: inv ${recs.filter((r) => r.mode === 'inv').length} / del ${recs.filter((r) => r.mode === 'del').length}; strict (no op may change a range's column span) ${recs.filter((r) => !r.allowAffecting).length} / loose ${recs.filter((r) => r.allowAffecting).length}; join full ${recs.filter((r) => /full/.test(r.variant)).length}, single-sheet output ${recs.filter((r) => /drop/.test(r.variant)).length}`);
  lines.push(`accepted ops: ${JSON.stringify(ops)}; refused by the doc ${recs.reduce((a, r) => a + r.trail.filter((t) => !t.ok && !/^\(not taken/.test(t.refused ?? '')).length, 0)}; not taken by the generator (span / loop guard) ${recs.reduce((a, r) => a + r.trail.filter((t) => !t.ok && /^\(not taken/.test(t.refused ?? '')).length, 0)}`);
  lines.push(`scenarios: ${recs.length} (${recs.filter((r) => r.skipped).length} skipped), with the Excel engine: ${Object.values(runner).filter((r) => r.ok).length} ok / ${Object.values(runner).filter((r) => !r.ok).length} failed`);
  lines.push(`cells compared: ${sum((x) => x.stats.cells)} (formula cells ${sum((x) => x.stats.formulaCells)}), mismatches vs baseline ${sum((x) => x.stats.mismatches)} (JS engine)`);
  lines.push(`row mapping (imported formulas vs the source file, same person): ${src.checks} (set, variant) pairs, ${src.keptCompared} formula cells compared (${src.widenedCompared} against a reduced source), ${src.skipped} skipped by rule, ${src.personRows} placed rows checked for the person`);
  lines.push(`formulas converted to values by deleteColumns(convertDependents): ${conv.ok} equal to Excel's value, ${conv.blank} blank where Excel has a value (no stored result, the doc warns), ${conv.stale} other value`);
  const ex: Record<string, number> = {};
  for (const p of per) for (const [k, v] of Object.entries(p.explained)) ex[k] = (ex[k] ?? 0) + v;
  lines.push(`explained changes: ${JSON.stringify(ex)}`);
  const byCls = new Map<string, Finding[]>();
  for (const f of findings) byCls.set(f.cls, [...(byCls.get(f.cls) ?? []), f]);
  lines.push('', '| class | kind | findings | scenarios | example |', '|---|---|---|---|---|');
  for (const [cls, fs_] of [...byCls].sort()) lines.push(`| ${cls} | ${INFO_CLASSES.test(cls) ? 'info' : KNOWN_OPEN.test(cls) ? 'open (reported)' : '**FAIL**'} | ${fs_.length} | ${new Set(fs_.map((x) => x.scenario)).size} | ${fs_[0].scenario}: ${fs_[0].detail.replace(/\|/g, '/').slice(0, 220)} |`);
  if (!byCls.size) lines.push('| (none) | | 0 | 0 | |');
  return lines.join('\n') + '\n';
}
