// Mutation check of the consistency oracle: does it SEE a deliberate bug? A copy of the repository (node_modules linked) gets one small
// defect at a time in the code the oracle is meant to guard; the oracle runs on the copy and must report findings of the expected class.
//   node tests/oracle/consistency/mutate.mjs --work <dir> [--count 40] [--only M1,M3] [--runner]
// The original sources are never touched. --runner also starts the app's Excel engine (needed for defects that doc, replayLog and
// excelRef share: only real Excel can tell, M1 / M2).
import { spawnSync } from 'node:child_process';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const repo = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..', '..', '..');
const arg = (k, d = '') => {
  const i = process.argv.indexOf(`--${k}`);
  return i >= 0 && process.argv[i + 1] ? process.argv[i + 1] : d;
};
const work = path.resolve(arg('work', path.join(repo, '..', '.tmp', 'v1', 'mut')));
const count = arg('count', '40');
const only = arg('only') ? new Set(arg('only').split(',')) : null;
const withRunner = process.argv.includes('--runner');

/** id, file, [find, replace], expected finding classes (any of), needs real Excel */
const MUTATIONS = [
  { id: 'M0', what: 'control: the unmutated copy must report NO failing finding', file: null, edit: null, expect: null, excel: true },
  {
    id: 'M1', what: 'excelRef.moveSpan: a block inside a range whose end it moves inward keeps "follow" instead of the delete+insert model',
    file: 'src/core/formula/excelRef.ts',
    edit: ['if (blockInside && ((inB(lo) && movingRight) || (inB(hi) && !movingRight))) return modelA();', 'if (false) return modelA();'],
    expect: /engine-(formula|value)|value-changed/, excel: true,
  },
  {
    id: 'M2', what: 'excelRef.deleteSpan: a range whose end lies inside the deleted block keeps one column too many',
    file: 'src/core/formula/excelRef.ts',
    edit: ['const h = hi === max && !single ? max : hi > s2 ? hi - count : hi >= s1 ? s1 - 1 : hi;', 'const h = hi === max && !single ? max : hi > s2 ? hi - count : hi >= s1 ? s1 : hi;'],
    expect: /engine-(formula|value)|value-changed|doc-vs/, excel: true,
  },
  {
    id: 'M3', what: 'replay (patch/formula.ts): references qualified with a sheet name (Sheet!A1) are not rewritten when a column moves',
    file: 'src/core/xlsx/patch/formula.ts',
    edit: ['const sheetL = p ? lower(p.sheet) : hostL;', 'const sheetL = p ? null : hostL;'],
    expect: /value-changed|doc-vs-replay|ref-identity|engine-/, excel: false,
  },
  {
    id: 'M4', what: 'doc: defined names are not adjusted by column operations',
    file: 'src/core/doc/struct.ts',
    edit: ['function adjustNames(names: DefinedName[], ctx: FormulaCtx): DefinedName[] {', 'function adjustNames(names: DefinedName[], ctx: FormulaCtx): DefinedName[] {\n  if (ctx) return names;'],
    expect: /doc-vs-replay/, excel: false, noExcel: true,
  },
  {
    id: 'M5', what: 'doc: formulas of other sheets are not adjusted (grid and file would disagree)',
    file: 'src/core/doc/struct.ts',
    edit: ['const externalFormulas = adjustExternal(doc.externalFormulas, ctx, counter);', 'const externalFormulas = doc.externalFormulas;'],
    expect: /doc-vs-replay/, excel: false, noExcel: true,
  },
  {
    id: 'M6', what: 'premerge row mapping (formula/remap.ts): the merged row of every 7th source row is one too far down (rows rearranged by key point at the wrong person)',
    file: 'src/core/formula/remap.ts',
    edit: ['this.t.set(sr, mr);', 'this.t.set(sr, mr + (sr % 7 === 0 ? 1 : 0));'],
    expect: /src-value/, excel: false, only: '^(rich-f04_lookups|rich-f06_running|rich-f05_ranges|row-(shuf|rev))',
  },
];

function copyRepo(dest) {
  fs.rmSync(dest, { recursive: true, force: true });
  fs.mkdirSync(dest, { recursive: true });
  for (const f of ['package.json', 'tsconfig.json', 'vitest.config.ts']) fs.copyFileSync(path.join(repo, f), path.join(dest, f));
  for (const d of ['src', 'tests', 'tools', 'samples', 'server']) {
    fs.cpSync(path.join(repo, d), path.join(dest, d), { recursive: true, filter: (s) => !/[\\/](node_modules|test-results|\.demo-dist)([\\/]|$)/.test(s) });
  }
  fs.symlinkSync(path.join(repo, 'node_modules'), path.join(dest, 'node_modules'), 'junction');
}

const results = [];
for (const m of MUTATIONS) {
  if (only && !only.has(m.id)) continue;
  if (m.excel && !withRunner) { results.push({ id: m.id, what: m.what, skipped: 'needs --runner' }); continue; }
  const dest = path.join(work, 'repo');
  copyRepo(dest);
  const file = m.file ? path.join(dest, m.file) : '';
  let text = m.file ? fs.readFileSync(file, 'utf8') : '';
  if (m.edit) {
    if (!text.includes(m.edit[0])) { results.push({ id: m.id, what: m.what, error: `pattern not found in ${m.file}` }); continue; }
    text = text.split(m.edit[0]).join(m.edit[1]);
    fs.writeFileSync(file, text);
  }
  const out = path.join(work, `out-${m.id}`);
  fs.rmSync(out, { recursive: true, force: true });
  fs.mkdirSync(out, { recursive: true });
  // gen/ (authored workbooks) is shared with the main run when it exists
  const genSrc = path.join(repo, '..', '.tmp', 'v1', 'big', 'gen');
  if (fs.existsSync(genSrc)) fs.cpSync(genSrc, path.join(out, 'gen'), { recursive: true });
  const args = ['tests/oracle/consistency/run.mjs', '--out', out, '--count', count, '--seed-base', '501'];
  if (m.only) args.push('--only', m.only);
  if (!m.excel || !withRunner) args.push('--no-runner');
  if (m.noExcel) args.push('--no-excel');
  console.log(`[${m.id}] ${m.what}`);
  const r = spawnSync(process.execPath, args, { cwd: dest, stdio: ['ignore', 'pipe', 'pipe'], encoding: 'utf8', maxBuffer: 1 << 28, timeout: 90 * 60_000, env: { ...process.env, TEMP: path.join(repo, '..', '.tmp'), TMP: path.join(repo, '..', '.tmp') } });
  const fpath = path.join(out, 'findings.json');
  const findings = fs.existsSync(fpath) ? JSON.parse(fs.readFileSync(fpath, 'utf8')) : [];
  const INFO = /^(skipped|runner-fallback|convert-blank|cf-quirk|circular|pivot-value-changed|pivot-engine-package-(cols|widths)|no-dump|append-range-running)/; // append-range-running: a reported open defect of premerge (check.ts KNOWN_OPEN)
  const hit = m.expect ? findings.filter((f) => m.expect.test(f.cls)) : findings.filter((f) => !INFO.test(f.cls));
  const byCls = {};
  for (const f of findings) byCls[f.cls] = (byCls[f.cls] ?? 0) + 1;
  results.push({ id: m.id, what: m.what, ...(m.expect ? { detected: hit.length > 0 } : { cleanControl: hit.length === 0 }), expectedClassFindings: hit.length, scenarios: new Set(hit.map((f) => f.scenario)).size, classes: byCls, exit: r.status });
  console.log(`[${m.id}] ${m.expect ? (hit.length > 0 ? 'DETECTED' : 'NOT DETECTED') : hit.length === 0 ? 'CLEAN' : 'NOT CLEAN'}: ${JSON.stringify(byCls)}`);
}
fs.writeFileSync(path.join(work, 'mutation-results.json'), JSON.stringify(results, null, 1));
console.log(JSON.stringify(results, null, 1));
process.exitCode = results.some((x) => x.detected === false || x.cleanControl === false || x.error) ? 1 : 0;
