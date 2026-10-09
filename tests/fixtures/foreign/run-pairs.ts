// node --import tsx tests/fixtures/foreign/run-pairs.ts [--pairs N] [--seed S] [--steps N] [--out <dir>]
// Cross-producer pairs: workbook A (as base) merged with workbook B (as other), both from the corpus (the run-corpus.ts runs always use
// partner.xlsx as the second file). Style tables, themes, date systems, tables / filters, shared strings of two different writers meet in one
// premerged package; then the random op log, replayLog and the harness invariants (incl. the structural lint) as usual.
import fs from 'node:fs';
import path from 'node:path';
import { CORPUS_DIR, PARTNER, load, runPipeline, type Loaded } from './harness';
import { rng } from '../../doc/fuzzcore';

const args = process.argv.slice(2);
const opt = (k: string, d?: string) => { const i = args.indexOf(k); return i >= 0 ? args[i + 1] : d; };
const pairs = Number(opt('--pairs', '200'));
const seed = Number(opt('--seed', '1'));
const steps = Number(opt('--steps', '15'));
const outDir = opt('--out');
if (outDir) fs.mkdirSync(outDir, { recursive: true });

const names = fs.readdirSync(CORPUS_DIR).filter((f) => /\.(xlsx|xlsm)$/i.test(f) && !/hc-24|hc-22/.test(f)).sort();
const cache = new Map<string, Promise<Loaded | null>>();
const get = (n: string) => {
  let p = cache.get(n);
  if (!p) cache.set(n, (p = load(CORPUS_DIR, n, 'x').catch(() => null)));
  return p;
};
const r = rng(seed * 104729);
const manifest: Record<string, { trail: string[]; unconvertedDelete: boolean }> = {};
let ok = 0, refused = 0, bugs = 0;
const t0 = Date.now();
for (let k = 0; k < pairs; k++) {
  const a = names[Math.floor(r() * names.length)];
  let b = names[Math.floor(r() * names.length)];
  if (b === a) b = PARTNER;
  const [la, lb] = [await get(a), await get(b)];
  if (!la || !lb) continue;
  const join = r() < 0.5 ? 'left' : 'full';
  const res = await runPipeline(la, lb, { role: 'base', join, seed: k + 1, steps, keep: !!outDir });
  if (res.verdict === 'ok') ok++; else if (res.verdict === 'refused') refused++; else bugs++;
  if (res.verdict === 'bug') console.log(`[${k}] ${a} + ${b} (${join}) BUG at ${res.stage}: ${res.problems.join(' | ').slice(0, 700)}\n    ops: ${(res.info.trail ?? []).join(' ; ').slice(0, 300)}`);
  else if (res.verdict === 'refused' && !/py-26/.test(a + b)) console.log(`[${k}] ${a} + ${b} (${join}) refused at ${res.stage}: ${res.refusal?.slice(0, 200)}`);
  if (outDir && res.premerged) {
    const base = `${a.replace(/\.xlsx?m?$/, '')}+${b.replace(/\.xlsx?m?$/, '')}-${join}`;
    fs.writeFileSync(path.join(outDir, `${base}.pre.xlsx`), res.premerged);
    if (res.output) fs.writeFileSync(path.join(outDir, `${base}.out.xlsx`), res.output);
    const trail = res.info.trail ?? [];
    manifest[`${base}.out.xlsx`] = { trail, unconvertedDelete: trail.some((t) => /^delete /.test(t) && !/\(convert\)/.test(t) && !/\[refused/.test(t)) };
  }
}
if (outDir) fs.writeFileSync(path.join(outDir, '_manifest.json'), JSON.stringify(manifest, null, 1));
console.log(`${pairs} pairs: ${ok} ok, ${refused} refused, ${bugs} BUGS, ${(Date.now() - t0) / 1000}s`);
process.exit(bugs ? 1 : 0);
