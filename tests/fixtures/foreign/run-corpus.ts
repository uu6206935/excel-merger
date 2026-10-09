// node --import tsx tests/fixtures/foreign/run-corpus.ts [corpusDir] [--out <dir>] [--seeds N] [--steps N] [--only <substring>] [--roles base,other] [--joins left,full]
// Runs every workbook of the corpus through the pipeline (harness.ts) and prints a table; with --out writes the premerged and
// the replayed outputs (for the Excel open check: tests/fixtures/foreign/excel-check.ps1).
import fs from 'node:fs';
import path from 'node:path';
import { CORPUS_DIR, PARTNER, load, runPipeline, type RunResult } from './harness';

const args = process.argv.slice(2);
const opt = (k: string, d?: string) => { const i = args.indexOf(k); return i >= 0 ? args[i + 1] : d; };
const dir = path.resolve(args.find((a, i) => !a.startsWith('--') && !(i > 0 && args[i - 1].startsWith('--'))) ?? CORPUS_DIR);
const outDir = opt('--out');
const seeds = Number(opt('--seeds', '3'));
const steps = Number(opt('--steps', '12'));
const onlyList = (opt('--only') ?? '').split(',').filter(Boolean);
const roles = (opt('--roles', 'base,other') as string).split(',') as Array<'base' | 'other'>;
const joins = (opt('--joins', 'left,full') as string).split(',') as Array<'left' | 'full'>;
if (outDir) fs.mkdirSync(outDir, { recursive: true });
/** per written output: the op trail (the Excel check tolerates #REF! after a delete that was NOT converted to values) */
const manifest: Record<string, { trail: string[]; unconvertedDelete: boolean }> = {};

const files = fs.readdirSync(dir).filter((f) => /\.(xlsx|xlsm)$/i.test(f) && f !== PARTNER && (!onlyList.length || onlyList.some((o) => f.includes(o)))).sort();
const partner = await load(dir, PARTNER, 'p');
let bugs = 0, refused = 0, ok = 0;
const t0 = Date.now();
for (const f of files) {
  let loaded;
  try {
    loaded = await load(dir, f, 'f');
  } catch (e) {
    console.log(`${f.padEnd(44)} READ FAILED: ${(e as Error).message.split('\n')[0]}`);
    refused++;
    continue;
  }
  const lines: string[] = [];
  for (const role of roles) for (const join of joins) for (let seed = 1; seed <= seeds; seed++) {
    const t = Date.now();
    const r: RunResult = await runPipeline(loaded, partner, { role, join, seed, steps, keep: !!outDir && seed === 1 });
    if (r.verdict === 'ok') ok++; else if (r.verdict === 'refused') refused++; else bugs++;
    if (r.verdict !== 'ok') lines.push(`   [${role}/${join}/seed${seed}] ${r.verdict.toUpperCase()} at ${r.stage}: ${(r.refusal ?? r.problems.join(' | ')).slice(0, 600)}\n        ops: ${(r.info.trail ?? []).join(' ; ').slice(0, 400)}`);
    if (outDir && seed === 1) {
      const base = `${f.replace(/\.xlsx$/, '')}__${role}-${join}`;
      if (r.premerged) fs.writeFileSync(path.join(outDir, `${base}.pre.xlsx`), r.premerged);
      if (r.output) fs.writeFileSync(path.join(outDir, `${base}.out.xlsx`), r.output);
      const trail = r.info.trail ?? [];
      manifest[`${base}.out.xlsx`] = { trail, unconvertedDelete: trail.some((t) => /^delete /.test(t) && !/\(convert\)/.test(t) && !/\[refused/.test(t)) };
    }
    if (Date.now() - t > 8000) lines.push(`   [${role}/${join}/seed${seed}] SLOW ${Date.now() - t} ms`);
  }
  console.log(`${f.padEnd(44)} region ${loaded.ks.region.sheet}!h${loaded.ks.region.headerRow} d${loaded.ks.region.firstDataRow}-${loaded.ks.region.lastDataRow} c${loaded.ks.region.firstCol}-${loaded.ks.region.lastCol}  warnings:${loaded.book.warnings.length}`);
  for (const l of lines) console.log(l);
}
if (outDir) fs.writeFileSync(path.join(outDir, '_manifest.json'), JSON.stringify(manifest, null, 1));
console.log(`\n${files.length} files, ${ok} ok, ${refused} clean refusals, ${bugs} BUGS, ${(Date.now() - t0) / 1000}s`);
process.exit(bugs ? 1 : 0);
