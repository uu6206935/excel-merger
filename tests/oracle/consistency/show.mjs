// Prints one scenario of an oracle run: its op trail (what the doc accepted / refused), the runner's answer and every finding.
//   node tests/oracle/consistency/show.mjs <out dir> <scenario id | regex> [--findings-only]
import fs from 'node:fs';
import path from 'node:path';

const [, , dir, pattern, ...flags] = process.argv;
if (!dir || !pattern) {
  console.error('usage: show.mjs <out dir> <scenario id | regex> [--findings-only]');
  process.exit(2);
}
const read = (f, d) => (fs.existsSync(path.join(dir, f)) ? JSON.parse(fs.readFileSync(path.join(dir, f), 'utf8')) : d);
const recs = read('scenarios.json', []);
const findings = read('findings.json', []);
const runner = read('runner.json', {});
const re = new RegExp(pattern);
for (const r of recs.filter((x) => re.test(x.id))) {
  console.log(`=== ${r.id}  mode=${r.mode} allowAffecting=${r.allowAffecting} affecting=${r.affecting} lookupWarned=${r.lookupWarned} ops=${r.logLength}`);
  if (!flags.includes('--findings-only')) {
    for (const t of r.trail) console.log(`  ${t.ok ? 'OK  ' : '--  '}${t.desc}${t.refused ? `   REFUSED: ${t.refused.slice(0, 90)}` : ''}${t.warnings.length ? `   [${t.warnings.map((w) => w.slice(0, 40)).join(' / ')}]` : ''}`);
    console.log(`  initial ids: ${r.initial.ids.join(' ')}`);
    console.log(`  final   ids: ${r.final.ids.join(' ')}`);
    const x = runner[r.id];
    if (x) console.log(`  runner: ${x.ok ? 'ok' : `FAILED ${x.code}: ${x.message}`} ${x.ms} ms${x.warnings?.length ? ` warnings: ${x.warnings.map((w) => `${w.code}: ${w.message}`.slice(0, 120)).join(' | ')}` : ''}`);
  }
  for (const f of findings.filter((y) => y.scenario === r.id)) console.log(`  FINDING ${f.cls}: ${f.detail}`);
}
