// Markdown tables for RESULTS.md from an oracle output folder: node tests/oracle/consistency/report.mjs <out dir>
import fs from 'node:fs';
import path from 'node:path';

const dir = process.argv[2];
if (!dir) {
  console.error('usage: report.mjs <out dir>');
  process.exit(2);
}
const read = (f, d) => (fs.existsSync(path.join(dir, f)) ? JSON.parse(fs.readFileSync(path.join(dir, f), 'utf8')) : d);
const recs = read('scenarios.json', []);
const per = new Map(read('results.json', []).map((p) => [p.id, p]));
const findings = read('findings.json', []);
const runner = read('runner.json', {});
const src = read('srcstats.json', null);

const family = (r) => (r.set.startsWith('smp-') ? 'samples (smp-*)' : r.set.startsWith('row') ? 'rows rearranged (row-*)' : r.set.includes('table') || r.set.includes('title') || r.set.includes('quoted') ? 'rich table / title / quoted-name base' : r.set.includes('multi') ? 'rich base + 3 sources' : 'rich base + F1 formula source (rich-*)');
const rows = new Map();
for (const r of recs) {
  const k = family(r);
  const p = per.get(r.id);
  const e = rows.get(k) ?? { n: 0, cells: 0, fcells: 0, mis: 0, ex: {}, ops: 0, fin: 0 };
  e.n++;
  e.ops += r.logLength;
  if (p) {
    e.cells += p.stats.cells;
    e.fcells += p.stats.formulaCells;
    e.mis += p.stats.mismatches;
    for (const [a, b] of Object.entries(p.explained)) e.ex[a] = (e.ex[a] ?? 0) + b;
  }
  rows.set(k, e);
}
console.log('| family | scenarios | ops in the logs | cells compared | formula cells | changed vs baseline (all explained) | explained by |');
console.log('|---|---|---|---|---|---|---|');
for (const [k, e] of [...rows].sort()) console.log(`| ${k} | ${e.n} | ${e.ops} | ${e.cells.toLocaleString('en')} | ${e.fcells.toLocaleString('en')} | ${e.mis.toLocaleString('en')} | ${Object.entries(e.ex).map(([a, b]) => `${a} ${b}`).join(', ')} |`);
const tot = [...rows.values()].reduce((a, e) => ({ n: a.n + e.n, ops: a.ops + e.ops, cells: a.cells + e.cells, fcells: a.fcells + e.fcells, mis: a.mis + e.mis }), { n: 0, ops: 0, cells: 0, fcells: 0, mis: 0 });
console.log(`| **total** | ${tot.n} | ${tot.ops} | ${tot.cells.toLocaleString('en')} | ${tot.fcells.toLocaleString('en')} | ${tot.mis.toLocaleString('en')} | |`);
const rk = Object.values(runner);
console.log(`\nExcel engine (runner): ${rk.filter((x) => x.ok).length} ok, ${rk.filter((x) => !x.ok).length} not run (${[...new Set(rk.filter((x) => !x.ok).map((x) => x.code))].join(', ') || '-'}); median ${Math.round(rk.map((x) => x.ms).sort((a, b) => a - b)[Math.floor(rk.length / 2)] ?? 0)} ms per job`);
if (src) console.log(`row mapping: ${src.checks} (set, variant) pairs, ${src.keptCompared} imported formula cells compared with the source file (${src.widenedCompared} against a reduced source), ${src.skipped} skipped by rule, ${src.personRows} placed rows checked for the person`);
const byCls = new Map();
for (const f of findings) byCls.set(f.cls, [...(byCls.get(f.cls) ?? []), f]);
console.log('\n| finding class | findings | scenarios |');
console.log('|---|---|---|');
for (const [c, l] of [...byCls].sort()) console.log(`| ${c} | ${l.length} | ${new Set(l.map((x) => x.scenario)).size} |`);
