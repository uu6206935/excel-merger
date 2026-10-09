// Writes a seeded random scenario list for excel-cf-probe.ps1 (tests/formula/cfRandom.test.ts): one or two rules
// (one <conditionalFormatting> block, `=5` constant or `=ref>0` / `=ref>ref` with relative / absolute-column references)
// on one or two areas of a 2-row band (columns C..), then one or more column inserts / deletes / moves (Excel runner's
// clipboard-free sequence for moves). The fixtures used by the tests were made with
//   node make-cf-random-spec.mjs p6 424242 240 1 1 <out>    (single rules, one op)
//   node make-cf-random-spec.mjs p7 777 360 2 1 <out>       (blocks of up to 2 rules, one op)
//   node make-cf-random-spec.mjs p8 31337 300 2 3 <out>     (blocks of up to 2 rules, 1 to 3 successive ops)
// Usage: node make-cf-random-spec.mjs <id prefix> <seed> <count> <max rules per block 1|2> <max ops 1..3> <out spec.json>
import fs from 'node:fs';

const [prefix, seedArg, countArg, maxRulesArg, maxOpsArg, dest] = process.argv.slice(2);
if (!dest) { console.error('usage: node make-cf-random-spec.mjs <prefix> <seed> <count> <max rules 1|2> <max ops 1..3> <out spec.json>'); process.exit(1); }
const numToCol = (n) => { let s = ''; for (; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s; return s; };
let seed = Number(seedArg);
const rnd = (n) => { seed = (seed * 1103515245 + 12345) & 0x7fffffff; return seed % n; };
const pick = (xs) => xs[rnd(xs.length)];
const out = [];
for (let k = 0; k < Number(countArg); k++) {
  const nAreas = rnd(4) === 0 ? 2 : 1;
  const areas = [];
  let c = 3 + rnd(2);
  for (let i = 0; i < nAreas; i++) { const w = 1 + rnd(4); areas.push([c, c + w - 1]); c += w + 1 + rnd(2); }
  const lo = areas[0][0];
  const addr = areas.map(([a, b]) => `${numToCol(a)}2:${numToCol(b)}3`).join(',');
  const relRef = () => `${numToCol(Math.max(1, lo + pick([-2, -1, 0, 1, 2, 3, 4, 5])))}2`;
  const absRef = () => `$${numToCol(1 + rnd(13))}2`;
  const ref = () => (rnd(3) === 0 ? absRef() : relRef());
  const mk = (kind) => kind === 0 ? { type: 'cellIs', addr, active: `${numToCol(lo)}2`, f: '=5' }
    : kind <= 6 ? { type: 'expr', addr, active: `${numToCol(lo)}2`, f: `=${ref()}>0` }
      : { type: 'expr', addr, active: `${numToCol(lo)}2`, f: `=${ref()}>${ref()}` };
  const rules = Number(maxRulesArg) >= 2 && rnd(3) === 0 ? [mk(rnd(10)), mk(1 + rnd(9))] : [mk(rnd(10))];
  const mkOp = () => {
    const t = rnd(3);
    return t === 0 ? { kind: 'ins', at: 2 + rnd(11), count: 1 + rnd(2) }
      : t === 1 ? { kind: 'del', at: 2 + rnd(11), count: 1 + rnd(2) }
        : (() => { const count = 1 + rnd(3), from = 2 + rnd(9); let to = 2 + rnd(12); if (to >= from && to <= from + count) to = from + count + 1 + rnd(2); return { kind: 'mov', from, count, to }; })();
  };
  const nOps = Number(maxOpsArg) > 1 ? 1 + rnd(Number(maxOpsArg)) : 1;
  out.push({ id: `${prefix}_${String(k).padStart(3, '0')}`, rules, ops: Array.from({ length: nOps }, mkOp) });
}
fs.writeFileSync(dest, JSON.stringify({ scenarios: out }, null, 1));
console.log(`wrote ${dest} (${out.length} scenarios)`);
