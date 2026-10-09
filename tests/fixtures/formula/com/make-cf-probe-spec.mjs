// Writes the scenario list of excel-cf-probe.ps1 for the "cf probe" families used by tests/formula/cfProbe.test.ts.
// Every scenario = one workbook with conditional-format rules (one <conditionalFormatting> block per distinct range) and
// ONE column move (Excel runner's clipboard-free sequence: insert blank columns at `to`, Range.Cut(Destination), delete
// the emptied source). Families (col letters; every area is 2 rows high):
//   p1  areas B:E / C:G, expression rule `=<col>2>0` whose relative reference points -1 / 0 / +1 / +2 columns from the
//       first cell; one column moved from the first / second / last cell, to right after the area or one column later
//   p2  area C:G, relative offsets -2 -1 +1 +2, every source column C..G, destination J (far) or H (right after G)
//   p3  area C:G, `=$<A..L>2>0` (absolute column anywhere around the area), source D or E, destination H
//   p4  area C:G, two references `=<r1>><r2>` taken from C2, D2, G2, $F2, $K2, $B2, source D, destination H
//   p5  rule blocks: several rules on the same range (and one on another range) with source B, destination F
// Usage: node make-cf-probe-spec.mjs <out spec.json>
import fs from 'node:fs';

const numToCol = (n) => { let s = ''; for (; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s; return s; };
const out = [];
const mov = (from, to) => [{ kind: 'mov', from, count: 1, to }];

// p1
for (const [lo, hi] of [[2, 5], [3, 7]]) for (const o of [-1, 0, 1, 2]) {
  const f = `=${numToCol(lo + o)}2>0`;
  const add = (src, X, tag) => out.push({
    id: `p1_${numToCol(lo)}${numToCol(hi)}_o${o < 0 ? 'm' + -o : o}_${tag}`,
    rules: [{ type: 'expr', addr: `${numToCol(lo)}2:${numToCol(hi)}3`, active: `${numToCol(lo)}2`, f }],
    ops: mov(src, X),
  });
  add(lo, hi + 1, 'lo_to_hi1');
  add(lo + 1, hi + 1, 'mid_to_hi1');
  add(hi, hi + 2, 'hi_to_hi2');
  add(lo + 1, hi + 2, 'mid_to_hi2');
}
// p2
for (const o of [-2, -1, 1, 2]) for (const X of [10, 8]) for (let src = 3; src <= 7; src++) {
  if (X === src + 1) continue;
  out.push({
    id: `p2_o${o < 0 ? 'm' + -o : o}_s${numToCol(src)}_x${numToCol(X)}`,
    rules: [{ type: 'expr', addr: 'C2:G3', active: 'C2', f: `=${numToCol(3 + o)}2>0` }],
    ops: mov(src, X),
  });
}
// p3
for (let c = 1; c <= 12; c++) for (const src of [4, 5]) {
  out.push({
    id: `p3_abs${numToCol(c)}_s${numToCol(src)}_xH`,
    rules: [{ type: 'expr', addr: 'C2:G3', active: 'C2', f: `=$${numToCol(c)}2>0` }],
    ops: mov(src, 8),
  });
}
// p4
{
  const refs = ['C2', 'D2', 'G2', '$F2', '$K2', '$B2'];
  let k = 0;
  for (const r1 of refs) for (const r2 of refs) {
    if (r1 === r2) continue;
    out.push({ id: `p4_${String(k++).padStart(2, '0')}`, rules: [{ type: 'expr', addr: 'C2:G3', active: 'C2', f: `=${r1}>${r2}` }], ops: mov(4, 8) });
  }
}
// p5
{
  const R = (f, addr = 'B2:E3') => ({ type: 'expr', addr, active: 'B2', f });
  const m = mov(2, 6);
  out.push(
    { id: 'p5_a_alone', rules: [R('=C2>0')], ops: m },
    { id: 'p5_b_alone', rules: [R('=A2>0')], ops: m },
    { id: 'p5_ab', rules: [R('=C2>0'), R('=A2>0')], ops: m },
    { id: 'p5_ba', rules: [R('=A2>0'), R('=C2>0')], ops: m },
    { id: 'p5_a_const', rules: [R('=C2>0'), { type: 'cellIs', addr: 'B2:E3', active: 'B2', f: '=5' }], ops: m },
    { id: 'p5_a_twice', rules: [R('=C2>0'), R('=D2>0')], ops: m },
    { id: 'p5_a_diffrange', rules: [R('=C2>0'), R('=A2>0', 'B2:D3')], ops: m },
  );
}

const dest = process.argv[2];
if (!dest) { console.error('usage: node make-cf-probe-spec.mjs <out spec.json>'); process.exit(1); }
fs.writeFileSync(dest, JSON.stringify({ scenarios: out }, null, 1));
console.log(`wrote ${dest} (${out.length} scenarios)`);
