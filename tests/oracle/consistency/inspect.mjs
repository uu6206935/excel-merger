// Prints cells (value + formula) of an oracle dump: node tests/oracle/consistency/inspect.mjs <dump.values.json> <sheet> <A1 | A1:C5> [more dumps ...]
// With several dumps the same range is printed for each (baseline / js / xl side by side by name).
import fs from 'node:fs';

const [, , ...rest] = process.argv;
const files = rest.filter((a) => a.endsWith('.json'));
const [sheet, range] = rest.filter((a) => !a.endsWith('.json'));
const colNum = (s) => [...s].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
const colLet = (n) => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
const m = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/.exec(range ?? '');
if (!m || !files.length) { console.error('usage: inspect.mjs <dump.json>... <sheet> <A1[:B2]>'); process.exit(2); }
const [c1, r1, c2, r2] = [colNum(m[1]), Number(m[2]), colNum(m[3] ?? m[1]), Number(m[4] ?? m[2])];
const ERR = { '-2146826281': '#DIV/0!', '-2146826246': '#N/A', '-2146826259': '#NAME?', '-2146826288': '#NULL!', '-2146826252': '#NUM!', '-2146826265': '#REF!', '-2146826273': '#VALUE!' };
for (const f of files) {
  const d = JSON.parse(fs.readFileSync(f, 'utf8').replace(/^﻿/, ''));
  console.log(`== ${f.split(/[\\/]/).pop()} open=${d.open} ${d.log ?? ''}`);
  const s = d.sheets.find((x) => x.name === sheet);
  if (!s) { console.log(`  (no sheet ${sheet}; sheets: ${d.sheets.map((x) => x.name).join(', ')})`); continue; }
  for (let r = r1; r <= r2; r++) {
    for (let c = c1; c <= c2; c++) {
      const v = s.v[r - 1]?.[c - 1];
      const fm = s.f[r - 1]?.[c - 1];
      const t = v && typeof v === 'object' ? ERR[v.e] ?? `#E${v.e}` : JSON.stringify(v);
      console.log(`  ${colLet(c)}${r}: ${t}${fm ? `   ${fm}` : ''}`);
    }
  }
}
