// Reads the workbooks saved by excel-cf-spread.ps1 (cfs-<op>.xlsx) and writes ../cf-spread-excel-results.json
// (used by tests/formula/cfSpread.test.ts): for every op the conditional formats (classic <conditionalFormatting>;
// the x14 twin of a dataBar has the same sqref and is skipped) and data validations of sheet P as
// {kind, type, sqref, formulas}.
// Usage: node extract-cf-spread.mjs <dir with cfs-*.xlsx>
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';

const here = path.dirname(fileURLToPath(import.meta.url));
const dir = process.argv[2];
if (!dir) { console.error('usage: node extract-cf-spread.mjs <dir>'); process.exit(1); }
const decode = (s) => s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
const attr = (tag, name) => { const m = new RegExp(`\\s${name}="([^"]*)"`).exec(tag); return m ? decode(m[1]) : undefined; };

// ops of excel-cf-spread.ps1 as excelRef StructOps (1-based columns; move.to = insert-before index)
const OPS = {
  before: null,
  insD: { type: 'insert', axis: 'col', at: 4, count: 1 },
  insF2: { type: 'insert', axis: 'col', at: 6, count: 2 },
  insC: { type: 'insert', axis: 'col', at: 3, count: 1 },
  delD: { type: 'delete', axis: 'col', at: 4, count: 1 },
  movGtoD: { type: 'move', axis: 'col', from: 7, count: 1, to: 4 },
  movAtoD: { type: 'move', axis: 'col', from: 1, count: 1, to: 4 },
  movHtoF: { type: 'move', axis: 'col', from: 8, count: 1, to: 6 },
  movIJtoF: { type: 'move', axis: 'col', from: 9, count: 2, to: 6 },
  movBHtoL: { type: 'move', axis: 'col', from: 2, count: 7, to: 12 },
  movFGtoC: { type: 'move', axis: 'col', from: 6, count: 2, to: 3 },
  movDtoH: { type: 'move', axis: 'col', from: 4, count: 1, to: 8 },
  insF: { type: 'insert', axis: 'col', at: 6, count: 1 },
  delF: { type: 'delete', axis: 'col', at: 6, count: 1 },
  delEF: { type: 'delete', axis: 'col', at: 5, count: 2 },
  insA: { type: 'insert', axis: 'col', at: 1, count: 1 },
  delB: { type: 'delete', axis: 'col', at: 2, count: 1 },
};

async function rulesOf(file) {
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  const xml = await zip.file('xl/worksheets/sheet1.xml').async('string');
  const out = [];
  for (const m of xml.matchAll(/<conditionalFormatting\b([^>]*)>([\s\S]*?)<\/conditionalFormatting>/g)) {
    const sqref = attr(m[0].slice(0, m[0].indexOf('>') + 1), 'sqref');
    for (const r of m[2].matchAll(/<cfRule\b([^>]*?)(?:\/>|>([\s\S]*?)<\/cfRule>)/g)) {
      const formulas = [...(r[2] ?? '').matchAll(/<formula>([\s\S]*?)<\/formula>/g)].map((f) => decode(f[1]));
      out.push({ kind: 'cf', type: attr(r[0], 'type'), sqref, formulas });
    }
  }
  // x14 rules (cross-sheet formulas); the x14 twin of a classic dataBar is skipped
  for (const m of xml.matchAll(/<x14:conditionalFormatting\b[^>]*>([\s\S]*?)<\/x14:conditionalFormatting>/g)) {
    const sqref = /<xm:sqref>([\s\S]*?)<\/xm:sqref>/.exec(m[1])?.[1] ?? '';
    for (const r of m[1].matchAll(/<x14:cfRule\b([^>]*?)(?:\/>|>([\s\S]*?)<\/x14:cfRule>)/g)) {
      const type = attr(r[0], 'type');
      if (type === 'dataBar') continue;
      const formulas = [...(r[2] ?? '').matchAll(/<xm:f>([\s\S]*?)<\/xm:f>/g)].map((f) => decode(f[1]));
      out.push({ kind: 'cf', type, sqref: decode(sqref), formulas, x14: true });
    }
  }
  for (const m of xml.matchAll(/<dataValidation\b([^>]*?)(?:\/>|>([\s\S]*?)<\/dataValidation>)/g)) {
    const formulas = [...(m[2] ?? '').matchAll(/<(formula1|formula2)>([\s\S]*?)<\/\1>/g)].map((f) => decode(f[2]));
    out.push({ kind: 'dv', type: attr(m[0], 'type'), sqref: attr(m[0], 'sqref'), formulas });
  }
  return out;
}

const files = {};
for (const op of Object.keys(OPS)) files[op] = await rulesOf(path.join(dir, `cfs-${op}.xlsx`));
const dest = path.join(here, '..', 'cf-spread-excel-results.json');
fs.writeFileSync(dest, JSON.stringify({
  meta: {
    description: 'Conditional formats / data validations of every rule type under column insert / delete / move in Excel 365 (moves = the Excel runner\'s clipboard-free sequence: insert blank + Range.Cut(Destination) + delete). One band of 2 rows per rule (area C or C:E) plus R2b replicas (J:K cellIs =$C$5, B list =$B$2:$H$2). Built by tests/fixtures/formula/com/{excel-cf-spread.ps1,extract-cf-spread.mjs}.',
    generated: new Date().toISOString(),
  },
  ops: OPS,
  files,
}, null, 1));
console.log(`wrote ${dest}`);
for (const [op, rules] of Object.entries(files)) {
  console.log(`== ${op}`);
  for (const r of rules) console.log(`  ${r.kind} ${r.type} ${r.sqref} ${JSON.stringify(r.formulas)}`);
}
