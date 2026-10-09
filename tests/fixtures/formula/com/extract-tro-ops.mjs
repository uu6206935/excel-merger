// Joins the COM read-back of excel-tro-ops.ps1 (tro-ops-com.json) with the <f> texts stored in the saved workbooks
// (tro-<op>.xlsx) and writes ../tro-ops-excel-results.json (used by tests/formula/tro.test.ts).
// Usage: node extract-tro-ops.mjs <dir with tro-ops-com.json and tro-*.xlsx>
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';

const here = path.dirname(fileURLToPath(import.meta.url));
const dir = process.argv[2];
if (!dir) { console.error('usage: node extract-tro-ops.mjs <dir>'); process.exit(1); }
const decode = (s) => s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
const attr = (tag, name) => { const m = new RegExp(`\\s${name}="([^"]*)"`).exec(tag); return m ? decode(m[1]) : undefined; };

async function storedFormulas(file) {
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  const read = (p) => zip.file(p).async('string');
  const wb = await read('xl/workbook.xml');
  const rels = await read('xl/_rels/workbook.xml.rels');
  const out = {};
  for (const m of wb.matchAll(/<sheet\b[^>]*>/g)) {
    const name = attr(m[0], 'name'), rid = attr(m[0], 'r:id');
    const rel = [...rels.matchAll(/<Relationship\b[^>]*>/g)].find((r) => attr(r[0], 'Id') === rid);
    const xml = await read('xl/' + attr(rel[0], 'Target').replace(/^\/?xl\//, ''));
    for (const c of xml.matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
      const fm = /<f\b([^>]*?)(?:\/>|>([\s\S]*?)<\/f>)/.exec(c[2] ?? '');
      if (fm && fm[2] != null) out[`${name}!${attr(c[1], 'r')}`] = decode(fm[2]);
    }
  }
  return out;
}

const com = JSON.parse(fs.readFileSync(path.join(dir, 'tro-ops-com.json'), 'utf8').replace(/^﻿/, ''));
const byOp = {};
for (const op of [...new Set(com.map((c) => c.op))]) byOp[op] = await storedFormulas(path.join(dir, `tro-${op}.xlsx`));
const cells = com.map((c) => ({ ...c, file: byOp[c.op][`${c.sheet}!${c.cell}`] ?? null }));
const dest = path.join(here, '..', 'tro-ops-excel-results.json');
fs.writeFileSync(dest, JSON.stringify({
  meta: {
    description: 'Trim-range references under structural ops in Excel 365: formulas typed via Range.Formula2 on a fresh workbook per op, op applied on sheet Data (clip-free moves), Formula2 + stored <f> read back. Built by tests/fixtures/formula/com/{excel-tro-ops.ps1,extract-tro-ops.mjs}.',
    generated: new Date().toISOString(),
  },
  cells,
}, null, 1));
console.log(`wrote ${dest}: ${cells.length} cells`);
