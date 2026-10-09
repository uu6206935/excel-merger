// Reads the workbook saved by excel-forms.ps1 and joins every cell's stored <f> text (and t / ref / cm attributes)
// with the COM read-back (forms-com.json). Writes ../forms-excel-results.json (the fixture used by forms.test.ts).
// Usage: node extract-forms.mjs <forms-excel.xlsx> <forms-com.json> [fixture name, default forms-excel-results.json]
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';

const here = path.dirname(fileURLToPath(import.meta.url));
const [, , xlsxPath, comPath, outName = 'forms-excel-results.json'] = process.argv;
const zip = await JSZip.loadAsync(fs.readFileSync(xlsxPath));
const read = (p) => zip.file(p).async('string');
const decode = (s) => s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&#(\d+);/g, (_, d) => String.fromCodePoint(+d)).replace(/&#x([0-9a-f]+);/gi, (_, h) => String.fromCodePoint(parseInt(h, 16))).replace(/&amp;/g, '&');
const attr = (tag, name) => { const m = new RegExp(`\\s${name}="([^"]*)"`).exec(tag); return m ? decode(m[1]) : undefined; };

const wb = await read('xl/workbook.xml');
const rels = await read('xl/_rels/workbook.xml.rels');
const sheetPath = {};
for (const m of wb.matchAll(/<sheet\b[^>]*>/g)) {
  const name = attr(m[0], 'name'), rid = attr(m[0], 'r:id');
  const rel = [...rels.matchAll(/<Relationship\b[^>]*>/g)].find((r) => attr(r[0], 'Id') === rid);
  sheetPath[name] = 'xl/' + attr(rel[0], 'Target').replace(/^\/?xl\//, '');
}
const stored = {};
for (const [name, p] of Object.entries(sheetPath)) {
  const xml = await read(p);
  for (const m of xml.matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
    const r = attr(m[1], 'r'), cm = attr(m[1], 'cm');
    const fm = /<f\b([^>]*?)(?:\/>|>([\s\S]*?)<\/f>)/.exec(m[2] ?? '');
    if (!fm) continue;
    stored[`${name}!${r}`] = { f: fm[2] != null ? decode(fm[2]) : null, t: attr(fm[1], 't'), ref: attr(fm[1], 'ref'), cm };
  }
}
const com = JSON.parse(fs.readFileSync(comPath, 'utf8').replace(/^﻿/, ''));
const out = com.map((c) => ({ ...c, file: stored[`${c.sheet}!${c.cell}`] ?? null }));
const dest = path.join(here, '..', outName);
fs.writeFileSync(dest, JSON.stringify({
  meta: { description: 'Formulas typed into Excel 365 via Range.Formula2, saved as .xlsx; file = stored <f>. Built by tests/fixtures/formula/com/{make-forms-input.mjs,excel-forms.ps1 | make-forms-extra-input.mjs,excel-forms-extra.ps1 ; extract-forms.mjs}.', generated: new Date().toISOString() },
  cells: out,
}, null, 1));
console.log(`wrote ${dest}: ${out.length} cells, ${out.filter((c) => c.file).length} with <f>`);
