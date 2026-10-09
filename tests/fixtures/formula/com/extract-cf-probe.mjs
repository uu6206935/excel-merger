// Reads the workbooks saved by excel-cf-probe.ps1 (<dir>/<id>.xlsx) and writes <out.json>:
//   { meta, scenarios: [ { id, rules, ops, excel: [ { kind, type, sqref, formulas } ] } ] }
// (spec + what Excel left behind: classic <conditionalFormatting>, x14 conditional formats and data validations of
// sheet P; the x14 twin of a dataBar is skipped).   Usage: node extract-cf-probe.mjs <spec.json> <dir with *.xlsx> <out.json>
import fs from 'node:fs';
import path from 'node:path';
import JSZip from 'jszip';

const [specFile, dir, outFile] = process.argv.slice(2);
if (!specFile || !dir || !outFile) { console.error('usage: node extract-cf-probe.mjs <spec.json> <dir> <out.json>'); process.exit(1); }
const decode = (s) => s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
const attr = (tag, name) => { const m = new RegExp(`\\s${name}="([^"]*)"`).exec(tag); return m ? decode(m[1]) : undefined; };

export async function rulesOf(file) {
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

const spec = JSON.parse(fs.readFileSync(specFile, 'utf8').replace(/^﻿/, ''));
const scenarios = [];
for (const sc of spec.scenarios) {
  const f = path.join(dir, `${sc.id}.xlsx`);
  if (!fs.existsSync(f)) { console.error(`missing ${f}`); process.exit(2); }
  scenarios.push({ ...sc, excel: await rulesOf(f) });
}
fs.writeFileSync(outFile, JSON.stringify({
  meta: {
    description: 'Conditional formats after ONE column move in Excel 365 (the clipboard-free sequence of the Excel runner: insert blank columns, Range.Cut(Destination), delete), for small rule families. Built by tests/fixtures/formula/com/{make-cf-probe-spec.mjs,excel-cf-probe.ps1,extract-cf-probe.mjs}; used by tests/formula/cfProbe.test.ts.',
    generated: new Date().toISOString(),
  },
  scenarios,
}, null, 1));
console.log(`wrote ${outFile} (${scenarios.length} scenarios)`);
