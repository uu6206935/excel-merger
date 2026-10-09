// Builds tests/doc/fixtures/r2b-doc-sample.json from the R2b Excel ground truth (research/r2b/excel-ref-fixtures.json):
// column ops (move / insert / delete) of the suites main, main_jp, table and merged, with the baseline state of the
// op sheet (formulas hosted on it, formulas on the other sheets, names, table, merges, array, autofilter, freeze,
// widths, row-1 markers) and Excel's result after each op. Cases where the excelRef engine itself documents a
// deviation from Excel (tests/formula/DEVIATIONS.md) are listed in `dev` (the doc must then equal excelRef).
//
//   node tests/doc/fixtures/build-r2b-sample.mjs <path to excel-ref-fixtures.json>
// (bundles src/core/formula/excelRef.ts in memory with esbuild and imports it as a data: URL)
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import * as esbuild from 'esbuild';

const here = path.dirname(fileURLToPath(import.meta.url));
const src = process.argv[2];
if (!src) throw new Error('usage: node build-r2b-sample.mjs <excel-ref-fixtures.json>');
const bundle = await esbuild.build({
  entryPoints: [path.resolve(here, '../../../src/core/formula/excelRef.ts')],
  bundle: true,
  format: 'esm',
  platform: 'neutral',
  write: false,
});
const X = await import('data:text/javascript;base64,' + Buffer.from(bundle.outputFiles[0].text).toString('base64'));
const fx = JSON.parse(fs.readFileSync(src, 'utf8').replace(/^﻿/, ''));

const SUITES = ['main', 'main_jp', 'table', 'merged'];
const COL_PRIMS = new Set(['moveCols', 'insertCols', 'deleteCols']);
const strip = (f) => (typeof f === 'string' && f.startsWith('=') ? f.slice(1) : f);
const fileText = (f) => (f == null ? null : typeof f === 'string' ? f : f.f);

function toOp(o) {
  const P = o.params;
  if (o.primitive === 'moveCols') return { type: 'move', axis: 'col', from: P.src.start, count: P.src.end - P.src.start + 1, to: P.before.index };
  if (o.primitive === 'insertCols') return { type: 'insert', axis: 'col', at: P.before.index, count: P.count };
  return { type: 'delete', axis: 'col', at: P.src.start, count: P.src.end - P.src.start + 1 };
}

const out = { source: 'research/r2b/excel-ref-fixtures.json (Excel 365 via COM)', suites: {}, ops: [], dev: [] };
const casesByOp = new Map();
for (const c of fx.cases) {
  if (!SUITES.includes(c.suite)) continue;
  let l = casesByOp.get(c.opId);
  if (!l) casesByOp.set(c.opId, (l = []));
  l.push(c);
}

for (const suite of SUITES) {
  const S = fx.suites[suite];
  const ops = fx.ops.filter((o) => o.suite === suite && COL_PRIMS.has(o.primitive));
  const first = casesByOp.get(ops[0].opId);
  const data = S.dataSheet;
  const formulas = [];
  const names = [];
  let arrays = [];
  let merges = [];
  let widths = null;
  let markers = null;
  let tables = [];
  let autofilter = null;
  let freeze = 0;
  for (const c of first) {
    if (c.kind === 'formula' && fileText(c.fileBefore) != null) {
      const f = { id: c.id.split(':')[1], where: c.where, cell: c.cellBefore, f: fileText(c.fileBefore) };
      if (typeof c.fileBefore === 'object' && c.fileBefore.t === 'array') {
        f.cell = c.fileBefore.ref.split(':')[0];
        f.arr = c.fileBefore.ref;
        if (c.where === data) arrays.push(c.fileBefore.ref);
      }
      formulas.push(f);
    } else if (c.kind === 'definedName' && c.before != null) {
      const bare = c.name.includes('!') ? c.name.slice(c.name.lastIndexOf('!') + 1) : c.name;
      const scope = /\.xlsx$/i.test(c.scope ?? '') ? null : c.scope ?? null;
      names.push({ key: c.name, name: bare, scope, refersTo: strip(c.before) });
    } else if (c.kind === 'mergedCells') merges = c.before;
    else if (c.kind === 'columnWidths') widths = c.before;
    else if (c.kind === 'columnOrder') markers = c.before;
    else if (c.kind === 'autoFilter') autofilter = c.before;
    else if (c.kind === 'frozenPane') freeze = Number(c.before?.xSplit ?? 0);
  }
  for (const t of S.baseline.tables ?? []) tables.push({ name: t.name, ref: t.range, columns: t.headers });
  out.suites[suite] = { dataSheet: data, sheets: S.sheets, formulas, names, tables, merges, arrays, autofilter, freeze, widths, markers };

  for (const o of ops) {
    const op = toOp(o);
    const cs = casesByOp.get(o.opId) ?? [];
    const rec = { opId: o.opId, suite, op, refused: !!o.error, silent: !!o.primarySilentBlankInsert, formulas: [], names: [] };
    const tablesRef = tables.map((t) => ({ name: t.name, sheet: data, ref: t.ref, columns: t.columns }));
    for (const c of cs) {
      if (c.kind === 'formula') {
        const fb = fileText(c.fileBefore);
        const fa = fileText(c.fileAfter);
        if (fb == null) continue;
        const id = c.id.split(':')[1];
        const cellAfter = c.cellAfter == null ? null : typeof c.fileAfter === 'object' && c.fileAfter?.ref ? c.fileAfter.ref.split(':')[0] : c.cellAfter;
        rec.formulas.push([id, cellAfter, fa]);
        if (fa != null && !o.error) {
          const hc = c.where === data && c.cellBefore ? X.parseAreaText(c.cellBefore.split(':')[0]) : null;
          const ours = X.adjustFormula(fb, op, { hostSheet: c.where, targetSheet: data, tables: tablesRef, hostCell: hc ? { row: hc.r1, col: hc.c1 } : undefined }).formula;
          // host cells on the op sheet travel with their column (excelRef.mapIndex)
          const nc = hc ? X.mapIndex(hc.c1, op) : null;
          const hostOurs = hc ? (nc == null ? null : X.numToCol(nc) + hc.r1) : cellAfter;
          if (ours !== fa || hostOurs !== cellAfter) out.dev.push(c.id);
        }
      } else if (c.kind === 'definedName' && c.before != null) {
        rec.names.push([c.name, c.after == null ? null : strip(c.after)]);
        if (c.after != null && !o.error) {
          const ours = X.adjustFormula(strip(c.before), op, { hostSheet: null, targetSheet: data, isName: true, tables: tablesRef }).formula;
          if (ours !== strip(c.after)) out.dev.push(c.id);
        }
      } else if (c.kind === 'mergedCells') rec.merges = c.after;
      else if (c.kind === 'columnWidths') rec.widths = c.after;
      else if (c.kind === 'columnOrder') rec.markers = c.after;
      else if (c.kind === 'autoFilter') rec.autofilter = c.after;
      else if (c.kind === 'frozenPane') rec.freeze = Number(c.after?.xSplit ?? 0);
      else if (c.kind === 'table') rec.table = c.after ? { ref: c.after.ref, columns: c.after.columns.map((x) => x.name) } : null;
    }
    out.ops.push(rec);
  }
}

const dest = path.join(here, 'r2b-doc-sample.json');
fs.writeFileSync(dest, JSON.stringify(out));
console.log(`wrote ${dest}: ${out.ops.length} ops, ${out.ops.reduce((n, o) => n + o.formulas.length, 0)} formula cases, ${out.dev.length} excelRef deviations`);
