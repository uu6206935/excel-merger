// Builds tests/fixtures/formula/r2b-fixtures.json (a trimmed copy of the R2b Excel ground truth) from
//   research/r2b/excel-ref-fixtures.json   (9 MB; schema: docs/research/R2b-excel-ground-truth.md "Fixture schema")
// Keeps every case of the kinds the formula engine can check (formula, definedName, conditionalFormat, tableCom,
// dataValidation, autoFilter, mergedCells, mergeTracked, table, columnOrder, columnWidths) and only the fields the
// tests read. (frozenPane cases are the patcher's: src/core/xlsx/patch/views.ts.)
// Usage: node tests/fixtures/formula/build-r2b-fixture.mjs <path to excel-ref-fixtures.json>
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const here = path.dirname(fileURLToPath(import.meta.url));
const src = process.argv[2];
if (!src) { console.error('usage: node build-r2b-fixture.mjs <excel-ref-fixtures.json>'); process.exit(1); }
const fx = JSON.parse(fs.readFileSync(src, 'utf8'));

const KINDS = new Set(['formula', 'definedName', 'conditionalFormat', 'dataValidation', 'autoFilter', 'mergedCells', 'mergeTracked', 'table',
  'tableCom', 'columnOrder', 'columnWidths']);
const CASE_FIELDS = ['id', 'opId', 'where', 'kind', 'tag', 'cellBefore', 'cellAfter', 'before', 'after', 'before2', 'after2',
  'fileBefore', 'fileAfter', 'arrayBefore', 'arrayAfter', 'source', 'name', 'scope', 'dvSheet', 'cfSheet', 'ext', 'color', 'notes'];

const ops = fx.ops.map((o) => ({
  opId: o.opId,
  suite: o.suite,
  primitive: o.primitive,
  sheet: o.sheet,
  params: o.params,
  afterSource: o.afterSource,
  error: o.error ? o.error.message ?? String(o.error) : null,
  noTruth: !!(o.hung || o.crashed || o.noGroundTruth || o.markersOk === null),
}));

const cases = [];
for (const c of fx.cases) {
  if (!KINDS.has(c.kind)) continue;
  const o = {};
  for (const k of CASE_FIELDS) if (c[k] !== undefined && c[k] !== null) o[k] = c[k];
  if (c.after === null) o.after = null; // host deleted / feature removed
  if (c.kind === 'conditionalFormat' && c.spec) o.cfType = c.spec.type;
  cases.push(o);
}

const suites = Object.fromEntries(Object.entries(fx.suites).map(([k, s]) => [k, {
  dataSheet: s.dataSheet,
  sheets: s.sheets,
  tables: (s.baseline?.tables ?? []).map((t) => ({ name: t.name, sheet: t.sheet, ref: t.range, columns: t.headers })),
}]));

const out = {
  meta: {
    description: 'Trimmed copy of the R2b Excel ground truth (research/r2b/excel-ref-fixtures.json). Real Excel 365 results of structural ops. See docs/research/R2b-excel-ground-truth.md.',
    excel: fx.meta.excel,
    generated: fx.meta.generated,
    trimmedBy: 'tests/fixtures/formula/build-r2b-fixture.mjs',
  },
  suites,
  ops,
  cases,
};
const dest = path.join(here, 'r2b-fixtures.json');
fs.writeFileSync(dest, JSON.stringify(out));
console.log(`wrote ${dest}: ${ops.length} ops, ${cases.length} cases, ${(fs.statSync(dest).size / 1e6).toFixed(2)} MB`);
