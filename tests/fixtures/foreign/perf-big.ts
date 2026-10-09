// node --max-old-space-size=2048 --import tsx tests/fixtures/foreign/perf-big.ts <big.xlsx> <partner.xlsx> [--role base|other] [--join left|full] [--out <dir>]
// Time and memory of the whole pipeline on a 50,000 x 30 workbook (make-big.py): readBook, region / keys / join, buildPremerged, readBook of the
// result, createEditor, 10 ops (moves incl. non-contiguous, inserts, a delete with convert, typed cells, a width), replayLog, readBook of the output.
// Peak heap / rss are sampled every 50 ms. Exit code 1 when a stage takes 20 s or more or the heap peak reaches 2 GB.
import fs from 'node:fs';
import path from 'node:path';
import { analyzeSources, defaultPlan, detectRegion, joinSources, pickDataSheet } from '../../../src/core/keys';
import { buildPremerged, readBook, replayLog } from '../../../src/core/xlsx';
import { createEditor, deleteColumns, insertColumns, moveColumns, setCellInput, setColumnWidth } from '../../../src/core/doc';
import type { KeySource, PremergeInput } from '../../../src/core/contracts';
import type { EditorSnapshot } from '../../../src/core/types';

const args = process.argv.slice(2);
const opt = (k: string, d: string) => { const i = args.indexOf(k); return i >= 0 ? args[i + 1] : d; };
const files = args.filter((a, i) => !a.startsWith('--') && !(i > 0 && args[i - 1].startsWith('--')));
const role = opt('--role', 'base') as 'base' | 'other';
const joinType = opt('--join', 'left') as 'left' | 'full';
const outDir = opt('--out', '');

let peakHeap = 0, peakRss = 0;
const sample = () => { const m = process.memoryUsage(); peakHeap = Math.max(peakHeap, m.heapUsed); peakRss = Math.max(peakRss, m.rss); };
const timer = setInterval(sample, 50);
const stages: Array<[string, number]> = [];
async function stage<T>(name: string, fn: () => T | Promise<T>): Promise<T> {
  const t = Date.now();
  const c0 = process.cpuUsage();
  const r = await fn();
  sample();
  const ms = Date.now() - t;
  const c1 = process.cpuUsage(c0);
  const cpu = Math.round((c1.user + c1.system) / 1000);
  stages.push([name, ms]);
  console.log(`${name.padEnd(34)} ${String(ms).padStart(7)} ms (cpu ${String(cpu).padStart(6)} ms)   heap ${(process.memoryUsage().heapUsed / 1048576).toFixed(0)} MB (peak ${(peakHeap / 1048576).toFixed(0)}), rss ${(process.memoryUsage().rss / 1048576).toFixed(0)} MB`);
  return r;
}

const [bigPath, partnerPath] = files;
const bigName = path.basename(bigPath);
const partnerName = path.basename(partnerPath);
const bigBytes = new Uint8Array(fs.readFileSync(bigPath));
const partnerBytes = new Uint8Array(fs.readFileSync(partnerPath));
console.log(`${bigName} ${(bigBytes.length / 1048576).toFixed(1)} MB + ${partnerName} ${(partnerBytes.length / 1048576).toFixed(1)} MB, big file is the ${role.toUpperCase()}, join ${joinType}, node ${process.version}`);

const bigBook = await stage('readBook big', () => readBook(bigBytes, bigName, { id: 'b' }));
const partnerBook = await stage('readBook partner', () => readBook(partnerBytes, partnerName, { id: 'p' }));
const src = (book: typeof bigBook, id: string, name: string): KeySource => {
  const sheet = pickDataSheet(book);
  return { id, book, region: detectRegion(book, sheet), label: name.replace(/\.[^.]+$/, '') };
};
const big = await stage('pickDataSheet + detectRegion', () => src(bigBook, role === 'base' ? 's1' : 's2', bigName));
const partner = src(partnerBook, role === 'base' ? 's2' : 's1', partnerName);
const base = role === 'base' ? big : partner;
const other = role === 'base' ? partner : big;
console.log(`  base region ${base.region.sheet}!h${base.region.headerRow} rows ${base.region.firstDataRow}-${base.region.lastDataRow} cols ${base.region.firstCol}-${base.region.lastCol}`);
const analysis = await stage('analyzeSources', () => analyzeSources(base, [other]));
const join = await stage('joinSources', () => joinSources(analysis, { ...defaultPlan(analysis), joinType }));
const input: PremergeInput = { base, sources: [other], rows: join.rows, imports: join.defaultImports, keyFill: join.keyFill, options: { formulaPolicy: 'keep-if-possible', extendTable: true, headerConflict: 'suffix' } };
console.log(`  ${join.rows.length} rows, ${join.defaultImports.length} imported columns`);
const pm = await stage('buildPremerged', () => buildPremerged(input));
const merged = await stage('readBook premerged', () => readBook(pm.bytes, base.book.fileName, { id: 'm' }));
let last = base.region.lastDataRow;
for (const a of pm.layout.appended) if (a.row > last) last = a.row;
let s: EditorSnapshot = await stage('createEditor', () => createEditor(merged, pm.layout, { s1: { label: 'a', color: '#2563eb' }, s2: { label: 'b', color: '#15803d' } }, { headerRow: base.region.headerRow, firstDataRow: base.region.firstDataRow, lastDataRow: last }));
const n = () => s.doc.columns.length;
const ids = (...i: number[]) => i.map((k) => s.doc.columns[Math.min(k, n() - 1)].id);
const apply = async (label: string, f: () => ReturnType<typeof moveColumns>) => {
  const r = await stage(label, f);
  if (!r.ok) console.log(`    refused: ${r.reason}`);
  else s = r.snapshot;
};
await apply('op: move 3 columns', () => moveColumns(s, ids(2, 3, 4), n(), { unmergeTitles: true }));
await apply('op: move non-contiguous', () => moveColumns(s, ids(0, 5, 9), 12, { unmergeTitles: true }));
await apply('op: insert 2 columns', () => insertColumns(s, 6, 2));
await apply('op: delete 2 columns (convert)', () => deleteColumns(s, ids(10, 11), { convertDependents: true }));
await apply('op: move to the front', () => moveColumns(s, ids(20, 21), 0, { unmergeTitles: true }));
for (const [i, text] of ['=SUM(I2:I100)', '2026/10/20', '50%', 'メモ', "'007"].entries()) {
  const r = setCellInput(s, s.doc.columns[i + 2].id, 2 + i, text);
  if (r.ok) s = r.snapshot;
}
await apply('op: column width', () => setColumnWidth(s, s.doc.columns[3].id, 21.5));
console.log(`  ${s.log.length} logged ops`);
const out = await stage('replayLog', () => replayLog(pm.bytes, s.log));
const back = await stage('readBook output', () => readBook(out, base.book.fileName, { id: 'o' }));
clearInterval(timer);
sample();
const sheet = back.sheets.find((g) => g.name === s.doc.sheetName)!;
console.log(`output ${(out.length / 1048576).toFixed(1)} MB, ${sheet.maxRow} rows x ${sheet.maxCol} columns`);
if (outDir) {
  fs.mkdirSync(outDir, { recursive: true });
  fs.writeFileSync(path.join(outDir, `perf-${role}-${joinType}.pre.xlsx`), pm.bytes);
  fs.writeFileSync(path.join(outDir, `perf-${role}-${joinType}.out.xlsx`), out);
}
const slowest = stages.reduce((a, b) => (b[1] > a[1] ? b : a));
const total = stages.reduce((a, b) => a + b[1], 0);
console.log(`\nslowest stage: ${slowest[0]} ${slowest[1]} ms; total ${total} ms; peak heap ${(peakHeap / 1048576).toFixed(0)} MB; peak rss ${(peakRss / 1048576).toFixed(0)} MB`);
process.exit(slowest[1] >= 20000 || peakHeap >= 2 * 1024 ** 3 ? 1 : 0);
