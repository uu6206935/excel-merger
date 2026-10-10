/**
 * Writes premerged outputs WITH copied sheets that own objects (task A2: charts, pictures, shapes, notes, tables, pivot tables, chart sheets;
 * PremergeOptions.includeSourceSheets) into a folder, together with a manifest, so that real Excel can be asked (one hidden instance,
 * tests/golden/com/check-addsheets-objects.ps1):
 *   npx tsx tests/golden/tools/make-addsheets-objects-outputs.ts <folder> [<regex of combination names>]
 *   (EXCEL_ENGINE=1 adds a combination whose edits are replayed by the app's Excel engine: one hidden Excel start)
 * <folder>/out/<name>.xlsx, <folder>/manifest.json
 * [{ name, out, expected: [sheet names in order], sources: [{ file, label, sheets: [{ original, name, complete, objects }] }] }]
 * where `objects` = what the source sheet owns (tests/xlsx/addsheets-objects-helpers.ts objectsOf), the check compares it with Excel's view of the copy.
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';
import { analyzeSources, defaultPlan, detectRegion, joinSources, pickDataSheet } from '../../../src/core/keys';
import type { KeySource, PremergeInput } from '../../../src/core/contracts';
import { createEditor } from '../../../src/core/doc';
import { buildPremerged, readBook, replayLog } from '../../../src/core/xlsx';
import { Engine } from '../../../src/core/xlsx/patch/engine';
import { sameTheme, themeXml } from '../../../src/core/xlsx/patch/bookinfo';
import { rng, step } from '../../doc/fuzzcore';
import { plannedNames } from '../../xlsx/addsheets-helpers';
import { lintObjects, lintObjectsText, objectsOf, type SheetObjects } from '../../xlsx/addsheets-objects-helpers';
import { runRunnerJob } from '../../oracle/consistency/excel';

const here = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(here, '../../..');
const out = path.resolve(process.argv[2] ?? 'addsheets-objects-out');
fs.mkdirSync(path.join(out, 'out'), { recursive: true });

const S = (n: string) => path.join(root, 'samples', n);
const X = (n: string) => path.join(root, 'tests/fixtures/addsheets', n);
const O = (n: string) => path.join(root, 'tests/fixtures/addsheets-objects', n);
const G = (n: string) => path.join(root, 'tests/fixtures/golden', n);
const C = (n: string) => path.join(root, 'tests/fixtures/foreign/corpus', n);

interface Entry {
  name: string;
  out: string;
  expected: string[];
  note: string;
  lint: string;
  /** the copy holds literal chart data for references that cannot follow: series formulas are not compared */
  literalCharts: boolean;
  sources: Array<{ file: string; label: string; themeDiffers: boolean; drawingsSkipped: boolean; sheets: Array<{ original: string; name: string; complete: boolean; objects: SheetObjects }> }>;
}
const manifest: Entry[] = [];

async function load(file: string, id: string, fileName = path.basename(file)): Promise<KeySource> {
  const bytes = new Uint8Array(fs.readFileSync(file));
  const book = await readBook(bytes, fileName, { id });
  return { id, book, region: detectRegion(book, pickDataSheet(book)), label: fileName.replace(/\.[^.]+$/, '') };
}

interface Combo {
  name: string;
  base: string;
  others: Array<string | { file: string; fileName: string }>;
  drop?: boolean;
  join?: 'left' | 'full';
  ops?: number;
  engine?: 'excel';
  note?: string;
  literalCharts?: boolean;
}

async function run(c: Combo): Promise<void> {
  const base = await load(c.base, 'b');
  const others: KeySource[] = [];
  const files: string[] = [];
  for (const [i, o] of c.others.entries()) {
    const file = typeof o === 'string' ? o : o.file;
    files.push(file);
    others.push(await load(file, `s${i + 2}`, typeof o === 'string' ? path.basename(o) : o.fileName));
  }
  let input: PremergeInput;
  if (c.join) {
    const analysis = analyzeSources(base, others);
    const join = joinSources(analysis, { ...defaultPlan(analysis), joinType: c.join });
    input = { base, sources: others, rows: join.rows, imports: join.defaultImports, keyFill: join.keyFill, keyPairs: join.keyPairs, options: { formulaPolicy: 'keep-if-possible', extendTable: true, headerConflict: 'suffix' } };
  } else {
    const rows = [];
    for (let r = base.region.firstDataRow; r <= base.region.lastDataRow; r++) rows.push({ baseRow: r, matches: {} });
    input = { base, sources: others, rows, imports: [], keyFill: {}, options: { formulaPolicy: 'keep-if-possible', extendTable: true, headerConflict: 'suffix' } };
  }
  const drop = c.drop ?? true;
  input.options = { ...input.options, includeSourceSheets: true, dropOtherSheets: drop };
  const res = await buildPremerged(input);
  let bytes = res.bytes;
  if (c.ops) {
    const premerged = await readBook(bytes, 'merged.xlsx', { id: 'merged' });
    const colors = ['#2563eb', '#15803d', '#c2410c', '#7c3aed', '#0e7490'];
    const info: Record<string, { label: string; color: string }> = {};
    [base, ...others].forEach((k, i) => (info[k.id] = { label: k.label, color: colors[i % colors.length] }));
    let last = base.region.lastDataRow;
    for (const a of res.layout.appended) if (a.row > last) last = a.row;
    let snap = createEditor(premerged, res.layout, info, { headerRow: base.region.headerRow, firstDataRow: base.region.firstDataRow, lastDataRow: last });
    const r = rng(c.ops * 7919);
    for (let i = 0; i < 10; i++) {
      const { res: sr } = step(snap, r, false, snap.doc.lastDataRow + 2);
      if (sr.ok) snap = sr.snapshot;
    }
    if (c.engine === 'excel') {
      const work = path.join(out, '_work');
      fs.mkdirSync(work, { recursive: true });
      const inp = path.join(work, `${c.name}-in.xlsx`), outp = path.join(work, `${c.name}-out.xlsx`);
      fs.writeFileSync(inp, bytes);
      const outcome = await runRunnerJob(inp, snap.log, outp, res.layout.sheet, { tempRoot: work, cacheDir: path.join(work, 'cache') });
      if (!outcome.ok || !fs.existsSync(outp)) throw new Error(`${c.name}: Excel engine failed: ${outcome.code} ${outcome.message}`);
      bytes = new Uint8Array(fs.readFileSync(outp));
      fs.rmSync(inp, { force: true });
      fs.rmSync(outp, { force: true });
    } else {
      bytes = await replayLog(bytes, snap.log);
    }
  }
  const added = res.layout.addedSheets ?? [];
  const keptBase = drop ? [] : base.book.sheets.map((s) => s.name).filter((n) => n !== res.layout.sheet);
  // the expected names: the naming rule over every sheet the copier takes (worksheets and chart sheets)
  const srcFiles = await Promise.all(others.map(async (o) => ({ fileName: o.book.fileName, sheets: (await objectsOf(o.book.bytes)).map((s) => s.name) })));
  const expectedCopies = plannedNames([res.layout.sheet, ...keptBase], srcFiles);
  const gotCopies = added.map((a) => a.name);
  if (JSON.stringify(gotCopies) !== JSON.stringify(expectedCopies)) throw new Error(`${c.name}: layout says ${gotCopies.join('|')}, rule says ${expectedCopies.join('|')}`);
  // (a base in Strict namespaces keeps its docProps/app.xml: A1's updateAppSheets does not know the Strict relationship type; the output of the Excel engine is Excel's own
  // file: calcChain, shared formulas ... are its business, not the patcher's)
  const strictBase = new TextDecoder().decode((await JSZip.loadAsync(base.book.bytes).then((z) => z.file('xl/workbook.xml')?.async('uint8array'))) ?? new Uint8Array()).includes('purl.oclc.org');
  const lint = c.engine === 'excel' ? '' : lintObjectsText((await lintObjects(bytes)).filter((l) => !(strictBase && /^(APP_|VML_ID_DUPLICATE)/.test(l.code))));
  const meta: Entry['sources'] = [];
  for (const [i, ks] of others.entries()) {
    const objs = await objectsOf(ks.book.bytes);
    // automatic series colours follow the theme of the workbook they are in: only meaningful when the two themes agree
    // (drawings are copied between packages of the same flavour only: Strict / transitional)
    const strictOf = async (b: Uint8Array) => new TextDecoder().decode((await JSZip.loadAsync(b).then((z) => z.file('xl/workbook.xml')?.async('uint8array'))) ?? new Uint8Array()).includes('purl.oclc.org');
    const drawingsSkipped = (await strictOf(ks.book.bytes)) !== (await strictOf(base.book.bytes));
    const themeDiffers = !sameTheme(await themeXml(await Engine.open(ks.book.bytes)), await themeXml(await Engine.open(base.book.bytes)));
    meta.push({
      file: files[i], label: ks.book.fileName, themeDiffers, drawingsSkipped,
      sheets: added.filter((a) => a.sourceId === ks.id).map((a) => ({ original: a.original, name: a.name, complete: a.complete, objects: objs.find((s) => s.name === a.original)! })),
    });
  }
  const file = `${String(manifest.length + 1).padStart(2, '0')}-${c.name}.xlsx`;
  fs.writeFileSync(path.join(out, 'out', file), bytes);
  manifest.push({ name: c.name, out: `out/${file}`, expected: [res.layout.sheet, ...keptBase, ...expectedCopies], note: c.note ?? '', lint, literalCharts: c.literalCharts ?? false, sources: meta });
  console.log('wrote', file, `${added.length} copies`, lint ? `LINT: ${lint.replace(/\n/g, ' | ')}` : '');
}


/** ob-charts with series references to workbook-level names ([0]!Name) of one plain range: the copy re-points them to the range (series formulas are not compared: the source's read the names) */
async function namesVariant(): Promise<string> {
  const zip = await JSZip.loadAsync(fs.readFileSync(O('ob-charts.xlsx')));
  let wb = await zip.file('xl/workbook.xml')!.async('string');
  wb = wb.replace('</sheets>', '</sheets><definedNames><definedName name="ChtVal">' + 'データ!$B$2:$B$7</definedName><definedName name="ChtName">データ!$B$1</definedName><definedName name="ChtCat">データ!$A$2:$A$7</definedName></definedNames>');
  zip.file('xl/workbook.xml', wb);
  let ch = await zip.file('xl/charts/chart1.xml')!.async('string');
  ch = ch.replace('<c:f>データ!$B$2:$B$7</c:f>', '<c:f>[0]!ChtVal</c:f>').replace('<c:f>データ!$B$1</c:f>', '<c:f>[0]!ChtName</c:f>').replace('<c:f>データ!$A$2:$A$7</c:f>', '<c:f>[0]!ChtCat</c:f>');
  zip.file('xl/charts/chart1.xml', ch);
  const file = path.join(out, 'ob-charts-names.xlsx');
  fs.writeFileSync(file, await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }));
  return file;
}

/** corpus / golden files that own objects (found by looking into the package) */
async function withObjects(dir: string): Promise<string[]> {
  const hit: string[] = [];
  for (const f of fs.readdirSync(dir).filter((x) => /\.xlsx$/i.test(x)).sort()) {
    try {
      const zip = await JSZip.loadAsync(fs.readFileSync(path.join(dir, f)));
      if (Object.keys(zip.files).some((n) => /^xl\/(drawings\/drawing|charts\/chart|tables\/table|comments|pivotTables\/|chartsheets\/)/i.test(n))) hit.push(f);
    } catch { /* not a package */ }
  }
  return hit;
}

const NAMES_VARIANT = await namesVariant();
const OB = ['ob-charts', 'ob-shapes', 'ob-comments', 'ob-table', 'ob-pivot', 'ob-slicer-control'].map((n) => O(`${n}.xlsx`));
const COMBOS: Combo[] = [
  // ---- every Excel-authored fixture alone ----
  ...OB.map((f): Combo => ({ name: path.basename(f, '.xlsx'), base: S('01_学生名簿.xlsx'), others: [f] })),
  { name: 'samples-01-chart-pivot-as-other', base: S('02_面接評価.xlsx'), others: [S('01_学生名簿.xlsx')], drop: false },
  { name: 'ob-all-keep', base: S('02_面接評価.xlsx'), others: OB, drop: false },
  { name: 'ob-charts-literal-data', base: S('01_学生名簿.xlsx'), others: [NAMES_VARIANT], literalCharts: true },
  // ---- the same file twice: names / ids that collide ----
  ...['ob-charts', 'ob-shapes', 'ob-comments', 'ob-table', 'ob-pivot'].map((n): Combo => ({ name: `${n}-x2`, base: S('02_面接評価.xlsx'), others: [O(`${n}.xlsx`), { file: O(`${n}.xlsx`), fileName: `${n}-2.xlsx` }] })),
  { name: 'ob-table-as-base-keep', base: O('ob-table.xlsx'), others: [{ file: O('ob-table.xlsx'), fileName: 'ob-table-2.xlsx' }], drop: false },
  { name: 'ob-pivot-as-base-keep', base: O('ob-pivot.xlsx'), others: [{ file: O('ob-pivot.xlsx'), fileName: 'ob-pivot-2.xlsx' }, O('ob-charts.xlsx')], drop: false },
  { name: 'ob-slicer-as-base-keep', base: O('ob-slicer-control.xlsx'), others: [O('ob-charts.xlsx'), O('ob-comments.xlsx')], drop: false },
  // ---- Strict Open XML Spreadsheet (other namespaces everywhere): as the other file, as the base, strict with strict ----
  { name: 'ob-strict-as-other', base: S('01_学生名簿.xlsx'), others: [O('ob-strict.xlsx')] },
  { name: 'ob-strict-as-base-keep', base: O('ob-strict.xlsx'), others: [O('ob-charts.xlsx'), O('ob-table.xlsx'), O('ob-comments.xlsx'), O('ob-pivot.xlsx')], drop: false },
  { name: 'ob-strict-x-strict', base: O('ob-strict.xlsx'), others: [{ file: O('ob-strict.xlsx'), fileName: 'ob-strict-2.xlsx' }], drop: false },
  // ---- a base with another theme: chart / shape colours must stay those of the source ----
  { name: 'theme-other-base', base: X('ex-theme-other.xlsx'), others: [O('ob-charts.xlsx'), O('ob-shapes.xlsx'), O('ob-table.xlsx')] },
  // ---- after edits of the merged sheet (replayLog) ----
  { name: 'copies-after-edits', base: S('01_学生名簿.xlsx'), others: [O('ob-charts.xlsx'), O('ob-table.xlsx'), O('ob-comments.xlsx')], drop: false, ops: 2 },
  { name: 'join-then-edits', base: S('01_学生名簿.xlsx'), others: [S('02_面接評価.xlsx'), O('ob-pivot.xlsx')], join: 'left', drop: false, ops: 3 },
  ...(process.env.EXCEL_ENGINE === '1' ? [
    { name: 'excel-engine-copies', base: S('01_学生名簿.xlsx'), others: [O('ob-charts.xlsx'), O('ob-pivot.xlsx'), O('ob-table.xlsx')], drop: false, ops: 2, engine: 'excel' } as Combo,
  ] : []),
  // ---- Excel-authored golden fixtures with charts / pictures / shapes / notes / threaded comments / pivots / tables ----
  ...(await withObjects(path.join(root, 'tests/fixtures/golden'))).map((f): Combo => ({ name: `golden-${path.basename(f, '.xlsx')}`, base: S('01_学生名簿.xlsx'), others: [G(f)] })),
  // ---- the foreign corpus files that own objects (Google Sheets, LibreOffice, openpyxl, XlsxWriter, ...) ----
  ...(await withObjects(path.join(root, 'tests/fixtures/foreign/corpus'))).map((f): Combo => ({ name: `corpus-${path.basename(f, '.xlsx')}`, base: S('01_学生名簿.xlsx'), others: [C(f)] })),
];

const only = process.argv[3] ? new RegExp(process.argv[3]) : null;
for (const c of COMBOS) {
  if (only && !only.test(c.name)) continue;
  try {
    await run(c);
  } catch (e) {
    console.log('FAILED', c.name, (e as Error).message);
  }
}
fs.writeFileSync(path.join(out, 'manifest.json'), JSON.stringify(manifest, null, 1));
console.log('manifest:', manifest.length, 'entries');
