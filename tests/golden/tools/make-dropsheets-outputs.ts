/**
 * Writes the single-sheet outputs (PremergeOptions.dropOtherSheets) of the Excel-authored fixtures, the samples and the multi-sheet
 * foreign corpus into a folder, together with the inputs they come from, so that real Excel can be asked (one hidden instance,
 * tests/golden/com/check-dropsheets.ps1) whether they open without repair, hold exactly one sheet, show no #REF! / #NAME? that the
 * input did not have and — where the output has the same layout as its input — the same values after a full recalculation:
 *   npx tsx tests/golden/tools/make-dropsheets-outputs.ts <folder>
 * <folder>/in/<name>.xlsx, <folder>/out/<name>.xlsx, <folder>/manifest.json [{ name, sheet, compare }]
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { analyzeSources, defaultPlan, detectRegion, joinSources, pickDataSheet } from '../../../src/core/keys';
import type { KeySource, PremergeInput } from '../../../src/core/contracts';
import { createEditor } from '../../../src/core/doc';
import { buildPremerged, readBook, replayLog } from '../../../src/core/xlsx';
import { rng, step } from '../../doc/fuzzcore';
import { baseOptions, premergeOne } from '../../xlsx/dropsheets-helpers';
import { source } from '../premerge-scenario';

const here = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(here, '../../..');
const out = path.resolve(process.argv[2] ?? 'dropsheets-out');
fs.mkdirSync(path.join(out, 'in'), { recursive: true });
fs.mkdirSync(path.join(out, 'out'), { recursive: true });

interface Entry { name: string; sheet: string; compare: boolean; note: string }
const manifest: Entry[] = [];
const rd = (...p: string[]) => new Uint8Array(fs.readFileSync(path.join(root, ...p)));

async function load(file: string, bytes: Uint8Array, id: string, sheet?: string): Promise<KeySource> {
  const book = await readBook(bytes, file, { id });
  return { id, book, region: detectRegion(book, sheet ?? pickDataSheet(book)), label: file.replace(/\.[^.]+$/, '') };
}

function save(name: string, input: Uint8Array, output: Uint8Array, sheet: string, compare: boolean, note: string) {
  fs.writeFileSync(path.join(out, 'in', `${name}.xlsx`), input);
  fs.writeFileSync(path.join(out, 'out', `${name}.xlsx`), output);
  manifest.push({ name, sheet, compare, note });
  console.log('wrote', name, compare ? '(values compared)' : '', note);
}

/**
 * The doc's random op sequences (tests/doc/fuzzcore.ts) replayed on a premerged package: what the app exports after the user's edits.
 * `seeds` outputs named <name>-ops<seed>; Excel only has to open them (no repair, one sheet, no new error).
 */
async function withOps(name: string, input: Uint8Array, premerged: Uint8Array, base: KeySource, layout: Parameters<typeof createEditor>[1], seeds: number[]) {
  const book = await readBook(premerged, 'merged.xlsx', { id: 'merged' });
  const snapshot = createEditor(book, layout, { [base.id]: { label: base.label, color: '#2563eb' } }, { headerRow: base.region.headerRow, firstDataRow: base.region.firstDataRow, lastDataRow: base.region.lastDataRow });
  for (const seed of seeds) {
    const r = rng(seed);
    let s = snapshot;
    for (let i = 0; i < 10; i++) {
      const o = step(s, r, false, s.doc.lastDataRow + 2);
      if (o.res.ok) s = o.res.snapshot;
    }
    if (!s.log.length) continue;
    save(`${name}-ops${seed}`, input, await replayLog(premerged, s.log), layout.sheet, false, `${s.log.length} logged ops`);
  }
}

/** the base alone: the output has the layout of the input, so Excel can compare every value */
async function alone(name: string, file: string, bytes: Uint8Array, sheet?: string, compare = true, seeds: number[] = []) {
  const base = await load(file, bytes, 'b', sheet);
  const res = await premergeOne(base, { dropOtherSheets: true });
  save(name, bytes, res.bytes, res.layout.sheet, compare, `detached ${res.report.formulasDetached}, dropped ${(res.report.droppedSheets ?? []).length}`);
  await withOps(name, bytes, res.bytes, base, res.layout, seeds);
}

// ---- samples ------------------------------------------------------------------------------------------------
{
  const samples = path.join(root, 'samples');
  const read = (f: string) => new Uint8Array(fs.readFileSync(path.join(samples, f)));
  await alone('samples-01-alone', '01_学生名簿.xlsx', read('01_学生名簿.xlsx'));
  for (const joinType of ['left', 'full'] as const) {
    const base = await load('01_学生名簿.xlsx', read('01_学生名簿.xlsx'), 's1');
    const others = [await load('02_面接評価.xlsx', read('02_面接評価.xlsx'), 's2'), await load('03_連絡先.csv', read('03_連絡先.csv'), 's3'), await load('05_研修出欠.xlsx', read('05_研修出欠.xlsx'), 's4')];
    const analysis = analyzeSources(base, others);
    const join = joinSources(analysis, { ...defaultPlan(analysis), joinType });
    const input: PremergeInput = { base, sources: others, rows: join.rows, imports: join.defaultImports, keyFill: join.keyFill, options: baseOptions({ dropOtherSheets: true }) };
    const res = await buildPremerged(input);
    save(`samples-01-merged-${joinType}`, read('01_学生名簿.xlsx'), res.bytes, res.layout.sheet, false, `merged with 02 03 05 (${joinType}), detached ${res.report.formulasDetached}`);
    await withOps(`samples-01-merged-${joinType}`, read('01_学生名簿.xlsx'), res.bytes, base, res.layout, [31, 32, 33]);
  }
  await alone('samples-02-alone', '02_面接評価.xlsx', read('02_面接評価.xlsx'));
}

// ---- Excel-authored fixtures ------------------------------------------------------------------------------------
await alone('torture', 'torture.xlsx', rd('tests/fixtures/golden/torture.xlsx'), 'Data', true, [11, 12, 13]);
for (const f of ['g-table.xlsx', 'pm-base-plain.xlsx', 'pm-base-table.xlsx', 'r2-arrays.xlsx', 'r2-table.xlsx']) await alone(f.replace(/\.xlsx$/, ''), f, rd('tests/fixtures/golden', f));
await alone('r3-misc-Tb', 'r3-misc.xlsx', rd('tests/fixtures/golden/r3-misc.xlsx'), 'Tb');
for (const f of ['ds-features.xlsx', 'ds-table.xlsx']) {
  if (fs.existsSync(path.join(root, 'tests/fixtures/dropsheets', f))) await alone(f.replace(/\.xlsx$/, ''), f, rd('tests/fixtures/dropsheets', f), '名簿', true, [21, 22, 23]);
}
// ds-table with a source: imported columns inside / next to the table, appended rows, the table extended
{
  const f = path.join(root, 'tests/fixtures/dropsheets/ds-table.xlsx');
  if (fs.existsSync(f)) {
    const bytes = rd('tests/fixtures/dropsheets/ds-table.xlsx');
    const base = await load('ds-table.xlsx', bytes, 'b', '名簿');
    const src = source('src', 'pm-src.xlsx', { sheet: '回答', headerRow: 1, firstDataRow: 2, lastDataRow: 6, firstCol: 1, lastCol: 9 }, 'pm-src');
    const rows = [];
    for (let r = base.region.firstDataRow; r <= base.region.lastDataRow; r++) rows.push({ baseRow: r, matches: { src: r - 1 <= 5 ? { row: r, via: 'ID', flags: [] as never[] } : null } });
    rows.push({ baseRow: null as number | null, appendedFrom: { sourceId: 'src', row: 4 }, matches: { src: { row: 4, via: 'ID', flags: [] as never[] } } } as never);
    const res = await buildPremerged({
      base, sources: [src], rows: rows as PremergeInput['rows'], imports: [3, 4, 5].map((col) => ({ sourceId: 'src', col })), keyFill: { src: [{ baseCol: 1, sourceCol: 1 }, { baseCol: 2, sourceCol: 9 }] },
      options: baseOptions({ dropOtherSheets: true }),
    });
    save('ds-table-merged', bytes, res.bytes, res.layout.sheet, false, `imports + appended row, detached ${res.report.formulasDetached}`);
  }
}

// ---- the foreign corpus: every file with several sheets ------------------------------------------------------------
{
  const dir = path.join(root, 'tests/fixtures/foreign/corpus');
  for (const f of fs.readdirSync(dir).filter((x) => /\.xlsx$/.test(x)).sort()) {
    try {
      const bytes = new Uint8Array(fs.readFileSync(path.join(dir, f)));
      const book = await readBook(bytes, f, { id: 'b' });
      if (book.sheets.length < 2 || book.fidelity !== 'full') continue;
      await alone(f.replace(/\.xlsx$/, ''), f, bytes, undefined, false);
    } catch (e) {
      console.log('skipped', f, (e as Error).message);
    }
  }
}

fs.writeFileSync(path.join(out, 'manifest.json'), JSON.stringify(manifest, null, 1), 'utf8');
console.log(`${manifest.length} pairs in ${out}`);
