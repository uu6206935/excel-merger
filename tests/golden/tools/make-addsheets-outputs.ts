/**
 * Writes premerged outputs WITH copied sheets (PremergeOptions.includeSourceSheets) of >= 40 combinations — the samples in several card orders,
 * the Excel-authored fixtures of tests/fixtures/addsheets, a sample of the foreign corpus (Google Sheets / LibreOffice / Mac 1904 / Strict /
 * openpyxl / XlsxWriter / macro-enabled / external links ...), three files with the same name, a long file name, a base whose data sheet
 * is not its first sheet, and exports after random edits (replayLog) — into a folder, together with a manifest, so that real Excel can be
 * asked (one hidden instance, tests/golden/com/check-addsheets.ps1):
 *   npx tsx tests/golden/tools/make-addsheets-outputs.ts <folder> [<regex of combination names>]
 *   (EXCEL_ENGINE=1 adds two combinations whose edits are replayed by the app's Excel engine: one hidden Excel start each)
 * <folder>/out/<name>.xlsx, <folder>/manifest.json
 * [{ name, out, merged, expected: [sheet names in order], drop, sources: [{ file, label, compare, sheets: [{ original, name, complete }] }] }]
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { analyzeSources, defaultPlan, detectRegion, joinSources, pickDataSheet } from '../../../src/core/keys';
import type { KeySource, PremergeInput } from '../../../src/core/contracts';
import { createEditor } from '../../../src/core/doc';
import { buildPremerged, readBook, replayLog } from '../../../src/core/xlsx';
import { rng, step } from '../../doc/fuzzcore';
import { plannedNames } from '../../xlsx/addsheets-helpers';
import { runRunnerJob } from '../../oracle/consistency/excel';

const here = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(here, '../../..');
const out = path.resolve(process.argv[2] ?? 'addsheets-out');
fs.mkdirSync(path.join(out, 'out'), { recursive: true });

const SAMPLES = path.join(root, 'samples');
const CORPUS = path.join(root, 'tests/fixtures/foreign/corpus');
const FX = path.join(root, 'tests/fixtures/addsheets');
const S = (n: string) => path.join(SAMPLES, n);
const C = (n: string) => path.join(CORPUS, n);
const X = (n: string) => path.join(FX, n);

interface Entry {
  name: string;
  out: string;
  merged: string;
  expected: string[];
  drop: boolean;
  note: string;
  sources: Array<{ file: string; label: string; compare: boolean; lenientValues: boolean; date1904: boolean; sheets: Array<{ original: string; name: string; complete: boolean }> }>;
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
  /** a premerge through the real join (imports, join full) instead of "copies only" */
  join?: 'left' | 'full';
  /** random edits (doc fuzz seeds), replayed with the built-in engine (or the Excel engine, see `engine`) */
  ops?: number;
  /** replay the edits with the app's Excel engine instead of replayLog */
  engine?: 'excel';
  note?: string;
}

async function run(c: Combo): Promise<void> {
  const base = await load(c.base, 'b');
  const others: KeySource[] = [];
  const meta: Entry['sources'] = [];
  for (const [i, o] of c.others.entries()) {
    const file = typeof o === 'string' ? o : o.file;
    const ks = await load(file, `s${i + 2}`, typeof o === 'string' ? path.basename(o) : o.fileName);
    others.push(ks);
    // formulas into other workbooks / tables (phase 1 of the copier does not copy tables) are copied as their stored values — a cell without a stored
    // result stays empty — while Excel recomputes the source's own: value differences of such sources are counted apart (lenientValueDiffs)
    meta.push({ file, label: ks.book.fileName, compare: !/\.(csv|tsv)$/i.test(file), lenientValues: ks.book.features.externalLinks > 0 || ks.book.features.tables > 0, date1904: ks.book.date1904, sheets: [] });
  }
  const drop = c.drop ?? true;
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
      // the app's "Excel本体" engine (server/excel/excel-runner.ps1) replays the same log in its own hidden Excel: one Excel start per output
      const work = path.join(out, '_work');
      fs.mkdirSync(work, { recursive: true });
      const input = path.join(work, `${c.name}-in.xlsx`), output = path.join(work, `${c.name}-out.xlsx`);
      fs.writeFileSync(input, bytes);
      const outcome = await runRunnerJob(input, snap.log, output, res.layout.sheet, { tempRoot: work, cacheDir: path.join(work, 'cache') });
      if (!outcome.ok || !fs.existsSync(output)) throw new Error(`${c.name}: Excel engine failed: ${outcome.code} ${outcome.message}`);
      bytes = new Uint8Array(fs.readFileSync(output));
      fs.rmSync(input, { force: true });
      fs.rmSync(output, { force: true });
    } else {
      bytes = await replayLog(bytes, snap.log);
    }
  }
  const added = res.layout.addedSheets ?? [];
  for (const m of meta) {
    const ks = others.find((k) => k.book.fileName === m.label)!;
    void ks;
  }
  // the expected names come from the naming rule alone (planSheetNames), not from the copier
  const baseBook = base.book;
  const mergedName = res.layout.sheet;
  const keptBase = drop ? [] : baseBook.sheets.map((s) => s.name).filter((n) => n !== mergedName);
  const files = await Promise.all(others.map(async (o) => ({ fileName: o.book.fileName, sheets: (await readBook(o.book.bytes, o.book.fileName)).sheets.map((s) => s.name) })));
  const expectedCopies = plannedNames([mergedName, ...keptBase], files);
  const expected = [mergedName, ...keptBase, ...expectedCopies];
  const gotCopies = added.map((a) => a.name);
  if (JSON.stringify(gotCopies) !== JSON.stringify(expectedCopies)) throw new Error(`${c.name}: layout says ${gotCopies.join('|')}, rule says ${expectedCopies.join('|')}`);
  for (const a of added) meta[Number(a.sourceId.slice(1)) - 2].sheets.push({ original: a.original, name: a.name, complete: a.complete });
  const file = `${String(manifest.length + 1).padStart(2, '0')}-${c.name}.${/\.xlsm$/i.test(c.base) ? 'xlsm' : 'xlsx'}`;
  fs.writeFileSync(path.join(out, 'out', file), bytes);
  manifest.push({ name: c.name, out: `out/${file}`, merged: mergedName, expected, drop, note: c.note ?? '', sources: meta });
  console.log('wrote', file, `${expected.length} sheets`, c.ops ? '(after edits)' : '');
}

const COMBOS: Combo[] = [
  // ---- the app's samples (card orders) ----
  { name: 'samples-01-all-keep', base: S('01_学生名簿.xlsx'), others: [S('02_面接評価.xlsx'), S('03_連絡先.csv'), S('04_適性検査.xls'), S('05_研修出欠.xlsx')], drop: false },
  { name: 'samples-01-all-drop', base: S('01_学生名簿.xlsx'), others: [S('02_面接評価.xlsx'), S('03_連絡先.csv'), S('04_適性検査.xls'), S('05_研修出欠.xlsx')] },
  { name: 'samples-01-reversed', base: S('01_学生名簿.xlsx'), others: [S('05_研修出欠.xlsx'), S('04_適性検査.xls'), S('03_連絡先.csv'), S('02_面接評価.xlsx')] },
  { name: 'samples-01-shuffled-keep', base: S('01_学生名簿.xlsx'), others: [S('03_連絡先.csv'), S('05_研修出欠.xlsx'), S('02_面接評価.xlsx'), S('04_適性検査.xls')], drop: false },
  { name: 'samples-02-base', base: S('02_面接評価.xlsx'), others: [S('01_学生名簿.xlsx'), S('05_研修出欠.xlsx')], drop: false },
  { name: 'samples-05-base', base: S('05_研修出欠.xlsx'), others: [S('01_学生名簿.xlsx')] },
  { name: 'samples-03-csv-base', base: S('03_連絡先.csv'), others: [S('01_学生名簿.xlsx')] },
  { name: 'samples-04-xls-base', base: S('04_適性検査.xls'), others: [S('01_学生名簿.xlsx')] },
  // ---- through the real join + after random edits (what the app exports) ----
  { name: 'join-left-imports', base: S('01_学生名簿.xlsx'), others: [S('02_面接評価.xlsx'), S('03_連絡先.csv'), S('05_研修出欠.xlsx')], join: 'left', drop: false },
  { name: 'join-full-imports', base: S('01_学生名簿.xlsx'), others: [S('02_面接評価.xlsx'), S('05_研修出欠.xlsx')], join: 'full' },
  { name: 'join-left-edits-1', base: S('01_学生名簿.xlsx'), others: [S('02_面接評価.xlsx'), S('03_連絡先.csv'), S('05_研修出欠.xlsx')], join: 'left', drop: false, ops: 1 },
  { name: 'join-full-edits-2', base: S('01_学生名簿.xlsx'), others: [S('02_面接評価.xlsx'), S('05_研修出欠.xlsx')], join: 'full', ops: 2 },
  { name: 'join-left-edits-3', base: S('02_面接評価.xlsx'), others: [S('01_学生名簿.xlsx'), S('05_研修出欠.xlsx')], join: 'left', ops: 3 },
  // ---- the same edits through the app's Excel engine (only with EXCEL_ENGINE=1: one Excel start each) ----
  ...(process.env.EXCEL_ENGINE === '1' ? [
    { name: 'excel-engine-join-left', base: S('01_学生名簿.xlsx'), others: [S('02_面接評価.xlsx'), S('05_研修出欠.xlsx')], join: 'left', drop: false, ops: 2, engine: 'excel' } as Combo,
    { name: 'excel-engine-ex-formats', base: S('01_学生名簿.xlsx'), others: [X('ex-formats.xlsx'), X('ex-two-sheets.xlsx')], join: 'left', ops: 3, engine: 'excel' } as Combo,
  ] : []),
  // ---- Excel-authored fixtures ----
  { name: 'ex-formats', base: S('01_学生名簿.xlsx'), others: [X('ex-formats.xlsx')] },
  { name: 'ex-theme-other', base: S('01_学生名簿.xlsx'), others: [X('ex-theme-other.xlsx')] },
  { name: 'ex-two-sheets', base: S('01_学生名簿.xlsx'), others: [X('ex-two-sheets.xlsx')] },
  { name: 'ex-1904', base: S('01_学生名簿.xlsx'), others: [X('ex-1904.xlsx')] },
  { name: 'ex-all-four', base: S('01_学生名簿.xlsx'), others: [X('ex-formats.xlsx'), X('ex-theme-other.xlsx'), X('ex-two-sheets.xlsx'), X('ex-1904.xlsx')], drop: false },
  { name: 'ex-formats-as-base', base: X('ex-formats.xlsx'), others: [X('ex-two-sheets.xlsx'), S('02_面接評価.xlsx')], drop: false },
  { name: 'ex-1904-as-base', base: X('ex-1904.xlsx'), others: [X('ex-formats.xlsx'), S('05_研修出欠.xlsx')], drop: false },
  // ---- naming ----
  { name: 'same-name-x3', base: S('02_面接評価.xlsx'), others: [S('01_学生名簿.xlsx'), { file: S('01_学生名簿.xlsx'), fileName: '01_学生名簿.xlsx' }, { file: S('01_学生名簿.xlsx'), fileName: 'C:\\Users\\x\\01_学生名簿.xlsx' }], drop: false },
  { name: 'long-file-name', base: S('01_学生名簿.xlsx'), others: [{ file: S('02_面接評価.xlsx'), fileName: '営業部門の月次実績報告書（2026年度第3四半期確定版）.xlsx' }] },
  { name: 'base-data-sheet-not-first', base: C('hc-17-sheetnames-hidden-first.xlsx'), others: [S('02_面接評価.xlsx')], drop: false },
  // ---- foreign corpus ----
  ...[
    'hc-01-google-sheets', 'hc-02-libreoffice', 'hc-05-mac-excel-2011-1904', 'hc-06-excel2007-ja-numfmts', 'hc-08-strict-ooxml', 'hc-10-shared-formulas',
    'hc-17-sheetnames-hidden-first', 'hc-19-merged-two-row-header', 'hc-29-formula-variety', 'hc-35-names-cf-dv', 'hc-36-theme2013-colours', 'hc-37-external-link',
    'js-04-exceljs-basic', 'js-10-write-excel-file', 'py-01-openpyxl-styled', 'py-06-openpyxl-date1904', 'py-12-openpyxl-protected', 'py-14-openpyxl-formats-errors',
    'py-17-xlsxwriter-arrays', 'py-04-openpyxl-table',
  ].map((n): Combo => ({ name: `corpus-${n}`, base: S('01_学生名簿.xlsx'), others: [C(`${n}.xlsx`)] })),
  // macro-enabled workbooks authored by Excel itself (hc-31 from the corpus is not openable by Excel even as an input, so it is only unit tested)
  { name: 'ex-macro-as-source', base: S('01_学生名簿.xlsx'), others: [X('ex-macro.xlsm')] },
  { name: 'ex-macro-as-base', base: X('ex-macro.xlsm'), others: [S('02_面接評価.xlsx')], drop: false },
  { name: 'ex-macro-as-base-drop', base: X('ex-macro.xlsm'), others: [X('ex-formats.xlsx'), S('05_研修出欠.xlsx')] },
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
void buildPremerged;
