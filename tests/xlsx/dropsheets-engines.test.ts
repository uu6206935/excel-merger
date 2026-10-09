/**
 * Both export engines on a SINGLE-SHEET premerged package (PremergeOptions.dropOtherSheets): the same op log is replayed by the JS engine
 * (replayLog) and by the real Excel through the runner (skipped unless EXCEL_TESTS=1; the runner starts its own hidden Excel). Excel must
 * run the job, hold exactly one sheet, show no error cell the package did not have, and the two outputs must agree (tools/lib/package-diff.mjs,
 * the normalized package diff that tests/com/parity.test.ts uses). tests/golden/RESULTS.md records the run.
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { afterAll, describe, expect, it } from 'vitest';
import { logToRunnerJob } from '../../src/core/export/comJob';
import { analyzeSources, defaultPlan, detectRegion, joinSources, pickDataSheet } from '../../src/core/keys';
import { buildPremerged, readBook, replayLog } from '../../src/core/xlsx';
import type { KeySource, PremergeInput } from '../../src/core/contracts';
import type { LoggedOp } from '../../src/core/types';
import { runRunner } from '../../server/excel/run-job.mjs';
import { comparePackages } from '../../tools/lib/package-diff.mjs';
import { EXCEL_TESTS, dumpBook, tempDir } from '../com/lib';
import { SAMPLES_DIR, samplesAvailable } from '../doc/sampleMerge';
import { baseOptions, baseSource, lintSingleSheet, lintText, premergeOne, readFixture } from './dropsheets-helpers';

const here = path.dirname(fileURLToPath(import.meta.url));
const RUNNER = path.resolve(here, '../../server/excel/excel-runner.ps1');
const DS_FIX = path.resolve(here, '../fixtures/dropsheets');

let dir: string | null = null;
afterAll(() => {
  if (dir && process.env.KEEP_COM_OUTPUT !== '1') fs.rmSync(dir, { recursive: true, force: true });
});

async function bothEngines(label: string, premerged: Uint8Array, sheet: string, log: LoggedOp[]) {
  dir ??= tempDir('dropsheets');
  const input = path.join(dir, `${label}.xlsx`);
  fs.writeFileSync(input, premerged);
  const job = logToRunnerJob(log, { format: 'xlsx', sheet });
  const out = path.join(dir, `${label}.excel.xlsx`);
  const jobFile = path.join(dir, `${label}.job.json`);
  fs.writeFileSync(jobFile, JSON.stringify({ ...job, input, output: out, workDir: path.join(dir, `w-${label}`) }));
  const events: Array<Record<string, unknown>> = [];
  const r = await runRunner({ runner: RUNNER, jobFile, env: { XLRUNNER_CACHE: path.join(dir, 'cache') }, onEvent: (e: Record<string, unknown>) => events.push(e) });
  const js = await replayLog(premerged, log);
  fs.writeFileSync(path.join(dir, `${label}.js.xlsx`), js);
  return { r, events, excel: r.ok ? new Uint8Array(fs.readFileSync(out)) : null, js, out };
}

/**
 * Differences of the two engines that are understood (not caused by dropOtherSheets; [regex over "[group] sheet: detail", why]):
 *  - Excel keeps the text of a hyperlink's `location` (a link to a cell of the sheet itself) when columns move, replayLog moves it with the
 *    cell like a reference (engine.adjustTailFormulas, tests/xlsx/patch-features.test.ts "locations are formulas"): the doc keeps the text, too;
 *  - Excel merges validations with identical settings that become adjacent into one rule (D2:E9 G2:G9), replayLog keeps them apart.
 */
const KNOWN: Array<[RegExp, string]> = [
  [/\[hyperlinks\]/, 'hyperlink location text is not adjusted by Excel'],
  [/\[(validations|dv)\]|入力規則/, 'Excel merges validations with the same settings'],
];
const realDiffs = (rows: Array<{ ok: boolean; group?: string; sheet?: string; detail?: string }>) =>
  rows.filter((x) => !x.ok).map((x) => `[${x.group}] ${x.sheet}: ${x.detail}`).filter((d) => !KNOWN.some(([re]) => re.test(d)));

const errorsOf = (d: Awaited<ReturnType<typeof dumpBook>>) => d.sheets.flatMap((s) => s.cells.filter((c) => typeof c.v === 'string' && c.v.startsWith('ERR:')).map((c) => `${s.name}!${c.r},${c.c} ${c.v}`));

describe.skipIf(!EXCEL_TESTS)('single-sheet premerged package: Excel runner vs JS engine (real Excel, one runner job per file)', () => {
  it('ds-features.xlsx (Excel-authored, every feature): insert / move / row insert / cell writes / width', async () => {
    const bytes = readFixture(DS_FIX, 'ds-features.xlsx');
    const base = await baseSource({ bytes, fileName: 'ds-features.xlsx', sheet: '名簿' });
    const res = await premergeOne(base);
    const log: LoggedOp[] = [
      { kind: 'struct', sheet: '名簿', op: { type: 'insert', axis: 'col', at: 3, count: 1 } },
      { kind: 'struct', sheet: '名簿', op: { type: 'move', axis: 'col', from: 11, count: 2, to: 2 } },
      { kind: 'struct', sheet: '名簿', op: { type: 'insert', axis: 'row', at: 12, count: 1 } },
      { kind: 'cells', sheet: '名簿', reason: 'edit', writes: [{ row: 3, col: 3, cell: { t: 'n', v: 77 } }, { row: 12, col: 1, cell: { t: 's', v: '追加' } }] },
      { kind: 'colWidth', sheet: '名簿', col: 2, width: 18.625 },
    ];
    const { r, events, excel, js } = await bothEngines('ds-features', res.bytes, '名簿', log);
    expect(r.ok, JSON.stringify(r)).toBe(true);
    expect(events.filter((e) => e.type === 'warn' && e.code !== 'MACROS_DROPPED').map((e) => `${e.code}: ${e.message}`)).toEqual([]);
    expect(lintText(await lintSingleSheet(js, ['設定', 'マスタ', '売上', '集計', 'ピボット']))).toBe('');
    expect((await readBook(excel!, 'excel.xlsx')).sheets.map((s) => s.name)).toEqual(['名簿']);
    const { rows } = await comparePackages(excel!, js);
    const real = realDiffs(rows);
    if (real.length) console.warn(['[dropsheets] ds-features: Excel runner vs replayLog', ...real].join('\n'));
    expect(real).toEqual([]);
  }, 420_000);

  it.skipIf(!samplesAvailable())('samples/01 merged with 02 03 05 (imported columns, appended rows): the same log', async () => {
    const load = async (name: string, id: string): Promise<KeySource> => {
      const book = await readBook(new Uint8Array(fs.readFileSync(path.join(SAMPLES_DIR, name))), name, { id });
      return { id, book, region: detectRegion(book, pickDataSheet(book)), label: name.replace(/\.[^.]+$/, '') };
    };
    const base = await load('01_学生名簿.xlsx', 's1');
    const others = [await load('02_面接評価.xlsx', 's2'), await load('03_連絡先.csv', 's3'), await load('05_研修出欠.xlsx', 's4')];
    const analysis = analyzeSources(base, others);
    const join = joinSources(analysis, { ...defaultPlan(analysis), joinType: 'full' });
    const input: PremergeInput = { base, sources: others, rows: join.rows, imports: join.defaultImports, keyFill: join.keyFill, options: baseOptions({ dropOtherSheets: true }) };
    const res = await buildPremerged(input);
    expect(res.report.droppedSheets).toEqual(['集計', 'ピボット', 'マスタ']);
    // The title row is merged over A:K (Excel refuses a move into the middle of it) and column L is the hidden helper column: everything happens
    // right of M. (An insert right of a hidden column is NOT done here: replayLog copies `hidden` from the left column into the new columns,
    // Excel does not — a deviation of the cols model that has nothing to do with dropOtherSheets, see tests/golden/RESULTS.md.)
    const log: LoggedOp[] = [
      { kind: 'struct', sheet: '名簿', op: { type: 'move', axis: 'col', from: 20, count: 2, to: 14 } },
      { kind: 'struct', sheet: '名簿', op: { type: 'insert', axis: 'col', at: 15, count: 2 } },
      { kind: 'struct', sheet: '名簿', op: { type: 'move', axis: 'col', from: 19, count: 1, to: 15 } },
      { kind: 'colWidth', sheet: '名簿', col: 16, width: 20.625 },
    ];
    const { r, events, excel, js } = await bothEngines('samples-01', res.bytes, '名簿', log);
    expect(r.ok, JSON.stringify(r)).toBe(true);
    expect(events.filter((e) => e.type === 'warn' && e.code !== 'MACROS_DROPPED').map((e) => `${e.code}: ${e.message}`)).toEqual([]);
    expect(lintText(await lintSingleSheet(js, ['集計', 'ピボット', 'マスタ']))).toBe('');
    expect((await readBook(excel!, 'excel.xlsx')).sheets.map((s) => s.name)).toEqual(['名簿']);
    expect(errorsOf(await dumpBook(path.join(dir!, 'samples-01.excel.xlsx')))).toEqual([]);
    const { rows } = await comparePackages(excel!, js);
    const real = realDiffs(rows);
    if (real.length) console.warn(['[dropsheets] samples-01: Excel runner vs replayLog', ...real].join('\n'));
    expect(real).toEqual([]);
  }, 420_000);
});
