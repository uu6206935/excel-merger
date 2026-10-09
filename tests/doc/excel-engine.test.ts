/**
 * The doc's logs against the REAL Excel (skipped unless EXCEL_TESTS=1; own hidden Excel instance, no clipboard):
 * a workbook authored by Excel itself is read back, edited through the doc (deleteColumns with 「値に変換」 over array /
 * spill masters and their readers on other sheets), and the op log is replayed by the Excel runner. Excel's result must
 * be free of #REF! and equal the doc (formulas + cached values of every cell the doc knows). The entries are ordered
 * readers-of-a-spill first (valid for any engine; the runner replays with manual calculation, so Excel itself
 * accepts either order — measured with the reverse order as a control).
 */
import fs from 'node:fs';
import path from 'node:path';
import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { numToCol } from '../../src/core/formula/excelRef';
import { logToRunnerJob } from '../../src/core/export/comJob';
import { createEditor, deleteColumns } from '../../src/core/doc';
import { readBook } from '../../src/core/xlsx';
import type { EditorSnapshot, LoggedOp } from '../../src/core/types';
import { runRunner } from '../../server/excel/run-job.mjs';
import { EXCEL_TESTS, cellAt, dumpBook, makeFixture, tempDir, type Dump, type FixtureSpec, type SpecCell } from '../com/lib';

const MAIN = '名簿';
const SUM = '集計';
const RUNNER = path.resolve(__dirname, '..', '..', 'server', 'excel', 'excel-runner.ps1');

function spec(): FixtureSpec {
  const cells: SpecCell[] = [['A1', 's', '値'], ['B1', 's', '備考'], ['D1', 's', '件数']];
  for (let r = 2; r <= 6; r++) cells.push([`A${r}`, 'n', r - 1], [`B${r}`, 's', `メモ${r}`]);
  cells.push(['C2', 'd', '=A2:A6&"x"']); // spill C2:C6 (reads column A)
  cells.push(['D2', 'f', '=COUNTA(C2#)']); // reader on the same sheet
  cells.push(['E2:E6', 'a', '=A2:A6*10']); // legacy array E2:E6 (reads column A) -- a one-column range
  const sum: SpecCell[] = [
    ['A1', 's', '集計'],
    ['B2', 'd', `=${MAIN}!C2#`], // a spill that reads the spill (itself a master)
    ['C2', 'f', `=COUNTA(${MAIN}!C2#)`], // a plain reader
    ['A8', 'f', `=SUM(${MAIN}!A2:A6)`], // breaks with column A as well
  ];
  return { sheets: [{ name: MAIN, cells }, { name: SUM, cells: sum }] };
}

async function runJob(dir: string, input: string, log: LoggedOp[], label: string) {
  const job = logToRunnerJob(log, { format: 'xlsx', sheet: MAIN });
  const out = path.join(dir, `${label}.xlsx`);
  const jobFile = path.join(dir, `${label}.job.json`);
  fs.writeFileSync(jobFile, JSON.stringify({ ...job, input, output: out, workDir: path.join(dir, `w-${label}`) }));
  const r = await runRunner({ runner: RUNNER, jobFile, env: { XLRUNNER_CACHE: path.join(dir, 'cache') }, onEvent: () => undefined });
  return { r, out };
}

const errors = (d: Dump) => d.sheets.flatMap((s) => s.cells.filter((c) => typeof c.v === 'string' && c.v.startsWith('ERR:')).map((c) => `${s.name}!${numToCol(c.c)}${c.r} ${c.v}`));
const valueAt = (d: Dump, sheet: string, r: number, c: number) => cellAt(d, sheet, r, c)?.v;

describe.skipIf(!EXCEL_TESTS)('doc log vs real Excel: converting array / spill masters and their readers', () => {
  let dir: string;
  let fixture: string;
  let before: Dump;
  let snap: EditorSnapshot;
  let excelDump: Dump;

  beforeAll(async () => {
    dir = tempDir('docspill');
    fixture = path.join(dir, 'fixture.xlsx');
    await makeFixture(spec(), fixture);
    before = await dumpBook(fixture);
    const book = await readBook(new Uint8Array(fs.readFileSync(fixture)), 'fixture.xlsx');
    const s0 = createEditor(book, { sheet: MAIN, insertAt: 0, imported: [], appendAt: null, appended: [] }, {}, { headerRow: 1, firstDataRow: 2, lastDataRow: 6 });
    const r = deleteColumns(s0, [s0.doc.columns[0].id], { convertDependents: true });
    if (!r.ok) throw new Error(r.reason);
    snap = r.snapshot;
  }, 240_000);
  afterAll(() => {
    if (dir && process.env.KEEP_COM_OUTPUT !== '1') fs.rmSync(dir, { recursive: true, force: true });
  });

  it('fixture sanity: Excel authored the spills, the CSE array and the reader', () => {
    expect(cellAt(before, MAIN, 2, 3)?.spill).toBe('C2:C6');
    expect(cellAt(before, MAIN, 2, 5)?.arr).toBe('E2:E6');
    expect(cellAt(before, SUM, 2, 2)?.spill).toBe('B2:B6');
  });

  it('the log converts the readers of the spill first, then the masters', () => {
    const cells = snap.log.filter((e) => e.kind === 'cells');
    const order = cells.map((e) => (e.kind === 'cells' ? `${e.sheet}:${e.writes.map((w) => `${numToCol(w.col)}${w.row}`).join(',')}` : ''));
    // 名簿!D2 (COUNTA(C2#)) and 集計!B2 / C2 read C2#: they come before the master 名簿!C2
    const iMaster = order.findIndex((o) => o.startsWith(`${MAIN}:`) && o.split(':')[1].split(',').includes('C2'));
    const iReaderSum = order.findIndex((o) => o.startsWith(`${SUM}:`) && o.split(':')[1].split(',').includes('B2'));
    expect(iMaster).toBeGreaterThan(-1);
    expect(iReaderSum).toBeGreaterThan(-1);
    expect(iReaderSum).toBeLessThan(iMaster);
  });

  it('Excel engine: no #REF!, every converted array keeps its values, and the result equals the doc', async () => {
    const { r, out } = await runJob(dir, fixture, snap.log, 'excel');
    expect(r.ok, JSON.stringify(r)).toBe(true);
    excelDump = await dumpBook(out);
    expect(errors(excelDump)).toEqual([]);
    // 名簿 after deleting column A: C -> B (spill master -> plain values), E -> D (CSE array -> values)
    const main = excelDump.sheets.find((s) => s.name === MAIN)!;
    const col = (c: number) => [2, 3, 4, 5, 6].map((row) => valueAt(excelDump, MAIN, row, c));
    expect(col(2)).toEqual(['1x', '2x', '3x', '4x', '5x']);
    expect(col(4)).toEqual([10, 20, 30, 40, 50]);
    expect(main.cells.filter((c) => c.hf && c.c <= 4).map((c) => `${numToCol(c.c)}${c.r}`)).toEqual([]);
    // 集計: the spill that read 名簿!C2# is values now (all five rows kept), the plain reader too
    expect([2, 3, 4, 5, 6].map((row) => valueAt(excelDump, SUM, row, 2))).toEqual(['1x', '2x', '3x', '4x', '5x']);
    expect(valueAt(excelDump, SUM, 2, 3)).toBe(5);
    expect(cellAt(excelDump, SUM, 2, 2)?.hf).toBe(false);
    // the doc's view of the merged sheet equals what Excel saved (values and formulas of every cell)
    snap.doc.columns.forEach((dc, i) => {
      dc.cells.forEach((cell, ri) => {
        if (!cell) return;
        const x = cellAt(excelDump, MAIN, ri + 1, i + 1);
        const where = `${MAIN}!${numToCol(i + 1)}${ri + 1}`;
        if (cell.f === undefined) {
          if (cell.v === null || cell.v === '') expect(x?.hf ?? false, where).toBe(false);
          else expect(x?.v, where).toEqual(cell.t === 'b' ? !!cell.v : cell.v);
        } else expect(x?.hf, where).toBe(true);
      });
    });
  }, 300_000);
});
