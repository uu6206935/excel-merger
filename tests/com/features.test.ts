/**
 * Round-2 runner features against the REAL Excel (skipped unless EXCEL_TESTS=1) with detailed assertions (the
 * JS-vs-Excel comparison of the same log is tests/com/invariant.test.ts, scenario "r2"):
 *   unmerge (+ 「選択範囲内で中央」) followed by a move across the former title, cell writes with numFmt, dynamic-array
 *   (Formula2) and legacy (Formula) formulas, toValues over a legacy array (CSE) and over a spill, a delete of the
 *   column those arrays referenced, a move inside an Excel table (Excel's Cut + Insert), inserts, stored widths.
 * The fixture is authored by Excel itself (tests/com/ps/make-fixture.ps1).
 */
import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import fs from 'node:fs';
import path from 'node:path';
import { numToCol } from '../../src/core/formula/excelRef';
import { logToRunnerJob } from '../../src/core/export/comJob';
import type { LoggedOp } from '../../src/core/types';
import { runRunner, type RunnerEvent, type RunnerResult } from '../../server/excel/run-job.mjs';
import { EXCEL_TESTS, cellAt, dumpBook, makeFixture, tempDir, type Dump } from './lib';
import { MAIN, SUM, TSHEET, buildR2Log, r2Spec } from './scenarios';

const RUNNER = path.resolve(__dirname, '..', '..', 'server', 'excel', 'excel-runner.ps1');

async function runJob(dir: string, input: string, log: LoggedOp[], label: string): Promise<{ r: RunnerResult; out: string; events: RunnerEvent[] }> {
  const job = logToRunnerJob(log, { format: 'xlsx', sheet: MAIN });
  const out = path.join(dir, `${label}.xlsx`);
  const jobFile = path.join(dir, `${label}.job.json`);
  fs.writeFileSync(jobFile, JSON.stringify({ ...job, input, output: out, workDir: path.join(dir, `w-${label}`) }));
  const events: RunnerEvent[] = [];
  const r = await runRunner({ runner: RUNNER, jobFile, env: { XLRUNNER_CACHE: path.join(dir, 'cache') }, onEvent: (e) => events.push(e) });
  return { r, out, events };
}

const sheetOf = (d: Dump, name: string) => d.sheets.find((s) => s.name === name)!;
const byHeader = (d: Dump, sheet: string, row: number) => {
  const m = new Map<string, number>();
  for (const c of sheetOf(d, sheet).cells) if (c.r === row && typeof c.v === 'string') m.set(c.v, c.c);
  return m;
};

describe.skipIf(!EXCEL_TESTS)('round 2 runner ops with the real Excel', () => {
  let dir: string;
  let fixture: string;
  let before: Dump;
  let built: ReturnType<typeof buildR2Log>;

  beforeAll(async () => {
    dir = tempDir('r2');
    fixture = path.join(dir, 'fixture.xlsx');
    await makeFixture(r2Spec(), fixture);
    before = await dumpBook(fixture);
    built = buildR2Log(before);
  }, 240_000);
  afterAll(() => {
    if (dir && process.env.KEEP_COM_OUTPUT !== '1') fs.rmSync(dir, { recursive: true, force: true });
    else if (dir) console.log(`[tests/com] outputs kept in ${dir}`);
  });

  it('fixture sanity: Excel authored the merged title, the CSE array, the spill, the table', () => {
    const m = sheetOf(before, MAIN);
    expect(m.merges).toEqual(['A1:F1']);
    expect(cellAt(before, MAIN, 3, 10)?.arr).toBe('J3:J12');
    expect(cellAt(before, MAIN, 3, 11)?.spill).toBe('K3:K12');
    expect(sheetOf(before, TSHEET).tables[0]).toMatchObject({ name: 'Tbl', range: 'A1:E6', columns: ['コード', '品名', '単価', '数量', '金額'] });
    expect(cellAt(before, SUM, 7, 2)?.f).toBe('=SUM(Tbl[金額])');
  });

  it('refuses the move across the merged title when the log has no unmerge first', async () => {
    const { r } = await runJob(dir, fixture, [{ kind: 'struct', sheet: MAIN, op: { type: 'move', axis: 'col', from: 3, count: 1, to: 7 } }], 'refuse');
    expect(r.ok).toBe(false);
    expect(r.code).toBe('REFUSED_MERGE');
    expect(r.message).toMatch(/結合/);
  }, 240_000);

  it('runner: unmerge + center across, numFmt / Formula2 / Formula writes, toValues over a CSE array and a spill, table move, widths', async () => {
    const { r, out, events } = await runJob(dir, fixture, built.log, 'excel');
    expect(r.ok, JSON.stringify(r)).toBe(true);
    const d = await dumpBook(out);
    const { b } = built;
    const col = (orig: number) => b.col(orig);
    // unmerge + 「選択範囲内で中央」 on every cell of the former title (the cells moved with their columns)
    expect(sheetOf(d, MAIN).merges).toEqual([]);
    const title = sheetOf(d, MAIN).cells.find((c) => c.v === '2026年度 社員名簿')!;
    expect(title.ha).toBe(7);
    // numFmt set before the value; dynamic formula spills; legacy formula keeps implicit intersection
    const { M, N } = built; // final positions (buildR2Log follows them through every struct op, incl. the last insert)
    expect(cellAt(d, MAIN, 3, M)).toMatchObject({ v: 46000, nf: 'yyyy/m/d', hf: false });
    expect(cellAt(d, MAIN, 4, M)).toMatchObject({ v: 0.25, nf: '0%' });
    expect(cellAt(d, MAIN, 5, M)).toMatchObject({ v: 1234 });
    expect(cellAt(d, MAIN, 5, M)?.nf).toMatch(/¥|\\/);
    expect(cellAt(d, MAIN, 6, M)).toMatchObject({ f2: '=SEQUENCE(3)', spill: `${numToCol(M)}6:${numToCol(M)}8` });
    expect(cellAt(d, MAIN, 10, M)?.f2).toMatch(/^=@/);
    expect(cellAt(d, MAIN, 11, M)).toMatchObject({ v: '001', nf: '@' });
    expect(cellAt(d, MAIN, 5, N)).toMatchObject({ v: true, hf: false });
    expect(cellAt(d, MAIN, 6, N)).toMatchObject({ v: 'ERR:#N/A', hf: false });
    expect(cellAt(d, MAIN, 7, N)).toMatchObject({ v: '=text', hf: false });
    expect(cellAt(d, MAIN, 4, N)?.f).toMatch(/^=XLOOKUP\(/);
    // a formula over date cells keeps the cell's own format: the log carries no numFmt (Excel alone would copy the
    // referenced cell's date format into the General cell)
    expect(cellAt(d, MAIN, 3, N)).toMatchObject({ hf: true, v: 46000 + 0.25 + 1234 });
    expect(cellAt(d, MAIN, 3, N)?.nf).toMatch(/^(G\/標準|General)$/);
    // number formats widen a standard-width column like typing does: no "########" in the output
    // (format applied AFTER the value; set before it Excel would not widen - tests/golden/RESULTS.md)
    const { P } = built;
    for (const [row, col, text] of [[3, M, /^\d{4}\/\d{1,2}\/\d{1,2}$/], [5, M, /^[¥\\]1,234$/], [3, P, /^\d{4}\/\d{1,2}\/\d{1,2}$/], [6, P, /^\d{4}\/\d{1,2}\/\d{1,2}$/], [7, P, /^123,456,789$/]] as const) {
      expect(cellAt(d, MAIN, row, col)?.t, `${numToCol(col)}${row}`).toMatch(text);
    }
    expect(sheetOf(d, MAIN).widths[M - 1]).toBeGreaterThan(9);
    expect(sheetOf(d, MAIN).widths[P - 1]).toBeGreaterThan(10.5); // 123,456,789 is wider than a date
    // toValues: the whole CSE array and the whole spill became values (same numbers), 係数x10 too
    const hdr = byHeader(d, MAIN, 2);
    for (const [h, origCol] of [['倍', 10], ['順位', 11], ['係数x10', 12]] as const) {
      const c = hdr.get(h)!;
      expect(c, h).toBe(col(origCol));
      for (let r2 = 3; r2 <= 12; r2++) {
        const cell = cellAt(d, MAIN, r2, c);
        expect(cell?.hf, `${h} row ${r2}`).toBe(false);
        expect(cell?.v, `${h} row ${r2}`).toBe(cellAt(before, MAIN, r2, origCol)?.v);
        expect(cell?.arr).toBeUndefined();
      }
    }
    expect(hdr.has('係数')).toBe(false);
    expect(d.names.find((n) => n.name === '係数列')?.refersTo).toMatch(/#REF!/);
    // table: header order, structured references unchanged, no blank column
    const t = sheetOf(d, TSHEET).tables[0];
    expect(t).toMatchObject({ name: 'Tbl', range: 'A1:E6', columns: ['コード', '品名', '数量', '単価', '金額'] });
    expect(cellAt(d, TSHEET, 2, 7)?.f).toBe('=SUM(Tbl[金額])');
    // a column span keeps covering the same two columns (Excel may write it in the new order)
    expect(['=SUM(Tbl[[単価]:[数量]])', '=SUM(Tbl[[数量]:[単価]])']).toContain(cellAt(d, TSHEET, 3, 7)?.f);
    expect(cellAt(d, TSHEET, 2, 5)?.f).toBe('=[@単価]*[@数量]');
    expect(cellAt(d, SUM, 7, 2)?.f).toBe('=SUM(Tbl[金額])');
    expect(cellAt(d, SUM, 8, 2)?.f).toBe('=INDEX(Tbl[単価],2)');
    const moves = events.filter((e) => e.type === 'progress' && e.phase === 'op' && e.op === 'moveCols');
    expect(moves.map((e) => e.method)).toContain('cutInsert');
    // widths: stored 20 -> ColumnWidth 20 - padding
    const pads = events.filter((e) => e.type === 'progress' && e.op === 'setColumnWidth').map((e) => Number(e.padding));
    expect(pads[0]).toBeGreaterThan(0.3);
    expect(pads[0]).toBeLessThan(1.2);
    // widths[] is 0-based; Excel reports ColumnWidth rounded to whole pixels (19.375 -> 19.38)
    expect(sheetOf(d, MAIN).widths[col(2) - 1]).toBeCloseTo(20 - pads[0], 1);
    expect(sheetOf(d, MAIN).widths[col(5) - 1]).toBeCloseTo(12.625 - pads[0], 1);
    expect(r.warnings?.map((w) => w.code)).toEqual(expect.arrayContaining(['REF_ERRORS']));
  }, 600_000);

  it('bad input does not abort the job: unparsable formula, invalid number format, unmerge of a non-merged area are reported', async () => {
    const ops = [
      { op: 'setCells', sheet: MAIN, at: 'P20', rows: [[{ f: '=SUM(' }], [{ t: 'n', v: 5, numFmt: '[' }], ['ok'], [{ f: '=1+1', numFmt: '0.00' }]] },
      { op: 'unmerge', sheet: MAIN, refs: ['H20:I20', 'Z1'], centerAcross: true },
      { op: 'setColumnWidth', sheet: MAIN, col: 16, width: 256, unit: 'stored' },
    ];
    const out = path.join(dir, 'bad.xlsx');
    const jobFile = path.join(dir, 'bad.job.json');
    fs.writeFileSync(jobFile, JSON.stringify({ version: 1, format: 'xlsx', options: { recalc: true, clipboard: 'restoreEachOp' }, ops, input: fixture, output: out, workDir: path.join(dir, 'w-bad') }));
    const r = await runRunner({ runner: RUNNER, jobFile, env: { XLRUNNER_CACHE: path.join(dir, 'cache') } });
    expect(r.ok, JSON.stringify(r)).toBe(true);
    const codes = (r.warnings ?? []).map((w) => w.code);
    expect(codes).toEqual(expect.arrayContaining(['BAD_FORMULA', 'BAD_NUMFMT']));
    for (const w of r.warnings ?? []) expect(w.message, w.code).toMatch(/[぀-ヿ一-鿿]/); // Japanese text
    const d = await dumpBook(out);
    expect(cellAt(d, MAIN, 20, 16)).toMatchObject({ v: '=SUM(', hf: false }); // kept as text, not lost
    expect(cellAt(d, MAIN, 21, 16)).toMatchObject({ v: 5 });
    expect(cellAt(d, MAIN, 22, 16)).toMatchObject({ v: 'ok' });
    expect(cellAt(d, MAIN, 23, 16)).toMatchObject({ v: 2, nf: '0.00' });
    expect(sheetOf(d, MAIN).widths[15]).toBeGreaterThan(250); // stored 256 -> 255.4: still in range
  }, 300_000);
});
