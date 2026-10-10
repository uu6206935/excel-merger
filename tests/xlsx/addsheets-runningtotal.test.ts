/**
 * The range-extension rule of join 'full' (patch/formula.ts extendArea): ranges that end at the old last data row grow over the appended rows so that
 * 合計 / 件数 formulas include the new people — but a formula INSIDE the data rows that merely ends at its own row (a running total) keeps its range
 * (consistency-oracle class append-range-running: =SUM($D$2:D31) in the last base row used to sum the appended rows too).
 * Rule: a formula inside the data rows keeps a RELATIVE range end that is its own row; every other range that ends at the old last data row and
 * starts at / above the first data row grows (total rows below / above the table, whole-column totals — also with an absolute end — in a data row,
 * formulas on other sheets).
 */
import { describe, expect, it } from 'vitest';
import { PEOPLE, merge, type BookSpec } from './formulaimport-helpers';

// base: rows 2..6 = P1..P5, C = points; D holds the formula under test (row r); E1 / E8 hold totals
const base = (dRow: (i: number, row: number) => string | null, extra: Partial<BookSpec> = {}): BookSpec => ({
  sheet: '名簿',
  rows: [
    ['ID', '氏名', '点', '式', '上'],
    ...PEOPLE.map((p, i) => [p, `名前${i + 1}`, 10 * (i + 1), dRow(i, i + 2) === null ? null : { f: dRow(i, i + 2)! }]),
  ],
  ...extra,
});
const src = (): BookSpec => ({ sheet: '評価', rows: [['ID', '氏名', '評価'], ...[...PEOPLE, 'P9'].map((p, i) => [p, `名前${i + 1}`, i])] });
const plan = () => [...PEOPLE.map((_p, i) => ({ base: i + 2, src: i + 2 })), { base: null, src: 7 }];
const keyFill = [{ baseCol: 1, sourceCol: 1 }, { baseCol: 2, sourceCol: 2 }];
const run = (b: BookSpec) => merge({ base: b, src: src(), imports: [3], rows: plan(), keyFill, baseRegion: { lastDataRow: 6 } });

describe("join 'full': which ranges grow over the appended row", () => {
  it('a running total in the LAST data row keeps its range; in the other data rows too', async () => {
    const m = await run(base((_i, r) => `SUM($C$2:C${r})`));
    for (let r = 2; r <= 6; r++) expect(m.cell(r, 4)!.f, `D${r}`).toBe(`SUM($C$2:C${r})`);
  });

  it('a running total whose start is the header row, or with a relative start, keeps its range', async () => {
    const m = await run(base((i, r) => (i === 4 ? `SUM(C1:C${r})` : `SUM(C$2:C${r})`)));
    expect(m.cell(6, 4)!.f).toBe('SUM(C1:C6)');
    expect(m.cell(3, 4)!.f).toBe('SUM(C$2:C3)');
  });

  it('a total of the WHOLE data column inside a data row (end row absolute, or the end lies below the formula) grows', async () => {
    const abs = await run(base((i) => (i === 4 ? 'SUM($C$2:$C$6)' : null)));
    expect(abs.cell(6, 4)!.f).toBe('SUM($C$2:$C$7)'); // the last data row, absolute end: not a running total
    const mid = await run(base((i) => (i === 2 ? 'SUM(C2:C6)' : null)));
    expect(mid.cell(4, 4)!.f).toBe('SUM(C2:C7)'); // a middle row that totals the column
    const midAbs = await run(base((i) => (i === 1 ? 'C3/SUM($C$2:$C$6)' : null)));
    expect(midAbs.cell(3, 4)!.f).toBe('C3/SUM($C$2:$C$7)');
  });

  it('a running total in a middle row keeps its range (its end is its own row)', async () => {
    const m = await run(base((i) => (i === 2 ? 'SUM($C$2:C4)' : null)));
    expect(m.cell(4, 4)!.f).toBe('SUM($C$2:C4)');
  });

  it('a total row BELOW the table and a total ABOVE it grow; the total row moves down with the appended row', async () => {
    const b = base(() => null);
    b.rows.push([], ['合計', null, { f: 'SUM(C2:C6)' }, { f: 'COUNTA(A2:A6)' }]); // row 8
    b.rows[0][5] = '上計';
    b.rows[1][5] = { f: 'SUM(C2:C6)' }; // F2 sits in the first data row: a whole-column total inside the data rows
    b.rows[0][6] = { f: 'SUM(C2:C6)' }; // G1 sits above the table
    const m = await run(b);
    expect(m.cell(9, 3)!.f).toBe('SUM(C2:C7)');
    expect(m.cell(9, 4)!.f).toBe('COUNTA(A2:A7)');
    expect(m.cell(1, 7)!.f).toBe('SUM(C2:C7)');
    expect(m.cell(2, 6)!.f).toBe('SUM(C2:C7)');
  });

  it("a formula of ANOTHER sheet that reads the whole column grows; so does a name (absolute) that points at it", async () => {
    const b = base(() => null, {
      extra: [{ name: '集計', rows: [[{ f: 'SUM(名簿!C2:C6)' }, { f: 'SUM(名簿!$C$2:C6)' }]] }],
      names: [{ name: 'データ列', refersTo: '名簿!$C$2:$C$6' }],
    });
    const m = await run(b);
    const out = m.book.sheets.find((s) => s.name === '集計')!;
    expect(out.rows[0]![0]!.f).toBe('SUM(名簿!C2:C7)');
    expect(out.rows[0]![1]!.f).toBe('SUM(名簿!$C$2:C7)');
    expect(m.book.names.find((n) => n.name === 'データ列')!.refersTo).toBe('名簿!$C$2:$C$7');
  });
});
