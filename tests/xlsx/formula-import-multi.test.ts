/**
 * buildPremerged with SEVERAL source files (the real use: a base + the 2nd, 3rd … file): every source has its own row mapping, its own
 * columns, its own report; rows appended from one source are matched by the next.
 */
import { describe, expect, it } from 'vitest';
import type { PremergeInput } from '../../src/core/contracts';
import type { MergedRowPlan } from '../../src/core/types';
import { buildPremerged, type FormulaTrace } from '../../src/core/xlsx/premerge';
import { readBook } from '../../src/core/xlsx/read';
import { PEOPLE, buildBook, keySource, type BookSpec, type G } from './formulaimport-helpers';

const base: BookSpec = { sheet: '名簿', rows: [['ID', '氏名', '大学'], ...PEOPLE.map((p, i) => [p, `名前${i + 1}`, `大学${i + 1}`])] };
const A_IDS = ['P3', 'P1', 'P5', 'P2', 'P4'];
const P1: Record<string, number> = { P1: 10, P2: 20, P3: 30, P4: 40, P5: 50, P6: 60 };
const srcA: BookSpec = {
  sheet: 'A',
  rows: [['ID', '氏名', '点1', '点2', '合計'], ...A_IDS.map((id, i): G[] => [id, id, P1[id], Number(id.slice(1)), { f: `C${i + 2}+D${i + 2}`, v: P1[id] + Number(id.slice(1)) }])],
};
/** B: another layout (ID, 備考, 評価 …), in the base order, plus a stranger P6 that only B has */
const B_IDS = ['P1', 'P2', 'P3', 'P4', 'P5', 'P6'];
const srcB: BookSpec = {
  sheet: 'B',
  rows: [['ID', '備考', '評価', '二倍', '累計', '順位'], ...B_IDS.map((id, i): G[] => [id, `備${id}`, P1[id], { f: `C${i + 2}*2`, v: P1[id] * 2 }, { f: `SUM(C$2:C${i + 2})`, v: B_IDS.slice(0, i + 1).reduce((a, x) => a + P1[x], 0) }, { f: `RANK(C${i + 2},$C$2:$C$7)`, v: 6 - i }])],
};

async function run(join: 'left' | 'full') {
  const [bb, ba, bbb] = await Promise.all([buildBook(base), buildBook(srcA), buildBook(srcB)]);
  const baseKs = await keySource('b', 'base', bb, { sheet: '名簿', headerRow: 1, firstDataRow: 2, lastDataRow: 6, firstCol: 1, lastCol: 3 });
  const a = await keySource('a', 'fileA', ba, { sheet: 'A', headerRow: 1, firstDataRow: 2, lastDataRow: 6, firstCol: 1, lastCol: 5 });
  const b = await keySource('bb', 'fileB', bbb, { sheet: 'B', headerRow: 1, firstDataRow: 2, lastDataRow: 7, firstCol: 1, lastCol: 6 });
  const aRow = (id: string) => 2 + A_IDS.indexOf(id);
  const bRow = (id: string) => 2 + B_IDS.indexOf(id);
  const rows: MergedRowPlan[] = PEOPLE.map((id, i) => ({
    baseRow: 2 + i,
    matches: { a: { row: aRow(id), via: 'ID', flags: [] }, bb: { row: bRow(id), via: 'ID', flags: [] } },
  }));
  if (join === 'full') rows.push({ baseRow: null, appendedFrom: { sourceId: 'bb', row: bRow('P6') }, matches: { a: null, bb: { row: bRow('P6'), via: 'ID', flags: [] } } });
  const input: PremergeInput = {
    base: baseKs,
    sources: [a, b],
    rows,
    imports: [{ sourceId: 'a', col: 3 }, { sourceId: 'a', col: 4 }, { sourceId: 'a', col: 5 }, { sourceId: 'bb', col: 3 }, { sourceId: 'bb', col: 4 }, { sourceId: 'bb', col: 5 }, { sourceId: 'bb', col: 6 }],
    keyFill: { bb: [{ baseCol: 1, sourceCol: 1 }] },
    options: { formulaPolicy: 'keep-if-possible', extendTable: true, headerConflict: 'suffix' },
  };
  const trace: FormulaTrace[] = [];
  const res = await buildPremerged(input, undefined, (t) => trace.push(t));
  const book = await readBook(res.bytes, 'm.xlsx', { id: 'm' });
  const sheet = book.sheets.find((s) => s.name === '名簿')!;
  return { res, trace, cell: (r: number, c: number) => sheet.rows[r - 1]?.[c - 1] };
}

describe('two sources', () => {
  it('each source remaps its own formulas with its own row mapping and columns (A: shuffled, D/E/F; B: base order, G/H/I/J)', async () => {
    const m = await run('left');
    // A 合計 (source col 5 -> F): P3 = A row 2 -> merged row 4: C2+D2 -> D4+E4
    expect(m.cell(4, 6)).toMatchObject({ f: 'D4+E4', v: 33 });
    expect(m.cell(2, 6)).toMatchObject({ f: 'D2+E2', v: 11 });
    // B 二倍 (source col 4 -> H), 累計 (I), 順位 (J): same row, running total over consecutive rows, rank over the whole merged table
    expect(m.cell(4, 8)).toMatchObject({ f: 'G4*2', v: 60 });
    expect(m.cell(4, 9)).toMatchObject({ f: 'SUM(G$2:G4)', v: 60 });
    expect(m.cell(4, 10)).toMatchObject({ f: 'RANK(G4,$G$2:$G$6)' });
    expect(m.res.report.formulaStats).toEqual([
      { sourceId: 'a', label: 'fileA', kept: 5, toValues: 0, empty: 0, reasons: {} },
      { sourceId: 'bb', label: 'fileB', kept: 15, toValues: 0, empty: 0, reasons: {} },
    ]);
    // B's stranger P6 matched no base row: the RANK range is widened (the merged table has no P6), so its stored result is dropped
    expect(m.trace.filter((t) => t.sourceId === 'bb' && t.srcCol === 6).every((t) => t.widened && t.cached === false)).toBe(true);
    expect(m.res.report.items.filter((i) => i.code === 'formulasKept').map((i) => i.sourceId)).toEqual(['a', 'bb']);
  });

  it('join "full": B\'s stranger is appended below the table; the ranges of B grow over it and keep their results', async () => {
    const m = await run('full');
    expect(m.res.layout.appended).toEqual([{ row: 7, sourceId: 'bb', sourceRow: 7 }]);
    expect(m.cell(4, 10)).toMatchObject({ f: 'RANK(G4,$G$2:$G$7)', v: 4 });
    expect(m.cell(7, 8)).toMatchObject({ f: 'G7*2', v: 120 });
    expect(m.cell(7, 9)).toMatchObject({ f: 'SUM(G$2:G7)', v: 210 });
    expect(m.cell(7, 1)).toMatchObject({ v: 'P6' }); // key cell from keyFill
    // A has no row for P6: its cells there are empty
    expect(m.cell(7, 6)?.f).toBeUndefined();
    expect(m.trace.filter((t) => t.sourceId === 'bb').every((t) => t.outcome === 'kept' && t.cached)).toBe(true);
  });
});
