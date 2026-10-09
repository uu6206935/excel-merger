/**
 * buildPremerged, round 8: imported formulas stay formulas. Hand-built source / base workbooks, a hand-made join plan (the keys module is
 * exercised on the corpus fixtures in formula-import-corpus.test.ts). Source persons in file order (rows 2..6): P3 P1 P5 P2 P4, so
 * the merged row of source row r is T = {2: 4, 3: 2, 4: 6, 5: 3, 6: 5} (base order P1..P5 on rows 2..6).
 * Merged columns: the base has ID 氏名 大学 (A..C), the imports start at D: source C -> D, D -> E, E -> F ...
 */
import { describe, expect, it } from 'vitest';
import { snapshot } from '../golden/snapshot';
import { PEOPLE, merge, type BookSpec, type F, type G, type RowPlan } from './formulaimport-helpers';

const SRC_IDS = ['P3', 'P1', 'P5', 'P2', 'P4'];
const P1: Record<string, number> = { P1: 10, P2: 20, P3: 30, P4: 40, P5: 50 };
const P2: Record<string, number> = { P1: 1, P2: 2, P3: 3, P4: 4, P5: 5 };
const T: Record<number, number> = { 2: 4, 3: 2, 4: 6, 5: 3, 6: 5 };
const PLAN: RowPlan[] = [{ base: 2, src: 3 }, { base: 3, src: 5 }, { base: 4, src: 2 }, { base: 5, src: 6 }, { base: 6, src: 4 }];
const IDENT_IDS = ['P1', 'P2', 'P3', 'P4', 'P5'];
const IDENT_PLAN: RowPlan[] = [2, 3, 4, 5, 6].map((r) => ({ base: r, src: r }));

type Col = { h: string; make: (r: number, id: string) => G };
const base = (): BookSpec => ({ sheet: '名簿', rows: [['ID', '氏名', '大学'], ...PEOPLE.map((p, i) => [p, `名前${i + 1}`, `大学${i + 1}`])] });
function src(cols: Col[], ids = SRC_IDS, extra: Partial<BookSpec> = {}): BookSpec {
  return {
    sheet: 'S',
    rows: [['ID', '氏名', '点1', '点2', ...cols.map((c) => c.h)], ...ids.map((id, i) => [id, `名前${id.slice(1)}`, P1[id], P2[id], ...cols.map((c) => c.make(i + 2, id))])],
    ...extra,
  };
}
const f = (text: string, v?: number | string): F => (v === undefined ? { f: text } : { f: text, v });

describe('same-row formulas (the 1st case that always worked) and the cached result', () => {
  it('kept live with the references moved to the merged columns / rows; the stored result stays', async () => {
    const m = await merge({ base: base(), src: src([{ h: '合計', make: (r, id) => f(`C${r}+D${r}`, P1[id] + P2[id]) }]), imports: [3, 4, 5], rows: PLAN });
    for (const r of [2, 3, 4, 5, 6]) {
      const c = m.cell(T[r], 6)!;
      expect(c.f).toBe(`D${T[r]}+E${T[r]}`);
      expect(c.v).toBe(P1[SRC_IDS[r - 2]] + P2[SRC_IDS[r - 2]]);
    }
    expect(m.report).toMatchObject({ formulasKept: 5, formulasToValues: 0 });
    expect(m.report.formulaStats).toEqual([{ sourceId: 's', label: 'src', kept: 5, toValues: 0, empty: 0, reasons: {} }]);
    expect(m.trace.every((t) => t.outcome === 'kept' && t.cached === true && t.exact === true)).toBe(true);
  });

  it('a source that stored no results: the formulas are kept, without a result (Excel computes them on load), and it is said', async () => {
    const m = await merge({ base: base(), src: src([{ h: '合計', make: (r) => f(`C${r}+D${r}`) }]), imports: [3, 4, 5], rows: PLAN });
    expect(m.cell(4, 6)).toMatchObject({ f: 'D4+E4' });
    expect(m.cell(4, 6)!.v ?? null).toBeNull();
    expect(m.report.formulaStats![0]).toMatchObject({ kept: 5, toValues: 0, empty: 0 });
    expect(m.items.find((i) => i.code === 'formulasRecalc')?.message).toMatch(/Excel で開いたとき/);
    expect(m.items.some((i) => i.code === 'formulasNoCache')).toBe(false);
    expect((await snapshot(m.res.bytes)).fullCalcOnLoad).toBe(true);
  });

  it('text, boolean and error results are kept next to the formula', async () => {
    const m = await merge({
      base: base(),
      src: src([
        { h: '判定', make: (r, id) => f(`IF(C${r}>=30,"合格","不可")`, P1[id] >= 30 ? '合格' : '不可') },
        { h: '真偽', make: (r, id) => ({ f: `C${r}>D${r}`, v: P1[id] > P2[id] }) },
        { h: 'エラー', make: (r) => ({ f: `NA()`, e: '#N/A', ...(r ? {} : {}) }) },
        { h: '空', make: (r, id) => ({ f: `IF(C${r}>1000,1,"")`, v: P1[id] > 1000 ? 1 : '' }) },
      ]),
      imports: [3, 4, 5, 6, 7, 8],
      rows: PLAN,
    });
    expect(m.cell(4, 6)).toMatchObject({ f: 'IF(D4>=30,"合格","不可")', v: '合格' });
    expect(m.cell(2, 6)).toMatchObject({ v: '不可' });
    expect(m.cell(4, 7)).toMatchObject({ f: 'D4>E4', v: true });
    expect(m.cell(4, 8)).toMatchObject({ f: 'NA()', t: 'e', v: '#N/A' });
    expect(m.cell(4, 9)).toMatchObject({ f: 'IF(D4>1000,1,"")', v: '' });
  });

  it('rows of the base that matched nothing hold no formula (and keep the column\'s look)', async () => {
    const plan: RowPlan[] = [{ base: 2, src: 3 }, { base: 3, src: null }, { base: 4, src: 2 }, { base: 5, src: null }, { base: 6, src: 4 }];
    const m = await merge({ base: base(), src: src([{ h: '合計', make: (r, id) => f(`C${r}+D${r}`, P1[id] + P2[id]) }]), imports: [3, 4, 5], rows: plan });
    expect(m.cell(3, 6)?.f).toBeUndefined();
    expect(m.cell(5, 6)?.f).toBeUndefined();
    expect(m.report.formulasKept).toBe(3);
  });
});

describe('references to OTHER rows go through the join plan (rows matched by key land on other rows)', () => {
  it('previous-row differences and next-row references follow the people, not the row numbers', async () => {
    const m = await merge({
      base: base(),
      src: src([
        { h: '前差', make: (r, id) => (r === 2 ? f(`C${r}`, P1[id]) : f(`C${r}-C${r - 1}`, P1[id] - P1[SRC_IDS[r - 3]])) },
        { h: '次', make: (r, id) => f(`C${r}+1`, P1[id] + 1) },
      ]),
      imports: [3, 4, 5, 6],
      rows: PLAN,
    });
    // source row 3 (P1) minus source row 2 (P3): merged row 2 minus merged row 4
    expect(m.cell(2, 6)).toMatchObject({ f: 'D2-D4', v: 10 - 30 });
    // source row 6 (P4) minus source row 5 (P2): merged row 5 minus merged row 3
    expect(m.cell(5, 6)).toMatchObject({ f: 'D5-D3', v: 40 - 20 });
    expect(m.report.formulasToValues).toBe(0);
  });

  it('a reference to a source row that was not placed cannot be remapped (rowRelation) and becomes the value', async () => {
    const plan: RowPlan[] = [{ base: 2, src: 3 }, { base: 3, src: 5 }, { base: 4, src: 2 }, { base: 5, src: null }, { base: 6, src: 4 }];
    // each row reads the NEXT source row: source row 5 (P2) reads row 6 (P4), which matched nothing
    const m = await merge({ base: base(), src: src([{ h: '次との差', make: (r, id) => f(`C${r}-C${r + 1}`, P1[id] - (SRC_IDS[r - 1] ? P1[SRC_IDS[r - 1]] : 0)) }]), imports: [3, 4, 5], rows: plan });
    expect(m.report.formulaStats![0]).toMatchObject({ kept: 3, toValues: 1, empty: 0, reasons: { rowRelation: 1 } });
    expect(m.cell(4, 6)).toMatchObject({ f: 'D4-D2' }); // P3 (row 2) minus the next source row 3 = P1 (merged row 2)
    expect(m.cell(3, 6)?.f).toBeUndefined(); // P2 (row 5)
    expect(m.cell(3, 6)?.v).toBe(20 - 40);
  });

  it('running totals: kept as a range of consecutive rows when the order is the base order, a value when it is not', async () => {
    const run = (r: number, ids: string[]) => ids.slice(0, r - 1).reduce((a, id) => a + P1[id], 0);
    const mk = (ids: string[]) => src([{ h: '累計', make: (r) => f(`SUM(C$2:C${r})`, run(r, ids)) }], ids);
    const same = await merge({ base: base(), src: mk(IDENT_IDS), imports: [3, 4, 5], rows: IDENT_PLAN });
    expect(same.cell(4, 6)).toMatchObject({ f: 'SUM(D$2:D4)', v: 10 + 20 + 30 });
    expect(same.report.formulasToValues).toBe(0);
    const shuf = await merge({ base: base(), src: mk(SRC_IDS), imports: [3, 4, 5], rows: PLAN });
    expect(shuf.cell(4, 6)).toMatchObject({ v: 30 }); // source row 2 (P3): SUM(C$2:C2) = its own value, a single cell -> live
    expect(shuf.cell(3, 6)?.f).toBeUndefined(); // P2 = source row 5: SUM(C$2:C5) over rows that are not consecutive in the merged sheet
    expect(shuf.cell(3, 6)?.v).toBe(30 + 10 + 50 + 20);
    expect(shuf.report.formulaStats![0].reasons).toEqual({ rowRelation: 3 });
    expect(shuf.items.find((i) => i.code === 'formulasToValues')?.message).toMatch(/前の行との差・途中までの合計/);
  });

  it('absolute references to one data row and to the header row', async () => {
    const m = await merge({
      base: base(),
      src: src([
        { h: '先頭比', make: (r, id) => f(`C${r}-C$2`, P1[id] - 30) },
        { h: '見出し', make: (r, id) => f(`C$1&":"&C${r}`, `点1:${P1[id]}`) },
      ]),
      imports: [3, 4, 5, 6],
      rows: PLAN,
    });
    expect(m.cell(2, 6)).toMatchObject({ f: 'D2-D$4', v: 10 - 30 });
    expect(m.cell(2, 7)).toMatchObject({ f: 'D$1&":"&D2', v: '点1:10' });
  });

  it('cells outside the table (a rate in a side cell, a footer) are externalCell and become values', async () => {
    const s = src([{ h: '税込', make: (r, id) => f(`C${r}*$J$1`, P1[id] * 1.1) }, { h: '構成比', make: (r, id) => f(`C${r}/$C$9`, P1[id] / 150) }]);
    s.rows[0][9] = 'rate';
    const m = await merge({ base: base(), src: s, imports: [3, 4, 5, 6], rows: PLAN, srcRegion: { lastCol: 6 } });
    expect(m.report.formulaStats![0]).toMatchObject({ kept: 0, toValues: 10, empty: 0, reasons: { externalCell: 10 } });
    expect(m.cell(4, 6)).toMatchObject({ v: 30 * 1.1 });
    expect(m.cell(4, 6)?.f).toBeUndefined();
    expect(m.items.find((i) => i.code === 'formulasToValues')?.message).toMatch(/表の外のセル（例: 税率）を参照している数式 10 件は値にしました/);
  });
});

describe('ranges over the whole source table become the whole merged table (rule b)', () => {
  const sum = (ids: string[]) => ids.reduce((a, id) => a + P1[id], 0);
  const cols = (ids: string[]): Col[] => [
    { h: '合計', make: () => f('SUM(C$2:C$6)', sum(ids)) },
    { h: '平均', make: () => f('AVERAGE(C2:C6)', sum(ids) / 5) },
    { h: '順位', make: (r, id) => f(`RANK(C${r},$C$2:$C$6)`, 6 - Math.round(P1[id] / 10)) },
    { h: '列全体', make: (r, id) => f(`C${r}/SUM(C:C)`, P1[id] / sum(ids)) },
    { h: '余裕', make: () => f('MAX(C2:C1000)', 50) },
  ];
  it('every source row placed: the same set of rows, so the stored results stay; no "widened" notice', async () => {
    const m = await merge({ base: base(), src: src(cols(SRC_IDS)), imports: [3, 4, 5, 6, 7, 8, 9], rows: PLAN });
    expect(m.cell(4, 6)).toMatchObject({ f: 'SUM(D$2:D$6)', v: 150 });
    expect(m.cell(4, 7)).toMatchObject({ f: 'AVERAGE(D2:D6)', v: 30 });
    expect(m.cell(4, 8)).toMatchObject({ f: 'RANK(D4,$D$2:$D$6)', v: 3 });
    expect(m.cell(4, 9)).toMatchObject({ f: 'D4/SUM(D:D)', v: 0.2 });
    expect(m.cell(4, 10)).toMatchObject({ f: 'MAX(D2:D6)', v: 50 });
    expect(m.items.some((i) => i.code === 'formulaRangesWidened')).toBe(false);
    expect(m.trace.every((t) => t.exact && !t.widened)).toBe(true);
  });

  it('source rows that matched nothing are not in the merged table: the range is widened, the old result dropped, and it is reported', async () => {
    const plan: RowPlan[] = [{ base: 2, src: 3 }, { base: 3, src: 5 }, { base: 4, src: 2 }, { base: 5, src: null }, { base: 6, src: null }];
    const m = await merge({ base: base(), src: src(cols(SRC_IDS)), imports: [3, 4, 5, 6, 7, 8, 9], rows: plan });
    expect(m.cell(4, 6)).toMatchObject({ f: 'SUM(D$2:D$6)' });
    expect(m.cell(4, 6)!.v ?? null).toBeNull(); // never a stale value next to a changed formula
    expect(m.trace.filter((t) => t.outcome === 'kept').every((t) => t.widened && !t.cached)).toBe(true);
    const w = m.items.find((i) => i.code === 'formulaRangesWidened')!;
    expect(w.severity).toBe('info');
    expect(w.message).toMatch(/表全体を参照している数式 15 件/);
    expect(m.items.some((i) => i.code === 'formulasRecalc')).toBe(true);
  });

  it('join "full": the appended rows are part of the merged table (the range grows) and, being source rows, keep the result valid', async () => {
    const ids = [...SRC_IDS, 'P6', 'P7'];
    P1.P6 = 60; P2.P6 = 6; P1.P7 = 70; P2.P7 = 7;
    const s = src([{ h: '合計', make: () => f('SUM(C$2:C$8)', sum(ids)) }, { h: '前行', make: (r) => (r === 2 ? f('C2', 30) : f(`C${r - 1}`, P1[ids[r - 3]])) }], ids);
    const plan: RowPlan[] = [...PLAN, { base: null, src: 7 }, { base: null, src: 8 }];
    const m = await merge({ base: base(), src: s, imports: [3, 4, 5, 6], rows: plan, keyFill: [{ baseCol: 1, sourceCol: 1 }, { baseCol: 2, sourceCol: 2 }] });
    expect(m.res.layout.appended.map((a) => a.row)).toEqual([7, 8]);
    expect(m.cell(4, 6)).toMatchObject({ f: 'SUM(D$2:D$8)', v: 280 });
    // appended row 7 = source row 7 (P6): its previous-row formula reads source row 6 (P4) -> merged row 5
    expect(m.cell(7, 7)).toMatchObject({ f: 'D5' });
    expect(m.cell(7, 4)).toMatchObject({ v: 60 });
    expect(m.cell(1 + 6, 1)).toMatchObject({ v: 'P6' });
  });

  it('a header-including range ends at the base header row; a range beyond the data is the data', async () => {
    const m = await merge({ base: base(), src: src([{ h: 'a', make: () => f('COUNTA(C1:C6)', 6) }, { h: 'b', make: () => f('SUM(C2:C100)', 150) }]), imports: [3, 4, 5, 6], rows: PLAN });
    expect(m.cell(2, 6)).toMatchObject({ f: 'COUNTA(D1:D6)', v: 6 });
    expect(m.cell(2, 7)).toMatchObject({ f: 'SUM(D2:D6)', v: 150 });
  });

  it('a footer below the table inside the range keeps it from being a "whole table" range (externalCell)', async () => {
    const s = src([{ h: 'a', make: () => f('SUM(C2:C20)', 150) }]);
    s.rows[7] = [null, '合計', 150];
    const m = await merge({ base: base(), src: s, imports: [3, 4, 5], rows: PLAN, srcRegion: { lastDataRow: 6 } });
    expect(m.report.formulaStats![0].reasons).toEqual({ externalCell: 5 });
  });
});

describe('key columns that are not imported (=A2&"様", COUNTIF over the ID column) read the base partner column (rule c)', () => {
  const cols: Col[] = [
    { h: '様', make: (r, id) => f(`A${r}&"様"`, `${id}様`) },
    { h: '重複', make: (r) => f(`IF(COUNTIF($A$2:$A$6,A${r})>1,"重複","")`, '') },
    { h: '氏名参照', make: (r, id) => f(`B${r}&"/"&C${r}`, `名前${id.slice(1)}/${P1[id]}`) },
    { h: '引当', make: (r, id) => f(`INDEX($D$2:$D$6,MATCH(A${r},$A$2:$A$6,0))`, P2[id]) },
  ];
  const pairs = [{ baseCol: 1, sourceCol: 1 }, { baseCol: 2, sourceCol: 2 }];

  it('with keyPairs and equal texts on every matched row: live, reading the base ID / 氏名 columns', async () => {
    const m = await merge({ base: base(), src: src(cols), imports: [3, 4, 5, 6, 7, 8], rows: PLAN, keyPairs: pairs });
    expect(m.cell(4, 6)).toMatchObject({ f: 'A4&"様"', v: 'P3様' });
    expect(m.cell(4, 7)).toMatchObject({ f: 'IF(COUNTIF($A$2:$A$6,A4)>1,"重複","")' });
    expect(m.cell(4, 8)).toMatchObject({ f: 'B4&"/"&D4', v: '名前3/30' });
    expect(m.cell(4, 9)).toMatchObject({ f: 'INDEX($E$2:$E$6,MATCH(A4,$A$2:$A$6,0))', v: 3 });
    expect(m.trace.filter((t) => t.substituted).length).toBe(20);
    expect(m.report.formulasToValues).toBe(0);
  });

  it('without keyPairs nothing is substituted: notImported, values', async () => {
    const m = await merge({ base: base(), src: src(cols), imports: [3, 4, 5, 6, 7, 8], rows: PLAN });
    expect(m.report.formulaStats![0].reasons).toEqual({ notImported: 20 });
    expect(m.cell(4, 6)).toMatchObject({ v: 'P3様' });
    expect(m.cell(4, 6)?.f).toBeUndefined();
    expect(m.items.find((i) => i.code === 'formulasToValues')?.message).toMatch(/取り込まない列を参照している数式/);
  });

  it('one matched row whose texts differ keeps the whole column from being substituted (a column-level decision)', async () => {
    const b = base();
    b.rows[3][0] = 'ｐ3'; // base row 4 holds a full-width spelling of P3
    const m = await merge({ base: b, src: src(cols), imports: [3, 4, 5, 6, 7, 8], rows: PLAN, keyPairs: pairs });
    // ID (A) differs -> not substituted; 氏名 (B) agrees -> still substituted
    expect(m.report.formulaStats![0].reasons).toEqual({ notImported: 15 });
    expect(m.cell(4, 8)).toMatchObject({ f: 'B4&"/"&D4' });
  });

  it('a different type (number vs text) also keeps the column out', async () => {
    const b = base();
    b.rows[1][0] = 1;
    const s = src([{ h: '様', make: (r, id) => f(`A${r}&"様"`, `${id}様`) }]);
    s.rows[2][0] = '1';
    const m = await merge({ base: b, src: s, imports: [3, 4, 5], rows: PLAN, keyPairs: pairs });
    expect(m.report.formulaStats![0].reasons).toEqual({ notImported: 5 });
  });

  it('appended rows: the base key cells come from keyFill; substitution needs them to be equal too', async () => {
    const ids = [...SRC_IDS, 'P6'];
    P1.P6 = 60; P2.P6 = 6;
    const s = src([{ h: '様', make: (r, id) => f(`A${r}&"様"`, `${id}様`) }], ids);
    const plan: RowPlan[] = [...PLAN, { base: null, src: 7 }];
    const ok = await merge({ base: base(), src: s, imports: [3, 4, 5], rows: plan, keyPairs: pairs, keyFill: [{ baseCol: 1, sourceCol: 1 }, { baseCol: 2, sourceCol: 2 }] });
    expect(ok.report.formulaStats![0].kept).toBe(6);
    expect(ok.cell(7, 6)).toMatchObject({ f: 'A7&"様"' });
    const noFill = await merge({ base: base(), src: s, imports: [3, 4, 5], rows: plan, keyPairs: pairs, keyFill: [] });
    expect(noFill.report.formulaStats![0].reasons).toEqual({ notImported: 6 });
  });

  it('imported key columns are read as imported columns (all-columns import)', async () => {
    const m = await merge({ base: base(), src: src(cols), imports: [1, 2, 3, 4, 5, 6, 7, 8], rows: PLAN, keyPairs: pairs });
    expect(m.cell(4, 8)).toMatchObject({ f: 'D4&"様"' });
    expect(m.trace.some((t) => t.substituted)).toBe(false);
    expect(m.report.formulasToValues).toBe(0);
  });
});
