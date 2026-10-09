/**
 * buildPremerged, round 8 — defined names, Excel Tables, shared formulas, array formulas, unsafe formulas, the policy, the reports,
 * single-sheet output, placeholder results and two sources. Same fixtures as formula-import.test.ts (source persons in file order
 * P3 P1 P5 P2 P4; T = {2: 4, 3: 2, 4: 6, 5: 3, 6: 5}; source C -> merged D, D -> E, E -> F ...).
 */
import { describe, expect, it } from 'vitest';
import { part } from '../golden/build';
import { snapshot } from '../golden/snapshot';
import { PEOPLE, merge, type BookSpec, type F, type G, type RowPlan } from './formulaimport-helpers';

const SRC_IDS = ['P3', 'P1', 'P5', 'P2', 'P4'];
const P1: Record<string, number> = { P1: 10, P2: 20, P3: 30, P4: 40, P5: 50 };
const P2: Record<string, number> = { P1: 1, P2: 2, P3: 3, P4: 4, P5: 5 };
const PLAN: RowPlan[] = [{ base: 2, src: 3 }, { base: 3, src: 5 }, { base: 4, src: 2 }, { base: 5, src: 6 }, { base: 6, src: 4 }];

type Col = { h: string; make: (r: number, id: string) => G };
const base = (): BookSpec => ({ sheet: '名簿', rows: [['ID', '氏名', '大学'], ...PEOPLE.map((p, i) => [p, `名前${i + 1}`, `大学${i + 1}`])] });
function src(cols: Col[], extra: Partial<BookSpec> = {}): BookSpec {
  return {
    sheet: 'S',
    rows: [['ID', '氏名', '点1', '点2', ...cols.map((c) => c.h)], ...SRC_IDS.map((id, i) => [id, `名前${id.slice(1)}`, P1[id], P2[id], ...cols.map((c) => c.make(i + 2, id))])],
    ...extra,
  };
}
const f = (text: string, v?: number | string): F => (v === undefined ? { f: text } : { f: text, v });

describe.each([['SourceBook read by the app', false], ['bare SourceBook (names / tables read from the package)', true]])('defined names — %s', (_label, bare) => {
  const names = [
    { name: '税率', refersTo: '0.08' },
    { name: '評価列', refersTo: 'S!$C$2:$C$6' },
    { name: '係数', refersTo: 'S!$N$1' },
    { name: '区分表', refersTo: 'Master!$A$2:$B$9' },
    { name: '動的', refersTo: 'OFFSET(S!$C$1,1,0,5,1)' },
    { name: 'ラベル', refersTo: '"点"' },
    { name: '範囲点', refersTo: 'S!$D$2:$D$6', localSheetId: 0 },
  ];
  const cols: Col[] = [
    { h: '税込', make: (r, id) => f(`C${r}*税率`, P1[id] * 0.08) },
    { h: '合計', make: () => f('SUM(評価列)', 150) },
    { h: '係数', make: (r, id) => f(`C${r}*係数`, P1[id] * 2) },
    { h: '区分', make: (r) => f(`VLOOKUP(C${r},区分表,2,FALSE)`, 'x') },
    { h: '動的', make: () => f('SUM(動的)', 150) },
    { h: '点付', make: (r, id) => f(`C${r}&ラベル`, `${P1[id]}点`) },
    { h: '範囲', make: () => f('SUM(範囲点)', 15) },
  ];

  it('resolved to the references / constants they stand for; the rest becomes values with the reason', async () => {
    const s = src(cols, { names });
    const m = await merge({ base: base(), src: s, imports: [3, 4, 5, 6, 7, 8, 9, 10, 11], rows: PLAN, bare });
    // 税込 (col 5 -> F), 合計 (G), 係数 (H), 区分 (I), 動的 (J), 点付 (K), 範囲 (L)
    expect(m.cell(4, 6)).toMatchObject({ f: 'D4*0.08', v: 30 * 0.08 });
    expect(m.cell(4, 7)).toMatchObject({ f: 'SUM($D$2:$D$6)', v: 150 });
    expect(m.cell(4, 11)).toMatchObject({ f: 'D4&"点"', v: '30点' });
    expect(m.cell(4, 12)).toMatchObject({ f: 'SUM($E$2:$E$6)', v: 15 });
    expect(m.cell(4, 8)?.f).toBeUndefined(); // 係数 -> a cell beside the table
    expect(m.cell(4, 9)?.f).toBeUndefined(); // 区分表 -> another sheet
    expect(m.cell(4, 10)?.f).toBeUndefined(); // 動的 -> OFFSET
  });

  it('the reasons are counted per source', async () => {
    const s = src(cols, { names });
    const m = await merge({ base: base(), src: s, imports: [3, 4, 5, 6, 7, 8, 9, 10, 11], rows: PLAN, bare });
    const r = m.report.formulaStats![0];
    expect(r.kept).toBe(20);
    expect(r.reasons.unresolvedName).toBe(5);
    expect(r.reasons.otherSheet).toBe(5);
    expect(r.reasons.externalCell).toBe(5); // 係数: a cell beside the table
    expect(r.toValues).toBe(15);
  });
});

describe.each([['SourceBook read by the app', false], ['bare SourceBook (tables read from the package)', true]])('Excel Tables — %s', (_label, bare) => {
  const tables = [{ name: 'Table1', ref: 'A1:G6', columns: ['ID', '氏名', '点1', '点2', '合計', '平均', '順位'] }];
  const cols: Col[] = [
    { h: '合計', make: (_r, id) => f('Table1[[#This Row],[点1]]+Table1[[#This Row],[点2]]', P1[id] + P2[id]) },
    { h: '平均', make: () => f('SUM(Table1[点1])/COUNT(Table1[点1])', 30) },
    { h: '順位', make: (_r, id) => f('RANK(Table1[[#This Row],[点1]],Table1[点1])', 6 - P1[id] / 10) },
  ];

  it('this-row and column references are resolved to A1 references and remapped', async () => {
    const m = await merge({ base: base(), src: src(cols, { tables }), imports: [3, 4, 5, 6, 7], rows: PLAN, bare });
    expect(m.cell(4, 6)).toMatchObject({ f: 'D4+E4', v: 33 });
    expect(m.cell(4, 7)).toMatchObject({ f: 'SUM($D$2:$D$6)/COUNT($D$2:$D$6)', v: 30 });
    expect(m.cell(4, 8)).toMatchObject({ f: 'RANK(D4,$D$2:$D$6)', v: 3 });
    expect(m.report.formulasToValues).toBe(0);
  });

  it('a table the formula reads but whose columns are not imported: notImported', async () => {
    const m = await merge({ base: base(), src: src(cols, { tables }), imports: [5, 6, 7], rows: PLAN, bare });
    expect(m.report.formulaStats![0].reasons).toEqual({ notImported: 15 });
  });

  it('a missing table is unresolvedTable', async () => {
    const m = await merge({ base: base(), src: src([{ h: 'x', make: () => f('SUM(Nope[a])', 1) }]), imports: [3, 4, 5], rows: PLAN, bare });
    expect(m.report.formulaStats![0].reasons).toEqual({ unresolvedTable: 5 });
  });
});

describe('shared formulas and arrays', () => {
  it('shared formulas: master and dependents (no text of their own) are expanded and kept', async () => {
    const cols: Col[] = [{
      h: '合計',
      make: (r, id) => (r === 2 ? { f: 'C2+D2', v: P1[id] + P2[id], shared: { si: 0, ref: 'E2:E6' } } : { f: '', v: P1[id] + P2[id], shared: { si: 0 } }),
    }];
    const m = await merge({ base: base(), src: src(cols), imports: [3, 4, 5], rows: PLAN });
    expect(m.report).toMatchObject({ formulasKept: 5, formulasToValues: 0 });
    expect(m.cell(2, 6)).toMatchObject({ f: 'D2+E2', v: 11 }); // P1 = source row 3: C3+D3 -> D2+E2
    expect(m.cell(6, 6)).toMatchObject({ f: 'D6+E6', v: 55 });
  });

  it('a shared dependent whose master is missing is a plain value, not an empty cell and not an array', async () => {
    const cols: Col[] = [{ h: '合計', make: (r, id) => (r === 2 ? { f: 'C2+D2', v: 33, shared: { si: 0, ref: 'E2:E3' } } : r === 3 ? { f: '', v: 11, shared: { si: 0 } } : { f: '', v: P1[id] + P2[id], shared: { si: 9 } }) }];
    const m = await merge({ base: base(), src: src(cols), imports: [3, 4, 5], rows: PLAN });
    expect(m.cell(3, 6)).toMatchObject({ v: 22 }); // P2 = source row 5 (dependent of the missing master si=9)
    expect(m.cell(3, 6)?.f).toBeUndefined();
    expect(m.report.formulaStats![0].empty).toBe(0);
    expect(m.report.formulaStats![0].reasons.array).toBeUndefined();
  });

  it('a single-cell legacy array formula stays an array formula (t="array"); a multi-cell one becomes values', async () => {
    const cols: Col[] = [
      { h: '配列', make: (r, id) => ({ f: 'SUM(IF($C$2:$C$6>C' + r + ',1,0))', v: 5 - P1[id] / 10, array: { ref: `E${r}` } }) },
      { h: '複数', make: (r, id) => (r === 2 ? { f: 'C2:C4*2', v: 60, array: { ref: 'F2:F4' } } : r <= 4 ? P1[id] * 2 : null) },
    ];
    const m = await merge({ base: base(), src: src(cols), imports: [3, 4, 5, 6], rows: PLAN });
    expect(m.cell(4, 6)).toMatchObject({ f: 'SUM(IF($D$2:$D$6>D4,1,0))', v: 2 });
    const xml = (await part(m.res.bytes, 'xl/worksheets/sheet1.xml'))!;
    expect(xml).toMatch(/<c r="F4"[^>]*><f t="array" ref="F4">SUM\(IF\(\$D\$2:\$D\$6&gt;D4,1,0\)\)<\/f>/);
    // the multi-cell array: its master (source row 2 = P3 -> merged row 4) is a value; the other cells were values already
    expect(m.cell(4, 7)).toMatchObject({ v: 60 });
    expect(m.cell(4, 7)?.f).toBeUndefined();
    expect(m.report.formulaStats![0].reasons).toEqual({ array: 1 });
  });

  it('a dynamic-array formula that returns one value stays a dynamic array (cm + metadata)', async () => {
    const cols: Col[] = [{ h: '動的', make: (r) => ({ f: `_xlfn.SUM(_xlfn._xlws.FILTER($D$2:$D$6,$C$2:$C$6>C${r}))`, v: 1, array: { ref: `E${r}`, cm: true } }) }];
    const m = await merge({ base: base(), src: src(cols), imports: [3, 4, 5], rows: PLAN });
    const xml = (await part(m.res.bytes, 'xl/worksheets/sheet1.xml'))!;
    expect(xml).toMatch(/<c r="F4"[^>]* cm="1"[^>]*><f t="array" ref="F4">/);
    expect(await part(m.res.bytes, 'xl/metadata.xml')).toContain('XLDAPR');
    expect(m.report.formulasKept).toBe(5);
  });

  it('a spill reference (F2#) reads a spill master that is imported as a value: array', async () => {
    const cols: Col[] = [{ h: '参照', make: () => f('SUM(_xlfn.ANCHORARRAY(E2))', 1) }];
    const m = await merge({ base: base(), src: src(cols), imports: [3, 4, 5], rows: PLAN });
    expect(m.report.formulaStats![0].reasons).toEqual({ array: 5 });
  });
});

describe('formulas that are never kept live, the policy, and the report', () => {
  it('external-workbook references and call-out functions: the cached value (unsafe)', async () => {
    const cols: Col[] = [
      { h: '外部', make: () => f("'C:\\data\\[o.xlsx]S'!A1", 7) },
      { h: '3D', make: () => f('SUM(S:Sheet2!B3)', 8) },
      { h: '位置', make: (r) => f(`INDIRECT("C"&${r})`, 9) },
      { h: '通常', make: (r) => f(`C${r}+1`, 10) },
    ];
    const m = await merge({ base: base(), src: src(cols), imports: [3, 4, 5, 6, 7, 8], rows: PLAN });
    expect(m.report.formulaStats![0]).toMatchObject({ kept: 5, toValues: 15, empty: 0, reasons: { unsafe: 5, otherSheet: 5, positional: 5 } });
    expect(m.cell(4, 6)).toMatchObject({ v: 7 });
    expect(m.cell(4, 6)?.f).toBeUndefined();
    const msgs = m.items.filter((i) => i.code === 'formulasToValues').map((i) => i.message);
    expect(msgs.some((t) => /外部のブックなどを参照している数式 5 件は値にしました/.test(t))).toBe(true);
    expect(msgs.some((t) => /他のシートを参照している数式 5 件は値にしました/.test(t))).toBe(true);
    expect(msgs.some((t) => /ROW・COLUMN・OFFSET・INDIRECT/.test(t))).toBe(true);
  });

  it('a defect in the remapper (here: a SourceBook whose names are broken) never fails the merge: the formulas become values', async () => {
    const cols: Col[] = [{ h: '合計', make: (r, id) => f(`C${r}+D${r}`, P1[id] + P2[id]) }];
    const m = await merge({ base: base(), src: src(cols), imports: [3, 4, 5], rows: PLAN, tweak: (i) => { (i.sources[0].book as unknown as { names: unknown }).names = null; } });
    expect(m.report.formulaStats![0]).toMatchObject({ kept: 0, toValues: 5, empty: 0, reasons: { unsupported: 5 } });
    expect(m.cell(4, 6)).toMatchObject({ v: 33 });
    expect(m.cell(4, 6)?.f).toBeUndefined();
  });

  it('an empty text result (Excel stores <v/>) is a result, not a missing one: no "no result" warning for a formula that cannot be kept', async () => {
    const cols: Col[] = [{ h: '外', make: () => ({ f: 'Sheet2!A1', v: '' }) }, { h: '内', make: (r) => ({ f: `IF(C${r}>1000,1,"")`, v: '' }) }];
    const m = await merge({ base: base(), src: src(cols), imports: [3, 4, 5, 6], rows: PLAN });
    expect(m.report.formulaStats![0]).toMatchObject({ kept: 5, toValues: 5, empty: 0, reasons: { otherSheet: 5 } });
    expect(m.items.some((i) => i.code === 'formulasNoCache')).toBe(false);
    expect(m.cell(4, 6)).toMatchObject({ v: '' });
    expect(m.cell(4, 7)).toMatchObject({ f: 'IF(D4>1000,1,"")', v: '' });
    expect(m.trace.filter((t) => t.outcome === 'kept').every((t) => t.cached === true)).toBe(true);
  });

  it('formulaPolicy "values": every formula becomes its value (unchanged behaviour)', async () => {
    const m = await merge({ base: base(), src: src([{ h: '合計', make: (r, id) => f(`C${r}+D${r}`, P1[id] + P2[id]) }, { h: '無', make: (r) => f(`C${r}*2`) }]), imports: [3, 4, 5, 6], rows: PLAN, options: { formulaPolicy: 'values' } });
    expect(m.report).toMatchObject({ formulasKept: 0, formulasToValues: 10 });
    expect(m.cell(4, 6)).toMatchObject({ v: 33 });
    expect(m.cell(4, 6)?.f).toBeUndefined();
    expect(m.cell(4, 7)?.v ?? null).toBeNull();
    expect(m.report.formulaStats![0]).toMatchObject({ kept: 0, toValues: 5, empty: 5, reasons: { policy: 10 } });
    expect(m.items.find((i) => i.code === 'formulasToValues')?.message).toMatch(/値に変換して取り込みました/);
  });

  it('a formula that cannot be kept AND has no stored result is the only way to get an empty cell, and it is said', async () => {
    const cols: Col[] = [{ h: '外', make: () => f('Sheet2!A1') }, { h: '内', make: (r) => f(`C${r}+1`) }];
    const m = await merge({ base: base(), src: src(cols), imports: [3, 4, 5, 6], rows: PLAN });
    expect(m.cell(4, 6)?.f).toBeUndefined();
    expect(m.cell(4, 6)?.v ?? null).toBeNull();
    expect(m.report.formulaStats![0]).toMatchObject({ kept: 5, toValues: 0, empty: 5, reasons: { otherSheet: 5 } });
    const w = m.items.find((i) => i.code === 'formulasNoCache')!;
    expect(w.severity).toBe('warning');
    expect(w.message).toContain('元のファイルに計算結果が保存されていない数式です（Excelで開いて保存し直してください）');
    expect(w.rows).toEqual([2, 3, 4, 5, 6]);
    expect(m.items.some((i) => i.code === 'formulasToValues')).toBe(false);
  });

  it('every message is Japanese and names the source; the counters are consistent', async () => {
    const cols: Col[] = [{ h: 'a', make: (r) => f(`C${r}+1`, 1) }, { h: 'b', make: () => f('Sheet2!A1', 2) }, { h: 'c', make: () => f('Sheet2!A1') }, { h: 'd', make: () => f('SUM(C2:C6)', 1) }];
    const m = await merge({ base: base(), src: src(cols), imports: [3, 4, 5, 6, 7, 8], rows: PLAN });
    const st = m.report.formulaStats![0];
    expect(st.kept + st.toValues + st.empty).toBe(20);
    expect(m.report.formulasKept).toBe(st.kept);
    expect(m.report.formulasToValues).toBe(st.toValues + st.empty);
    expect(Object.values(st.reasons).reduce((a, b) => a + b, 0)).toBe(st.toValues + st.empty);
    for (const it of m.items) expect(it.message).toMatch(/[ぁ-んァ-ン一-龥]/);
  });
});

describe('single-sheet output (dropOtherSheets) and the other sheets of the base', () => {
  it('imported formulas are written after the drop step: kept formulas stay, formulas that read a sheet of the source become values', async () => {
    const b = base();
    b.extra = [{ name: 'Summary', rows: [[{ f: 'SUM(名簿!A2:A6)', v: 0 }]] }];
    const s = src([{ h: '合計', make: (r, id) => f(`C${r}+D${r}`, P1[id] + P2[id]) }, { h: '集計', make: (r) => f(`Summary!A1+C${r}`, 1) }]);
    s.extra = [{ name: 'Summary', rows: [[1]] }];
    const m = await merge({ base: b, src: s, imports: [3, 4, 5, 6], rows: PLAN, options: { dropOtherSheets: true } });
    expect(m.report.droppedSheets).toEqual(['Summary']);
    expect(m.book.sheets.map((x) => x.name)).toEqual(['名簿']);
    expect(m.cell(4, 6)).toMatchObject({ f: 'D4+E4', v: 33 });
    expect(m.cell(4, 7)?.f).toBeUndefined();
    expect(m.cell(4, 7)?.v).toBe(1);
    expect(m.report.formulaStats![0].reasons).toEqual({ otherSheet: 5 });
    // the same input without dropping keeps the base's Summary and gives the same imported formulas
    const keep = await merge({ base: b, src: s, imports: [3, 4, 5, 6], rows: PLAN, options: { dropOtherSheets: false } });
    expect(keep.book.sheets.map((x) => x.name)).toEqual(['名簿', 'Summary']);
    expect(keep.cell(4, 6)).toMatchObject({ f: 'D4+E4', v: 33 });
    expect(keep.report.formulaStats).toEqual(m.report.formulaStats);
  });
});

describe('results that are only placeholders (XlsxWriter writes 0 and asks Excel to recalculate)', () => {
  const cols: Col[] = [{ h: '合計', make: (r) => ({ f: `C${r}+D${r}`, v: 0 }) }, { h: '外', make: () => ({ f: 'Sheet2!A1', v: 0 }) }];

  it('all-zero results in a fullCalcOnLoad book are "no result stored": kept formulas get none, the others stay empty (not 0)', async () => {
    const m = await merge({ base: base(), src: src(cols, { calcPr: '<calcPr calcId="124519" fullCalcOnLoad="1"/>' }), imports: [3, 4, 5, 6], rows: PLAN });
    expect(m.cell(4, 6)).toMatchObject({ f: 'D4+E4' });
    expect(m.cell(4, 6)?.v ?? null).toBeNull();
    expect(m.cell(4, 7)?.v ?? null).toBeNull();
    expect(m.report.formulaStats![0]).toMatchObject({ kept: 5, toValues: 0, empty: 5 });
    expect(m.items.find((i) => i.code === 'formulasPlaceholderCache')?.message).toMatch(/結果が 0 のまま保存されている/);
  });

  it('a real zero among other results, or a book without fullCalcOnLoad, is trusted', async () => {
    const real = await merge({ base: base(), src: src([{ h: '合計', make: (r, id) => ({ f: `C${r}+D${r}`, v: r === 2 ? 0 : P1[id] }) }], { calcPr: '<calcPr calcId="124519" fullCalcOnLoad="1"/>' }), imports: [3, 4, 5], rows: PLAN });
    expect(real.items.some((i) => i.code === 'formulasPlaceholderCache')).toBe(false);
    expect(real.cell(2, 6)?.v).toBe(10);
    const plain = await merge({ base: base(), src: src(cols), imports: [3, 4, 5, 6], rows: PLAN });
    expect(plain.cell(4, 6)).toMatchObject({ f: 'D4+E4', v: 0 });
    expect(plain.cell(4, 7)).toMatchObject({ v: 0 });
  });
});

describe('two sources and base columns beside the table', () => {
  it('each source has its own row mapping and columns; base columns right of the table move out of the way', async () => {
    const b = base();
    b.rows[0].push('メモ');
    b.rows.slice(1).forEach((r, i) => r.push(`m${i + 1}`));
    const sA = src([{ h: '合計', make: (r, id) => f(`C${r}+D${r}`, P1[id] + P2[id]) }]);
    const bytesS = sA;
    const m = await merge({ base: b, src: bytesS, imports: [3, 4, 5], rows: PLAN, baseRegion: { lastCol: 3 } });
    expect(m.res.layout.insertAt).toBe(4);
    expect(m.cell(4, 6)).toMatchObject({ f: 'D4+E4', v: 33 });
    expect(m.cell(4, 7)).toMatchObject({ v: 'm3' }); // the base's メモ column moved right by the 3 imported columns
    expect(m.cell(1, 7)).toMatchObject({ v: 'メモ' });
  });
});

describe('the premerged package', () => {
  it('asks Excel to recalculate on load and has no calcChain', async () => {
    const m = await merge({ base: base(), src: src([{ h: '合計', make: (r, id) => f(`C${r}+D${r}`, P1[id] + P2[id]) }]), imports: [3, 4, 5], rows: PLAN });
    const snap = await snapshot(m.res.bytes);
    expect(snap.fullCalcOnLoad).toBe(true);
    expect(snap.calcChain).toBe(false);
  });
});
