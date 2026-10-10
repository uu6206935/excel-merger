/**
 * buildPremerged with includeSourceSheets (round 9, task F2): the imported formulas that the merged table cannot keep read the COPIES of the
 * source's sheets and so stay formulas. Hand-built source / base workbooks and a hand-made join plan (helpers: formulaimport-helpers.ts).
 * Source persons in file order (rows 2..6): P3 P1 P5 P2 P4 -> merged row T = {2: 4, 3: 2, 4: 6, 5: 3, 6: 5} (base order P1..P5 on rows 2..6).
 * Source columns: A ID, B 氏名, C 点1, D 点2, E.. formulas; side cell J1, footer total in row 9; other sheets レート and マスタ.
 * File name "src.xlsx" -> the copies are src_S, src_レート, src_マスタ.
 */
import { describe, expect, it } from 'vitest';
import { tokenize } from '../../src/core/formula/excelRef';
import type { PremergeInput } from '../../src/core/contracts';
import type { MergedRowPlan } from '../../src/core/types';
import { buildPremerged, type FormulaTrace } from '../../src/core/xlsx/premerge';
import { readBook } from '../../src/core/xlsx/read';
import { PEOPLE, buildBook, keySource, merge, type BookSpec, type F, type G, type RowPlan } from './formulaimport-helpers';

const SRC_IDS = ['P3', 'P1', 'P5', 'P2', 'P4'];
const P1: Record<string, number> = { P1: 10, P2: 20, P3: 30, P4: 40, P5: 50, P6: 60, P7: 70 };
const P2: Record<string, number> = { P1: 1, P2: 2, P3: 3, P4: 4, P5: 5, P6: 6, P7: 7 };
const T: Record<number, number> = { 2: 4, 3: 2, 4: 6, 5: 3, 6: 5 };
const PLAN: RowPlan[] = [{ base: 2, src: 3 }, { base: 3, src: 5 }, { base: 4, src: 2 }, { base: 5, src: 6 }, { base: 6, src: 4 }];
const OPT = { includeSourceSheets: true };

const base = (): BookSpec => ({ sheet: '名簿', rows: [['ID', '氏名', '大学'], ...PEOPLE.map((p, i) => [p, `名前${i + 1}`, `大学${i + 1}`])] });
const RATES: G[][] = [['項目', '値'], ['税率', 1.1], ['割引', 0.9]];
const MASTER: G[][] = [['ID', '区分'], ...PEOPLE.map((p, i): G[] => [p, i % 2 ? 'B' : 'A'])];
const GRADE = (id: string) => (PEOPLE.indexOf(id) % 2 ? 'B' : 'A');

type Col = { h: string; make: (r: number, id: string) => G };
function src(cols: Col[], ids = SRC_IDS, extra: Partial<BookSpec> = {}): BookSpec {
  const rows: G[][] = [['ID', '氏名', '点1', '点2', ...cols.map((c) => c.h)], ...ids.map((id, i): G[] => [id, `名前${id.slice(1)}`, P1[id], P2[id], ...cols.map((c) => c.make(i + 2, id))])];
  return { sheet: 'S', rows, extra: [{ name: 'レート', rows: RATES }, { name: 'マスタ', rows: MASTER }], ...extra };
}
const f = (text: string, v?: number | string | boolean): F => (v === undefined ? { f: text } : { f: text, v });
const sheetNames = (m: { book: { sheets: Array<{ name: string }> } }) => m.book.sheets.map((s) => s.name);

describe('the formulas that read other sheets of the source (round 8: values) stay formulas, reading the copies', () => {
  const cols: Col[] = [
    { h: '税込', make: (r, id) => f(`C${r}*レート!$B$2`, P1[id] * 1.1) },
    { h: '区分', make: (r, id) => f(`VLOOKUP(A${r},マスタ!$A$2:$B$6,2,FALSE)`, GRADE(id)) },
    { h: '合算', make: (r, id) => f(`C${r}+SUM(レート!$B$2:$B$3)`, P1[id] + 2) },
  ];

  it('without includeSourceSheets nothing changes: values (round 8), reason otherSheet, no sheet is added', async () => {
    const m = await merge({ base: base(), src: src(cols), imports: [3, 4, 5, 6, 7], rows: PLAN });
    expect(m.report.formulaStats![0]).toMatchObject({ kept: 0, toValues: 15, empty: 0, reasons: { otherSheet: 10, notImported: 5 } }); // (the key column A is not imported)
    expect(sheetNames(m)).toEqual(['名簿', 'レート', 'マスタ'].slice(0, 1));
    expect(m.cell(4, 6)?.f).toBeUndefined();
    expect(m.report.addedSheets).toBeUndefined();
  });

  it('with includeSourceSheets: kept as formulas reading the copies, the stored results stay (same cells, same values)', async () => {
    const m = await merge({ base: base(), src: src(cols), imports: [3, 4, 5, 6, 7], rows: PLAN, keyPairs: [{ baseCol: 1, sourceCol: 1 }], options: OPT });
    expect(sheetNames(m)).toEqual(['名簿', 'src_S', 'src_レート', 'src_マスタ']);
    for (const r of [2, 3, 4, 5, 6]) {
      const id = SRC_IDS[r - 2];
      expect(m.cell(T[r], 6)).toMatchObject({ f: `D${T[r]}*src_レート!$B$2`, v: P1[id] * 1.1 });
      expect(m.cell(T[r], 7)).toMatchObject({ f: `VLOOKUP(A${T[r]},src_マスタ!$A$2:$B$6,2,FALSE)`, v: GRADE(id) }); // A = the base's ID column (same text)
      expect(m.cell(T[r], 8)).toMatchObject({ f: `D${T[r]}+SUM(src_レート!$B$2:$B$3)`, v: P1[id] + 2 });
    }
    expect(m.report).toMatchObject({ formulasKept: 15, formulasToValues: 0, addedSheets: ['src_S', 'src_レート', 'src_マスタ'] });
    expect(m.report.formulaStats).toEqual([{ sourceId: 's', label: 'src', kept: 15, toValues: 0, empty: 0, reasons: { viaCopy: 15 } }]);
    expect(m.trace.every((t) => t.outcome === 'kept' && t.viaCopy === true && t.cached === true && t.exact === true && !t.widened)).toBe(true);
  });

  it('the report says it in plain words: an info item with the number, the copies read and example rows', async () => {
    const m = await merge({ base: base(), src: src(cols), imports: [3, 4, 5, 6, 7], rows: PLAN, options: OPT });
    const it = m.items.find((i) => i.code === 'formulasViaCopy')!;
    expect(it).toMatchObject({ severity: 'info', sourceId: 's' });
    expect(it.message).toContain('15 件の数式は元のシート（コピー）を参照して数式のまま残しました');
    for (const n of ['src_S', 'src_レート', 'src_マスタ']) expect(it.message).toContain(`「${n}」`);
    expect(it.rows!.length).toBeGreaterThan(0);
    expect(m.items.some((i) => i.code === 'formulasToValues' || i.code === 'formulasNoCache')).toBe(false);
    expect(m.items.find((i) => i.code === 'formulasKept')).toBeTruthy();
  });

  it('a source without stored results: kept without a result (Excel computes them), nothing becomes an empty cell', async () => {
    const nores: Col[] = [{ h: '税込', make: (r) => f(`C${r}*レート!$B$2`) }, { h: '区分', make: (r) => f(`VLOOKUP(A${r},マスタ!$A$2:$B$6,2,FALSE)`) }];
    const m = await merge({ base: base(), src: src(nores), imports: [3, 4, 5, 6], rows: PLAN, options: OPT });
    expect(m.report.formulaStats![0]).toMatchObject({ kept: 10, toValues: 0, empty: 0, reasons: { viaCopy: 10 } });
    // (the key column A is not imported and there are no keyPairs: it is read from the copy of the data sheet)
    expect(m.cell(4, 6)).toMatchObject({ f: 'D4*src_レート!$B$2' });
    expect(m.cell(4, 7)).toMatchObject({ f: 'VLOOKUP(src_S!A2,src_マスタ!$A$2:$B$6,2,FALSE)' });
    expect(m.items.some((i) => i.code === 'formulasNoCache')).toBe(false);
  });

  it('a hidden sheet of the source is copied and read like any other', async () => {
    const s = src(cols);
    const bytes = await buildBook(s);
    const JSZip = (await import('jszip')).default;
    const z = await JSZip.loadAsync(bytes);
    z.file('xl/workbook.xml', (await z.file('xl/workbook.xml')!.async('string')).replace('<sheet name="レート"', '<sheet state="hidden" name="レート"'));
    const hidden = await z.generateAsync({ type: 'uint8array' });
    const m = await merge({ base: base(), src: s, imports: [3, 4, 5, 6, 7], rows: PLAN, options: OPT, tweak: (input) => { input.sources[0].book.bytes = hidden; } });
    expect(m.cell(4, 6)).toMatchObject({ f: 'D4*src_レート!$B$2' });
    expect(m.report.formulasToValues).toBe(0);
  });
});

describe('cells the merged table cannot represent: side cell, footer, running totals, previous rows, columns that are not imported', () => {
  const sideCell = (s: BookSpec): BookSpec => { s.rows[0][9] = '税率'; s.rows[1][9] = 1.1; s.rows[7] = ['合計', null, 150, 15]; return s; };

  it('a rate in a side cell and a footer total are read from the copy of the data sheet (original coordinates)', async () => {
    const s = sideCell(src([{ h: '税込', make: (r, id) => f(`C${r}*$J$2`, P1[id] * 1.1) }, { h: '構成比', make: (r, id) => f(`C${r}/$C$8`, P1[id] / 150) }]));
    const m = await merge({ base: base(), src: s, imports: [3, 4, 5, 6], rows: PLAN, srcRegion: { lastCol: 6 }, options: OPT });
    expect(m.cell(4, 6)).toMatchObject({ f: 'D4*src_S!$J$2', v: 30 * 1.1 });
    expect(m.cell(4, 7)).toMatchObject({ f: 'D4/src_S!$C$8', v: 0.2 });
    expect(m.report.formulaStats![0]).toMatchObject({ kept: 10, toValues: 0, reasons: { viaCopy: 10 } });
    // the copy holds the source's cells where they were
    const copy = m.book.sheets.find((x) => x.name === 'src_S')!;
    expect(copy.rows[1]?.[9]).toMatchObject({ v: 1.1 });
    expect(copy.rows[7]?.[2]).toMatchObject({ v: 150 });
  });

  it('running totals and previous-row differences over rows the base put in another order: the copy still has the source order', async () => {
    const s = src([
      { h: '累計', make: (r) => f(`SUM(C$2:C${r})`, SRC_IDS.slice(0, r - 1).reduce((a, id) => a + P1[id], 0)) },
      { h: '前差', make: (r, id) => (r === 2 ? f('C2', 30) : f(`C${r}-C${r - 1}`, P1[id] - P1[SRC_IDS[r - 3]])) },
    ]);
    const m = await merge({ base: base(), src: s, imports: [3, 4, 5, 6], rows: PLAN, options: OPT });
    // source row 5 (P2) -> merged row 3: a partial range over source rows that are not consecutive in the merged table reads the copy ...
    expect(m.cell(3, 6)).toMatchObject({ f: 'SUM(src_S!C$2:C5)', v: 30 + 10 + 50 + 20 });
    // ... a single cell of another row follows the person (round 8): source row 4 (P5) is merged row 6
    expect(m.cell(3, 7)).toMatchObject({ f: 'D3-D6', v: 20 - 50 });
    // row 2 of the source is read by its own formula: the own row is always the merged table
    expect(m.cell(4, 7)).toMatchObject({ f: 'D4', v: 30 });
    expect(m.report.formulaStats![0]).toMatchObject({ toValues: 0, empty: 0 });
    expect(m.report.formulaStats![0].reasons.viaCopy).toBe(3); // the running totals of source rows 3..5 (row 2 is the formula's own row, row 6 ends the table: the whole merged table)
  });

  it('a running total over rows that are consecutive in the merged table stays in the merged table (round 8 first)', async () => {
    const ids = ['P1', 'P2', 'P3', 'P4', 'P5'];
    const s = src([{ h: '累計', make: (r) => f(`SUM(C$2:C${r})`, ids.slice(0, r - 1).reduce((a, id) => a + P1[id], 0)) }], ids);
    const m = await merge({ base: base(), src: s, imports: [3, 4, 5], rows: [2, 3, 4, 5, 6].map((r) => ({ base: r, src: r })), options: OPT });
    expect(m.cell(4, 6)).toMatchObject({ f: 'SUM(D$2:D4)', v: 60 });
    expect(m.trace.every((t) => !t.viaCopy)).toBe(true);
    expect(m.report.formulaStats![0].reasons).toEqual({});
  });

  it('columns of the source that are not imported: the formula reads them from the copy, the imported ones from the merged table', async () => {
    const s = src([{ h: '合計', make: (r, id) => f(`C${r}+D${r}`, P1[id] + P2[id]) }, { h: '比', make: (r, id) => f(`D${r}/C${r}`, P2[id] / P1[id]) }]);
    // import only 点1 (C) and the 比 column (F): 点2 (D) and 合計 (E) stay out
    const m = await merge({ base: base(), src: s, imports: [3, 6], rows: PLAN, options: OPT });
    expect(m.cell(4, 5)).toMatchObject({ f: 'src_S!D2/D4', v: 3 / 30 });
    expect(m.report.formulaStats![0]).toMatchObject({ kept: 5, toValues: 0, reasons: { viaCopy: 5 } });
    expect(m.cell(4, 4)).toMatchObject({ v: 30 });
  });

  it('two ranges that are not aligned (one in the merged table, one in a copy) are never mixed: the whole formula reads the copies', async () => {
    const s = src([{ h: '加重', make: () => f('SUMPRODUCT(C$2:C$6,D$2:D$6)', 30 * 3 + 10 * 1 + 50 * 5 + 20 * 2 + 40 * 4) }]);
    const both = await merge({ base: base(), src: s, imports: [3, 4, 5], rows: PLAN, options: OPT });
    expect(both.cell(4, 6)).toMatchObject({ f: 'SUMPRODUCT(D$2:D$6,E$2:E$6)', v: 30 * 3 + 10 * 1 + 50 * 5 + 20 * 2 + 40 * 4 });
    const half = await merge({ base: base(), src: s, imports: [3, 5], rows: PLAN, options: OPT });
    expect(half.cell(4, 5)).toMatchObject({ f: 'SUMPRODUCT(src_S!C$2:C$6,src_S!D$2:D$6)', v: 30 * 3 + 10 * 1 + 50 * 5 + 20 * 2 + 40 * 4 });
    expect(half.trace.every((t) => t.viaCopy && t.exact && !t.widened)).toBe(true);
  });

  it('the formula\'s own cell and header: a reference to the header text of an imported column stays in the merged table', async () => {
    const s = src([{ h: '見出し', make: (r, id) => f(`C$1&":"&C${r}`, `点1:${P1[id]}`) }, { h: '他', make: (r) => f(`C${r}&レート!$A$2`) }]);
    const m = await merge({ base: base(), src: s, imports: [3, 4, 5, 6], rows: PLAN, options: OPT });
    expect(m.cell(4, 6)).toMatchObject({ f: 'D$1&":"&D4' });
    expect(m.cell(4, 7)).toMatchObject({ f: 'D4&src_レート!$A$2' });
  });
});

describe('join "full": appended rows read the copy at the source row they come from', () => {
  it('the people only the source has: their formulas keep reading their own source row', async () => {
    const ids = [...SRC_IDS, 'P6', 'P7'];
    const cols: Col[] = [
      { h: '税込', make: (r, id) => f(`C${r}*レート!$B$2`, P1[id] * 1.1) },
      { h: '前差', make: (r, id) => (r === 2 ? f('C2', 30) : f(`C${r}-C${r - 1}`, P1[id] - P1[ids[r - 3]])) },
      { h: '累計', make: (r) => f(`SUM(C$2:C${r})`, ids.slice(0, r - 1).reduce((a, id) => a + P1[id], 0)) },
    ];
    const plan: RowPlan[] = [...PLAN, { base: null, src: 7 }, { base: null, src: 8 }];
    const m = await merge({ base: base(), src: src(cols, ids), imports: [3, 4, 5, 6, 7], rows: plan, keyFill: [{ baseCol: 1, sourceCol: 1 }, { baseCol: 2, sourceCol: 2 }], options: OPT });
    expect(m.res.layout.appended.map((a) => a.row)).toEqual([7, 8]);
    expect(m.cell(7, 6)).toMatchObject({ f: 'D7*src_レート!$B$2', v: 60 * 1.1 });
    expect(m.cell(8, 6)).toMatchObject({ f: 'D8*src_レート!$B$2', v: 70 * 1.1 });
    // P7 (source row 8, merged row 8): its previous row is the source's row 7 = merged row 7: the merged table can follow it
    expect(m.cell(8, 7)).toMatchObject({ f: 'D8-D7', v: 10 });
    // P6 (source row 7): previous row = source row 6 (P4 -> merged row 5)
    expect(m.cell(7, 7)).toMatchObject({ f: 'D7-D5', v: 60 - 40 });
    expect(m.cell(8, 8)).toMatchObject({ f: 'SUM(D$2:D8)', v: 280 }); // the last row: the whole table (round 8)
    expect(m.cell(7, 8)).toMatchObject({ f: 'SUM(src_S!C$2:C7)', v: 210 }); // a middle row: the copy
    expect(m.report.formulasToValues).toBe(0);
  });
});

describe('priority and mixing', () => {
  it('formulas the merged table keeps are untouched by includeSourceSheets (same text, same results, no viaCopy)', async () => {
    const cols: Col[] = [
      { h: '合計', make: (r, id) => f(`C${r}+D${r}`, P1[id] + P2[id]) },
      { h: '最大', make: () => f('MAX(C$2:C$6)', 50) },
      { h: '順位', make: (r, id) => f(`RANK(C${r},C$2:C$6)`, 6 - P1[id] / 10) },
    ];
    const a = await merge({ base: base(), src: src(cols), imports: [3, 4, 5, 6, 7], rows: PLAN });
    const b = await merge({ base: base(), src: src(cols), imports: [3, 4, 5, 6, 7], rows: PLAN, options: OPT });
    for (const r of [2, 3, 4, 5, 6]) for (const c of [4, 5, 6, 7, 8]) expect(b.cell(r, c)).toEqual(a.cell(r, c));
    expect(b.trace.map((t) => ({ ...t })).every((t) => !t.viaCopy)).toBe(true);
    expect(b.report.formulaStats).toEqual(a.report.formulaStats);
    expect(b.items.some((i) => i.code === 'formulasViaCopy')).toBe(false);
  });

  it('a mixed formula: some references in the merged table, others in the copies', async () => {
    const s = src([{ h: '混在', make: (r, id) => f(`C${r}*レート!$B$2+D${r}`, P1[id] * 1.1 + P2[id]) }]);
    const m = await merge({ base: base(), src: s, imports: [3, 4, 5], rows: PLAN, options: OPT });
    expect(m.cell(4, 6)).toMatchObject({ f: 'D4*src_レート!$B$2+E4', v: 30 * 1.1 + 3 });
  });

  it('positional functions: ROW(ref) / OFFSET(ref) read the copies, ROW() alone stays a value', async () => {
    const s = src([
      { h: '行', make: (r) => f(`ROW(C${r})`, r) },
      { h: 'ずらし', make: (r, id) => f(`OFFSET(C${r},0,1)`, P2[id]) },
      { h: '行番号', make: (r) => f('ROW()-1', r - 1) },
    ]);
    const m = await merge({ base: base(), src: s, imports: [3, 4, 5, 6, 7], rows: PLAN, options: OPT });
    expect(m.cell(4, 6)).toMatchObject({ f: 'ROW(src_S!C2)', v: 2 });
    expect(m.cell(4, 7)).toMatchObject({ f: 'OFFSET(src_S!C2,0,1)', v: 3 });
    expect(m.cell(4, 8)?.f).toBeUndefined();
    expect(m.cell(4, 8)).toMatchObject({ v: 1 });
    expect(m.report.formulaStats![0]).toMatchObject({ kept: 10, toValues: 5, reasons: { viaCopy: 10, positional: 5 } });
  });

  it('a single-cell array formula that reads another sheet stays an array formula', async () => {
    const s = src([{ h: '配列', make: (r, id) => ({ f: `SUM(レート!$B$2:$B$3*C${r})`, v: P1[id] * 2, array: { ref: `E${r}` } }) }]);
    const m = await merge({ base: base(), src: s, imports: [3, 4, 5], rows: PLAN, options: OPT });
    expect(m.cell(4, 6)).toMatchObject({ f: 'SUM(src_レート!$B$2:$B$3*D4)', v: 60 });
    expect(m.report.formulaStats![0]).toMatchObject({ kept: 5, toValues: 0 });
  });

  it('a bare range of another sheet (implicit intersection) is not kept: a value, reason otherSheet', async () => {
    const s = src([{ h: '暗黙', make: (r, id) => f(`レート!$B$2:$B$3*C${r}`, P1[id] * 1.1) }]);
    const m = await merge({ base: base(), src: s, imports: [3, 4, 5], rows: PLAN, options: OPT });
    expect(m.report.formulaStats![0]).toMatchObject({ kept: 0, toValues: 5, reasons: { otherSheet: 5 } });
    expect(m.cell(4, 6)).toMatchObject({ v: 30 * 1.1 });
  });

  it('3-D references, external workbooks and sheets that do not exist stay values', async () => {
    const s = src([
      { h: '三次元', make: (r) => f(`SUM(レート:マスタ!B2)+C${r}`, 30) },
      { h: '外部', make: () => f("'[Other.xlsx]Sheet1'!A1", 5) },
      { h: '無い', make: () => f('Nope!A1', 7) },
    ]);
    const m = await merge({ base: base(), src: s, imports: [3, 4, 5, 6, 7], rows: PLAN, options: OPT });
    expect(m.report.formulaStats![0]).toMatchObject({ kept: 0, toValues: 15, reasons: { otherSheet: 10, unsafe: 5 } });
  });

  it('defined names: a name of the source that points at another sheet or a side cell is inlined as a reference to the copy', async () => {
    const s = src([{ h: '税込', make: (r, id) => f(`C${r}*税率`, P1[id] * 1.1) }, { h: '割引', make: (r, id) => f(`C${r}*割引率`, P1[id] * 0.9) }]);
    s.names = [{ name: '税率', refersTo: 'レート!$B$2' }, { name: '割引率', refersTo: 'S!$J$2' }];
    s.rows[1][9] = 0.9;
    const m = await merge({ base: base(), src: s, imports: [3, 4, 5, 6], rows: PLAN, options: OPT });
    expect(m.cell(4, 6)).toMatchObject({ f: 'D4*src_レート!$B$2', v: 33 });
    expect(m.cell(4, 7)).toMatchObject({ f: 'D4*src_S!$J$2', v: 27 });
    expect(m.report.formulaStats![0]).toMatchObject({ kept: 10, toValues: 0 });
  });
});

describe('several sources, names with "(1)", names shortened to 31 characters', () => {
  async function two(labels: [string, string], sheets: [string, string]) {
    const rates = (n: number): G[][] => [['項目', '値'], ['税率', n]];
    const mk = (sheet: string, rate: number): BookSpec => ({
      sheet,
      rows: [['ID', '氏名', '点1', '税込'], ...PEOPLE.map((p, i): G[] => [p, `名前${i + 1}`, (i + 1) * 10, { f: `C${i + 2}*マスタ!$B$2`, v: (i + 1) * 10 * rate }])],
      extra: [{ name: 'マスタ', rows: rates(rate) }],
    });
    const [bb, b1, b2] = await Promise.all([buildBook(base()), buildBook(mk(sheets[0], 2)), buildBook(mk(sheets[1], 3))]);
    const baseKs = await keySource('b', 'base', bb, { sheet: '名簿', headerRow: 1, firstDataRow: 2, lastDataRow: 6, firstCol: 1, lastCol: 3 });
    const s1 = await keySource('s1', labels[0], b1, { sheet: sheets[0], headerRow: 1, firstDataRow: 2, lastDataRow: 6, firstCol: 1, lastCol: 4 });
    const s2 = await keySource('s2', labels[1], b2, { sheet: sheets[1], headerRow: 1, firstDataRow: 2, lastDataRow: 6, firstCol: 1, lastCol: 4 });
    const rows: MergedRowPlan[] = PEOPLE.map((_p, i) => ({ baseRow: 2 + i, matches: { s1: { row: 2 + i, via: 'ID', flags: [] }, s2: { row: 2 + i, via: 'ID', flags: [] } } }));
    const input: PremergeInput = {
      base: baseKs, sources: [s1, s2], rows,
      imports: [{ sourceId: 's1', col: 3 }, { sourceId: 's1', col: 4 }, { sourceId: 's2', col: 3 }, { sourceId: 's2', col: 4 }],
      keyFill: {},
      options: { formulaPolicy: 'keep-if-possible', extendTable: true, headerConflict: 'suffix', includeSourceSheets: true },
    };
    const trace: FormulaTrace[] = [];
    const res = await buildPremerged(input, undefined, (t) => trace.push(t));
    const book = await readBook(res.bytes, 'm.xlsx', { id: 'm' });
    const sheet = book.sheets.find((s) => s.name === '名簿')!;
    return { res, trace, book, cell: (r: number, c: number) => sheet.rows[r - 1]?.[c - 1] };
  }

  it('two files with the same name and the same sheet names: the second one\'s copies get "(1)", and each file\'s formulas read their own copy', async () => {
    const m = await two(['面接', '面接'], ['評価', '評価']);
    expect(m.book.sheets.map((s) => s.name)).toEqual(['名簿', '面接_評価', '面接_マスタ', '面接_評価(1)', '面接_マスタ(1)']);
    // first file: columns D (点1) and E (税込); second: F and G
    expect(m.cell(2, 5)).toMatchObject({ f: 'D2*面接_マスタ!$B$2', v: 20 });
    expect(m.cell(2, 7)).toMatchObject({ f: "F2*'面接_マスタ(1)'!$B$2", v: 30 });
    expect(m.res.report.formulaStats).toEqual([
      { sourceId: 's1', label: '面接', kept: 5, toValues: 0, empty: 0, reasons: { viaCopy: 5 } },
      { sourceId: 's2', label: '面接', kept: 5, toValues: 0, empty: 0, reasons: { viaCopy: 5 } },
    ]);
    // the values behind the two copies are the two files' own: the rates 2 and 3
    const rate = (name: string) => m.book.sheets.find((s) => s.name === name)!.rows[1]?.[1];
    expect(rate('面接_マスタ')).toMatchObject({ v: 2 });
    expect(rate('面接_マスタ(1)')).toMatchObject({ v: 3 });
  });

  it('long file and sheet names are shortened to Excel\'s 31 characters; the formulas read exactly the shortened names', async () => {
    const long = '令和七年度第三回新卒採用面接評価シート（最終版）_確定';
    const m = await two([long, long], ['とても長い名前の評価シートです二十三文字', 'とても長い名前の評価シートです二十三文字']);
    const names = m.book.sheets.map((s) => s.name);
    for (const n of names) expect(n.length).toBeLessThanOrEqual(31);
    expect(new Set(names.map((n) => n.toLowerCase())).size).toBe(names.length);
    const refs = [m.cell(2, 5)!.f!, m.cell(2, 7)!.f!].map((x) => tokenize(x).find((t) => t.type === 'ref' && t.prefix)!.prefix!.sheet!);
    expect(refs[0]).not.toBe(refs[1]);
    for (const r of refs) expect(names).toContain(r);
    expect(m.res.report.formulasToValues).toBe(0);
  });

  it('the base\'s own sheets are never renamed: a copy that would collide with one gets the suffix', async () => {
    // a base sheet that is already called 面接_マスタ
    const b = base();
    b.extra = [{ name: '面接_マスタ', rows: [['x']] }];
    const rates: G[][] = [['項目', '値'], ['税率', 2]];
    const s: BookSpec = { sheet: '評価', rows: [['ID', '氏名', '点1', '税込'], ...PEOPLE.map((p, i): G[] => [p, `名前${i + 1}`, (i + 1) * 10, { f: `C${i + 2}*マスタ!$B$2`, v: (i + 1) * 20 }])], extra: [{ name: 'マスタ', rows: rates }] };
    const m = await merge({ base: b, src: s, imports: [3, 4], rows: [2, 3, 4, 5, 6].map((r) => ({ base: r, src: r })), options: OPT });
    expect(sheetNames(m)).toEqual(['名簿', '面接_マスタ', 'src_評価', 'src_マスタ']);
    expect(m.cell(2, 5)).toMatchObject({ f: 'D2*src_マスタ!$B$2' });
    // (the base's sheet keeps its name and its content)
    expect(m.book.sheets.find((x) => x.name === '面接_マスタ')!.rows[0]?.[0]).toMatchObject({ v: 'x' });
  });
});

describe('around the edges: dropOtherSheets, the cached values, big tables, hostile sources', () => {
  it('dropOtherSheets (the base\'s sheets) and includeSourceSheets together: the source\'s formulas still read the copies', async () => {
    const b = base();
    b.extra = [{ name: 'メモ', rows: [['x']] }];
    const s = src([{ h: '税込', make: (r, id) => f(`C${r}*レート!$B$2`, P1[id] * 1.1) }]);
    const m = await merge({ base: b, src: s, imports: [3, 4, 5], rows: PLAN, options: { includeSourceSheets: true, dropOtherSheets: true } });
    expect(sheetNames(m)).toEqual(['名簿', 'src_S', 'src_レート', 'src_マスタ']);
    expect(m.cell(4, 6)).toMatchObject({ f: 'D4*src_レート!$B$2', v: 33 });
    expect(m.report.droppedSheets).toEqual(['メモ']);
  });

  it('XlsxWriter-style placeholder results (0 everywhere): not imported as values, the formulas read the copies and Excel computes them', async () => {
    const s = src([{ h: '税込', make: (r) => f(`C${r}*レート!$B$2`, 0) }, { h: '合計', make: (r) => f(`C${r}+D${r}`, 0) }]);
    s.calcPr = '<calcPr calcId="191029" fullCalcOnLoad="1"/>';
    const m = await merge({ base: base(), src: s, imports: [3, 4, 5, 6], rows: PLAN, options: OPT });
    expect(m.cell(4, 6)).toMatchObject({ f: 'D4*src_レート!$B$2' });
    expect(m.cell(4, 6)!.v ?? null).toBeNull();
    expect(m.report.formulaStats![0]).toMatchObject({ kept: 10, toValues: 0, empty: 0 });
  });

  it('a big source (3,000 rows, formulas on another sheet, a side cell, a lookup into a 3,000-row master) merges fast and keeps every formula', async () => {
    const n = 3000;
    const ids = Array.from({ length: n }, (_x, i) => `Q${i + 1}`);
    const rows: G[][] = [['ID', '氏名', '点', '税込', '区分', '構成比'], ...ids.map((id, i): G[] => [id, `n${i}`, i + 1, f(`C${i + 2}*レート!$B$2`, (i + 1) * 1.1), f(`VLOOKUP(A${i + 2},マスタ!$A$2:$B$${n + 1},2,FALSE)`, i % 2 ? 'B' : 'A'), f(`C${i + 2}/$H$1`, 1)])];
    rows[0][7] = 5;
    const master: G[][] = [['ID', '区分'], ...ids.map((id, i): G[] => [id, i % 2 ? 'B' : 'A'])];
    const b: BookSpec = { sheet: '名簿', rows: [['ID', '氏名'], ...ids.map((id, i): G[] => [id, `n${i}`])] };
    const s: BookSpec = { sheet: 'S', rows, extra: [{ name: 'レート', rows: RATES }, { name: 'マスタ', rows: master }] };
    const t0 = Date.now();
    const m = await merge({ base: b, src: s, imports: [3, 4, 5, 6], rows: ids.map((_id, i) => ({ base: i + 2, src: i + 2 })), keyPairs: [{ baseCol: 1, sourceCol: 1 }], srcRegion: { lastCol: 6 }, baseRegion: { lastCol: 2 }, options: OPT });
    expect(Date.now() - t0).toBeLessThan(60_000);
    expect(m.report.formulaStats![0]).toMatchObject({ kept: 3 * n, toValues: 0, empty: 0, reasons: { viaCopy: 3 * n } });
    expect(m.cell(1500, 4)).toMatchObject({ f: 'C1500*src_レート!$B$2' });
    expect(m.cell(1500, 5)).toMatchObject({ f: `VLOOKUP(A1500,src_マスタ!$A$2:$B$${n + 1},2,FALSE)` });
    expect(m.cell(1500, 6)).toMatchObject({ f: 'C1500/src_S!$H$1' });
  }, 120_000);

  it('formulaPolicy "values": every formula is a value (reason policy), nothing reads the copies, the sheets are still copied', async () => {
    const s = src([{ h: '税込', make: (r, id) => f(`C${r}*レート!$B$2`, P1[id] * 1.1) }, { h: '合計', make: (r, id) => f(`C${r}+D${r}`, P1[id] + P2[id]) }]);
    const m = await merge({ base: base(), src: s, imports: [3, 4, 5, 6], rows: PLAN, options: { includeSourceSheets: true, formulaPolicy: 'values' } });
    expect(sheetNames(m)).toEqual(['名簿', 'src_S', 'src_レート', 'src_マスタ']);
    expect(m.report.formulaStats![0]).toMatchObject({ kept: 0, toValues: 10, reasons: { policy: 10 } });
    expect(m.trace.every((t) => t.outcome === 'value' && !t.viaCopy)).toBe(true);
    expect(m.cell(4, 6)).toMatchObject({ v: 33 });
    expect(m.cell(4, 6)?.f).toBeUndefined();
  });

  it('a source whose formulas refer to its data sheet by name and to itself through names: nothing is left reading a source sheet name', async () => {
    const s = src([{ h: '自分', make: (r, id) => f(`S!C${r}*S!$J$2`, P1[id]) }]);
    s.rows[1][9] = 1;
    const m = await merge({ base: base(), src: s, imports: [3, 4, 5], rows: PLAN, options: OPT });
    expect(m.cell(4, 6)).toMatchObject({ f: 'D4*src_S!$J$2', v: 30 });
  });

  it('every kept formula only reads sheets that exist in the output', async () => {
    const cols: Col[] = [
      { h: 'a', make: (r) => f(`C${r}*レート!$B$2`) },
      { h: 'b', make: (r) => f(`SUM(C$2:C${r})`) },
      { h: 'c', make: (r) => f(`E${r}+マスタ!$A$1&"x"`) },
      { h: 'd', make: (r) => f(`INDEX(マスタ!$B$2:$B$6,MATCH(A${r},マスタ!$A$2:$A$6,0))`) },
    ];
    const m = await merge({ base: base(), src: src(cols), imports: [3, 4, 5, 6, 7, 8], rows: PLAN, options: OPT });
    const names = new Set(sheetNames(m));
    let n = 0;
    for (const t of m.trace.filter((x) => x.outcome === 'kept')) {
      const c = m.cell(t.row, t.col)!;
      for (const tk of tokenize(c.f!)) if ((tk.type === 'ref' || tk.type === 'name' || tk.type === 'sref') && tk.prefix) { expect(names.has(tk.prefix.sheet!), c.f!).toBe(true); n++; }
    }
    expect(n).toBeGreaterThan(15);
  });
});

describe('after the premerge the merged sheet is an ordinary sheet: column edits adjust what it reads of itself, never what it reads of the copies', () => {
  it('moving / inserting / deleting columns of the merged sheet (the app\'s own engine): the references to the copies stay, the merged ones follow', async () => {
    const { replayLog } = await import('../../src/core/xlsx');
    const cols: Col[] = [
      { h: '税込', make: (r, id) => f(`C${r}*レート!$B$2`, P1[id] * 1.1) },
      { h: '合計', make: (r, id) => f(`C${r}+D${r}`, P1[id] + P2[id]) },
      { h: '側', make: (r, id) => f(`C${r}*$J$2`, P1[id]) },
    ];
    const s = src(cols);
    s.rows[1][9] = 1;
    const m = await merge({ base: base(), src: s, imports: [3, 4, 5, 6, 7], rows: PLAN, options: OPT, srcRegion: { lastCol: 7 } });
    const sheet = m.res.layout.sheet;
    expect(m.cell(4, 6)).toMatchObject({ f: 'D4*src_レート!$B$2' });
    // move the first imported column (D, 点1) behind the last one, then insert a column in front of everything
    const out = await replayLog(m.res.bytes, [
      { kind: 'struct', sheet, op: { type: 'move', axis: 'col', from: 4, count: 1, to: 9 } },
      { kind: 'struct', sheet, op: { type: 'insert', axis: 'col', at: 1, count: 1 } },
    ]);
    const book = await readBook(out, 'out.xlsx', { id: 'o' });
    const g = book.sheets.find((x) => x.name === sheet)!;
    // 点1 (D) went behind the last column (H) and the insert moved everything one to the right: it is column I; 点2 is E, 税込 F, 合計 G, 側 H
    const row = g.rows[3]!; // merged row 4 = source row 2 (P3)
    expect(row[8]).toMatchObject({ v: 30 });
    expect(row[5]).toMatchObject({ f: 'I4*src_レート!$B$2', v: 33 }); // the formula follows the moved column and still reads the copy
    expect(row[6]).toMatchObject({ f: 'I4+E4', v: 33 });
    expect(row[7]).toMatchObject({ f: 'I4*src_S!$J$2', v: 30 });
    // the copies are untouched: same sheets, same cells
    const names = book.sheets.map((x) => x.name);
    expect(names).toEqual([sheet, 'src_S', 'src_レート', 'src_マスタ']);
    expect(book.sheets.find((x) => x.name === 'src_S')!.rows[1]?.[9]).toMatchObject({ v: 1 });
    expect(book.sheets.find((x) => x.name === 'src_レート')!.rows[1]?.[1]).toMatchObject({ v: 1.1 });
  });
});
