/** detectRegion / pickDataSheet / the SourceBook → Table adapter on hand-built books. */
import { describe, expect, it } from 'vitest';
import { analyzeSources, defaultPlan, detectRegion, joinSources, pickDataSheet } from '../../src/core/keys';
import { buildTable, keyValue } from '../../src/core/keys/region';
import { type CellSpec, bookFromRows, sourceFromRows, sourceOf } from './helpers';

const book = (rows: CellSpec[][], name = 'Sheet1') => bookFromRows('t.xlsx', [{ name, rows }]);
const people = (n: number, from = 1): CellSpec[][] =>
  Array.from({ length: n }, (_, i) => [i + from, `A${String(1000 + i + from)}`, `山田 太郎${i + from}`, '東京大学', (i + 1) * 10]);
const HEAD: CellSpec[] = ['No', '応募者ID', '氏名', '大学', '点数'];

describe('detectRegion', () => {
  it('title rows above the header', () => {
    const b = bookFromRows('t.xlsx', [{ name: 'S', rows: [['2027卒 応募者名簿（人事部）'], [], HEAD, ...people(6)], merges: [{ r1: 1, c1: 1, r2: 1, c2: 5 }] }]);
    expect(detectRegion(b, 'S')).toEqual({ sheet: 'S', headerRow: 3, firstDataRow: 4, lastDataRow: 9, firstCol: 1, lastCol: 5 });
  });

  it('trailing 合計 rows (label or SUM formula) and blank rows are not data', () => {
    const b = book([HEAD, ...people(5), [], ['合計', null, null, null, 150], [null, null, null, null, { v: 150, f: 'SUM(E2:E6)' }]]);
    expect(detectRegion(b, 'Sheet1')).toMatchObject({ headerRow: 1, firstDataRow: 2, lastDataRow: 6 });
    const b2 = book([HEAD, ...people(5), [null, null, null, null, { v: 150, f: 'SUM(E2:E6)' }]]);
    expect(detectRegion(b2, 'Sheet1').lastDataRow).toBe(6);
    const b3 = book([HEAD, ...people(5), [null, 'Total', null, null, 150]]);
    expect(detectRegion(b3, 'Sheet1').lastDataRow).toBe(6);
  });

  it('a single blank row inside the list is kept; a note line or 3 blank rows end the table', () => {
    const b = book([HEAD, ...people(3), [], ...people(3, 4)]);
    expect(detectRegion(b, 'Sheet1').lastDataRow).toBe(8);
    const note = book([HEAD, ...people(4), [], ['※ 2026年4月1日時点の情報です']]);
    expect(detectRegion(note, 'Sheet1').lastDataRow).toBe(5);
    const gap = book([HEAD, ...people(4), [], [], [], ...people(2, 10)]);
    expect(detectRegion(gap, 'Sheet1').lastDataRow).toBe(5);
  });

  it('小計 rows inside the region stay in place but are never matched', () => {
    const rows: CellSpec[][] = [HEAD, ...people(2), ['小計', null, null, null, 30], ...people(2, 3)];
    const src = sourceFromRows('base', rows);
    expect(src.region.lastDataRow).toBe(6);
    const t = buildTable(src);
    expect(t.excelRow).toEqual([2, 3, 5, 6]);
    expect(t.skippedRows).toEqual([4]);
    const other = sourceFromRows('O', [['応募者ID', '評価'], ['A1001', 'x'], ['A1002', 'y'], ['A1003', 'z'], ['A1004', 'w']]);
    const an = analyzeSources(src, [other]);
    const res = joinSources(an, defaultPlan(an));
    expect(res.rows.map((r) => [r.baseRow, r.matches.O?.row ?? null])).toEqual([[2, 2], [3, 3], [4, null], [5, 4], [6, 5]]);
    expect(res.summary.baseRows).toBe(4);
  });

  it('header block: the contiguous block with the most dictionary hits; header-less data columns are bridged', () => {
    const b = book([
      ['メモ', null, '応募者ID', '氏名', null, '大学', '点数'],
      ['x', null, 'A1', '山田 太郎', 'ﾒﾓ', '東京大学', 10],
      [null, null, 'A2', '鈴木 花子', null, '京都大学', 20],
      [null, null, 'A3', '佐藤 翔太', 'ﾒﾓ', '大阪大学', 30],
    ]);
    expect(detectRegion(b, 'Sheet1')).toMatchObject({ headerRow: 1, firstCol: 3, lastCol: 7 });
  });

  it('merged header cells give their text to the whole merge', () => {
    const b = bookFromRows('t.xlsx', [{
      name: 'S',
      rows: [['応募者ID', '氏名', null, '大学'], ['A1', '山田', '太郎', '東京大学'], ['A2', '鈴木', '花子', '京都大学'], ['A3', '佐藤', '翔太', '大阪大学']],
      merges: [{ r1: 1, c1: 2, r2: 1, c2: 3 }],
    }]);
    const src = sourceOf(b, 'b', 'b');
    expect(src.region).toMatchObject({ firstCol: 1, lastCol: 4 });
    expect(buildTable(src).headers).toEqual(['応募者ID', '氏名', '氏名', '大学']);
  });

  it('single-column lists, empty and missing sheets', () => {
    expect(detectRegion(book([['氏名'], ['山田 太郎'], ['鈴木 花子']]), 'Sheet1')).toEqual({ sheet: 'Sheet1', headerRow: 1, firstDataRow: 2, lastDataRow: 3, firstCol: 1, lastCol: 1 });
    expect(detectRegion(book([]), 'Sheet1')).toEqual({ sheet: 'Sheet1', headerRow: 1, firstDataRow: 2, lastDataRow: 1, firstCol: 1, lastCol: 1 });
    expect(detectRegion(book([['a']]), 'none')).toEqual({ sheet: 'none', headerRow: 1, firstDataRow: 2, lastDataRow: 1, firstCol: 1, lastCol: 1 });
    const headerOnly = detectRegion(book([['応募者ID', '氏名']]), 'Sheet1');
    expect(headerOnly.lastDataRow).toBeLessThan(headerOnly.firstDataRow);
  });

  it('never throws on malformed sheets', () => {
    const b = book([HEAD, ...people(3)]);
    const s = b.sheets[0];
    s.maxRow = NaN as unknown as number;
    (s.rows as unknown[])[1] = undefined;
    (s.features as { merges: unknown }).merges = [{ r1: 5, c1: 9, r2: 2, c2: 1 }];
    expect(() => detectRegion(b, 'Sheet1')).not.toThrow();
    expect(() => pickDataSheet({ ...b, sheets: undefined as never })).not.toThrow();
  });
});

describe('footer rows (集計行) end the table', () => {
  // like samples/01: 5 people in rows 2–6, a blank row, 人数 / 平均 and 志望度4以上 rows
  const HEAD6: CellSpec[] = ['学生ID', '氏名', '大学', '年齢', '志望度', '重複チェック'];
  const person = (i: number): CellSpec[] => [
    `00010${i}`, `山田 太郎${i}`, '東京大学', 20 + i, 1 + (i % 5), { v: 1, f: `COUNTIF($B$2:$B$6,B${i + 2})` },
  ];
  const five = () => Array.from({ length: 5 }, (_, i) => person(i));
  const f = (v: number, formula: string) => ({ v, f: formula });

  it('人数 / 平均 / COUNTIF footers below a blank row (samples/01 layout)', () => {
    const b = book([
      HEAD6, ...five(), [],
      ['人数', f(5, 'COUNTA(B2:B6)'), null, '平均', f(22, 'AVERAGE(D2:D6)'), null],
      ['志望度4以上', f(2, 'COUNTIF(E2:E6,">=4")')],
    ]);
    expect(detectRegion(b, 'Sheet1')).toEqual({ sheet: 'Sheet1', headerRow: 1, firstDataRow: 2, lastDataRow: 6, firstCol: 1, lastCol: 6 });
  });

  it('footers directly below the data, below 2 blank rows, and label-less aggregate rows', () => {
    const direct = book([HEAD6, ...five(), ['平均', null, null, f(22, 'AVERAGE(D2:D6)'), f(3, 'AVERAGE(E2:E6)')]]);
    expect(detectRegion(direct, 'Sheet1').lastDataRow).toBe(6);
    const gap2 = book([HEAD6, ...five(), [], [], [null, f(5, 'COUNTA(B2:B6)'), null, f(110, 'SUM(D2:D6)')], [null, null, null, f(24, 'MAX(D2:D6)')]]);
    expect(detectRegion(gap2, 'Sheet1').lastDataRow).toBe(6);
    const stats = book([HEAD6, ...five(), [], ['最大', null, null, 24, 5], ['最小', null, null, 20, 1], ['Average', null, null, 22, 3]]);
    expect(detectRegion(stats, 'Sheet1').lastDataRow).toBe(6);
    const subtotal = book([HEAD6, ...five(), [null, null, '計', f(110, '_xlfn.AGGREGATE(9,6,D2:D6)')]]);
    expect(detectRegion(subtotal, 'Sheet1').lastDataRow).toBe(6);
  });

  it('per-row formulas are data: same-row SUM, running totals, COUNTIF over the rows above with the own row as criteria', () => {
    const HEAD: CellSpec[] = ['氏名', '国語', '数学', '英語', '合計点', '累計'];
    const rows: CellSpec[][] = [HEAD];
    for (let i = 0; i < 5; i++) {
      const r = i + 2;
      rows.push([`山田 太郎${i}`, 60 + i, 70 + i, 80 + i, f(210 + 3 * i, `SUM(B${r}:D${r})`), f(0, `SUM($E$2:E${r})`)]);
    }
    expect(detectRegion(book(rows), 'Sheet1').lastDataRow).toBe(6);
    // two-column list whose check column counts the rows above (own row = criteria)
    const narrow = book([['氏名', '既出'], ...Array.from({ length: 5 }, (_, i): CellSpec[] => [`佐藤 花子${i}`, f(0, `COUNTIF($A$1:A${i + 1},A${i + 2})`)])]);
    expect(detectRegion(narrow, 'Sheet1').lastDataRow).toBe(6);
    // data rows aggregating ANOTHER sheet (集計 sheet style) are data too
    const other = book([['大学', '人数', '平均'], ...['東京大学', '京都大学', '大阪大学'].map((u, i): CellSpec[] => [u, f(i, `COUNTIF(名簿!$D:$D,A${i + 2})`), f(3, `AVERAGEIF(名簿!$D$4:$D$63,A${i + 2},名簿!$K$4:$K$63)`)])]);
    expect(detectRegion(other, 'Sheet1').lastDataRow).toBe(4);
  });

  it('formulas filled down the column are data even in narrow lists: No =MAX($A$1:A1)+1, running totals', () => {
    // No column: 1, then =MAX($A$1:A2)+1 … (aggregate over the rows above, the classic auto-numbering)
    const numbered: CellSpec[][] = [['No', '氏名', '大学']];
    for (let i = 0; i < 6; i++) numbered.push([i === 0 ? 1 : f(i + 1, `MAX($A$1:A${i + 1})+1`), `山田 太郎${i}`, '東京大学']);
    expect(detectRegion(book(numbered), 'Sheet1')).toMatchObject({ firstDataRow: 2, lastDataRow: 7, lastCol: 3 });
    expect(buildTable(sourceOf(book(numbered), 'n', 'n')).excelRow).toEqual([2, 3, 4, 5, 6, 7]);
    const numbered2: CellSpec[][] = [['No', '氏名']];
    for (let i = 0; i < 6; i++) numbered2.push([f(i + 1, `MAX($A$1:A${i + 1})+1`), `山田 太郎${i}`]);
    expect(detectRegion(book(numbered2), 'Sheet1').lastDataRow).toBe(7);
    // …and the real footer below them is still found
    const withFooter = book([...numbered, [], ['人数', f(6, 'COUNTA(B2:B7)')]]);
    expect(detectRegion(withFooter, 'Sheet1').lastDataRow).toBe(7);
    const running: CellSpec[][] = [['氏名', '売上', '累計']];
    for (let i = 0; i < 6; i++) running.push([`山田 太郎${i}`, 10, f(10 * (i + 1), `SUM($B$2:B${i + 2})`)]);
    expect(detectRegion(book([...running, ['合計', f(60, 'SUM(B2:B7)')]]), 'Sheet1').lastDataRow).toBe(7);
  });

  it('a footer below three blank rows is outside the table; below one to two it is trimmed', () => {
    const gap3 = book([HEAD6, ...five(), [], [], [], ['人数', f(5, 'COUNTA(B2:B6)')]]);
    expect(detectRegion(gap3, 'Sheet1').lastDataRow).toBe(6);
    const gap1 = book([HEAD6, ...five(), [], ['合計', null, null, f(110, 'SUM(D2:D6)'), f(15, 'SUM(E2:E6)')]]);
    expect(detectRegion(gap1, 'Sheet1').lastDataRow).toBe(6);
    // the footer rows are never matched even when the user extends the region over them
    const t = buildTable({ id: 'x', book: gap3, region: detectRegion(gap3, 'Sheet1', { lastDataRow: 10 }), label: 'x' });
    expect(t.excelRow).toEqual([2, 3, 4, 5, 6]);
  });

  it('whole-column aggregates (=COUNTIF(C:C,…), =SUM(F:F)) are footers; per-row whole-column formulas are data', () => {
    const last = (rows: CellSpec[][]) => detectRegion(book(rows), 'Sheet1').lastDataRow;
    expect(last([HEAD6, ...five(), [null, null, null, f(5, 'COUNTIF(C:C,"東京大学")')]])).toBe(6);
    expect(last([HEAD6, ...five(), [], [null, null, null, f(5, 'COUNTIF(C:C,"東京大学")'), f(15, 'SUM(E:E)')]])).toBe(6);
    // a whole-column formula in its own column is circular, never a footer
    expect(last([HEAD6, ...five(), [null, null, null, f(110, 'SUM(D:D)')]])).toBe(7);
    // share of the total / duplicate check: the own row is referenced
    const share: CellSpec[][] = [['氏名', '売上', '構成比']];
    for (let i = 0; i < 4; i++) share.push([`山田 太郎${i}`, 10, f(0.25, `B${i + 2}/SUM(B:B)`)]);
    expect(last(share)).toBe(5);
    const dupchk: CellSpec[][] = [['氏名', '大学', '重複']];
    for (let i = 0; i < 4; i++) dupchk.push([`山田 太郎${i}`, '東京大学', f(1, `COUNTIF(A:A,A${i + 2})`)]);
    expect(last(dupchk)).toBe(5);
    // the same constant formula filled down every row
    const fill: CellSpec[][] = [['氏名', '大学', '大学別人数']];
    for (let i = 0; i < 4; i++) fill.push([`山田 太郎${i}`, '東京大学', f(4, 'COUNTIF(B:B,"東京大学")')]);
    expect(last(fill)).toBe(5);
  });

  it('a note line directly below the footer is not data', () => {
    const last = (rows: CellSpec[][]) => detectRegion(book(rows), 'Sheet1').lastDataRow;
    expect(last([HEAD6, ...five(), ['人数', f(5, 'COUNTA(B2:B6)')], ['※ 集計は4/1時点']])).toBe(6);
    expect(last([HEAD6, ...five(), [], ['人数', f(5, 'COUNTA(B2:B6)')], ['※ 集計は4/1時点'], ['※ 二行目']])).toBe(6);
    // 小計 followed by more data of the table stays inside it
    expect(last([HEAD6, ...five(), ['小計', null, null, 110, 15], ...five()])).toBe(12);
  });

  it('a 平均 / 最大 / Max value in an ordinary row is data: 区分 = 平均, first name "Max"; real Latin footers still end the table', () => {
    const last = (rows: CellSpec[][]) => detectRegion(book(rows), 'Sheet1').lastDataRow;
    expect(last([['First', 'Last', '年齢'], ['John', 'Smith', 30], ['Anna', 'Lee', 28], ['Max', 'Mueller', 35]])).toBe(4);
    expect(last([['氏名', '区分', '点数'], ['山田 太郎', '通常', 60], ['鈴木 花子', '通常', 70], ['佐藤 次郎', '平均', 80]])).toBe(4);
    expect(last([HEAD6, ...five(), [], ['Max', null, null, 24, 4], ['Min', null, null, 20, 1]])).toBe(6);
  });

  it('a list of statistics (項目 / 値) is data, not a footer', () => {
    const kv = book([['項目', '値'], ['人数', 60], ['平均', 20.9], ['最大', 24]]);
    expect(detectRegion(kv, 'Sheet1')).toMatchObject({ firstDataRow: 2, lastDataRow: 4 });
    const t = buildTable(sourceOf(kv, 'kv', 'kv'));
    expect(t.excelRow).toEqual([2, 3, 4]);
  });

  it('a table below the footer is not part of this table; groups with 小計 rows are', () => {
    const second = book([HEAD, ...people(4), ['合計', null, null, null, 100], [], ['応募者ID', '氏名', '面接日'], ['A1', '山田 太郎', '4/1']]);
    expect(detectRegion(second, 'Sheet1').lastDataRow).toBe(5);
    const groups = book([HEAD, ...people(3), ['小計', null, null, null, 60], [], ...people(3, 4), ['小計', null, null, null, 150], [], ['総計', null, null, null, 210]]);
    // rows 2–4, 小計 5, blank, rows 7–9, 小計 10, blank, 総計 12: the trailing 小計 / 総計 block is not data
    const r = detectRegion(groups, 'Sheet1');
    expect(r.lastDataRow).toBe(9);
    expect(buildTable(sourceOf(groups, 'g', 'g')).excelRow).toEqual([2, 3, 4, 7, 8, 9]);
  });

  it('Excel table totals rows and SUBTOTAL over a table column', () => {
    const b = book([HEAD, ...people(4), [null, null, null, null, { v: 100, f: 'SUBTOTAL(109,名簿表[点数])' }]]);
    expect(detectRegion(b, 'Sheet1').lastDataRow).toBe(5);
    const t = book([HEAD, ...people(4), [null, '4', null, null, 100]]);
    t.sheets[0].features.tables.push({ name: 'T', ref: 'A1:E6', headerRow: true, totalsRow: true, columns: ['No', '応募者ID', '氏名', '大学', '点数'] });
    expect(detectRegion(t, 'Sheet1').lastDataRow).toBe(5);
  });

  it('inside a user-extended region trailing footers are skipped; a labelled SUM line inside the table is data', () => {
    const b = book([
      HEAD6, ...five(), [],
      ['人数', f(5, 'COUNTA(B2:B6)'), null, '平均', f(22, 'AVERAGE(D2:D6)'), null],
      ['志望度4以上', f(2, 'COUNTIF(E2:E6,">=4")')],
    ]);
    const region = detectRegion(b, 'Sheet1', { lastDataRow: 9 });
    expect(region).toMatchObject({ firstDataRow: 2, lastDataRow: 9 });
    const t = buildTable({ id: 'x', book: b, region, label: 'x' });
    expect(t.excelRow).toEqual([2, 3, 4, 5, 6]);
    expect(t.skippedRows).toEqual([7, 8, 9]);
    // P/L style: 売上高合計 =SUM(...) between the lines is a row of its own (a key when merging by 科目)
    const pl = book([
      ['科目', '当期', '前期'],
      ['国内売上', 100, 90], ['海外売上', 50, 40], ['売上高合計', f(150, 'SUM(B2:B3)'), f(130, 'SUM(C2:C3)')],
      ['売上原価', 80, 70], ['販管費', 30, 25],
    ]);
    expect(buildTable(sourceOf(pl, 'pl', 'pl')).excelRow).toEqual([2, 3, 4, 5, 6]);
  });
});

describe('detectRegion options (header row / last row chosen by the user)', () => {
  const titled = () => bookFromRows('t.xlsx', [{ name: 'S', rows: [['2027卒 応募者名簿'], ['作成', '人事部'], HEAD, ...people(6), [], ['合計', null, null, null, 210]] }]);

  it('headerRow: columns from that row, data below it, same end rules', () => {
    const b = titled();
    expect(detectRegion(b, 'S')).toEqual({ sheet: 'S', headerRow: 3, firstDataRow: 4, lastDataRow: 9, firstCol: 1, lastCol: 5 });
    expect(detectRegion(b, 'S', { headerRow: 3 })).toEqual({ sheet: 'S', headerRow: 3, firstDataRow: 4, lastDataRow: 9, firstCol: 1, lastCol: 5 });
    // row 2 as the header: its two cells are the columns, rows 3.. are data (the real header becomes a data row)
    expect(detectRegion(b, 'S', { headerRow: 2 })).toEqual({ sheet: 'S', headerRow: 2, firstDataRow: 3, lastDataRow: 9, firstCol: 1, lastCol: 2 });
    // a data row as the header
    expect(detectRegion(b, 'S', { headerRow: 5 })).toMatchObject({ headerRow: 5, firstDataRow: 6, lastDataRow: 9, firstCol: 1, lastCol: 5 });
  });

  it('headerRow at / beyond the end, blank header rows, invalid values', () => {
    const b = titled();
    expect(detectRegion(b, 'S', { headerRow: 11 })).toMatchObject({ headerRow: 11, firstDataRow: 12, lastDataRow: 11 });
    expect(detectRegion(b, 'S', { headerRow: 50 })).toMatchObject({ headerRow: 50, firstDataRow: 51, lastDataRow: 50 });
    const blankHeader = bookFromRows('t.xlsx', [{ name: 'S', rows: [[], [null, 'A1', '山田 太郎', 10], [null, 'A2', '鈴木 花子', 20]] }]);
    expect(detectRegion(blankHeader, 'S', { headerRow: 1 })).toEqual({ sheet: 'S', headerRow: 1, firstDataRow: 2, lastDataRow: 3, firstCol: 2, lastCol: 4 });
    expect(detectRegion(b, 'S', { headerRow: 0 })).toEqual(detectRegion(b, 'S'));
    expect(detectRegion(b, 'S', { headerRow: NaN })).toEqual(detectRegion(b, 'S'));
    expect(detectRegion(b, 'none', { headerRow: 4 })).toMatchObject({ sheet: 'none', headerRow: 4, firstDataRow: 5, lastDataRow: 4 });
  });

  it('lastDataRow: taken as given (clamped to the sheet), no footer trimming', () => {
    const b = titled();
    expect(detectRegion(b, 'S', { lastDataRow: 11 })).toEqual({ sheet: 'S', headerRow: 3, firstDataRow: 4, lastDataRow: 11, firstCol: 1, lastCol: 5 });
    expect(detectRegion(b, 'S', { lastDataRow: 6 }).lastDataRow).toBe(6);
    expect(detectRegion(b, 'S', { lastDataRow: 500 }).lastDataRow).toBe(11);
    expect(detectRegion(b, 'S', { lastDataRow: 1 })).toMatchObject({ headerRow: 3, firstDataRow: 4, lastDataRow: 3 });
    expect(detectRegion(b, 'S', { headerRow: 2, lastDataRow: 7 })).toMatchObject({ headerRow: 2, firstDataRow: 3, lastDataRow: 7, lastCol: 2 });
  });
});

describe('pickDataSheet', () => {
  it('the list-like visible sheet wins; hidden sheets are ignored', () => {
    const b = bookFromRows('t.xlsx', [
      { name: '表紙', rows: [['2027卒 採用管理'], ['作成日', '2026/4/1']] },
      { name: '集計', rows: [['区分', '人数'], ['A', 3], ['B', 5]] },
      { name: '名簿', rows: [HEAD, ...people(20)] },
      { name: '旧名簿', rows: [HEAD, ...people(40)], state: 'hidden' },
    ]);
    expect(pickDataSheet(b)).toBe('名簿');
    expect(pickDataSheet(bookFromRows('e.xlsx', []))).toBe('');
    // only hidden sheets → still returns one
    expect(pickDataSheet(bookFromRows('h.xlsx', [{ name: 'H', rows: [HEAD, ...people(2)], state: 'hidden' }]))).toBe('H');
  });
});

describe('cell values', () => {
  it('contract cells → key values', () => {
    expect(keyValue(undefined)).toBeNull();
    expect(keyValue({ t: 'e', v: '#N/A' })).toBeNull();
    expect(keyValue({ t: 'z', v: null, f: 'A1' })).toBeNull();
    expect(keyValue({ t: 'n', v: 115 })).toBe(115);
    expect(keyValue({ t: 'b', v: true })).toBe(true);
    expect((keyValue({ t: 'd', v: 36982 }) as Date).toISOString()).toBe('2001-04-01T00:00:00.000Z');
    expect(keyValue({ t: 's', v: '00123' })).toBe('00123');
  });

  it('date cells are dates for 生年月日 composites (serial with a date format)', () => {
    const b = sourceFromRows('base', [
      ['氏名', '生年月日'],
      ['佐藤 翔太', new Date(Date.UTC(2001, 3, 1))],
      ['佐藤 翔太', new Date(Date.UTC(2002, 4, 2))],
      ['山田 太郎', new Date(Date.UTC(2001, 5, 3))],
      ['鈴木 花子', new Date(Date.UTC(2003, 6, 4))],
    ]);
    const o = sourceFromRows('O', [
      ['お名前', '誕生日', '評価'],
      ['佐藤 翔太', 'H14.5.2', 'A'],
      ['佐藤 翔太', '2001/4/1', 'B'],
      ['山田 太郎', '平成13年6月3日', 'C'],
      ['鈴木 花子', '2003-07-04', 'D'],
    ]);
    const an = analyzeSources(b, [o]);
    expect(an.pairs.O.candidates.find((c) => c.selected)?.label).toBe('氏名＋生年月日');
    const res = joinSources(an, defaultPlan(an));
    expect(res.rows.map((r) => r.matches.O?.row)).toEqual([3, 2, 4, 5]);
  });
});
