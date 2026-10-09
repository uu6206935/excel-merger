// Round-3 doc parity fixes: unmerge look (fill + borders of the top-left cell), the quote flag of typed text, moving a
// block inside a table to the gap right after its last column, constants typed into a table's totals row, customWidth
// (auto-widening only of auto widths) and the siblings of a CF block.
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { adjustAnchoredRule, parseAreaText, renderArea } from '../../src/core/formula/excelRef';
import { createEditor, insertColumns, moveColumns, setCellInput, setColumnWidth } from '../../src/core/doc';
import { readBook, replayLog } from '../../src/core/xlsx';
import { unmergeInDoc } from '../../src/core/doc/unmerge';
import { needsQuotePrefix, parseInput } from '../../src/core/doc/parse';
import type { FeatureRule, RangeBox, XBorderSide, XStyle } from '../../src/core/types';
import type { CellWrite } from '../../src/core/types';
import { golden } from './compare';
import { F, XF, at, ext, fAt, ids, makeEditor, ok, refused, roster } from './helpers';

const line = (style: XBorderSide['style'], color: string): XBorderSide => ({ style, color });
const TOP = line('thin', '#000000');
const BOTTOM = line('medium', '#FF0000');
const LEFT = line('thin', '#0000FF');
const RIGHT = line('dashed', '#00FF00');
/** a report title as openpyxl writes it: the style is on the top-left cell only */
const BAND: XStyle = { bold: true, size: 14, fill: '#FFC000', hAlign: 'center', indent: 1, border: { top: TOP, bottom: BOTTOM, left: LEFT, right: RIGHT } };
const PLAIN_TITLE: XStyle = { bold: true, size: 14, hAlign: 'center' };

/** 報告: title merged over `merge` in rows 1..; header row 3, data rows 4..6; the title cell A1 only carries the style */
function titled(merge: RangeBox = { r1: 1, c1: 1, r2: 1, c2: 5 }, styles: XStyle[] = [{}, BAND], cells: Record<string, unknown> = {}) {
  const c: Record<string, unknown> = { A1: { t: 's', v: '2026年度 名簿', s: 1 }, A3: '番号', B3: '氏名', C3: '点数', D3: '倍', E3: '備考', F3: '外', ...cells };
  for (let r = 4; r <= 6; r++) {
    c[`A${r}`] = r - 3;
    c[`B${r}`] = `名${r}`;
    c[`C${r}`] = r * 10;
    c[`D${r}`] = F(`C${r}*2`, r * 20);
  }
  return makeEditor({
    styles,
    sheets: [{ name: '報告', cells: c as never, maxCol: 6, features: { merges: [merge], tables: [], rules: [], arrays: [], hyperlinks: [] } }],
    region: { headerRow: 3, firstDataRow: 4, lastDataRow: 6 },
  });
}

// ---------------------------------------------------------------------------------------------
describe('unmerge: the top-left cell\'s fill and border sides go to the other cells of the area', () => {
  const styleAt = (s: { doc: { styles: XStyle[] } }, cell: { s?: number } | undefined): XStyle => s.doc.styles[cell?.s ?? 0];
  const CC = 'centerContinuous' as const;

  it('title band with the fill on the top-left cell only (centerAcross)', () => {
    const s0 = titled();
    const doc = unmergeInDoc(s0.doc, [{ r1: 1, c1: 1, r2: 1, c2: 5 }], true);
    expect(doc.features.merges).toEqual([]);
    const s = { doc };
    // the top-left cell keeps its own style (the indent goes with centerContinuous, like Excel)
    expect(styleAt(s, at(doc, 'A1'))).toEqual({ bold: true, size: 14, fill: '#FFC000', hAlign: CC, border: { top: TOP, bottom: BOTTOM, left: LEFT, right: RIGHT } });
    // the middle cells: fill + top + bottom (no inner left / right)
    for (const a of ['B1', 'C1', 'D1']) {
      expect(at(doc, a), a).toMatchObject({ t: 'z', v: null });
      expect(styleAt(s, at(doc, a)), a).toEqual({ fill: '#FFC000', hAlign: CC, border: { top: TOP, bottom: BOTTOM } });
    }
    // the last cell of the row closes the band with the right side
    expect(styleAt(s, at(doc, 'E1'))).toEqual({ fill: '#FFC000', hAlign: CC, border: { top: TOP, bottom: BOTTOM, right: RIGHT } });
    expect(at(doc, 'F1')).toBeUndefined();
    // the input doc is untouched
    expect(s0.doc.features.merges.length).toBe(1);
    expect(at(s0.doc, 'B1')).toBeUndefined();
    expect(s0.doc.styles.length).toBe(2);
  });

  it('the look is spread even without centerAcross (no alignment change)', () => {
    const doc = unmergeInDoc(titled().doc, [{ r1: 1, c1: 1, r2: 1, c2: 5 }], false);
    const s = { doc };
    expect(at(doc, 'A1')!.s).toBe(1); // untouched
    expect(styleAt(s, at(doc, 'C1'))).toEqual({ fill: '#FFC000', border: { top: TOP, bottom: BOTTOM } });
    expect(styleAt(s, at(doc, 'E1'))).toEqual({ fill: '#FFC000', border: { top: TOP, bottom: BOTTOM, right: RIGHT } });
  });

  it('a multi-row area: each side goes to the cells on that edge (top: first row, bottom: last row, left / right: first / last column)', () => {
    const doc = unmergeInDoc(titled({ r1: 1, c1: 1, r2: 2, c2: 3 }).doc, [{ r1: 1, c1: 1, r2: 2, c2: 3 }], false);
    const s = { doc };
    expect(styleAt(s, at(doc, 'A1'))).toMatchObject({ fill: '#FFC000' }); // the top-left cell keeps its own style
    expect(at(doc, 'A1')!.s).toBe(1);
    expect(styleAt(s, at(doc, 'A2'))).toEqual({ fill: '#FFC000', border: { bottom: BOTTOM, left: LEFT } });
    expect(styleAt(s, at(doc, 'B1'))).toEqual({ fill: '#FFC000', border: { top: TOP } });
    expect(styleAt(s, at(doc, 'B2'))).toEqual({ fill: '#FFC000', border: { bottom: BOTTOM } });
    expect(styleAt(s, at(doc, 'C1'))).toEqual({ fill: '#FFC000', border: { top: TOP, right: RIGHT } });
    expect(styleAt(s, at(doc, 'C2'))).toEqual({ fill: '#FFC000', border: { bottom: BOTTOM, right: RIGHT } });
  });

  it('keeps what the other cells have of their own and only adds what the top-left cell has', () => {
    const own: XStyle = { italic: true, numFmt: '0', border: { top: line('dotted', '#111111'), left: line('hair', '#222222') } };
    const s0 = titled({ r1: 1, c1: 1, r2: 1, c2: 5 }, [{}, BAND, own, { fill: '#00B050' }], { B1: { t: 'n', v: 7, s: 2 }, C1: { t: 'z', v: null, s: 3 } });
    const doc = unmergeInDoc(s0.doc, [{ r1: 1, c1: 1, r2: 1, c2: 5 }], true);
    const s = { doc };
    expect(at(doc, 'B1')).toMatchObject({ t: 'n', v: 7 });
    // own style + the top-left look: top replaced by the title's, bottom added, the inner left side stays
    expect(styleAt(s, at(doc, 'B1'))).toEqual({ italic: true, numFmt: '0', fill: '#FFC000', hAlign: CC, border: { top: TOP, bottom: BOTTOM, left: line('hair', '#222222') } });
    // the title's fill wins over the cell's own fill (the merged cell showed the top-left cell's fill)
    expect(styleAt(s, at(doc, 'C1'))).toEqual({ fill: '#FFC000', hAlign: CC, border: { top: TOP, bottom: BOTTOM } });
  });

  it('a title without fill or borders only gets the alignment; nothing at all without centerAcross', () => {
    const s0 = titled({ r1: 1, c1: 1, r2: 1, c2: 5 }, [{}, PLAIN_TITLE]);
    const c = unmergeInDoc(s0.doc, [{ r1: 1, c1: 1, r2: 1, c2: 5 }], true);
    expect(styleAt({ doc: c }, at(c, 'C1'))).toEqual({ hAlign: CC });
    const n = unmergeInDoc(s0.doc, [{ r1: 1, c1: 1, r2: 1, c2: 5 }], false);
    expect(n.features.merges).toEqual([]);
    expect(n.columns).toBe(s0.doc.columns); // no column was copied
    expect(at(n, 'C1')).toBeUndefined();
  });

  it('a pathological merge (over 100,000 cells) restyles only the cells that exist', () => {
    const s0 = titled({ r1: 1, c1: 1, r2: 30000, c2: 5 }, [{}, BAND], { C1: { t: 'n', v: 1, s: 0 } });
    const doc = unmergeInDoc(s0.doc, [{ r1: 1, c1: 1, r2: 30000, c2: 5 }], false);
    expect(styleAt({ doc }, at(doc, 'C1'))).toMatchObject({ fill: '#FFC000' });
    expect(at(doc, 'B1')).toBeUndefined();
  });

  it('end to end: moveColumns with unmergeTitles logs the op and leaves a visible band', () => {
    const s = ok(moveColumns(titled(), ['b2'], 4, { unmergeTitles: true }));
    expect(s.log[0]).toEqual({ kind: 'unmerge', sheet: '報告', refs: ['A1:E1'], centerAcross: true });
    for (const a of ['A1', 'B1', 'C1', 'D1', 'E1']) expect(styleAt(s, at(s.doc, a)).fill, a).toBe('#FFC000');
    expect(styleAt(s, at(s.doc, 'E1')).border).toEqual({ top: TOP, bottom: BOTTOM, right: RIGHT });
    expect(styleAt(s, at(s.doc, 'D1')).border).toEqual({ top: TOP, bottom: BOTTOM });
    expect(at(s.doc, 'F1')).toBeUndefined();
  });
});

// ---------------------------------------------------------------------------------------------
describe('quote flag (CellWrite.cell.quote)', () => {
  it('needsQuotePrefix: true for text Excel would re-parse, false for ordinary text', () => {
    const NOW = new Date(2026, 8, 30);
    const yes = ['00123', '123', '１２３', ' 42 ', '2026/10/20', '10/20', '50%', '¥1,000', '$5', '1,000', '1e3', '3/4', '9:30', 'R8.10.20', '令和8年10月20日', '2026年10月20日',
      'TRUE', 'false', 'ＴＲＵＥ', '#N/A', '=A1', '=1+1', '(100)', '-7', '1234567890123456789'];
    for (const t of yes) expect(needsQuotePrefix(t, undefined, NOW), t).toBe(true);
    const no = ['', '   ', 'abc', '東京大学', '03-1234-5678', '090-1234-5678', 'ABC-123', '12-34', '1.2.3', '1,23', '2026/2/30', '13/40', '=', 'A1', '1 2 3'];
    for (const t of no) expect(needsQuotePrefix(t, undefined, NOW), t).toBe(false);
    // a text-formatted cell keeps everything as text by itself
    for (const t of ['00123', '=A1', 'TRUE', '2026/10/20']) expect(needsQuotePrefix(t, '@', NOW), t).toBe(false);
    // the same helpers and the same "today" as typing: 2/29 is a date only in a leap year
    expect(needsQuotePrefix('2/29', undefined, new Date(2026, 0, 1))).toBe(false);
    expect(needsQuotePrefix('2/29', undefined, new Date(2028, 0, 1))).toBe(true);
    // the cell format matters for fractions only in which class the text falls into, never in whether it is re-parsed
    expect(needsQuotePrefix('3/4', '0.00', NOW)).toBe(true);
  });

  it('parseInput flags exactly the text results Excel would re-parse', () => {
    const P = (t: string, ctx = {}) => parseInput(t, { now: new Date(2026, 8, 30), ...ctx });
    expect(P("'00123")).toEqual({ kind: 'value', t: 's', v: '00123', quote: true });
    expect(P("'abc")).toEqual({ kind: 'value', t: 's', v: 'abc', quote: true }); // the prefix is kept whatever follows
    expect(P("''abc")).toEqual({ kind: 'value', t: 's', v: "'abc", quote: true });
    expect(P("'=A1")).toEqual({ kind: 'value', t: 's', v: '=A1', quote: true });
    expect(P('00123', { textIds: true })).toEqual({ kind: 'value', t: 's', v: '00123', quote: true });
    expect(P(' 00123 ', { textIds: true })).toMatchObject({ t: 's', v: ' 00123 ', quote: true });
    expect(P('03-1234-5678')).toEqual({ kind: 'value', t: 's', v: '03-1234-5678' });
    expect(P('abc')).toEqual({ kind: 'value', t: 's', v: 'abc' });
    expect(P('ID-001', { textIds: true })).toEqual({ kind: 'value', t: 's', v: 'ID-001' });
    expect(P('   ')).toEqual({ kind: 'value', t: 's', v: '   ' });
    // text-formatted cells: no prefix (the cell format keeps text; an apostrophe is content there)
    expect(P('00123', { numFmt: '@' })).toEqual({ kind: 'value', t: 's', v: '00123' });
    expect(P("'00123", { numFmt: '@' })).toEqual({ kind: 'value', t: 's', v: "'00123" });
    // numbers, dates, booleans are values, never flagged
    expect(P('00123')).toEqual({ kind: 'value', t: 'n', v: 123 });
    expect(P('TRUE')).toEqual({ kind: 'value', t: 'b', v: true });
    // > 15 digits stay text and are flagged
    expect(P('1234567890123456789')).toMatchObject({ t: 's', quote: true });
  });

  const writeOf = (s: { log: Array<{ kind: string }> }, i = s.log.length - 1): CellWrite =>
    (s.log[i] as unknown as { writes: CellWrite[] }).writes[0];

  it('a leading apostrophe: quote = true, stored text without the apostrophe, the grid shows the text', () => {
    const s = ok(setCellInput(roster(), 'b2', 3, "'00123"));
    expect(at(s.doc, 'B3')).toEqual({ t: 's', v: '00123' });
    expect(writeOf(s)).toEqual({ row: 3, col: 2, cell: { t: 's', v: '00123', quote: true } });
    expect(s.log).toHaveLength(1);
    const p = ok(setCellInput(roster(), 'b2', 4, "'2026/10/20"));
    expect(at(p.doc, 'B4')).toEqual({ t: 's', v: '2026/10/20' }); // not a date, no number format change
    expect(writeOf(p).cell).toEqual({ t: 's', v: '2026/10/20', quote: true });
    const a = ok(setCellInput(roster(), 'b2', 4, "'abc"));
    expect(writeOf(a).cell).toEqual({ t: 's', v: 'abc', quote: true });
  });

  it('digit-only input kept as text in an ID column is quoted; ordinary text is not', () => {
    const idCol = () =>
      roster((sp) => {
        for (let r = 2; r <= 6; r++) (sp.sheets[0].cells![`A${r}`] as { s?: number }).s = undefined; // General cells holding '00123'...
      });
    const s = ok(setCellInput(idCol(), 'b1', 4, '200'));
    expect(at(s.doc, 'A4')).toMatchObject({ t: 's', v: '200' });
    expect(writeOf(s).cell).toEqual({ t: 's', v: '200', quote: true });
    const t = ok(setCellInput(idCol(), 'b1', 4, 'A-200'));
    expect(writeOf(t).cell).toEqual({ t: 's', v: 'A-200' });
    expect('quote' in writeOf(t).cell!).toBe(false);
    // a number in an ordinary number column is a number
    expect(writeOf(ok(setCellInput(roster(), 'b4', 3, '200'))).cell).toEqual({ t: 'n', v: 200 });
  });

  it("ordinary text is never flagged ('03-1234-5678' is no date, 'abc' is text)", () => {
    for (const text of ['03-1234-5678', 'abc', '東京大学', '090-1234-5678']) {
      const s = ok(setCellInput(roster(), 'b2', 3, text));
      expect(writeOf(s).cell, text).toEqual({ t: 's', v: text });
    }
  });

  it('a text-formatted column keeps text without the prefix (the cell format does that job in Excel)', () => {
    const s = ok(setCellInput(roster(), 'b1', 4, '00200')); // column A is '@' in the roster
    expect(at(s.doc, 'A4')).toEqual({ t: 's', v: '00200', s: XF.text });
    expect(writeOf(s).cell).toEqual({ t: 's', v: '00200' });
    const apo = ok(setCellInput(roster(), 'b1', 4, "'00200"));
    expect(at(apo.doc, 'A4')).toEqual({ t: 's', v: "'00200", s: XF.text }); // the apostrophe is content in a text cell
    expect(writeOf(apo).cell).toEqual({ t: 's', v: "'00200" });
  });

  it('a quoted write is not a no-op (the doc cell cannot show whether the prefix is stored); plain re-typing still is', () => {
    const s0 = roster();
    const same = setCellInput(s0, 'b2', 3, '鈴木');
    expect(same.ok && same.snapshot).toBe(s0);
    const q = ok(setCellInput(s0, 'b2', 3, "'鈴木"));
    expect(q).not.toBe(s0);
    expect(writeOf(q).cell).toEqual({ t: 's', v: '鈴木', quote: true });
  });

  it('a table header typed with an apostrophe is the text without it (header cells are always text: no flag)', () => {
    const t = roster((sp) => {
      sp.sheets[0].features = {
        merges: [], rules: [], arrays: [], hyperlinks: [],
        tables: [{ name: '名簿T', ref: 'A1:E6', headerRow: true, totalsRow: false, columns: ['学生ID', '氏名', '大学', '点数', '倍'] }],
      };
    });
    const s = ok(setCellInput(t, 'b4', 1, "'2026"));
    expect(at(s.doc, 'D1')).toMatchObject({ t: 's', v: '2026' });
    expect(s.doc.features.tables[0].columns[3]).toBe('2026');
    expect(writeOf(s).cell).toEqual({ t: 's', v: '2026' });
  });

  it('formulas, numbers and dates never carry the flag; the doc display of numFmt typing is unchanged', () => {
    const s = ok(setCellInput(roster(), 'b6', 2, '2026/10/20'));
    expect(writeOf(s).cell).toEqual({ t: 'd', v: 46315, numFmt: 'yyyy/m/d' });
    expect(writeOf(ok(setCellInput(roster(), 'b6', 2, '=A1&"x"'))).cell).toEqual({ t: 'z', v: null, f: 'A1&"x"' });
  });
});

// ---------------------------------------------------------------------------------------------
describe('tables: a block inside the table may land in the gap right after its last column (an in-table reorder)', () => {
  /** 名簿 A1:E6 = table 名簿T (学生ID 氏名 大学 点数 倍); F, G imported; H 備考 */
  const tbl = (ref = 'A1:E6', columns = ['学生ID', '氏名', '大学', '点数', '倍']) =>
    roster((sp) => {
      sp.sheets[0].features = { merges: [], rules: [], arrays: [], hyperlinks: [], tables: [{ name: '名簿T', ref, headerRow: true, totalsRow: false, columns }] };
      for (let r = 2; r <= 6; r++) sp.sheets[0].cells![`E${r}`] = F('名簿T[[#This Row],[点数]]*2', 0);
    });

  it('moves one column to c2 + 1: the table columns reorder, the range stays', () => {
    const r = moveColumns(tbl(), ['b2'], 5); // 氏名 before F (= right after the table)
    const s = ok(r);
    expect(s.log).toEqual([{ kind: 'struct', sheet: '名簿', op: { type: 'move', axis: 'col', from: 2, count: 1, to: 6 } }]);
    expect(ids(s.doc).slice(0, 6)).toEqual(['b1', 'b3', 'b4', 'b5', 'b2', 'i:s2:3']);
    expect(s.doc.features.tables[0]).toMatchObject({ ref: 'A1:E6', columns: ['学生ID', '大学', '点数', '倍', '氏名'] });
    // the header cells travel with their columns: row 1 and the table columns agree
    expect([1, 2, 3, 4, 5].map((c) => s.doc.columns[c - 1].cells[0]?.v)).toEqual(['学生ID', '大学', '点数', '倍', '氏名']);
    // plain references follow (D moved to C), structured references do not change
    expect(ext(s.doc, '集計', 'B2')).toBe('SUM(名簿!C2:C6)');
    expect(fAt(s.doc, 'D2')).toBe('名簿T[[#This Row],[点数]]*2');
  });

  it('moves several columns, and non-contiguous selections (planMove: every block move is checked on its own)', () => {
    const two = ok(moveColumns(tbl(), ['b2', 'b3'], 5));
    expect(two.doc.features.tables[0]).toMatchObject({ ref: 'A1:E6', columns: ['学生ID', '点数', '倍', '氏名', '大学'] });
    const picked = ok(moveColumns(tbl(), ['b1', 'b3'], 5));
    expect(picked.doc.features.tables[0]).toMatchObject({ ref: 'A1:E6', columns: ['氏名', '点数', '倍', '学生ID', '大学'] });
    expect(picked.log.length).toBeGreaterThanOrEqual(1);
    expect(picked.log.every((e) => e.kind === 'struct' && e.op.type === 'move')).toBe(true);
    // the last column "moved" to the end is nothing
    const nothing = moveColumns(tbl(), ['b5'], 5);
    expect(nothing.ok && nothing.snapshot.log.length).toBe(0);
  });

  it('still refuses what Excel refuses: further than c2 + 1, before the first column, into the table, partly over it', () => {
    expect(refused(moveColumns(tbl(), ['b2'], 6))).toMatch(/テーブルの外/); // before G: out of the table
    expect(refused(moveColumns(tbl('B1:E6', ['氏名', '大学', '点数', '倍']), ['b3'], 0))).toMatch(/テーブルの外/); // before the first column A of a table at B
    expect(refused(moveColumns(tbl(), ['i:s2:3'], 2))).toMatch(/の中へ/); // from outside into the middle
    expect(refused(moveColumns(tbl(), ['b5', 'i:s2:3'], 0))).toMatch(/一部の列/); // non-contiguous selection: the first block is a partial cover
  });

  it('a block from outside landing right after the table does not extend it; the whole table may move', () => {
    const out = ok(moveColumns(tbl(), ['b6'], 5)); // 備考 (H) before F
    expect(out.doc.features.tables[0].ref).toBe('A1:E6');
    const whole = ok(moveColumns(tbl(), ['b1', 'b2', 'b3', 'b4', 'b5'], 8));
    expect(whole.doc.features.tables[0]).toMatchObject({ ref: 'D1:H6', columns: ['学生ID', '氏名', '大学', '点数', '倍'] });
  });

  it('a table that starts right of column A: c1 .. c2 + 1 is the legal range', () => {
    const t = () => tbl('B1:E6', ['氏名', '大学', '点数', '倍']);
    const a = ok(moveColumns(t(), ['b3'], 1)); // before the first table column
    expect(a.doc.features.tables[0]).toMatchObject({ ref: 'B1:E6', columns: ['大学', '氏名', '点数', '倍'] });
    const b = ok(moveColumns(t(), ['b3'], 5));
    expect(b.doc.features.tables[0]).toMatchObject({ ref: 'B1:E6', columns: ['氏名', '点数', '倍', '大学'] });
  });
});

// ---------------------------------------------------------------------------------------------
describe('table totals row: a typed constant becomes text (the column label), formulas stay formulas', () => {
  /** 名簿T A1:E7 with a totals row (row 7): A7 label, D7 = SUBTOTAL(109, 点数) */
  const totals = () =>
    roster((sp) => {
      sp.sheets[0].features = {
        merges: [], rules: [], arrays: [], hyperlinks: [],
        tables: [{ name: '名簿T', ref: 'A1:E7', headerRow: true, totalsRow: true, columns: ['学生ID', '氏名', '大学', '点数', '倍'] }],
      };
      const g = sp.sheets[0].cells!;
      g.A7 = { t: 's', v: '合計', s: XF.text };
      g.D7 = F('SUBTOTAL(109,名簿T[点数])', 350, { s: XF.int });
    });
  const writeOf = (s: { log: Array<{ kind: string }> }): CellWrite => (s.log[s.log.length - 1] as unknown as { writes: CellWrite[] }).writes[0];

  it('numbers, dates, percentages and booleans typed into a totals cell stay TEXT as typed', () => {
    const s0 = totals();
    const styleCount = s0.doc.styles.length;
    for (const [col, text] of [['b2', '123'], ['b2', '2026/10/20'], ['b3', '50%'], ['b2', 'TRUE'], ['b4', '¥1,000'], ['b2', '１，２３４'], ['b2', '#N/A']] as const) {
      const s = ok(setCellInput(s0, col, 7, text));
      const colIndex = Number(col.slice(1));
      const c = s.doc.columns[colIndex - 1].cells[6];
      expect(c, text).toMatchObject({ t: 's', v: text });
      expect(c!.f, text).toBeUndefined();
      expect(writeOf(s), text).toEqual({ row: 7, col: colIndex, cell: { t: 's', v: text } }); // no numFmt, no quote
      expect(s.doc.styles.length, text).toBe(styleCount); // the cell format is untouched
    }
  });

  it('warns that the value became a label; plain text does not', () => {
    const num = setCellInput(totals(), 'b2', 7, '123');
    expect(num.ok && num.warnings.join()).toMatch(/集計行/);
    const txt = setCellInput(totals(), 'b2', 7, '平均');
    expect(txt.ok && txt.warnings).toEqual([]);
  });

  it('an apostrophe-prefixed label is stored without the apostrophe; a text-formatted cell keeps the text', () => {
    expect(at(ok(setCellInput(totals(), 'b2', 7, "'123")).doc, 'B7')).toMatchObject({ t: 's', v: '123' });
    expect(writeOf(ok(setCellInput(totals(), 'b2', 7, "'123"))).cell).toEqual({ t: 's', v: '123' });
    expect(at(ok(setCellInput(totals(), 'b1', 7, '00123')).doc, 'A7')).toMatchObject({ t: 's', v: '00123' });
  });

  it('formulas stay formulas (SUBTOTAL of the own column, anything else); clearing clears', () => {
    const own = ok(setCellInput(totals(), 'b4', 7, '=SUBTOTAL(104,[点数])'));
    expect(at(own.doc, 'D7')).toMatchObject({ t: 'z', v: null, f: 'SUBTOTAL(104,名簿T[点数])' });
    expect(writeOf(own).cell).toEqual({ t: 'z', v: null, f: 'SUBTOTAL(104,名簿T[点数])' });
    const other = ok(setCellInput(totals(), 'b3', 7, '=SUM(D2:D6)'));
    expect(at(other.doc, 'C7')).toMatchObject({ f: 'SUM(D2:D6)' });
    const gone = ok(setCellInput(totals(), 'b4', 7, ''));
    expect(at(gone.doc, 'D7')).toMatchObject({ t: 'z', v: null });
    expect(writeOf(gone).cell).toBeNull();
  });

  it('a formula typed into a totals cell is never a dynamic array (the totals entry of the column)', () => {
    const s = ok(setCellInput(totals(), 'b3', 7, '=SEQUENCE(2)'));
    expect(at(s.doc, 'C7')).toEqual({ t: 'z', v: null, f: '_xlfn.SEQUENCE(2)' });
    expect(writeOf(s).cell).toEqual({ t: 'z', v: null, f: '_xlfn.SEQUENCE(2)' });
    expect(s.doc.features.arrays).toEqual([]);
    // in a data row of the same table it is one
    const d = ok(setCellInput(totals(), 'b3', 6, '=SEQUENCE(2)'));
    expect(at(d.doc, 'C6')).toMatchObject({ dyn: true, arr: 'C6' });
  });

  it('data rows and tables without a totals row are not affected', () => {
    const s = ok(setCellInput(totals(), 'b4', 6, '85'));
    expect(at(s.doc, 'D6')).toMatchObject({ t: 'n', v: 85 });
    const date = ok(setCellInput(totals(), 'b2', 6, '2026/10/20'));
    expect(at(date.doc, 'B6')).toMatchObject({ t: 'd', v: 46315 });
    const plain = roster((sp) => {
      sp.sheets[0].features = {
        merges: [], rules: [], arrays: [], hyperlinks: [],
        tables: [{ name: '名簿T', ref: 'A1:E6', headerRow: true, totalsRow: false, columns: ['学生ID', '氏名', '大学', '点数', '倍'] }],
      };
    });
    expect(at(ok(setCellInput(plain, 'b4', 6, '85')).doc, 'D6')).toMatchObject({ t: 'n', v: 85 });
    // the row below the table is not a totals row either
    expect(at(ok(setCellInput(totals(), 'b4', 8, '85')).doc, 'D8')).toMatchObject({ t: 'n', v: 85 });
  });
});

// ---------------------------------------------------------------------------------------------
describe('customWidth: only auto widths are widened by typed dates and numbers', () => {
  const withCol = (props: { width?: number; hidden?: boolean; customWidth?: boolean } | null) =>
    roster((sp) => {
      if (props) sp.sheets[0].colProps!.H = props;
      else delete sp.sheets[0].colProps!.H;
    });
  const widthAfter = (s0: ReturnType<typeof withCol>, text: string) => ok(setCellInput(s0, 'b6', 2, text)).doc.columns[7].props;

  it('an auto width (stored, no customWidth) is widened to fit, like the patcher and Excel', () => {
    const p = widthAfter(withCol({ width: 8.38 }), '2026/10/20');
    expect(p.width).toBe(11.375);
    expect(p.customWidth).toBeUndefined(); // the flag is left alone: a later, wider value may widen it again
    // a stored auto width that already fits is kept
    expect(widthAfter(withCol({ width: 15 }), '2026/10/20').width).toBe(15);
    // no <col> at all: same as before
    expect(widthAfter(withCol(null), '2026/10/20').width).toBe(11.375);
    // values that fit never widen
    expect(widthAfter(withCol({ width: 8.38 }), '9/30').width).toBe(8.38);
    expect(widthAfter(withCol({ width: 8.38 }), '50%').width).toBe(8.38);
  });

  it('a custom width is never touched; nor is a hidden column', () => {
    expect(widthAfter(withCol({ width: 8.38, customWidth: true }), '2026/10/20').width).toBe(8.38);
    expect(widthAfter(withCol({ width: 0, hidden: true, customWidth: true }), '2026/10/20').width).toBe(0);
    expect(widthAfter(withCol({ width: 8.38, hidden: true }), '2026/10/20').width).toBe(8.38);
  });

  it('the widening is computed from the column\'s own stored width, growing only when the text does not fit', () => {
    const s = ok(setCellInput(withCol({ width: 8.38 }), 'b6', 2, '2026/10/20'));
    expect(s.doc.columns[7].props.width).toBe(11.375);
    // a wider value in the same column widens it again (the patcher marks it bestFit: auto-fitted columns keep growing)
    const t = ok(setCellInput(s, 'b6', 3, '2026年10月20日'));
    expect(t.doc.columns[7].props.width).toBeGreaterThan(11.375);
    // a narrower one never shrinks it
    const u = ok(setCellInput(t, 'b6', 4, '10/20'));
    expect(u.doc.columns[7].props.width).toBe(t.doc.columns[7].props.width);
  });

  it('setColumnWidth marks the width as a custom width; an inserted column takes the flag from its left neighbour', () => {
    const s = ok(setColumnWidth(withCol({ width: 8.38 }), 'b6', 9));
    expect(s.doc.columns[7].props).toEqual({ width: 9, customWidth: true });
    expect(widthAfter(s, '2026/10/20').width).toBe(9);
    const ins = ok(insertColumns(roster(), 2, 1)); // right of B (custom width 12)
    expect(ins.doc.columns[2].props).toEqual({ width: 12, customWidth: true });
    const auto = ok(insertColumns(withCol({ width: 8.38 }), 8, 1)); // right of H (auto)
    expect(auto.doc.columns[8].props).toEqual({ width: 8.38 });
  });
});

// ---------------------------------------------------------------------------------------------
describe('CF blocks: the other rules of a block (siblings) take part in the move decisions', () => {
  const rule = (formulas: string[], extra: Partial<FeatureRule> = {}): FeatureRule => ({ sqref: 'B2:B6', formulas, kind: 'cf', type: 'expression', ...extra });
  const withRules = (rules: FeatureRule[]) =>
    roster((sp) => {
      sp.sheets[0].features = { merges: [], tables: [], arrays: [], hyperlinks: [], rules };
    });
  const sq = (s: { doc: { features: { rules: FeatureRule[] } } }) => s.doc.features.rules.map((r) => `${r.sqref} [${r.formulas.join(';')}]`);
  // 点数 (D) moves before 大学 (C): it lands right after the rules' area B2:B6
  const moveD = (s0: ReturnType<typeof withRules>) => ok(moveColumns(s0, ['b4'], 2));

  it('alone, the rule spreads onto the block that lands right after its area', () => {
    expect(sq(moveD(withRules([rule(['B2="x"'])])))).toEqual(['B2:C6 [B2="x"]']);
  });

  it('a sibling that refers to the moved block stops the spread for every rule of the block (one range object per block)', () => {
    const s = moveD(withRules([rule(['B2="x"']), rule(['$D2>5'])]));
    expect(sq(s)).toEqual(['B2:B6 [B2="x"]', 'B2:B6 [$C2>5]']);
    // and it is exactly what excelRef answers for the block
    const area = [parseAreaText('B2:B6')!];
    const op = { type: 'move', axis: 'col', from: 4, count: 1, to: 3 } as const;
    const want = adjustAnchoredRule({ sqref: area, formulas: ['B2="x"'] }, op, '名簿', { kind: 'cf', type: 'expression', siblings: [['$D2>5']] });
    expect(want.map((r) => r.sqref.map((a) => renderArea(a)).join(' '))).toEqual(['B2:B6']);
  });

  it('a sibling that does not refer to the moved cells changes nothing', () => {
    expect(sq(moveD(withRules([rule(['B2="x"']), rule(['$F2>5'])])))).toEqual(['B2:C6 [B2="x"]', 'B2:C6 [$F2>5]']);
  });

  it('only rules of the same CF block are siblings: same sqref, same storage; never a data validation', () => {
    const other = moveD(withRules([rule(['B2="x"']), rule(['$D2>5'], { sqref: 'B2:B7' })]));
    expect(sq(other)[0]).toBe('B2:C6 [B2="x"]');
    const x14 = moveD(withRules([rule(['B2="x"']), rule(['$D2>5'], { ext: true })]));
    expect(sq(x14)[0]).toBe('B2:C6 [B2="x"]');
    const dv = moveD(withRules([rule(['B2="x"']), { sqref: 'B2:B6', formulas: ['$D2>5'], kind: 'dv', type: 'custom' }]));
    expect(sq(dv)[0]).toBe('B2:C6 [B2="x"]');
    // spacing / $ in the sqref do not matter
    const spaced = moveD(withRules([rule(['B2="x"']), rule(['$D2>5'], { sqref: ' $B$2:$B$6 ' })]));
    expect(sq(spaced)[0]).toBe('B2:B6 [B2="x"]');
  });
  it('invariant: the doc and the package patcher fragment a two-rule block the same way (A moved before D, rule 2 reads A)', async () => {
    const base = golden('pm-base-plain.xlsx'); // 学生: ID 氏名 大学 点数, data rows 2..6
    const blockXml = (rules: string[]) => `<conditionalFormatting sqref="C2:C6">${rules.map((f, i) => `<cfRule type="expression" priority="${i + 1}"><formula>${f.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')}</formula></cfRule>`).join('')}</conditionalFormatting>`;
    const variant = async (rules: string[]) => {
      const z = await JSZip.loadAsync(base);
      const name = 'xl/worksheets/sheet1.xml';
      const xml = await z.file(name)!.async('string');
      z.file(name, xml.replace(/<conditionalFormatting[\s\S]*?<\/conditionalFormatting>/, blockXml(rules)));
      const bytes = await z.generateAsync({ type: 'uint8array' });
      const book = await readBook(bytes, 'cf.xlsx');
      const s0 = createEditor(book, { sheet: '学生', insertAt: 0, imported: [], appendAt: null, appended: [] }, {}, { headerRow: 1, firstDataRow: 2, lastDataRow: 6 });
      const s = ok(moveColumns(s0, ['b1'], 3)); // ID (the column rule 2 reads) lands between 大学 and 点数: right after the rules' area
      const after = await readBook(await replayLog(bytes, s.log), 'cf-after.xlsx');
      const file = after.sheets.find((x) => x.name === '学生')!.features.rules;
      const fmt = (l: FeatureRule[]) => l.map((r) => `${r.kind}|${r.type}|${r.sqref}|${r.formulas.join(';')}`).sort();
      return { doc: fmt(s.doc.features.rules), file: fmt(file) };
    };
    const block = await variant(['C2="東大"', '$A2<>""']);
    expect(block.doc).toEqual(block.file);
    const alone = await variant(['C2="東大"']);
    expect(alone.doc).toEqual(alone.file);
    // the sibling that reads the moved column keeps rule 1 from spreading onto it; alone it spreads
    expect(alone.doc).toEqual(['cf|expression|B2:C6|B2="東大"']);
    expect(block.doc).toContain('cf|expression|B2:B6|B2="東大"');
  });
});

// ---------------------------------------------------------------------------------------------
describe('a width the user sets ends auto-fit (the patcher drops bestFit, like Excel)', () => {
  // The patcher marks a column it auto-widened bestFit="1" customWidth="1" and keeps growing such columns; a width the
  // user sets afterwards (colWidth op) replaces that: bestFit is dropped, so a later wider date no longer widens it.
  // The doc says the same (ColProps.customWidth = user-set WITHOUT bestFit: never widened).
  it('a width set after an auto-widening is not widened again by the JS engine', async () => {
    const bytes = golden('g-table.xlsx'); // column J: an auto width (9, no customWidth)
    const book = await readBook(bytes, 'g-table.xlsx');
    let s = createEditor(book, { sheet: '名簿', insertAt: 0, imported: [], appendAt: null, appended: [] }, {}, { headerRow: 3, firstDataRow: 4, lastDataRow: 13 });
    const id = s.doc.columns[9].id;
    s = ok(setCellInput(s, id, 15, '2026年10月20日')); // widens J
    s = ok(setColumnWidth(s, id, 9)); // the user narrows it again
    s = ok(setCellInput(s, id, 16, '2026年10月20日')); // the doc keeps 9
    expect(s.doc.columns[9].props.width).toBe(9);
    const after = await readBook(await replayLog(bytes, s.log), 'after.xlsx');
    expect(after.sheets[0].colProps[9]?.width).toBe(9);
  });
});
