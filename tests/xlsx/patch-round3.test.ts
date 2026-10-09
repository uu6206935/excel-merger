/**
 * Round-3 patcher parity fixes (no Excel needed; the Excel comparisons are golden cases u-*, q-*, h2-totals-q, n-*,
 * tests/golden/cases.ts):
 *  1. 'unmerge' spreads the top-left cell's fill and border sides over the rest of each former area;
 *  2. CellWrite.cell.quote is honoured exactly (quotePrefix yes / no);
 *  3. tables: a block from inside a table may land right after its last column (in-table reorder, verified in Excel with
 *     native Insert Cut Cells: tests/fixtures/formula/table-move-excel-results.json), columns inserted into a table are
 *     named like Excel 365 does;
 *  4. totals-row writes; 5. the other rules of a CF block (siblings); 6. zip-bomb safety of the package.
 */
import fs from 'node:fs';
import path from 'node:path';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import type { CellWrite, LoggedOp, StructOp } from '../../src/core/types';
import { excelNewColumnNames as docNewColumnNames } from '../../src/core/doc/tables';
import { excelNewColumnNames } from '../../src/core/xlsx/patch/engine';
import { PatchError } from '../../src/core/xlsx/patch/pkg';
import { adjustUnit } from '../../src/core/xlsx/patch/rules';
import { findAll, getAttr } from '../../src/core/xlsx/patch/xml';
import { buildPremerged } from '../../src/core/xlsx/premerge';
import { replayLog } from '../../src/core/xlsx/replay';
import { readFixture } from '../formula/helpers';
import { buildXlsx, part } from '../golden/build';
import { GOLDEN_FIX, scenario } from '../golden/premerge-scenario';
import { snapshot } from '../golden/snapshot';
import { R3_STYLES, buildR3Fixture } from '../golden/tools/make-r3-fixture';

const JP = /[ぁ-んァ-ン一-龥]/;
const read = (f: string) => new Uint8Array(fs.readFileSync(path.join(GOLDEN_FIX, f)));
const cells = (sheet: string, writes: CellWrite[], reason: 'edit' | 'toValues' = 'edit'): LoggedOp => ({ kind: 'cells', sheet, writes, reason });
const w = (row: number, col: number, cell: CellWrite['cell']): CellWrite => ({ row, col, cell });
const unmerge = (sheet: string, refs: string[], centerAcross = false): LoggedOp => ({ kind: 'unmerge', sheet, refs, centerAcross });
const struct = (sheet: string, op: StructOp): LoggedOp => ({ kind: 'struct', sheet, op });
const insC = (sheet: string, at: number, count = 1) => struct(sheet, { type: 'insert', axis: 'col', at, count });
const movC = (sheet: string, from: number, count: number, to: number) => struct(sheet, { type: 'move', axis: 'col', from, count, to });

const sheetOf = async (snap: ReturnType<typeof snapshot>, name: string) => (await snap).sheets.find((s) => s.name === name)!;

/** fill element + border sides ('LRTB' order) of a snapshot style signature */
function look(sig: string | undefined): { fill: string; sides: string } {
  if (!sig) return { fill: '', sides: '' };
  const fill = /patternType="solid">(<fgColor[^>]*>)/.exec(sig)?.[1] ?? '';
  const border = /border=(<border[^>]*>.*?<\/border>)/.exec(sig)?.[1] ?? '';
  const sides = (['left', 'right', 'top', 'bottom'] as const).filter((s) => new RegExp(`<${s} style=`).test(border)).map((s) => s[0] === 'l' ? 'L' : s[0] === 'r' ? 'R' : s[0] === 't' ? 'T' : 'B').join('');
  return { fill, sides };
}
/** the style element (<xf>) a cell of sheet `n` uses */
async function xfOf(bytes: Uint8Array, n: number, ref: string): Promise<string> {
  const xml = (await part(bytes, `xl/worksheets/sheet${n}.xml`))!;
  const s = Number(new RegExp(`<c r="${ref}"[^>]*?\\ss="(\\d+)"`).exec(xml)?.[1] ?? 0);
  const styles = (await part(bytes, 'xl/styles.xml'))!;
  const xfs = findAll(/<cellXfs\b[\s\S]*?<\/cellXfs>/.exec(styles)![0], 'xf');
  return xfs[s].text;
}
const xfCount = async (bytes: Uint8Array) => Number(/<cellXfs count="(\d+)"/.exec((await part(bytes, 'xl/styles.xml'))!)![1]);
const cellXml = async (bytes: Uint8Array, n: number, ref: string) => new RegExp(`<c r="${ref}"[^>]*?(?:/>|>[\\s\\S]*?</c>)`).exec((await part(bytes, `xl/worksheets/sheet${n}.xml`))!)?.[0];

// ---------------------------------------------------------------------------------------------------------------
describe("'unmerge' spreads the top-left cell's fill and border sides over the former area", () => {
  it('a title band styled on its top-left cell only (fill): every other cell shows the same fill after the unmerge (A1:K1)', async () => {
    const band = await buildXlsx([{
      name: 'S',
      data: '<row r="1"><c r="A1" s="4" t="s"><v>0</v></c></row>',
      tail: '<mergeCells count="1"><mergeCell ref="A1:K1"/></mergeCells>',
    }], { sst: ['タイトル'], styles: R3_STYLES });
    const out = await replayLog(band, [unmerge('S', ['A1:K1'])]);
    const s = await sheetOf(snapshot(out), 'S');
    expect(s.merges).toEqual([]);
    const a1 = look(s.styles.A1);
    expect(a1.fill).toContain('theme="4"');
    expect(s.styles.A1).toContain('horizontal="center"'); // the top-left cell itself is untouched
    for (const c of 'BCDEFGHIJK') {
      expect(look(s.styles[`${c}1`]), c).toEqual({ fill: a1.fill, sides: '' });
      // only the fill: no font / alignment of the top-left cell
      expect(s.styles[`${c}1`], c).not.toContain('<alignment');
      expect(s.styles[`${c}1`], c).not.toContain('<b/>');
    }
    // the cells exist in the sheet (styled, empty) and share one xf
    const xml = (await part(out, 'xl/worksheets/sheet1.xml'))!;
    const ids = new Set([...xml.matchAll(/<c r="[B-K]1" s="(\d+)"\/>/g)].map((m) => m[1]));
    expect([...xml.matchAll(/<c r="[B-K]1" s="\d+"\/>/g)].length).toBe(10);
    expect(ids.size).toBe(1);
    expect(xml).not.toContain('mergeCell');
    expect(s.values.A1).toBe('s:タイトル');
  });

  it('with centerAcross the fill comes on top of centerContinuous (every cell)', async () => {
    const band = await buildXlsx([{
      name: 'S',
      data: '<row r="1"><c r="A1" s="4" t="s"><v>0</v></c></row>',
      tail: '<mergeCells count="1"><mergeCell ref="A1:D1"/></mergeCells>',
    }], { sst: ['タイトル'], styles: R3_STYLES });
    const s = await sheetOf(snapshot(await replayLog(band, [unmerge('S', ['A1:D1'], true)])), 'S');
    const a1 = look(s.styles.A1);
    for (const c of 'BCD') {
      expect(look(s.styles[`${c}1`]), c).toEqual({ fill: a1.fill, sides: '' });
      expect(s.styles[`${c}1`], c).toContain('horizontal="centerContinuous"');
    }
    expect(s.styles.A1).toContain('horizontal="centerContinuous"');
    expect(s.styles.A1).toContain('vertical="center"'); // the top-left cell keeps its other alignment settings
  });

  it('a bordered area: each side goes to the cells on that edge, the cells keep their own number formats', async () => {
    // Band!A3:C5: thin box + yellow fill on A3; B4 has a date format, C5 its own thin box + 0.0"pt" format
    const out = await replayLog(await buildR3Fixture(), [unmerge('Band', ['A3:C5'])]);
    const s = await sheetOf(snapshot(out), 'Band');
    const yellow = look(s.styles.A3).fill;
    expect(yellow).toContain('FFFFFF00');
    const want: Record<string, string> = {
      A3: 'LRTB', B3: 'T', C3: 'RT',
      A4: 'L', B4: '', C4: 'R',
      A5: 'LB', B5: 'B', C5: 'LRTB',
    };
    for (const [ref, sides] of Object.entries(want)) expect(look(s.styles[ref]), ref).toEqual({ fill: yellow, sides });
    // own styles survive: B4 a date, C5 the number format
    expect(s.styles.B4).toMatch(/^nf=#14 /);
    expect(s.styles.C5).toMatch(/^nf=0\.0"pt" /);
    expect(s.values.B4).toBe('n:46295');
    expect(s.values.C5).toBe('n:12.5');
    // the other merges are untouched
    expect(s.merges).toContain('A1:F1');
    expect(s.merges).not.toContain('A3:C5');
  });

  it('a one-row band with a box on its first cell: top / bottom on every cell, right on the last, fill everywhere', async () => {
    const s = await sheetOf(snapshot(await replayLog(await buildR3Fixture(), [unmerge('Band', ['A1:F1'])])), 'Band');
    const yellow = look(s.styles.A1).fill;
    expect(look(s.styles.A1)).toEqual({ fill: yellow, sides: 'LRTB' }); // unchanged
    for (const c of 'BCDE') expect(look(s.styles[`${c}1`]), c).toEqual({ fill: yellow, sides: 'TB' });
    expect(look(s.styles.F1)).toEqual({ fill: yellow, sides: 'RTB' });
  });

  it('an outline as Excel stores it (left / top / bottom on the first cell): only those sides, per row and edge', async () => {
    const s = await sheetOf(snapshot(await replayLog(await buildR3Fixture(), [unmerge('Outline', ['A1:F1', 'A3:C5'])])), 'Outline');
    const yellow = look(s.styles.A1).fill;
    expect(look(s.styles.A1)).toEqual({ fill: yellow, sides: 'LTB' });
    for (const c of 'BCDEF') expect(look(s.styles[`${c}1`]), c).toEqual({ fill: yellow, sides: 'TB' });
    // A3:C5 (left / top): top on the first row, left on the first column of every row, nothing at the bottom / right
    const want: Record<string, string> = { A3: 'LT', B3: 'T', C3: 'T', A4: 'L', B4: '', C4: '', A5: 'L', B5: '', C5: '' };
    for (const [ref, sides] of Object.entries(want)) expect(look(s.styles[ref]), ref).toEqual({ fill: yellow, sides });
  });

  it('a top-left cell with an underline only: the underline (and the theme fill) on every cell, with centerContinuous on top', async () => {
    const s = await sheetOf(snapshot(await replayLog(await buildR3Fixture(), [unmerge('Outline', ['A7:C7'], true)])), 'Outline');
    const tl = look(s.styles.A7);
    expect(tl.sides).toBe('B');
    expect(tl.fill).toContain('theme="4"');
    for (const c of 'BC') {
      expect(look(s.styles[`${c}7`]), c).toEqual(tl);
      expect(s.styles[`${c}7`], c).toContain('horizontal="centerContinuous"');
      expect(s.styles[`${c}7`], c).toContain('<bottom style="medium"><color rgb="FFFF0000"/></bottom>');
    }
  });

  it('a top-left cell without fill / border changes nothing (no cells created, other cells keep their own look)', async () => {
    const input = await buildR3Fixture();
    const out = await replayLog(input, [unmerge('Band', ['A9:C9'])]);
    const s = await sheetOf(snapshot(out), 'Band');
    expect(s.merges).not.toContain('A9:C9');
    expect(look(s.styles.B9).fill).toContain('FFDDEBF7'); // B9's own fill stays
    expect(s.styles.C9).toBeUndefined();
    expect(await part(out, 'xl/styles.xml')).toBe(await part(input, 'xl/styles.xml'));
  });

  it("a side of the top-left cell REPLACES the cell's own side; sides it does not have leave the cell's own alone", async () => {
    // Band!A11:C11: box on A11; C11 has a medium red bottom of its own -> thin (the top-left side), right / top added
    const s = await sheetOf(snapshot(await replayLog(await buildR3Fixture(), [unmerge('Band', ['A11:C11'])])), 'Band');
    expect(look(s.styles.C11).sides).toBe('RTB');
    expect(s.styles.C11).toContain('<bottom style="thin">');
    expect(look(s.styles.B11).sides).toBe('TB');
    // no look spread from an area whose top-left cell has none: C11's own medium red underline stays in u-nolook-like areas
    const only = await sheetOf(snapshot(await replayLog(await buildR3Fixture(), [unmerge('Band', ['A9:C9'])])), 'Band');
    expect(only.styles.C11).toContain('<bottom style="medium">');
  });

  it('a missing cell takes the row format (customFormat) first, like a cell typed there: its number format stays', async () => {
    const book = await buildXlsx([{
      name: 'S',
      data: '<row r="1" s="3" customFormat="1"><c r="A1" s="4" t="s"><v>0</v></c></row>',
      tail: '<mergeCells count="1"><mergeCell ref="A1:C1"/></mergeCells>',
    }], { sst: ['x'], styles: R3_STYLES });
    const s = await sheetOf(snapshot(await replayLog(book, [unmerge('S', ['A1:C1'])])), 'S');
    for (const c of 'BC') {
      expect(s.styles[`${c}1`], c).toMatch(/^nf=#14 /);
      expect(look(s.styles[`${c}1`]).fill, c).toContain('theme="4"');
    }
  });

  it('several areas, a cell of the area that holds a value, and an unknown area', async () => {
    const book = await buildXlsx([{
      name: 'S',
      data: '<row r="1"><c r="A1" s="1" t="s"><v>0</v></c><c r="C1"><v>42</v></c></row><row r="3"><c r="A3" s="1" t="s"><v>0</v></c></row>',
      tail: '<mergeCells count="2"><mergeCell ref="A1:C1"/><mergeCell ref="A3:B4"/></mergeCells>',
    }], { sst: ['t'], styles: R3_STYLES });
    const s = await sheetOf(snapshot(await replayLog(book, [unmerge('S', ['A1:C1', 'A3:B4', 'F9:G9'])])), 'S');
    expect(s.merges).toEqual([]);
    expect(s.values.C1).toBe('n:42'); // a hidden value stays
    expect(look(s.styles.C1).sides).toBe('RTB');
    expect(look(s.styles.B4).sides).toBe('RB');
    expect(look(s.styles.A4).sides).toBe('LB');
  });
});

// ---------------------------------------------------------------------------------------------------------------
describe('CellWrite.cell.quote: honoured exactly', () => {
  it("quote === true -> an xf with quotePrefix (derived from the cell's own xf); anything else -> none", async () => {
    const input = await buildR3Fixture();
    const before = await xfCount(input);
    const out = await replayLog(input, [
      cells('Q', [
        w(2, 1, { t: 's', v: '00123', quote: true }), // General cell
        w(3, 1, { t: 's', v: '=text', quote: true }),
        w(4, 1, { t: 's', v: '03-1234-5678' }), // quote false: plain text
        w(5, 1, { t: 's', v: '00123', quote: true }), // A5 is styled (bold, yellow, box)
        w(7, 1, { t: 's', v: '00123', quote: false }),
      ]),
    ]);
    const xf0 = await xfOf(out, 2, 'A2');
    expect(xf0).toMatch(/quotePrefix="1"/);
    expect(xf0).toMatch(/numFmtId="0" fontId="0" fillId="0" borderId="0"/);
    expect(await xfOf(out, 2, 'A3')).toBe(xf0);
    expect(await xfOf(out, 2, 'A4')).not.toContain('quotePrefix');
    expect(await xfOf(out, 2, 'A7')).not.toContain('quotePrefix');
    expect(await cellXml(out, 2, 'A4')).not.toMatch(/\ss="/); // not even styled
    const xf5 = await xfOf(out, 2, 'A5');
    expect(xf5).toMatch(/quotePrefix="1"/);
    expect(xf5).toMatch(/fontId="1" fillId="2" borderId="1"/); // bold, yellow fill, thin box kept
    // one new xf only: General + prefix already exists in the fixture (xf 7, reused), styled + prefix is new
    expect(await cellXml(out, 2, 'A2')).toContain(' s="7"');
    expect(await xfCount(out)).toBe(before + 1);
    // the text itself is unchanged and stays a shared string
    const s = await sheetOf(snapshot(out), 'Q');
    expect(s.values).toMatchObject({ A2: 's:00123', A3: 's:=text', A4: 's:03-1234-5678', A5: 's:00123', A7: 's:00123' });
    expect(s.styles.A2).toMatch(/ quote$/);
    expect(s.styles.A4).toBeUndefined();
    expect(s.styles.A7).toBeUndefined();
  });

  it('a quote-prefixed cell overwritten by a non-quote write loses the prefix (the rest of its xf stays); a clear keeps the style', async () => {
    const input = await buildR3Fixture();
    // Q!B1 '00123' (s=7: quote), B2 '=text' (s=7), B3 '00456' (s=7), E1 '007' (s=9: text format + quote)
    const out = await replayLog(input, [cells('Q', [w(1, 2, { t: 's', v: 'abc' }), w(2, 2, { t: 'n', v: 5 }), w(3, 2, null), w(1, 5, { t: 's', v: 'abc' })])]);
    const s = await sheetOf(snapshot(out), 'Q');
    expect(s.values).toMatchObject({ B1: 's:abc', B2: 'n:5' });
    expect(s.styles.B1).toBeUndefined(); // xf 0 again
    expect(s.styles.B2).toBeUndefined();
    expect(s.styles.E1).toMatch(/^nf=#49 /); // the text format stays ...
    expect(s.styles.E1).not.toMatch(/ quote$/); // ... the prefix is gone
    // cleared: value gone, style (with the prefix) kept: types.ts "null clears the value / formula (style kept)"
    expect(s.values.B3).toBeUndefined();
    expect(s.styles.B3).toMatch(/ quote$/);
    expect(await xfOf(out, 2, 'E1')).toBe((await xfOf(input, 2, 'C1'))); // the same xf as the plain text-format cell C1
  });

  it("quote with a number format: '@' + quote -> text format AND prefix; the prefix is only for text constants", async () => {
    const out = await replayLog(await buildR3Fixture(), [cells('Q', [
      w(1, 3, { t: 's', v: '00123', numFmt: '@' }),
      w(1, 4, { t: 's', v: '00123', numFmt: '@', quote: true }),
      w(10, 1, { t: 'n', v: 5, quote: true }), // not text: ignored
      w(11, 1, { t: 'z', v: null, f: 'A1+1', quote: true }), // formula: ignored
      w(12, 1, { t: 's', v: 'x', f: 'A1&"x"', quote: true }), // formula with a string result: ignored
      w(13, 1, { t: 'b', v: true, quote: true }),
    ])]);
    const s = await sheetOf(snapshot(out), 'Q');
    expect(s.styles.C1).toMatch(/^nf=#49 /);
    expect(s.styles.C1).not.toMatch(/ quote$/);
    expect(s.styles.D1).toMatch(/^nf=#49 .* quote$/);
    for (const ref of ['A10', 'A11', 'A12', 'A13']) expect(s.styles[ref], ref).toBeUndefined();
    expect(s.values).toMatchObject({ C1: 's:00123', D1: 's:00123' });
  });

  it('the flag is per write: a later plain write to the same cell removes the prefix again; later ops keep it', async () => {
    const out = await replayLog(await buildR3Fixture(), [
      cells('Q', [w(6, 1, { t: 's', v: '00123', quote: true }), w(8, 1, { t: 's', v: '00456', quote: true })]),
      insC('Q', 1),
      cells('Q', [w(6, 2, { t: 's', v: 'x' })]),
    ]);
    const s = await sheetOf(snapshot(out), 'Q');
    expect(s.styles.B6).toBeUndefined(); // column A moved to B by the insert
    expect(s.styles.B8).toMatch(/ quote$/);
    expect(s.values).toMatchObject({ B6: 's:x', B8: 's:00456' });
  });

  it("a constant typed into a table's totals row is the column label; a quote flag keeps the prefix on that cell like Excel, none otherwise", async () => {
    const out = await replayLog(read('r2-table.xlsx'), [cells('T', [w(5, 1, { t: 's', v: '5', quote: true }), w(5, 3, { t: 's', v: '計' })])]);
    const x = (await part(out, 'xl/tables/table1.xml'))!;
    expect(/<tableColumn\b[^>]*name="ID"[^>]*>/.exec(x)![0]).toContain('totalsRowLabel="5"');
    expect(/<tableColumn\b[^>]*name="Twice"[^>]*>/.exec(x)![0]).toContain('totalsRowLabel="計"');
    const s = await sheetOf(snapshot(out), 'T');
    expect(s.values).toMatchObject({ A5: 's:5', C5: 's:計' });
    expect(s.styles.A5).toMatch(/ quote$/);
    expect(s.styles.C5 ?? '').not.toMatch(/ quote$/);
  });
});

// ---------------------------------------------------------------------------------------------------------------
describe('tables: in-table reorder to the column right after the table, names of inserted columns', () => {
  interface Scenario { id: string; cut: string; at: string; error: string | null; table: string | null; headers: string[]; markers: string[] }
  const FX = readFixture<{ scenarios: Scenario[] }>('table-move-excel-results.json');
  const colNo = (s: string) => [...s].reduce((v, ch) => v * 26 + ch.charCodeAt(0) - 64, 0);
  const range = (a1: string) => { const [x, y] = a1.split(':'); return { from: colNo(x.replace(/\d/g, '')), count: colNo(y.replace(/\d/g, '')) - colNo(x.replace(/\d/g, '')) + 1 }; };
  const L = (n: number) => String.fromCharCode(64 + n);

  /** the fixture of tableMove.test.ts: Tbl on C3:H10, headers tC..tH, row-1 markers hA..hN (A..N) */
  async function tableBook(): Promise<Uint8Array> {
    const sst = [...'ABCDEFGHIJKLMN'].map((c) => `h${c}`).concat([...'CDEFGH'].map((c) => `t${c}`));
    const head = [...'ABCDEFGHIJKLMN'].map((c, i) => `<c r="${c}1" t="s"><v>${i}</v></c>`).join('');
    const hdr = [...'CDEFGH'].map((c, i) => `<c r="${c}3" t="s"><v>${14 + i}</v></c>`).join('');
    let data = '';
    for (let r = 4; r <= 10; r++) data += `<row r="${r}">${[...'CDEFGH'].map((c, i) => `<c r="${c}${r}"><v>${r * 10 + i}</v></c>`).join('')}</row>`;
    const table = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Tbl" displayName="Tbl" ref="C3:H10" totalsRowShown="0"><autoFilter ref="C3:H10"/>' +
      `<tableColumns count="6">${[...'CDEFGH'].map((c, i) => `<tableColumn id="${i + 1}" name="t${c}"/>`).join('')}</tableColumns></table>`;
    return buildXlsx([{
      name: 'S',
      data: `<row r="1">${head}</row><row r="3">${hdr}</row>${data}`,
      tail: '<tableParts count="1"><tablePart r:id="rId1"/></tableParts>',
      rels: [{ type: 'table', path: 'xl/tables/table1.xml', xml: table, contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml', id: 'rId1' }],
    }], { sst });
  }

  it('equals Excel (native Insert Cut Cells) for every scenario of the table-move probe, incl. landing right after the table', async () => {
    expect(FX.scenarios.length).toBe(9);
    const bad: string[] = [];
    for (const sc of FX.scenarios) {
      expect(sc.error).toBeNull();
      const c = range(sc.cut), at = range(sc.at);
      const out = await replayLog(await tableBook(), [movC('S', c.from, c.count, at.from)]);
      const x = (await part(out, 'xl/tables/table1.xml'))!;
      const ref = getAttr(/<table\b[^>]*>/.exec(x)![0], 'ref');
      if (ref !== sc.table) bad.push(`${sc.id}: table ${ref}, Excel ${sc.table}`);
      const names = [...x.matchAll(/<tableColumn\b[^>]*\sname="([^"]*)"/g)].map((m) => m[1]);
      const s = await sheetOf(snapshot(out), 'S');
      const tb = sc.table ? sc.table.split(':') : [];
      const c1 = tb.length ? colNo(tb[0].replace(/\d/g, '')) : 0;
      // header cells row 3 = the table column names = Excel's header row
      const headerCells = names.map((_n, i) => (s.values[`${L(c1 + i)}3`] ?? '').replace(/^s:/, ''));
      if (headerCells.join() !== sc.headers.join()) bad.push(`${sc.id}: headers ${headerCells}, Excel ${sc.headers}`);
      if (names.join() !== sc.headers.join()) bad.push(`${sc.id}: tableColumns ${names}, Excel ${sc.headers}`);
      const markers = [...Array(14)].map((_v, i) => (s.values[`${L(i + 1)}1`] ?? '').replace(/^s:/, ''));
      if (markers.join() !== sc.markers.join()) bad.push(`${sc.id}: markers ${markers}, Excel ${sc.markers}`);
    }
    expect(bad).toEqual([]);
  });

  it('the moved block becomes the last table column(s): data, autoFilter and structured references follow', async () => {
    // g-table: tblMain A3:G13 (ID 氏名 大学 点数1 点数2 合計 列1), H hidden. 氏名 (B) -> right after the table (before H = 8)
    const out = await replayLog(read('g-table.xlsx'), [movC('名簿', 2, 1, 8)]);
    const x = (await part(out, 'xl/tables/table1.xml'))!;
    expect(getAttr(/<table\b[^>]*>/.exec(x)![0], 'ref')).toBe('A3:G13');
    expect([...x.matchAll(/<tableColumn\b[^>]*\sname="([^"]*)"/g)].map((m) => m[1])).toEqual(['ID', '大学', '点数1', '点数2', '合計', '列1', '氏名']);
    expect(getAttr(/<autoFilter\b[^>]*>/.exec(x)![0], 'ref')).toBe('A3:G13');
    const before = await sheetOf(snapshot(read('g-table.xlsx')), '名簿');
    const s = await sheetOf(snapshot(out), '名簿');
    expect([...'ABCDEFG'].map((c) => s.values[`${c}3`])).toEqual(['s:ID', 's:大学', 's:点数1', 's:点数2', 's:合計', 's:列1', 's:氏名']);
    for (const r of [4, 7, 13]) expect(s.values[`G${r}`], `G${r}`).toBe(before.values[`B${r}`]); // the moved column's data
    // structured references do not change on a move inside the table; H and the merge I1:K1 did not move
    expect(s.merges).toContain('I1:K1');
    const summary = await sheetOf(snapshot(out), 'Summary');
    const was = await sheetOf(snapshot(read('g-table.xlsx')), 'Summary');
    const structured = Object.keys(was.formulas).filter((k) => was.formulas[k].includes('tblMain['));
    expect(structured.length).toBeGreaterThanOrEqual(3);
    for (const k of structured) expect(summary.formulas[k], k).toBe(was.formulas[k]);
    // plain A1 references follow the move like any range: 点数1 (D) is column C now
    expect(summary.formulas.A1).toBe('SUM(名簿!C4:C13)');
    expect(s.tables[0]).toContain('ref=A3:G13');
  });

  it('one column further (right after H) still leaves the table: refused with a Japanese message; rows keep the old rule', async () => {
    await expect(replayLog(read('g-table.xlsx'), [movC('名簿', 2, 1, 9)])).rejects.toThrow(/テーブル/);
    await expect(replayLog(read('g-table.xlsx'), [movC('名簿', 9, 1, 3)])).rejects.toThrow(JP); // from outside into the table
    await expect(replayLog(read('g-table.xlsx'), [movC('名簿', 6, 3, 1)])).rejects.toThrow(JP); // partly over the table (F..H)
    // rows: a block inside the table may not land right after its last row (only the columns rule changed)
    await expect(replayLog(read('g-table.xlsx'), [struct('名簿', { type: 'move', axis: 'row', from: 5, count: 1, to: 15 })])).rejects.toThrow(JP);
  });

  it("names of columns inserted into a table are Excel 365's: left neighbour 'Q1' -> 'Q2' when free, else 'Q12' ...; header cells hold the names", async () => {
    const names = async (ops: LoggedOp[]) => {
      const out = await replayLog(await buildR3Fixture(), ops);
      const x = (await part(out, 'xl/tables/table1.xml'))!;
      const cols = [...x.matchAll(/<tableColumn\b[^>]*\sname="([^"]*)"/g)].map((m) => m[1]);
      const s = await sheetOf(snapshot(out), 'Tb');
      const ref = getAttr(/<table\b[^>]*>/.exec(x)![0], 'ref');
      const c1 = colNo(/^([A-Z]+)/.exec(ref!)![1]);
      const header = [...cols].map((_n, i) => (s.values[`${L(c1 + i)}1`] ?? '').replace(/^s:/, ''));
      expect(header).toEqual(cols);
      return { cols, ref };
    };
    // Tb: TblQ A1:E4 = ID Q1 Q2 Q3 合計
    expect(await names([insC('Tb', 3)])).toEqual({ cols: ['ID', 'Q1', 'Q12', 'Q2', 'Q3', '合計'], ref: 'A1:F4' }); // Q2 taken
    expect(await names([insC('Tb', 4)])).toEqual({ cols: ['ID', 'Q1', 'Q2', 'Q22', 'Q3', '合計'], ref: 'A1:F4' });
    expect(await names([insC('Tb', 5)])).toEqual({ cols: ['ID', 'Q1', 'Q2', 'Q3', 'Q4', '合計'], ref: 'A1:F4' }); // Q4 free
    expect(await names([insC('Tb', 2, 2)])).toEqual({ cols: ['ID', '列1', '列2', 'Q1', 'Q2', 'Q3', '合計'], ref: 'A1:G4' }); // no number on the left
    expect(await names([insC('Tb', 3, 3)])).toEqual({ cols: ['ID', 'Q1', 'Q12', 'Q13', 'Q14', 'Q2', 'Q3', '合計'], ref: 'A1:H4' });
    // outside the table (right after 合計 / before ID): the table does not grow, no new column
    expect(await names([insC('Tb', 6)])).toEqual({ cols: ['ID', 'Q1', 'Q2', 'Q3', '合計'], ref: 'A1:E4' });
    expect(await names([insC('Tb', 1)])).toEqual({ cols: ['ID', 'Q1', 'Q2', 'Q3', '合計'], ref: 'B1:F4' });
  });

  it('g-table (the editor invariant case): right of 点数1 while 点数2 exists -> 点数12, 点数13', async () => {
    const one = await sheetOf(snapshot(await replayLog(read('g-table.xlsx'), [insC('名簿', 5)])), '名簿');
    expect(one.values.E3).toBe('s:点数12');
    expect(one.tables[0]).toContain('点数1|点数12|点数2');
    const two = await sheetOf(snapshot(await replayLog(read('g-table.xlsx'), [insC('名簿', 5, 2)])), '名簿');
    expect([two.values.E3, two.values.F3]).toEqual(['s:点数12', 's:点数13']);
  });

  it("the rule is the editor's (doc/tables.ts excelNewColumnNames) for every combination", () => {
    const lefts = [null, 'ID', 'Q1', 'Q9', 'Q09', 'h1', '点数1', '2024', 'a1b2', '列1', '列3'];
    const existings = [[], ['ID'], ['ID', 'Q1', 'Q2'], ['ID', 'Q1', 'Q12'], ['q2', 'Q12', 'Q13', '点数2', '列1', '列2', '2025', 'a1b3', 'Q10']];
    let n = 0;
    for (const left of lefts) for (const ex of existings) for (const count of [1, 2, 3, 5]) {
      expect(excelNewColumnNames(ex, left, count), `${left} ${ex} ${count}`).toEqual(docNewColumnNames(ex, left, count));
      n++;
    }
    expect(n).toBe(lefts.length * existings.length * 4);
  });

  it('premerge still names the imported columns by their headers (the inserted 列N names never leak into the file)', async () => {
    const r = await buildPremerged(scenario('table'));
    const x = (await part(r.bytes, 'xl/tables/table1.xml'))!;
    expect([...x.matchAll(/<tableColumn\b[^>]*\sname="([^"]*)"/g)].map((m) => m[1]).every((n) => !/^列\d+$/.test(n))).toBe(true);
  });
});

// ---------------------------------------------------------------------------------------------------------------
describe('totals row (round 3: the doc converts constants to text itself)', () => {
  const col = (xml: string, name: string) => new RegExp(`<tableColumn\\b[^>]*name="${name}"[^>]*?(?:/>|>[\\s\\S]*?</tableColumn>)`).exec(xml)![0];

  it("text that looks like a number / date / TRUE is the column's label (t='s'), the function / formula of the column is dropped", async () => {
    const out = await replayLog(read('r2-table.xlsx'), [cells('T', [w(5, 1, { t: 's', v: '5' }), w(5, 2, { t: 's', v: '2026/10/20' }), w(5, 3, { t: 's', v: 'TRUE' })])]);
    const x = (await part(out, 'xl/tables/table1.xml'))!;
    expect(col(x, 'ID')).toContain('totalsRowLabel="5"');
    expect(col(x, 'Score')).toContain('totalsRowLabel="2026/10/20"');
    expect(col(x, 'Score')).not.toMatch(/totalsRowFunction|totalsRowFormula/);
    expect(col(x, 'Twice')).toContain('totalsRowLabel="TRUE"');
    expect(col(x, 'Twice')).not.toContain('totalsRowFunction');
    const t = await sheetOf(snapshot(out), 'T');
    expect(t.values).toMatchObject({ A5: 's:5', B5: 's:2026/10/20', C5: 's:TRUE' });
    expect(t.formulas.B5).toBeUndefined();
    // the cell format is not touched (General stays General), no number is stored
    expect(t.styles.B5 ?? '').not.toMatch(/nf=#14/);
  });

  it('a text label written after a function, and the formula forms, in one log', async () => {
    const f = (c: number, text: string) => w(5, c, { t: 'z', v: null, f: text });
    const out = await replayLog(read('r2-table.xlsx'), [
      cells('T', [f(2, 'SUBTOTAL(109,Tbl[Score])'), f(3, 'SUM(Tbl[Twice])')]),
      cells('T', [w(5, 2, { t: 's', v: '合計' })]),
    ]);
    const x = (await part(out, 'xl/tables/table1.xml'))!;
    expect(col(x, 'Score')).toContain('totalsRowLabel="合計"');
    expect(col(x, 'Score')).not.toContain('totalsRowFunction');
    expect(col(x, 'Twice')).toContain('totalsRowFunction="custom"');
    expect(col(x, 'Twice')).toContain('<totalsRowFormula>SUM(Tbl[Twice])</totalsRowFormula>');
  });
});

// ---------------------------------------------------------------------------------------------------------------
describe('CF blocks: the other rules of the block take part in the move decisions (siblings)', () => {
  const block = (rules: string[], sqref = 'B2:B6') => `<conditionalFormatting sqref="${sqref}">${rules.map((f, i) => `<cfRule type="expression" priority="${i + 1}"><formula>${f.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')}</formula></cfRule>`).join('')}</conditionalFormatting>`;
  const op: StructOp = { type: 'move', axis: 'col', from: 4, count: 1, to: 3 }; // D before C: lands right after B2:B6
  const sqrefs = (xml: string) => [...xml.matchAll(/<conditionalFormatting sqref="([^"]*)">/g)].map((m) => m[1]);
  const formulas = (xml: string) => [...xml.matchAll(/<formula>([^<]*)<\/formula>/g)].map((m) => m[1]);

  it('alone, a rule spreads onto the block that lands right after its area', () => {
    const r = adjustUnit(block(['B2="x"']), 'cf', op, 'S', { leak: true });
    expect(sqrefs(r.xml)).toEqual(['B2:C6']);
  });

  it('a sibling that refers to the moved column keeps EVERY rule of the block from spreading; one that does not, changes nothing', () => {
    const cut = adjustUnit(block(['B2="x"', '$D2>5']), 'cf', op, 'S', { leak: true });
    expect(sqrefs(cut.xml)).toEqual(['B2:B6']);
    expect(formulas(cut.xml)).toEqual(['B2="x"', '$C2&gt;5']);
    const free = adjustUnit(block(['B2="x"', '$F2>5']), 'cf', op, 'S', { leak: true });
    expect(sqrefs(free.xml)).toEqual(['B2:C6']);
    // a clean insert never spreads (premerge): unchanged
    expect(adjustUnit(block(['B2="x"', '$D2>5']), 'cf', op, 'S', { leak: false }).xml).toContain('sqref="B2:B6"');
  });

  it('replayLog runs the same decision on a package (classic <conditionalFormatting> of a real sheet)', async () => {
    const book = await buildXlsx([{
      name: 'S',
      data: '<row r="1"><c r="A1"><v>1</v></c><c r="D1"><v>1</v></c></row><row r="2"><c r="B2" t="s"><v>0</v></c><c r="D2"><v>9</v></c></row>',
      tail: block(['B2="x"', '$D2>5']),
    }], { sst: ['x'] });
    const x = (await part(await replayLog(book, [movC('S', 4, 1, 3)]), 'xl/worksheets/sheet1.xml'))!;
    expect(sqrefs(x)).toEqual(['B2:B6']);
    const alone = await buildXlsx([{ name: 'S', data: '<row r="1"><c r="A1"><v>1</v></c></row>', tail: block(['B2="x"']) }]);
    expect(sqrefs((await part(await replayLog(alone, [movC('S', 4, 1, 3)]), 'xl/worksheets/sheet1.xml'))!)).toEqual(['B2:C6']);
  });
});

// ---------------------------------------------------------------------------------------------------------------
describe('package zip safety (reader R3: loadZip / readEntryBytes)', () => {
  const CDIR = [0x50, 0x4b, 0x01, 0x02];
  /** patch the central directory: `usize` (+24) and / or `csize` (+20) of every record */
  function patchDirectory(bytes: Uint8Array, fields: { usize?: number; csize?: number }): Uint8Array {
    const out = new Uint8Array(bytes);
    const dv = new DataView(out.buffer);
    for (let i = 0; i + 46 < out.length; i++) {
      if (CDIR.every((b, k) => out[i + k] === b)) {
        if (fields.csize !== undefined) dv.setUint32(i + 20, fields.csize, true);
        if (fields.usize !== undefined) dv.setUint32(i + 24, fields.usize, true);
      }
    }
    return out;
  }
  const smallBook = () => buildXlsx([{ name: 'S', data: '<row r="1"><c r="A1"><v>1</v></c></row>' }]);

  it('refuses a package whose central directory declares more than 1.5 GB in total (before inflating anything)', async () => {
    const bomb = patchDirectory(await smallBook(), { usize: 0x7fffffff });
    const e = await replayLog(bomb, []).then(() => null, (x: unknown) => x);
    expect(e).toBeInstanceOf(PatchError);
    expect((e as Error).message).toMatch(/大きすぎる/);
    await expect(buildPremerged({ ...scenario('plain'), base: { ...scenario('plain').base, book: { ...scenario('plain').base.book, bytes: bomb } } })).rejects.toThrow(PatchError);
  });

  it('an entry that inflates beyond the size it declares is a PatchError, not an out-of-memory', async () => {
    // declared 10 bytes, real deflate stream of the whole part
    const lying = patchDirectory(await smallBook(), { usize: 10 });
    const e = await replayLog(lying, []).then(() => null, (x: unknown) => x);
    expect(e).toBeInstanceOf(PatchError);
    expect((e as Error).message).toMatch(JP);
  });

  it('not a zip / a zip without [Content_Types].xml: Japanese PatchError', async () => {
    await expect(replayLog(new Uint8Array([1, 2, 3, 4, 5]), [])).rejects.toThrow(PatchError);
    const z = new JSZip();
    z.file('hello.txt', 'hi');
    await expect(replayLog(await z.generateAsync({ type: 'uint8array' }), [])).rejects.toThrow(JP);
  });

  it('normal packages are unaffected (same bytes out for an empty log as the plain round trip keeps every part)', async () => {
    const out = await replayLog(read('torture.xlsx'), []);
    const a = await JSZip.loadAsync(read('torture.xlsx'));
    const b = await JSZip.loadAsync(out);
    // every part is still there (directory entries are JSZip's, the calcChain is dropped on purpose)
    const parts = (z: JSZip) => Object.keys(z.files).filter((n) => !z.files[n].dir && !/calcChain/.test(n)).sort();
    expect(parts(b)).toEqual(parts(a));
  });
});
