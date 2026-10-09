/**
 * replayLog unit tests (no Excel needed): cell writes, widths, op ordering, package hygiene (calcChain,
 * untouched parts), refusals with Japanese messages, robustness, performance.
 * Excel-exactness of the structural ops is covered by tests/golden/golden.test.ts.
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { planMove, type StructOp } from '../../src/core/formula/excelRef';
import type { LoggedOp } from '../../src/core/types';
import { exportWithJs } from '../../src/core/export/js';
import { replayLog } from '../../src/core/xlsx/replay';
import { buildXlsx, listParts, part, partBytes } from '../golden/build';
import { snapshot } from '../golden/snapshot';

const here = path.dirname(fileURLToPath(import.meta.url));
const GF = path.resolve(here, '../fixtures/golden');
const XF = path.resolve(here, '../fixtures/xlsx');
const read = (p: string) => new Uint8Array(fs.readFileSync(p));
const JP = /[ぁ-んァ-ン一-龥]/;
const st = (sheet: string, op: StructOp): LoggedOp => ({ kind: 'struct', sheet, op });

async function small(): Promise<Uint8Array> {
  return buildXlsx([
    {
      name: 'S',
      head: '<cols><col min="2" max="2" width="20.625" customWidth="1"/></cols>',
      data: '<row r="1"><c r="A1" s="1" t="s"><v>0</v></c><c r="B1" s="1" t="s"><v>1</v></c><c r="C1" s="1"/></row>' +
        '<row r="2"><c r="A2" s="2"><v>1.5</v></c><c r="B2"><f>A2*2</f><v>3</v></c><c r="C2" s="2"/></row>' +
        '<row r="3"><c r="A3"><v>2</v></c><c r="B3"><f t="shared" ref="B3:B5" si="0">A3*2</f><v>4</v></c></row>' +
        '<row r="4"><c r="A4"><v>3</v></c><c r="B4"><f t="shared" si="0"/><v>6</v></c></row>' +
        '<row r="5"><c r="A5"><v>4</v></c><c r="B5"><f t="shared" si="0"/><v>8</v></c></row>',
    },
    { name: 'T', data: '<row r="1"><c r="A1"><f>SUM(S!A2:A5)</f><v>10.5</v></c><c r="B1" t="s"><v>2</v></c></row>' },
    { name: 'U', data: '<row r="1"><c r="A1"><v>1</v></c></row>' },
  ], { sst: ['見出し', '名前', 'x'], calcChain: true, names: '<definedName name="Rng">S!$A$2:$A$5</definedName>' });
}

describe('cell writes', () => {
  it('writes every value type, keeps the cell style, clears with null', async () => {
    const log: LoggedOp[] = [{
      kind: 'cells', sheet: 'S', reason: 'edit', writes: [
        { row: 1, col: 1, cell: { t: 's', v: '新しい見出し' } },
        { row: 2, col: 1, cell: { t: 'n', v: 42 } },
        { row: 2, col: 3, cell: { t: 'b', v: true } },
        { row: 6, col: 1, cell: { t: 'e', v: '#N/A' } },
        { row: 6, col: 2, cell: { t: 'n', v: null, f: 'SUM(A2:A5)' } },
        { row: 6, col: 3, cell: { t: 's', v: 'a<b&"c"', f: '"a<b&""c"""' } },
        { row: 7, col: 1, cell: { t: 'd', v: 46000 } },
        { row: 1, col: 2, cell: null },
      ],
    }];
    const out = await replayLog(await small(), log);
    const s = (await snapshot(out)).sheets[0];
    expect(s.values.A1).toBe('s:新しい見出し');
    expect(s.values.A2).toBe('n:42');
    expect(s.values.C2).toBe('b:1');
    expect(s.values.A6).toBe('e:#N/A');
    expect(s.formulas.B6).toBe('SUM(A2:A5)');
    expect(s.formulas.C6).toBe('"a<b&""c"""');
    expect(s.values.A7).toBe('n:46000');
    expect(s.values.B1).toBeUndefined();
    const xml = (await part(out, 'xl/worksheets/sheet1.xml'))!;
    expect(xml).toMatch(/<c r="A1" s="1" t="s"><v>\d+<\/v><\/c>/); // style kept, shared string
    expect(xml).toMatch(/<c r="B1" s="1"\/>/); // cleared, style kept
    expect(xml).toMatch(/<c r="A2" s="2"><v>42<\/v><\/c>/);
    expect(xml).toMatch(/<c r="C6" t="str"><f>"a&lt;b&amp;""c"""<\/f><v>a&lt;b&amp;"c"<\/v><\/c>/);
    const sst = (await part(out, 'xl/sharedStrings.xml'))!;
    expect(sst).toContain('<t>新しい見出し</t>');
    expect(sst).toMatch(/uniqueCount="4"/);
  });

  it('a written formula only receives the ops logged after the write', async () => {
    const before: LoggedOp[] = [
      { kind: 'cells', sheet: 'S', reason: 'edit', writes: [{ row: 8, col: 1, cell: { t: 'z', v: null, f: 'B2+S!B3' } }] },
      st('S', { type: 'insert', axis: 'col', at: 1, count: 1 }),
    ];
    const after: LoggedOp[] = [before[1], { kind: 'cells', sheet: 'S', reason: 'edit', writes: [{ row: 8, col: 2, cell: { t: 'z', v: null, f: 'B2+S!B3' } }] }];
    expect((await snapshot(await replayLog(await small(), before))).sheets[0].formulas.B8).toBe('C2+S!C3');
    expect((await snapshot(await replayLog(await small(), after))).sheets[0].formulas.B8).toBe('B2+S!B3');
  });

  it('column widths are written as <col width customWidth>', async () => {
    const out = await replayLog(await small(), [{ kind: 'colWidth', sheet: 'S', col: 4, width: 12.5 }, { kind: 'colWidth', sheet: 'S', col: 2, width: 9 }]);
    const s = (await snapshot(out)).sheets[0];
    expect(s.cols.D).toBe('w=12.5');
    expect(s.cols.B).toBe('w=9');
  });

  it('toValues writes replace formulas by values', async () => {
    const out = await replayLog(await small(), [{ kind: 'cells', sheet: 'S', reason: 'toValues', writes: [{ row: 2, col: 2, cell: { t: 'n', v: 3 } }] }]);
    const s = (await snapshot(out)).sheets[0];
    expect(s.formulas.B2).toBeUndefined();
    expect(s.values.B2).toBe('n:3');
  });
});

describe('structural ops on a small book', () => {
  it('moves cells with their styles, unshares shared formulas, rewrites other sheets and names', async () => {
    const out = await replayLog(await small(), [st('S', { type: 'insert', axis: 'col', at: 1, count: 2 })]);
    const snap = await snapshot(out);
    const s = snap.sheets[0];
    expect(s.formulas).toMatchObject({ D2: 'C2*2', D3: 'C3*2', D4: 'C4*2', D5: 'C5*2' });
    expect(s.values.C1).toBe('s:見出し');
    expect(s.cols.D).toBe('w=20.625');
    expect(snap.sheets[1].formulas.A1).toBe('SUM(S!C2:C5)');
    expect(snap.names).toEqual(['Rng = S!$C$2:$C$5']);
    const xml = (await part(out, 'xl/worksheets/sheet1.xml'))!;
    expect(xml).not.toContain('t="shared"');
    expect(xml).toContain('<dimension ref="C1:E5"/>');
  });

  it('removes calcChain (part, relationship, content type) and sets fullCalcOnLoad', async () => {
    const out = await replayLog(await small(), []);
    expect(await listParts(out)).not.toContain('xl/calcChain.xml');
    expect(await part(out, 'xl/_rels/workbook.xml.rels')).not.toContain('calcChain');
    expect(await part(out, '[Content_Types].xml')).not.toContain('calcChain');
    expect(await part(out, 'xl/workbook.xml')).toMatch(/<calcPr calcId="191029" fullCalcOnLoad="1"\/>/);
  });

  it('keeps every byte of the parts it does not need to touch', async () => {
    const input = read(path.join(GF, 'torture.xlsx'));
    const out = await replayLog(input, [st('Data', { type: 'insert', axis: 'col', at: 3, count: 1 })]);
    for (const p of ['xl/theme/theme1.xml', 'xl/media/image1.png', 'xl/worksheets/sheet5.xml', 'xl/styles.xml', 'xl/metadata.xml', 'xl/persons/person.xml', 'docProps/core.xml']) {
      expect(Buffer.from((await partBytes(out, p))!).equals(Buffer.from((await partBytes(input, p))!)), p).toBe(true);
    }
    // sheet 'Tbl' has no reference to Data: untouched as well
    expect(await part(out, 'xl/worksheets/sheet3.xml')).toBe(await part(input, 'xl/worksheets/sheet3.xml'));
  });

  it('removes a table whose columns are all deleted (part, relationship, content type, tablePart)', async () => {
    const input = read(path.join(GF, 'g-table.xlsx'));
    const out = await replayLog(input, [st('名簿', { type: 'delete', axis: 'col', at: 1, count: 7 })]);
    expect(await listParts(out)).not.toContain('xl/tables/table1.xml');
    expect(await part(out, 'xl/worksheets/_rels/sheet1.xml.rels')).not.toContain('table1.xml');
    expect(await part(out, '[Content_Types].xml')).not.toContain('table1.xml');
    expect(await part(out, 'xl/worksheets/sheet1.xml')).not.toContain('<tablePart');
    const sum = (await snapshot(out)).sheets.find((s) => s.name === 'Summary')!;
    expect(sum.formulas.A2).toBe('AVERAGE(#REF!)');
    expect(sum.formulas.A11).toBe('ROWS(#REF!)');
  });

  it('deletes notes and threaded comments together with their cells', async () => {
    const input = read(path.join(GF, 'torture.xlsx'));
    const out = await replayLog(input, [st('Data', { type: 'delete', axis: 'row', at: 4, count: 2 })]);
    const s = (await snapshot(out)).sheets[0];
    expect(s.comments).toEqual([]);
    expect(s.threaded).toEqual([]);
    expect(await part(out, 'xl/drawings/vmlDrawing1.vml')).not.toMatch(/<x:Row>/);
  });

  it('planMove sequences for non-contiguous selections', async () => {
    const input = await small();
    const ops = planMove([1, 3], 3).map((op) => st('S', op));
    const s = (await snapshot(await replayLog(input, ops))).sheets[0];
    // A and C now sit before (old) C: order B, A, C  -> A moves to B
    expect(s.values.B1).toBe('s:見出し');
    expect(s.values.A1).toBe('s:名前');
  });
});

describe('refusals and robustness (Japanese messages)', () => {
  it('refuses moving part of a merged area', async () => {
    const input = read(path.join(GF, 'torture.xlsx'));
    await expect(replayLog(input, [st('Data', { type: 'move', axis: 'col', from: 3, count: 1, to: 8 })])).rejects.toThrow(/結合/);
  });
  it('refuses moves into / out of a table and deleting a table header row', async () => {
    const input = read(path.join(GF, 'g-table.xlsx'));
    await expect(replayLog(input, [st('名簿', { type: 'move', axis: 'col', from: 12, count: 1, to: 3 })])).rejects.toThrow(/テーブル/);
    await expect(replayLog(input, [st('名簿', { type: 'move', axis: 'col', from: 2, count: 1, to: 9 })])).rejects.toThrow(/テーブル/);
    await expect(replayLog(input, [st('名簿', { type: 'delete', axis: 'row', at: 3, count: 1 })])).rejects.toThrow(/見出し行/);
  });
  it('refuses changing part of a legacy array formula', async () => {
    const input = await buildXlsx([{ name: 'A', data: '<row r="2"><c r="B2"><f t="array" ref="B2:C3">A1:B2*2</f><v>1</v></c><c r="C2"><v>1</v></c></row><row r="3"><c r="B3"><v>1</v></c><c r="C3"><v>1</v></c></row>' }]);
    await expect(replayLog(input, [st('A', { type: 'insert', axis: 'col', at: 3, count: 1 })])).rejects.toThrow(/配列/);
    await expect(replayLog(input, [st('A', { type: 'delete', axis: 'row', at: 3, count: 1 })])).rejects.toThrow(/配列/);
    await expect(replayLog(input, [{ kind: 'cells', sheet: 'A', reason: 'edit', writes: [{ row: 3, col: 3, cell: { t: 'n', v: 1 } }] }])).rejects.toThrow(/配列/);
    // after an insert and a write elsewhere, the array is still known at its new place
    await expect(replayLog(input, [
      st('A', { type: 'insert', axis: 'row', at: 1, count: 2 }),
      { kind: 'cells', sheet: 'A', reason: 'edit', writes: [{ row: 1, col: 1, cell: { t: 'n', v: 1 } }] },
      st('A', { type: 'delete', axis: 'row', at: 5, count: 1 }),
    ])).rejects.toThrow(/配列/);
    // the whole array may move / be deleted
    const out = await replayLog(input, [st('A', { type: 'insert', axis: 'col', at: 1, count: 1 })]);
    expect((await snapshot(out)).sheets[0].formulas.C2).toBe('B1:C2*2 {array C2:D3}');
  });
  it('rejects broken input with a Japanese message', async () => {
    await expect(replayLog(new Uint8Array([1, 2, 3]), [])).rejects.toThrow(JP);
    await expect(replayLog(await small(), [st('ない', { type: 'insert', axis: 'col', at: 1, count: 1 })])).rejects.toThrow(/シート「ない」/);
    await expect(replayLog(await small(), [st('S', { type: 'insert', axis: 'col', at: 0, count: 1 })])).rejects.toThrow(JP);
    await expect(replayLog(await small(), [st('S', { type: 'move', axis: 'col', from: 1, count: 0, to: 5 })])).rejects.toThrow(JP);
    await expect(replayLog(await small(), [{ kind: 'colWidth', sheet: 'S', col: 1, width: -3 }])).rejects.toThrow(JP);
    await expect(replayLog(await small(), [st('S', { type: 'insert', axis: 'col', at: 16384, count: 2 })])).resolves.toBeInstanceOf(Uint8Array);
  });
  it('refuses to push content off the sheet', async () => {
    const input = await buildXlsx([{ name: 'E', data: '<row r="1"><c r="XFD1"><v>1</v></c></row>' }]);
    await expect(replayLog(input, [st('E', { type: 'insert', axis: 'col', at: 1, count: 1 })])).rejects.toThrow(/押し出/);
  });
  it('exportWithJs is replayLog', async () => {
    const a = await exportWithJs(await small(), [st('S', { type: 'delete', axis: 'col', at: 1, count: 1 })]);
    expect((await snapshot(a)).sheets[0].formulas.A2).toBe('#REF!*2');
  });
  it('reports progress', async () => {
    const seen: string[] = [];
    await replayLog(await small(), [st('S', { type: 'insert', axis: 'row', at: 1, count: 1 })], (p) => seen.push(p.message));
    expect(seen.length).toBeGreaterThan(2);
    expect(seen.every((m) => JP.test(m))).toBe(true);
  });
});

describe('performance', () => {
  it('replays 20 ops on a 10,000 x 40 sheet (50,000 formulas) in < 10 s', async () => {
    const input = read(path.join(XF, 'perf_10000x40.xlsx'));
    const ops: LoggedOp[] = [];
    for (let i = 0; i < 6; i++) ops.push(st('Big', { type: 'move', axis: 'col', from: 3 + i, count: 2, to: 30 - i }));
    for (let i = 0; i < 4; i++) ops.push(st('Big', { type: 'insert', axis: 'col', at: 5 + i * 3, count: 1 }));
    for (let i = 0; i < 3; i++) ops.push(st('Big', { type: 'delete', axis: 'col', at: 10 + i, count: 1 }));
    ops.push(st('Big', { type: 'move', axis: 'col', from: 20, count: 3, to: 2 }));
    ops.push(st('Big', { type: 'insert', axis: 'row', at: 500, count: 3 }));
    ops.push(st('Big', { type: 'delete', axis: 'row', at: 900, count: 2 }));
    ops.push({ kind: 'colWidth', sheet: 'Big', col: 4, width: 22 });
    ops.push({ kind: 'cells', sheet: 'Big', reason: 'edit', writes: [{ row: 2, col: 3, cell: { t: 'n', v: 5, f: 'A2+B2' } }] });
    ops.push(st('Big', { type: 'move', axis: 'col', from: 1, count: 1, to: 40 }));
    ops.push(st('Big', { type: 'insert', axis: 'col', at: 2, count: 2 }));
    expect(ops.length).toBe(20);
    const t0 = Date.now();
    const out = await replayLog(input, ops);
    const ms = Date.now() - t0;
    expect(out.length).toBeGreaterThan(100000);
    expect(ms).toBeLessThan(10000);
  }, 60000);
});
