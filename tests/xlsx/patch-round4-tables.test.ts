/**
 * Deleting the totals row of a table (found by the random op-log fuzz with Excel as the oracle: the file Excel was asked to
 * open stopped Excel itself). Excel's own result is in the goldens h2-del-totals-row / h2-del-last-data-and-totals:
 * ref shrinks, totalsRowCount goes (totalsRowShown="0"), the filter covers the remaining rows, the columns lose their
 * totals function / label / formula. Fixture r2-table.xlsx: Tbl A1:C5 (ID Score Twice; 3 data rows; totals row 5).
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import type { CellWrite, LoggedOp } from '../../src/core/types';
import { replayLog } from '../../src/core/xlsx/replay';
import { PatchError } from '../../src/core/xlsx/patch/pkg';
import { packageProblems, text, unpack } from '../golden/package-tools';

const here = path.dirname(fileURLToPath(import.meta.url));
const fixture = () => new Uint8Array(fs.readFileSync(path.join(here, '../fixtures/golden/r2-table.xlsx')));
const delR = (at: number, count = 1): LoggedOp => ({ kind: 'struct', sheet: 'T', op: { type: 'delete', axis: 'row', at, count } });
const cells = (sheet: string, writes: CellWrite[]): LoggedOp => ({ kind: 'cells', sheet, writes, reason: 'edit' });
const w = (row: number, col: number, cell: CellWrite['cell']): CellWrite => ({ row, col, cell });

async function table(bytes: Uint8Array): Promise<string> {
  return text(await unpack(bytes), 'xl/tables/table1.xml');
}
const open = (xml: string) => /<table\b[^>]*>/.exec(xml)![0];

describe('deleting the totals row of a table', () => {
  it('the table loses its totals row: ref, totalsRowCount / totalsRowShown, filter range, column definitions', async () => {
    const out = await replayLog(fixture(), [delR(5)]);
    expect(await packageProblems(out)).toEqual([]);
    const t = await table(out);
    expect(open(t)).toContain('ref="A1:C4"');
    expect(open(t)).not.toContain('totalsRowCount');
    expect(open(t)).toContain('totalsRowShown="0"');
    expect(t).toMatch(/<autoFilter ref="A1:C4"/);
    expect(t).not.toMatch(/totalsRow(Function|Label|Formula|DxfId)/);
    // what is not about the totals stays: the calculated column, the ids and names
    expect(t).toContain('<calculatedColumnFormula>Tbl[[#This Row],[Score]]*2</calculatedColumnFormula>');
    expect(t).toMatch(/<tableColumns count="3"><tableColumn id="1" [^>]*name="ID"\/><tableColumn id="2" [^>]*name="Score"\/>/);
  });

  it('the last data row and the totals row together; a data row alone keeps the totals', async () => {
    const both = await table(await replayLog(fixture(), [delR(4, 2)]));
    expect(open(both)).toContain('ref="A1:C3"');
    expect(open(both)).not.toContain('totalsRowCount');
    expect(both).toMatch(/<autoFilter ref="A1:C3"/);
    const dataOnly = await table(await replayLog(fixture(), [delR(3)]));
    expect(open(dataOnly)).toContain('ref="A1:C4"');
    expect(open(dataOnly)).toContain('totalsRowCount="1"');
    expect(dataOnly).toMatch(/<autoFilter ref="A1:C3"/);
    expect(dataOnly).toContain('totalsRowFunction="custom"');
  });

  it('afterwards the former totals row is an ordinary data row: a typed text is no totals label, later ops see no totals', async () => {
    const out = await replayLog(fixture(), [delR(5), cells('T', [w(4, 1, { t: 's', v: 'x' }), w(4, 2, { t: 'n', v: 5 })])]);
    const t = await table(out);
    expect(t).not.toMatch(/totalsRow(Count|Function|Label|Formula)/);
    // and a row inserted at the end of the data now grows the table like any data row
    const more = await replayLog(fixture(), [delR(5), { kind: 'struct', sheet: 'T', op: { type: 'insert', axis: 'row', at: 4, count: 1 } }]);
    const t2 = await table(more);
    expect(open(t2)).toContain('ref="A1:C5"');
    expect(t2).not.toMatch(/totalsRow(Function|Label|Formula)/);
    expect(await packageProblems(more)).toEqual([]);
  });

  it('deleting the whole table (header to totals) still removes the table part; a table without a totals row is unaffected', async () => {
    const all = await replayLog(fixture(), [delR(1, 5)]);
    expect((await unpack(all)).has('xl/tables/table1.xml')).toBe(false);
    expect(await packageProblems(all)).toEqual([]);
    // the same op on a table with no totals row (g-table.xlsx: tblMain A3:G13, the 合計 row 15 lies below it)
    const g = new Uint8Array(fs.readFileSync(path.join(here, '../fixtures/golden/g-table.xlsx')));
    const out = await replayLog(g, [{ kind: 'struct', sheet: '名簿', op: { type: 'delete', axis: 'row', at: 13, count: 1 } }]);
    const t = text(await unpack(out), 'xl/tables/table1.xml');
    expect(open(t)).toMatch(/ref="A3:G12"/);
    expect(t).not.toContain('totalsRowShown="0" totalsRowShown');
  });
});

describe('moving rows inside a table: only data rows, only between the header row and the end', () => {
  const mov = (from: number, to: number, count = 1): LoggedOp => ({ kind: 'struct', sheet: 'T', op: { type: 'move', axis: 'row', from, count, to } });
  const refused = async (log: LoggedOp[]) => {
    let err: unknown;
    try { await replayLog(fixture(), log); } catch (e) { err = e; }
    expect(err).toBeInstanceOf(PatchError);
    expect((err as PatchError).internal).toBe(false);
  };

  it('a data row dropped at the header row, or the header / totals row moved, is refused (Japanese PatchError)', async () => {
    await refused([mov(3, 1)]); // before the header row: the table would get a data row as its header
    await refused([mov(1, 3)]); // the header row itself
    await refused([mov(5, 3)]); // the totals row
    await refused([mov(3, 6)]); // below the totals row = out of the table
  });

  it('a data row between the header and the totals row is an ordinary reorder: the table keeps its range and its totals', async () => {
    const out = await replayLog(fixture(), [mov(4, 2)]); // the last data row becomes the first
    expect(await packageProblems(out)).toEqual([]);
    const t = await table(out);
    expect(open(t)).toContain('ref="A1:C5"');
    expect(open(t)).toContain('totalsRowCount="1"');
    const before = await replayLog(fixture(), [mov(2, 5)]); // the first data row lands right before the totals row
    expect(open(await table(before))).toContain('ref="A1:C5"');
  });
});
