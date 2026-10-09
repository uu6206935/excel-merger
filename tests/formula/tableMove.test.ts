// Whole-column moves around a ListObject against Excel 365's native Insert Cut Cells:
//   fixtures/formula/table-move-excel-results.json (generator com/excel-table-move.ps1: table Tbl on C3:H10, headers
//   tC..tH, row-1 markers hA..) and the refusals of the R2b table suite (r2b-fixtures.json, ops t01..t22).
// The question of round 2: a block from INSIDE the table landing right after its last column (to = c2 + 1). Excel
// accepts it and reorders the table (the block becomes the last columns, the range stays) — tableMoveKind says so, and
// adjustTableRef keeps the range.
import { describe, expect, it } from 'vitest';
import { adjustTableRef, mapIndex, movePerm, tableMoveKind, numToCol, type StructOp } from '../../src/core/formula/excelRef';
import { readFixture } from './helpers';

interface Scenario { id: string; cut: string; at: string; error: string | null; table: string | null; headers: string[]; markers: string[] }
const FX = readFixture<{ scenarios: Scenario[] }>('table-move-excel-results.json');
const colRange = (a1: string) => { const [x, y] = a1.split(':'); const n = (s: string) => [...s].reduce((v, ch) => v * 26 + ch.charCodeAt(0) - 64, 0); return { from: n(x), count: n(y) - n(x) + 1 }; };
const opOf = (sc: Scenario): Extract<StructOp, { type: 'move' }> => {
  const c = colRange(sc.cut), at = colRange(sc.at);
  return { type: 'move', axis: 'col', from: c.from, count: c.count, to: at.from };
};
const HEADERS = ['tC', 'tD', 'tE', 'tF', 'tG', 'tH'];
const KIND: Record<string, string> = {
  in_to_I: 'reorder', in2_to_I: 'reorder', first_to_I: 'reorder', in_to_H: 'reorder', in_to_C: 'reorder',
  whole_to_K: 'whole', cover_to_L: 'whole', out_to_I: 'shift', outL_to_C: 'shift',
};

describe('moves around a table vs Excel (native Insert Cut Cells)', () => {
  it('every scenario: table range, header order and column permutation equal Excel', () => {
    expect(FX.scenarios.length).toBe(9);
    const bad: string[] = [];
    for (const sc of FX.scenarios) {
      expect(sc.error).toBeNull(); // none of these is refused
      const op = opOf(sc);
      const kind = tableMoveKind(3, 8, op);
      if (kind !== KIND[sc.id]) bad.push(`${sc.id}: kind ${kind}, expected ${KIND[sc.id]}`);
      const range = adjustTableRef('C3:H10', op);
      if (range !== sc.table) bad.push(`${sc.id}: table ${range}, Excel ${sc.table}`);
      // header order: reorder inside the table permutes the columns, otherwise the table only shifts
      let names = [...HEADERS];
      if (kind === 'reorder') {
        const p = movePerm(op.from, op.count, op.to);
        names = new Array(6);
        HEADERS.forEach((h, i) => { names[p(3 + i) - 3] = h; });
      }
      if (names.join() !== sc.headers.join()) bad.push(`${sc.id}: headers ${names}, Excel ${sc.headers}`);
      // the row-1 markers show the column permutation of the op (mapIndex is what cells, widths and rules follow)
      const after: string[] = [];
      for (let i = 1; i <= 14; i++) { const j = mapIndex(i, op); if (j != null && j <= 14) after[j - 1] = 'h' + numToCol(i); }
      if (after.join() !== sc.markers.join()) bad.push(`${sc.id}: markers ${after}, Excel ${sc.markers}`);
    }
    expect(bad).toEqual([]);
  });

  it('a block from inside the table landing right after its last column is a reorder (to = c2 + 1)', () => {
    const op: StructOp = { type: 'move', axis: 'col', from: 5, count: 1, to: 9 }; // E before I, table C:H
    expect(tableMoveKind(3, 8, op)).toBe('reorder');
    expect(adjustTableRef('C3:H10', op)).toBe('C3:H10');
    expect(adjustTableRef('Data!$C$3:$H$10', op)).toBe('Data!$C$3:$H$10');
    // one column further: Excel refuses (R2b t03: E before K)
    expect(tableMoveKind(3, 8, { type: 'move', axis: 'col', from: 5, count: 1, to: 10 })).toBe('refused');
  });

  it('tableMoveKind classification', () => {
    const mv = (from: number, count: number, to: number): Extract<StructOp, { type: 'move' }> => ({ type: 'move', axis: 'col', from, count, to });
    // inside: destination within [lo, hi + 1]
    for (const to of [3, 4, 6, 8, 9]) expect(tableMoveKind(3, 8, mv(5, 1, to))).toBe(to >= 5 && to <= 6 ? 'shift' : 'reorder');
    expect(tableMoveKind(3, 8, mv(5, 1, 2))).toBe('refused');
    expect(tableMoveKind(3, 8, mv(5, 1, 10))).toBe('refused');
    // covering block: the table travels with it
    expect(tableMoveKind(3, 8, mv(3, 6, 11))).toBe('whole');
    expect(tableMoveKind(3, 8, mv(2, 8, 12))).toBe('whole');
    expect(tableMoveKind(3, 8, mv(3, 7, 1))).toBe('whole');
    // outside: anywhere but strictly inside
    expect(tableMoveKind(3, 8, mv(10, 1, 9))).toBe('shift');
    expect(tableMoveKind(3, 8, mv(1, 1, 3))).toBe('shift');
    expect(tableMoveKind(3, 8, mv(1, 1, 12))).toBe('shift');
    expect(tableMoveKind(3, 8, mv(1, 1, 5))).toBe('refused');
    expect(tableMoveKind(3, 8, mv(10, 1, 4))).toBe('refused');
    expect(tableMoveKind(3, 8, mv(10, 1, 8))).toBe('refused');
    // partly over the table
    expect(tableMoveKind(3, 8, mv(2, 2, 6))).toBe('refused');
    expect(tableMoveKind(3, 8, mv(7, 3, 1))).toBe('refused');
    // nothing happens
    expect(tableMoveKind(3, 8, mv(5, 2, 5))).toBe('shift');
    expect(tableMoveKind(3, 8, mv(5, 2, 7))).toBe('shift');
  });

  it('the refusals of the R2b table suite (t03 t06 t07 t21) are exactly the moves classified as refused', () => {
    const R2B = readFixture<{ ops: Array<{ opId: string; suite: string; primitive: string; params: { src?: { start: number; end: number }; before?: { index: number } }; error: string | null }> }>('r2b-fixtures.json');
    const ops = R2B.ops.filter((o) => o.suite === 'table' && o.primitive === 'moveCols');
    expect(ops.length).toBeGreaterThanOrEqual(10);
    const refused: string[] = [], ours: string[] = [];
    for (const o of ops) {
      const op: Extract<StructOp, { type: 'move' }> = { type: 'move', axis: 'col', from: o.params.src!.start, count: o.params.src!.end - o.params.src!.start + 1, to: o.params.before!.index };
      if (o.error) refused.push(o.opId);
      if (tableMoveKind(3, 8, op) === 'refused') ours.push(o.opId);
    }
    expect(ours.sort()).toEqual(refused.sort());
    expect(refused.sort()).toEqual(['t03', 't06', 't07', 't21']);
  });
});
