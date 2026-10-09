// Performance budget: a column move on a 10,000 x 60 doc with 50,000 formulas must take < 500 ms
// (structural sharing: only columns whose cells change are copied; formulas are tokenized once).
// Wall-clock timing on a shared machine is noisy (other test files run in parallel), so every operation is timed
// up to 3 times on the same input (ops are pure) and the best time must meet the budget; the first, cold run
// (JIT not warmed up yet, like the user's first drag) is logged and sanity-checked with a generous bound.
// Typical numbers (plain node, idle machine): create ~130 ms, first move ~200 ms, later moves ~150 ms.
import { describe, expect, it } from 'vitest';
import { analyzeDelete, createEditor, deleteColumns, formulaForDisplay, insertColumns, moveColumns, setCellInput } from '../../src/core/doc';
import { ok } from './helpers';
import { ROWS, bigBook } from './bigbook';

interface Timing {
  first: number;
  best: number;
}
function measure<T>(fn: () => T, budget: number): [T, Timing] {
  let out!: T;
  let first = 0;
  let best = Infinity;
  for (let k = 0; k < 3; k++) {
    const t0 = performance.now();
    out = fn();
    const t = performance.now() - t0;
    if (k === 0) first = t;
    best = Math.min(best, t);
    if (best < budget / 2) break;
  }
  return [out, { first, best }];
}

describe('performance (10,000 rows x 60 columns, 50,000 formulas)', () => {
  it('moves, inserts and deletes within budget', () => {
    const book = bigBook();
    const [s0, tCreate] = measure(() => createEditor(book, { sheet: 'Data', insertAt: 0, imported: [], appendAt: null, appended: [] }, {}, { headerRow: 1, firstDataRow: 2, lastDataRow: ROWS + 1 }), 0);
    let formulas = 0;
    for (const c of s0.doc.columns) for (const cell of c.cells) if (cell?.f) formulas++;
    expect(formulas).toBe(50_000);

    const ids = s0.doc.columns.map((c) => c.id);
    const [m1, tMove] = measure(() => moveColumns(s0, [ids[2], ids[3], ids[4]], 40), 500);
    const s1 = ok(m1);
    expect(s1.doc.columns[55].cells[1]?.f).toBe('A2*2');
    // C:E moved out of A:F: each endpoint follows its column (R2b R3), F is now column C
    expect(s1.doc.columns[58].cells[1]?.f).toBe('VLOOKUP(A2,$A$2:$C$10001,3,FALSE)');
    const [m2, tMove2] = measure(() => moveColumns(s1, [ids[0], ids[10], ids[20], ids[59]], 30), 500);
    ok(m2);
    const [m3, tMoveBack] = measure(() => moveColumns(s0, [ids[0]], 60), 500);
    ok(m3);
    const [ins, tInsert] = measure(() => insertColumns(s1, 5, 2), 500);
    ok(ins);
    const [imp, tAnalyze] = measure(() => analyzeDelete(s0, [ids[0]]), 1500);
    expect(imp.breakCount).toBe(20_020);
    const [del, tDelete] = measure(() => deleteColumns(s0, [ids[0]], { convertDependents: true }), 2000);
    const sd = ok(del);
    expect(sd.doc.columns.some((c) => c.cells.some((x) => x?.f?.includes('#REF!')))).toBe(false);
    const [ed, tEdit] = measure(() => setCellInput(s1, ids[7], 500, '123'), 50);
    ok(ed);
    const [, tDisplay] = measure(() => {
      for (let r = 2; r < 60; r++) for (let c = 50; c < 60; c++) formulaForDisplay(s1.doc, c, r);
    }, 50);

    const report = { tCreate, tMove, tMove2, tMoveBack, tInsert, tAnalyze, tDelete, tEdit, tDisplay };
    console.log(Object.entries(report).map(([k, v]) => `${k}=${v.first.toFixed(0)}/${v.best.toFixed(0)}ms`).join(' '));
    expect(tMove.best).toBeLessThan(500);
    expect(tMove2.best).toBeLessThan(500);
    expect(tMoveBack.best).toBeLessThan(500);
    expect(tInsert.best).toBeLessThan(500);
    expect(tEdit.best).toBeLessThan(50);
    expect(tAnalyze.best).toBeLessThan(1500);
    expect(tDelete.best).toBeLessThan(2000);
    // the cold first move: generous sanity bound (the machine may be busy with other work)
    expect(tMove.first).toBeLessThan(1500);
  }, 120_000);
});
