// Robustness of adjustAnchoredRule: random rules / ops / options never throw and keep the invariants that do not
// depend on Excel's quirks (the per-cell agreement with Excel is cfSpread / cfProbe / r2b / cf tests).
//  - every result area is a valid area of the sheet and every result has one formula per input formula;
//  - a clean op (inheritFormat: false) only ever covers cells that come from covered cells: mapIndex of every covered
//    cell, never more (no growth, no spread), and exactly the surviving cells for insert / move;
//  - with format inheritance the covered cells are a superset of the clean result (growth / spread only ADD cells),
//    except for moves, where the destination cells pasted over by the block may drop out of the rule;
//  - siblings (other rules of the block) never change which cells this rule covers in a clean op.
import { describe, expect, it } from 'vitest';
import {
  adjustAnchoredRule, mapIndex, MAX_COL, MAX_ROW, parseAreaText, type Area, type AnchoredRule, type StructOp,
} from '../../src/core/formula/excelRef';

let seed = Number(process.env.FUZZ_SEED) || 20260930;
const rnd = (n: number) => { seed = (seed * 1103515245 + 12345) & 0x7fffffff; return seed % n; };
const pick = <T,>(xs: T[]): T => xs[rnd(xs.length)];
const col = (c: number) => { let s = ''; for (let n = c; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s; return s; };

const FORMULAS = ['C2>0', '$C2>D2', 'SUM(B2:F2)>10', 'A$1="x"', 'COUNTIF($B:$B,C2)>1', 'F2>E2', '$A$1', '5', 'OR(B2>0,H2<0)', 'Other!C2>0', 'AA2>0', '#REF!>0', '$E2>$D$1'];
const cells = (areas: Area[]): Set<string> => {
  const out = new Set<string>();
  for (const a of areas) {
    if (a.kind === 'cols' || a.kind === 'rows') continue;
    for (let c = Math.min(a.c1, a.c2); c <= Math.max(a.c1, a.c2); c++) for (let r = Math.min(a.r1, a.r2); r <= Math.max(a.r1, a.r2); r++) out.add(`${c},${r}`);
  }
  return out;
};
function randomRule(): AnchoredRule {
  const sqref: Area[] = [];
  for (let i = 0, n = 1 + rnd(3); i < n; i++) {
    const c1 = 1 + rnd(14), r1 = 1 + rnd(8), c2 = c1 + rnd(5), r2 = r1 + rnd(4);
    const a = parseAreaText(c1 === c2 && r1 === r2 ? `${col(c1)}${r1}` : `${col(c1)}${r1}:${col(c2)}${r2}`);
    if (a) sqref.push(a);
  }
  const formulas = Array.from({ length: 1 + rnd(2) }, () => pick(FORMULAS));
  return { sqref, formulas };
}
function randomOp(): StructOp {
  const axis = pick(['col', 'row'] as const);
  const span = 16;
  switch (rnd(3)) {
    case 0: return { type: 'insert', axis, at: 1 + rnd(span), count: 1 + rnd(3) };
    case 1: return { type: 'delete', axis, at: 1 + rnd(span), count: 1 + rnd(3) };
    default: {
      const count = 1 + rnd(3), from = 1 + rnd(span);
      let to = 1 + rnd(span + 1);
      if (to > from && to <= from + count) to = from + count + 1; // skip no-ops and destinations inside the block
      return { type: 'move', axis, from, count, to };
    }
  }
}
const validArea = (a: Area): boolean => {
  const okC = a.kind === 'rows' || (a.c1 >= 1 && a.c2 >= a.c1 && a.c2 <= MAX_COL);
  const okR = a.kind === 'cols' || (a.r1 >= 1 && a.r2 >= a.r1 && a.r2 <= MAX_ROW);
  return okC && okR;
};

describe('adjustAnchoredRule fuzz', () => {
  it('3,000 random rules x ops x options: valid areas, clean ops never add cells, inheritance only adds', () => {
    let checked = 0;
    for (let k = 0; k < 3000; k++) {
      const rule = randomRule(), op = randomOp();
      const before = cells(rule.sqref);
      const isCol = op.axis === 'col';
      const image = (c: number, r: number): [number, number] | null => {
        const v = mapIndex(isCol ? c : r, op);
        return v == null ? null : isCol ? [v, r] : [c, v];
      };
      const expectClean = new Set<string>();
      for (const key of before) { const [c, r] = key.split(',').map(Number); const m = image(c, r); if (m) expectClean.add(`${m[0]},${m[1]}`); }
      if (op.type === 'insert') {
        // cells inserted strictly inside an area always join it (range semantics), whatever the options
        for (const a of rule.sqref) {
          const lo = isCol ? Math.min(a.c1, a.c2) : Math.min(a.r1, a.r2), hi = isCol ? Math.max(a.c1, a.c2) : Math.max(a.r1, a.r2);
          if (!(op.at > lo && op.at <= hi)) continue;
          const o1 = isCol ? Math.min(a.r1, a.r2) : Math.min(a.c1, a.c2), o2 = isCol ? Math.max(a.r1, a.r2) : Math.max(a.c1, a.c2);
          for (let v = op.at; v < op.at + op.count; v++) for (let o = o1; o <= o2; o++) expectClean.add(isCol ? `${v},${o}` : `${o},${v}`);
        }
      }

      const clean = adjustAnchoredRule(rule, op, 'S', { kind: 'cf', type: 'expression', inheritFormat: false });
      for (const res of [clean]) for (const x of res) { expect(x.formulas.length).toBe(rule.formulas.length); for (const a of x.sqref) expect(validArea(a)).toBe(true); }
      const cleanCells = cells(clean.flatMap((x) => x.sqref));
      if (op.type === 'move') {
        // every covered cell comes from a covered cell (the block lands on blank cells of a clean move)
        for (const key of cleanCells) expect(expectClean.has(key)).toBe(true);
        expect(cleanCells.size).toBe(expectClean.size);
      } else {
        for (const key of cleanCells) expect(expectClean.has(key)).toBe(true);
        if (op.type === 'insert') expect(cleanCells.size).toBe(expectClean.size);
      }
      // with siblings a clean op covers the same cells
      const sib = adjustAnchoredRule(rule, op, 'S', { kind: 'cf', type: 'expression', inheritFormat: false, siblings: [[pick(FORMULAS)]] });
      expect([...cells(sib.flatMap((x) => x.sqref))].sort()).toEqual([...cleanCells].sort());
      for (const x of sib) expect(x.formulas.length).toBe(rule.formulas.length);

      const inh = adjustAnchoredRule(rule, op, 'S', { kind: pick(['cf', 'dv'] as const), type: pick(['expression', 'cellIs', 'colorScale', 'list', undefined as unknown as string]) });
      for (const x of inh) { expect(x.formulas.length).toBe(rule.formulas.length); for (const a of x.sqref) expect(validArea(a)).toBe(true); }
      if (op.type !== 'move') {
        const inhCells = cells(inh.flatMap((x) => x.sqref));
        for (const key of cleanCells) expect(inhCells.has(key)).toBe(true); // growth only adds
      }
      checked++;
    }
    expect(checked).toBe(3000);
  });

  it('never throws on junk formulas and whole-axis areas', () => {
    const areas = ['A:A', 'B:D', '1:1', '2:5', 'XFD1', 'A1:XFD1', 'A1048576', 'C3:E1048576'].map((t) => parseAreaText(t)!);
    const formulas = ['', '(', '"abc', 'SUM(', 'A1:', '#REF!', 'Data!#REF!:.', 'A1.:.B2', 'Table1[[#This Row],[x]]'];
    for (let k = 0; k < 2000; k++) {
      const rule: AnchoredRule = { sqref: [pick(areas), ...(rnd(2) ? [pick(areas)] : [])], formulas: [pick(formulas)] };
      const op = randomOp();
      expect(() => adjustAnchoredRule(rule, op, 'S', { kind: pick(['cf', 'dv'] as const), siblings: rnd(2) ? [[pick(formulas)]] : undefined })).not.toThrow();
    }
  });
});
