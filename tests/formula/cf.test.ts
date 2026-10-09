// CF / DV rule adjustment vs Excel 365 (fixtures/formula/cf-results.json), compared PER CELL (effective R1C1
// formula of every cell), because Excel fragments rules arbitrarily. Also: planMove() permutation checks.
import { describe, expect, it } from 'vitest';
import {
  adjustAnchoredRule, parseAreaText, relKey, translateFormula, planMove, movePerm, adjustFormulaSeq,
  type StructOp, type Area,
} from '../../src/core/formula/excelRef';
import { readFixture } from './helpers';

const R = readFixture<Record<string, { cf: string[]; dv: string | null }>>('cf-results.json');
const OPS: Record<string, StructOp> = {
  del_B: { type: 'delete', axis: 'col', at: 2, count: 1 }, del_C: { type: 'delete', axis: 'col', at: 3, count: 1 },
  ins_B: { type: 'insert', axis: 'col', at: 2, count: 1 }, ins_C: { type: 'insert', axis: 'col', at: 3, count: 1 },
  mv_B_to_F: { type: 'move', axis: 'col', from: 2, count: 1, to: 6 }, mv_E_to_B: { type: 'move', axis: 'col', from: 5, count: 1, to: 2 },
  mv_C_to_B: { type: 'move', axis: 'col', from: 3, count: 1, to: 2 }, del_row2: { type: 'delete', axis: 'row', at: 2, count: 1 },
};
// Excel's own move results contain #REF! artifacts from its internal insert+cut+delete; we keep intent instead
// (DEVIATIONS.md "CF move artifacts"). mv_B_to_F|=A2>0: in this 8-rule block Excel fragments the range into one object
// per column, so every cell of the part right of the moved column keeps its own formula instead of sharing the first
// cell's (DEVIATIONS.md "CF rule blocks").
const KNOWN_EXCEL_ARTIFACT = new Set(['mv_E_to_B|=A2>0', 'mv_C_to_B|=A2>0', 'mv_B_to_F|=A2>0']);

/** cell -> the set of effective formulas covering it (a cell can be covered twice: a moved cell and a spread one) */
function perCell(rules: { sqref: Area[]; formula: string }[]) {
  const m = new Map<string, Set<string>>();
  for (const r of rules) {
    const a0 = r.sqref[0]; const ar = Math.min(a0.r1, a0.r2), ac = Math.min(a0.c1, a0.c2);
    for (const a of r.sqref) for (let c = Math.min(a.c1, a.c2); c <= Math.max(a.c1, a.c2); c++) for (let row = Math.min(a.r1, a.r2); row <= Math.max(a.r1, a.r2); row++) {
      const k = `${c},${row}`;
      (m.get(k) ?? m.set(k, new Set()).get(k)!).add(relKey(translateFormula(r.formula, row - ar, c - ac), row, c));
    }
  }
  return new Map([...m].map(([k, v]) => [k, [...v].sort().join(' || ')]));
}

describe('CF / DV anchored rules vs Excel', () => {
  it('per-cell effective formulas equal Excel (61 checks, 3 known Excel artifacts)', () => {
    let pass = 0, known = 0;
    const fails: string[] = [];
    const base = R.none.cf.map((l) => l.split(' | '));
    for (const [opName, op] of Object.entries(OPS)) {
      const ex = R[opName].cf.map((l) => l.split(' | '));
      const perRule = base.length, frag = ex.length / perRule;
      base.forEach(([sq, f], i) => {
        const exRules = Array.from({ length: frag }, (_, k) => ex[k * perRule + i]).map(([s, ff]) => ({ sqref: s.split(',').map((x) => parseAreaText(x)!), formula: ff.slice(1) }));
        const want = perCell(exRules);
        // the 8 rules share one range (= one <conditionalFormatting> block): each is a sibling of the others
        const siblings = base.filter((_, j) => j !== i).map(([, g]) => [g.slice(1)]);
        const got = perCell(adjustAnchoredRule({ sqref: sq.split(',').map((x) => parseAreaText(x)!), formulas: [f.slice(1)] }, op, 'Data', { kind: 'cf', type: 'expression', siblings }).map((r) => ({ sqref: r.sqref, formula: r.formulas[0] })));
        const same = got.size === want.size && [...want].every(([k, v]) => got.get(k) === v);
        if (same) pass++;
        else if (KNOWN_EXCEL_ARTIFACT.has(`${opName}|${f}`)) known++;
        else fails.push(`${opName} ${f}\n     excel ${JSON.stringify([...want].filter(([k]) => k.endsWith(',2')))}\n     ours  ${JSON.stringify([...got].filter(([k]) => k.endsWith(',2')))}`);
      });
    }
    for (const [opName, op] of Object.entries(OPS)) {
      const dv = R[opName].dv; if (!dv) continue;
      const [cell, f] = dv.split(' | ');
      const cells = perCell(adjustAnchoredRule({ sqref: [parseAreaText('B2:E10')!], formulas: ['C2<>""'] }, op, 'Data').map((r) => ({ sqref: r.sqref, formula: r.formulas[0] })));
      const cc = parseAreaText(cell)!;
      const want = relKey(f.slice(1), cc.r1, cc.c1);
      if (cells.get(`${cc.c1},${cc.r1}`) === want) pass++; else fails.push(`DV ${opName}: ${cell} want ${want} got ${cells.get(`${cc.c1},${cc.r1}`)}`);
    }
    expect(fails).toEqual([]);
    expect(pass).toBe(61);
    expect(known).toBe(3);
  });
});

describe('planMove', () => {
  function simulate(n: number, ops: StructOp[]) {
    let arr = Array.from({ length: n }, (_, i) => i + 1);
    for (const op of ops) {
      if (op.type !== 'move') continue;
      const p = movePerm(op.from, op.count, op.to);
      const next = new Array<number>(n);
      arr.forEach((orig, i) => { next[p(i + 1) - 1] = orig; });
      arr = next;
    }
    return arr;
  }
  it('produces the intended final order for 5,000 random selections', () => {
    let seed = 7; const rnd = () => { seed = (seed * 16807) % 2147483647; return seed / 2147483647; };
    let okCount = 0;
    for (let trial = 0; trial < 5000; trial++) {
      const n = 12;
      const sel = Array.from({ length: n }, (_, i) => i + 1).filter(() => rnd() < 0.3);
      if (!sel.length) continue;
      const target = 1 + Math.floor(rnd() * (n + 1));
      const got = simulate(n, planMove(sel, target));
      const rest = Array.from({ length: n }, (_, i) => i + 1).filter((c) => !sel.includes(c));
      const want = [...rest.filter((c) => c < target), ...sel, ...rest.filter((c) => c >= target)];
      expect(got).toEqual(want);
      okCount++;
    }
    expect(okCount).toBeGreaterThan(4900);
  });
  it('decomposes [B,E,F,I] before D as documented', () => {
    const ops = planMove([2, 5, 6, 9], 4);
    expect(ops).toEqual([
      { type: 'move', axis: 'col', from: 2, count: 1, to: 4 },
      { type: 'move', axis: 'col', from: 5, count: 2, to: 4 },
      { type: 'move', axis: 'col', from: 9, count: 1, to: 6 },
    ]);
    // final order A C B E F I D G H J: every single-cell ref follows its cell
    expect(adjustFormulaSeq('A1+B1+C1+D1+E1+F1+G1+H1+I1+J1', ops, { hostSheet: 'S', targetSheet: 'S' }).formula)
      .toBe('A1+C1+B1+G1+D1+E1+H1+I1+F1+J1');
  });
  it('returns no ops for a no-op drop', () => {
    expect(planMove([3, 4], 3)).toEqual([]);
    expect(planMove([3, 4], 5)).toEqual([]);
    expect(planMove([], 5)).toEqual([]);
  });
});
