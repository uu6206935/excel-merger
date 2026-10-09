import { describe, expect, it } from 'vitest';
import {
  applyBlockMove,
  cellAddress,
  colLetter,
  computeMove,
  isNoopMove,
  letterToIndex,
  planBlockMoves,
  rangeBetween,
  toRuns,
} from '../../src/components/grid/colops';

const L = (s: string) => s.split('');

describe('computeMove (non-contiguous, gap in pre-move coordinates)', () => {
  it('hand-checked cases', () => {
    expect(computeMove(L('ABCDEFGHIJ'), new Set(L('BDG')), 9).join('')).toBe('ACEFHIBDGJ'); // before J
    expect(computeMove(L('ABCDEFGHIJ'), new Set(L('BDG')), 10).join('')).toBe('ACEFHIJBDG'); // at the end
    expect(computeMove(L('ABCDEFGHIJ'), new Set(L('HJ')), 0).join('')).toBe('HJABCDEFGI'); // to the front
    expect(computeMove(L('ABCDEFGHIJ'), new Set(L('BDG')), 4).join('')).toBe('ACBDGEFHIJ'); // inside the span
    expect(computeMove(L('ABCDEFGHIJ'), new Set(L('CD')), 3).join('')).toBe('ABCDEFGHIJ'); // inside a block
  });
  it('clamps the gap and ignores unknown ids', () => {
    expect(computeMove(L('ABC'), new Set(L('A')), 99).join('')).toBe('BCA');
    expect(computeMove(L('ABC'), new Set(L('C')), -5).join('')).toBe('CAB');
    expect(computeMove(L('ABC'), new Set(['X']), 1).join('')).toBe('ABC');
    expect(computeMove([], new Set(['X']), 0)).toEqual([]);
  });
  it('isNoopMove', () => {
    expect(isNoopMove(L('ABCDE'), new Set(L('CD')), 2)).toBe(true);
    expect(isNoopMove(L('ABCDE'), new Set(L('CD')), 4)).toBe(true);
    expect(isNoopMove(L('ABCDE'), new Set(L('CD')), 5)).toBe(false);
    expect(isNoopMove(L('ABCDE'), new Set<string>(), 0)).toBe(true);
  });
});

describe('letters and addresses', () => {
  it('colLetter / letterToIndex round trip', () => {
    expect(colLetter(0)).toBe('A');
    expect(colLetter(25)).toBe('Z');
    expect(colLetter(26)).toBe('AA');
    expect(colLetter(701)).toBe('ZZ');
    expect(colLetter(702)).toBe('AAA');
    expect(colLetter(16383)).toBe('XFD');
    for (const i of [0, 1, 25, 26, 51, 52, 701, 702, 16383]) expect(letterToIndex(colLetter(i))).toBe(i);
    expect(letterToIndex('xfd')).toBe(16383);
    expect(letterToIndex('')).toBe(-1);
    expect(letterToIndex('A1')).toBe(-1);
    expect(letterToIndex('ABCD')).toBe(-1);
  });
  it('cellAddress', () => {
    expect(cellAddress(2, 5)).toBe('C5');
    expect(cellAddress(27, 10001)).toBe('AB10001');
    expect(cellAddress(-1, 5)).toBe('');
    expect(cellAddress(0, 0)).toBe('');
  });
  it('rangeBetween / toRuns', () => {
    expect(rangeBetween(L('ABCDEF'), 'E', 'B')).toEqual(L('BCDE'));
    expect(rangeBetween(L('ABCDEF'), 'X', 'B')).toEqual(['B']);
    expect(rangeBetween(L('ABCDEF'), 'A', 'X')).toEqual([]);
    expect(toRuns([9, 1, 2, 3, 7, 10])).toEqual([
      [1, 3],
      [7, 7],
      [9, 10],
    ]);
    expect(toRuns([])).toEqual([]);
  });
});

describe('Excel block moves ("Insert Cut Cells")', () => {
  it('applyBlockMove', () => {
    expect(applyBlockMove(L('ABCDEFGHIJ'), { from: 2, count: 2, to: 7 }).join('')).toBe('ABEFGCDHIJ');
    expect(applyBlockMove(L('ABCDEFGHIJ'), { from: 7, count: 2, to: 2 }).join('')).toBe('ABHICDEFGJ');
  });
  it('planBlockMoves hand cases', () => {
    const p1 = planBlockMoves(L('AXBYC'), new Set(L('XY')), 5);
    expect(p1.final.join('')).toBe('ABCXY');
    const p2 = planBlockMoves(L('XabYc'), new Set(L('XY')), 4);
    expect(p2.final.join('')).toBe('abXYc');
    expect(p2.steps).toHaveLength(1);
    expect(planBlockMoves(L('ABCDE'), new Set(L('CD')), 2).steps).toHaveLength(0);
  });
  it('200,000 random moves: plan reproduces computeMove, never more steps than runs, 0 steps for no-ops', () => {
    let seed = 12345;
    const rnd = () => (seed = (seed * 1103515245 + 12345) & 0x7fffffff) / 0x7fffffff;
    let total = 0;
    for (let t = 0; t < 200000; t++) {
      const n = 1 + Math.floor(rnd() * 14);
      const order = Array.from({ length: n }, (_, i) => i);
      const moving = new Set(order.filter(() => rnd() < 0.35));
      const gap = Math.floor(rnd() * (n + 1));
      const { final, steps } = planBlockMoves(order, moving, gap);
      let cur: number[] = order;
      for (const s of steps) cur = applyBlockMove(cur, s);
      if (cur.join(',') !== final.join(',')) throw new Error(`mismatch ${JSON.stringify({ order, moving: [...moving], gap })}`);
      const runs = toRuns(order.filter((i) => moving.has(i))).length;
      if (steps.length > runs) throw new Error('too many steps');
      if (isNoopMove(order, moving, gap) && steps.length) throw new Error('steps for a no-op');
      // relative orders are kept
      const movedAfter = final.filter((i) => moving.has(i));
      if (movedAfter.join(',') !== order.filter((i) => moving.has(i)).join(',')) throw new Error('moved order changed');
      total += steps.length;
    }
    expect(total / 200000).toBeLessThan(2);
  });
});
