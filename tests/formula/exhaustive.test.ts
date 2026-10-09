// Exhaustive check of the column-move rule vs Excel 365: all 36 ranges inside A..H x all 128 block moves.
import { describe, expect, it } from 'vitest';
import { adjustFormula, type StructOp } from '../../src/core/formula/excelRef';
import { readFixture } from './helpers';

const R = readFixture<Record<string, { S: string[] }>>('exhaustive-results.json');
const inp = readFixture<{ ops: { name: string }[]; formulas: string[] }>('exhaustive-input.json');

describe('exhaustive block moves (Excel 365 ground truth)', () => {
  it('matches Excel on 4,608 / 4,608 (range x move) pairs', () => {
    let pass = 0;
    const fails: string[] = [];
    for (const o of inp.ops) {
      const [, from, count, to] = o.name.split('_').map(Number);
      const op: StructOp = { type: 'move', axis: 'col', from, count, to };
      inp.formulas.forEach((f, i) => {
        const want = R[o.name].S[i];
        const got = '=' + adjustFormula(f.slice(1), op, { hostSheet: 'F', targetSheet: 'Data' }).formula;
        if (got === want) pass++;
        else fails.push(`${o.name} (block ${String.fromCharCode(64 + from)}x${count} -> before ${String.fromCharCode(64 + to)}) ${f} got ${got} want ${want}`);
      });
    }
    expect(fails.slice(0, 50)).toEqual([]);
    expect(pass).toBe(4608);
  });
});
