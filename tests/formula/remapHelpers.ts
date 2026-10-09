/**
 * Helpers of the FormulaRemapper tests (tests/formula/remap*.test.ts).
 * Source layout: sheet 評価, header row 1, data rows 2..11, columns A..F (A 学生ID and B 氏名 are key columns that are NOT imported unless a
 * test says so; C..F are imported to the merged columns E..H). The default row mapping is the identity.
 */
import { FormulaRemapper, type RemapReason, type RemapSpec } from '../../src/core/formula/remap';

export const SHEET = '評価';

export function range(a: number, b: number): number[] {
  return Array.from({ length: b - a + 1 }, (_x, i) => a + i);
}

/** [source row, merged row] pairs: identity / reversed (r -> 13 - r) / shifted (r -> r + 10) */
export const identity = (): Array<[number, number]> => range(2, 11).map((r) => [r, r]);
export const reversed = (): Array<[number, number]> => range(2, 11).map((r) => [r, 13 - r]);
export const shifted = (): Array<[number, number]> => range(2, 11).map((r) => [r, r + 10]);

export function spec(over: Partial<RemapSpec> = {}): RemapSpec {
  return {
    sheet: SHEET,
    region: { headerRow: 1, firstDataRow: 2, lastDataRow: 11, firstCol: 1, lastCol: 6 },
    target: { headerRow: 1, firstDataRow: 2, lastDataRow: 11 },
    colMap: new Map([[3, 5], [4, 6], [5, 7], [6, 8]]),
    rowPairs: identity(),
    ...over,
  };
}

export type Over = Partial<RemapSpec>;

export function run(f: string, sr = 5, tr = sr, over: Over = {}) {
  return new FormulaRemapper(spec(over)).remap(f, sr, tr);
}

/** the remapped formula; throws when the formula could not be remapped */
export function ok(f: string, sr = 5, tr = sr, over: Over = {}): string {
  const r = run(f, sr, tr, over);
  if (!r.ok) throw new Error(`expected ok for ${f}: ${r.reason} ${r.detail ?? ''}`);
  return r.formula;
}

/** 'ok' or the reason the formula could not be remapped */
export function why(f: string, sr = 5, tr = sr, over: Over = {}): RemapReason | 'ok' {
  const r = run(f, sr, tr, over);
  return r.ok ? 'ok' : r.reason;
}

export function names(list: Array<[string, string | null, string]>): Over {
  return { names: list.map(([name, scope, refersTo]) => ({ name, scope, refersTo })) };
}
