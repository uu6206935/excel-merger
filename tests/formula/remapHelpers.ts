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

// ---------------------------------------------------------------------------------------------------------------------
// round 9: the copies of the source's sheets (PremergeOptions.includeSourceSheets), see remap-copies*.test.ts
// ---------------------------------------------------------------------------------------------------------------------

/** lower-cased sheet of the source -> its copy in the output (names as planSheetNames makes them: "<file>_<sheet>", "(1)" for a later duplicate) */
export const COPIES: ReadonlyMap<string, string> = new Map([
  ['評価', 'src_評価'],
  ['レート', 'src_レート'],
  ['master', 'src_Master'],
  ['my sheet', 'src_My Sheet'],
  ["o'brien", "src_O'Brien"],
  ['sheet2', 'src_Sheet2(1)'],
]);

/** only C and D are imported (to E and F): the columns A, B (keys) and E, F of the source are NOT in the merged table */
export const PARTIAL: Over = { colMap: new Map([[3, 5], [4, 6]]) };
/** the key columns A and B are not imported, but the base has the same text in its columns A and B */
export const KEYS: Over = { partnerCols: new Map([[1, 1], [2, 2]]) };
/** source rows 2..11 without row 4 / 7: those people matched nothing */
export const without = (...rows: number[]): Array<[number, number]> => identity().filter(([r]) => !rows.includes(r));

/** remap with the copies (`array`: the source cell holds an array / dynamic-array formula) */
export function runC(f: string, sr = 5, tr = sr, over: Over = {}, array = false, srcCol?: number) {
  return new FormulaRemapper(spec({ sheetCopies: COPIES, ...over })).remap(f, sr, tr, srcCol, array);
}

export function okC(f: string, sr = 5, tr = sr, over: Over = {}, array = false): string {
  const r = runC(f, sr, tr, over, array);
  if (!r.ok) throw new Error(`expected ok for ${f}: ${r.reason} ${r.detail ?? ''}`);
  return r.formula;
}

export function whyC(f: string, sr = 5, tr = sr, over: Over = {}, array = false): RemapReason | 'ok' {
  const r = runC(f, sr, tr, over, array);
  return r.ok ? 'ok' : r.reason;
}

/** did the formula come out reading a copy? (undefined for a failure) */
export function viaC(f: string, sr = 5, tr = sr, over: Over = {}, array = false): boolean | undefined {
  const r = runC(f, sr, tr, over, array);
  return r.ok ? r.viaCopy === true : undefined;
}
