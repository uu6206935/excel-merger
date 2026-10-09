// Types for tools/lib/compare-core.mjs (owner: T).
import type { InspectResult, InspectSheet } from './inspect.mjs';

export interface Box {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}
export interface FormulaRef {
  /** explicit sheet, or the current sheet passed to extractRefs (null when none was given) */
  sheet: string | null;
  box: Box;
  text: string;
  qualified: boolean;
}

export const MAX_ROW: number;
export const MAX_COL: number;
export function colToNum(letters: string): number;
export function numToCol(n: number): string;
export function parseCell(addr: string): { row: number; col: number } | null;
export function cellAddr(row: number, col: number): string;
export function parseRange(text: string): Box | null;
export function parseRanges(text: string | null | undefined): Box[];
export function extractRefs(formula: string, currentSheet?: string | null): FormulaRef[];
/** shifts the relative parts of every A1 reference (shared-formula expansion); out-of-sheet -> #REF! */
export function shiftFormula(formula: string, dr: number, dc: number): string;
/** NOW / TODAY / RAND / RANDBETWEEN / RANDARRAY / CELL / INFO calls */
export const VOLATILE_RE: RegExp;
export function sheetByName(dump: InspectResult | null | undefined, name: string): InspectSheet | null;
export function headerRowGuess(sheet: InspectSheet | null | undefined): number | null;
export function cellValue(sheet: InspectSheet | null | undefined, addr: string): string | number | boolean | undefined;
export function headerText(sheet: InspectSheet | null | undefined, headerRow: number | null | undefined, col: number): string | null;
export function headerIndex(sheet: InspectSheet | null | undefined, headerRow: number | null | undefined): Map<string, number[]>;
export function arr<T>(x: T | T[] | null | undefined): T[];
export function totalErrors(dump: InspectResult): Record<string, number>;
export function summarize(dump: InspectResult): string;

export interface ComparisonRow {
  group: string;
  sheet: string;
  ok: boolean;
  detail: string;
  /** a known-benign difference that was NOT counted (listed separately) */
  filtered?: boolean;
}
export function compareInspections(a: InspectResult, b: InspectResult, opts?: { ignoreValues?: boolean; pivotLayout?: boolean }): ComparisonRow[];
export function formatComparison(rows: ComparisonRow[], opts?: { onlyDiffs?: boolean; showFiltered?: boolean }): string;
/** 'formulas 2, styles 1': groups with real differences (filtered rows excluded) */
export function diffSummary(rows: ComparisonRow[]): string;
