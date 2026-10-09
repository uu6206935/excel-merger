// Types for tools/lib/verdict.mjs (owner: T).
import type { InspectResult } from './inspect.mjs';

export interface VerdictItem {
  level: 'fail' | 'warn' | 'info' | 'pass';
  code: string;
  message: string;
  examples?: string[];
}

export interface Verdict {
  ok: boolean;
  verdict: 'PASS' | 'FAIL';
  items: VerdictItem[];
  stats: {
    mergedSheet?: string | null;
    headerRow?: number | null;
    refText?: { base: number; out: number };
    errors?: { base: Record<string, number>; out: Record<string, number> };
    formulasChecked?: number;
    rows?: { base: number; out: number };
    cols?: { base: number; out: number };
    addedHeaders?: string[];
    removedHeaders?: string[];
    /** data block (rows under the header with values) of the merged sheet */
    dataRows?: { base: { first: number; last: number }; out: { first: number; last: number } | null };
    /** rows appended at the end of the table (join 'full') and how many bounded ranges over the table grew with it */
    appended?: { rows: number; firstRow: number; lastRow: number; ranges: number; extended: number; notExtended: number };
    /** CF / DV rules whose range grew over a column next to a column carrying them (info, not a failure) */
    ruleGrowth?: { cf: number; dv: number };
  };
}

export function checkOutput(
  base: InspectResult | null | undefined,
  out: InspectResult | null | undefined,
  opts?: { sheet?: string; headerRow?: number; maxExamples?: number },
): Verdict;
export function formatVerdict(v: Verdict, title?: string): string;
export function dataBlock(sheet: import('./inspect.mjs').InspectSheet | null | undefined, headerRow: number | null | undefined): { first: number; last: number } | null;
