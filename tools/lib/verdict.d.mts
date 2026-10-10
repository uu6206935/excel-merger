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
    /** withCopies: how many copied sheets were compared with their source sheet */
    copies?: { compared: number };
  };
}

/** an other file whose sheets were copied into the output: the name the app saw (the dropped file's name) and its inspect dump */
export interface CopySource {
  name: string;
  dump: InspectResult | null | undefined;
  /** the source is a Strict Open XML workbook (tools/lib/strict.mjs): its charts / pictures cannot be copied into a normal workbook */
  strict?: boolean;
}

export interface CheckOptions {
  sheet?: string;
  headerRow?: number;
  maxExamples?: number;
  /**
   * the output is expected to hold ONLY the merged sheet (the app's default: 「他のシートも出力に含める」 off): the base's other sheets and what
   * belonged to them (names, charts, rules that read them) are INFO; a formula that reads a sheet that is not in the output FAILS
   */
  singleSheet?: boolean;
  /** the output also holds the sheets of the other files (<file>_<sheet>): every one must be there under its name and equal its source (`sources`) */
  withCopies?: boolean;
  /** the other files in card order (needed to compare the copies with their sources; names alone are checked without) */
  sources?: CopySource[];
  /** the output is a Strict Open XML workbook (default false) */
  outStrict?: boolean;
}

export function checkOutput(base: InspectResult | null | undefined, out: InspectResult | null | undefined, opts?: CheckOptions): Verdict;
export function formatVerdict(v: Verdict, title?: string): string;
export function dataBlock(sheet: import('./inspect.mjs').InspectSheet | null | undefined, headerRow: number | null | undefined): { first: number; last: number } | null;
