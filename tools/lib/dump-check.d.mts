// Types for tools/lib/dump-check.mjs (owner: T).
import type { InspectResult } from './inspect.mjs';

export const MAX_CHECK_BYTES: number;

export interface PackageFacts {
  names: Array<{ name: string; scope: string | null }>;
  sheets: Array<{
    name: string;
    /** A1 -> sorted reference signature of the formula ('名簿!1,4,1,63;集計!2,5,2,5') */
    formulas: Map<string, string>;
    /** A1 -> text of the string constants (the first 600, row order) */
    texts: Map<string, string>;
    cf: number;
    dv: number;
    tables: number;
    charts: number;
    pivots: number;
  }>;
}

export function refSignature(formula: string, sheet: string | null): string;
export function packageFacts(file: string): Promise<PackageFacts | null>;
export function dumpIssues(facts: PackageFacts | null, dump: InspectResult | null | undefined): string[];
export function dumpDifferences(a: InspectResult, b: InspectResult): string[];
