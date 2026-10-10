// Types for tools/lib/copies.mjs (owner: T): the sheets copied from the other files, judged against their sources.
import type { InspectResult } from './inspect.mjs';
import type { PlannedSheetName } from './sheetnames.mjs';

export function stripPrefix(f: string): string;
export function textSheetName(fileName: string): string;
export function sheetNamesOf(dump: InspectResult | null | undefined): string[];
export function expectedCopies(merged: string, keptBase: string[], sources: Array<{ name: string; dump: InspectResult; strict?: boolean }>): Array<PlannedSheetName & { dump: InspectResult; strict: boolean }>;
export function checkCopies(ctx: {
  base: InspectResult;
  out: InspectResult;
  merged: string;
  sources: Array<{ name: string; dump: InspectResult | null | undefined; strict?: boolean }>;
  /** the OUTPUT is a Strict Open XML workbook (default false) */
  outStrict?: boolean;
  add: (level: 'fail' | 'warn' | 'info' | 'pass', code: string, message: string, examples?: string[]) => void;
}): { compared: number; expected: string[] };
