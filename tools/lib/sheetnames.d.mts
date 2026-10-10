// Types for tools/lib/sheetnames.mjs (owner: T): a plain-JavaScript port of src/core/xlsx/sheetnames.ts.
export const SHEET_NAME_MAX: number;
export function fileBaseName(fileName: string): string;
export function sheetNameKey(name: string): string;
export function cutUnits(s: string, units: number): string;
export function sanitizeSheetName(raw: string, fallback?: string): string;
export function joinSheetName(file: string, sheet: string, max?: number): string;
export function withSuffix(name: string, n: number, max?: number): string;
export interface PlannedSheetName {
  file: string;
  original: string;
  name: string;
  suffixed: boolean;
}
export function planSheetNames(existing: readonly string[], requests: ReadonlyArray<{ file: string; sheets: string[] }>): PlannedSheetName[];
