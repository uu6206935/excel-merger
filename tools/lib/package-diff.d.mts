// Types for tools/lib/package-diff.mjs (owner: T).
import type { ComparisonRow } from './compare-core.mjs';

export interface XmlNode {
  name: string;
  local: string;
  attrs: Record<string, string>;
  children: XmlNode[];
  text: string;
}

export function decodeXml(s: string): string;
export function parseXml(text: string): XmlNode | null;

/** one cell of a parsed sheet (formula text is expanded for shared formulas) */
export interface PackageCell {
  row: number;
  col: number;
  s: number;
  /** 's' text, 'n' number, 'b' boolean, 'e' error, 'd' date, 'z' empty */
  type: string;
  value: unknown;
  f?: string;
  fType?: string;
}

/** the parts of a parsed sheet that callers read (see package-diff.mjs parseSheet for the rest) */
export interface PackageSheetModel {
  name: string;
  part: string;
  cells: Map<string, PackageCell>;
  cf: string[];
  dv: string[];
  tables: string[];
  drawings: string[];
  pivots: Array<{ name: string; [k: string]: unknown }>;
  merges: string[];
  [k: string]: unknown;
}

/** normalized package model (see package-diff.mjs) */
export interface PackageModel {
  files: string[];
  sheets: Array<{ name: string; state: string; part: string | null; kind: string }>;
  names: Array<{ name: string; scope: string | null; text: string; hidden: boolean }>;
  date1904: boolean;
  sheetModels: Map<string, PackageSheetModel>;
  [k: string]: unknown;
}

export interface FilteredCategory {
  key: string;
  label: string;
  /** null = always ignored (not counted) */
  count: number | null;
  examples: string[];
}

/** the serial number of an ISO 8601 date / date-time text (a t="d" cell), or null (not a date, or before 1900-03-01) */
export function isoDateSerial(text: string, date1904?: boolean): number | null;
export const FILTER_LABELS: Record<string, string>;
export function loadPackageModel(input: string | Uint8Array): Promise<PackageModel>;
export function comparePackageModels(a: PackageModel, b: PackageModel, opts?: { pivotLayout?: boolean }): { rows: ComparisonRow[]; filtered: FilteredCategory[] };
export function comparePackages(a: string | Uint8Array, b: string | Uint8Array, opts?: { pivotLayout?: boolean }): Promise<{ rows: ComparisonRow[]; filtered: FilteredCategory[] }>;
