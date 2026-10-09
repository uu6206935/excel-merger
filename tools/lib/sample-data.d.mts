// Types for tools/lib/sample-data.mjs (owner: T).

export type CellInput = string | number | boolean | null;

export interface Student {
  index: number;
  id: string;
  family: string;
  given: string;
  name: string;
  kana: string;
  romaji: string;
  uni: string;
  uniShort: string;
  faculty: string;
  grade: string;
  birth: number;
  age: number;
  email: string | null;
  phone: string;
  aspiration: number;
  note: string | null;
}

export interface Rng {
  next(): number;
  int(lo: number, hi: number): number;
  pick<T>(arr: readonly T[]): T;
  weighted<T>(pairs: ReadonlyArray<readonly [T, number]>): T;
  shuffle<T>(arr: readonly T[]): T[];
}

export function serial(y: number, m: number, d: number): number;
export function fromSerial(n: number): { y: number; m: number; d: number };
export function rng(seed: number): Rng;
export function toHalfKana(s: string): string;
export function toFullDigits(s: string): string;
export function foldName(s: string): string;
export function toCsv(headers: string[], rows: CellInput[][]): string;

export const FAMILIES: Array<[string, string, string]>;
export const GIVENS: Array<[string, string, string]>;
export const UNIVERSITIES: Array<{ name: string; short: string; domain: string; faculties: string[] }>;
export const GRADES: Array<[string, number]>;
export const REF_DATE: { y: number; m: number; d: number };
export const INTERVIEWERS: string[];
export const ITAIJI: Record<string, string>;
export const SPECIAL: Record<number, { family: [string, string, string]; given: [string, string, string]; uni: string; faculty?: string; note: string }>;
export const N_STUDENTS: number;
export const NO_EMAIL: number[];
export const FILES: { base: string; interview: string; contacts: string; aptitude: string; attendance: string };

export const ROSTER: { sheet: string; titleRow: number; headerRow: number; firstDataRow: number; headers: string[]; helperCol: number };
export const INTERVIEW: {
  sheet: string;
  tableSheet: string;
  tableName: string;
  headers: string[];
  missing: number[];
  resubmitted: number[];
  extras: Array<{ id: string; name: string }>;
};
export const CONTACTS: { headers: string[]; extras: Array<{ name: string; kana: string; email: string; phone: string }> };
export const APTITUDE: { sheet: string; headers: string[]; extras: Array<{ kana: string }> };
export const ATTENDANCE: { sheet: string; title: string; headerRow: number; headers: string[]; extras: Array<{ name: string; uni: string }> };

export function makeStudents(): Student[];
export function rosterRows(students: Student[]): CellInput[][];

export interface SummaryBlock {
  sheet: string;
  title: string;
  note: string;
  headerRow: number;
  headers: string[];
  uniRows: CellInput[][];
  totalRow: number;
  total: CellInput[];
  gradeHeaderRow: number;
  gradeRows: CellInput[][];
  otherHeaderRow: number;
  others: CellInput[][];
}
export function summaryBlock(students: Student[]): SummaryBlock;

export interface InterviewMeta {
  row: number;
  baseIndex: number | null;
  idKind: 'text' | 'number' | 'fullwidth';
  extra: boolean;
  resubmitted: boolean;
  name: string;
}
export function interviewRows(students: Student[]): { rows: CellInput[][]; meta: InterviewMeta[] };
export function contactRows(students: Student[]): { rows: CellInput[][]; meta: Array<{ row: number; baseIndex: number | null }> };
export function aptitudeRows(students: Student[]): {
  rows: CellInput[][];
  meta: Array<{ row: number; baseIndex: number | null; kanaKind: 'halfwidth' | 'nospace' | 'fullwidth'; kana: string }>;
};
export function attendanceRows(students: Student[]): { rows: CellInput[][]; meta: Array<{ row: number; baseIndex: number | null; name: string; uni: string }> };

export interface SampleSpec {
  students: Student[];
  base: {
    title: string;
    note: string;
    roster: typeof ROSTER;
    rows: CellInput[][];
    master: { universities: string[]; faculties: string[]; grades: string[]; aspiration: Array<[number, string]>; refDate: number };
    summary: SummaryBlock;
  };
  interview: typeof INTERVIEW & { rows: CellInput[][]; meta: InterviewMeta[] };
  contacts: { headers: string[]; rows: CellInput[][]; meta: Array<{ row: number; baseIndex: number | null }> };
  aptitude: typeof APTITUDE & { rows: CellInput[][]; meta: Array<{ row: number; baseIndex: number | null; kanaKind: string; kana: string }> };
  attendance: typeof ATTENDANCE & { rows: CellInput[][]; meta: Array<{ row: number; baseIndex: number | null; name: string; uni: string }> };
}
export function buildSpec(): SampleSpec;

export interface TruthExpected {
  baseRows?: number;
  matchedBaseRows: number;
  unmatchedBaseRows: number[];
  extraRows: number[];
  duplicateIds?: string[];
}
export interface Truth {
  note: string;
  files: typeof FILES;
  base: {
    file: string;
    sheet: string;
    sheets: string[];
    headerRow: number;
    firstDataRow: number;
    lastDataRow: number;
    footerRows: number[];
    headers: string[];
    hiddenColumns: string[];
    freeze: { cell: string; rows: number; cols: number };
    names: Record<string, string>;
    summarySheet: string;
    summaryRefs: Array<{ cell: string; headers: string[] }>;
    students: Array<{ row: number; id: string; name: string; kana: string; uni: string; email: string | null; note: string | null }>;
    sameNamePair: number[];
  };
  interview: { file: string; sheet: string; headerRow: number; key: string; rows: Array<{ row: number; baseRow: number | null; idKind: string; resubmitted: boolean; extra: boolean }>; expected: TruthExpected };
  contacts: { file: string; encoding: string; headerRow: number; key: string; rows: Array<{ row: number; baseRow: number | null }>; expected: TruthExpected };
  aptitude: { file: string; sheet: string; headerRow: number; key: string; rows: Array<{ row: number; baseRow: number | null; kanaKind: string }>; expected: TruthExpected };
  attendance: { file: string; sheet: string; headerRow: number; key: string; rows: Array<{ row: number; baseRow: number | null; uni: string }>; expected: TruthExpected };
}
export function buildTruth(spec: SampleSpec): Truth;
