/**
 * Cross-module API types (the second half of the contract; see types.ts and docs/ARCHITECTURE.md).
 *
 * Module facades and who owns them:
 *   src/core/xlsx/index.ts     readBook, cellText, buildPremerged, replayLog      (reader + package patcher)
 *   src/core/keys/index.ts     pickDataSheet, detectRegion, analyzeSources, defaultPlan, joinSources
 *   src/core/doc/index.ts      createEditor, moveColumns, insertColumns, analyzeDelete, deleteColumns,
 *                              setCellInput, setColumnWidth, formulaForDisplay, history helpers
 *   src/core/export/index.ts   exportWithJs, logToRunnerJob
 *   src/api.ts                 client for the local server (Excel本体 engine, save/open, heartbeat)
 * Function signatures are declared in each facade (stubs throw until implemented).
 */

import type {
  CellType,
  CheckItem,
  EditorSnapshot,
  LoggedOp,
  MergeReport,
  MergedRowPlan,
  SourceBook,
  TableRegion,
  JoinType,
  DuplicatePolicy,
} from './types';

export interface Progress {
  /** 0..1, or null when unknown */
  ratio: number | null;
  /** short Japanese message for the UI ('シートを読み込み中…') */
  message: string;
}

export type ProgressFn = (p: Progress) => void;

// ---------------------------------------------------------------------------
// keys
// ---------------------------------------------------------------------------

/** One loaded file as the keys module sees it. */
export interface KeySource {
  id: string;
  book: SourceBook;
  region: TableRegion;
  /** short label for messages (file name without extension) */
  label: string;
}

export interface NormalizeOptions {
  /** 異体字を同一視 (default true) */
  itaiji: boolean;
  /** 先頭ゼロを無視 (default true) */
  leadingZeros: boolean;
  /** 敬称を除去 (様/さん/殿/君, default true) */
  honorifics: boolean;
  /** 旧姓の括弧を除去 (default true) */
  maidenName: boolean;
  /** 大学名の略称・表記ゆれを統一 (default true) */
  universityAbbrev: boolean;
}

export interface ColumnDTO {
  col: number;
  header: string;
  type: string;
  typeLabel: string;
  confidence: string;
  typeScore: number;
  keyScore: number;
  uniqueness: number;
  fill: number;
  blank: number;
  dupRows: number;
  notes: string[];
}

export interface SheetDTO {
  sourceId: string;
  file: string;
  sheet: string;
  headerRow: number;
  rows: number;
  layoutHint?: string;
  columns: ColumnDTO[];
}

/** A key candidate between the base and ONE other source (R6 CandidateDTO). */
export interface CandidateDTO {
  /** unique within the pair */
  id: string;
  kind: 'single' | 'composite' | 'discovered';
  /** 'ID', '氏名', '氏名＋大学', '学籍番号 ⟷ 学生番号' */
  label: string;
  /** key family for the main-screen chips: 'id' | 'name' | 'other' */
  family: 'id' | 'name' | 'other';
  baseColumns: Array<{ col: number; header: string }>;
  otherColumns: Array<{ col: number; header: string }>;
  /** share of the other file's non-blank keys found in the base */
  matchRate: number;
  /** share of the base's non-blank keys found in the other file */
  coverage: number;
  matchedOther: number;
  nonBlankOther: number;
  matchedBase: number;
  nonBlankBase: number;
  uniqBase: number;
  uniqOther: number;
  dupRowsBase: number;
  dupRowsOther: number;
  blankBase: number;
  blankOther: number;
  ambiguous: number;
  cardinality: string;
  score: number;
  verdict: '良好' | '一部一致' | '不一致多数' | '使用不可' | string;
  /** pre-selected by the auto plan */
  selected: boolean;
  /** 1-based priority in the auto plan */
  tier?: number;
}

export interface PairAnalysis {
  sourceId: string;
  other: SheetDTO;
  mode: 'oneToOne' | 'lookup';
  candidates: CandidateDTO[];
}

/** Main-screen chip: ID / 氏名 are shown from the start. */
export interface KeyChip {
  family: 'id' | 'name';
  label: string;
  /** per other source: the best candidate of this family (null = not available in that file) */
  perSource: Record<string, { candidateId: string; label: string; matchRate: number } | null>;
}

export interface KeyAnalysis {
  baseId: string;
  base: SheetDTO;
  pairs: Record<string, PairAnalysis>;
  chips: KeyChip[];
  options: NormalizeOptions;
  /** opaque internal state (profiles etc.) reused by joinSources */
  internal?: unknown;
}

export interface JoinPlan {
  /** per other source: enabled candidate ids in priority order (first = tried first) */
  tiers: Record<string, string[]>;
  joinType: JoinType;
  duplicatePolicy: DuplicatePolicy;
  /** user-confirmed links (manual or accepted suggestions): base row <-> other row (1-based Excel rows) */
  manualLinks: Record<string, Array<{ baseRow: number; otherRow: number }>>;
}

export interface RowRef {
  /** 1-based Excel row */
  excelRow: number;
  /** first few cell texts, for lists in the UI */
  preview: string[];
}

export interface PairJoinReport {
  sourceId: string;
  mode: 'oneToOne' | 'lookup';
  plan: Array<{ order: number; name: string }>;
  matched: number;
  perTier: Record<string, number>;
  unmatchedBase: RowRef[];
  unmatchedOther: RowRef[];
  duplicatesBase: Array<{ key: string; count: number }>;
  duplicatesOther: Array<{ key: string; count: number }>;
  ambiguous: Array<{ base: RowRef; candidates: RowRef[]; tier: number }>;
  conflicts: Array<{ base: RowRef; other: RowRef; message: string }>;
  flags: Record<string, { level: string; count: number; examples: string[] }>;
  suggestions: Array<{ base: RowRef; other: RowRef; score: number; reasons: string[] }>;
}

export interface ImportColumnSpec {
  sourceId: string;
  /** column in the source data sheet */
  col: number;
  /** override of the header text written into the merged sheet */
  header?: string;
}

export interface JoinResult {
  /** base rows in base order, then appended rows (join 'full') */
  rows: MergedRowPlan[];
  reports: Record<string, PairJoinReport>;
  summary: { baseRows: number; matchedAny: number; matchedAll: number; appended: number };
  items: CheckItem[];
  /** for appended rows: which base column receives which source column (key columns) */
  keyFill: Record<string, KeyFillEntry[]>;
  /** default columns to import: every non-key column of every other source, in source order */
  defaultImports: ImportColumnSpec[];
  /**
   * per source: the aligned (base column, source column) pairs of EVERY enabled key tier (ID, 氏名, 氏名＋大学 …), not only
   * those needed for appended rows. A matched row has equal keys on such a pair, so an imported formula that reads a
   * non-imported key column of its source (=A2&"様", VLOOKUP(B2, …)) can read the base's partner column instead
   * (premerge uses it for formula remapping; absent = no substitution).
   */
  keyPairs?: Record<string, Array<{ baseCol: number; sourceCol: number }>>;
}

// ---------------------------------------------------------------------------
// premerge (JS package patcher builds the merged package BEFORE any user edit)
// ---------------------------------------------------------------------------

export interface PremergeOptions {
  /** 'keep-if-possible': imported formulas stay live when all refs can be remapped; else cached value */
  formulaPolicy: 'keep-if-possible' | 'values';
  /** base ListObject / AutoFilter grows over the imported columns and appended rows (default true) */
  extendTable: boolean;
  /** 'suffix': a header equal to an existing one becomes 「見出し（ファイル名）」 (default) */
  headerConflict: 'suffix' | 'keep';
  /**
   * Output ONLY the merged sheet (the base data sheet): every other sheet of the base workbook — with its charts,
   * pivot tables, comments, drawings, tables — is removed. Nothing may be left dangling, no #REF! is created:
   *  - formulas of the merged sheet (and of imported columns) that reference a removed sheet, directly or through a
   *    defined name, are replaced by their cached values (arrays / spills as a whole);
   *  - defined names that belong to / point at a removed sheet are removed (localSheetId of the others re-indexed);
   *  - data validation lists that read a removed sheet become literal lists when they fit Excel's 255 characters
   *    (items without commas), else the validation is dropped; CF rules whose formulas reference a removed sheet are
   *    dropped; hyperlinks to a removed sheet lose their link; charts on the merged sheet that read a removed sheet are
   *    removed (self-contained charts / images / shapes stay);
   *  - workbook-level parts only needed by removed sheets (pivot caches, slicer caches, calcChain) are removed and
   *    docProps/app.xml / bookViews / definedNames are made consistent.
   * Default false at the API level (the app's UI passes true unless the user ticks 「他のシートも出力に含める」).
   */
  dropOtherSheets?: boolean;
}

export interface PremergeInput {
  base: KeySource;
  sources: KeySource[];
  rows: MergedRowPlan[];
  imports: ImportColumnSpec[];
  keyFill: Record<string, KeyFillEntry[]>;
  /** JoinResult.keyPairs (see there); premerge substitutes refs to non-imported key columns when the texts agree on every matched row */
  keyPairs?: Record<string, Array<{ baseCol: number; sourceCol: number }>>;
  options: PremergeOptions;
}

/**
 * For appended rows: which base column receives which source value.
 * sourceCols (+ joiner, default '　' full-width space) combines several source columns into one base cell
 * (base 氏名 ← source 姓 + 名). splitOf splits one source cell by whitespace (base 姓 ← part 0, base 名 ← part 1).
 * Exactly one of sourceCol / sourceCols must be set.
 */
export interface KeyFillEntry {
  baseCol: number;
  sourceCol?: number;
  sourceCols?: number[];
  joiner?: string;
  /** take part `part` (0-based) of the source text split on whitespace */
  splitPart?: number;
}

export interface PremergeLayout {
  /** merged sheet name (= base data sheet) */
  sheet: string;
  /** first imported column (all imported columns are contiguous at first) */
  insertAt: number;
  imported: Array<{ col: number; sourceId: string; sourceCol: number; header: string }>;
  /** first appended row, or null */
  appendAt: number | null;
  appended: Array<{ row: number; sourceId: string; sourceRow: number }>;
}

export interface PremergeResult {
  bytes: Uint8Array;
  layout: PremergeLayout;
  report: MergeReport;
}

// ---------------------------------------------------------------------------
// doc (editor)
// ---------------------------------------------------------------------------

export type OpResult =
  | { ok: true; snapshot: EditorSnapshot; warnings: string[] }
  /**
   * fix: the refusal can be resolved automatically. 'unmergeTitles': the only obstacles are merged areas lying
   * entirely ABOVE the header row (titles); calling the same op again with { unmergeTitles: true } logs
   * {kind:'unmerge', refs, centerAcross:true} first (same look, 「選択範囲内で中央」) and then performs the op,
   * all in ONE snapshot (one undo step). The UI asks the user before doing that.
   */
  | { ok: false; reason: string; fix?: { kind: 'unmergeTitles'; refs: string[]; message: string } };

export interface DeleteImpact {
  /** formulas / names / rules that would become #REF! (examples, UI form) */
  breaks: Array<{ where: string; text: string }>;
  breakCount: number;
  /** VLOOKUP/INDEX/OFFSET column-index style warnings (Excel does not adjust those) */
  lookupWarnings: string[];
}

export interface History {
  past: EditorSnapshot[];
  present: EditorSnapshot;
  future: EditorSnapshot[];
}

// ---------------------------------------------------------------------------
// export
// ---------------------------------------------------------------------------

export type ExportEngine = 'excel' | 'js';

/** Job for server/excel/excel-runner.ps1. The op and setCells encodings are documented in server/excel/README.md. */
export interface RunnerJob {
  version: 1;
  /** filled in by the server (temp ASCII paths) */
  input?: string;
  output?: string;
  format: 'xlsx' | 'xlsm';
  ops: Array<Record<string, unknown>>;
  options: { recalc: boolean; clipboard: 'restoreEachOp' | 'restoreAtEnd' };
}

/** value written by a runner setCells op (logical form; logToRunnerJob sends the compact encoding of the README) */
export interface RunnerCell {
  t: CellType;
  v: string | number | boolean | null;
  /** UI-form formula WITH leading '=' (for Range.Formula / Formula2) */
  f?: string;
  dyn?: boolean;
  numFmt?: string;
}

/** Result of an Excel本体 export: the bytes plus runner warnings (MACROS_DROPPED, REF_ERRORS, ...) in Japanese. */
export interface ExcelJobResult {
  bytes: Uint8Array;
  warnings: string[];
}

export type { LoggedOp, SourceBook, TableRegion, MergeReport, EditorSnapshot };
