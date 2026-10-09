/**
 * Core shared types — THE CONTRACT between modules. See docs/ARCHITECTURE.md.
 * Change only with care: loader, keys, premerge/patcher, doc (editor), exporters and UI all depend on it.
 *
 * Conventions
 *  - Rows and columns are 1-based everywhere (column 1 = "A"), like Excel.
 *  - Formula text is stored WITHOUT the leading "=", exactly in xlsx FILE form:
 *    English function names, "_xlfn." / "_xlws." / "_xlpm." prefixes kept, spill refs as
 *    _xlfn.ANCHORARRAY(X), structured refs as Table1[[#This Row],[col]]. See formula/forms.ts
 *    for conversion to/from what the user sees/types (UI form).
 *  - Dates are numbers (Excel serial, 1900 system; a 1904 book is converted on read) with t = 'd'
 *    when the cell's number format is a date/time format. Display formatting uses the numFmt code.
 *  - Everything in a MergedDoc / EditorSnapshot is immutable: operations return new objects and share
 *    unchanged parts (undo/redo keeps whole snapshots).
 */

import type { StructOp } from './formula/excelRef';

export type { StructOp };

// ---------------------------------------------------------------------------
// Coordinates
// ---------------------------------------------------------------------------

export interface CellPos {
  row: number;
  col: number;
}

/** Inclusive rectangle, normalized so r1 <= r2 and c1 <= c2. Whole columns: r1 = 1, r2 = 1048576. */
export interface RangeBox {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

// ---------------------------------------------------------------------------
// Cells
// ---------------------------------------------------------------------------

export type CellType =
  | 's' // string
  | 'n' // number
  | 'b' // boolean
  | 'd' // date/time (v = Excel serial number)
  | 'e' // error (v = '#N/A', '#REF!', ...)
  | 'z'; // empty / no cached result

export interface Cell {
  t: CellType;
  /** value, or the cached result of a formula */
  v: string | number | boolean | null;
  /** formula text without '=' (FILE form). t/v hold the cached result. Shared formulas are already expanded. */
  f?: string;
  /** array / dynamic-array formula master only: the A1 range it covers ("C2:C10") */
  arr?: string;
  /** dynamic-array (spill) formula (cm="1" in the file) */
  dyn?: boolean;
  /** xf index into the owning book's StyleSheet.xfs (cellXfs). undefined = 0 */
  s?: number;
  /** rich text runs (v holds the plain text) — display only */
  rich?: Array<{ text: string; style?: XStyle }>;
}

// ---------------------------------------------------------------------------
// Styles (display side). The raw XML side lives in xlsx/styles.ts (StyleSheet).
// ---------------------------------------------------------------------------

export type XBorderStyle =
  | 'thin'
  | 'medium'
  | 'thick'
  | 'dashed'
  | 'dotted'
  | 'double'
  | 'hair'
  | 'mediumDashed'
  | 'dashDot'
  | 'mediumDashDot'
  | 'dashDotDot'
  | 'mediumDashDotDot'
  | 'slantDashDot';

export interface XBorderSide {
  style: XBorderStyle;
  /** '#RRGGBB' */
  color?: string;
}

/** Flattened display style; colours already resolved (theme + tint, indexed) to '#RRGGBB'. */
export interface XStyle {
  fill?: string;
  color?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  /** points */
  size?: number;
  font?: string;
  hAlign?: 'general' | 'left' | 'center' | 'right' | 'fill' | 'justify' | 'centerContinuous' | 'distributed';
  vAlign?: 'top' | 'middle' | 'bottom';
  wrap?: boolean;
  indent?: number;
  border?: { top?: XBorderSide; right?: XBorderSide; bottom?: XBorderSide; left?: XBorderSide };
  /** number format CODE (built-in ids already resolved to their ja-JP code, e.g. 14 -> 'yyyy/m/d') */
  numFmt?: string;
}

// ---------------------------------------------------------------------------
// Sheets and books (reader output)
// ---------------------------------------------------------------------------

export interface ColProps {
  /** Excel character-width units */
  width?: number;
  hidden?: boolean;
  outlineLevel?: number;
  /** column default xf */
  s?: number;
  /**
   * The width was set by the user: <col customWidth="1"> WITHOUT bestFit. An auto width (no customWidth, or an
   * auto-fitted column: bestFit="1") is widened by Excel when a wider value (e.g. a typed date) is entered; the doc
   * and the JS patcher mirror that. Setting a width explicitly (colWidth op) ends auto-fit (bestFit is dropped).
   */
  customWidth?: boolean;
}

export interface RowProps {
  /** points */
  height?: number;
  hidden?: boolean;
  outlineLevel?: number;
  s?: number;
}

/** Formula-bearing sheet features that must follow structural ops (text kept in FILE form). */
export interface FeatureRule {
  /** space separated A1 ranges ("A2:A100 C2:C100") */
  sqref: string;
  /**
   * CF: the <cfRule type> ('cellIs', 'expression', 'colorScale', 'dataBar', 'iconSet', 'top10', 'uniqueValues',
   * 'duplicateValues', 'aboveAverage', ...); DV: the <dataValidation type> ('list', 'whole', ...). Needed for the
   * move "leak" semantics (colorScale/dataBar/iconSet/top10/uniqueValues/duplicateValues/aboveAverage do not spread).
   */
  type?: string;
  /** formulas (relative refs anchored at the top-left cell of the sqref bounding box), no '=' */
  formulas: string[];
  /** 'cf' conditional formatting, 'dv' data validation; `ext` = stored in x14 extLst */
  kind: 'cf' | 'dv';
  ext?: boolean;
}

export interface TableInfo {
  name: string;
  /** A1 range incl. header (and totals) row */
  ref: string;
  headerRow: boolean;
  totalsRow: boolean;
  columns: string[];
}

export interface SheetFeatures {
  merges: RangeBox[];
  rules: FeatureRule[];
  tables: TableInfo[];
  /** array / spill formula ranges (masters' `arr`) */
  arrays: RangeBox[];
  autoFilter?: string;
  /** frozen columns / rows */
  freeze?: { cols: number; rows: number };
  hyperlinks: Array<{ ref: string; target?: string; location?: string }>;
}

export interface SheetGrid {
  name: string;
  state: 'visible' | 'hidden' | 'veryHidden';
  /** used extents */
  maxRow: number;
  maxCol: number;
  /** rows[r - 1][c - 1]; missing rows / cells = empty */
  rows: Array<Array<Cell | undefined> | undefined>;
  colProps: Array<ColProps | undefined>;
  rowProps: Array<RowProps | undefined>;
  features: SheetFeatures;
  /** default column width of the sheet (sheetFormatPr), Excel char units */
  defaultColWidth?: number;
  defaultRowHeight?: number;
}

export interface DefinedName {
  name: string;
  /** sheet name for sheet-scoped names (localSheetId), null for workbook scope */
  scope: string | null;
  /** formula text without '=' (FILE form, e.g. "'名簿'!$C$2:$C$40") */
  refersTo: string;
  hidden?: boolean;
}

/** Things only the Excel本体 engine can adjust perfectly / that some paths may lose; shown as notes. */
export interface BookFeatureCounts {
  charts: number;
  pivotTables: number;
  vba: boolean;
  drawings: number;
  slicers: number;
  sparklineGroups: number;
  externalLinks: number;
  threadedComments: number;
  formControls: number;
  tables: number;
  images: number;
}

export type SourceFormat = 'xlsx' | 'xlsm' | 'xls' | 'xlsb' | 'csv' | 'ods' | 'other';

export interface SourceBook {
  id: string;
  fileName: string;
  size: number;
  format: SourceFormat;
  /**
   * original file bytes. For 'values' fidelity inputs (xls/csv/ods) this is a freshly generated
   * minimal .xlsx package holding the values, so every later step can treat all books as xlsx packages.
   */
  bytes: Uint8Array;
  /** 'values' when the file could only be read as values (xls/csv/ods via SheetJS) */
  fidelity: 'full' | 'values';
  /** display styles, index = xf index (cellXfs). styles[0] is the default style. */
  styles: XStyle[];
  sheets: SheetGrid[];
  names: DefinedName[];
  features: BookFeatureCounts;
  date1904: boolean;
  warnings: string[];
  /** detected text encoding of a CSV/TXT input ('Shift_JIS', 'UTF-8', ...) */
  encoding?: string;
}

/** The data table inside the chosen sheet (auto-detected, user adjustable). */
export interface TableRegion {
  sheet: string;
  headerRow: number;
  firstDataRow: number;
  lastDataRow: number;
  firstCol: number;
  lastCol: number;
}

// ---------------------------------------------------------------------------
// Join / merge planning (keys module output -> premerge input)
// ---------------------------------------------------------------------------

export type JoinType =
  | 'left' // keep base rows only (default)
  | 'full'; // also append rows that exist only in other files (at the end of the table)

export type DuplicatePolicy = 'first' | 'last' | 'skip';

export type MatchFlag =
  | 'fallback' // matched by a lower-priority key (要確認)
  | 'conflict' // matched by a key but another key disagrees (ID一致・氏名不一致)
  | 'duplicateInSource' // several source rows had this key; policy picked one
  | 'duplicateInBase' // several base rows share this key (lookup mode)
  | 'variant' // matched only after 異体字/カナ/全半角 folding (info)
  | 'manual'; // linked by the user / accepted suggestion

export interface RowMatch {
  /** sheet row (1-based Excel row) in the source data sheet */
  row: number;
  /** label of the key (tier) that produced the match, e.g. 'ID', '氏名＋大学' */
  via: string;
  flags: MatchFlag[];
}

export interface MergedRowPlan {
  /** sheet row in the base data sheet; null for rows appended from other files (join 'full') */
  baseRow: number | null;
  /** for appended rows: which source row created it */
  appendedFrom?: { sourceId: string; row: number };
  /** per NON-base source */
  matches: Record<string, RowMatch | null>;
}

// ---------------------------------------------------------------------------
// Merged document (the editable result shown in the grid)
// ---------------------------------------------------------------------------

export type ColumnOrigin =
  | { kind: 'base'; col: number } // column number in the ORIGINAL base sheet
  | { kind: 'import'; sourceId: string; col: number } // column number in the source data sheet
  | { kind: 'new' }; // inserted by the user in the editor

export interface DocColumn {
  /** stable id, survives moves: 'b3' (base col 3), 'i:s2:5' (source s2 col 5), 'n1' (user added) */
  id: string;
  origin: ColumnOrigin;
  props: ColProps;
  /** cells[row - 1]; length = doc.rowCount */
  cells: Array<Cell | undefined>;
}

/** A formula that lives outside the merged sheet (other sheet of the output book) and may reference it. */
export interface ExternalFormulaCell {
  sheet: string;
  row: number;
  col: number;
  f: string;
}

export interface MergedDoc {
  /** the merged sheet = the base data sheet (same name) */
  sheetName: string;
  /** all sheet names of the output book, in order (for formula context) */
  sheetNames: string[];
  rowCount: number;
  headerRow: number;
  firstDataRow: number;
  lastDataRow: number;
  /** current column order: columns[0] is column A of the merged sheet */
  columns: DocColumn[];
  rowProps: Array<RowProps | undefined>;
  /** display styles of the OUTPUT book (premerged styles.xml), index = xf */
  styles: XStyle[];
  /** merged sheet features in CURRENT coordinates */
  features: SheetFeatures;
  /** formula cells on the other sheets of the output book (current text) */
  externalFormulas: ExternalFormulaCell[];
  /** CF / DV rules of other sheets (for dependents checks; adjusted like formulas) */
  externalRules: Array<{ sheet: string; rule: FeatureRule }>;
  names: DefinedName[];
  /** per source: UI colour / label (header band in the grid) */
  sourceInfo: Record<string, { label: string; color: string }>;
  /** increments on every change (cheap memo key) */
  version: number;
}

// ---------------------------------------------------------------------------
// Operation log (shared by the JS package patcher and the Excel COM replay)
// ---------------------------------------------------------------------------

/**
 * A positional cell write (positions as of the time of the write). The cell's style is kept, except that
 * `numFmt` (when given) changes only the number format of the cell (Excel does this when a date / 50% / ¥1,000
 * is typed into a General cell): the JS patcher reuses or creates an xf = the cell's xf with that numFmt, the
 * Excel runner sets Range.NumberFormat.
 */
export interface CellWrite {
  row: number;
  col: number;
  /**
   * null clears the cell's value/formula (style kept).
   * dyn: the formula must be stored as a dynamic-array (spilling) formula — JS patcher writes
   * <c cm="1"><f t="array" ref="<cell>">…</f></c> + the dynamic-array metadata part; Excel runner uses Formula2.
   * Set by the doc from formula/dynamic.ts isDynamicArrayFormula for formulas the user typed.
   * A write with reason 'toValues' over an array/spill MASTER cell replaces the whole array by its values.
   * numFmt: applied AFTER the value is written (a format applied after the value widens an auto-width column the
   * way Excel does); '@' (text) is applied BEFORE the value. 'General' resets the format.
   * quote: store a TEXT constant with quotePrefix (the apostrophe Excel shows in the formula bar) because it would
   * otherwise be re-parsed as a number / date / formula / boolean when the user later edits the cell. The doc sets it
   * for strings that look like such values and for text typed with a leading apostrophe; both engines honour it
   * exactly: quote === true → quotePrefix="1" / Excel "'" prefix; otherwise NO prefix is added. (A prefix that the
   * cell already had is invisible for numbers/empty cells; Excel COM cannot clear it, the JS engine does — the two
   * engines may differ in that one invisible flag.) Ignored for numbers, booleans and formulas.
   * A text constant containing a line feed (LF or CR LF, not a lone CR) turns the cell's WrapText on (Excel does it
   * through COM Value2 too); both engines and the doc mirror it.
   * Totals row of an Excel table: a constant written into a totals cell becomes the column's totalsRowLabel (the
   * doc already converts it to t='s' text); a formula SUBTOTAL(n,Tbl[own column]) sets the totals function.
   */
  cell: { t: CellType; v: string | number | boolean | null; f?: string; dyn?: boolean; numFmt?: string; quote?: boolean } | null;
}

export type LoggedOp =
  /** excelRef StructOp (1-based; move.to = insert-before index in ORIGINAL coordinates) on `sheet` */
  | { kind: 'struct'; sheet: string; op: StructOp }
  | { kind: 'cells'; sheet: string; writes: CellWrite[]; reason: 'edit' | 'toValues' }
  /**
   * width = the value STORED in <col width> (same unit as ColProps.width; includes Excel's ~5 px padding).
   * The Excel runner converts: Range.ColumnWidth = stored width - padding (as the golden runner does).
   * The column becomes user-set (customWidth, no bestFit): later wider values do not widen it.
   */
  | { kind: 'colWidth'; sheet: string; col: number; width: number }
  /**
   * Unmerge merged areas (A1 refs) of `sheet`. centerAcross: give every cell of each former area the horizontal
   * alignment 'centerContinuous' (Excel「選択範囲内で中央」) so a merged title looks the same without blocking
   * column moves. JS patcher: remove <mergeCell>, derive xfs; Excel runner: Range.UnMerge + HorizontalAlignment = 7.
   * ALWAYS (independent of centerAcross): the other cells of each former area receive the top-left cell's FILL and the
   * border sides of the PERIMETER they lie on — top side only on the first row, bottom only on the last row, left only
   * on the first column, right only on the last column — and only sides the top-left cell HAS (nothing is erased; a
   * cell keeps its own font / number format / alignment). For a one-row title band that means top+bottom on every
   * cell, left on the first, right on the last. So a title band whose style was stored on the top-left cell only
   * (openpyxl-style files) stays visible. Excel's own Range.UnMerge spreads nothing: the Excel runner UnMerges first,
   * reads the top-left cell (it keeps its look) and then writes the others per edge.
   */
  | { kind: 'unmerge'; sheet: string; refs: string[]; centerAcross: boolean };

export interface EditorSnapshot {
  doc: MergedDoc;
  /** user ops since the merge (the premerged package already contains the merge itself) */
  log: LoggedOp[];
  /** human readable label of the command that produced this snapshot ('列を移動' ...) */
  label: string;
}

// ---------------------------------------------------------------------------
// Checks / reports
// ---------------------------------------------------------------------------

export type Severity = 'error' | 'warning' | 'info';

export interface CheckItem {
  severity: Severity;
  code: string;
  message: string;
  sourceId?: string;
  /** example rows of the BASE sheet (1-based Excel rows) for jumping to them */
  rows?: number[];
  /** example rows of the OTHER file named by sourceId (1-based Excel rows), shown in messages / lists */
  otherRows?: number[];
}

export interface MergeReport {
  mergedRows: number;
  matchedRows: number;
  baseOnlyRows: number;
  appendedRows: number;
  importedColumns: number;
  formulasKept: number;
  formulasToValues: number;
  /** names of the sheets removed by PremergeOptions.dropOtherSheets (undefined / empty = none) */
  droppedSheets?: string[];
  /** formulas of the merged sheet turned into values because they read a dropped sheet (dropOtherSheets) */
  formulasDetached?: number;
  /**
   * Imported formulas, per source file: how many stayed live formulas (remapped to the merged sheet), how many became values and
   * why, and how many ended up EMPTY (the source stored no result and the formula could not be kept). reasons: counts per code —
   * 'otherSheet' (reads another sheet of the source), 'externalCell' (a cell outside the data table, e.g. a rate in a side cell),
   * 'rowRelation' (reads other rows whose merged positions are not contiguous / not in the same order), 'unresolvedName',
   * 'unresolvedTable', 'positional' (ROW()/OFFSET/INDIRECT …), 'notImported' (reads a source column that is not imported and is no
   * key), 'unsafe', 'array', 'unsupported'.
   */
  formulaStats?: Array<{ sourceId: string; label: string; kept: number; toValues: number; empty: number; reasons: Record<string, number> }>;
  items: CheckItem[];
}
