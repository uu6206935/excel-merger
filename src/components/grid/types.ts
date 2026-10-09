/**
 * Grid view types (owner: grid G). See the GridProps documentation at the top of Grid.tsx.
 * Display styles use the contract XStyle (src/core/types.ts); it is re-exported here for convenience.
 */
import type { CellType, XBorderSide, XBorderStyle, XStyle } from '../../core/types';

export type { XBorderSide, XBorderStyle, XStyle };

/** Band colour / label of one input file (built from MergedDoc.sourceInfo). */
export interface SourceInfo {
  /** file name shown in the header band and its tooltip (e.g. '学生名簿.xlsx（基準）') */
  name: string;
  /** '#RRGGBB' band colour (white text is drawn on it) */
  color: string;
}

export interface GridColumn {
  /** stable id; never changes when the column moves (DocColumn.id) */
  id: string;
  /** header text (the sheet's header-row cell text) */
  title: string;
  /** width in CSS px (use excelWidthToPx from core/display/styles for Excel widths) */
  width: number;
  /** index into GridProps.sources; -1 / undefined = column added in the app (appSourceColor band) */
  source?: number;
  /** join-key role ('ID' / '氏名' / '氏名＋大学' …) -> key badge in the header */
  keyRole?: string;
  /** Excel hidden column: drawn hatched (give it a small width, e.g. 8 px) so it stays attached to its neighbours */
  hidden?: boolean;
}

/**
 * What the grid needs to draw one cell. Build it from the contract `Cell` with cellViewOf() (cellview.ts).
 *  - value: the value or the CACHED result of a formula; null = empty / no cached result.
 *  - formula: the formula as the user sees it (UI form WITH '=', e.g. from doc.formulaForDisplay).
 *    A formula cell whose value is null is drawn as grey italic formula text (never calculated yet).
 *  - t: the contract cell type when known ('e' errors, 'd' dates, 's' text that must not be parsed, 'b').
 */
export interface CellView {
  value: string | number | boolean | null | undefined;
  formula?: string;
  /** xf index -> GridProps.cssOf(styleId) */
  styleId?: number;
  t?: CellType;
}

/** Active cell: `row` is the EXCEL row number (1-based, one of GridProps.rows), `col` the column id. */
export interface CellPos {
  row: number;
  col: string;
}

/** Imperative API (pass `ref` to <Grid>). */
export interface GridHandle {
  /** scroll an Excel row into view (centred); `select` also makes it the active row (active column kept). false if the row is not shown. */
  scrollToRow(row: number, opts?: { select?: boolean }): boolean;
  /** scroll a column into view; `select` also selects that column (header selection) and makes it the active column. false if unknown. */
  scrollToColumn(id: string, opts?: { select?: boolean }): boolean;
  /** scroll a cell into view and make it the active cell. false if unknown. */
  scrollToCell(pos: CellPos): boolean;
  /**
   * Scroll horizontally so that the columns `ids` are around the CENTRE of the visible area (the first contiguous run of
   * them when they are split; a group wider than the viewport is aligned to the left edge instead) and to the top.
   * Smooth unless the user prefers reduced motion (or opts.behavior says otherwise). `select` also selects all of `ids`
   * (non-contiguous ok, anchor = the first id in `ids`) and makes the first one's column the active column (first data row).
   * Columns inside the frozen pane need no scrolling. false if none of the ids exists.
   */
  scrollToColumnsCentered(ids: readonly string[], opts?: { select?: boolean; behavior?: 'smooth' | 'instant' }): boolean;
  /** give keyboard focus to the grid (its IME-safe input sink) */
  focus(): void;
  /** true while a cell is being edited in place */
  isEditing(): boolean;
  /** commit a running in-cell edit (e.g. before an export); no-op when not editing */
  commitEdit(): void;
}

/** In-cell editing state, reported (onEditingChange) so the formula bar can mirror what is typed in the cell. */
export interface EditingState {
  pos: CellPos;
  text: string;
}
