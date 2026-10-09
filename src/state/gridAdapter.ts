/**
 * MergedDoc -> grid view model (pure). The grid shows the DATA rows of the merged sheet with their Excel row
 * numbers; the header row becomes the column titles. Types come from the grid module (owner G).
 */
import type { CellView, GridColumn, SourceInfo } from '../components/grid/types';
import { cellViewOf } from '../components/grid/cellview';
import type { JoinPlan, KeyAnalysis } from '../core/contracts';
import type { Cell, DocColumn, MergedDoc } from '../core/types';
import { defaultColumnPx, excelWidthToPx, mdwFor } from '../core/display/styles';
import { colLetter } from './helpers';
import { findCandidate, pairIds, shortLabel } from './keysPlan';

export interface GridModel {
  columns: GridColumn[];
  /** band colour / name per source, base first (GridColumn.source indexes into it) */
  sources: SourceInfo[];
  /** source id per entry of `sources` */
  sourceIds: string[];
  /** Excel rows shown by the grid (the data rows, contiguous) */
  rows: number[];
  rowCount: number;
  firstRow: number;
  frozenCols: number;
  /** column id -> current 0-based index */
  colIndex: Map<string, number>;
  colById: Map<string, DocColumn>;
  /** max digit width of the book's default font (px), for width conversions */
  mdw: number;
}

/** Base column (1-based, original base sheet) -> key badge text, from the plan's tiers. */
export function keyRolesFromPlan(analysis: KeyAnalysis | null, plan: JoinPlan | null): Map<number, string> {
  const roles = new Map<number, string>();
  if (!analysis || !plan) return roles;
  for (const s of pairIds(analysis)) {
    for (const id of plan.tiers[s] ?? []) {
      const c = findCandidate(analysis, s, id);
      if (!c) continue;
      const label = c.family === 'id' ? 'ID' : c.family === 'name' && c.kind !== 'composite' ? '氏名' : shortLabel(c);
      for (const bc of c.baseColumns) if (!roles.has(bc.col)) roles.set(bc.col, label);
    }
  }
  return roles;
}

export function dataRowCount(doc: MergedDoc): number {
  return Math.max(0, doc.lastDataRow - doc.firstDataRow + 1);
}

/** Max digit width of the workbook default font (styles[0]). */
export function docMdw(doc: MergedDoc): number {
  const s0 = doc.styles[0];
  return mdwFor(s0?.font, s0?.size ?? 11);
}

export function buildGridModel(
  doc: MergedDoc,
  opts: {
    baseId: string | null;
    keyRoles: Map<number, string>;
    headerText: (cell: Cell | undefined) => string;
    liveWidths?: Record<string, number>;
    /** sheet default column width (Excel units), when the file has one */
    defaultColWidth?: number;
  },
): GridModel {
  const mdw = docMdw(doc);
  const ids = Object.keys(doc.sourceInfo);
  const sourceIds = opts.baseId && ids.includes(opts.baseId) ? [opts.baseId, ...ids.filter((i) => i !== opts.baseId)] : ids;
  const sources: SourceInfo[] = sourceIds.map((id) => {
    const info = doc.sourceInfo[id];
    return { name: id === opts.baseId ? `${info.label}（基準）` : info.label, color: info.color };
  });
  const sourceIndex = new Map(sourceIds.map((id, i) => [id, i]));
  const baseIdx = opts.baseId ? (sourceIndex.get(opts.baseId) ?? -1) : -1;
  const defPx = opts.defaultColWidth && opts.defaultColWidth > 0 ? excelWidthToPx(opts.defaultColWidth, mdw) : defaultColumnPx(mdw);
  const colIndex = new Map<string, number>();
  const colById = new Map<string, DocColumn>();
  const live = opts.liveWidths ?? {};
  const columns: GridColumn[] = doc.columns.map((c, i) => {
    colIndex.set(c.id, i);
    colById.set(c.id, c);
    let title = '';
    try {
      title = opts.headerText(c.cells[doc.headerRow - 1]).replace(/\s+/g, ' ').trim();
    } catch {
      title = '';
    }
    let source = -1;
    let keyRole: string | undefined;
    if (c.origin.kind === 'base') {
      source = baseIdx;
      keyRole = opts.keyRoles.get(c.origin.col);
    } else if (c.origin.kind === 'import') {
      source = sourceIndex.get(c.origin.sourceId) ?? -1;
    }
    const stored = c.props.width;
    let px = live[c.id] ?? (stored !== undefined && stored > 0 ? excelWidthToPx(stored, mdw) : defPx);
    if (c.props.hidden && live[c.id] === undefined) px = 8;
    const col: GridColumn = { id: c.id, title: title || `（${colLetter(i + 1)}列）`, width: Math.max(8, px), source };
    if (keyRole) col.keyRole = keyRole;
    if (c.props.hidden) col.hidden = true;
    return col;
  });
  const rowCount = dataRowCount(doc);
  const rows = new Array<number>(rowCount);
  for (let i = 0; i < rowCount; i++) rows[i] = doc.firstDataRow + i;
  return {
    columns,
    sources,
    sourceIds,
    rows,
    rowCount,
    firstRow: doc.firstDataRow,
    frozenCols: Math.max(0, Math.min(doc.features.freeze?.cols ?? 0, columns.length)),
    colIndex,
    colById,
    mdw,
  };
}

/**
 * getCell for the grid: (Excel row, column) -> CellView, O(1). Views are cached per Cell object (cells are
 * immutable and shared between snapshots), so scrolling does not allocate. Empty cells of a column / row with
 * a default style get that style.
 * @param formulaText formulaForDisplay(doc, colIndex, row): UI form with '=' (null = use '=' + file form)
 */
export function makeGetCell(
  doc: MergedDoc,
  model: GridModel,
  formulaText: (colIndex: number, row: number) => string | null,
): (row: number, col: GridColumn) => CellView | undefined {
  const cache = new WeakMap<Cell, CellView | undefined>();
  const emptyByStyle = new Map<number, CellView | undefined>();
  return (row: number, col: GridColumn) => {
    const dc = model.colById.get(col.id);
    if (!dc) return undefined;
    const cell = dc.cells[row - 1];
    if (!cell) {
      const s = doc.rowProps[row - 1]?.s ?? dc.props.s;
      if (s === undefined || s === 0) return undefined;
      if (!emptyByStyle.has(s)) emptyByStyle.set(s, cellViewOf(undefined, null, s));
      return emptyByStyle.get(s);
    }
    if (cache.has(cell)) return cache.get(cell);
    let ftext: string | null = null;
    if (cell.f) {
      try {
        ftext = formulaText(model.colIndex.get(col.id) ?? 0, row);
      } catch {
        ftext = null;
      }
    }
    const view = cellViewOf(cell, ftext, dc.props.s);
    cache.set(cell, view);
    return view;
  };
}

/** 'C12' for the address box. */
export function cellAddress(colIdx: number, excelRow: number): string {
  return `${colLetter(colIdx + 1)}${excelRow}`;
}

/** true when the Excel row is one of the data rows shown by the grid. */
export function isDataRow(doc: MergedDoc, excelRow: number): boolean {
  return excelRow >= doc.firstDataRow && excelRow <= doc.lastDataRow;
}

/** Ids of the columns that exist in `after` but not in `before` (e.g. inserted columns), in order. */
export function addedColumnIds(before: MergedDoc, after: MergedDoc): string[] {
  const had = new Set(before.columns.map((c) => c.id));
  return after.columns.filter((c) => !had.has(c.id)).map((c) => c.id);
}

/** Column list entries for the side panel. */
export interface ColumnListItem {
  id: string;
  index: number;
  letter: string;
  title: string;
  color: string;
  sourceLabel: string;
  keyRole?: string;
  hidden: boolean;
}

export function columnListItems(model: GridModel, appColor: string): ColumnListItem[] {
  return model.columns.map((c, i) => {
    const src = c.source !== undefined && c.source >= 0 ? model.sources[c.source] : undefined;
    return {
      id: c.id,
      index: i,
      letter: colLetter(i + 1),
      title: c.title,
      color: src?.color ?? appColor,
      sourceLabel: src?.name ?? 'アプリで追加',
      keyRole: c.keyRole,
      hidden: !!c.hidden,
    };
  });
}
