/**
 * Grid module entry (owner: grid G). See the documentation at the top of Grid.tsx for GridProps / GridHandle.
 *
 * Typical App wiring (MergedDoc -> grid):
 *   const cssOf   = useMemo(() => makeStyleCache(doc.styles), [doc.styles]);          // core/display/styles
 *   const rows    = useMemo(() => range(doc.firstDataRow, doc.lastDataRow), [doc.firstDataRow, doc.lastDataRow]);
 *   const getCell = useCallback((row, col) => {                                       // row = EXCEL row
 *     const dc = colById.get(col.id); const cell = dc?.cells[row - 1];
 *     return cellViewOf(cell, cell?.f ? formulaForDisplay(doc, colIndex.get(col.id)!, row) : null, dc?.props.s);
 *   }, [doc]);
 *   <FormulaBar pos={active} address={cellAddress(activeIdx, active.row)} value={cellEditText(view, cssOf(view?.styleId))}
 *               mirror={editing?.text ?? null} onCommitEdit={commitEdit} onDone={() => gridRef.current?.focus()} />
 *   <Grid ref={gridRef} columns={…} sources={…} rows={rows} headerRow={doc.headerRow} frozenCols={doc.features.freeze?.cols}
 *         getCell={getCell} cssOf={cssOf} dataVersion={doc.version} … callbacks … onEditingChange={setEditing} />
 */
export { Grid, ROW_H, ROWHDR_W, HEADER_H, MIN_COL_W, type GridProps } from './Grid';
export { FormulaBar, type FormulaBarProps } from './FormulaBar';
export { ContextMenu, type ContextMenuProps, type MenuEntry, type MenuItem } from './ContextMenu';
export { cellViewOf, cellEditText, cellCopyText, renderCell, lacksCachedValue, isBlankCell, type CellRender } from './cellview';
export { computeMove, isNoopMove, planBlockMoves, applyBlockMove, colLetter, letterToIndex, cellAddress, rangeBetween, toRuns, type BlockMove } from './colops';
export { clickSelect, orderedSelection, selectOnly, extendTo, sameSelection, EMPTY_SEL, type ColSelection } from './selection';
export { insertTarget, makeRowIndex } from './geometry';
export type { CellPos, CellView, EditingState, GridColumn, GridHandle, SourceInfo, XStyle } from './types';
