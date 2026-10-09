/**
 * doc facade: the editable merged document (what the grid shows), user operations, undo/redo.
 * Owner: doc module.
 *
 *  createEditor(premerged: SourceBook, layout: PremergeLayout, sourceInfo: MergedDoc['sourceInfo'],
 *               region: { headerRow: number; firstDataRow: number; lastDataRow: number }): EditorSnapshot
 *    - builds the MergedDoc from the premerged package as READ BACK by readBook (single source of truth),
 *      column ids from the layout ('b<n>' base, 'i:<src>:<col>' imported), log = [].
 *
 *  moveColumns(s: EditorSnapshot, ids: string[], gap: number, opts?: { unmergeTitles?: boolean }): OpResult
 *    - gap = insertion slot 0..n in the CURRENT order. Uses excelRef.planMove -> contiguous block moves, each
 *      applied with Excel semantics to every formula (merged sheet + other sheets), names, rules, merges, tables,
 *      arrays, autofilter, freeze; appended to the log as {kind:'struct'} entries. Refuses like Excel:
 *      splitting merged cells, moving into/out of a table, part of an array formula.
 *    - When the only obstacles are merged areas entirely ABOVE the header row, the refusal carries
 *      fix {kind:'unmergeTitles', refs}; with opts.unmergeTitles the op first logs {kind:'unmerge', centerAcross:true}
 *      (display style: centerContinuous) and then moves — one snapshot, one undo step.
 *
 *  insertColumns(s: EditorSnapshot, index: number, count: number): OpResult
 *    - index = 0-based position in the current order (insert before it). Excel "format from left".
 *
 *  analyzeDelete(s: EditorSnapshot, ids: string[]): DeleteImpact
 *  deleteColumns(s: EditorSnapshot, ids: string[], opts: { convertDependents: boolean }): OpResult
 *    - convertDependents: first write cached values over every formula that would become #REF!
 *      (logged as {kind:'cells', reason:'toValues'}; an array/spill master is converted as a whole), then delete
 *      (right-to-left contiguous runs).
 *
 *  setCellInput(s: EditorSnapshot, colId: string, row: number, text: string): OpResult
 *    - Excel-like parsing: '=...' formula (UI form -> file form, adds _xlfn.; `dyn` from isDynamicArrayFormula),
 *      number, date, %, currency, TRUE/FALSE, text. Like Excel, a date / % / currency typed into a General cell sets
 *      the cell's number format (CellWrite numFmt: 'yyyy/m/d', '0%', '"¥"#,##0' …). Keeps text for text-formatted /
 *      ID-like columns ('00123'). Refuses cells inside an array range.
 *
 *  setColumnWidth(s: EditorSnapshot, colId: string, width: number): OpResult
 *    - width in the STORED <col width> unit (same as ColProps.width), not the character count Excel's UI shows.
 *  formulaForDisplay(doc: MergedDoc, colIndex: number, row: number): string | null   // UI form with '='
 *
 *  initHistory(s) / commit(h, s) / undo(h) / redo(h) / canUndo(h) / canRedo(h)
 */
export {
  createEditor,
  moveColumns,
  insertColumns,
  analyzeDelete,
  deleteColumns,
  setCellInput,
  setColumnWidth,
  formulaForDisplay,
} from './ops';
export { initHistory, commit, undo, redo, canUndo, canRedo } from './history';
