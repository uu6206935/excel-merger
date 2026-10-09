/**
 * xlsx facade: package reader + package patcher (premerge / replay).
 * Owner: reader (read.ts, styles.ts, cellText.ts, legacy.ts) and patcher (premerge.ts, replay.ts, patch/*).
 * Everything here runs in the browser (main thread or worker) AND in Node (vitest) — no DOM APIs.
 */
import type { Cell, SourceBook } from '../types';
import type { LoggedOp } from '../types';
import type { PremergeInput, PremergeResult, ProgressFn } from '../contracts';

export { readBook } from './read';
export { cellText } from './cellText';
export { buildPremerged } from './premerge';
export { replayLog } from './replay';

/** Signatures (implemented in the files above):
 *
 *  readBook(bytes: Uint8Array, fileName: string, opts?: { id?: string; onProgress?: ProgressFn }): Promise<SourceBook>
 *    - .xlsx/.xlsm: own OOXML reader (JSZip + string scanning; NO DOMParser). Shared formulas expanded with
 *      excelRef.translateFormula; array/dynamic masters get `arr`/`dyn`; cached values typed; dates -> 'd' when
 *      the xf number format is a date format; styles -> XStyle[] (theme + tint resolved).
 *    - .xls/.ods/.csv: SheetJS (+ encoding-japanese for CSV, raw:true so 00123 stays text); fidelity 'values';
 *      `bytes` replaced by a generated minimal .xlsx package holding the values.
 *
 *  cellText(book: SourceBook, cell: Cell | undefined): string
 *    - the text Excel would display (number format applied: 00123, 2026/10/20, ¥1,234).
 *
 *  buildPremerged(input: PremergeInput, onProgress?: ProgressFn): Promise<PremergeResult>
 *    - takes the base package bytes and produces the merged package: appended rows inserted (Excel insert-row
 *      semantics + range extension), imported columns inserted after the base table (clean insert), imported cells
 *      written with their source styles merged into the base styles.xml, header cells, widths.
 *
 *  replayLog(premerged: Uint8Array, log: LoggedOp[], onProgress?: ProgressFn): Promise<Uint8Array>
 *    - replays the user's op log on the package with Excel semantics (every part: all sheets' formulas, names,
 *      CF/DV incl. x14, merges, tables, autofilter, drawings anchors, charts <c:f>, pivot sources, comments,
 *      hyperlinks, cols, panes ...), deletes calcChain, sets fullCalcOnLoad.
 */
export type _Sigs = [Cell, SourceBook, LoggedOp, PremergeInput, PremergeResult, ProgressFn];
