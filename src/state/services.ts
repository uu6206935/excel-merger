/**
 * Everything the UI calls in other modules, behind one object so that `?mock=1` (dev only) can
 * swap in fake data (./mock.ts). `realServices` wires the documented facades exactly.
 */
import * as api from '../api';
import type { ServerStatus } from '../api';
import type {
  DeleteImpact,
  ExcelJobResult,
  JoinPlan,
  JoinResult,
  KeyAnalysis,
  KeySource,
  NormalizeOptions,
  OpResult,
  PremergeInput,
  PremergeLayout,
  PremergeResult,
  ProgressFn,
  RunnerJob,
} from '../core/contracts';
import type { Cell, EditorSnapshot, LoggedOp, MergedDoc, SourceBook, TableRegion } from '../core/types';
import * as doc from '../core/doc';
import * as exportFacade from '../core/export';
import * as keys from '../core/keys';
import * as xlsx from '../core/xlsx';

export interface Services {
  /** true for the ?mock=1 fake data */
  mock: boolean;
  // xlsx
  readBook(bytes: Uint8Array, fileName: string, opts?: { id?: string; onProgress?: ProgressFn }): Promise<SourceBook>;
  cellText(book: SourceBook, cell: Cell | undefined): string;
  buildPremerged(input: PremergeInput, onProgress?: ProgressFn): Promise<PremergeResult>;
  // keys
  pickDataSheet(book: SourceBook): string;
  /** opts: the header row / last data row the user picked (see keys facade) */
  detectRegion(book: SourceBook, sheet: string, opts?: { headerRow?: number; lastDataRow?: number }): TableRegion;
  analyzeSources(base: KeySource, others: KeySource[], options?: Partial<NormalizeOptions>): KeyAnalysis;
  defaultPlan(analysis: KeyAnalysis): JoinPlan;
  joinSources(analysis: KeyAnalysis, plan: JoinPlan): JoinResult;
  /**
   * 「列を選んで結合キーに追加」: a NEW analysis whose pair for `sourceId` has one more candidate (selected:false) built from
   * sheet columns of the base (`baseCols`) and of that file (`otherCols`), 1..3 each, same count. Throws an Error with a
   * Japanese message for invalid columns.
   */
  addManualCandidate(analysis: KeyAnalysis, sourceId: string, baseCols: number[], otherCols: number[]): KeyAnalysis;
  // doc
  createEditor(
    premerged: SourceBook,
    layout: PremergeLayout,
    sourceInfo: MergedDoc['sourceInfo'],
    region: { headerRow: number; firstDataRow: number; lastDataRow: number },
  ): EditorSnapshot;
  moveColumns(s: EditorSnapshot, ids: string[], gap: number, opts?: { unmergeTitles?: boolean }): OpResult;
  insertColumns(s: EditorSnapshot, index: number, count: number): OpResult;
  analyzeDelete(s: EditorSnapshot, ids: string[]): DeleteImpact;
  deleteColumns(s: EditorSnapshot, ids: string[], opts: { convertDependents: boolean }): OpResult;
  setCellInput(s: EditorSnapshot, colId: string, row: number, text: string): OpResult;
  /** width in the STORED <col width> unit (display/styles pxToExcelWidth), not the character count Excel shows */
  setColumnWidth(s: EditorSnapshot, colId: string, width: number): OpResult;
  formulaForDisplay(d: MergedDoc, colIndex: number, row: number): string | null;
  // export
  exportWithJs(premerged: Uint8Array, log: LoggedOp[], onProgress?: ProgressFn): Promise<Uint8Array>;
  logToRunnerJob(log: LoggedOp[], opts: { format: 'xlsx' | 'xlsm'; sheet: string }): RunnerJob;
  // server
  getServerStatus(): Promise<ServerStatus | null>;
  runExcelJob(premerged: Uint8Array, job: RunnerJob, onProgress?: ProgressFn, signal?: AbortSignal): Promise<ExcelJobResult>;
  saveToOutput(bytes: Uint8Array, fileName: string): Promise<{ path: string; name: string }>;
  openFile(path: string): Promise<void>;
  openOutputFolder(): Promise<void>;
  startHeartbeat(): void;
}

export const realServices: Services = {
  mock: false,
  readBook: xlsx.readBook,
  cellText: xlsx.cellText,
  buildPremerged: xlsx.buildPremerged,
  pickDataSheet: keys.pickDataSheet,
  detectRegion: keys.detectRegion,
  analyzeSources: keys.analyzeSources,
  defaultPlan: keys.defaultPlan,
  joinSources: keys.joinSources,
  addManualCandidate: keys.addManualCandidate,
  createEditor: doc.createEditor,
  moveColumns: doc.moveColumns,
  insertColumns: doc.insertColumns,
  analyzeDelete: doc.analyzeDelete,
  deleteColumns: doc.deleteColumns,
  setCellInput: doc.setCellInput,
  setColumnWidth: doc.setColumnWidth,
  formulaForDisplay: doc.formulaForDisplay,
  exportWithJs: exportFacade.exportWithJs,
  logToRunnerJob: exportFacade.logToRunnerJob,
  getServerStatus: api.getServerStatus,
  runExcelJob: async (premerged, job, onProgress, signal) => normalizeJobResult(await api.runExcelJob(premerged, job, onProgress, signal)),
  saveToOutput: api.saveToOutput,
  openFile: api.openFile,
  openOutputFolder: api.openOutputFolder,
  startHeartbeat: api.startHeartbeat,
};

/** {bytes, warnings} of an Excel本体 export (also accepts bare bytes from an older api client). */
export function normalizeJobResult(r: ExcelJobResult | Uint8Array): ExcelJobResult {
  if (r instanceof Uint8Array) return { bytes: r, warnings: [] };
  const warnings = Array.isArray(r?.warnings) ? r.warnings.filter((w): w is string => typeof w === 'string' && w.trim() !== '') : [];
  return { bytes: r.bytes, warnings };
}

/** True when the page was opened with ?mock=1 (only honoured in dev builds, see main.tsx). */
export function wantsMock(search: string): boolean {
  const p = new URLSearchParams(search);
  const v = p.get('mock');
  return v !== null && v !== '0' && v !== 'false';
}
