/**
 * App state (UI owner). Everything here is plain data; the reducer in ./reducer.ts is pure.
 * Heavy / async work (reading files, analysis, premerge, export) runs in ./controller.ts and only
 * dispatches results.
 */
import type { ServerStatus } from '../api';
import type {
  DeleteImpact,
  History,
  JoinPlan,
  JoinResult,
  KeyAnalysis,
  NormalizeOptions,
  PremergeLayout,
} from '../core/contracts';
import type { MergeReport, SourceBook, TableRegion } from '../core/types';

export type Phase = 'files' | 'merged';

export type FileStatus = 'loading' | 'ready' | 'error';

export interface FileEntry {
  /** stable source id ('s1', 's2', ...) — also the KeySource id */
  id: string;
  fileName: string;
  size: number;
  /** File.lastModified (duplicate drop detection) */
  lastModified?: number;
  /** source colour (card strip + grid header band) */
  color: string;
  status: FileStatus;
  /** progress text while loading */
  message?: string;
  /** Japanese error shown on the card (status 'error', or a region problem on a loaded file) */
  error?: string;
  /** Japanese remark shown on the card of a loaded file that is still used (e.g. a text file with an Excel extension) */
  note?: string;
  book?: SourceBook;
  /** chosen data sheet */
  sheet?: string;
  /** data table region used for the merge (detected, or adjusted by the user's header row / last data row) */
  region?: TableRegion;
  /** region as auto-detected for the current sheet */
  detected?: TableRegion;
  /** the user picked the last data row (最終行) by hand; 「自動」 clears it */
  lastRowManual?: boolean;
}

/**
 * A key candidate the user built from columns of the base and of ONE other file (詳細画面 「列を選んで結合キーに追加」).
 * keys.addManualCandidate adds it to an analysis; a re-analysis (options, sheet …) starts from scratch, so the
 * controller re-applies these (same files only) — and removing one is a re-analysis without it.
 */
export interface ManualKey {
  /** candidate id returned by keys.addManualCandidate ('manual:<baseCols>|<otherCols>') */
  id: string;
  /** sheet column numbers (1-based), 1..3 each, same count */
  baseCols: number[];
  otherCols: number[];
}

export interface ToastState {
  id: number;
  text: string;
  kind: 'info' | 'error';
  /** longer explanation (tooltip of the toast) */
  detail?: string;
}

export interface BusyState {
  message: string;
  /** 0..1 or null (indeterminate) */
  ratio: number | null;
}

/** The merged document being edited (phase 'merged'). */
export interface MergedState {
  /** premerged package (both export engines start from it) */
  bytes: Uint8Array;
  layout: PremergeLayout;
  report: MergeReport;
  /** premerged package as read back by readBook (styles / cellText) */
  book: SourceBook;
  history: History;
  /** column widths right after the merge (「列幅を元に戻す」), column id -> Excel width (undefined = default) */
  initialWidths: Record<string, number | undefined>;
  /** file name of the base file (export default name) */
  baseFileName: string;
  format: 'xlsx' | 'xlsm';
}

/** Grid active cell: row = EXCEL row number of the merged sheet (like the grid's CellPos), col = column id. */
export interface GridPos {
  row: number;
  col: string;
}

export interface ColSelectionState {
  cols: ReadonlySet<string>;
  anchor: string | null;
}

/** A column move that is only blocked by merged title cells above the header row (fix 'unmergeTitles'). */
export interface PendingMove {
  ids: string[];
  gap: number;
  /** label of the undo step / of a failure toast (default 「列を移動」; a card drop says 「<ファイル名>の列を移動」) */
  label?: string;
  /** merged areas that will be unmerged (A1 refs) */
  refs: string[];
  /** doc.version the move was computed for (the move is dropped when the doc changed meanwhile) */
  version: number;
}

export type ModalState =
  | { kind: 'keyDetail'; tab: string | null }
  | { kind: 'deleteConfirm'; ids: string[]; impact: DeleteImpact }
  | { kind: 'export' }
  | { kind: 'help' }
  | {
      kind: 'confirm';
      purpose: 'remerge' | 'addFiles' | 'unmergeMove';
      title: string;
      message: string;
      okLabel: string;
      /** red OK button (discards work) */
      danger?: boolean;
      move?: PendingMove;
    };

export type ExportEngineChoice = 'excel' | 'js';

export type ExportPhase = 'idle' | 'running' | 'done' | 'error';

export interface ExportState {
  phase: ExportPhase;
  engine: ExportEngineChoice | null;
  message: string;
  ratio: number | null;
  /** server mode: saved path */
  savedPath?: string;
  savedName?: string;
  /** standalone: how the file was handed over */
  delivered?: 'picker' | 'download';
  /** last exported bytes (standalone: download again) */
  bytes?: Uint8Array;
  fileName?: string;
  error?: string;
  /** the Excel engine failed and the JS engine was used instead */
  fellBack?: boolean;
  /** warnings of the export (Excel本体 runner warnings, Japanese) */
  warnings?: string[];
}

export interface ServerState {
  checked: boolean;
  status: ServerStatus | null;
  /** heartbeat connection to the local server (server mode only); false = the server stopped answering */
  connected: boolean;
}

export interface AppState {
  files: FileEntry[];
  /** next number for file ids */
  seq: number;
  baseId: string | null;
  options: NormalizeOptions;
  /** key of the inputs the current analysis was computed from (see helpers.analysisInputKey) */
  analysisKey: string | null;
  analysis: KeyAnalysis | null;
  plan: JoinPlan | null;
  /** filesKey() the plan was made for (manual links stay valid while it is unchanged) */
  planFilesKey: string | null;
  /** manual key candidates per source id (see ManualKey); valid for planFilesKey only */
  manualKeys: Record<string, ManualKey[]>;
  joinResult: JoinResult | null;
  /** analysis / join failure (Japanese) */
  keyError: string | null;
  /** user choices in 取り込む列, key = importKey(sourceId, sheet, col) */
  importOverrides: Record<string, boolean>;
  /**
   * 「他のシートも出力に含める」 (default false = the output holds ONLY the merged sheet). Part of the merge settings: it is
   * handed to buildPremerged as options.dropOtherSheets = !includeOtherSheets, so it can only change before the merge.
   */
  includeOtherSheets: boolean;
  /**
   * 「他のファイルのシートも含める」 (round 9, plan A; default false): copy every sheet of the 2nd, 3rd … files into the output,
   * handed to buildPremerged as options.includeSourceSheets. Coupled to includeOtherSheets (see ./sheetPlan.ts: ticking it ticks
   * the first one, clearing the first one clears it) and, like it, a merge setting that only changes before the merge.
   */
  includeSourceSheets: boolean;
  phase: Phase;
  merged: MergedState | null;
  selection: ColSelectionState;
  active: GridPos | null;
  /** live column widths (px) while a resize handle is being dragged */
  liveWidths: Record<string, number>;
  modal: ModalState | null;
  toast: ToastState | null;
  toastSeq: number;
  busy: BusyState | null;
  exportState: ExportState;
  server: ServerState;
}
