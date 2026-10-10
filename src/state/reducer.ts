/**
 * Pure reducer of the whole app (no I/O, no module calls except the history helpers).
 * StrictMode may call it twice per action: it must stay free of side effects.
 */
import type { ServerStatus } from '../api';
import type { DeleteImpact, JoinPlan, JoinResult, KeyAnalysis, NormalizeOptions, PremergeLayout } from '../core/contracts';
import type { EditorSnapshot, MergeReport, SourceBook, TableRegion } from '../core/types';
import { canRedo, canUndo, commit, initHistory, redo, undo } from '../core/doc';
import { moveBefore } from './groups';
import { DEFAULT_OPTIONS, pickColor, sameServerStatus } from './helpers';
import { analysisInputKey, effectiveBaseId, filesKey } from './keysPlan';
import { sameSheetOptions, withOtherSheets, withSourceSheets, type SheetOptions } from './sheetPlan';
import type { AppState, BusyState, ColSelectionState, ExportState, FileEntry, GridPos, ManualKey, ModalState } from './types';

export const EMPTY_SELECTION: ColSelectionState = { cols: new Set<string>(), anchor: null };
const IDLE_EXPORT: ExportState = { phase: 'idle', engine: null, message: '', ratio: null };

export const initialState: AppState = {
  files: [],
  seq: 1,
  baseId: null,
  options: DEFAULT_OPTIONS,
  analysisKey: null,
  analysis: null,
  plan: null,
  planFilesKey: null,
  manualKeys: {},
  joinResult: null,
  keyError: null,
  importOverrides: {},
  includeOtherSheets: false,
  includeSourceSheets: false,
  phase: 'files',
  merged: null,
  selection: EMPTY_SELECTION,
  active: null,
  liveWidths: {},
  modal: null,
  toast: null,
  toastSeq: 1,
  busy: null,
  exportState: IDLE_EXPORT,
  server: { checked: false, status: null, connected: true },
};

export type Action =
  | { type: 'server/status'; status: ServerStatus | null }
  | { type: 'server/connection'; connected: boolean }
  /** a fresh /api/status answer (e.g. when the export dialog opens): replaces the status of a known server, never creates one */
  | { type: 'server/refresh'; status: ServerStatus }
  | { type: 'files/add'; entries: Array<{ id: string; fileName: string; size: number; lastModified?: number }> }
  | { type: 'file/progress'; id: string; message: string }
  | { type: 'file/loaded'; id: string; book: SourceBook; sheet: string; region: TableRegion | null; error?: string; note?: string }
  | { type: 'file/error'; id: string; error: string }
  | { type: 'file/sheet'; id: string; sheet: string; region: TableRegion | null; error?: string }
  | { type: 'file/region'; id: string; region: TableRegion; error?: string; lastRowManual?: boolean }
  | { type: 'file/remove'; id: string }
  /**
   * Drag / Alt+← → of a file card BEFORE the merge: `id` moves right before `beforeId` (null = to the end). The files are the
   * same, so the key plan / manual keys / regions / import choices stay (filesKey ignores the order); only what depends on the
   * order of the files (analysis pairs, default import order, merge preview, appended rows) is recomputed.
   */
  | { type: 'files/move'; id: string; beforeId: string | null }
  | { type: 'base/set'; id: string }
  /** 「他のシートも出力に含める」 (also sets / clears includeSourceSheets, see withOtherSheets) */
  | { type: 'output/otherSheets'; include: boolean }
  /** 「他のファイルのシートも含める」 (ticking it also ticks the first one, see withSourceSheets) */
  | { type: 'output/sourceSheets'; include: boolean }
  | { type: 'options/set'; options: Partial<NormalizeOptions> }
  | {
      type: 'analysis/done';
      key: string;
      analysis: KeyAnalysis;
      plan: JoinPlan;
      joinResult: JoinResult | null;
      error?: string;
      /** manual key candidates that are part of `analysis` (see ManualKey); default none */
      manualKeys?: Record<string, ManualKey[]>;
    }
  | { type: 'analysis/error'; key: string; error: string }
  | { type: 'plan/set'; plan: JoinPlan; joinResult: JoinResult | null; error?: string }
  | { type: 'imports/set'; changes: Record<string, boolean> }
  | {
      type: 'merge/done';
      bytes: Uint8Array;
      layout: PremergeLayout;
      report: MergeReport;
      book: SourceBook;
      snapshot: EditorSnapshot;
      baseFileName: string;
      format: 'xlsx' | 'xlsm';
    }
  | { type: 'merge/reset' }
  | { type: 'doc/commit'; snapshot: EditorSnapshot; selection?: ColSelectionState; active?: GridPos | null }
  | { type: 'doc/undo' }
  | { type: 'doc/redo' }
  | { type: 'grid/selection'; selection: ColSelectionState }
  | { type: 'grid/active'; active: GridPos | null }
  | { type: 'grid/liveWidth'; id: string; px: number | null }
  | { type: 'modal/open'; modal: ModalState }
  | { type: 'modal/close' }
  | { type: 'toast/show'; text: string; kind: 'info' | 'error'; detail?: string }
  | { type: 'toast/clear'; id?: number }
  | { type: 'busy/set'; busy: BusyState | null }
  | { type: 'export/update'; patch: Partial<ExportState> }
  | { type: 'export/reset' };

// ---------------------------------------------------------------------------------------------

function updateFile(state: AppState, id: string, fn: (f: FileEntry) => FileEntry): AppState {
  let hit = false;
  const files = state.files.map((f) => {
    if (f.id !== id) return f;
    hit = true;
    return fn(f);
  });
  return hit ? { ...state, files } : state;
}

/**
 * A finished (or failed) export describes the document as it was then: once the document changes, the export is history. The step
 * indicator goes back to 「列の編集」 and the window asks before it is closed with unsaved edits again.
 */
function staleExport(ex: ExportState): ExportState {
  return ex.phase === 'done' || ex.phase === 'error' ? IDLE_EXPORT : ex;
}

/** Drop analysis results that no longer match the inputs. */
function invalidateAnalysis(state: AppState): AppState {
  const baseId = effectiveBaseId(state.files, state.baseId);
  const s = baseId === state.baseId ? state : { ...state, baseId };
  const key = analysisInputKey(s.files, s.baseId, s.options);
  if (key === s.analysisKey) return s;
  return { ...s, analysis: null, joinResult: null, keyError: null, analysisKey: null };
}

/** Keeps the active cell inside the doc (row = Excel row within the data rows, col = an existing id). */
export function clampActive(active: GridPos | null, snapshot: EditorSnapshot): GridPos | null {
  if (!active) return null;
  const { doc } = snapshot;
  const cols = doc.columns;
  if (cols.length === 0) return null;
  const last = Math.max(doc.firstDataRow, doc.lastDataRow);
  const row = Math.max(doc.firstDataRow, Math.min(last, active.row));
  if (cols.some((c) => c.id === active.col)) return row === active.row ? active : { row, col: active.col };
  return { row, col: cols[0].id };
}

function pruneSelection(sel: ColSelectionState, snapshot: EditorSnapshot): ColSelectionState {
  if (sel.cols.size === 0) return sel;
  const ids = new Set(snapshot.doc.columns.map((c) => c.id));
  const cols = new Set([...sel.cols].filter((id) => ids.has(id)));
  if (cols.size === sel.cols.size) return sel;
  return { cols, anchor: sel.anchor && ids.has(sel.anchor) ? sel.anchor : null };
}

export function reducer(state: AppState, action: Action): AppState {
  switch (action.type) {
    case 'server/status':
      return { ...state, server: { ...state.server, checked: true, status: action.status } };
    case 'server/refresh':
      // standalone (no server was found at start) stays standalone; an unchanged answer keeps the same state object
      return !state.server.status || sameServerStatus(state.server.status, action.status) ? state : { ...state, server: { ...state.server, status: action.status } };
    case 'server/connection':
      return state.server.connected === action.connected ? state : { ...state, server: { ...state.server, connected: action.connected } };

    // ----------------------------------------------------------------- files
    case 'files/add': {
      const used = state.files.map((f) => f.color);
      const added: FileEntry[] = [];
      for (const e of action.entries) {
        const color = pickColor([...used, ...added.map((a) => a.color)]);
        added.push({ id: e.id, fileName: e.fileName, size: e.size, lastModified: e.lastModified, color, status: 'loading', message: '読み込み中…' });
      }
      const files = [...state.files, ...added];
      const seq = state.seq + action.entries.length;
      const baseId = state.baseId && files.some((f) => f.id === state.baseId) ? state.baseId : (files[0]?.id ?? null);
      return invalidateAnalysis({ ...state, files, seq, baseId });
    }
    case 'file/progress':
      return updateFile(state, action.id, (f) => (f.status === 'loading' ? { ...f, message: action.message } : f));
    case 'file/loaded': {
      const next = updateFile(state, action.id, (f) => ({
        ...f,
        status: 'ready',
        message: undefined,
        error: action.error,
        note: action.note,
        book: action.book,
        sheet: action.sheet,
        region: action.region ?? undefined,
        detected: action.region ?? undefined,
        lastRowManual: undefined,
        size: f.size || action.book.size,
      }));
      return invalidateAnalysis(next);
    }
    case 'file/error': {
      const next = updateFile(state, action.id, (f) => ({ ...f, status: 'error', message: undefined, error: action.error, book: undefined, region: undefined, detected: undefined }));
      return invalidateAnalysis(next);
    }
    case 'file/sheet': {
      const next = updateFile(state, action.id, (f) => ({
        ...f,
        sheet: action.sheet,
        region: action.region ?? undefined,
        detected: action.region ?? undefined,
        lastRowManual: undefined,
        error: action.error,
      }));
      return invalidateAnalysis(next);
    }
    case 'file/region': {
      const next = updateFile(state, action.id, (f) => ({ ...f, region: action.region, error: action.error, lastRowManual: action.lastRowManual || undefined }));
      return invalidateAnalysis(next);
    }
    case 'file/remove': {
      const files = state.files.filter((f) => f.id !== action.id);
      const baseId = state.baseId === action.id ? (files.find((f) => f.status === 'ready')?.id ?? files[0]?.id ?? null) : state.baseId;
      const s = invalidateAnalysis({ ...state, files, baseId });
      return files.length < 2 ? { ...s, analysis: null, plan: null, planFilesKey: null, manualKeys: {}, joinResult: null, analysisKey: null, keyError: null } : s;
    }
    case 'files/move': {
      if (state.phase !== 'files') return state;
      const files = moveBefore(state.files, action.id, action.beforeId);
      return files ? invalidateAnalysis({ ...state, files }) : state;
    }
    case 'output/otherSheets':
    case 'output/sourceSheets': {
      // merge settings: they cannot change once the merge is done (the controls are disabled; see SettingsBar)
      if (state.phase !== 'files') return state;
      const cur: SheetOptions = { includeOtherSheets: state.includeOtherSheets, includeSourceSheets: state.includeSourceSheets };
      const next = action.type === 'output/otherSheets' ? withOtherSheets(cur, action.include) : withSourceSheets(cur, action.include);
      return sameSheetOptions(cur, next) ? state : { ...state, ...next };
    }
    case 'base/set': {
      if (!state.files.some((f) => f.id === action.id) || state.baseId === action.id) return state;
      return invalidateAnalysis({ ...state, baseId: action.id, plan: null, planFilesKey: null, manualKeys: {} });
    }
    case 'options/set':
      return invalidateAnalysis({ ...state, options: { ...state.options, ...action.options } });

    // ----------------------------------------------------------------- keys
    case 'analysis/done': {
      const current = analysisInputKey(state.files, state.baseId, state.options);
      if (action.key !== current) return state; // stale result
      // the controller already carried the user's plan choices over (carryOverPlan) before joining
      return {
        ...state,
        analysisKey: action.key,
        analysis: action.analysis,
        plan: action.plan,
        planFilesKey: filesKey(state.files, state.baseId),
        manualKeys: action.manualKeys ?? {},
        joinResult: action.joinResult,
        keyError: action.error ?? null,
      };
    }
    case 'analysis/error': {
      const current = analysisInputKey(state.files, state.baseId, state.options);
      if (action.key !== current) return state;
      return { ...state, analysisKey: action.key, analysis: null, joinResult: null, keyError: action.error };
    }
    case 'plan/set':
      if (!state.analysis) return state;
      return { ...state, plan: action.plan, joinResult: action.joinResult, keyError: action.error ?? null };
    case 'imports/set':
      return { ...state, importOverrides: { ...state.importOverrides, ...action.changes } };

    // ----------------------------------------------------------------- merge / doc
    case 'merge/done': {
      const snapshot = action.snapshot;
      const first = snapshot.doc.columns[0]?.id;
      const initialWidths: Record<string, number | undefined> = {};
      for (const c of snapshot.doc.columns) initialWidths[c.id] = c.props.width;
      return {
        ...state,
        phase: 'merged',
        merged: {
          bytes: action.bytes,
          layout: action.layout,
          report: action.report,
          book: action.book,
          history: initHistory(snapshot),
          initialWidths,
          baseFileName: action.baseFileName,
          format: action.format,
        },
        selection: EMPTY_SELECTION,
        active: first ? { row: snapshot.doc.firstDataRow, col: first } : null,
        liveWidths: {},
        modal: null,
        exportState: IDLE_EXPORT,
      };
    }
    case 'merge/reset':
      return { ...state, phase: 'files', merged: null, selection: EMPTY_SELECTION, active: null, liveWidths: {}, modal: null, exportState: IDLE_EXPORT };
    case 'doc/commit': {
      if (!state.merged) return state;
      const history = commit(state.merged.history, action.snapshot);
      return {
        ...state,
        // an error toast about an earlier refused edit must not stay on screen next to the edit that has just worked
        toast: state.toast?.kind === 'error' ? null : state.toast,
        exportState: staleExport(state.exportState),
        merged: { ...state.merged, history },
        selection: action.selection ?? pruneSelection(state.selection, action.snapshot),
        active: action.active !== undefined ? action.active : clampActive(state.active, action.snapshot),
        liveWidths: {},
      };
    }
    case 'doc/undo':
    case 'doc/redo': {
      if (!state.merged) return state;
      const h = state.merged.history;
      if (action.type === 'doc/undo' ? !canUndo(h) : !canRedo(h)) return state;
      const history = action.type === 'doc/undo' ? undo(h) : redo(h);
      return {
        ...state,
        exportState: staleExport(state.exportState),
        merged: { ...state.merged, history },
        selection: pruneSelection(state.selection, history.present),
        active: clampActive(state.active, history.present),
        liveWidths: {},
      };
    }
    case 'grid/selection':
      return { ...state, selection: action.selection };
    case 'grid/active':
      return { ...state, active: action.active };
    case 'grid/liveWidth': {
      const liveWidths = { ...state.liveWidths };
      if (action.px === null) delete liveWidths[action.id];
      else liveWidths[action.id] = action.px;
      return { ...state, liveWidths };
    }

    // ----------------------------------------------------------------- ui chrome
    case 'modal/open':
      return { ...state, modal: action.modal };
    case 'modal/close':
      return state.modal ? { ...state, modal: null } : state;
    case 'toast/show':
      return {
        ...state,
        toast: { id: state.toastSeq, text: action.text, kind: action.kind, ...(action.detail && action.detail !== action.text ? { detail: action.detail } : {}) },
        toastSeq: state.toastSeq + 1,
      };
    case 'toast/clear':
      if (!state.toast || (action.id !== undefined && action.id !== state.toast.id)) return state;
      return { ...state, toast: null };
    case 'busy/set':
      return { ...state, busy: action.busy };
    case 'export/update':
      return { ...state, exportState: { ...state.exportState, ...action.patch } };
    case 'export/reset':
      return { ...state, exportState: IDLE_EXPORT };
    default:
      return state;
  }
}

// ---------------------------------------------------------------------------------------------
// selectors

export function presentSnapshot(state: AppState): EditorSnapshot | null {
  return state.merged?.history.present ?? null;
}

export function hasEdits(state: AppState): boolean {
  const h = state.merged?.history;
  return !!h && (h.past.length > 0 || h.present.log.length > 0);
}

/** Step indicator: 0 ファイル, 1 結合キー, 2 列の編集, 3 出力 */
export function currentStep(state: AppState): number {
  if (state.phase === 'merged') {
    return state.modal?.kind === 'export' || state.exportState.phase === 'done' ? 3 : 2;
  }
  const ready = state.files.filter((f) => f.status === 'ready').length;
  return ready >= 2 ? 1 : 0;
}

export type { DeleteImpact };
