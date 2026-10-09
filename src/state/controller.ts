/**
 * Async flows of the app (reading files, analysis, merge, editor ops, export). Talks to the other modules
 * only through `Services` (real facades or the ?mock=1 fakes) and only changes state through `dispatch`.
 * No React here: App.tsx creates one controller and calls its methods from event handlers / effects.
 */
import type { ServerStatus } from '../api';
import type { JoinPlan, JoinResult, KeyAnalysis, KeySource, OpResult, PremergeInput, Progress } from '../core/contracts';
import type { Cell, EditorSnapshot, JoinType, MergedDoc, SourceBook, TableRegion } from '../core/types';
import type { CellPos as GridCellPos } from '../components/grid/types';
import { cellEditText, cellViewOf } from '../components/grid/cellview';
import { defaultColumnPx, makeStyleCache, pxToExcelWidth, type CellCss } from '../core/display/styles';
import { addedColumnIds, docMdw } from './gridAdapter';
import {
  errorLine,
  errorText,
  exportFileName,
  extOf,
  formatSize,
  friendlyError,
  isAcceptedFileName,
  isExcelLockFile,
  normalizeExportName,
  opWarningsMessage,
  readErrorMessage,
  refusalMessage,
  regionWithHeaderRow,
  regionWithLastRow,
  sheetByName,
  stripExt,
  UNSUPPORTED_FILE_MESSAGE,
  yieldToUI,
} from './helpers';
import {
  addedCandidate,
  analysisInputKey,
  anyLoading,
  candidateByColumns,
  carryOverPlan,
  effectiveBaseId,
  effectiveImports,
  excludedSourceIds,
  filesKey,
  joinSummary,
  keyPairsFor,
  keySources,
  readyFiles,
  toggleCandidate,
  withoutSources,
} from './keysPlan';
import { formulaReportViews, formulaToastLines, formulaToastParts } from './formulaReport';
import { folderScanMessage, scanFolders, type DropPayload } from './dropFiles';
import { fileColumnIndexes, groupMove, orderFilesByGroups } from './groups';
import { droppedSheetsLine, formulasDetachedLine } from './notes';
import { hasEdits, type Action } from './reducer';
import { normalizeJobResult, type Services } from './services';
import type { AppState, ExportEngineChoice, GridPos, ManualKey } from './types';

/** Files bigger than this are refused (the browser would run out of memory). */
export const MAX_FILE_BYTES = 300 * 1024 * 1024;

export const DISCARD_EDITS_TITLE = '結合をやり直しますか？';
export const DISCARD_EDITS_MESSAGE = '列の並べ替え・挿入・削除・セルの編集など、結合後に行った変更はすべて失われます。';

export const UNMERGE_MOVE_TITLE = 'タイトルの結合を解除して移動しますか？';
export const UNMERGE_MOVE_OK = '結合を解除して移動';

export interface ExportRequest {
  fileName: string;
  engine: ExportEngineChoice;
}

// Minimal File System Access API typing (not in lib.dom for this TS version).
interface SaveFileHandle {
  createWritable(): Promise<{ write(data: BufferSource | Blob): Promise<void>; close(): Promise<void> }>;
  name?: string;
}
type ShowSaveFilePicker = (opts: {
  suggestedName?: string;
  types?: Array<{ description: string; accept: Record<string, string[]> }>;
}) => Promise<SaveFileHandle>;

export interface Controller {
  init(): void;
  toast(text: string, kind?: 'info' | 'error', detail?: string): void;
  // files
  addFiles(files: ArrayLike<File>): Promise<void>;
  /** files dropped onto the window: plain files plus dropped folders (their Excel files are read) */
  addDropped(payload: DropPayload): Promise<void>;
  removeFile(id: string): void;
  setBase(id: string): void;
  changeSheet(id: string, sheet: string): void;
  changeHeaderRow(id: string, headerRow: number): void;
  /** 最終行: the last data row picked by the user; null = back to the automatic one */
  changeLastDataRow(id: string, lastDataRow: number | null): void;
  // keys
  ensureAnalysis(): Promise<void>;
  updatePlan(plan: JoinPlan): void;
  setJoinType(t: JoinType): void;
  /**
   * 「追加」 of the manual key picker: adds a key built from sheet columns of the base and of `sourceId` (checked, lowest
   * priority) and recomputes the join. `existing`: the same columns were already a candidate (only checked).
   * Returns the Japanese error message instead of throwing.
   */
  addManualKey(sourceId: string, baseCols: number[], otherCols: number[]): { ok: true; id: string; existing: boolean } | { ok: false; error: string };
  /** × of a manual candidate (a re-analysis without it; automatic candidates can only be unchecked) */
  removeManualKey(sourceId: string, candidateId: string): Promise<void>;
  // merge
  merge(): Promise<void>;
  requestRemerge(): void;
  confirmModal(): void;
  /** 「他のシートも出力に含める」 (before the merge only; false = the output holds only the merged sheet) */
  setIncludeOtherSheets(include: boolean): void;
  // editor
  /** false when refused (a title-merge refusal opens the 「結合を解除して移動」 confirmation instead) */
  moveColumns(ids: string[], gap: number): boolean;
  /**
   * A file card was dropped before the card `beforeId` (null = after the last one) or moved with Alt+← →. Files phase: reorders
   * the files (the order of the imported column groups). Merged phase: moves the file's whole column group in the grid with
   * doc.moveColumns, one undo step 「<ファイル名>の列を移動」; refusals go through the same handling as a column drag.
   * true = something changed (merged phase: the move was committed).
   */
  reorderFile(id: string, beforeId: string | null): boolean;
  insertColumns(index: number, count: number): void;
  requestDelete(ids: string[]): void;
  confirmDelete(): void;
  editCell(pos: GridCellPos, text: string): boolean;
  renameColumn(id: string, text: string): void;
  resizeColumn(id: string, px: number, commit: { from: number } | null): void;
  resetColumnWidths(ids: string[]): void;
  undo(): void;
  redo(): void;
  // export
  /** opens the export dialog (and asks the server for its disk-space status again, see refreshServerStatus) */
  openExport(): void;
  /** GET /api/status again: the low-disk notice appears / disappears without reloading. Keeps the old status when the server does not answer. */
  refreshServerStatus(): Promise<void>;
  runExport(req: ExportRequest): Promise<void>;
  cancelExport(): void;
  openSavedFile(): Promise<void>;
  openOutputFolder(): Promise<void>;
  downloadAgain(): void;
  /** true when a pending async flow is running (for tests) */
  idle(): Promise<void>;
}

export function outputFormatOf(book: SourceBook | undefined): 'xlsx' | 'xlsm' {
  if (!book) return 'xlsx';
  return book.format === 'xlsm' || book.features.vba ? 'xlsm' : 'xlsx';
}

/** Region of the merged sheet handed to createEditor (appended rows extend the data rows). */
export function mergedRegion(
  base: { headerRow: number; firstDataRow: number; lastDataRow: number },
  appended: ReadonlyArray<{ row: number }>,
): { headerRow: number; firstDataRow: number; lastDataRow: number } {
  let last = base.lastDataRow;
  for (const a of appended) if (a.row > last) last = a.row;
  return { headerRow: base.headerRow, firstDataRow: base.firstDataRow, lastDataRow: last };
}

/** Browser download of bytes (standalone mode). */
export function downloadBytes(bytes: Uint8Array, fileName: string, mime: string): void {
  const blob = new Blob([bytes as BlobPart], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = fileName;
  a.rel = 'noopener';
  a.style.display = 'none';
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 60_000);
}

const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const XLSM_MIME = 'application/vnd.ms-excel.sheet.macroEnabled.12';

/** Japanese card error for a detected region without data rows (undefined = fine). */
export function regionProblem(region: { firstDataRow: number; lastDataRow: number } | null | undefined): string | undefined {
  if (!region) return undefined;
  if (region.lastDataRow < region.firstDataRow) return 'このシートにはデータ行がありません。シートか「見出し行」を選び直してください';
  return undefined;
}

/**
 * A file whose extension says Excel but whose content was read as plain text (a CSV saved as .xlsx, a text export renamed):
 * it is used, but the card says what happened (the 「値のみ」 tag alone does not explain it).
 */
export function formatNote(fileName: string, book: { format: string }): string | undefined {
  const ext = extOf(fileName);
  if (book.format === 'csv' && ['.xlsx', '.xlsm', '.xls', '.xlsb', '.ods'].includes(ext)) {
    return `拡張子は ${ext} ですが、中身はExcelではなくテキスト（CSV）だったため、値だけを読み込みました。書式・数式はありません`;
  }
  return undefined;
}

/** AbortError (fetch / mock) or the api client's cancellation error (code 'CANCELLED'). */
function isAbort(e: unknown): boolean {
  if (!e || typeof e !== 'object') return false;
  const x = e as { name?: string; code?: string };
  return x.name === 'AbortError' || x.code === 'CANCELLED';
}

export function createController(deps: { services: Services; dispatch: (a: Action) => void; getState: () => AppState }): Controller {
  const { services: svc, dispatch, getState } = deps;
  let seq = 0;
  let readChain: Promise<void> = Promise.resolve();
  let analysisInFlight: string | null = null;
  let analysisPromise: Promise<void> = Promise.resolve();
  let pendingFiles: File[] | null = null;
  let exportAbort: AbortController | null = null;
  let exportRun = 0;
  const pending = new Set<Promise<unknown>>();
  const styleCaches = new WeakMap<object, (id: number | undefined) => CellCss>();

  const track = <T>(p: Promise<T>): Promise<T> => {
    pending.add(p);
    p.finally(() => pending.delete(p)).catch(() => undefined);
    return p;
  };

  const toast = (text: string, kind: 'info' | 'error' = 'info', detail?: string) =>
    dispatch(detail ? { type: 'toast/show', text, kind, detail } : { type: 'toast/show', text, kind });

  /** Error toast: a Japanese sentence; the raw (usually English) text of an unexpected error is only the tooltip. */
  const toastError = (prefix: string, e: unknown) => {
    const f = friendlyError(e);
    toast(`${prefix}: ${f.text}`, 'error', f.detail);
  };

  /** Throttled progress -> busy overlay. */
  const busyProgress = (prefix?: string) => {
    let last = 0;
    return (p: Progress) => {
      const now = Date.now();
      if (now - last < 80 && p.ratio !== 1) return;
      last = now;
      dispatch({ type: 'busy/set', busy: { message: prefix ? `${prefix}${p.message ? `（${p.message}）` : ''}` : p.message, ratio: p.ratio } });
    };
  };

  const present = (): EditorSnapshot | null => getState().merged?.history.present ?? null;

  const cssFor = (doc: MergedDoc) => {
    let c = styleCaches.get(doc.styles);
    if (!c) {
      c = makeStyleCache(doc.styles);
      styleCaches.set(doc.styles, c);
    }
    return c;
  };

  // ------------------------------------------------------------------------------------------ files

  const readOne = async (id: string, file: File) => {
    try {
      if (file.size > MAX_FILE_BYTES) {
        dispatch({ type: 'file/error', id, error: `${file.name}: ファイルが大きすぎます（${formatSize(file.size)}）。${formatSize(MAX_FILE_BYTES)} 以下にしてください` });
        return;
      }
      if (file.size === 0) {
        dispatch({ type: 'file/error', id, error: `${file.name}: 空のファイルです` });
        return;
      }
      await yieldToUI();
      let bytes: Uint8Array;
      try {
        bytes = new Uint8Array(await file.arrayBuffer());
      } catch {
        dispatch({ type: 'file/error', id, error: `${file.name}: ファイルを開けませんでした。Excel などで開いている場合は閉じてから、もう一度選んでください。OneDrive の「オンライン専用」のファイルは、先にPCにダウンロードしてください` });
        return;
      }
      let lastMsg = 0;
      const book = await svc.readBook(bytes, file.name, {
        id,
        onProgress: (p) => {
          const now = Date.now();
          if (now - lastMsg < 120) return;
          lastMsg = now;
          dispatch({ type: 'file/progress', id, message: p.message || '読み込み中…' });
        },
      });
      if (!book.sheets.length) {
        dispatch({ type: 'file/error', id, error: `${file.name}: シートがありません` });
        return;
      }
      let sheet = book.sheets.find((s) => s.state === 'visible')?.name ?? book.sheets[0].name;
      try {
        const picked = svc.pickDataSheet(book);
        if (picked && book.sheets.some((s) => s.name === picked)) sheet = picked;
      } catch {
        /* keep the first visible sheet */
      }
      let region = null;
      let error: string | undefined;
      try {
        region = svc.detectRegion(book, sheet);
        error = regionProblem(region);
      } catch (e) {
        error = `表の見出し行を自動で見つけられませんでした。「見出し行」を選んでください（${errorLine(e)}）`;
      }
      dispatch({ type: 'file/loaded', id, book, sheet, region, error, note: formatNote(file.name, book) });
    } catch (e) {
      dispatch({ type: 'file/error', id, error: readErrorMessage(file.name, e) });
    }
  };

  const addFilesNow = (list: File[]) => {
    const st = getState();
    seq = Math.max(seq, st.seq);
    const entries: Array<{ id: string; fileName: string; size: number; lastModified?: number; file: File }> = [];
    for (const f of list) {
      const id = `s${seq++}`;
      entries.push({ id, fileName: f.name, size: f.size, lastModified: f.lastModified, file: f });
    }
    if (!entries.length) return Promise.resolve();
    // a card that failed to load is replaced when the same file is dropped again (typically after fixing it): no stale error cards
    for (const old of st.files) if (old.status === 'error' && entries.some((e) => e.fileName === old.fileName)) dispatch({ type: 'file/remove', id: old.id });
    dispatch({ type: 'files/add', entries: entries.map(({ id, fileName, size, lastModified }) => ({ id, fileName, size, lastModified })) });
    // read one after the other (memory), but never block the caller
    for (const e of entries) readChain = readChain.then(() => readOne(e.id, e.file));
    return track(readChain);
  };

  const filterFiles = (files: ArrayLike<File>): File[] => {
    const st = getState();
    const out: File[] = [];
    const rejected: string[] = [];
    let dup = 0;
    let locks = 0;
    for (const f of Array.from(files)) {
      if (isExcelLockFile(f.name)) {
        locks++;
        continue;
      }
      if (!isAcceptedFileName(f.name)) {
        rejected.push(f.name);
        continue;
      }
      const same = st.files.some((x) => x.fileName === f.name && x.size === f.size && x.lastModified === f.lastModified && x.status !== 'error');
      if (same || out.some((x) => x.name === f.name && x.size === f.size && x.lastModified === f.lastModified)) {
        dup++;
        continue;
      }
      out.push(f);
    }
    if (rejected.length) toast(`${rejected.slice(0, 3).join('、')}${rejected.length > 3 ? ` ほか${rejected.length - 3}件` : ''}: ${UNSUPPORTED_FILE_MESSAGE}`, 'error');
    else if (dup) toast(`同じファイルはすでに読み込まれています（${dup}件）`);
    else if (locks && !out.length) toast('「~$」で始まるファイルは、Excelでファイルを開いている間だけ作られる一時ファイルです。「~$」のない元のファイルをドロップしてください', 'error');
    return out;
  };

  const addFiles = async (files: ArrayLike<File>) => {
    const list = filterFiles(files);
    if (!list.length) return;
    if (getState().phase === 'merged') {
      pendingFiles = list;
      dispatch({
        type: 'modal/open',
        modal: {
          kind: 'confirm',
          purpose: 'addFiles',
          title: 'ファイルを追加しますか？',
          message: `${list.length} 個のファイルを追加して、結合キーの画面に戻ります。${hasEdits(getState()) ? DISCARD_EDITS_MESSAGE : ''}`,
          okLabel: '追加して戻る',
          danger: hasEdits(getState()),
        },
      });
      return;
    }
    await addFilesNow(list);
  };

  const addDropped = async (payload: DropPayload) => {
    const list: File[] = [...payload.files];
    if (payload.folders.length) {
      const names = payload.folders.map((f) => f.name);
      dispatch({ type: 'busy/set', busy: { message: 'フォルダの中のファイルを探しています…', ratio: null } });
      let scanned;
      try {
        scanned = await scanFolders(payload.folders);
      } finally {
        dispatch({ type: 'busy/set', busy: null });
      }
      const msg = folderScanMessage(names, scanned);
      list.push(...scanned.files);
      if (msg) toast(msg.text, msg.kind);
    }
    if (!list.length) return;
    await addFiles(list);
  };

  const changeSheet = (id: string, sheet: string) => {
    const f = getState().files.find((x) => x.id === id);
    if (!f?.book || f.sheet === sheet) return;
    try {
      const region = svc.detectRegion(f.book, sheet);
      dispatch({ type: 'file/sheet', id, sheet, region, error: regionProblem(region) });
    } catch (e) {
      dispatch({ type: 'file/sheet', id, sheet, region: null, error: `このシートでは表を見つけられませんでした。「見出し行」を選んでください（${errorLine(e)}）` });
    }
  };

  /** keys.detectRegion with the user's choices; null when it throws or does not honour them (then: local fallback). */
  const detectWith = (book: SourceBook, sheet: string, opts: { headerRow: number; lastDataRow?: number }): TableRegion | null => {
    try {
      const r = svc.detectRegion(book, sheet, opts);
      if (!r || r.sheet !== sheet || r.headerRow !== opts.headerRow) return null;
      if (!(r.firstDataRow > r.headerRow) || !(r.lastDataRow >= r.headerRow) || !(r.lastCol >= r.firstCol) || r.firstCol < 1) return null;
      return r;
    } catch {
      return null;
    }
  };

  const changeHeaderRow = (id: string, headerRow: number) => {
    const f = getState().files.find((x) => x.id === id);
    if (!f?.book || !f.sheet || !Number.isInteger(headerRow) || headerRow < 1) return;
    const sheet = sheetByName(f.book, f.sheet);
    if (!sheet) return;
    let region = detectWith(f.book, f.sheet, { headerRow });
    if (!region) {
      // without a detected region (detection failed) start from "nothing known": header row 0, no data rows
      const detected = f.detected ?? { sheet: f.sheet, headerRow: 0, firstDataRow: 1, lastDataRow: 0, firstCol: 1, lastCol: Math.max(1, sheet.maxCol) };
      region = regionWithHeaderRow(f.book, detected, headerRow);
    }
    const error = region.lastDataRow < region.firstDataRow ? `${headerRow} 行目より下にデータがありません。見出し行を選び直してください` : undefined;
    dispatch({ type: 'file/region', id, region, error });
  };

  const changeLastDataRow = (id: string, lastDataRow: number | null) => {
    const f = getState().files.find((x) => x.id === id);
    if (!f?.book || !f.sheet || !f.region) return;
    const sheet = sheetByName(f.book, f.sheet);
    if (!sheet) return;
    const cur = f.region;
    if (lastDataRow === null) {
      // 「自動」: the table end as detected for the current header row
      const auto = detectWith(f.book, f.sheet, { headerRow: cur.headerRow }) ?? (f.detected && f.detected.headerRow === cur.headerRow ? f.detected : null) ?? regionWithHeaderRow(f.book, { ...cur, lastDataRow: 0 }, cur.headerRow);
      const error = auto.lastDataRow < auto.firstDataRow ? `${cur.headerRow} 行目より下にデータがありません。見出し行を選び直してください` : undefined;
      dispatch({ type: 'file/region', id, region: auto, error });
      return;
    }
    if (!Number.isFinite(lastDataRow)) return;
    const want = Math.max(cur.firstDataRow, Math.min(Math.max(sheet.maxRow, cur.firstDataRow), Math.round(lastDataRow)));
    if (want === cur.lastDataRow && f.lastRowManual) return;
    let region = detectWith(f.book, f.sheet, { headerRow: cur.headerRow, lastDataRow: want });
    if (!region || region.lastDataRow !== want) region = regionWithLastRow(cur, want, sheet.maxRow);
    dispatch({ type: 'file/region', id, region, error: undefined, lastRowManual: true });
  };

  // ------------------------------------------------------------------------------------------ keys

  /**
   * keys.joinSources, minus the files without a key: they take no part in the merge (no appended rows of theirs, no
   * noKey error, no columns) and the summary counts only the files that join.
   */
  const joinFor = (analysis: KeyAnalysis, plan: JoinPlan): JoinResult => withoutSources(svc.joinSources(analysis, plan), excludedSourceIds(analysis, plan));

  /** Manual key candidates put back into a fresh analysis (their columns may be gone: those are dropped). */
  const reapplyManualKeys = (analysis: KeyAnalysis, wanted: Record<string, ManualKey[]>): { analysis: KeyAnalysis; manualKeys: Record<string, ManualKey[]> } => {
    let cur = analysis;
    const kept: Record<string, ManualKey[]> = {};
    for (const [s, list] of Object.entries(wanted)) {
      if (!cur.pairs[s]) continue;
      for (const m of list) {
        try {
          const next = svc.addManualCandidate(cur, s, m.baseCols, m.otherCols);
          const added = addedCandidate(cur.pairs[s], next.pairs[s], m.baseCols, m.otherCols);
          cur = next;
          (kept[s] ??= []).push({ ...m, id: added?.id ?? m.id });
        } catch {
          /* the columns do not exist any more */
        }
      }
    }
    return { analysis: cur, manualKeys: kept };
  };

  const ensureAnalysis = async (): Promise<void> => {
    const st = getState();
    if (st.phase !== 'files' || anyLoading(st.files)) return;
    const key = analysisInputKey(st.files, st.baseId, st.options);
    if (key === st.analysisKey || key === analysisInFlight) return;
    const srcs = keySources(st.files, st.baseId);
    if (!srcs) return;
    analysisInFlight = key;
    const run = (async () => {
      dispatch({ type: 'busy/set', busy: { message: '結合キーを分析しています…', ratio: null } });
      try {
        await yieldToUI();
        const first = svc.analyzeSources(srcs.base, srcs.others, st.options);
        const cur = getState();
        const sameFiles = cur.planFilesKey !== null && cur.planFilesKey === filesKey(st.files, st.baseId);
        // the keys the user built by hand survive a re-analysis of the same files (options, …)
        const { analysis, manualKeys } = reapplyManualKeys(first, sameFiles ? cur.manualKeys : {});
        const fresh = svc.defaultPlan(analysis);
        const plan = carryOverPlan(fresh, cur.plan, { sameFiles, analysis });
        await yieldToUI();
        let joinResult = null;
        let error: string | undefined;
        try {
          joinResult = joinFor(analysis, plan);
        } catch (e) {
          error = `結合の準備でエラーが発生しました: ${errorLine(e)}`;
        }
        dispatch({ type: 'analysis/done', key, analysis, plan, joinResult, error, manualKeys });
      } catch (e) {
        dispatch({ type: 'analysis/error', key, error: `結合キーを分析できませんでした: ${errorLine(e)}` });
      } finally {
        analysisInFlight = null;
        dispatch({ type: 'busy/set', busy: null });
      }
    })();
    analysisPromise = track(run);
    await run;
    // inputs changed while we were busy -> the effect already fired and returned; run again
    const now = getState();
    if (analysisInputKey(now.files, now.baseId, now.options) !== key) await ensureAnalysis();
  };

  const updatePlan = (plan: JoinPlan) => {
    const st = getState();
    if (!st.analysis) return;
    try {
      const joinResult = joinFor(st.analysis, plan);
      dispatch({ type: 'plan/set', plan, joinResult });
    } catch (e) {
      dispatch({ type: 'plan/set', plan, joinResult: null, error: `結合の準備でエラーが発生しました: ${errorLine(e)}` });
    }
  };

  const addManualKey: Controller['addManualKey'] = (sourceId, baseCols, otherCols) => {
    const st = getState();
    const { analysis, plan } = st;
    if (st.phase !== 'files') return { ok: false, error: '結合後はキーを変えられません（「結合をやり直す」で戻れます）' };
    if (!analysis || !plan || st.analysisKey === null || !analysis.pairs[sourceId]) return { ok: false, error: '結合キーを分析しています。少し待ってから、もう一度お試しください' };
    const pair = analysis.pairs[sourceId];
    const same = candidateByColumns(pair, baseCols, otherCols);
    if (same) {
      // the same columns are already in the table (automatic or added before): just use it
      updatePlan(toggleCandidate(plan, sourceId, same.id, true));
      return { ok: true, id: same.id, existing: true };
    }
    let next: KeyAnalysis;
    try {
      next = svc.addManualCandidate(analysis, sourceId, baseCols, otherCols);
    } catch (e) {
      return { ok: false, error: friendlyError(e).text };
    }
    const added = addedCandidate(pair, next.pairs[sourceId], baseCols, otherCols);
    if (!added) return { ok: false, error: '列の組み合わせを追加できませんでした' };
    const nextPlan = toggleCandidate(plan, sourceId, added.id, true); // checked, lowest priority
    let joinResult: JoinResult | null = null;
    let error: string | undefined;
    try {
      joinResult = joinFor(next, nextPlan);
    } catch (e) {
      error = `結合の準備でエラーが発生しました: ${errorLine(e)}`;
    }
    const manualKeys = { ...st.manualKeys, [sourceId]: [...(st.manualKeys[sourceId] ?? []), { id: added.id, baseCols: baseCols.slice(), otherCols: otherCols.slice() }] };
    dispatch({ type: 'analysis/done', key: st.analysisKey, analysis: next, plan: nextPlan, joinResult, error, manualKeys });
    return { ok: true, id: added.id, existing: false };
  };

  const removeManualKey = async (sourceId: string, candidateId: string): Promise<void> => {
    const st = getState();
    if (st.phase !== 'files' || !st.analysis || !st.plan || st.analysisKey === null || st.busy) return;
    const srcs = keySources(st.files, st.baseId);
    if (!srcs) return;
    const key = st.analysisKey;
    const prevPlan = st.plan;
    const wanted = { ...st.manualKeys, [sourceId]: (st.manualKeys[sourceId] ?? []).filter((m) => m.id !== candidateId) };
    const run = (async () => {
      dispatch({ type: 'busy/set', busy: { message: '結合キーを更新しています…', ratio: null } });
      try {
        await yieldToUI();
        // the keys module cannot take a candidate away: analyse again and put the remaining manual keys back
        const first = svc.analyzeSources(srcs.base, srcs.others, st.options);
        const { analysis, manualKeys } = reapplyManualKeys(first, wanted);
        const without: JoinPlan = { ...prevPlan, tiers: { ...prevPlan.tiers, [sourceId]: (prevPlan.tiers[sourceId] ?? []).filter((id) => id !== candidateId) } };
        const plan = carryOverPlan(svc.defaultPlan(analysis), without, { sameFiles: true, analysis });
        let joinResult: JoinResult | null = null;
        let error: string | undefined;
        try {
          joinResult = joinFor(analysis, plan);
        } catch (e) {
          error = `結合の準備でエラーが発生しました: ${errorLine(e)}`;
        }
        dispatch({ type: 'analysis/done', key, analysis, plan, joinResult, error, manualKeys });
      } catch (e) {
        toastError('結合キーを更新できませんでした', e);
      } finally {
        dispatch({ type: 'busy/set', busy: null });
      }
    })();
    await track(run);
  };

  // ------------------------------------------------------------------------------------------ merge

  const merge = async () => {
    const st = getState();
    if (st.phase !== 'files' || st.busy) return; // a second click while merging
    const srcs = keySources(st.files, st.baseId);
    if (!srcs || !st.analysis || !st.plan || !st.joinResult) {
      toast('2つ以上のファイルを読み込み、結合キーを確認してから結合してください', 'error');
      return;
    }
    // files without a key take no part in the merge (they do not block it); with no usable key at all nothing can merge
    const excluded = excludedSourceIds(st.analysis, st.plan);
    const others: KeySource[] = srcs.others.filter((o) => !excluded.includes(o.id));
    if (!others.length) {
      toast('どのファイルにも結合キーが選ばれていません。ファイルカードの「キーを選ぶ…」か「詳細設定…」でキーを選んでください', 'error');
      return;
    }
    const excludedNames = srcs.others.filter((o) => excluded.includes(o.id)).map((o) => o.label);
    const imports = effectiveImports(st.analysis, st.plan, st.joinResult, st.importOverrides);
    if (!imports.length && st.plan.joinType === 'left') {
      toast('取り込む列がありません。「詳細設定…」の「取り込む列」で選んでください', 'error');
      return;
    }
    const baseFile = readyFiles(st.files).find((f) => f.id === srcs.base.id);
    const baseBook = srcs.base.book;
    const keyPairs = keyPairsFor(st.joinResult, others.map((o) => o.id));
    const input: PremergeInput = {
      base: srcs.base,
      sources: others,
      rows: st.joinResult.rows,
      imports,
      keyFill: st.joinResult.keyFill,
      // formulas of the imported files that read a key column which is not imported (=A2&"様", VLOOKUP(A2,…)) read the base's partner column
      // instead — only for the files that take part in the merge
      ...(keyPairs ? { keyPairs } : {}),
      // the output holds ONLY the merged sheet unless the user ticked 「他のシートも出力に含める」
      options: { formulaPolicy: 'keep-if-possible', extendTable: true, headerConflict: 'suffix', dropOtherSheets: !st.includeOtherSheets },
    };
    const run = (async () => {
      dispatch({ type: 'busy/set', busy: { message: '結合しています…', ratio: 0 } });
      try {
        await yieldToUI();
        const res = await svc.buildPremerged(input, busyProgress('結合しています…'));
        dispatch({ type: 'busy/set', busy: { message: '結合結果を読み込んでいます…', ratio: null } });
        await yieldToUI();
        const book = await svc.readBook(res.bytes, baseBook.fileName, { id: 'merged' });
        await yieldToUI();
        const sourceInfo: MergedDoc['sourceInfo'] = {};
        for (const f of readyFiles(st.files)) sourceInfo[f.id] = { label: stripExt(f.fileName), color: f.color };
        const region = mergedRegion(srcs.base.region, res.layout.appended);
        const snapshot = svc.createEditor(book, res.layout, sourceInfo, region);
        dispatch({
          type: 'merge/done',
          bytes: res.bytes,
          layout: res.layout,
          report: res.report,
          book,
          snapshot,
          baseFileName: baseFile?.fileName ?? baseBook.fileName,
          format: outputFormatOf(baseBook),
        });
        const sum = joinSummary(st.joinResult!);
        const parts = [`${res.report.importedColumns || imports.length} 列を追加`, `一致 ${sum.matched} 行`];
        if (sum.partial) parts.push(`一部一致 ${sum.partial} 行`);
        if (sum.unmatched) parts.push(`未一致 ${sum.unmatched} 行`);
        if (res.report.appendedRows) parts.push(`末尾に ${res.report.appendedRows} 行追加`);
        // the imported formulas: how many stayed formulas / became values / ended up empty (the side panel's 「数式の取り込み」 has the details)
        const fileName = (sid: string, label: string) => st.files.find((f) => f.id === sid)?.fileName ?? label;
        const fviews = formulaReportViews(res.report, fileName);
        if (fviews.length) parts.push(...formulaToastParts(fviews));
        else if (res.report.formulasToValues) parts.push(`数式→値 ${res.report.formulasToValues} 件`);
        if (excludedNames.length) parts.push(`結合対象外 ${excludedNames.length} ファイル（${excludedNames.join('、')}）`);
        const dropped = res.report.droppedSheets ?? [];
        const detached = res.report.formulasDetached ?? 0;
        // single-sheet output: say what is NOT in the output (the full sentences are in the side panel and in the 警告 list)
        if (dropped.length) parts.push(`他のシート ${dropped.length} 枚は含めません`);
        if (detached > 0) parts.push(`他のシートを参照していた数式 ${detached} 個は値にしました`);
        const detail = [dropped.length ? droppedSheetsLine(dropped) : '', detached > 0 ? formulasDetachedLine(detached) : '', ...formulaToastLines(fviews)].filter(Boolean).join('\n');
        toast(`結合しました: ${parts.join(' / ')}`, 'info', detail || undefined);
      } catch (e) {
        toastError('結合できませんでした', e);
      } finally {
        dispatch({ type: 'busy/set', busy: null });
      }
    })();
    await track(run);
  };

  const requestRemerge = () => {
    const st = getState();
    if (st.phase !== 'merged') return;
    if (!hasEdits(st)) {
      dispatch({ type: 'merge/reset' });
      return;
    }
    dispatch({ type: 'modal/open', modal: { kind: 'confirm', purpose: 'remerge', title: DISCARD_EDITS_TITLE, message: DISCARD_EDITS_MESSAGE, okLabel: '変更を破棄してやり直す', danger: true } });
  };

  const confirmModal = () => {
    const m = getState().modal;
    if (!m || m.kind !== 'confirm') return;
    dispatch({ type: 'modal/close' });
    if (m.purpose === 'remerge') {
      dispatch({ type: 'merge/reset' });
    } else if (m.purpose === 'addFiles') {
      const files = pendingFiles ?? [];
      pendingFiles = null;
      dispatch({ type: 'merge/reset' });
      void addFilesNow(files);
    } else if (m.purpose === 'unmergeMove' && m.move) {
      confirmUnmergeMove(m.move);
    }
  };

  // ------------------------------------------------------------------------------------------ editor

  type AfterFn = (before: EditorSnapshot, next: EditorSnapshot) => Partial<{ selection: AppState['selection']; active: GridPos | null }>;

  /**
   * Runs a doc op on the present snapshot and commits it (or shows the refusal, short and friendly).
   * `onRefused` may take over a refusal (return true = handled, no toast).
   */
  const runOp = (label: string, fn: (s: EditorSnapshot) => OpResult, after?: AfterFn, onRefused?: (r: Extract<OpResult, { ok: false }>, s: EditorSnapshot) => boolean): boolean => {
    const s = present();
    if (!s) return false;
    let r: OpResult;
    try {
      r = fn(s);
    } catch (e) {
      toast(`${label}できませんでした`, 'error', errorText(e));
      return false;
    }
    if (!r || typeof r !== 'object') {
      toast(`${label}できませんでした`, 'error');
      return false;
    }
    if (!r.ok) {
      if (onRefused && onRefused(r, s)) return false;
      const m = refusalMessage(r.reason, `${label}できませんでした`);
      toast(m.text, 'error', m.detail);
      return false;
    }
    if (r.snapshot === s) return false; // no-op (e.g. dropped where the columns already are)
    const extra = after ? after(s, r.snapshot) : {};
    dispatch({ type: 'doc/commit', snapshot: r.snapshot, ...extra });
    const w = opWarningsMessage(r.warnings);
    if (w) toast(w.text, 'info', w.detail);
    return true;
  };

  /**
   * The undo step of a card move is named after the file (「02_面接評価の列を移動」) instead of the doc's 「列を移動」 / 「3 列を移動」
   * (a 「結合を解除して」 prefix stays). A no-op (same snapshot) is returned untouched so that runOp still sees it.
   */
  const relabel = (r: OpResult, before: EditorSnapshot, label: string): OpResult => {
    if (!r.ok || r.snapshot === before || label === '列を移動' || !/列を移動$/.test(r.snapshot.label)) return r;
    const prefix = r.snapshot.label.startsWith('結合を解除して') ? '結合を解除して' : '';
    return { ...r, snapshot: { ...r.snapshot, label: prefix + label } };
  };

  const moveColumns = (ids: string[], gap: number, label = '列を移動'): boolean => {
    if (!ids.length) return false;
    return runOp(
      label,
      (s) => relabel(svc.moveColumns(s, ids, gap), s, label),
      () => ({ selection: { cols: new Set(ids), anchor: ids[0] } }),
      (r, s) => {
        // only merged TITLE cells above the header row are in the way: ask, then unmerge + move as one step
        if (r.fix?.kind !== 'unmergeTitles') return false;
        dispatch({
          type: 'modal/open',
          modal: {
            kind: 'confirm',
            purpose: 'unmergeMove',
            title: UNMERGE_MOVE_TITLE,
            message: r.fix.message || `見出しの上の結合セル（${r.fix.refs.join('、')}）があるため、このままでは移動できません。結合を解除して「選択範囲内で中央」にすると、見た目を保ったまま移動できます。`,
            okLabel: UNMERGE_MOVE_OK,
            move: { ids: ids.slice(), gap, label, refs: r.fix.refs.slice(), version: s.doc.version },
          },
        });
        return true;
      },
    );
  };

  /** Merged phase, card drop: the whole column group of file `id` goes where the card was dropped (see groupMove). */
  const moveFileGroup = (id: string, beforeId: string | null): boolean => {
    const st = getState();
    const s = st.merged?.history.present;
    if (!s) return false;
    const baseId = effectiveBaseId(st.files, st.baseId);
    const groups = fileColumnIndexes(s.doc.columns, baseId);
    const cards = orderFilesByGroups(st.files, groups).map((f) => f.id);
    const order = s.doc.columns.map((c) => c.id);
    const m = groupMove({ order, groups, cards, id, beforeId });
    if (m.kind === 'noop') return false;
    const name = st.files.find((f) => f.id === id)?.fileName ?? id;
    if (m.kind === 'nocols') {
      toast(`「${name}」の列は表にありません（結合に取り込まれていないため、移動できません）`);
      return false;
    }
    return moveColumns(m.ids, m.gap, `${stripExt(name)}の列を移動`);
  };

  const reorderFile = (id: string, beforeId: string | null): boolean => {
    const st = getState();
    if (st.busy) return false;
    if (st.phase === 'merged') return moveFileGroup(id, beforeId);
    const before = st.files;
    dispatch({ type: 'files/move', id, beforeId });
    return getState().files !== before;
  };

  /** 「結合を解除して移動」 confirmed: the same move with { unmergeTitles: true } (one snapshot, one undo step). */
  const confirmUnmergeMove = (move: NonNullable<Extract<AppState['modal'], { kind: 'confirm' }>['move']>) => {
    const s = present();
    if (!s) return;
    if (s.doc.version !== move.version) {
      toast('表が変更されたため、移動を取りやめました。もう一度ドラッグしてください', 'error');
      return;
    }
    const label = move.label ?? '列を移動';
    runOp(label, (snap) => relabel(svc.moveColumns(snap, move.ids, move.gap, { unmergeTitles: true }), snap, label), () => ({
      selection: { cols: new Set(move.ids), anchor: move.ids[0] },
    }));
  };

  const insertColumns = (index: number, count: number) => {
    runOp('列を挿入', (s) => svc.insertColumns(s, index, Math.max(1, count)), (before, next) => {
      const added = addedColumnIds(before.doc, next.doc);
      const st = getState();
      return added.length
        ? { selection: { cols: new Set(added), anchor: added[0] }, active: { row: st.active?.row ?? next.doc.firstDataRow, col: added[0] } }
        : {};
    });
  };

  const doDelete = (ids: string[], convertDependents: boolean) => {
    runOp('列を削除', (s) => svc.deleteColumns(s, ids, { convertDependents }), () => ({ selection: { cols: new Set<string>(), anchor: null } }));
  };

  const requestDelete = (ids: string[]) => {
    const s = present();
    if (!s || !ids.length) return;
    if (ids.length >= s.doc.columns.length) {
      toast('すべての列を削除することはできません', 'error');
      return;
    }
    let impact;
    try {
      impact = svc.analyzeDelete(s, ids);
    } catch (e) {
      toastError('削除の影響を確認できませんでした', e);
      return;
    }
    if (impact.breakCount > 0 || impact.lookupWarnings.length > 0) {
      dispatch({ type: 'modal/open', modal: { kind: 'deleteConfirm', ids, impact } });
      return;
    }
    doDelete(ids, false);
  };

  const confirmDelete = () => {
    const m = getState().modal;
    if (!m || m.kind !== 'deleteConfirm') return;
    dispatch({ type: 'modal/close' });
    doDelete(m.ids, m.impact.breakCount > 0);
  };

  const editCell = (pos: GridCellPos, text: string): boolean => {
    const s = present();
    if (!s) return false;
    const col = s.doc.columns.find((c) => c.id === pos.col);
    if (!col) return false;
    const cell: Cell | undefined = col.cells[pos.row - 1];
    const idx = s.doc.columns.indexOf(col);
    let current = '';
    try {
      const view = cellViewOf(cell, cell?.f ? svc.formulaForDisplay(s.doc, idx, pos.row) : null, col.props.s);
      current = cellEditText(view, cssFor(s.doc)(view?.styleId));
    } catch {
      current = '';
    }
    if (current === text) return false;
    return runOp('セルを編集', (snap) => svc.setCellInput(snap, pos.col, pos.row, text));
  };

  const renameColumn = (id: string, text: string) => {
    const s = present();
    if (!s) return;
    const t = text.trim();
    if (!t) {
      toast('見出しを入力してください', 'error');
      return;
    }
    runOp('見出しを変更', (snap) => svc.setCellInput(snap, id, snap.doc.headerRow, t));
  };

  const resizeColumn = (id: string, px: number, commit: { from: number } | null) => {
    if (!commit) {
      dispatch({ type: 'grid/liveWidth', id, px });
      return;
    }
    const s = present();
    if (!s || Math.round(px) === Math.round(commit.from)) {
      dispatch({ type: 'grid/liveWidth', id, px: null });
      return;
    }
    const ok = runOp('列幅を変更', (snap) => svc.setColumnWidth(snap, id, pxToExcelWidth(px, docMdw(snap.doc))));
    if (!ok) dispatch({ type: 'grid/liveWidth', id, px: null });
  };

  /** 「列幅を元に戻す」: the widths right after the merge (inserted columns: the sheet default), one undo step. */
  const resetColumnWidths = (ids: string[]) => {
    const merged = getState().merged;
    if (!merged || !ids.length) return;
    runOp('列幅を元に戻す', (snap) => {
      let cur = snap;
      const warnings: string[] = [];
      const fallback = pxToExcelWidth(defaultColumnPx(docMdw(snap.doc)), docMdw(snap.doc));
      for (const id of ids) {
        const col = cur.doc.columns.find((c) => c.id === id);
        if (!col) continue;
        const target = merged.initialWidths[id] ?? fallback;
        if (col.props.width === target || (col.props.width === undefined && merged.initialWidths[id] === undefined)) continue;
        const r = svc.setColumnWidth(cur, id, target);
        if (!r.ok) return r;
        warnings.push(...r.warnings);
        cur = r.snapshot;
      }
      return { ok: true, snapshot: cur, warnings };
    });
  };

  // ------------------------------------------------------------------------------------------ export

  const refreshServerStatus = (): Promise<void> => {
    // standalone (no server was found at start): nothing to ask — and a server that shows up later has no heartbeat, so it is not adopted
    if (!getState().server.status) return Promise.resolve();
    return track(
      (async () => {
        let status: ServerStatus | null = null;
        try {
          status = await svc.getServerStatus();
        } catch {
          status = null;
        }
        // null = the server did not answer: the heartbeat reports that; the last known status stays
        if (status) dispatch({ type: 'server/refresh', status });
      })(),
    );
  };

  const openExport = () => {
    const st = getState();
    if (!st.merged) return;
    if (st.exportState.phase === 'running') {
      dispatch({ type: 'modal/open', modal: { kind: 'export' } });
      return;
    }
    dispatch({ type: 'export/reset' });
    dispatch({ type: 'modal/open', modal: { kind: 'export' } });
    // free space changes while the app is open (and the dialog is where it matters): ask again
    void refreshServerStatus();
  };

  const runExport = async (req: ExportRequest) => {
    const st = getState();
    const merged = st.merged;
    if (!merged || st.exportState.phase === 'running') return;
    const format = merged.format;
    const name = normalizeExportName(req.fileName || exportFileName(merged.baseFileName, format), format);
    // the server stopped answering (window of the black console closed): export like in standalone mode (built-in engine + browser save)
    const serverMode = !!st.server.status && st.server.connected;
    const excelOk = serverMode && !!st.server.status?.excel.installed;
    const engine: ExportEngineChoice = req.engine === 'excel' && excelOk ? 'excel' : 'js';
    const snapshot = merged.history.present;
    const mime = format === 'xlsm' ? XLSM_MIME : XLSX_MIME;
    const run = ++exportRun;

    // standalone: ask for the destination FIRST (the picker needs the click's user activation)
    let handle: SaveFileHandle | null = null;
    if (!serverMode) {
      const picker = (globalThis as unknown as { showSaveFilePicker?: ShowSaveFilePicker }).showSaveFilePicker;
      // automated browsers (E2E) cannot answer the native dialog: they get the download instead
      const automated = typeof navigator !== 'undefined' && !!(navigator as { webdriver?: boolean }).webdriver;
      if (typeof picker === 'function' && !automated) {
        try {
          handle = await picker.call(globalThis, {
            suggestedName: name,
            types: [{ description: format === 'xlsm' ? 'Excel マクロ有効ブック' : 'Excel ブック', accept: { [mime]: [`.${format}`] } }],
          });
        } catch (e) {
          if (isAbort(e)) return; // the user closed the save dialog
          handle = null; // picker unavailable (e.g. file://) -> download
        }
      }
    }

    const ctrl = new AbortController();
    exportAbort = ctrl;
    const progress = (p: Progress) => {
      if (run !== exportRun) return;
      dispatch({ type: 'export/update', patch: { message: p.message || '出力しています…', ratio: p.ratio } });
    };
    dispatch({
      type: 'export/update',
      patch: { phase: 'running', engine, message: engine === 'excel' ? 'Excel本体で出力しています…' : '出力しています…', ratio: null, error: undefined, fellBack: false, savedPath: undefined, savedName: undefined, delivered: undefined, bytes: undefined, fileName: name, warnings: undefined },
    });
    const task = (async () => {
      let bytes: Uint8Array;
      let fellBack = false;
      let usedEngine = engine;
      let warnings: string[] = [];
      try {
        if (engine === 'excel') {
          try {
            const job = svc.logToRunnerJob(snapshot.log, { format, sheet: snapshot.doc.sheetName });
            const res = normalizeJobResult(await svc.runExcelJob(merged.bytes, job, progress, ctrl.signal));
            bytes = res.bytes;
            warnings = res.warnings;
          } catch (e) {
            if (ctrl.signal.aborted || isAbort(e)) throw e;
            fellBack = true;
            usedEngine = 'js';
            toast('Excel本体で出力できなかったため、内蔵エンジンで出力します', 'error', errorText(e));
            dispatch({ type: 'export/update', patch: { engine: 'js', fellBack: true, message: '内蔵エンジンで出力しています…', ratio: null } });
            bytes = await svc.exportWithJs(merged.bytes, snapshot.log, progress);
          }
        } else {
          bytes = await svc.exportWithJs(merged.bytes, snapshot.log, progress);
        }
        if (run !== exportRun || ctrl.signal.aborted) return;
        dispatch({ type: 'export/update', patch: { message: '保存しています…', ratio: null } });
        if (serverMode) {
          try {
            const saved = await svc.saveToOutput(bytes, name);
            dispatch({ type: 'export/update', patch: { phase: 'done', engine: usedEngine, fellBack, savedPath: saved.path, savedName: saved.name, bytes, fileName: saved.name, message: '', warnings } });
          } catch (e) {
            // could not write into the output folder: hand the file to the browser instead
            downloadBytes(bytes, name, mime);
            dispatch({ type: 'export/update', patch: { phase: 'done', engine: usedEngine, fellBack, delivered: 'download', bytes, fileName: name, message: '', warnings } });
            toast('出力フォルダに保存できなかったため、ダウンロードしました', 'error', errorText(e));
          }
        } else if (handle) {
          const w = await handle.createWritable();
          await w.write(bytes as BufferSource);
          await w.close();
          dispatch({ type: 'export/update', patch: { phase: 'done', engine: usedEngine, fellBack, delivered: 'picker', bytes, fileName: handle.name ?? name, message: '', warnings } });
        } else {
          downloadBytes(bytes, name, mime);
          dispatch({ type: 'export/update', patch: { phase: 'done', engine: usedEngine, fellBack, delivered: 'download', bytes, fileName: name, message: '', warnings } });
        }
      } catch (e) {
        if (run !== exportRun) return;
        if (ctrl.signal.aborted || isAbort(e)) {
          dispatch({ type: 'export/update', patch: { phase: 'idle', message: '', ratio: null } });
          toast('出力を中止しました');
          return;
        }
        dispatch({ type: 'export/update', patch: { phase: 'error', error: `出力できませんでした: ${errorLine(e)}`, message: '', ratio: null } });
      } finally {
        if (exportAbort === ctrl) exportAbort = null;
      }
    })();
    await track(task);
  };

  const cancelExport = () => {
    if (getState().exportState.phase !== 'running') return;
    exportAbort?.abort();
    exportRun++; // results of the JS engine (not abortable) are ignored
    exportAbort = null;
    dispatch({ type: 'export/update', patch: { phase: 'idle', message: '', ratio: null } });
    toast('出力を中止しました');
  };

  const openSavedFile = async () => {
    const p = getState().exportState.savedPath;
    if (!p) return;
    try {
      await svc.openFile(p);
      toast('Excel で開いています…');
    } catch (e) {
      toastError('ファイルを開けませんでした', e);
    }
  };

  const openOutputFolder = async () => {
    try {
      await svc.openOutputFolder();
    } catch (e) {
      toastError('フォルダを開けませんでした', e);
    }
  };

  const downloadAgain = () => {
    const ex = getState().exportState;
    const merged = getState().merged;
    if (!ex.bytes || !ex.fileName) return;
    downloadBytes(ex.bytes, ex.fileName, merged?.format === 'xlsm' ? XLSM_MIME : XLSX_MIME);
  };

  // ------------------------------------------------------------------------------------------

  return {
    init() {
      track(
        (async () => {
          let status = null;
          try {
            status = await svc.getServerStatus();
          } catch {
            status = null;
          }
          dispatch({ type: 'server/status', status });
          if (status) {
            try {
              svc.startHeartbeat();
            } catch {
              /* heartbeat is best effort */
            }
          }
        })(),
      );
    },
    toast,
    addFiles,
    addDropped,
    removeFile: (id) => dispatch({ type: 'file/remove', id }),
    setBase: (id) => dispatch({ type: 'base/set', id }),
    changeSheet,
    changeHeaderRow,
    changeLastDataRow,
    ensureAnalysis,
    updatePlan,
    setJoinType(t) {
      const plan = getState().plan;
      if (plan && plan.joinType !== t) updatePlan({ ...plan, joinType: t });
    },
    addManualKey,
    removeManualKey,
    merge,
    requestRemerge,
    confirmModal,
    setIncludeOtherSheets: (include) => dispatch({ type: 'output/otherSheets', include }),
    moveColumns: (ids, gap) => moveColumns(ids, gap),
    reorderFile,
    insertColumns,
    requestDelete,
    confirmDelete,
    editCell,
    renameColumn,
    resizeColumn,
    resetColumnWidths,
    undo: () => dispatch({ type: 'doc/undo' }),
    redo: () => dispatch({ type: 'doc/redo' }),
    openExport,
    refreshServerStatus,
    runExport,
    cancelExport,
    openSavedFile,
    openOutputFolder,
    downloadAgain,
    async idle() {
      // settle everything that is running now (and what it started)
      for (let i = 0; i < 20 && pending.size; i++) await Promise.allSettled([...pending, analysisPromise]);
    },
  };
}
