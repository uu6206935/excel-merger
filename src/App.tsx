/**
 * Excel 統合ツール — single-page UI (docs/ARCHITECTURE.md §5).
 *
 * State: pure reducer (src/state/reducer.ts). Async flows (read files, analysis, merge, editor ops, export):
 * src/state/controller.ts, talking to the modules only through `Services` (real facades or ?mock=1 fakes).
 * Flow: ファイルを投げる → 結合キー (automatic, adjustable) → 結合する → 列を並べる (grid) → Excel出力.
 */
import { useCallback, useEffect, useMemo, useReducer, useRef, useState, type ReactNode } from 'react';
import type { GridHandle } from './components/grid/types';
import { clickSelect } from './components/grid/selection';
import { AppHeader } from './components/AppHeader';
import { BusyOverlay } from './components/BusyOverlay';
import { CheckPanel } from './components/CheckPanel';
import { ColumnList, PlannedColumnList } from './components/ColumnList';
import { ConfirmDialog } from './components/ConfirmDialog';
import { DeleteConfirmDialog } from './components/DeleteConfirmDialog';
import { EditorPane } from './components/EditorPane';
import { ErrorBoundary } from './components/ErrorBoundary';
import { ExportDialog } from './components/ExportDialog';
import { AddFileCard, EmptyState, FileCard } from './components/FileCard';
import { FileCardRow, type CardGhostInfo } from './components/FileCardRow';
import { FormulaReport } from './components/FormulaReport';
import { HelpDialog } from './components/HelpPanel';
import { KeyDetailModal } from './components/KeyDetailModal';
import { MergePreview } from './components/MergePreview';
import { KeyControls, OtherSheetsOption, SettingsBar } from './components/SettingsBar';
import { SidePanel } from './components/SidePanel';
import { Toast } from './components/Toast';
import type { Cell, SourceBook } from './core/types';
import { createController } from './state/controller';
import { readDropPayload } from './state/dropFiles';
import { buildGridModel, columnListItems, keyRolesFromPlan } from './state/gridAdapter';
import { exampleCellColumn, formulaCheckRows, formulaColumnsOf, formulaReportViews } from './state/formulaReport';
import { fileColumnIds, fileColumnIndexes, orderFilesByGroups, selectionOwner } from './state/groups';
import { ACCEPT_ATTR, APP_SOURCE_COLOR, exportFileName, isTextEditingTarget, stripExt, undoRedoAction } from './state/helpers';
import {
  checkGroups,
  effectiveBaseId,
  effectiveImports,
  excludedSourceIds,
  fileKeyColumns,
  joinSummary,
  joinedSourceCount,
  plannedColumns,
  readyFiles,
  summaryText,
  summaryUnit,
  tierLabels,
  toggleFamily,
  unmatchedEverywhere,
} from './state/keysPlan';
import { droppedSheetsLine, fidelityNotes, formulasDetachedLine } from './state/notes';
import { currentStep, hasEdits, initialState, reducer, type Action } from './state/reducer';
import type { Services } from './state/services';
import type { AppState, ColSelectionState } from './state/types';
import { canRedo, canUndo } from './core/doc';

const TOAST_MS = 5000;
const TOAST_ERROR_MS = 9000;

export interface DemoOptions {
  files: () => File[];
  stage: 'files' | 'merged';
}

export default function App({ services, demo }: { services: Services; demo?: DemoOptions | null }) {
  const [state, rawDispatch] = useReducer(reducer, initialState);
  // always-current state for async flows (updated synchronously on every dispatch)
  const stateRef = useRef<AppState>(initialState);
  const dispatch = useCallback((a: Action) => {
    stateRef.current = reducer(stateRef.current, a);
    rawDispatch(a);
  }, []);
  const ctrl = useMemo(() => createController({ services, dispatch, getState: () => stateRef.current }), [services, dispatch]);
  const gridRef = useRef<GridHandle>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [dragging, setDragging] = useState(false);
  /** files phase: the card the user clicked last (highlight only) */
  const [pickedFile, setPickedFile] = useState<string | null>(null);
  const phaseNow = state.phase;
  useEffect(() => setPickedFile(null), [phaseNow]);

  // ------------------------------------------------------------------ lifecycle
  useEffect(() => ctrl.init(), [ctrl]);

  useEffect(() => {
    void ctrl.ensureAnalysis();
  }, [ctrl, state.files, state.baseId, state.options, state.phase]);

  useEffect(() => {
    if (!state.toast) return;
    const id = state.toast.id;
    // errors and toasts with a 「詳しく」 detail stay longer (time to hover / read)
    const long = state.toast.kind === 'error' || !!state.toast.detail;
    const t = setTimeout(() => dispatch({ type: 'toast/clear', id }), long ? TOAST_ERROR_MS : TOAST_MS);
    return () => clearTimeout(t);
  }, [state.toast, dispatch]);

  // Ctrl+Z / Ctrl+Y / Ctrl+Shift+Z anywhere (not while typing in a form field; the grid handles its own sink).
  // Ctrl+S / Ctrl+O: the Excel habits. The browser's own 「名前を付けて保存」/「開く」 would only confuse: Ctrl+S opens the export
  // dialog (after the merge), Ctrl+O the file chooser.
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      const st = stateRef.current;
      if ((e.ctrlKey || e.metaKey) && !e.altKey && !e.shiftKey && (e.key === 's' || e.key === 'S' || e.key === 'o' || e.key === 'O')) {
        e.preventDefault();
        if (st.modal || st.busy) return;
        if (e.key === 's' || e.key === 'S') {
          if (st.phase === 'merged') {
            gridRef.current?.commitEdit();
            ctrl.openExport();
          } else {
            ctrl.toast('保存は「結合する」のあと、「Excel出力」から行います');
          }
        } else {
          fileInputRef.current?.click();
        }
        return;
      }
      // F6 (Excel: next pane) jumps between the grid and the main button of the side panel; Tab stays inside the grid until its last column
      if (e.key === 'F6' && !e.ctrlKey && !e.altKey && !st.modal && !st.busy) {
        e.preventDefault();
        const primary = document.querySelector<HTMLElement>('[data-testid=export-button], [data-testid=merge-button]:not([disabled])');
        if (e.shiftKey) {
          if (st.phase === 'merged') gridRef.current?.focus();
          else document.querySelector<HTMLElement>('[data-testid=file-card-0] select, [data-testid=drop-zone] button')?.focus();
        } else primary?.focus();
        return;
      }
      const act = undoRedoAction(e);
      if (!act) return;
      if (st.phase !== 'merged' || st.modal || st.busy || isTextEditingTarget(e.target)) return;
      e.preventDefault();
      if (act === 'undo') ctrl.undo();
      else ctrl.redo();
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [ctrl]);

  // the whole window accepts file drops (and a stray drop never navigates away from the app)
  useEffect(() => {
    let depth = 0;
    const hasFiles = (e: DragEvent) => !!e.dataTransfer && Array.from(e.dataTransfer.types ?? []).includes('Files');
    const onEnter = (e: DragEvent) => {
      if (!hasFiles(e)) return;
      e.preventDefault();
      depth++;
      setDragging(true);
    };
    const onOver = (e: DragEvent) => {
      if (!hasFiles(e)) return;
      e.preventDefault();
      if (e.dataTransfer) e.dataTransfer.dropEffect = 'copy';
    };
    const onLeave = (e: DragEvent) => {
      if (!hasFiles(e)) return;
      depth = Math.max(0, depth - 1);
      if (depth === 0) setDragging(false);
    };
    const onDrop = (e: DragEvent) => {
      if (!hasFiles(e)) return;
      e.preventDefault();
      depth = 0;
      setDragging(false);
      const st = stateRef.current;
      if (st.busy || st.exportState.phase === 'running') {
        ctrl.toast('処理中です。終わってからもう一度ドロップしてください', 'error');
        return;
      }
      if (st.modal) {
        ctrl.toast('開いている画面を閉じてから、もう一度ドロップしてください', 'error');
        return;
      }
      // folders are only readable during the drop event itself: separate them from the plain files synchronously
      const payload = readDropPayload(e.dataTransfer);
      if (payload.files.length || payload.folders.length) void ctrl.addDropped(payload);
    };
    window.addEventListener('dragenter', onEnter);
    window.addEventListener('dragover', onOver);
    window.addEventListener('dragleave', onLeave);
    window.addEventListener('drop', onDrop);
    return () => {
      window.removeEventListener('dragenter', onEnter);
      window.removeEventListener('dragover', onOver);
      window.removeEventListener('dragleave', onLeave);
      window.removeEventListener('drop', onDrop);
    };
  }, [ctrl]);

  // heartbeat connection to the local server (fired by api.startHeartbeat)
  useEffect(() => {
    const onConn = (e: Event) => {
      const connected = !!(e as CustomEvent<{ connected: boolean }>).detail?.connected;
      const before = stateRef.current.server.connected;
      dispatch({ type: 'server/connection', connected });
      if (!connected && before && stateRef.current.server.status) {
        ctrl.toast('ローカルサーバーとの接続が切れました。黒い画面（サーバー）を閉じた場合は start.bat から起動し直してください', 'error');
      }
    };
    window.addEventListener('excel-merger-server', onConn);
    return () => window.removeEventListener('excel-merger-server', onConn);
  }, [ctrl, dispatch]);

  // closing the window with unsaved edits asks first
  useEffect(() => {
    const onBefore = (e: BeforeUnloadEvent) => {
      const st = stateRef.current;
      if (st.phase === 'merged' && hasEdits(st) && st.exportState.phase !== 'done') {
        e.preventDefault();
        e.returnValue = '';
      }
    };
    window.addEventListener('beforeunload', onBefore);
    return () => window.removeEventListener('beforeunload', onBefore);
  }, []);

  // ?mock=1&demo=files|merged (dev only): preload the demo files (and merge)
  const demoRef = useRef<{ loaded: boolean; merged: boolean }>({ loaded: false, merged: false });
  useEffect(() => {
    if (!demo || demoRef.current.loaded) return;
    demoRef.current.loaded = true;
    void ctrl.addFiles(demo.files());
  }, [demo, ctrl]);
  useEffect(() => {
    if (demo?.stage === 'merged' && !demoRef.current.merged && state.phase === 'files' && state.joinResult && !state.busy) {
      demoRef.current.merged = true;
      void ctrl.merge();
    }
  }, [demo, ctrl, state.phase, state.joinResult, state.busy]);

  // ------------------------------------------------------------------ derived data
  const { files, analysis, plan, joinResult, merged, phase } = state;
  const baseId = effectiveBaseId(files, state.baseId);
  const labels = useMemo(() => Object.fromEntries(files.map((f) => [f.id, stripExt(f.fileName)])), [files]);
  const colors = useMemo(() => Object.fromEntries(files.map((f) => [f.id, f.color])), [files]);
  const readyCount = readyFiles(files).length;
  const loadingAny = files.some((f) => f.status === 'loading');
  const analyzing = phase === 'files' && readyCount >= 2 && !analysis && !state.keyError;
  const step = currentStep(state);
  const excelAvailable = !!state.server.status?.excel.installed;
  const cellText = useCallback((book: SourceBook, cell: Cell | undefined) => services.cellText(book, cell), [services]);

  // files without a selected key take no part in the merge (they join it as soon as a key is selected)
  const excluded = useMemo(() => excludedSourceIds(analysis, plan), [analysis, plan]);
  const excludedLabels = useMemo(() => excluded.map((id) => labels[id] ?? id), [excluded, labels]);
  const noKeyAtAll = !!analysis && !!plan && excluded.length > 0 && excluded.length === Object.keys(analysis.pairs).length;
  const summary = useMemo(() => (joinResult && joinedSourceCount(joinResult) > 0 ? joinSummary(joinResult) : null), [joinResult]);
  const unit = summaryUnit(analysis, plan);
  const summaryLine = summary ? summaryText(summary, unit, excluded.length) : noKeyAtAll && phase === 'files' ? 'どのファイルにも結合キーが選ばれていません' : '';

  // 「数式の取り込み」: per file with formulas (MergeReport.formulaStats), after the merge
  const report = merged?.report ?? null;
  const formulaViews = useMemo(
    () => (phase === 'merged' ? formulaReportViews(report, (sid, label) => files.find((f) => f.id === sid)?.fileName ?? label) : []),
    [phase, report, files],
  );
  const toValuesRows = useMemo(
    () => (phase === 'merged' ? formulaCheckRows(report, (sid, label) => files.find((f) => f.id === sid)?.fileName ?? label) : []),
    [phase, report, files],
  );
  const toValuesGroup = merged && merged.report.formulasToValues > 0 ? { count: merged.report.formulasToValues, rows: toValuesRows } : undefined;
  const groups = useMemo(() => (joinResult ? checkGroups(joinResult, labels, toValuesGroup) : null), [joinResult, labels, toValuesGroup?.count, toValuesRows]); // eslint-disable-line react-hooks/exhaustive-deps
  const exportWarnings = state.exportState.phase === 'done' ? state.exportState.warnings : undefined;
  const notes = useMemo(
    () =>
      fidelityNotes(files, baseId, excelAvailable, merged?.report ?? null, {
        labels,
        exportWarnings,
        excluded,
        merged: phase === 'merged',
        dropOtherSheets: !state.includeOtherSheets,
      }),
    [files, baseId, excelAvailable, merged?.report, labels, exportWarnings, excluded, phase, state.includeOtherSheets],
  );
  const keyRoles = useMemo(() => keyRolesFromPlan(analysis, plan), [analysis, plan]);
  const keyColsByFile = useMemo(() => Object.fromEntries(files.map((f) => [f.id, fileKeyColumns(analysis, plan, f.id)])), [files, analysis, plan]);
  const highlightRows = useMemo(() => unmatchedEverywhere(joinResult) as ReadonlySet<number>, [joinResult]);

  // the detail modal stays open (with the previous numbers) while an option change re-runs the analysis
  const lastAnalysisRef = useRef(analysis);
  if (analysis) lastAnalysisRef.current = analysis;
  const detailAnalysis = analysis ?? (readyCount >= 2 ? lastAnalysisRef.current : null);

  const present = merged?.history.present ?? null;
  const doc = present?.doc ?? null;
  const textBook = useMemo(() => (merged && doc ? { ...merged.book, styles: doc.styles } : null), [merged, doc]);
  const listModel = useMemo(() => {
    if (!doc || !textBook) return null;
    return buildGridModel(doc, {
      baseId,
      keyRoles,
      headerText: (c) => {
        try {
          return services.cellText(textBook, c);
        } catch {
          return c?.v === null || c?.v === undefined ? '' : String(c.v);
        }
      },
    });
  }, [doc, textBook, baseId, keyRoles, services]);
  const listItems = useMemo(() => (listModel ? columnListItems(listModel, APP_SOURCE_COLOR) : []), [listModel]);
  const planned = useMemo(
    () => (phase === 'files' ? plannedColumns(analysis, plan, joinResult, state.importOverrides) : []),
    [phase, analysis, plan, joinResult, state.importOverrides],
  );

  // After the merge the card row is a mini-map of the grid: the cards stand in the order of their column groups
  const fileGroups = useMemo(() => (doc ? fileColumnIndexes(doc.columns, baseId) : null), [doc, baseId]);
  const cardFiles = useMemo(() => (phase === 'merged' && fileGroups ? orderFilesByGroups(files, fileGroups) : files), [phase, fileGroups, files]);
  /** the card whose complete column group is the current column selection */
  const activeFileId = useMemo(() => (phase === 'merged' && doc ? selectionOwner(doc.columns, baseId, state.selection.cols) : null), [phase, doc, baseId, state.selection]);

  // ------------------------------------------------------------------ handlers
  const pickFiles = useCallback(() => {
    const st = stateRef.current;
    if (st.busy) return;
    fileInputRef.current?.click();
  }, []);

  const focusGrid = useCallback(() => gridRef.current?.focus(), []);

  // ---- file cards: reorder by drag / Alt+arrows, click
  const cardsDisabled = !!state.busy || !!state.modal || state.exportState.phase === 'running';
  const canDragCard = useCallback((id: string) => stateRef.current.phase !== 'merged' || !!fileGroups?.has(id), [fileGroups]);
  const describeCard = useCallback((id: string): CardGhostInfo | undefined => {
    const f = stateRef.current.files.find((x) => x.id === id);
    return f ? { name: f.fileName, color: f.color, base: id === effectiveBaseId(stateRef.current.files, stateRef.current.baseId) } : undefined;
  }, []);
  const onReorderCard = useCallback(
    (id: string, beforeId: string | null): boolean => {
      const ok = ctrl.reorderFile(id, beforeId);
      const st = stateRef.current;
      const d = st.merged?.history.present.doc;
      if (ok && st.phase === 'merged' && d) {
        // show where the group went (it may be far away from where it was)
        const ids = fileColumnIds(d.columns, effectiveBaseId(st.files, st.baseId), id);
        requestAnimationFrame(() => gridRef.current?.scrollToColumnsCentered(ids));
      }
      return ok;
    },
    [ctrl],
  );
  const onCardClick = useCallback(
    (id: string) => {
      const st = stateRef.current;
      if (st.phase !== 'merged') {
        setPickedFile(id);
        return;
      }
      const d = st.merged?.history.present.doc;
      const f = st.files.find((x) => x.id === id);
      if (!d || !f) return;
      const ids = fileColumnIds(d.columns, effectiveBaseId(st.files, st.baseId), id);
      if (!ids.length) {
        ctrl.toast(`「${f.fileName}」の列は表にありません（結合に取り込まれていません）`);
        return;
      }
      // the file's columns to the middle of the grid, all of them selected; the grid gets the focus (Ctrl+C / arrow keys)
      gridRef.current?.commitEdit();
      gridRef.current?.scrollToColumnsCentered(ids, { select: true });
      gridRef.current?.focus();
    },
    [ctrl],
  );

  const editorHandlers = useMemo(
    () => ({
      onSelection: (sel: ColSelectionState) => dispatch({ type: 'grid/selection', selection: sel }),
      onActive: (pos: { row: number; col: string }) => dispatch({ type: 'grid/active', active: pos }),
      onMoveColumns: (ids: string[], gap: number) => ctrl.moveColumns(ids, gap),
      onResizeColumn: (id: string, px: number, commit: { from: number } | null) => ctrl.resizeColumn(id, px, commit),
      onCommitEdit: (pos: { row: number; col: string }, text: string) => {
        ctrl.editCell(pos, text);
      },
      onDeleteColumns: (ids: string[]) => ctrl.requestDelete(ids),
      onInsertColumns: (index: number, count: number) => ctrl.insertColumns(index, count),
      onResetColumnWidths: (ids: string[]) => ctrl.resetColumnWidths(ids),
      onUndo: () => ctrl.undo(),
      onRedo: () => ctrl.redo(),
      onNotice: (text: string) => ctrl.toast(text),
    }),
    [ctrl, dispatch],
  );

  const onColumnSelect = useCallback(
    (id: string, mods: { ctrl: boolean; shift: boolean }) => {
      const st = stateRef.current;
      const d = st.merged?.history.present.doc;
      if (!d) return;
      const order = d.columns.map((c) => c.id);
      dispatch({ type: 'grid/selection', selection: clickSelect(order, st.selection, id, mods) });
      dispatch({ type: 'grid/active', active: { row: st.active?.row ?? d.firstDataRow, col: id } });
      requestAnimationFrame(() => gridRef.current?.scrollToColumn(id));
    },
    [dispatch],
  );

  const jumpToRow = useCallback((row: number) => {
    const ok = gridRef.current?.scrollToRow(row, { select: true });
    if (ok === false) ctrl.toast(`${row}行目は表示されていません（非表示の行です）`);
  }, [ctrl]);

  /** 「数式の取り込み」 example row -> the cell of that file's column that held (or still holds) the formula */
  const jumpToFormulaCell = useCallback(
    (sourceId: string, row: number, cell: 'plain' | 'formula') => {
      const st = stateRef.current;
      const d = st.merged?.history.present.doc;
      if (!d) return;
      const f = st.files.find((x) => x.id === sourceId);
      const col = exampleCellColumn(d.columns, sourceId, row, new Set(formulaColumnsOf(f?.book, f?.region).keys()), cell);
      if (!col) {
        ctrl.toast(`「${f?.fileName ?? sourceId}」の列は表にありません（削除されています）`);
        return;
      }
      gridRef.current?.commitEdit();
      if (gridRef.current?.scrollToCell({ row, col }) === false) ctrl.toast(`${row}行目は表示されていません（非表示の行です）`);
    },
    [ctrl],
  );

  const closeModal = useCallback(() => {
    dispatch({ type: 'modal/close' });
    requestAnimationFrame(() => gridRef.current?.focus());
  }, [dispatch]);

  const openExport = useCallback(() => {
    gridRef.current?.commitEdit();
    ctrl.openExport();
  }, [ctrl]);

  // ------------------------------------------------------------------ render pieces
  const mergeDisabledReason =
    readyCount < 2
      ? 'ファイルを2つ以上読み込んでください'
      : loadingAny
        ? 'ファイルを読み込み中です'
        : !joinResult
          ? (state.keyError ?? '結合キーを確認しています')
          : noKeyAtAll
            ? 'どのファイルにも結合キーが選ばれていません。ファイルカードの「キーを選ぶ…」か「詳細設定…」でキーを選んでください'
            : null;

  let settings: ReactNode = null;
  if (phase === 'files' && analysis && plan) {
    settings = (
      <KeyControls
        analysis={analysis}
        plan={plan}
        labels={labels}
        onPlan={ctrl.updatePlan}
        onToggleFamily={(family, on) => ctrl.updatePlan(toggleFamily(analysis, plan, family, on))}
        onOpenDetail={() => dispatch({ type: 'modal/open', modal: { kind: 'keyDetail', tab: null } })}
        onJoinType={ctrl.setJoinType}
        includeOtherSheets={state.includeOtherSheets}
        onIncludeOtherSheets={ctrl.setIncludeOtherSheets}
      />
    );
  } else if (phase === 'merged' && analysis && plan) {
    const ids = Object.keys(analysis.pairs);
    settings = (
      <span className="settings-keys-readonly" title="結合キーを変えるには「結合をやり直す」を押してください">
        <span className="settings-label">結合キー</span>
        {ids.map((s) => (
          <span key={s} className={`key-chip on readonly${excluded.includes(s) ? ' excluded' : ''}`} title={excluded.includes(s) ? `「${labels[s] ?? s}」は結合キーがないため結合の対象外です` : undefined}>
            {ids.length > 1 && <span className="src-dot" style={{ background: colors[s] }} aria-hidden />}
            {excluded.includes(s) ? '結合対象外' : tierLabels(analysis, plan, s)}
          </span>
        ))}
        <OtherSheetsOption checked={state.includeOtherSheets} disabled />
      </span>
    );
  } else if (phase === 'files' && readyCount >= 2 && analyzing) {
    settings = (
      <span className="muted">
        <span className="spinner spinner-sm" aria-hidden /> 結合キーを探しています…
      </span>
    );
  }

  const sideSummary =
    phase === 'merged' && doc && merged ? (
      <>
        <div>
          <b>{(doc.lastDataRow - doc.firstDataRow + 1).toLocaleString()}</b> 行 × <b>{doc.columns.length}</b> 列
          <span className="muted">（追加 {merged.report.importedColumns} 列{merged.report.appendedRows ? `・追加行 ${merged.report.appendedRows}` : ''}）</span>
        </div>
        <div className="muted">
          結合後の変更 {merged.history.past.length} 件{merged.history.future.length ? `（やり直し可能 ${merged.history.future.length}）` : ''}
        </div>
        {(merged.report.droppedSheets?.length ?? 0) > 0 && (
          <div className="side-note" data-testid="side-dropped-sheets">
            {droppedSheetsLine(merged.report.droppedSheets!)}
          </div>
        )}
        {(merged.report.formulasDetached ?? 0) > 0 && (
          <div className="side-note" data-testid="side-formulas-detached">
            {formulasDetachedLine(merged.report.formulasDetached!)}
          </div>
        )}
      </>
    ) : summary ? (
      <>
        <div>
          基準 <b>{summary.baseRows.toLocaleString()}</b> {unit} / 一致 <b>{summary.matched.toLocaleString()}</b>
          {summary.partial > 0 && (
            <>
              {' '}
              / 一部 <b>{summary.partial.toLocaleString()}</b>
            </>
          )}{' '}
          / 未一致 <b className={summary.unmatched ? 'text-drop' : ''}>{summary.unmatched.toLocaleString()}</b>
        </div>
        <div className="muted">追加する列 {effectiveImports(analysis, plan, joinResult, state.importOverrides).length} 列</div>
        {excluded.length > 0 && (
          <div className="side-excluded" data-testid="side-excluded">
            結合対象外: {excludedLabels.map((l) => `「${l}」`).join('、')}（結合キーなし）
          </div>
        )}
      </>
    ) : noKeyAtAll ? (
      <div className="side-excluded" data-testid="side-excluded">
        どのファイルにも結合キーが選ばれていません。ファイルカードの「キーを選ぶ…」でキーを選んでください。
      </div>
    ) : (
      <span className="muted">{files.length ? 'ファイルを2つ以上読み込むと、結合の見込みを表示します' : 'ファイルをドロップすると始まります'}</span>
    );

  const deleteModal = state.modal?.kind === 'deleteConfirm' ? state.modal : null;

  return (
    <div className={`app${dragging ? ' is-dragging-files' : ''}`}>
      <AppHeader server={state.server} step={step} mock={services.mock} onHelp={() => dispatch({ type: 'modal/open', modal: { kind: 'help' } })} onAddFiles={files.length > 0 ? pickFiles : undefined} />

      {files.length > 0 && (
        <FileCardRow
          className={`upload-row${phase === 'merged' ? ' compact' : ''}`}
          disabled={cardsDisabled}
          canDrag={canDragCard}
          describe={describeCard}
          onReorder={onReorderCard}
          onCardClick={onCardClick}
        >
          {cardFiles.map((f, i) => (
            <FileCard
              key={f.id}
              index={i}
              file={f}
              isBase={f.id === baseId}
              single={files.length === 1}
              keyCols={keyColsByFile[f.id] ?? []}
              excluded={excluded.includes(f.id)}
              onPickKey={(id) => dispatch({ type: 'modal/open', modal: { kind: 'keyDetail', tab: id } })}
              analyzing={analyzing}
              locked={phase === 'merged'}
              compact={phase === 'merged'}
              cellText={cellText}
              onSetBase={ctrl.setBase}
              onRemove={ctrl.removeFile}
              onSheet={ctrl.changeSheet}
              onHeaderRow={ctrl.changeHeaderRow}
              onLastRow={ctrl.changeLastDataRow}
              reorderable={phase !== 'merged' || !!fileGroups?.has(f.id)}
              active={activeFileId === f.id}
              picked={phase === 'files' && pickedFile === f.id}
              gestureHint={
                phase === 'merged'
                  ? fileGroups?.has(f.id)
                    ? 'クリック: このファイルの列を表の中央に表示して選択／ドラッグ: 列をまとめて移動（Alt+←/→ でも移動できます）'
                    : 'このファイルの列は表にありません'
                  : 'ドラッグで並べ替え（Alt+←/→ でも移動できます）'
              }
            />
          ))}
          <AddFileCard onPick={pickFiles} dragging={dragging} compact={phase === 'merged'} />
        </FileCardRow>
      )}

      {(files.length > 0 || state.toast) && (
        <SettingsBar summary={summaryLine} summaryTitle={summary ? '一致: すべてのファイルに該当する行がある基準の行' : undefined} toast={<Toast toast={state.toast} onClose={() => dispatch({ type: 'toast/clear' })} />}>
          {settings}
        </SettingsBar>
      )}
      {files.length === 0 && !state.toast && <div className="settings-spacer" aria-hidden />}

      <main className="main-area">
        <div className="table-column">
          <ErrorBoundary area={phase === 'merged' ? '表' : '画面'} resetKey={doc?.version ?? phase} onReset={phase === 'merged' ? ctrl.undo : undefined} resetLabel="直前の操作を元に戻す">
            {files.length === 0 ? (
              <EmptyState onPick={pickFiles} dragging={dragging} />
            ) : phase === 'merged' && merged ? (
              <EditorPane
                merged={merged}
                services={services}
                baseId={baseId}
                keyRoles={keyRoles}
                selection={state.selection}
                active={state.active}
                liveWidths={state.liveWidths}
                highlightRows={highlightRows}
                gridRef={gridRef}
                onFocusGrid={focusGrid}
                {...editorHandlers}
              />
            ) : (
              <MergePreview
                files={files}
                analysis={analysis}
                plan={plan}
                joinResult={joinResult}
                overrides={state.importOverrides}
                keyError={state.keyError}
                analyzing={analyzing}
                excluded={excluded}
                includeOtherSheets={state.includeOtherSheets}
                onImports={(changes) => dispatch({ type: 'imports/set', changes })}
                onOpenDetail={(tab) => dispatch({ type: 'modal/open', modal: { kind: 'keyDetail', tab: tab ?? null } })}
              />
            )}
          </ErrorBoundary>
        </div>

        <ErrorBoundary area="サイドパネル">
          <SidePanel
            phase={phase}
            primaryDisabled={!!mergeDisabledReason || !!state.busy}
            primaryTitle={mergeDisabledReason ?? undefined}
            onMerge={() => void ctrl.merge()}
            onExport={openExport}
            exportRunning={state.exportState.phase === 'running'}
            canUndo={!!merged && canUndo(merged.history)}
            canRedo={!!merged && canRedo(merged.history)}
            undoLabel={merged && canUndo(merged.history) ? merged.history.present.label : undefined}
            redoLabel={merged && canRedo(merged.history) ? merged.history.future[0]?.label : undefined}
            onUndo={ctrl.undo}
            onRedo={ctrl.redo}
            onRemerge={ctrl.requestRemerge}
            summary={sideSummary}
            formulas={phase === 'merged' && formulaViews.length > 0 ? <FormulaReport views={formulaViews} colors={colors} canJump onJump={jumpToFormulaCell} /> : undefined}
            checks={
              <CheckPanel
                groups={groups}
                canJump={phase === 'merged'}
                onJump={jumpToRow}
                emptyText={files.length ? '結合キーが決まると、未一致・重複などを表示します' : '―'}
              />
            }
            columns={
              phase === 'merged' ? (
                <ColumnList items={listItems} selected={state.selection.cols} onMove={ctrl.moveColumns} onSelect={onColumnSelect} onRename={ctrl.renameColumn} />
              ) : (
                <PlannedColumnList items={planned} colors={colors} />
              )
            }
            columnCount={phase === 'merged' ? listItems.length : planned.length}
            columnHelp={phase === 'merged' ? 'ドラッグまたは Alt+↑↓ で並べ替え／クリックで表の列へ移動／ダブルクリックで見出しを変更' : '結合後にドラッグで並べ替えできます'}
            notes={notes}
          />
        </ErrorBoundary>
      </main>

      <input
        ref={fileInputRef}
        type="file"
        multiple
        accept={ACCEPT_ATTR}
        className="visually-hidden"
        data-testid="file-input"
        aria-label="Excelファイルを選択"
        tabIndex={-1}
        onChange={(e) => {
          const list = e.target.files ? Array.from(e.target.files) : [];
          e.target.value = ''; // allow choosing the same file again
          if (list.length) void ctrl.addFiles(list);
        }}
      />

      {state.modal?.kind === 'keyDetail' && detailAnalysis && plan && (
        <KeyDetailModal
          analysis={detailAnalysis}
          plan={plan}
          joinResult={joinResult}
          labels={labels}
          colors={colors}
          options={state.options}
          overrides={state.importOverrides}
          initialTab={state.modal.tab}
          excluded={excluded}
          manualKeys={state.manualKeys}
          onAddManual={ctrl.addManualKey}
          onRemoveManual={(s, id) => void ctrl.removeManualKey(s, id)}
          onTab={(tab) => dispatch({ type: 'modal/open', modal: { kind: 'keyDetail', tab } })}
          onPlan={ctrl.updatePlan}
          onOptions={(o) => dispatch({ type: 'options/set', options: o })}
          onImports={(changes) => dispatch({ type: 'imports/set', changes })}
          onClose={closeModal}
        />
      )}
      {deleteModal && doc && (
        <DeleteConfirmDialog
          titles={deleteModal.ids.map((id) => listItems.find((c) => c.id === id)?.title ?? id)}
          impact={deleteModal.impact}
          onConvert={() => {
            ctrl.confirmDelete();
            requestAnimationFrame(() => gridRef.current?.focus());
          }}
          onCancel={closeModal}
        />
      )}
      {state.modal?.kind === 'export' && merged && (
        <ExportDialog
          defaultName={exportFileName(merged.baseFileName, merged.format)}
          format={merged.format}
          server={state.server}
          exportState={state.exportState}
          editCount={merged.history.past.length}
          onExport={(fileName, engine) => void ctrl.runExport({ fileName, engine })}
          onCancelExport={ctrl.cancelExport}
          onOpenFile={() => void ctrl.openSavedFile()}
          onOpenFolder={() => void ctrl.openOutputFolder()}
          onDownloadAgain={ctrl.downloadAgain}
          onClose={closeModal}
        />
      )}
      {state.modal?.kind === 'help' && <HelpDialog onClose={closeModal} />}
      {state.modal?.kind === 'confirm' && (
        <ConfirmDialog
          title={state.modal.title}
          message={state.modal.message}
          okLabel={state.modal.okLabel}
          danger={!!state.modal.danger}
          testId={state.modal.purpose === 'unmergeMove' ? 'unmerge-confirm-dialog' : 'confirm-dialog'}
          onOk={() => {
            const purpose = stateRef.current.modal?.kind === 'confirm' ? stateRef.current.modal.purpose : null;
            ctrl.confirmModal();
            if (purpose === 'unmergeMove') requestAnimationFrame(() => gridRef.current?.focus());
          }}
          onCancel={closeModal}
        />
      )}

      <BusyOverlay busy={state.busy} />
      {dragging && !state.modal && (
        <div className="drop-overlay" aria-hidden>
          <div className="drop-overlay-card">ここにドロップしてファイルを追加</div>
        </div>
      )}
    </div>
  );
}
