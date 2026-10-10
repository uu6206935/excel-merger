/**
 * Right-hand panel (280px, like the reference): primary button (結合する / Excel出力), 元に戻す / やり直す,
 * 結合をやり直す, summary box, チェック結果, 列の一覧 and the collapsible 警告 section.
 */
import type { ReactNode } from 'react';
import type { Note } from '../state/notes';
import { noteCounts } from '../state/notes';

export interface SidePanelProps {
  phase: 'files' | 'merged';
  primaryDisabled: boolean;
  primaryTitle?: string;
  onMerge: () => void;
  onExport: () => void;
  exportRunning: boolean;
  canUndo: boolean;
  canRedo: boolean;
  undoLabel?: string;
  redoLabel?: string;
  onUndo: () => void;
  onRedo: () => void;
  onRemerge: () => void;
  summary: ReactNode;
  /** 「シート一覧」 (after the merge, when the output has more than one sheet; absent = nothing to show) */
  sheets?: ReactNode;
  /** 「数式の取り込み」 block (after the merge; absent = nothing to show) */
  formulas?: ReactNode;
  checks: ReactNode;
  columns: ReactNode;
  columnCount: number;
  columnHelp: string;
  notes: Note[];
}

export function SidePanel(props: SidePanelProps) {
  const { phase, notes } = props;
  const counts = noteCounts(notes);
  const merged = phase === 'merged';
  return (
    <aside className="side-panel" aria-label="操作と状況" data-testid="side-panel">
      <div className="side-actions">
        {merged ? (
          <button
            type="button"
            className="btn btn-primary btn-autofill"
            data-testid="export-button"
            onClick={props.onExport}
            aria-haspopup="dialog"
            title="結合した表を、数式・書式・列幅を保ったまま .xlsx に出力します（他のシートは、設定で含めたときだけ出力されます）"
          >
            {props.exportRunning ? '出力中…' : 'Excel出力'}
          </button>
        ) : (
          <button
            type="button"
            className="btn btn-primary btn-autofill"
            data-testid="merge-button"
            onClick={props.onMerge}
            disabled={props.primaryDisabled}
            title={props.primaryTitle ?? '基準ファイルの表に、他のファイルの列を追加します'}
          >
            結合する
          </button>
        )}
        <div className="btn-row">
          <button
            type="button"
            className="btn"
            data-testid="undo-button"
            onClick={props.onUndo}
            disabled={!props.canUndo}
            aria-label="元に戻す"
            title={props.undoLabel ? `元に戻す: ${props.undoLabel} (Ctrl+Z)` : '元に戻す (Ctrl+Z)'}
          >
            ↶ 元に戻す
          </button>
          <button
            type="button"
            className="btn"
            data-testid="redo-button"
            onClick={props.onRedo}
            disabled={!props.canRedo}
            aria-label="やり直す"
            title={props.redoLabel ? `やり直す: ${props.redoLabel} (Ctrl+Y)` : 'やり直す (Ctrl+Y)'}
          >
            ↷ やり直す
          </button>
        </div>
        {merged && (
          <button type="button" className="btn btn-sm remerge-btn" data-testid="remerge-button" onClick={props.onRemerge} title="結合キーの画面に戻ります（結合後の変更は失われます）">
            ← 結合をやり直す
          </button>
        )}
      </div>

      <div className="side-summary" data-testid="side-summary">
        {props.summary}
      </div>

      {props.sheets}

      {props.formulas}

      {props.checks}

      <section className="side-section column-section" aria-label="列の一覧">
        <h3>
          列の一覧 <span className="count-badge">{props.columnCount}</span>
        </h3>
        <p className="side-help">{props.columnHelp}</p>
        <div className="column-list-wrap" data-testid="column-list">
          {props.columns}
        </div>
      </section>

      <section className="side-section warning-section">
        <details className="side-details" open={counts.error > 0} data-testid="warning-details">
          <summary title="出力したときに保てないもの・確認してほしい点の一覧です">
            警告 <span className={`count-badge${counts.error > 0 ? ' badge-error' : counts.warning > 0 ? ' badge-warn' : ''}`}>{notes.length}</span>
          </summary>
          <ul className="side-list warning-list" data-testid="warning-list">
            {notes.length === 0 && <li className="muted">なし</li>}
            {notes.map((n) => (
              <li key={n.key} className={`sev-${n.severity}`}>
                {n.text}
              </li>
            ))}
          </ul>
        </details>
      </section>
    </aside>
  );
}
