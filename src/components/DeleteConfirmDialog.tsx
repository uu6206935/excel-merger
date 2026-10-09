/**
 * Shown before deleting columns that formulas depend on: never produce #REF! — the only way forward is
 * 「参照している数式を値に変換して削除」 (or cancel). Also lists VLOOKUP-style warnings.
 */
import type { DeleteImpact } from '../core/contracts';
import { Modal } from './Modal';

/** examples listed in the dialog (the doc may give 50; the rest is 「ほか N 件」) */
const MAX_LISTED = 10;

export interface DeleteConfirmDialogProps {
  titles: string[];
  impact: DeleteImpact;
  onConvert: () => void;
  onCancel: () => void;
}

export function DeleteConfirmDialog({ titles, impact, onConvert, onCancel }: DeleteConfirmDialogProps) {
  const n = titles.length;
  const shown = titles.slice(0, 5).map((t) => `「${t}」`).join('');
  const breaks = impact.breakCount > 0;
  return (
    <Modal
      title={breaks ? 'この列を参照している数式があります' : '列を削除します'}
      testId="delete-confirm-dialog"
      size="md"
      onClose={onCancel}
      initialFocus="[data-testid=delete-convert-button]"
      footer={
        <>
          <button type="button" className="btn" data-testid="delete-cancel-button" onClick={onCancel}>
            キャンセル
          </button>
          <button type="button" className="btn btn-danger" data-testid="delete-convert-button" onClick={onConvert}>
            {breaks ? '参照している数式を値に変換して削除' : '削除する'}
          </button>
        </>
      }
    >
      <p className="modal-text">
        {shown}
        {n > 5 ? ` ほか${n - 5}列` : ''}（{n} 列）を削除します。
        {breaks && (
          <>
            <br />
            このままでは <b>{impact.breakCount.toLocaleString()} 件</b>の数式が <code>#REF!</code> エラーになるため、それらの数式は<b>今の計算結果の値</b>に置き換えてから削除します。
          </>
        )}
      </p>
      {breaks && (
        <ul className="side-list delete-list">
          {impact.breaks.slice(0, MAX_LISTED).map((b, i) => (
            <li key={i}>
              <code className="where">{b.where}</code>
              <span className="formula">{b.text}</span>
            </li>
          ))}
          {impact.breakCount > Math.min(impact.breaks.length, MAX_LISTED) && (
            <li className="muted">ほか {(impact.breakCount - Math.min(impact.breaks.length, MAX_LISTED)).toLocaleString()} 件</li>
          )}
        </ul>
      )}
      {impact.lookupWarnings.length > 0 && (
        <div className="card-note">
          {impact.lookupWarnings.map((w, i) => (
            <div key={i}>{w}</div>
          ))}
        </div>
      )}
      <p className="modal-note muted">削除した後でも「元に戻す」（Ctrl+Z）で戻せます。</p>
    </Modal>
  );
}
