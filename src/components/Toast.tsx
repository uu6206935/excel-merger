/**
 * Toast (inline at the right end of the settings bar, like the reference app). A short text; the longer
 * explanation (`detail`) is the tooltip (「詳しく」).
 */
import type { ToastState } from '../state/types';

export function Toast({ toast, onClose }: { toast: ToastState | null; onClose: () => void }) {
  if (!toast) return null;
  return (
    <div
      className={`toast${toast.kind === 'error' ? ' toast-error' : ''}`}
      role={toast.kind === 'error' ? 'alert' : 'status'}
      aria-live={toast.kind === 'error' ? 'assertive' : 'polite'}
      data-testid="toast"
      data-kind={toast.kind}
      title={toast.detail ?? toast.text}
    >
      <span className="toast-text">{toast.text}</span>
      {toast.detail && (
        <>
          <span className="toast-more" aria-hidden>
            詳しく
          </span>
          <span className="visually-hidden">{toast.detail}</span>
        </>
      )}
      <button type="button" className="toast-close" onClick={onClose} aria-label="閉じる">
        ×
      </button>
    </div>
  );
}
