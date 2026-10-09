/**
 * Progress overlay for heavy work (reading, analysis, merge). Blocks input while visible.
 */
import type { BusyState } from '../state/types';

export function ProgressBar({ ratio }: { ratio: number | null }) {
  const pct = ratio === null ? null : Math.round(Math.max(0, Math.min(1, ratio)) * 100);
  return (
    <div
      className={`progress${pct === null ? ' progress-indeterminate' : ''}`}
      role="progressbar"
      aria-valuemin={0}
      aria-valuemax={100}
      aria-valuenow={pct ?? undefined}
    >
      <div className="progress-fill" style={pct === null ? undefined : { width: `${pct}%` }} />
    </div>
  );
}

export function BusyOverlay({ busy }: { busy: BusyState | null }) {
  if (!busy) return null;
  return (
    <div className="busy-backdrop" data-testid="busy-overlay" aria-busy="true">
      <div className="busy-card" role="status" aria-live="polite">
        <div className="spinner" aria-hidden />
        <div className="busy-text">{busy.message}</div>
        <ProgressBar ratio={busy.ratio} />
      </div>
    </div>
  );
}
