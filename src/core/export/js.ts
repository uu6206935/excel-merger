// JS export engine = package patcher replay (owner: patcher P).
import type { LoggedOp } from '../types';
import type { ProgressFn } from '../contracts';
import { replayLog, type ReplayWarning } from '../xlsx/replay';

/** `onWarning` (optional): UNSAFE_FORMULA notes — call-out formulas were written as text (the Excel engine reports the same warning). */
export function exportWithJs(premerged: Uint8Array, log: LoggedOp[], onProgress?: ProgressFn, onWarning?: (w: ReplayWarning) => void): Promise<Uint8Array> {
  return replayLog(premerged, log, onProgress, onWarning);
}
