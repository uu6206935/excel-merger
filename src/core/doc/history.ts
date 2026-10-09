// Undo/redo over immutable snapshots (owner: doc D; this minimal version is part of the contract).
import type { EditorSnapshot } from '../types';
import type { History } from '../contracts';

const LIMIT = 200;

export function initHistory(s: EditorSnapshot): History {
  return { past: [], present: s, future: [] };
}

export function commit(h: History, s: EditorSnapshot): History {
  const past = [...h.past, h.present];
  if (past.length > LIMIT) past.shift();
  return { past, present: s, future: [] };
}

export function undo(h: History): History {
  if (h.past.length === 0) return h;
  const prev = h.past[h.past.length - 1];
  return { past: h.past.slice(0, -1), present: prev, future: [h.present, ...h.future] };
}

export function redo(h: History): History {
  if (h.future.length === 0) return h;
  const [next, ...rest] = h.future;
  return { past: [...h.past, h.present], present: next, future: rest };
}

export const canUndo = (h: History): boolean => h.past.length > 0;
export const canRedo = (h: History): boolean => h.future.length > 0;
