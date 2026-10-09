/**
 * Column selection model (pure). Unit-tested in tests/grid/selection.test.ts.
 */
import { rangeBetween } from './colops';

export interface ColSelection {
  cols: ReadonlySet<string>;
  /** anchor for Shift+click ranges (last plain / Ctrl click) */
  anchor: string | null;
}

export const EMPTY_SEL: ColSelection = Object.freeze({ cols: new Set<string>(), anchor: null }) as ColSelection;

/**
 * Header click semantics (same as Excel / Explorer):
 *   click        -> only this column, anchor = it
 *   Ctrl+click   -> toggle this column (non-contiguous), anchor = it
 *   Shift+click  -> range anchor..this (replaces), anchor unchanged
 *   Ctrl+Shift   -> range anchor..this added to the existing selection
 */
export function clickSelect(order: readonly string[], sel: ColSelection, id: string, mods: { ctrl: boolean; shift: boolean }): ColSelection {
  if (mods.shift && sel.anchor && order.includes(sel.anchor)) {
    const range = rangeBetween(order, sel.anchor, id);
    const cols = mods.ctrl ? new Set([...sel.cols, ...range]) : new Set(range);
    return { cols, anchor: sel.anchor };
  }
  if (mods.ctrl) {
    const cols = new Set(sel.cols);
    if (cols.has(id)) cols.delete(id);
    else cols.add(id);
    return { cols, anchor: id };
  }
  return { cols: new Set([id]), anchor: id };
}

/**
 * Click on a source band (a run of columns = one file's group): the same modifier semantics as a header click, applied
 * to the whole run.
 *   click        -> exactly the run's columns, anchor = its first column
 *   Ctrl+click   -> toggle the run: all its columns are removed when they are all selected, else all are added; anchor = first
 *   Shift+click  -> range from the anchor to the run (the span covering the anchor and the whole run), anchor unchanged
 *   Ctrl+Shift   -> that range added to the existing selection
 * `ids` = the run's columns in current order (non-empty).
 */
export function selectRun(order: readonly string[], sel: ColSelection, ids: readonly string[], mods: { ctrl: boolean; shift: boolean }): ColSelection {
  if (ids.length === 0) return sel;
  const first = ids[0];
  if (mods.shift && sel.anchor && order.includes(sel.anchor)) {
    const a = order.indexOf(sel.anchor);
    const lo = Math.min(a, order.indexOf(first));
    const hi = Math.max(a, order.indexOf(ids[ids.length - 1]));
    const range = order.slice(lo, hi + 1);
    return { cols: mods.ctrl ? new Set([...sel.cols, ...range]) : new Set(range), anchor: sel.anchor };
  }
  if (mods.ctrl) {
    const cols = new Set(sel.cols);
    if (ids.every((id) => cols.has(id))) for (const id of ids) cols.delete(id);
    else for (const id of ids) cols.add(id);
    return { cols, anchor: first };
  }
  return { cols: new Set(ids), anchor: first };
}

/** Selected ids in current display order (ids that no longer exist are dropped). */
export function orderedSelection(order: readonly string[], cols: ReadonlySet<string>): string[] {
  if (cols.size === 0) return [];
  return order.filter((id) => cols.has(id));
}

/** Selection of a single column. */
export function selectOnly(id: string): ColSelection {
  return { cols: new Set([id]), anchor: id };
}

/** Shift+Arrow with a column selection: the range anchor..`to` (anchor kept). */
export function extendTo(order: readonly string[], sel: ColSelection, to: string): ColSelection {
  const anchor = sel.anchor && order.includes(sel.anchor) ? sel.anchor : to;
  return { cols: new Set(rangeBetween(order, anchor, to)), anchor };
}

/** Same members (order-insensitive). */
export function sameSelection(a: ColSelection, b: ColSelection): boolean {
  if (a.anchor !== b.anchor || a.cols.size !== b.cols.size) return false;
  for (const id of a.cols) if (!b.cols.has(id)) return false;
  return true;
}
