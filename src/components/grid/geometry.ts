/**
 * Pure layout / navigation helpers of the grid (no React, no DOM). Unit-tested in tests/grid/geometry.test.ts.
 */

/** starts[i] = x of item i, starts[n] = total. */
export function prefixSums(sizes: ArrayLike<number>): Float64Array {
  const n = sizes.length;
  const a = new Float64Array(n + 1);
  for (let i = 0; i < n; i++) a[i + 1] = a[i] + sizes[i];
  return a;
}

/**
 * Insertion gap under position `x` (data-area coordinates) restricted to items [lo, hi):
 * the first index i in [lo, hi) whose centre is right of x, else hi. Binary search over prefix sums.
 */
export function gapAt(x: number, starts: ArrayLike<number>, lo: number, hi: number): number {
  while (lo < hi) {
    const mid = (lo + hi) >> 1;
    const centre = (starts[mid] + starts[mid + 1]) / 2;
    if (x < centre) hi = mid;
    else lo = mid + 1;
  }
  return lo;
}

/**
 * scrollLeft that puts the middle of the range [x0, x1] (canvas coordinates) into the middle of the visible, scrolling
 * area [viewLeft, viewRight] (viewport coordinates: the row-number gutter and the frozen pane are excluded). A range that
 * is wider than the area is aligned to the area's left edge (with a small margin) so its beginning stays visible.
 * The result is clamped to [0, maxScroll].
 */
export function centerScrollLeft(x0: number, x1: number, viewLeft: number, viewRight: number, maxScroll: number, margin = 12): number {
  const viewW = Math.max(0, viewRight - viewLeft);
  const w = Math.max(0, x1 - x0);
  const target = w + 2 * margin > viewW ? x0 - viewLeft - margin : (x0 + x1) / 2 - (viewLeft + viewRight) / 2;
  return Math.max(0, Math.min(Math.max(0, maxScroll), Math.round(target)));
}

/** Index of the item containing position x (clamped to [0, n-1]); -1 when n = 0. */
export function indexAt(x: number, starts: ArrayLike<number>): number {
  const n = starts.length - 1;
  if (n <= 0) return -1;
  let lo = 0;
  let hi = n - 1;
  while (lo < hi) {
    const mid = (lo + hi + 1) >> 1;
    if (starts[mid] <= x) lo = mid;
    else hi = mid - 1;
  }
  return lo;
}

/** Excel row number -> index into `rows` (-1 when not shown). O(1) for contiguous rows, Map otherwise. */
export function makeRowIndex(rows: readonly number[]): (row: number) => number {
  const n = rows.length;
  if (n === 0) return () => -1;
  const first = rows[0];
  let contiguous = true;
  for (let i = 1; i < n; i++) {
    if (rows[i] !== first + i) {
      contiguous = false;
      break;
    }
  }
  if (contiguous) return (row) => (Number.isInteger(row) && row >= first && row < first + n ? row - first : -1);
  const m = new Map<number, number>();
  for (let i = 0; i < n; i++) if (!m.has(rows[i])) m.set(rows[i], i);
  return (row) => m.get(row) ?? -1;
}

export type NavKey = 'ArrowDown' | 'ArrowUp' | 'ArrowLeft' | 'ArrowRight' | 'Tab' | 'Enter' | 'Home' | 'End' | 'PageDown' | 'PageUp';

export function isNavKey(k: string): k is NavKey {
  return (
    k === 'ArrowDown' || k === 'ArrowUp' || k === 'ArrowLeft' || k === 'ArrowRight' || k === 'Tab' || k === 'Enter' ||
    k === 'Home' || k === 'End' || k === 'PageDown' || k === 'PageUp'
  );
}

/**
 * Active-cell navigation (Excel-like, without "End mode"):
 *   arrows ±1 (Ctrl: to the edge), Tab/Shift+Tab ±1 column, Enter/Shift+Enter ±1 row,
 *   Home = first column (Ctrl: first cell), End = last column (Ctrl: last cell), PageDown/PageUp ±page rows.
 * Returns clamped indexes, or null when there is nothing to navigate.
 */
export function navigate(
  key: NavKey,
  mods: { ctrl: boolean; shift: boolean },
  at: { rowIdx: number; colIdx: number },
  size: { rows: number; cols: number; page: number },
): { rowIdx: number; colIdx: number } | null {
  if (size.rows <= 0 || size.cols <= 0) return null;
  let r = at.rowIdx;
  let c = at.colIdx;
  const page = Math.max(1, size.page);
  switch (key) {
    case 'ArrowDown':
      r = mods.ctrl ? size.rows - 1 : r + 1;
      break;
    case 'ArrowUp':
      r = mods.ctrl ? 0 : r - 1;
      break;
    case 'ArrowRight':
      c = mods.ctrl ? size.cols - 1 : c + 1;
      break;
    case 'ArrowLeft':
      c = mods.ctrl ? 0 : c - 1;
      break;
    case 'Tab':
      c += mods.shift ? -1 : 1;
      break;
    case 'Enter':
      r += mods.shift ? -1 : 1;
      break;
    case 'Home':
      c = 0;
      if (mods.ctrl) r = 0;
      break;
    case 'End':
      c = size.cols - 1;
      if (mods.ctrl) r = size.rows - 1;
      break;
    case 'PageDown':
      r += page;
      break;
    case 'PageUp':
      r -= page;
      break;
  }
  return { rowIdx: Math.max(0, Math.min(size.rows - 1, r)), colIdx: Math.max(0, Math.min(size.cols - 1, c)) };
}

/**
 * Where 「左に列を挿入」/「右に列を挿入」 inserts, Excel style: when the target columns form ONE contiguous run
 * that many columns are inserted before (left) / after (right) the run; otherwise one column next to the
 * clicked column. `index` = 0-based insert-before position in the current order.
 */
export function insertTarget(
  order: readonly string[],
  targets: readonly string[],
  clicked: string,
  side: 'left' | 'right',
): { index: number; count: number } {
  const ci = order.indexOf(clicked);
  const idx = targets.map((id) => order.indexOf(id)).filter((i) => i >= 0).sort((a, b) => a - b);
  const contiguous = idx.length > 0 && idx[idx.length - 1] - idx[0] === idx.length - 1 && idx.includes(ci);
  if (contiguous) return side === 'left' ? { index: idx[0], count: idx.length } : { index: idx[idx.length - 1] + 1, count: idx.length };
  const at = ci >= 0 ? ci : order.length;
  return side === 'left' ? { index: at, count: 1 } : { index: Math.min(order.length, at + 1), count: 1 };
}
