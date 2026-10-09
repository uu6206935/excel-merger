/**
 * Pure column-order helpers (no React). Unit-tested in tests/grid/colops.test.ts.
 * NOTE: the doc module decides the real decomposition of a move (excelRef.planMove); computeMove here is the
 * grid-side preview / no-op check and planBlockMoves is kept for tests and tooling.
 *
 * Coordinates: a "gap" is an insertion slot in the CURRENT order, 0..n
 * (gap g = "before the column that is now at index g"; gap n = after the last column).
 * This is what the drag indicator shows, and it is expressed in pre-move coordinates,
 * exactly like Excel's "Insert Cut Cells" target.
 */

/**
 * Final order after moving every id in `moving` (non-contiguous allowed) to `gap`.
 * The moved ids keep their relative order; everything else keeps its relative order.
 */
export function computeMove<T>(order: readonly T[], moving: ReadonlySet<T>, gap: number): T[] {
  const g = Math.max(0, Math.min(order.length, gap));
  const moved: T[] = [];
  const rest: T[] = [];
  let insertAt = 0;
  for (let i = 0; i < order.length; i++) {
    const id = order[i];
    if (moving.has(id)) moved.push(id);
    else {
      rest.push(id);
      if (i < g) insertAt++;
    }
  }
  return [...rest.slice(0, insertAt), ...moved, ...rest.slice(insertAt)];
}

/** true when computeMove would not change anything (used to avoid no-op undo entries). */
export function isNoopMove<T>(order: readonly T[], moving: ReadonlySet<T>, gap: number): boolean {
  const next = computeMove(order, moving, gap);
  for (let i = 0; i < order.length; i++) if (next[i] !== order[i]) return false;
  return true;
}

/**
 * One Excel-native step: cut `count` contiguous columns starting at 0-based `from`
 * and "Insert Cut Cells" before 0-based column `to` (to is measured BEFORE the cut
 * columns are removed, like Excel; to === n means "after the last column").
 * In Excel COM (1-based):  ws.Range(ws.Columns(from+1), ws.Columns(from+count)).Cut();
 *                          ws.Columns(to+1).Insert(xlShiftToRight = -4161)
 */
export interface BlockMove {
  from: number;
  count: number;
  to: number;
}

/** Apply one BlockMove to an array (returns a new array). */
export function applyBlockMove<T>(order: readonly T[], m: BlockMove): T[] {
  const cur = order.slice();
  const block = cur.splice(m.from, m.count);
  const insertAt = m.to > m.from ? m.to - m.count : m.to;
  cur.splice(insertAt, 0, ...block);
  return cur;
}

/**
 * Decompose a (possibly non-contiguous) multi-column move into a short sequence of
 * contiguous BlockMoves. Excel refuses to Cut a multi-area selection, so the
 * "Excel本体" export engine must replay a non-contiguous move as several block moves.
 *
 * Strategy: every moved run is inserted immediately before an anchor (the first
 * non-moving column at/after the gap, or the end). Runs are processed left-to-right so
 * their relative order is preserved. Runs already sitting in place directly before
 * the anchor (a trailing suffix of the moved set) are left alone and become the anchor.
 */
export function planBlockMoves<T>(order: readonly T[], moving: ReadonlySet<T>, gap: number): { final: T[]; steps: BlockMove[] } {
  const final = computeMove(order, moving, gap);
  const mv = order.filter((id) => moving.has(id)); // moved ids in (unchanged) relative order
  if (mv.length === 0) return { final, steps: [] };
  const g = Math.max(0, Math.min(order.length, gap));
  let anchor: T | null = null;
  for (let i = g; i < order.length; i++) {
    if (!moving.has(order[i])) {
      anchor = order[i];
      break;
    }
  }
  let cur = order.slice();
  // Trailing suffix of mv already contiguous right before the anchor -> keep in place.
  const anchorIdx = anchor === null ? cur.length : cur.indexOf(anchor);
  let s = mv.length;
  while (s > 0) {
    const pos = anchorIdx - (mv.length - s) - 1;
    if (pos >= 0 && cur[pos] === mv[s - 1]) s--;
    else break;
  }
  const effAnchor: T | null = s < mv.length ? mv[s] : anchor;
  const steps: BlockMove[] = [];
  let j = 0;
  while (j < s) {
    const from = cur.indexOf(mv[j]);
    let count = 1;
    while (j + count < s && cur[from + count] === mv[j + count]) count++;
    const to = effAnchor === null ? cur.length : cur.indexOf(effAnchor);
    if (from + count !== to) {
      const step = { from, count, to };
      steps.push(step);
      cur = applyBlockMove(cur, step);
    }
    j += count;
  }
  // Safety net: the plan must reproduce computeMove exactly.
  for (let i = 0; i < final.length; i++) {
    if (cur[i] !== final[i]) throw new Error('planBlockMoves: internal mismatch');
  }
  return { final, steps };
}

/** 0-based column index -> Excel letters (0 -> A, 25 -> Z, 26 -> AA). */
export function colLetter(index: number): string {
  let n = index + 1;
  let s = '';
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/** Excel letters -> 0-based column index ('A' -> 0, 'AA' -> 26); -1 for invalid input. */
export function letterToIndex(letters: string): number {
  if (!/^[A-Za-z]{1,3}$/.test(letters)) return -1;
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

/** A1 address of a cell from its 0-based column index and Excel row ("C5"); '' when unknown. */
export function cellAddress(colIndex: number, row: number): string {
  if (!(colIndex >= 0) || !(row >= 1)) return '';
  return `${colLetter(colIndex)}${row}`;
}

/** Inclusive range of ids between a and b in `order` (for Shift+click). */
export function rangeBetween<T>(order: readonly T[], a: T, b: T): T[] {
  const i = order.indexOf(a);
  const j = order.indexOf(b);
  if (i < 0 || j < 0) return j >= 0 ? [b] : [];
  const lo = Math.min(i, j);
  const hi = Math.max(i, j);
  return order.slice(lo, hi + 1);
}

/** Group consecutive indexes into runs, e.g. [1,2,3,7,9,10] -> [[1,3],[7,7],[9,10]]. */
export function toRuns(indexes: number[]): Array<[number, number]> {
  const s = [...indexes].sort((a, b) => a - b);
  const runs: Array<[number, number]> = [];
  for (const i of s) {
    const last = runs[runs.length - 1];
    if (last && i === last[1] + 1) last[1] = i;
    else runs.push([i, i]);
  }
  return runs;
}
