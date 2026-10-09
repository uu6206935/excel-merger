/**
 * Source bands of the grid header (pure, no React / DOM). Unit-tested in tests/grid/runs.test.ts.
 *
 * A "run" is a maximal sequence of CONTIGUOUS columns with the same source (GridColumn.source; -1 / undefined =
 * columns added in the app). Runs are computed from the complete `columns` array, never from the virtualised
 * (visible) columns, so `count` is always the true number of columns of the group. Runs are logical: the frozen pane
 * does not split them (a band is only DRAWN in two pieces when the frozen boundary lies inside a run, see bandPieces).
 */

export interface BandRun {
  /** index of the first column (0-based, current order) */
  s: number;
  /** index of the last column (inclusive) */
  e: number;
  /** index into GridProps.sources, -1 = added in the app */
  source: number;
  /** true number of columns of the run (e - s + 1) */
  count: number;
}

/** Maximal runs of contiguous columns that share a source. */
export function computeRuns(sources: ReadonlyArray<number | undefined>): BandRun[] {
  const out: BandRun[] = [];
  for (let i = 0; i < sources.length; i++) {
    const src = sources[i] ?? -1;
    const last = out[out.length - 1];
    if (last && last.source === src) {
      last.e = i;
      last.count++;
    } else {
      out.push({ s: i, e: i, source: src, count: 1 });
    }
  }
  return out;
}

/** The part of a run that is drawn as one band element. */
export interface BandPiece {
  /** index into the runs array */
  run: number;
  s: number;
  e: number;
  /** drawn in the frozen (sticky) pane */
  pinned: boolean;
  /** the run continues on the other side of the frozen boundary */
  split: boolean;
}

/** Splits the runs at the frozen boundary (columns [0, frozen) are pinned). Every run keeps its own index. */
export function bandPieces(runs: readonly BandRun[], frozen: number): BandPiece[] {
  const out: BandPiece[] = [];
  runs.forEach((r, i) => {
    if (frozen > 0 && r.s < frozen && r.e >= frozen) {
      out.push({ run: i, s: r.s, e: frozen - 1, pinned: true, split: true });
      out.push({ run: i, s: frozen, e: r.e, pinned: false, split: true });
    } else {
      out.push({ run: i, s: r.s, e: r.e, pinned: r.e < frozen, split: false });
    }
  });
  return out;
}

/** Ids of the columns of a run, in current order. */
export function runColumnIds(order: readonly string[], run: Pick<BandRun, 's' | 'e'>): string[] {
  return order.slice(run.s, run.e + 1);
}

/** True when every column of the run is selected. */
export function runFullySelected(order: readonly string[], run: Pick<BandRun, 's' | 'e'>, cols: ReadonlySet<string>): boolean {
  if (cols.size < run.e - run.s + 1) return false;
  for (let i = run.s; i <= run.e; i++) if (!cols.has(order[i])) return false;
  return true;
}
