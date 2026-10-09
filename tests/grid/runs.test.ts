/**
 * Source bands of the grid header (round 7): runs and their TRUE counts, the pieces drawn around the frozen pane, the band
 * click selection (selectRun), the "centre this group" scroll position and the drop rules of a run drag.
 */
import { describe, expect, it } from 'vitest';
import { computeMove, isNoopMove } from '../../src/components/grid/colops';
import { centerScrollLeft, prefixSums } from '../../src/components/grid/geometry';
import { bandPieces, computeRuns, runColumnIds, runFullySelected } from '../../src/components/grid/runs';
import { clickSelect, EMPTY_SEL, orderedSelection, selectRun, type ColSelection } from '../../src/components/grid/selection';

const none = { ctrl: false, shift: false };
const ctrl = { ctrl: true, shift: false };
const shift = { ctrl: false, shift: true };
const both = { ctrl: true, shift: true };

describe('computeRuns', () => {
  it('is the maximal contiguous runs of the same source, with the true count of each', () => {
    const runs = computeRuns([0, 0, 1, 1, 1, 2, 2, -1, 0]);
    expect(runs).toEqual([
      { s: 0, e: 1, source: 0, count: 2 },
      { s: 2, e: 4, source: 1, count: 3 },
      { s: 5, e: 6, source: 2, count: 2 },
      { s: 7, e: 7, source: -1, count: 1 },
      { s: 8, e: 8, source: 0, count: 1 },
    ]);
    // the same source in two places is two runs (a file whose columns were split by moves)
    expect(runs.filter((r) => r.source === 0)).toHaveLength(2);
  });

  it('undefined source = columns added in the app (-1), and an empty table has no run', () => {
    expect(computeRuns([undefined, -1, undefined, 0])).toEqual([
      { s: 0, e: 2, source: -1, count: 3 },
      { s: 3, e: 3, source: 0, count: 1 },
    ]);
    expect(computeRuns([])).toEqual([]);
  });

  it('counts every column of a long run (not a virtualised window)', () => {
    const sources = [...Array(1200).fill(0), ...Array(800).fill(1)];
    const runs = computeRuns(sources);
    expect(runs.map((r) => r.count)).toEqual([1200, 800]);
    expect(runs[1]).toMatchObject({ s: 1200, e: 1999 });
  });
});

describe('bandPieces (drawing around the frozen pane)', () => {
  it('a run that crosses the frozen boundary is drawn in two pieces that both belong to the same run', () => {
    const runs = computeRuns([0, 0, 0, 0, 0, 1, 1]);
    const pieces = bandPieces(runs, 2);
    expect(pieces).toEqual([
      { run: 0, s: 0, e: 1, pinned: true, split: true },
      { run: 0, s: 2, e: 4, pinned: false, split: true },
      { run: 1, s: 5, e: 6, pinned: false, split: false },
    ]);
    // both pieces carry the logical run: its count is the TRUE 5, not 2 + 3
    expect(runs[pieces[0].run].count).toBe(5);
    expect(runs[pieces[1].run].count).toBe(5);
  });

  it('no frozen columns / a run entirely inside or outside the pane is one piece', () => {
    const runs = computeRuns([0, 0, 1, 1]);
    expect(bandPieces(runs, 0).map((p) => [p.pinned, p.split])).toEqual([[false, false], [false, false]]);
    expect(bandPieces(runs, 2).map((p) => [p.pinned, p.split])).toEqual([[true, false], [false, false]]);
  });
});

describe('run helpers', () => {
  const order = ['a', 'b', 'c', 'd', 'e', 'f'];
  it('runColumnIds / runFullySelected', () => {
    expect(runColumnIds(order, { s: 2, e: 4 })).toEqual(['c', 'd', 'e']);
    expect(runFullySelected(order, { s: 2, e: 4 }, new Set(['c', 'd', 'e']))).toBe(true);
    expect(runFullySelected(order, { s: 2, e: 4 }, new Set(['c', 'd']))).toBe(false);
    expect(runFullySelected(order, { s: 2, e: 4 }, new Set(['a', 'c', 'd', 'e']))).toBe(true);
  });
});

describe('selectRun (band click)', () => {
  const order = ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h'];
  const ids = (s: ColSelection) => orderedSelection(order, s.cols);
  const run1 = ['c', 'd', 'e'];
  const run2 = ['g', 'h'];

  it('click selects exactly the run, anchor = its first column', () => {
    const s = selectRun(order, clickSelect(order, EMPTY_SEL, 'a', none), run1, none);
    expect(ids(s)).toEqual(['c', 'd', 'e']);
    expect(s.anchor).toBe('c');
  });

  it('Ctrl+click adds the run, and toggles it off when all its columns are selected', () => {
    let s = clickSelect(order, EMPTY_SEL, 'a', none);
    s = selectRun(order, s, run1, ctrl);
    expect(ids(s)).toEqual(['a', 'c', 'd', 'e']);
    expect(s.anchor).toBe('c');
    s = selectRun(order, s, run2, ctrl);
    expect(ids(s)).toEqual(['a', 'c', 'd', 'e', 'g', 'h']);
    s = selectRun(order, s, run1, ctrl);
    expect(ids(s)).toEqual(['a', 'g', 'h']);
    // a partly selected run is completed, not toggled off
    let p = clickSelect(order, EMPTY_SEL, 'd', none);
    p = selectRun(order, p, run1, ctrl);
    expect(ids(p)).toEqual(['c', 'd', 'e']);
  });

  it('Shift+click extends from the anchor to the whole run (both directions), the anchor stays', () => {
    const base = clickSelect(order, EMPTY_SEL, 'b', none);
    let s = selectRun(order, base, run1, shift);
    expect(ids(s)).toEqual(['b', 'c', 'd', 'e']);
    expect(s.anchor).toBe('b');
    s = selectRun(order, base, run2, shift);
    expect(ids(s)).toEqual(['b', 'c', 'd', 'e', 'f', 'g', 'h']);
    const right = clickSelect(order, EMPTY_SEL, 'h', none);
    s = selectRun(order, right, run1, shift);
    expect(ids(s)).toEqual(['c', 'd', 'e', 'f', 'g', 'h']);
    expect(s.anchor).toBe('h');
    // the anchor inside the run: just the run
    const inside = clickSelect(order, EMPTY_SEL, 'd', none);
    expect(ids(selectRun(order, inside, run1, shift))).toEqual(['c', 'd', 'e']);
  });

  it('Ctrl+Shift+click adds the range to the existing selection; Shift without an anchor is a plain click', () => {
    let s = clickSelect(order, EMPTY_SEL, 'a', none);
    s = clickSelect(order, s, 'b', both); // anchor a..b added (anchor stays a)
    s = selectRun(order, s, run2, both);
    expect(ids(s)).toEqual(['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h']);
    const s2 = selectRun(order, EMPTY_SEL, run1, shift);
    expect(ids(s2)).toEqual(['c', 'd', 'e']);
    expect(s2.anchor).toBe('c');
  });

  it('an empty run changes nothing', () => {
    const s = clickSelect(order, EMPTY_SEL, 'a', none);
    expect(selectRun(order, s, [], none)).toBe(s);
  });
});

describe('run drag drop rules', () => {
  const order = ['a', 'b', 'c', 'd', 'e', 'f', 'g'];
  const run = new Set(['c', 'd', 'e']);

  it('dropping inside the dragged run (or right at its edges) is a no-op', () => {
    for (const gap of [2, 3, 4, 5]) expect(isNoopMove(order, run, gap), `gap ${gap}`).toBe(true);
  });

  it('dropping elsewhere moves the whole run, in its own order', () => {
    expect(computeMove(order, run, 0)).toEqual(['c', 'd', 'e', 'a', 'b', 'f', 'g']);
    expect(computeMove(order, run, 7)).toEqual(['a', 'b', 'f', 'g', 'c', 'd', 'e']);
    expect(computeMove(order, run, 6)).toEqual(['a', 'b', 'f', 'c', 'd', 'e', 'g']);
    expect(isNoopMove(order, run, 0)).toBe(false);
  });
});

describe('centerScrollLeft', () => {
  // canvas: row-number gutter 56, columns of 100 px; viewport 1000 px wide, frozen pane 200 px -> visible area [256, 1000]
  const starts = prefixSums(Array(30).fill(100));
  const max = 56 + 3000 - 1000;

  it('puts the middle of the group into the middle of the visible area', () => {
    const x0 = 56 + starts[10];
    const x1 = 56 + starts[13]; // 3 columns
    const sl = centerScrollLeft(x0, x1, 256, 1000, max);
    // the group's centre on screen = (x0 + x1) / 2 - sl = the middle of [256, 1000]
    expect((x0 + x1) / 2 - sl).toBeCloseTo(628, 0);
  });

  it('clamps at the left end (a group near the start stays where it is) and at the right end', () => {
    expect(centerScrollLeft(56, 156, 256, 1000, max)).toBe(0);
    expect(centerScrollLeft(56 + starts[28], 56 + starts[30], 256, 1000, max)).toBe(max);
  });

  it('a group wider than the visible area is aligned to its left edge (the beginning stays visible)', () => {
    const x0 = 56 + starts[5];
    const x1 = 56 + starts[20]; // 1500 px
    const sl = centerScrollLeft(x0, x1, 256, 1000, max);
    expect(x0 - sl).toBeGreaterThanOrEqual(256);
    expect(x0 - sl).toBeLessThan(256 + 20);
  });

  it('nothing to scroll (content narrower than the viewport) gives 0', () => {
    expect(centerScrollLeft(300, 400, 100, 900, 0)).toBe(0);
  });
});
