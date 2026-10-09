import { describe, expect, it } from 'vitest';
import { gapAt, indexAt, insertTarget, isNavKey, makeRowIndex, navigate, prefixSums } from '../../src/components/grid/geometry';
import { autoFitWidth, measureText } from '../../src/components/grid/autofit';

describe('prefixSums / gapAt / indexAt', () => {
  const starts = prefixSums([100, 50, 200, 10]); // 0 100 150 350 360
  it('prefix sums', () => {
    expect(Array.from(starts)).toEqual([0, 100, 150, 350, 360]);
  });
  it('gapAt = first column whose centre is right of x', () => {
    expect(gapAt(0, starts, 0, 4)).toBe(0);
    expect(gapAt(49, starts, 0, 4)).toBe(0);
    expect(gapAt(50, starts, 0, 4)).toBe(1); // centre of col 0 reached
    expect(gapAt(124, starts, 0, 4)).toBe(1);
    expect(gapAt(126, starts, 0, 4)).toBe(2);
    expect(gapAt(355, starts, 0, 4)).toBe(4);
    expect(gapAt(10_000, starts, 0, 4)).toBe(4);
    expect(gapAt(-50, starts, 0, 4)).toBe(0);
  });
  it('gapAt restricted to a pane', () => {
    expect(gapAt(0, starts, 2, 4)).toBe(2); // scrolling pane starts at 2
    expect(gapAt(10_000, starts, 0, 2)).toBe(2); // frozen pane ends at 2
  });
  it('indexAt', () => {
    expect(indexAt(0, starts)).toBe(0);
    expect(indexAt(99, starts)).toBe(0);
    expect(indexAt(100, starts)).toBe(1);
    expect(indexAt(355, starts)).toBe(3);
    expect(indexAt(9999, starts)).toBe(3);
    expect(indexAt(5, prefixSums([]))).toBe(-1);
  });
});

describe('makeRowIndex', () => {
  it('contiguous rows are O(1) arithmetic', () => {
    const idx = makeRowIndex([5, 6, 7, 8]);
    expect(idx(5)).toBe(0);
    expect(idx(8)).toBe(3);
    expect(idx(4)).toBe(-1);
    expect(idx(9)).toBe(-1);
    expect(idx(6.5)).toBe(-1);
  });
  it('non-contiguous rows (hidden rows left out)', () => {
    const idx = makeRowIndex([2, 3, 7, 10]);
    expect(idx(7)).toBe(2);
    expect(idx(4)).toBe(-1);
    expect(idx(10)).toBe(3);
  });
  it('empty', () => {
    expect(makeRowIndex([])(1)).toBe(-1);
  });
});

describe('navigate', () => {
  const size = { rows: 100, cols: 10, page: 20 };
  const nav = (k: Parameters<typeof navigate>[0], at: { rowIdx: number; colIdx: number }, mods = { ctrl: false, shift: false }) =>
    navigate(k, mods, at, size);
  it('arrows and clamping', () => {
    expect(nav('ArrowDown', { rowIdx: 0, colIdx: 0 })).toEqual({ rowIdx: 1, colIdx: 0 });
    expect(nav('ArrowUp', { rowIdx: 0, colIdx: 0 })).toEqual({ rowIdx: 0, colIdx: 0 });
    expect(nav('ArrowRight', { rowIdx: 0, colIdx: 9 })).toEqual({ rowIdx: 0, colIdx: 9 });
    expect(nav('ArrowLeft', { rowIdx: 3, colIdx: 4 })).toEqual({ rowIdx: 3, colIdx: 3 });
  });
  it('Ctrl+arrows jump to the edges', () => {
    const c = { ctrl: true, shift: false };
    expect(nav('ArrowDown', { rowIdx: 5, colIdx: 2 }, c)).toEqual({ rowIdx: 99, colIdx: 2 });
    expect(nav('ArrowUp', { rowIdx: 5, colIdx: 2 }, c)).toEqual({ rowIdx: 0, colIdx: 2 });
    expect(nav('ArrowRight', { rowIdx: 5, colIdx: 2 }, c)).toEqual({ rowIdx: 5, colIdx: 9 });
    expect(nav('ArrowLeft', { rowIdx: 5, colIdx: 2 }, c)).toEqual({ rowIdx: 5, colIdx: 0 });
  });
  it('Tab / Enter (with Shift), Home / End, Ctrl+Home / Ctrl+End, PageUp / PageDown', () => {
    const s = { ctrl: false, shift: true };
    const c = { ctrl: true, shift: false };
    expect(nav('Tab', { rowIdx: 1, colIdx: 1 })).toEqual({ rowIdx: 1, colIdx: 2 });
    expect(nav('Tab', { rowIdx: 1, colIdx: 1 }, s)).toEqual({ rowIdx: 1, colIdx: 0 });
    expect(nav('Enter', { rowIdx: 1, colIdx: 1 })).toEqual({ rowIdx: 2, colIdx: 1 });
    expect(nav('Enter', { rowIdx: 1, colIdx: 1 }, s)).toEqual({ rowIdx: 0, colIdx: 1 });
    expect(nav('Home', { rowIdx: 7, colIdx: 5 })).toEqual({ rowIdx: 7, colIdx: 0 });
    expect(nav('End', { rowIdx: 7, colIdx: 5 })).toEqual({ rowIdx: 7, colIdx: 9 });
    expect(nav('Home', { rowIdx: 7, colIdx: 5 }, c)).toEqual({ rowIdx: 0, colIdx: 0 });
    expect(nav('End', { rowIdx: 7, colIdx: 5 }, c)).toEqual({ rowIdx: 99, colIdx: 9 });
    expect(nav('PageDown', { rowIdx: 90, colIdx: 5 })).toEqual({ rowIdx: 99, colIdx: 5 });
    expect(nav('PageDown', { rowIdx: 10, colIdx: 5 })).toEqual({ rowIdx: 30, colIdx: 5 });
    expect(nav('PageUp', { rowIdx: 10, colIdx: 5 })).toEqual({ rowIdx: 0, colIdx: 5 });
  });
  it('empty grid', () => {
    expect(navigate('ArrowDown', { ctrl: false, shift: false }, { rowIdx: 0, colIdx: 0 }, { rows: 0, cols: 3, page: 5 })).toBeNull();
  });
  it('isNavKey', () => {
    expect(isNavKey('PageDown')).toBe(true);
    expect(isNavKey('a')).toBe(false);
    expect(isNavKey('Delete')).toBe(false);
  });
});

describe('insertTarget (Excel: insert as many columns as the contiguous selection)', () => {
  const order = ['a', 'b', 'c', 'd', 'e'];
  it('single column', () => {
    expect(insertTarget(order, ['c'], 'c', 'left')).toEqual({ index: 2, count: 1 });
    expect(insertTarget(order, ['c'], 'c', 'right')).toEqual({ index: 3, count: 1 });
    expect(insertTarget(order, ['e'], 'e', 'right')).toEqual({ index: 5, count: 1 });
  });
  it('contiguous selection', () => {
    expect(insertTarget(order, ['b', 'c', 'd'], 'c', 'left')).toEqual({ index: 1, count: 3 });
    expect(insertTarget(order, ['b', 'c', 'd'], 'c', 'right')).toEqual({ index: 4, count: 3 });
  });
  it('non-contiguous selection -> one column next to the clicked one', () => {
    expect(insertTarget(order, ['a', 'c'], 'c', 'left')).toEqual({ index: 2, count: 1 });
    expect(insertTarget(order, ['a', 'c'], 'a', 'right')).toEqual({ index: 1, count: 1 });
  });
  it('clicked column outside the selection', () => {
    expect(insertTarget(order, ['a', 'b'], 'd', 'left')).toEqual({ index: 3, count: 1 });
  });
});

describe('autofit (no DOM: estimated widths)', () => {
  it('measureText estimates without a canvas', () => {
    expect(measureText('', '11pt x')).toBe(0);
    expect(measureText('ab', '11pt x')).toBeGreaterThan(0);
    expect(measureText('日本', '11pt x')).toBeGreaterThan(measureText('ab', '11pt x'));
  });
  it('autoFitWidth takes the widest text / line, the header and clamps', () => {
    const texts = ['a', 'abcdefghij', null, 'x\nabcdefghijklmnopqrst'];
    const w = autoFitWidth(texts.length, (i) => (texts[i] ? { text: texts[i]!, font: '11pt x' } : null), { text: 'h', font: 'x', extra: 0 }, { min: 10, max: 1000, padding: 0 });
    expect(w).toBe(Math.ceil(measureText('abcdefghijklmnopqrst', 'x')));
    const clamped = autoFitWidth(1, () => ({ text: 'x'.repeat(500), font: 'x' }), { text: '', font: 'x', extra: 0 }, { min: 10, max: 300, padding: 8 });
    expect(clamped).toBe(300);
    const min = autoFitWidth(0, () => null, { text: '', font: 'x', extra: 0 }, { min: 24, max: 300, padding: 8 });
    expect(min).toBe(24);
  });
  it('samples very tall columns instead of freezing', () => {
    let calls = 0;
    autoFitWidth(1_000_000, () => (calls++, { text: 'abc', font: 'x' }), { text: '', font: 'x', extra: 0 }, { min: 1, max: 999, padding: 0 });
    expect(calls).toBeLessThanOrEqual(30000);
  });
});
