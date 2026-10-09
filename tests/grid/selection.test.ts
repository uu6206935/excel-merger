import { describe, expect, it } from 'vitest';
import { clickSelect, EMPTY_SEL, extendTo, orderedSelection, sameSelection, selectOnly, type ColSelection } from '../../src/components/grid/selection';

const order = ['a', 'b', 'c', 'd', 'e', 'f', 'g'];
const ids = (s: ColSelection) => orderedSelection(order, s.cols);
const none = { ctrl: false, shift: false };
const ctrl = { ctrl: true, shift: false };
const shift = { ctrl: false, shift: true };
const both = { ctrl: true, shift: true };

describe('clickSelect', () => {
  it('plain click selects only that column and sets the anchor', () => {
    const s = clickSelect(order, EMPTY_SEL, 'c', none);
    expect(ids(s)).toEqual(['c']);
    expect(s.anchor).toBe('c');
    const s2 = clickSelect(order, clickSelect(order, s, 'e', ctrl), 'a', none);
    expect(ids(s2)).toEqual(['a']);
  });
  it('Ctrl+click toggles (non-contiguous) and moves the anchor', () => {
    let s = clickSelect(order, EMPTY_SEL, 'b', none);
    s = clickSelect(order, s, 'd', ctrl);
    s = clickSelect(order, s, 'f', ctrl);
    expect(ids(s)).toEqual(['b', 'd', 'f']);
    expect(s.anchor).toBe('f');
    s = clickSelect(order, s, 'd', ctrl);
    expect(ids(s)).toEqual(['b', 'f']);
    expect(s.anchor).toBe('d');
  });
  it('Shift+click selects anchor..target (replacing), anchor kept; both directions', () => {
    let s = clickSelect(order, EMPTY_SEL, 'c', none);
    s = clickSelect(order, s, 'f', shift);
    expect(ids(s)).toEqual(['c', 'd', 'e', 'f']);
    expect(s.anchor).toBe('c');
    s = clickSelect(order, s, 'a', shift);
    expect(ids(s)).toEqual(['a', 'b', 'c']);
    expect(s.anchor).toBe('c');
  });
  it('Ctrl+Shift+click adds the range to the existing selection', () => {
    let s = clickSelect(order, EMPTY_SEL, 'a', none);
    s = clickSelect(order, s, 'e', ctrl);
    s = clickSelect(order, s, 'g', both);
    expect(ids(s)).toEqual(['a', 'e', 'f', 'g']);
  });
  it('Shift+click without an anchor behaves like a click', () => {
    const s = clickSelect(order, EMPTY_SEL, 'd', shift);
    expect(ids(s)).toEqual(['d']);
    expect(s.anchor).toBe('d');
  });
  it('Shift+click with a stale anchor (column deleted) behaves like a click', () => {
    const s = clickSelect(order, { cols: new Set(['zz']), anchor: 'zz' }, 'b', shift);
    expect(ids(s)).toEqual(['b']);
  });
  it('never mutates the input selection', () => {
    const s = clickSelect(order, EMPTY_SEL, 'a', none);
    const before = [...s.cols];
    clickSelect(order, s, 'b', ctrl);
    expect([...s.cols]).toEqual(before);
    expect(EMPTY_SEL.cols.size).toBe(0);
  });
});

describe('helpers', () => {
  it('orderedSelection follows the display order and drops unknown ids', () => {
    expect(orderedSelection(order, new Set(['f', 'x', 'b']))).toEqual(['b', 'f']);
    expect(orderedSelection(order, new Set())).toEqual([]);
  });
  it('selectOnly / extendTo / sameSelection', () => {
    expect(ids(selectOnly('c'))).toEqual(['c']);
    const s = extendTo(order, selectOnly('c'), 'e');
    expect(ids(s)).toEqual(['c', 'd', 'e']);
    expect(s.anchor).toBe('c');
    expect(ids(extendTo(order, s, 'b'))).toEqual(['b', 'c']);
    expect(ids(extendTo(order, EMPTY_SEL, 'b'))).toEqual(['b']);
    expect(sameSelection(selectOnly('a'), selectOnly('a'))).toBe(true);
    expect(sameSelection(selectOnly('a'), selectOnly('b'))).toBe(false);
  });
});
