import { describe, expect, it } from 'vitest';
import { cellCopyText, cellEditText, cellViewOf, isBlankCell, lacksCachedValue, renderCell } from '../../src/components/grid/cellview';
import { makeStyleCache } from '../../src/core/display/styles';
import type { XStyle } from '../../src/core/types';

const styles: XStyle[] = [{}, { numFmt: 'yyyy/m/d' }, { numFmt: '0.0%', hAlign: 'center' }, { numFmt: '#,##0;[Red]-#,##0' }, { numFmt: '@' }];
const cssOf = makeStyleCache(styles);

describe('cellViewOf (contract Cell -> CellView)', () => {
  it('values, types, styles, formulas', () => {
    expect(cellViewOf({ t: 'n', v: 3, s: 2 })).toEqual({ value: 3, t: 'n', styleId: 2 });
    expect(cellViewOf({ t: 'n', v: 3, f: 'A1+1' })).toEqual({ value: 3, t: 'n', styleId: 0, formula: '=A1+1' });
    expect(cellViewOf({ t: 'n', v: 3, f: 'A1+1' }, '=A1+1 (ui)')?.formula).toBe('=A1+1 (ui)');
    expect(cellViewOf({ t: 'z', v: null, f: 'SUM(A:A)' })).toEqual({ value: null, t: 'z', styleId: 0, formula: '=SUM(A:A)' });
    expect(cellViewOf(undefined)).toBeUndefined();
    expect(cellViewOf(undefined, null, 4)).toEqual({ value: null, styleId: 4 });
    expect(cellViewOf({ t: 's', v: 'x' }, null, 3)?.styleId).toBe(3);
  });
});

describe('renderCell', () => {
  it('formats with the number format and applies General alignment', () => {
    expect(renderCell({ value: 45566, t: 'd', styleId: 1 }, cssOf(1))).toMatchObject({ text: '2024/10/1', align: 'right', formulaText: false });
    expect(renderCell({ value: 'abc', styleId: 0 }, cssOf(0))).toMatchObject({ text: 'abc', align: 'left', kind: 'text' });
    expect(renderCell({ value: true }, cssOf(0))).toMatchObject({ text: 'TRUE', align: 'center' });
    expect(renderCell({ value: '#N/A', t: 'e' }, cssOf(0))).toMatchObject({ text: '#N/A', align: 'center', kind: 'error' });
    expect(renderCell({ value: 0.125, styleId: 2 }, cssOf(2))).toMatchObject({ text: '12.5%', align: 'center' });
    expect(renderCell({ value: -5, styleId: 3 }, cssOf(3))).toMatchObject({ text: '-5', color: '#FF0000' });
    expect(renderCell(undefined, cssOf(0))).toMatchObject({ text: '', kind: 'empty' });
  });
  it('a formula without a cached value shows its text (grey italic in the grid)', () => {
    const r = renderCell({ value: null, formula: '=SUM(A2:C2)' }, cssOf(3));
    expect(r).toMatchObject({ text: '=SUM(A2:C2)', formulaText: true, align: 'left' });
    expect(lacksCachedValue({ value: null, formula: '=1' })).toBe(true);
    expect(lacksCachedValue({ value: 0, formula: '=1' })).toBe(false);
    expect(lacksCachedValue({ value: '', formula: '=""', t: 's' })).toBe(false);
    expect(lacksCachedValue({ value: 5 })).toBe(false);
  });
});

describe('edit / copy text', () => {
  it('F2 / formula bar text', () => {
    expect(cellEditText({ value: 3, formula: '=A1+2' }, cssOf(0))).toBe('=A1+2');
    expect(cellEditText({ value: 45566, t: 'd', styleId: 1 }, cssOf(1))).toBe('2024/10/1');
    expect(cellEditText({ value: 0.125, styleId: 2 }, cssOf(2))).toBe('12.5%');
    expect(cellEditText({ value: '00123', t: 's', styleId: 4 }, cssOf(4))).toBe('00123');
    expect(cellEditText({ value: 1234567, styleId: 3 }, cssOf(3))).toBe('1234567');
    expect(cellEditText(undefined, cssOf(0))).toBe('');
    expect(cellEditText({ value: null }, undefined)).toBe('');
  });
  it('Ctrl+C copies the displayed text', () => {
    expect(cellCopyText({ value: 1234567, styleId: 3 }, cssOf(3))).toBe('1,234,567');
    expect(cellCopyText({ value: 45566, t: 'd', styleId: 1 }, cssOf(1))).toBe('2024/10/1');
  });
  it('isBlankCell', () => {
    expect(isBlankCell(undefined)).toBe(true);
    expect(isBlankCell({ value: null, styleId: 3 })).toBe(true);
    expect(isBlankCell({ value: '' })).toBe(true);
    expect(isBlankCell({ value: 0 })).toBe(false);
    expect(isBlankCell({ value: null, formula: '=1' })).toBe(false);
  });
});
