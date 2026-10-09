/**
 * What a grid cell shows / edits / copies (pure; no React, no DOM). Unit-tested in tests/grid/cellview.test.ts.
 */
import type { Cell } from '../../core/types';
import type { CellCss } from '../../core/display/styles';
import { displayValue, editText, type FormattedValue } from '../../core/display/numfmt';
import type { CellView } from './types';

/**
 * Contract `Cell` -> CellView (for GridProps.getCell).
 * @param formulaText the formula in UI form WITH '=' (e.g. doc.formulaForDisplay); defaults to '=' + cell.f
 * @param styleId style to use when the cell has none (column / row default style), default 0
 */
export function cellViewOf(cell: Cell | undefined, formulaText?: string | null, styleId?: number): CellView | undefined {
  if (!cell) return styleId !== undefined && styleId !== 0 ? { value: null, styleId } : undefined;
  const value = cell.t === 'z' || cell.v === undefined ? null : cell.v;
  const view: CellView = { value, t: cell.t, styleId: cell.s ?? styleId ?? 0 };
  if (cell.f) view.formula = formulaText ?? (cell.f.startsWith('=') ? cell.f : '=' + cell.f);
  return view;
}

export interface CellRender {
  text: string;
  color?: string;
  align: 'left' | 'center' | 'right' | 'justify';
  kind: FormattedValue['kind'];
  /** true = the formula text is shown because the formula has no cached value (grey italic) */
  formulaText: boolean;
}

/** true when a formula cell has no cached result to display (never calculated / written by the tool). */
export function lacksCachedValue(cell: CellView | undefined): boolean {
  if (!cell || !cell.formula) return false;
  return cell.value === null || cell.value === undefined || cell.t === 'z';
}

function generalAlign(kind: FormattedValue['kind']): CellRender['align'] {
  return kind === 'number' ? 'right' : kind === 'bool' || kind === 'error' ? 'center' : 'left';
}

/** Displayed text, colour (number-format section colour) and alignment of a cell (Excel "General" alignment rules). */
export function renderCell(cell: CellView | undefined, css: CellCss): CellRender {
  if (!cell) return { text: '', align: css.textAlign ?? 'left', kind: 'empty', formulaText: false };
  if (lacksCachedValue(cell)) {
    return { text: cell.formula!, align: 'left', kind: 'text', formulaText: true };
  }
  let f: FormattedValue;
  try {
    f = displayValue({ t: cell.t, v: cell.value }, css.numFmt);
  } catch {
    f = { text: String(cell.value ?? ''), kind: 'text' };
  }
  const align = css.textAlign ?? generalAlign(f.kind);
  return f.color
    ? { text: f.text, color: f.color, align, kind: f.kind, formulaText: false }
    : { text: f.text, align, kind: f.kind, formulaText: false };
}

/** Text for F2 / double-click editing and the formula bar: the formula (UI form) or the constant as Excel's formula bar shows it. */
export function cellEditText(cell: CellView | undefined, css: Pick<CellCss, 'numFmt'> | undefined): string {
  if (!cell) return '';
  if (cell.formula) return cell.formula;
  try {
    return editText({ t: cell.t, v: cell.value }, css?.numFmt);
  } catch {
    return cell.value === null || cell.value === undefined ? '' : String(cell.value);
  }
}

/** Text put on the clipboard by Ctrl+C: the displayed text (like Excel's plain-text clipboard format). */
export function cellCopyText(cell: CellView | undefined, css: CellCss): string {
  return renderCell(cell, css).text;
}

/** true when the cell shows nothing and holds nothing (Delete on it would be a no-op). */
export function isBlankCell(cell: CellView | undefined): boolean {
  if (!cell) return true;
  if (cell.formula) return false;
  return cell.value === null || cell.value === undefined || cell.value === '';
}
