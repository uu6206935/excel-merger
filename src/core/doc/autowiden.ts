/**
 * Excel widens a column of standard width when a number format makes a typed number no longer fit (a date typed into
 * a General cell would otherwise show #####). The JS patcher does the same on replay (xlsx/patch/engine.ts
 * autoWiden, measured: tests/golden r-numfmt-w) and Excel itself does it for the Excel engine; the doc mirrors it so
 * the grid shows the width the output file will have.
 *
 * "Standard width" = an AUTO width: the column is not hidden and its width was not set by the user, i.e. ColProps.
 * customWidth is undefined. The reader reports customWidth only for <col customWidth="1"> WITHOUT bestFit, so a
 * column Excel itself auto-fitted (bestFit="1" customWidth="1") and a stored width without the flag (what Excel wrote
 * for an untouched column) are both auto: they are widened when the formatted text does not fit (the patcher then
 * writes bestFit + customWidth, which reads back as auto again). A column the user sized keeps its width, and so does
 * a width set by setColumnWidth (which sets the flag; the patcher drops bestFit when it applies that width).
 *
 * The doc does NOT set customWidth when it widens a column: the patcher's result (bestFit="1" customWidth="1") may be
 * widened again by a later, wider value, exactly like an auto-fitted column in Excel, so the doc leaves the flag alone.
 */
import { formatValue } from '../display/numfmt';
import { defaultColumnPx, excelWidthToPx } from '../display/styles';
import { fontMdw, numberTextPx } from '../xlsx/patch/autowidth';
import type { DocColumn, MergedDoc } from '../types';

/**
 * sheetFormatPr defaultColWidth of the merged sheet, kept beside the doc (MergedDoc has no field for it). Keyed by
 * doc.rowProps, the one array createEditor creates and no op replaces.
 */
const defaultWidths = new WeakMap<object, number>();

export function setDefaultColWidth(doc: MergedDoc, width: number | undefined): void {
  if (width !== undefined && width > 0) defaultWidths.set(doc.rowProps, width);
}

/**
 * New stored width of `col` after `value` got number format `code` in a cell of style `xf`, or null (no change).
 * `styles` = the style list `xf` indexes: when the typed format created a new xf (the cell's xf + the number format),
 * that is the derived list, not doc.styles (the text is measured in the cell's own font, as the patcher does).
 */
export function autoWidenedWidth(
  doc: MergedDoc,
  col: DocColumn,
  xf: number | undefined,
  value: number,
  code: string,
  styles: MergedDoc['styles'] = doc.styles,
): number | null {
  // a user-set width (customWidth = set by the user, no bestFit) and a hidden column are never touched
  if (col.props.customWidth || col.props.hidden || !Number.isFinite(value)) return null;
  // General never shows ##### (it shortens the decimals to the column), so Excel does not widen for it
  if (/^general$/i.test(code.trim())) return null;
  let text: string;
  try {
    text = formatValue(value, code).text;
  } catch {
    return null;
  }
  if (!text || /^#+$/.test(text) || text.includes('\n')) return null;
  const def = styles[0] ?? {};
  const mdw = fontMdw(def.font, def.size);
  const st = styles[xf ?? 0] ?? {};
  const need = numberTextPx(text, fontMdw(st.font, st.size));
  // the column's own stored width, else the sheet's default column width (like the patcher)
  const cur = col.props.width;
  const dw = defaultWidths.get(doc.rowProps);
  const px = cur !== undefined && Number.isFinite(cur) && cur > 0 ? excelWidthToPx(cur, mdw) : dw ? excelWidthToPx(dw, mdw) : defaultColumnPx(mdw, 8);
  if (need <= px - 5) return null;
  return Math.min(255, Math.round(((need + 5) / mdw) * 1e6) / 1e6);
}
