/**
 * Display text of a cell as Excel shows it (owner: reader R): number format applied via display/numfmt.formatValue
 * (00123, 2026/10/20, ¥1,234, 12.5%, 和暦 …). Column-width effects (#### / General shrinking) are not simulated.
 */
import type { Cell, SourceBook } from '../types';
import { formatValue } from '../display/numfmt';

export function cellText(book: SourceBook, cell: Cell | undefined): string {
  if (!cell) return '';
  const v = cell.v;
  if (v === null || v === undefined) return '';
  switch (cell.t) {
    case 'z':
      return '';
    case 'e':
      return String(v);
    case 'b':
      return v === true || v === 1 || v === 'TRUE' ? 'TRUE' : 'FALSE';
    default:
      break;
  }
  const fmt = book.styles[cell.s ?? 0]?.numFmt;
  if (typeof v === 'string') {
    if (!fmt || fmt === '@' || fmt === 'General') return v;
    return formatValue(v, fmt).text;
  }
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (cell.t === 'd' && (!fmt || fmt === 'General')) {
    // a date without a date format (cannot come from xlsx, but keep it readable)
    return formatValue(v, Math.abs(v - Math.floor(v)) > 1e-9 ? 'yyyy/m/d h:mm' : 'yyyy/m/d').text;
  }
  return formatValue(v, fmt ?? 'General').text;
}
