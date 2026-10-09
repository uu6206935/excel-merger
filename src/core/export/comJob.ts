/**
 * Op log -> job for server/excel/excel-runner.ps1 (the "Excel本体" export engine). Owner: server/COM (S).
 *
 * The runner replays the SAME log the JS engine replays (docs/ARCHITECTURE.md §2), in the user's own Excel:
 *   struct insert  -> insertCols / insertRows   { sheet, at, count, formatFrom: 'left' | 'above' }
 *   struct delete  -> deleteCols / deleteRows   { sheet, at, count }
 *   struct move    -> moveCols / moveRows       { sheet, start: from, count, to, method: 'clipFree' }
 *                     (to = insert-before index in ORIGINAL coordinates, exactly like excelRef's StructOp)
 *   cells          -> setCells                  { sheet, at: 'C5', rows: [[cell], ...], toValues? }   (one op per
 *                     contiguous vertical run of written cells; last write to a position wins; toValues: true for
 *                     reason 'toValues' = an array / spill formula under a written cell is replaced by its values)
 *   colWidth       -> setColumnWidth            { sheet, col, width, unit: 'stored' }  (the logged width is the
 *                     value stored in <col width>; the runner subtracts the workbook's padding for ColumnWidth)
 *   unmerge        -> unmerge                   { sheet, refs, centerAcross }
 * All positions are 1-based. Cell encoding of setCells (see server/excel/README.md):
 *   null = clear the value (format kept) · string = text written as it is (Range.Value2 = text; the runner reads it back
 *   and, if Excel converted it anyway, rewrites it with the apostrophe and warns TEXT_COERCED) · number · boolean ·
 *   { t: 's', v: text, quote: true } = text WITH the quote prefix ("'" + text, quotePrefix="1") — exactly when
 *   CellWrite.quote === true, never guessed by the runner · { f: '=SUM(A1:A3)' } = legacy formula in UI form, written with
 *   Range.Formula · { f, dyn: true } = dynamic-array formula, written with Range.Formula2 · { e: '#N/A' } = error literal ·
 *   any value with a number format: { t, v, numFmt } / { f, dyn?, numFmt } (Range.NumberFormat: "@" before the value, every
 *   other code after it, so Excel widens a too-narrow standard-width column for a date exactly like typing it does;
 *   'General' resets the format: the runner translates it to the name this Excel knows, "G/標準" in Japanese Excel).
 */
import type { CellWrite, LoggedOp } from '../types';
import type { RunnerJob } from '../contracts';
import { numToCol } from '../formula/excelRef';
import { toUiForm } from '../formula/forms';

/** Compact setCells cell value (what excel-runner.ps1 / XlRunnerHelper.cs Cells.Build accepts). */
export type RunnerCellValue =
  | null
  | string
  | number
  | boolean
  | { f: string; dyn?: true; numFmt?: string }
  | { e: string }
  | { t: 's' | 'n' | 'b' | 'e' | 'z'; v: string | number | boolean | null; numFmt: string }
  | { t: 's'; v: string; quote: true; numFmt?: string };

const ERROR_LITERALS = new Set([
  '#NULL!',
  '#DIV/0!',
  '#VALUE!',
  '#REF!',
  '#NAME?',
  '#NUM!',
  '#N/A',
  '#GETTING_DATA',
  '#SPILL!',
  '#CALC!',
  '#FIELD!',
  '#BLOCKED!',
  '#CONNECT!',
  '#BUSY!',
  '#UNKNOWN!',
]);

/** Value part of a cell (no formula, no number format). */
function plainValue(cell: NonNullable<CellWrite['cell']>): null | string | number | boolean | { e: string } {
  const v = cell.v;
  switch (cell.t) {
    case 'z':
      return null;
    case 'b':
      return typeof v === 'boolean' ? v : String(v).toUpperCase() === 'TRUE';
    case 'n':
    case 'd': {
      const n = typeof v === 'number' ? v : Number(v);
      if (Number.isFinite(n)) return n;
      return v == null ? null : String(v);
    }
    case 'e': {
      const e = String(v ?? '').toUpperCase();
      return ERROR_LITERALS.has(e) ? { e } : e.length > 0 ? String(v) : null;
    }
    case 's':
    default:
      if (v == null) return null;
      return typeof v === 'string' ? (v.length === 0 ? null : v) : String(v);
  }
}

const EXT_BARE = /\[[^[\]]+\][A-Za-z0-9_.\u0080-\uFFFF]*!/;
const DANGER_FN = /(?:^|[^A-Za-z0-9_])(?:_xlfn\.)?(?:WEBSERVICE|IMAGE|RTD|CALL|REGISTER(?:\.ID)?|EXEC)\s*\(/i;

/**
 * true when entering this formula text would make Excel reach outside the workbook by itself (CSV-injection style attacks on text
 * that came from untrusted files): DDE (`=cmd|' /C calc'!A0`: a `|` outside string literals and quoted names), external workbook
 * references (`[Book.xlsx]Sheet1!A1`, `'C:\dir\[Book.xlsx]S'!A1`, UNC / URL paths) and WEBSERVICE( IMAGE( RTD( CALL( REGISTER( EXEC(.
 * HYPERLINK and structured references are fine. Same rule, same test vectors as `XlRunner.Cells.UnsafeFormula` in
 * server/excel/XlRunnerHelper.cs (tests/server/fixtures/unsafe-formulas.json): the Excel本体 runner writes such a formula as plain
 * text and warns UNSAFE_FORMULA; the built-in engine (replayLog) may use this predicate to do the same. Quoting follows Excel's
 * tokenizer: "..." strings and '...' sheet names are recognised only where a token starts, "" / '' are escapes.
 */
export function isUnsafeFormula(f: string): boolean {
  if (!f) return false;
  const s = f.normalize('NFKC'); // the full-width forms are read like the ASCII ones
  let outside = '';
  const n = s.length;
  let i = 0;
  while (i < n) {
    const ch = s[i];
    if (ch === '"') {
      i++;
      while (i < n) {
        if (s[i] === '"') {
          if (i + 1 < n && s[i + 1] === '"') {
            i += 2;
            continue;
          }
          break;
        }
        i++;
      }
      i++;
      outside += '""';
      continue;
    }
    if (ch === "'") {
      let name = '';
      let j = i + 1;
      while (j < n) {
        if (s[j] === "'") {
          if (j + 1 < n && s[j + 1] === "'") {
            name += "'";
            j += 2;
            continue;
          }
          break;
        }
        name += s[j];
        j++;
      }
      // an internal sheet name never contains [ ] \ / (Excel refuses them): this is a path / UNC / URL / book reference
      if (j + 1 < n && s[j + 1] === '!' && /[[\]\\/]/.test(name)) return true;
      i = j + 1;
      outside += "'_'";
      continue;
    }
    if (ch === '|') return true;
    outside += ch;
    i++;
  }
  return EXT_BARE.test(outside) || DANGER_FN.test(outside);
}

/** Converts one CellWrite payload into the runner's compact cell encoding. */
export function toRunnerCell(cell: CellWrite['cell']): RunnerCellValue {
  if (cell == null) return null;
  const numFmt = typeof cell.numFmt === 'string' && cell.numFmt.length > 0 ? cell.numFmt : undefined;
  if (typeof cell.f === 'string' && cell.f.length > 0) {
    // FILE form (no '=') -> UI form with '=' (_xlfn. dropped, ANCHORARRAY -> #, ...): exactly what Range.Formula
    // shows for a legacy cell and Range.Formula2 for a dynamic-array one.
    const ui = toUiForm(cell.f);
    const out: { f: string; dyn?: true; numFmt?: string } = { f: ui.startsWith('=') ? ui : '=' + ui };
    if (cell.dyn) out.dyn = true;
    if (numFmt) out.numFmt = numFmt;
    return out;
  }
  const value = plainValue(cell);
  // CellWrite.quote === true (strictly): a text with the quote prefix. Only a non-empty string can carry it.
  if (cell.quote === true && typeof value === 'string') {
    const out: { t: 's'; v: string; quote: true; numFmt?: string } = { t: 's', v: value, quote: true };
    if (numFmt) out.numFmt = numFmt;
    return out;
  }
  if (!numFmt) return value;
  // with a number format: the typed RunnerCell form { t, v, numFmt }
  if (value === null) return { t: 'z', v: null, numFmt };
  if (typeof value === 'object') return { t: 'e', v: value.e, numFmt };
  if (typeof value === 'boolean') return { t: 'b', v: value, numFmt };
  if (typeof value === 'number') return { t: 'n', v: value, numFmt };
  return { t: 's', v: value, numFmt };
}

/** Groups the writes of ONE cells op into setCells ops: one per contiguous vertical run (last write wins). */
function cellsToOps(sheet: string, writes: CellWrite[], toValues: boolean): Array<Record<string, unknown>> {
  const byPos = new Map<string, CellWrite>();
  for (const w of writes ?? []) {
    if (!w || !Number.isInteger(w.row) || !Number.isInteger(w.col) || w.row < 1 || w.col < 1) continue;
    byPos.set(`${w.col}:${w.row}`, w);
  }
  if (byPos.size === 0) return [];
  const byCol = new Map<number, CellWrite[]>();
  for (const w of byPos.values()) {
    let list = byCol.get(w.col);
    if (!list) byCol.set(w.col, (list = []));
    list.push(w);
  }
  const ops: Array<Record<string, unknown>> = [];
  for (const col of [...byCol.keys()].sort((a, b) => a - b)) {
    const list = byCol.get(col)!.sort((a, b) => a.row - b.row);
    let i = 0;
    while (i < list.length) {
      let j = i;
      while (j + 1 < list.length && list[j + 1].row === list[j].row + 1) j++;
      const rows: RunnerCellValue[][] = [];
      for (let k = i; k <= j; k++) rows.push([toRunnerCell(list[k].cell)]);
      const op: Record<string, unknown> = { op: 'setCells', sheet, at: `${numToCol(col)}${list[i].row}`, rows };
      if (toValues) op.toValues = true;
      ops.push(op);
      i = j + 1;
    }
  }
  return ops;
}

const A1_REF = /^\$?[A-Za-z]{1,3}\$?[0-9]{1,7}(:\$?[A-Za-z]{1,3}\$?[0-9]{1,7})?$/;

export function logToRunnerJob(log: LoggedOp[], opts: { format: 'xlsx' | 'xlsm'; sheet: string }): RunnerJob {
  const ops: Array<Record<string, unknown>> = [];
  for (const entry of log ?? []) {
    if (!entry) continue;
    const sheet = entry.sheet || opts.sheet;
    switch (entry.kind) {
      case 'struct': {
        const op = entry.op;
        const cols = op.axis === 'col';
        if (op.type === 'insert') {
          if (op.count > 0) ops.push({ op: cols ? 'insertCols' : 'insertRows', sheet, at: op.at, count: op.count, formatFrom: cols ? 'left' : 'above' });
        } else if (op.type === 'delete') {
          if (op.count > 0) ops.push({ op: cols ? 'deleteCols' : 'deleteRows', sheet, at: op.at, count: op.count });
        } else if (op.type === 'move') {
          // A move whose target lies inside / right after the block changes nothing (Excel does nothing either).
          if (op.count > 0 && !(op.to >= op.from && op.to <= op.from + op.count)) {
            ops.push({ op: cols ? 'moveCols' : 'moveRows', sheet, start: op.from, count: op.count, to: op.to, method: 'clipFree' });
          }
        }
        break;
      }
      case 'cells':
        ops.push(...cellsToOps(sheet, entry.writes, entry.reason === 'toValues'));
        break;
      case 'colWidth':
        if (Number.isFinite(entry.width) && entry.width >= 0) {
          // stored <col width> value (includes the ~5 px padding); the runner converts it to Range.ColumnWidth
          ops.push({ op: 'setColumnWidth', sheet, col: entry.col, width: Math.min(256, entry.width), unit: 'stored' });
        }
        break;
      case 'unmerge': {
        const refs = (entry.refs ?? []).map((r) => String(r).trim()).filter((r) => A1_REF.test(r));
        if (refs.length) ops.push({ op: 'unmerge', sheet, refs, centerAcross: !!entry.centerAcross });
        break;
      }
      default:
        break;
    }
  }
  return {
    version: 1,
    format: opts.format,
    ops,
    // restoreEachOp: the user's clipboard is put back right after every cut, which also takes clipboard
    // ownership away from the hidden Excel (clipboard managers would otherwise stall it).
    options: { recalc: true, clipboard: 'restoreEachOp' },
  };
}
