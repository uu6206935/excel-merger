/**
 * SourceBook adapter for the key detector:
 *   - cell VALUES as the keys module sees them (cached formula results, dates as UTC Date objects, errors blank);
 *   - pickDataSheet / detectRegion (header row by dictionary hits or opts.headerRow, contiguous header block; the table
 *     ends at a blank-row run, a note line or trailing summary rows — 合計 / 人数 / 平均 … labels and/or aggregate
 *     formulas over the column above, also below 1–3 blank rows — or at opts.lastDataRow);
 *   - buildTable: KeySource (book + region) → R6 Table (headers, matchable rows, lazy display text, numFmt per column).
 *
 * Pure: no DOM, no Node APIs. Never throws on malformed input (missing sheet, inverted region …): it degrades to an
 * empty table instead.
 */
import type { Cell as XCell, SheetGrid, SourceBook, TableRegion } from '../types';
import type { KeySource } from '../contracts';
import { displayValue, isDateFormat } from '../display/numfmt';
import { type Token, parseAreaText, relKey, tokenize } from '../formula/excelRef';
import { type Cell, cellText, excelSerialToDate, isExampleMark, isNoteLineText, isPageMarkerText, nfkc, summaryLabelKind } from './normalize';
import { headerNorm, bestHeaderType, isNonKeyHeader } from './detect';
import { type Table, MAX_HEADER_SCAN, detectHeaderRow, headerText, wideLayoutHint } from './table';

/** Wider sheets are cut here for header detection / region scanning (formatting-only columns are never data). */
const MAX_SCAN_COLS = 4096;
/** A run of this many blank rows ends the table. */
const BLANK_RUN_END = 3;

// ---------------------------------------------------------------------------
// cell values
// ---------------------------------------------------------------------------

/** Contract cell → key value (cached formula result; 'd' serial → UTC Date; errors / no cached value → null). */
export function keyValue(c: XCell | undefined): Cell {
  if (!c) return null;
  const v = c.v;
  if (v === null || v === undefined) return null;
  switch (c.t) {
    case 'z':
    case 'e':
      return null;
    case 'd':
      if (typeof v === 'number') return excelSerialToDate(v);
      return typeof v === 'string' ? v : null;
    case 'b':
      return typeof v === 'boolean' ? v : v === 1 || v === 'TRUE';
    case 'n':
      if (typeof v === 'number') return Number.isFinite(v) ? v : null;
      return typeof v === 'string' ? v : null;
    default:
      return typeof v === 'string' ? v : typeof v === 'number' || typeof v === 'boolean' ? v : String(v);
  }
}

export function numFmtOf(book: SourceBook, c: XCell | undefined): string {
  if (!c) return 'General';
  return book.styles?.[c.s ?? 0]?.numFmt ?? 'General';
}

/** The text Excel displays for a cell (number format applied: 000123, 2001/4/1). */
export function displayText(book: SourceBook, c: XCell | undefined): string {
  if (!c || c.v === null || c.v === undefined) return '';
  try {
    return displayValue({ t: c.t, v: c.v }, numFmtOf(book, c)).text;
  } catch {
    return c.v === null ? '' : String(c.v);
  }
}

function rawCell(sheet: SheetGrid, r: number, c: number): XCell | undefined {
  return sheet.rows?.[r - 1]?.[c - 1];
}

function isBlankValue(v: Cell): boolean {
  return v === null || v === undefined || (typeof v === 'string' && v.trim() === '');
}

export function findSheet(book: SourceBook, name: string): SheetGrid | undefined {
  return book.sheets?.find((s) => s.name === name);
}

function sheetCols(sheet: SheetGrid, r1: number, r2: number): number {
  let n = Math.min(Math.max(0, sheet.maxCol | 0), MAX_SCAN_COLS);
  // maxCol may be missing / too small in hand-built books: trust the row arrays too
  for (let r = r1; r <= r2; r++) {
    const row = sheet.rows?.[r - 1];
    if (row && row.length > n) n = Math.min(row.length, MAX_SCAN_COLS);
  }
  return n;
}

function sheetRows(sheet: SheetGrid): number {
  return Math.max(sheet.maxRow | 0, sheet.rows?.length ?? 0);
}

/** Value grid of rows r1..r2 (1-based, inclusive) × columns 1..nCols; merged cells inside it take the top-left value. */
function valueGrid(sheet: SheetGrid, r1: number, r2: number, nCols: number): Cell[][] {
  const grid: Cell[][] = [];
  for (let r = r1; r <= r2; r++) {
    const row: Cell[] = new Array(nCols).fill(null);
    const src = sheet.rows?.[r - 1];
    if (src) {
      const n = Math.min(src.length, nCols);
      for (let c = 0; c < n; c++) row[c] = keyValue(src[c]);
    }
    grid.push(row);
  }
  for (const m of sheet.features?.merges ?? []) {
    if (m.r2 < r1 || m.r1 > r2 || m.c1 > nCols) continue;
    const top = m.r1 >= r1 ? grid[m.r1 - r1]?.[m.c1 - 1] : keyValue(rawCell(sheet, m.r1, m.c1));
    if (isBlankValue(top ?? null)) continue;
    for (let r = Math.max(m.r1, r1); r <= Math.min(m.r2, r2); r++) {
      for (let c = m.c1; c <= Math.min(m.c2, nCols); c++) {
        const row = grid[r - r1];
        if (row && isBlankValue(row[c - 1])) row[c - 1] = top;
      }
    }
  }
  return grid;
}

/** true when the header cell counts as a dictionary hit (key synonym ≥ 0.7 or a known non-key header word). */
function isHeaderHit(v: Cell): boolean {
  if (typeof v !== 'string') return false;
  const hn = headerNorm(nfkc(v));
  return hn !== '' && (bestHeaderType(hn).score >= 0.7 || isNonKeyHeader(hn));
}

/** the header row names a person (氏名 / ID / 学籍番号 / フリガナ / メール …): the table is a list of people */
const PERSON_HEADER_TYPES: ReadonlySet<string> = new Set(['name', 'lastName', 'firstName', 'nameKana', 'nameRoman', 'id', 'studentId', 'employeeId', 'email', 'phone']);
function isPersonList(headers: readonly Cell[]): boolean {
  for (const h of headers) {
    if (typeof h !== 'string') continue;
    const b = bestHeaderType(headerNorm(nfkc(h)));
    if (b.score >= 0.7 && b.type && PERSON_HEADER_TYPES.has(b.type)) return true;
  }
  return false;
}

// ---------------------------------------------------------------------------
// region detection
// ---------------------------------------------------------------------------

interface Block {
  c1: number;
  c2: number;
  hits: number;
  filled: number;
}

/**
 * Contiguous header block of row h: consecutive columns with a header text; a header-less column is bridged when
 * its column has data below AND the block continues after it (a forgotten title inside a list).
 */
function headerBlocks(headerRow: Cell[], hasDataBelow: (c: number) => boolean): Block[] {
  const n = headerRow.length;
  const filled = (c: number) => !isBlankValue(headerRow[c - 1] ?? null);
  const blocks: Block[] = [];
  let c = 1;
  while (c <= n) {
    if (!filled(c)) {
      c++;
      continue;
    }
    const b: Block = { c1: c, c2: c, hits: 0, filled: 0 };
    let k = c;
    while (k <= n) {
      if (filled(k)) {
        b.c2 = k;
        b.filled++;
        if (isHeaderHit(headerRow[k - 1])) b.hits++;
        k++;
        continue;
      }
      // bridge a single header-less data column when the header continues right after it
      if (k + 1 <= n && filled(k + 1) && hasDataBelow(k)) {
        k++;
        continue;
      }
      break;
    }
    blocks.push(b);
    c = k + 1;
  }
  return blocks;
}

function emptyRegion(sheet: string): TableRegion {
  return { sheet, headerRow: 1, firstDataRow: 2, lastDataRow: 1, firstCol: 1, lastCol: 1 };
}

function rowNonEmpty(sheet: SheetGrid, r: number, c1: number, c2: number): number {
  const row = sheet.rows?.[r - 1];
  if (!row) return 0;
  let n = 0;
  const end = Math.min(c2, row.length);
  for (let c = c1; c <= end; c++) if (!isBlankValue(keyValue(row[c - 1]))) n++;
  return n;
}

// ---------------------------------------------------------------------------
// summary (footer) rows: 合計 / 小計 / 人数 / 平均 … and aggregate formulas over the column above
// ---------------------------------------------------------------------------

/** Functions that aggregate a range (a footer's =SUM(D4:D63), =COUNTA(B4:B63), =COUNTIF(K4:K63,">=4") …). */
const AGG_FUNCS = new Set([
  'SUM', 'SUMIF', 'SUMIFS', 'SUMPRODUCT', 'SUMSQ', 'SUBTOTAL', 'AGGREGATE', 'PRODUCT',
  'COUNT', 'COUNTA', 'COUNTBLANK', 'COUNTIF', 'COUNTIFS',
  'AVERAGE', 'AVERAGEA', 'AVERAGEIF', 'AVERAGEIFS', 'MEDIAN', 'MODE', 'MODE.SNGL', 'MODE.MULT', 'GEOMEAN', 'HARMEAN', 'TRIMMEAN',
  'MAX', 'MAXA', 'MAXIFS', 'MIN', 'MINA', 'MINIFS', 'LARGE', 'SMALL',
  'STDEV', 'STDEV.S', 'STDEV.P', 'STDEVA', 'STDEVP', 'STDEVPA', 'VAR', 'VAR.S', 'VAR.P', 'VARA', 'VARP', 'VARPA', 'AVEDEV', 'DEVSQ',
  'PERCENTILE', 'PERCENTILE.INC', 'PERCENTILE.EXC', 'QUARTILE', 'QUARTILE.INC', 'QUARTILE.EXC',
]);
/** cheap pre-filter before tokenizing a formula */
const AGG_HINT_RE = /(SUM|SUBTOTAL|AGGREGATE|PRODUCT|COUNT|AVERAGE|MEDIAN|MODE|MEAN|MAX|MIN|LARGE|SMALL|STDEV|VAR|AVEDEV|DEVSQ|PERCENTILE|QUARTILE)[A-Z.]*\s*\(/i;
const FUNC_PREFIX_RE = /^_xl(?:fn|ws)\./i;
const THIS_ROW_RE = /@|#this\s*row/i;

/** R1C1-like key of a formula (null when it cannot be tokenized) */
function safeRelKey(f: string, r: number, c: number): string | null {
  try {
    return relKey(f, r, c);
  } catch {
    return null;
  }
}

/**
 * true when the formula of cell (r, c) is filled down: the cell directly above or below holds the same relative
 * formula (No =MAX($A$3:A4)+1, running totals =SUM($E$4:E5), 既出 =COUNTIF($B$4:B4,B5)). A footer formula is
 * unique in its column.
 */
function isFilledDown(sheet: SheetGrid, f: string, r: number, c: number): boolean {
  let key: string | null | undefined;
  for (const rr of [r - 1, r + 1]) {
    const g = rr >= 1 ? rawCell(sheet, rr, c)?.f : undefined;
    if (typeof g !== 'string' || !g) continue;
    key ??= safeRelKey(f, r, c);
    if (key !== null && safeRelKey(g, rr, c) === key) return true;
  }
  return false;
}

/**
 * true when formula `f` in cell (r, c) aggregates a range of the SAME sheet that lies entirely above the row and
 * reaches into the data (rows ≥ top): =SUM(D4:D63) in row 65, =SUBTOTAL(109,Table1[数量]). A formula that references
 * its own row or rows below it is a per-row formula (合計点 =SUM(D5:F5), =COUNTIF($B$4:$B$63,B5)); so is a formula
 * filled down the column (numbering =MAX($A$3:A4)+1, running totals =SUM($E$4:E5)).
 */
function isFooterAggregate(sheet: SheetGrid, f: string, r: number, c: number, top: number): boolean {
  if (!AGG_HINT_RE.test(f)) return false;
  const sheetName = sheet.name;
  let toks: Token[];
  try {
    toks = tokenize(f);
  } catch {
    return false;
  }
  let agg = false;
  let above = false;
  for (const t of toks) {
    if (t.type === 'func') {
      if (AGG_FUNCS.has(t.text.replace(FUNC_PREFIX_RE, '').toUpperCase())) agg = true;
    } else if (t.type === 'ref' && t.area) {
      const p = t.prefix;
      // other sheets / books / 3D refs say nothing about this sheet's rows
      if (p && (p.book !== undefined || p.sheet2 !== undefined || (p.sheet !== null && p.sheet !== sheetName))) continue;
      const a = t.area;
      if (a.kind === 'cols') {
        // =SUMIF(C:C,"東京大学") / =COUNTA(B:B): a whole-column aggregate (never in its own column: that is circular)
        if (c >= a.c1 && c <= a.c2) return false;
        above = true;
        continue;
      }
      if (a.r2 >= r) return false;
      if (a.kind !== 'rows' && a.r2 >= top && (a.r2 > a.r1 || a.c1 === a.c2)) above = true;
    } else if (t.type === 'sref') {
      if (THIS_ROW_RE.test(t.text)) return false;
      above = true; // a whole table column
    }
  }
  return agg && above && !isFilledDown(sheet, f, r, c);
}

const totalsRowCache = new WeakMap<SheetGrid, Set<number>>();
/** last rows of the sheet's Excel tables that have a totals row */
function tableTotalsRows(sheet: SheetGrid): Set<number> {
  let s = totalsRowCache.get(sheet);
  if (!s) {
    s = new Set<number>();
    for (const t of sheet.features?.tables ?? []) {
      if (!t?.totalsRow || typeof t.ref !== 'string') continue;
      const a = parseAreaText(t.ref.replace(/\$/g, ''));
      if (a && a.kind === 'area') s.add(a.r2);
    }
    totalsRowCache.set(sheet, s);
  }
  return s;
}

/**
 * Kind of a summary row (null = not one):
 *   'label'     a 合計 / 小計 / 総計 / Total label among the first three non-blank cells;
 *   'table'     the totals row of an Excel table;
 *   'aggregate' aggregate formulas over the column above and no text except summary labels (=SUM(E2:E6));
 *   'weak'      a 人数 / 平均 / 件数 / 最大 … label with only numbers / formulas beside it, or aggregate formulas
 *               with a free-text label (志望度4以上 | =COUNTIF(K4:K63,">=4")).
 * 'weak' rows end a table (trailing footer) but are not skipped inside it (a P/L line 売上高合計 =SUM(…) is data there).
 * @param top first data row of the table (aggregate ranges must reach into the data)
 */
export type SummaryKind = 'label' | 'table' | 'aggregate' | 'weak';

export function summaryRowKind(sheet: SheetGrid, r: number, c1: number, c2: number, top = 1, person = false): SummaryKind | null {
  const row = sheet.rows?.[r - 1];
  if (!row) return null;
  if (tableTotalsRows(sheet).has(r)) return 'table';
  let seen = 0;
  let weakFirst = false;
  let weakLatin = false;
  let suffixFirst = false;
  let otherConst = 0;
  let otherText = 0;
  let numeric = 0;
  const formulas: Array<{ f: string; c: number }> = [];
  const end = Math.min(c2, row.length);
  for (let c = c1; c <= end; c++) {
    const cell = row[c - 1];
    if (!cell) continue;
    const v = keyValue(cell);
    const f = typeof cell.f === 'string' && cell.f ? cell.f : undefined;
    if (isBlankValue(v) && !f) continue;
    seen++;
    let label = typeof v === 'string' ? summaryLabelKind(v) : null;
    // a 平均 / 最大 after free text is a VALUE of that row (区分 = 平均, first name "Max"), not a footer label
    if (label === 'weak' && otherText > 0) label = null;
    if (label === 'suffix') {
      // 東京大学 小計: a footer label only when no other text sits in the row (a person's name may end in 計)
      if (seen <= 3 && otherText === 0) suffixFirst = true;
      else otherText++;
      continue;
    }
    if (label) {
      if (seen <= 3) {
        if (label === 'strong') return 'label';
        weakFirst = true;
        // Latin labels (Max / Min / Count / Mean) are also first names: no free text may sit beside them
        if (typeof v === 'string' && /^[A-Za-z]/.test(nfkc(v).trim())) weakLatin = true;
      }
      continue;
    }
    if (f) {
      formulas.push({ f, c });
      continue;
    }
    otherConst++;
    if (typeof v === 'string') otherText++;
    else numeric++;
    // an ordinary data row: too many values for an aggregate footer, no 人数/平均 label in front
    if (seen > 3 && otherConst > Math.max(2, formulas.length + (end - c)) && (!weakFirst || otherText > 1)) return null;
  }
  if (!seen) return null;
  // 東京大学 小計 in a list of PEOPLE is a subtotal line; in another kind of table (売上高合計 in a P/L) it is a row of its own
  if (suffixFirst && otherText === 0 && otherConst <= Math.max(2, formulas.length + 1)) return person ? 'label' : 'weak';
  let agg = 0;
  if (formulas.length && otherConst <= Math.max(2, formulas.length)) {
    // aggregates needed for the verdict below (agg > 0 && otherConst <= max(2, agg)): stop once out of reach
    const need = Math.max(1, otherConst > 2 ? otherConst : 1);
    for (let k = 0; k < formulas.length && agg < need && agg + (formulas.length - k) >= need; k++) {
      if (isFooterAggregate(sheet, formulas[k].f, r, formulas[k].c, top)) agg++;
    }
  }
  if (agg > 0 && otherConst <= Math.max(2, agg)) return otherText === 0 ? 'aggregate' : 'weak';
  if (weakFirst && (weakLatin ? otherText === 0 : otherText <= 1) && numeric + formulas.length > 0) return 'weak';
  return null;
}

/** A 合計 / 小計 / Total / 人数 / 平均 … row (any SummaryKind). */
export function isSummaryRow(sheet: SheetGrid, r: number, c1: number, c2: number, top = 1): boolean {
  return summaryRowKind(sheet, r, c1, c2, top) !== null;
}

/** Summary rows skipped INSIDE a table (not only at its end). */
function isInnerSummaryKind(k: SummaryKind | null): boolean {
  return k === 'label' || k === 'table' || k === 'aggregate';
}

// ---------------------------------------------------------------------------
// region detection
// ---------------------------------------------------------------------------

export interface RegionOptions {
  /** the user picked the header row: columns from that row, data below it */
  headerRow?: number;
  /** the user picked the last data row (clamped to the sheet; no footer trimming) */
  lastDataRow?: number;
}

function median(xs: number[]): number {
  if (!xs.length) return 0;
  const s = [...xs].sort((a, b) => a - b);
  return s[Math.floor(s.length / 2)];
}

/** a row that looks like the header of another table (dictionary hits, texts only) */
function isHeaderLike(sheet: SheetGrid, r: number, c1: number, c2: number): boolean {
  const row = sheet.rows?.[r - 1];
  if (!row) return false;
  let hits = 0;
  let texts = 0;
  let nonBlank = 0;
  const end = Math.min(c2, row.length);
  for (let c = c1; c <= end; c++) {
    const v = keyValue(row[c - 1]);
    if (isBlankValue(v)) continue;
    nonBlank++;
    if (typeof v === 'string') {
      texts++;
      if (isHeaderHit(v)) hits++;
    }
  }
  return nonBlank > 0 && hits >= 2 && texts / nonBlank >= 0.8;
}

// ---------------------------------------------------------------------------
// two-row headers: 「面接評価」 merged over 評価1 / 評価2 / 評価3
// ---------------------------------------------------------------------------

/**
 * true when row `up` + 1 is the SUB-header row of a two-row header whose upper row is `up`: the lower row holds only short
 * texts, and they sit under cells the upper row leaves blank or merges horizontally (a group title over several columns).
 * Key headers may be merged over both rows (学籍番号 A1:A2) or only written in the upper row — both are fine.
 */
export function subHeaderRow(sheet: SheetGrid, up: number): boolean {
  const lo = up + 1;
  if (up < 1 || lo > sheetRows(sheet)) return false;
  const nCols = sheetCols(sheet, up, lo);
  const merges = (sheet.features?.merges ?? []).filter((m) => m.r1 === up && m.r2 === up && m.c2 > m.c1);
  const inGroup = (c: number) => merges.some((m) => c >= m.c1 && c <= m.c2);
  let upNonBlank = 0;
  let loCount = 0;
  let holes = 0;
  let loMin = 0;
  let loMax = 0;
  for (let c = 1; c <= nCols; c++) {
    if (!isBlankValue(keyValue(rawCell(sheet, up, c)))) upNonBlank++;
    const v = keyValue(rawCell(sheet, lo, c));
    if (isBlankValue(v)) continue;
    loMin ||= c;
    loMax = c;
    // a number / date under a header is data; so is an ID-shaped text
    if (typeof v !== 'string') return false;
    const t = nfkc(v).trim();
    if (t.length > 24 || (t.length >= 5 && /\d/.test(t) && /^[A-Za-z0-9][A-Za-z0-9\-_/.#]*$/.test(t))) return false;
    loCount++;
    if (isBlankValue(keyValue(rawCell(sheet, up, c))) || inGroup(c)) holes++;
  }
  if (loCount < 2 || upNonBlank === 0 || holes < loCount * 0.7) return false;
  // a title / note line above the header (one cell, or one merge over the whole width) is not a group header: a group
  // title governs SOME of the columns; without such a merge the upper row must carry two or more texts of its own
  const span = loMax - loMin + 1;
  const group = merges.some((m) => m.c2 - m.c1 + 1 <= Math.max(2, Math.floor(span * 0.7)));
  return group || upNonBlank >= 2;
}

// ---------------------------------------------------------------------------
// paged exports: the header repeated on every page, page-number lines
// ---------------------------------------------------------------------------

const headerKey = (v: Cell): string => (typeof v === 'string' ? nfkc(v).replace(/\s+/g, '').toLowerCase() : '');

/** normalised texts of the header row's cells (index 0 = column c1) — raw cells, so a merged header compares like its copy */
function headerKeys(sheet: SheetGrid, headerRow: number, c1: number, c2: number): string[] {
  const out: string[] = [];
  for (let c = c1; c <= c2; c++) out.push(headerKey(keyValue(rawCell(sheet, headerRow, c))));
  return out;
}

/** a row that repeats the header row (a paged export prints it on every page) */
function isRepeatedHeader(sheet: SheetGrid, r: number, hdr: string[], c1: number, c2: number): boolean {
  const row = sheet.rows?.[r - 1];
  if (!row) return false;
  let nonBlank = 0;
  let same = 0;
  const end = Math.min(c2, row.length);
  for (let c = c1; c <= end; c++) {
    const v = keyValue(row[c - 1]);
    if (isBlankValue(v)) continue;
    nonBlank++;
    const k = headerKey(v);
    if (k && k === hdr[c - c1]) same++;
    else if (nonBlank >= 6 && same === 0) return false; // an ordinary data row: decided after a few cells
  }
  return nonBlank >= 2 && same >= 2 && same >= nonBlank * 0.6;
}

/**
 * a row that is furniture, not a person: its only value is a page-number line ('- 2 -', 'Page 2 of 5', '2/5') or a remark
 * line ('※ 以下は追加分'), or it is the sample row of a form template (例） 山田 太郎 …: the first value is 例 / サンプル)
 */
function isPageLine(sheet: SheetGrid, r: number, c1: number, c2: number): boolean {
  const row = sheet.rows?.[r - 1];
  if (!row) return false;
  let first: Cell = null;
  let n = 0;
  const end = Math.min(c2, row.length);
  for (let c = c1; c <= end; c++) {
    const v = keyValue(row[c - 1]);
    if (isBlankValue(v)) continue;
    if (++n === 1) first = v;
    else if (n > 2 && !isExampleMark(first)) return false;
  }
  if (n === 1) return isPageMarkerText(first) || isNoteLineText(first);
  return n >= 2 && isExampleMark(first);
}

/**
 * Data rows below the header: the table ends at a run of BLANK_RUN_END blank rows, at a note line after a blank row,
 * or at a footer block (summary rows, possibly below 1–3 blank rows) that is not followed by more rows of the same
 * table. Returns the first data row (-1 = none) and the last scanned row (footers / blank rows not trimmed yet).
 */
function scanDataRows(sheet: SheetGrid, headerRow: number, c1: number, c2: number, maxRow: number, person: boolean): { first: number; last: number } {
  let first = -1;
  let last = headerRow;
  const fills: number[] = [];
  const hdr = headerKeys(sheet, headerRow, c1, c2);
  const isNoise = (r: number) => isRepeatedHeader(sheet, r, hdr, c1, c2) || isPageLine(sheet, r, c1, c2);
  /**
   * A paged export goes on after a page break (blank rows, footer/subtotal rows, page-number lines): the next page starts
   * with the repeated header (or, after a page-number line, simply with more rows). Returns that row, or -1.
   */
  const nextPage = (from: number): number => {
    let sawPageLine = false;
    for (let k = from, steps = 0; k <= maxRow && steps < 400; k++, steps++) {
      const nf = rowNonEmpty(sheet, k, c1, c2);
      if (nf === 0) continue;
      if (isPageLine(sheet, k, c1, c2)) {
        sawPageLine = true;
        continue;
      }
      if (isRepeatedHeader(sheet, k, hdr, c1, c2)) return k;
      if (sawPageLine && nf >= Math.max(2, Math.ceil(median(fills) * 0.5)) && !isHeaderLike(sheet, k, c1, c2) && !summaryRowKind(sheet, k, c1, c2, first, person)) return k;
      return -1;
    }
    return -1;
  };
  const blankRunEnd = (from: number): number => {
    let k = from;
    while (k <= maxRow && k - from < BLANK_RUN_END && rowNonEmpty(sheet, k, c1, c2) === 0) k++;
    return k;
  };
  let r = headerRow + 1;
  while (r <= maxRow) {
    const n = rowNonEmpty(sheet, r, c1, c2);
    if (n > 0) {
      if (isNoise(r)) {
        r++; // a repeated header / page-number line is neither data nor the end of the table
        continue;
      }
      if (first < 0) {
        first = r;
      } else if (summaryRowKind(sheet, r, c1, c2, first, person)) {
        // footer block: summary rows, possibly separated by blank rows
        let end = r;
        let next = -1;
        for (;;) {
          const k = blankRunEnd(end + 1);
          if (k > maxRow || k - (end + 1) >= BLANK_RUN_END) break;
          if (!summaryRowKind(sheet, k, c1, c2, first, person)) {
            next = k;
            break;
          }
          end = k;
        }
        last = end;
        const np = nextPage(end + 1);
        if (np >= 0) {
          r = np; // the next page of a paged export
          continue;
        }
        if (next < 0) break;
        // 小計 rows directly followed by more data belong to the table; after a blank gap only rows that look like
        // more data of the same table continue it (not a note line, not another table's header)
        const nf = rowNonEmpty(sheet, next, c1, c2);
        if (next > end + 1) {
          if (nf < Math.max(2, Math.ceil(median(fills) * 0.5)) || isHeaderLike(sheet, next, c1, c2)) break;
        } else if ((nf <= 1 && median(fills) >= 3) || isHeaderLike(sheet, next, c1, c2)) {
          break; // a note line (※ 集計は4/1時点) or another table right below the footer
        }
        r = next;
        continue;
      }
      last = r;
      fills.push(n);
      r++;
      continue;
    }
    if (first < 0) {
      // blank rows between the header and the first data row
      if (r - headerRow > BLANK_RUN_END + 2) break;
      r++;
      continue;
    }
    // blank row inside / after the data: look at the next non-blank row
    const k = blankRunEnd(r);
    if (k > maxRow || k - r >= BLANK_RUN_END) {
      const np = nextPage(r);
      if (np >= 0) {
        r = np;
        continue;
      }
      break;
    }
    if (!summaryRowKind(sheet, k, c1, c2, first, person)) {
      const nextFill = rowNonEmpty(sheet, k, c1, c2);
      if (nextFill <= 1 && median(fills) >= 3) {
        // a note line (※ …) below the list — or a page-number line between two pages
        const np = nextPage(r);
        if (np >= 0) {
          r = np;
          continue;
        }
        break;
      }
    }
    r = k;
  }
  return { first, last };
}

/** Last data row: trailing blank rows and summary rows (and the blank rows above them) removed. */
function trimFooter(sheet: SheetGrid, first: number, last: number, c1: number, c2: number, person: boolean): number {
  let end = last;
  let onlyLabels = true;
  while (end >= first) {
    if (rowNonEmpty(sheet, end, c1, c2) === 0) {
      end--;
      continue;
    }
    const k = summaryRowKind(sheet, end, c1, c2, first, person);
    if (!k) break;
    if (k !== 'label' && k !== 'table') onlyLabels = false;
    end--;
  }
  if (end < first && !onlyLabels) {
    // nothing but 人数 / 平均 … rows: that is the data (a key-value list), not a footer — only blank rows are trimmed
    end = last;
    while (end >= first && rowNonEmpty(sheet, end, c1, c2) === 0) end--;
  }
  return end;
}

/**
 * More rows of the same table BELOW the detected end: a gap of blank rows (3 or more) ends a table, but a list pasted in two
 * pieces carries on after it. Counts the rows below `region.lastDataRow` that are as full as the table's own rows (≥ 70% of
 * their median fill) and are not footers, remark / page lines or another header. null when there are fewer than 3.
 */
export function dataBelowTable(sheet: SheetGrid, region: TableRegion): { count: number; firstRow: number } | null {
  if (region.lastDataRow < region.firstDataRow) return null;
  const c1 = region.firstCol;
  const c2 = region.lastCol;
  const fills: number[] = [];
  for (let r = Math.max(region.firstDataRow, region.lastDataRow - 19); r <= region.lastDataRow; r++) {
    const n = rowNonEmpty(sheet, r, c1, c2);
    if (n > 0) fills.push(n);
  }
  const need = Math.max(2, Math.ceil(median(fills) * 0.7));
  const maxRow = Math.min(sheetRows(sheet), region.lastDataRow + 5000);
  const hdr = headerKeys(sheet, region.headerRow, c1, c2);
  let count = 0;
  let firstRow = 0;
  for (let r = region.lastDataRow + 1; r <= maxRow; r++) {
    const n = rowNonEmpty(sheet, r, c1, c2);
    if (n < need) continue;
    if (isRepeatedHeader(sheet, r, hdr, c1, c2) || isHeaderLike(sheet, r, c1, c2) || summaryRowKind(sheet, r, c1, c2, region.firstDataRow)) continue;
    if (!firstRow) firstRow = r;
    count++;
  }
  return count >= 3 ? { count, firstRow } : null;
}

const MAX_ROWS = 1048576;

function validRow(x: number | undefined): number | undefined {
  return typeof x === 'number' && Number.isFinite(x) && x >= 1 ? Math.min(Math.floor(x), MAX_ROWS) : undefined;
}

/**
 * Header row (dictionary hits in the first 30 rows, not the densest row — or opts.headerRow), first/last data row
 * (the table ends at a run of blank rows, a note line after a blank row, or trailing summary rows — 合計 / 人数 / 平均 …
 * labels and/or aggregate formulas over the column above, also below blank rows — or at opts.lastDataRow) and the
 * contiguous header block's first/last column.
 */
export function detectRegion(book: SourceBook, sheetName: string, opts?: RegionOptions): TableRegion {
  const wantHeader = validRow(opts?.headerRow);
  const wantLast = typeof opts?.lastDataRow === 'number' && Number.isFinite(opts.lastDataRow) ? Math.round(opts.lastDataRow) : undefined;
  const sheet = findSheet(book, sheetName);
  if (!sheet) return wantHeader ? noData(sheetName, wantHeader, 1, 1) : emptyRegion(sheetName);
  const maxRow = sheetRows(sheet);
  if (maxRow <= 0) return wantHeader ? noData(sheetName, wantHeader, 1, 1) : emptyRegion(sheetName);

  let headerRow: number;
  let headerCells: Cell[];
  if (wantHeader) {
    headerRow = wantHeader;
    const nCols = sheetCols(sheet, headerRow, Math.min(maxRow, headerRow + 50));
    headerCells = nCols > 0 ? valueGrid(sheet, headerRow, headerRow, nCols)[0] : [];
  } else {
    const scanEnd = Math.min(maxRow, MAX_HEADER_SCAN + 6);
    const nCols = sheetCols(sheet, 1, Math.min(maxRow, MAX_HEADER_SCAN + 50));
    if (nCols <= 0) return emptyRegion(sheetName);
    const grid = valueGrid(sheet, 1, scanEnd, nCols);
    let h = detectHeaderRow(grid);
    if (h < 0) {
      // single-column lists / tiny sheets: the first non-empty row
      h = grid.findIndex((row) => row.some((v) => !isBlankValue(v)));
      if (h < 0) return emptyRegion(sheetName);
    }
    headerRow = h + 1;
    headerCells = grid[h];
    // a two-row header: the sub-header row (評価1 / 評価2 …) is the header row, the group titles above it are merged in
    if (h + 1 < grid.length && subHeaderRow(sheet, headerRow)) {
      const lo = grid[h + 1];
      headerCells = headerCells.map((u, c) => (isBlankValue(lo[c] ?? null) ? u : lo[c]));
      headerRow += 1;
    }
  }

  const hasDataBelow = (c: number) => {
    for (let r = headerRow + 1; r <= Math.min(maxRow, headerRow + 200); r++) if (!isBlankValue(keyValue(rawCell(sheet, r, c)))) return true;
    return false;
  };
  const blocks = headerBlocks(headerCells, hasDataBelow);
  let firstCol = 1;
  let lastCol = 1;
  if (blocks.length) {
    let best = blocks[0];
    for (const b of blocks) {
      if (b.hits > best.hits || (b.hits === best.hits && b.filled > best.filled)) best = b;
    }
    firstCol = best.c1;
    lastCol = best.c2;
  } else if (wantHeader) {
    // a blank header row picked by the user: the columns the data below uses
    const ext = dataColumns(sheet, headerRow + 1, Math.min(maxRow, headerRow + 200));
    if (ext) [firstCol, lastCol] = ext;
  }

  // ---- data rows ----
  if (wantLast !== undefined) {
    const last = Math.max(headerRow, Math.min(wantLast, Math.max(maxRow, headerRow + 1)));
    let first = headerRow + 1;
    for (let r = headerRow + 1; r <= Math.min(last, maxRow); r++) {
      if (rowNonEmpty(sheet, r, firstCol, lastCol) > 0) {
        first = r;
        break;
      }
    }
    return { sheet: sheetName, headerRow, firstDataRow: first, lastDataRow: last, firstCol, lastCol };
  }
  if (headerRow >= maxRow) return noData(sheetName, headerRow, firstCol, lastCol);
  const person = isPersonList(headerCells.slice(firstCol - 1, lastCol));
  const { first, last } = scanDataRows(sheet, headerRow, firstCol, lastCol, maxRow, person);
  if (first < 0) return noData(sheetName, headerRow, firstCol, lastCol);
  const lastDataRow = trimFooter(sheet, first, last, firstCol, lastCol, person);
  return { sheet: sheetName, headerRow, firstDataRow: first, lastDataRow: Math.max(lastDataRow, headerRow), firstCol, lastCol };
}

/** header row without data rows below it */
function noData(sheet: string, headerRow: number, firstCol: number, lastCol: number): TableRegion {
  return { sheet, headerRow, firstDataRow: headerRow + 1, lastDataRow: headerRow, firstCol, lastCol };
}

/** first / last column with a value in rows r1..r2 */
function dataColumns(sheet: SheetGrid, r1: number, r2: number): [number, number] | null {
  let lo = 0;
  let hi = 0;
  for (let r = r1; r <= r2; r++) {
    const row = sheet.rows?.[r - 1];
    if (!row) continue;
    const n = Math.min(row.length, MAX_SCAN_COLS);
    for (let c = 1; c <= n; c++) {
      if (isBlankValue(keyValue(row[c - 1]))) continue;
      if (!lo || c < lo) lo = c;
      if (c > hi) hi = c;
    }
  }
  return lo ? [lo, hi] : null;
}

/** The most list-like visible sheet (dictionary hits in its header, ID/氏名 headers, rows; 横持ち sheets penalised). */
export function pickDataSheet(book: SourceBook): string {
  const sheets = book?.sheets ?? [];
  if (!sheets.length) return '';
  const visible = sheets.filter((s) => s.state === 'visible' || s.state === undefined);
  const pool = visible.length ? visible : sheets;
  let best = pool[0].name;
  let bestScore = -Infinity;
  for (const s of pool) {
    let score: number;
    try {
      score = sheetListScore(book, s);
    } catch {
      score = -1000;
    }
    if (score > bestScore) {
      bestScore = score;
      best = s.name;
    }
  }
  return best;
}

function sheetListScore(book: SourceBook, s: SheetGrid): number {
  const region = detectRegion(book, s.name);
  const rows = Math.max(0, region.lastDataRow - region.firstDataRow + 1);
  if (rows === 0) return -500 + (region.lastCol - region.firstCol);
  const width = region.lastCol - region.firstCol + 1;
  let hits = 0;
  let keyHeader = 0;
  const headers: Cell[] = [];
  for (let c = region.firstCol; c <= region.lastCol; c++) {
    const v = keyValue(rawCell(s, region.headerRow, c));
    headers.push(v);
    if (typeof v !== 'string') continue;
    const hn = headerNorm(nfkc(v));
    const b = bestHeaderType(hn);
    if (b.score >= 0.7 || isNonKeyHeader(hn)) hits++;
    if (b.score >= 0.7 && b.type && ['id', 'studentId', 'name', 'nameKana', 'email', 'lastName'].includes(b.type)) keyHeader++;
  }
  let score = (keyHeader ? 100 : 0) + keyHeader * 10 + hits * 5 + Math.min(rows, 2000) / 100 + (width >= 2 ? 5 : 0);
  // persons laid out across columns (横持ち): row-wise keys are meaningless there
  const scan = valueGrid(s, Math.max(1, region.headerRow - 1), region.headerRow, Math.min(region.lastCol, MAX_SCAN_COLS));
  const hint = wideLayoutHint(scan.map((row) => row.slice(region.firstCol - 1)), scan.length - 1, width);
  if (hint) score -= 200;
  return score;
}

// ---------------------------------------------------------------------------
// KeySource → Table
// ---------------------------------------------------------------------------

export function colLetter(c: number): string {
  let s = '';
  let n = c;
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/** Region clamped to the sheet (inverted / out-of-range regions become empty). */
export function clampRegion(sheet: SheetGrid | undefined, region: TableRegion): TableRegion {
  const r = { ...region };
  const toInt = (x: number, d: number) => (Number.isFinite(x) ? Math.floor(x) : d);
  r.headerRow = Math.max(1, toInt(r.headerRow, 1));
  r.firstCol = Math.max(1, toInt(r.firstCol, 1));
  r.lastCol = Math.min(Math.max(r.firstCol, toInt(r.lastCol, r.firstCol)), r.firstCol + MAX_SCAN_COLS);
  r.firstDataRow = Math.max(r.headerRow + 1, toInt(r.firstDataRow, r.headerRow + 1));
  const maxRow = sheet ? sheetRows(sheet) : 0;
  r.lastDataRow = Math.min(toInt(r.lastDataRow, r.headerRow), maxRow);
  if (r.lastDataRow < r.firstDataRow) r.lastDataRow = r.firstDataRow - 1;
  return r;
}

export interface SourceTable extends Table {
  /** Excel rows of the region that are blank / 合計 rows (kept in place, never matched) */
  skippedRows: number[];
  /** raw contract cell of table row i / table column j (for display and notes) */
  raw: (row: number, col: number) => XCell | undefined;
  region: TableRegion;
  /** key cells that are formulas without a cached result (keys are blank there) */
  formulasWithoutValue: number;
  /** text of the header row's own cell per column ('' for the continuation of a merge / a blank cell) */
  headerRaw: string[];
  /** headers[j] is the made-up 列C of a blank header cell */
  headerFallback: boolean[];
}

function mode<T>(xs: T[]): T | undefined {
  const m = new Map<T, number>();
  let best: T | undefined;
  let bn = 0;
  for (const x of xs) {
    const n = (m.get(x) ?? 0) + 1;
    m.set(x, n);
    if (n > bn) {
      bn = n;
      best = x;
    }
  }
  return best;
}

/** Build the R6 Table of one source (rows = matchable data rows of the region). */
export function buildTable(src: KeySource): SourceTable {
  const book = src.book;
  const sheet = findSheet(book, src.region.sheet);
  const region = clampRegion(sheet, src.region);
  const cols: number[] = [];
  if (sheet) for (let c = region.firstCol; c <= region.lastCol; c++) cols.push(c);
  const width = cols.length;

  // header texts (a merged header cell gives its top-left value to the whole merge)
  const headerCells = sheet ? valueGrid(sheet, region.headerRow, region.headerRow, region.lastCol)[0] : [];
  const headers = cols.map((c) => headerText(headerCells[c - 1] ?? null) || `列${colLetter(c)}`);
  // the text of the header cell itself (no merge spread, no two-row join): what a premerge would copy
  const headerRaw = cols.map((c) => (sheet ? headerText(keyValue(rawCell(sheet, region.headerRow, c))) : ''));
  const headerFallback = cols.map((c, j) => headers[j] === `列${colLetter(c)}` && headerRaw[j] === '' && isBlankValue(headerCells[c - 1] ?? null));
  // two-row header: a group title above the sub-header joins it ('面接評価 評価1'); a title merged over both rows stays as is
  if (sheet && region.headerRow > 1 && subHeaderRow(sheet, region.headerRow - 1)) {
    const upCells = valueGrid(sheet, region.headerRow - 1, region.headerRow - 1, region.lastCol)[0];
    cols.forEach((c, j) => {
      const lo = headerText(headerCells[c - 1] ?? null);
      const up = headerText(upCells[c - 1] ?? null);
      const t = !lo ? up : !up || up === lo || lo.includes(up) ? lo : `${up} ${lo}`;
      if (t) {
        headers[j] = t;
        headerFallback[j] = false;
      }
    });
  }

  const rows: Cell[][] = [];
  const rawRows: Array<Array<XCell | undefined> | undefined> = [];
  const excelRow: number[] = [];
  const skippedRows: number[] = [];
  let formulasWithoutValue = 0;
  // xlsx books in the 1904 date system keep raw 1904 serials (legacy / csv readers convert them to 1900 ones)
  const shift1904 = book.date1904 && book.fidelity === 'full' ? 1462 * 86400000 : 0;
  const rowIndexOf = new Map<number, number>();
  if (sheet && width) {
    // trailing footer rows inside the region (a user-chosen last row may include 人数 / 平均 rows) are skipped like
    // 合計 rows; inside the table only 合計-type / pure aggregate rows are (see summaryRowKind)
    const hdrKeys = headerKeys(sheet, region.headerRow, region.firstCol, region.lastCol);
    const person = isPersonList(headers);
    const tailStart = region.lastDataRow >= region.firstDataRow
      ? trimFooter(sheet, region.firstDataRow, region.lastDataRow, region.firstCol, region.lastCol, person) + 1
      : region.firstDataRow;
    for (let r = region.firstDataRow; r <= region.lastDataRow; r++) {
      const src = sheet.rows?.[r - 1];
      const vals: Cell[] = new Array(width);
      let blank = true;
      for (let j = 0; j < width; j++) {
        const cell = src?.[cols[j] - 1];
        let v = keyValue(cell);
        if (shift1904 && v instanceof Date) v = new Date(v.getTime() + shift1904);
        vals[j] = v;
        if (!isBlankValue(v)) blank = false;
        else if (cell?.f && (cell.v === null || cell.v === undefined || cell.t === 'z')) formulasWithoutValue++;
      }
      const kind = blank ? null : summaryRowKind(sheet, r, region.firstCol, region.lastCol, region.firstDataRow, person);
      // blank rows, footer rows, and the repeated header / page-number lines of a paged export are never matched
      if (blank || (r >= tailStart ? kind !== null : isInnerSummaryKind(kind)) || isRepeatedHeader(sheet, r, hdrKeys, region.firstCol, region.lastCol) || isPageLine(sheet, r, region.firstCol, region.lastCol)) {
        skippedRows.push(r);
        continue;
      }
      rowIndexOf.set(r, rows.length);
      rows.push(vals);
      rawRows.push(src);
      excelRow.push(r);
    }
    // a merged cell (大学 merged over the rows of its students) is the value of every row it covers
    for (const m of sheet.features?.merges ?? []) {
      if (m.r2 <= m.r1 || m.r2 < region.firstDataRow || m.r1 > region.lastDataRow || m.c2 < region.firstCol || m.c1 > region.lastCol) continue;
      let top = keyValue(rawCell(sheet, m.r1, m.c1));
      if (isBlankValue(top)) continue;
      if (shift1904 && top instanceof Date) top = new Date(top.getTime() + shift1904);
      for (let r = Math.max(m.r1, region.firstDataRow); r <= Math.min(m.r2, region.lastDataRow); r++) {
        const i = rowIndexOf.get(r);
        if (i === undefined) continue;
        for (let c = Math.max(m.c1, region.firstCol); c <= Math.min(m.c2, region.lastCol); c++) {
          if (isBlankValue(rows[i][c - region.firstCol])) rows[i][c - region.firstCol] = top;
        }
      }
    }
  }

  const raw = (i: number, j: number): XCell | undefined => rawRows[i]?.[cols[j] - 1];
  const dispCache = new Map<number, string>();
  const display = (i: number, j: number): string => {
    const k = i * (width + 1) + j;
    let s = dispCache.get(k);
    if (s === undefined) {
      const cell = raw(i, j);
      s = cell ? displayText(book, cell) : cellText(rows[i]?.[j] ?? null);
      dispCache.set(k, s);
    }
    return s;
  };

  // number format per column: mode over (a sample of) the data cells
  const numFmt: string[] = [];
  const dateFmt: boolean[] = [];
  const step = Math.max(1, Math.floor(rows.length / 400));
  for (let j = 0; j < width; j++) {
    const fmts: string[] = [];
    let dCells = 0;
    let nCells = 0;
    for (let i = 0; i < rows.length; i += step) {
      const cell = raw(i, j);
      if (!cell || cell.v === null || cell.v === undefined) continue;
      nCells++;
      if (cell.t === 'd') dCells++;
      fmts.push(numFmtOf(book, cell));
    }
    const f = mode(fmts) ?? 'General';
    numFmt.push(f);
    let isDate = false;
    try {
      isDate = isDateFormat(f);
    } catch {
      isDate = false;
    }
    dateFmt.push(isDate || (nCells > 0 && dCells / nCells >= 0.5));
  }

  const t: SourceTable = {
    file: src.label || book.fileName || src.id,
    sheet: region.sheet,
    headerRow: region.headerRow,
    headers,
    cols,
    rows,
    display,
    numFmt,
    dateFmt,
    excelRow,
    skippedRows,
    raw,
    region,
    formulasWithoutValue,
    headerRaw,
    headerFallback,
  };
  if (sheet && width) {
    const top = Math.max(1, region.headerRow - 1);
    const scan = valueGrid(sheet, top, region.headerRow, region.lastCol).map((row) => row.slice(region.firstCol - 1));
    t.layoutHint = wideLayoutHint(scan, scan.length - 1, width, top - 1);
  }
  return t;
}
