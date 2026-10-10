/**
 * Worksheet model for the package patcher (owner: patcher P).
 * Only <sheetData> is parsed into rows/cells (string attributes kept verbatim); everything before it (`head`:
 * dimension, sheetViews, cols ...) and after it (`tail`: mergeCells, CF, DV, hyperlinks, extLst ...) stays text
 * and is rewritten element by element by the struct engine. Serialization writes cells in Excel's canonical form
 * (<c r="A1" ...>) and drops row `spans` (optional; Excel recomputes them).
 */
import { MAX_COL, MAX_ROW, colToNum, numToCol, translateFormula } from '../../formula/excelRef';
import { PatchError } from './pkg';
import { esc, getAttr, setAttr, unesc } from './xml';

export interface CellM {
  /** 1-based column */
  c: number;
  /** attributes without r, each with a leading space (' s="3" t="s"') */
  a: string;
  /** inner xml; null = self-closing */
  x: string | null;
  /** index of the first op entry that still has to be applied to this cell's formula */
  ep: number;
}

export interface RowM {
  r: number;
  /** attributes without r / spans, each with a leading space */
  a: string;
  cells: CellM[];
}

export interface ParsedF {
  /** attributes of <f> (leading spaces) */
  attrs: string;
  /** formula text (unescaped); '' for a self-closing <f/> */
  text: string;
  /** index of <f in the inner xml and end index */
  start: number;
  end: number;
}

const reRow = /<row\b([^>]*?)(\/>|>([\s\S]*?)<\/row>)/g;
const reCell = /<c(?=[\s/>])([^>]*?)(\/>|>([\s\S]*?)<\/c>)/g;
const reF = /<f\b([^>]*?)(\/>|>([\s\S]*?)<\/f>)/;

export function parseF(x: string | null): ParsedF | null {
  if (!x || x.indexOf('<f') < 0) return null;
  const m = reF.exec(x);
  if (!m) return null;
  return { attrs: m[1], text: m[3] !== undefined ? unesc(m[3]) : '', start: m.index, end: m.index + m[0].length };
}

export function cellRef(c: number, r: number): string {
  return numToCol(c) + r;
}

/** Parse "AB12" -> {c, r} */
export function parseCellRef(ref: string): { c: number; r: number } | null {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(ref);
  if (!m) return null;
  return { c: colToNum(m[1]), r: Number(m[2]) };
}

export function cellStyle(cell: CellM): number {
  const s = getAttr(cell.a, 's');
  return s ? Number(s) || 0 : 0;
}

export class SheetModel {
  head: string;
  tail: string;
  /** the <sheetData ...> open tag */
  sdOpen: string;
  rows: RowM[];
  /** true once cells/rows changed (serialize needed) */
  dirty = false;
  /**
   * Something in the input that the text patching cannot carry through (two cells at one position, a cell whose
   * reference lies in another row or beyond XFD / row 1,048,576 ...): reading is tolerant, but serializing such a
   * sheet is refused with a PatchError (writing it back would silently move or duplicate cells).
   */
  problem: string | null = null;
  private rowIndex: Map<number, RowM> | null = null;

  constructor(public xml: string) {
    const s = findSheetData(xml);
    if (s < 0) {
      // no <sheetData> (empty sheet written by a sloppy generator): insert an empty one where the schema wants it,
      // before the first element that follows the sheet data (mergeCells, conditionalFormatting, pageMargins ...)
      const at = firstAfterSheetData(xml);
      this.head = xml.slice(0, at);
      this.tail = xml.slice(at);
      this.sdOpen = '<sheetData>';
      this.rows = [];
      this.dirty = true;
      return;
    }
    const openEnd = xml.indexOf('>', s);
    if (openEnd < 0) throw new PatchError('シートのデータが途中で切れているため処理できません（ファイルが壊れています）。');
    const open = xml.slice(s, openEnd + 1);
    this.head = xml.slice(0, s);
    if (open.endsWith('/>')) {
      this.sdOpen = open.slice(0, -2).replace(/\s+$/, '') + '>';
      this.tail = xml.slice(openEnd + 1);
      this.rows = [];
      return;
    }
    this.sdOpen = open;
    const close = xml.indexOf('</sheetData>', openEnd);
    if (close < 0) throw new PatchError('シートのデータが途中で切れているため処理できません（ファイルが壊れています）。');
    const body = xml.slice(openEnd + 1, close);
    this.tail = xml.slice(close + '</sheetData>'.length);
    const parsed = parseRows(body);
    this.rows = parsed.rows;
    this.problem = parsed.problem;
  }

  /** Find a row (binary search; rows are sorted). */
  row(r: number): RowM | undefined {
    if (!this.rowIndex) {
      this.rowIndex = new Map();
      for (const row of this.rows) this.rowIndex.set(row.r, row);
    }
    return this.rowIndex.get(r);
  }

  invalidateIndex(): void {
    this.rowIndex = null;
  }

  /** Get or create a row. */
  ensureRow(r: number): RowM {
    const ex = this.row(r);
    if (ex) return ex;
    const nr: RowM = { r, a: '', cells: [] };
    let lo = 0, hi = this.rows.length;
    while (lo < hi) { const mid = (lo + hi) >> 1; if (this.rows[mid].r < r) lo = mid + 1; else hi = mid; }
    this.rows.splice(lo, 0, nr);
    this.rowIndex?.set(r, nr);
    this.dirty = true;
    return nr;
  }

  cell(r: number, c: number): CellM | undefined {
    const row = this.row(r);
    if (!row) return undefined;
    return findCell(row, c);
  }

  ensureCell(r: number, c: number): CellM {
    const row = this.ensureRow(r);
    const cells = row.cells;
    let lo = 0, hi = cells.length;
    while (lo < hi) { const mid = (lo + hi) >> 1; if (cells[mid].c < c) lo = mid + 1; else hi = mid; }
    if (lo < cells.length && cells[lo].c === c) return cells[lo];
    const nc: CellM = { c, a: '', x: null, ep: 0 };
    cells.splice(lo, 0, nc);
    this.dirty = true;
    return nc;
  }

  /** Expand every shared formula into a plain formula (translateFormula of the master). */
  unshare(): boolean {
    const masters = new Map<string, { r: number; c: number; text: string }>();
    let any = false;
    for (const row of this.rows) {
      for (const cell of row.cells) {
        const f = parseF(cell.x);
        if (!f || !/\bt\s*=\s*"shared"/.test(f.attrs)) continue;
        any = true;
        const si = getAttr(f.attrs, 'si');
        if (si != null && getAttr(f.attrs, 'ref') != null && f.text) masters.set(si, { r: row.r, c: cell.c, text: f.text });
      }
    }
    if (!any) return false;
    for (const row of this.rows) {
      for (const cell of row.cells) {
        const f = parseF(cell.x);
        if (!f || !/\bt\s*=\s*"shared"/.test(f.attrs)) continue;
        const si = getAttr(f.attrs, 'si');
        const m = si != null ? masters.get(si) : undefined;
        let text = f.text;
        if (m && !(m.r === row.r && m.c === cell.c)) text = translateFormula(m.text, row.r - m.r, cell.c - m.c);
        let attrs = setAttr(setAttr(setAttr(f.attrs, 't', null), 'ref', null), 'si', null);
        if (!/\S/.test(attrs)) attrs = '';
        const nf = text ? `<f${attrs}>${esc(text)}</f>` : '';
        cell.x = cell.x!.slice(0, f.start) + nf + cell.x!.slice(f.end);
      }
    }
    this.dirty = true;
    return true;
  }

  /** Used range from the cells ("A1" when empty). */
  usedRange(): string {
    let r1 = Infinity, r2 = 0, c1 = Infinity, c2 = 0;
    for (const row of this.rows) {
      if (!row.cells.length) continue;
      if (row.r < r1) r1 = row.r;
      if (row.r > r2) r2 = row.r;
      const a = row.cells[0].c, b = row.cells[row.cells.length - 1].c;
      if (a < c1) c1 = a;
      if (b > c2) c2 = b;
    }
    if (!r2) return 'A1';
    const tl = cellRef(c1, r1), br = cellRef(c2, r2);
    return tl === br ? tl : tl + ':' + br;
  }

  serializeRows(): string {
    const parts: string[] = [];
    for (const row of this.rows) {
      if (!row.cells.length && !/\S/.test(row.a)) continue; // empty row without attributes: drop
      const cells = row.cells;
      if (!cells.length) { parts.push(`<row r="${row.r}"${row.a}/>`); continue; }
      let s = `<row r="${row.r}"${row.a}>`;
      const rs = String(row.r);
      for (const c of cells) {
        s += `<c r="${numToCol(c.c)}${rs}"${c.a}` + (c.x == null || c.x === '' ? '/>' : `>${c.x}</c>`);
      }
      parts.push(s + '</row>');
    }
    return parts.join('');
  }

  /**
   * Refuses (PatchError, Japanese) a sheet whose input had something that cannot be written back (see `problem`).
   * The engine calls it before an op edits the sheet so that the message names the real cause instead of a follow-up
   * symptom (a cell beyond XFD would otherwise be reported as "data pushed off the sheet" by an insert).
   */
  assertWritable(sheetName = ''): void {
    if (this.problem) {
      throw new PatchError(`${sheetName ? `シート「${sheetName}」` : 'シート'}の内容に、そのまま書き戻せない部分があります（${this.problem}）。ファイルが壊れている可能性があります。Excel でファイルを開いて保存し直してから、もう一度お試しください。`);
    }
  }

  /**
   * The sheet XML. `sheetName` is only used for the message when the input had something that cannot be written back
   * (see `problem`).
   */
  serialize(sheetName = ''): string {
    this.assertWritable(sheetName);
    const body = this.serializeRows();
    const sd = body ? `${this.sdOpen}${body}</sheetData>` : this.sdOpen.replace(/>$/, '/>');
    return this.head + sd + this.tail;
  }
}

/** Index of the <sheetData> start tag (not <sheetDataX>), -1 when the sheet has none. */
function findSheetData(xml: string): number {
  let i = xml.indexOf('<sheetData');
  while (i >= 0) {
    const c = xml.charCodeAt(i + 10);
    if (c === 62 || c === 47 || c === 32 || c === 9 || c === 10 || c === 13) return i;
    i = xml.indexOf('<sheetData', i + 10);
  }
  return -1;
}

/** Elements that follow <sheetData> in the worksheet schema (CT_Worksheet), without extLst. */
const AFTER_SHEET_DATA = /<(?:sheetCalcPr|sheetProtection|protectedRanges|scenarios|autoFilter|sortState|dataConsolidate|customSheetViews|mergeCells|phoneticPr|conditionalFormatting|dataValidations|hyperlinks|printOptions|pageMargins|pageSetup|headerFooter|rowBreaks|colBreaks|customProperties|cellWatches|ignoredErrors|smartTags|drawing|legacyDrawing|legacyDrawingHF|drawingHF|picture|oleObjects|controls|webPublishItems|tableParts)(?=[\s/>])/;

/** Where a missing <sheetData> belongs: before the first element that the schema puts after it. */
function firstAfterSheetData(xml: string): number {
  const m = AFTER_SHEET_DATA.exec(xml);
  if (m) return m.index;
  // only an <extLst> of the worksheet itself (after the sheet views) or nothing at all
  const views = xml.lastIndexOf('</sheetViews>');
  const ext = xml.lastIndexOf('<extLst');
  if (ext >= 0 && ext > views) return ext;
  const end = xml.lastIndexOf('</worksheet>');
  return end >= 0 ? end : xml.length;
}

export function findCell(row: RowM, c: number): CellM | undefined {
  const cells = row.cells;
  let lo = 0, hi = cells.length;
  while (lo < hi) { const mid = (lo + hi) >> 1; if (cells[mid].c < c) lo = mid + 1; else hi = mid; }
  return lo < cells.length && cells[lo].c === c ? cells[lo] : undefined;
}

function parseRows(body: string): { rows: RowM[]; problem: string | null } {
  const rows: RowM[] = [];
  let problem: string | null = null;
  let prevR = 0;
  reRow.lastIndex = 0;
  let m: RegExpExecArray | null;
  while ((m = reRow.exec(body))) {
    let attrs = m[1];
    const rAttr = /\sr\s*=\s*(["'])(\d+)\1/.exec(attrs);
    const r = rAttr ? Number(rAttr[2]) : prevR + 1;
    if (!(r >= 1 && r <= MAX_ROW)) problem ??= `${r} 行目は Excel の範囲外です`;
    else if (/\sr\s*=/.test(attrs) && !rAttr) problem ??= '行番号の書き方が不正です';
    prevR = r;
    attrs = attrs.replace(/\sr\s*=\s*(["'])\d+\1/, '').replace(/\sspans\s*=\s*(["'])[^"']*\1/, '');
    attrs = attrs.replace(/\s+$/, '');
    const cells: CellM[] = [];
    const inner = m[3];
    if (inner) {
      let prevC = 0;
      reCell.lastIndex = 0;
      let cm: RegExpExecArray | null;
      let sorted = true;
      while ((cm = reCell.exec(inner))) {
        let ca = cm[1];
        let c = 0;
        let fast = false;
        // fast path (what Excel and most writers produce): ` r="AB12"` first, upper-case letters, then the row digits
        if (ca.charCodeAt(0) === 32 && ca.charCodeAt(1) === 114 && ca.charCodeAt(2) === 61 && ca.charCodeAt(3) === 34) {
          const q = ca.indexOf('"', 4);
          let k = 4;
          let col = 0;
          while (k < q) { const ch = ca.charCodeAt(k); if (ch < 65 || ch > 90) break; col = col * 26 + ch - 64; k++; }
          let rowNo = 0;
          let d = k;
          while (d < q) { const ch = ca.charCodeAt(d); if (ch < 48 || ch > 57) break; rowNo = rowNo * 10 + ch - 48; d++; }
          if (q > 0 && k - 4 >= 1 && k - 4 <= 3 && d === q && d > k && d - k <= 7) {
            if (!(col >= 1 && col <= MAX_COL)) problem ??= `${ca.slice(4, k)}${rowNo} は Excel の範囲（XFD 列）外のセルです`;
            else if (rowNo !== r) problem ??= `${r} 行目の中に ${ca.slice(4, q)} のセルがあります`;
            c = col;
            ca = ca.slice(q + 1);
            fast = true;
          }
        }
        if (!fast) {
          const rr = /\sr\s*=\s*(["'])\$?([A-Za-z]+)\$?(\d*)\1/.exec(ca);
          c = rr ? colToNum(rr[2]) : prevC + 1;
          if (rr) {
            if (!(c >= 1 && c <= MAX_COL)) problem ??= `${rr[2]}${rr[3]} は Excel の範囲（XFD 列）外のセルです`;
            else if (rr[3] !== '' && Number(rr[3]) !== r) problem ??= `${r} 行目の中に ${rr[2]}${rr[3]} のセルがあります`;
          } else if (/\sr\s*=/.test(ca)) problem ??= 'セル番地の書き方が不正です';
          if (rr) ca = ca.slice(0, rr.index) + ca.slice(rr.index + rr[0].length);
        }
        if (!Number.isFinite(c) || c < 1) c = prevC + 1;
        if (c <= prevC) sorted = false;
        prevC = c;
        if (ca.length && ca.charCodeAt(ca.length - 1) <= 32) ca = ca.replace(/\s+$/, '');
        cells.push({ c, a: ca, x: cm[3] !== undefined ? cm[3] : null, ep: 0 });
      }
      if (!sorted) {
        cells.sort((a, b) => a.c - b.c);
        for (let i = 1; i < cells.length; i++) if (cells[i].c === cells[i - 1].c) { problem ??= `${cellRef(cells[i].c, r)} のセルが重複しています`; break; }
      }
    }
    rows.push({ r, a: attrs, cells });
  }
  let sorted = true;
  for (let i = 1; i < rows.length; i++) if (rows[i].r <= rows[i - 1].r) { sorted = false; break; }
  if (!sorted) {
    rows.sort((a, b) => a.r - b.r);
    for (let i = 1; i < rows.length; i++) if (rows[i].r === rows[i - 1].r) { problem ??= `${rows[i].r} 行目が重複しています`; break; }
  }
  return { rows, problem };
}
