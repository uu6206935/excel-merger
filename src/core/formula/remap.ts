/**
 * Formula remapping for imported columns (owner: patcher P, round 8 task F1).
 *
 * buildPremerged copies the cells of a source workbook (the 2nd, 3rd … dropped file) into the base sheet. A formula of the source
 * reads cells of the SOURCE sheet; in the merged sheet the same data lives somewhere else (other columns, other rows because
 * the rows are matched by key, not by position). `FormulaRemapper.remap` rewrites such a formula so that it reads the same data
 * in the merged sheet, or says precisely why it cannot (the caller then imports the cached value).
 *
 * Everything works on excelRef tokens (tokenize / parsePrefix / renderArea): only reference tokens (and the names / structured
 * references that are replaced by the references they stand for) are rewritten, `$` flags and every other character of the formula
 * are kept as they are. The result is FILE form (no '='), unqualified (the merged sheet is the formula's own sheet).
 *
 * Rules (see docs/ARCHITECTURE.md "Round 8"):
 *  a. Row mapping T(r): every source data row r that is placed in the merged sheet (matched base row, or a row appended by join 'full')
 *     has the merged row T(r). A reference to ONE row is rewritten by T (the formula's own row is the special case T(sr) = target; the
 *     header row maps to the base header row). A row of the source that was not placed has no T: the reference cannot be remapped.
 *  b. Ranges over rows. A range that spans the whole data table (first data row … last data row, optionally from the header row, or a
 *     whole-column reference, or the Excel-user habit D2:D1000 that runs past the data) becomes the whole merged data table
 *     (first … last merged data row incl. appended rows; whole columns stay whole columns). Why that is right: the merged sheet IS
 *     the table the formulas now belong to — an average / total / rank / lookup over "the table" must see the table as it is after the
 *     merge (rows of the source that matched nothing are not in it, appended rows are), and Excel recalculates it on load. A partial
 *     range [a..b] is kept only when T(a)..T(b) are exactly consecutive and ascending (then it is [T(a)..T(b)] and reads the same
 *     rows); anything else (running totals after the rows were put in another order …) cannot be remapped.
 *  c. Columns: imported columns map by `colMap` (a multi-column area needs consecutive merged columns in the same order, so a
 *     VLOOKUP column index stays right). A non-imported source column that the caller proved equal to a base column on every matched
 *     row (key columns: 氏名 / ID, `partnerCols`) reads that base column. Other non-imported columns: 'notImported'.
 *  d. Defined names (workbook scope or the data sheet's scope) whose refersTo is one reference to the data sheet, one structured
 *     reference, or a constant are inlined; every other name cannot be remapped. Table structured references of the source
 *     (this row, [col], [#Data], [#Headers], the whole table inside a function, a column range) are resolved to A1 references on the
 *     source sheet and then remapped like any other reference.
 *  e. Functions: lookup / aggregate functions are fine when all their references remap. ROW / COLUMN / OFFSET / INDIRECT / ADDRESS /
 *     CELL / FORMULATEXT … depend on where things are: 'positional'. ROWS / COLUMNS with an explicit reference are fine.
 *  g. `exact`: the cached result of the source stays valid (every reference reads the very same source values). It is false when a
 *     whole-table range is not the same set of rows any more (rows of the source not placed), when a position-sensitive function
 *     (MATCH / INDEX / ROWS …) reads a table whose layout changed, or when a substituted key column / renamed header is read.
 */
import { MAX_ROW, parseAreaText, renderArea, tokenize, type Area, type Token } from './excelRef';
import { funcName, matchBrackets, parseCalls, trimRange, type CallInfo } from './calls';
import { POSITIONAL_FUNCTIONS, isBuiltinFunction } from './forms';

export type RemapReason =
  | 'otherSheet'
  | 'externalCell'
  | 'rowRelation'
  | 'unresolvedName'
  | 'unresolvedTable'
  | 'positional'
  | 'notImported'
  | 'unsafe'
  | 'array'
  | 'unsupported';

export interface RemapTable {
  name: string;
  /** sheet the table lives on */
  sheet: string;
  /** A1 range incl. header (and totals) row */
  ref: string;
  headerRow: boolean;
  totalsRow: boolean;
  columns: string[];
}

export interface RemapName {
  name: string;
  /** sheet name of a sheet-scoped name, null for workbook scope */
  scope: string | null;
  /** FILE form formula text, with or without '=' */
  refersTo: string;
}

export interface RemapSpec {
  /** the source data sheet */
  sheet: string;
  /** the source data table (firstCol / lastCol: references entirely outside these columns read cells outside the table: 'externalCell') */
  region: { headerRow: number; firstDataRow: number; lastDataRow: number; firstCol?: number; lastCol?: number };
  /** the data table of the merged sheet (appended rows included in lastDataRow) */
  target: { headerRow: number; firstDataRow: number; lastDataRow: number };
  /** imported source column -> merged column */
  colMap: ReadonlyMap<number, number>;
  /** non-imported source column -> merged column of the base column that holds the same text on every matched row */
  partnerCols?: ReadonlyMap<number, number>;
  /** [source row, merged row] of every placed source row (matched or appended); a source row may appear twice (duplicate in the base) */
  rowPairs: ReadonlyArray<readonly [number, number]>;
  /** source rows of the data table that hold data (default: every row firstDataRow..lastDataRow) */
  dataRows?: ReadonlyArray<number>;
  names?: ReadonlyArray<RemapName>;
  tables?: ReadonlyArray<RemapTable>;
  /** does column `col` hold anything (value or formula) in the source rows r1..r2? Only asked for rows OUTSIDE [headerRow, lastDataRow] */
  hasContentOutside?: (col: number, r1: number, r2: number) => boolean;
  /** is the merged header text of imported source column `col` the source's header text? (a ref to the header cell is exact only then) */
  headerSame?: (col: number) => boolean;
}

export interface RemapOk {
  ok: true;
  /** FILE form, no '=' */
  formula: string;
  /** the cached result of the source is still the result of this formula */
  exact: boolean;
  /** a whole-table range was rewritten to a merged table that is not the source table (rows not placed / added rows) */
  widened: boolean;
  /** a non-imported key column was replaced by the base's partner column */
  substituted: boolean;
}

export interface RemapFail {
  ok: false;
  reason: RemapReason;
  detail?: string;
}

export type RemapResult = RemapOk | RemapFail;

class Fail extends Error {
  constructor(readonly reason: RemapReason, readonly detail?: string) {
    super(reason);
  }
}
const fail = (reason: RemapReason, detail?: string): never => {
  throw new Fail(reason, detail);
};

/** Functions whose value depends on where a reference / the formula is (beyond excelRef POSITIONAL_FUNCTIONS). */
const POSITIONAL: ReadonlySet<string> = new Set([...POSITIONAL_FUNCTIONS, 'INDIRECT', 'ADDRESS', 'FORMULATEXT', 'SHEETS', 'INFO', 'ISFORMULA', 'GETPIVOTDATA']);
/** the result is a POSITION inside the range (or the number of rows): changes when the layout of the table changes */
const POSITION_RESULT: ReadonlySet<string> = new Set(['MATCH', 'XMATCH', 'INDEX', 'ROWS', 'COUNTBLANK', 'AREAS', 'CHOOSEROWS', 'TAKE', 'DROP']);
/** the result depends on how many rows the range has */
const COUNT_RESULT: ReadonlySet<string> = new Set(['ROWS', 'COUNTBLANK']);
/** functions that take ranges: a whole Table (data body) is only meaningful as their argument */
const RANGE_FUNCTIONS: ReadonlySet<string> = new Set(`SUM AVERAGE AVERAGEA COUNT COUNTA COUNTBLANK COUNTIF COUNTIFS SUMIF SUMIFS AVERAGEIF AVERAGEIFS MAX MIN MAXA MINA MAXIFS MINIFS MEDIAN MODE STDEV STDEVA STDEVP STDEVPA VAR VARA VARP VARPA PRODUCT SUMPRODUCT SUMSQ LARGE SMALL RANK PERCENTILE QUARTILE PERCENTRANK GEOMEAN HARMEAN KURT SKEW DEVSQ AVEDEV TRIMMEAN CORREL SLOPE INTERCEPT RSQ FORECAST VLOOKUP HLOOKUP LOOKUP INDEX MATCH XLOOKUP XMATCH CHOOSE ROWS COLUMNS AREAS SUBTOTAL AGGREGATE TEXTJOIN CONCAT CONCATENATE UNIQUE FILTER SORT SORTBY TAKE DROP TOCOL TOROW HSTACK VSTACK CHOOSEROWS CHOOSECOLS`.split(/\s+/));

const eqi = (a: string | null | undefined, b: string | null | undefined): boolean => a != null && b != null && a.toLowerCase() === b.toLowerCase();

/** parsed structured reference (FILE form `Table1[[#This Row],[a]:[b]]`, `Table1[a]`, `Table1[#Data]`, UI form `[@a]`, `Table1[@a]` also accepted) */
export interface Sref {
  table: string;
  /** lower-case special items: 'this row', 'data', 'headers', 'all', 'totals' */
  specials: string[];
  /** first / last column of the column selection (names unescaped); null = every column */
  colFrom: string | null;
  colTo: string | null;
}

const unescapeItem = (s: string): string => s.replace(/'(.)/gs, '$1');

/** null = not a structured reference this module understands */
export function parseSref(text: string): Sref | null {
  const lb = text.indexOf('[');
  if (lb < 0 || !text.endsWith(']')) return null;
  const table = text.slice(0, lb);
  let body = text.slice(lb + 1, -1);
  const specials: string[] = [];
  const cols: string[] = [];
  let range = false;
  const addItem = (raw: string): boolean => {
    if (raw.startsWith('#')) {
      const k = raw.slice(1).trim().toLowerCase();
      if (!['this row', 'data', 'headers', 'all', 'totals'].includes(k)) return false;
      specials.push(k);
    } else {
      cols.push(unescapeItem(raw));
    }
    return true;
  };
  if (body.startsWith('@')) {
    specials.push('this row');
    body = body.slice(1);
    if (body === '') return { table, specials, colFrom: null, colTo: null };
  }
  if (body === '') return { table, specials, colFrom: null, colTo: null };
  if (body[0] !== '[') {
    // Table1[col], Table1[#Data], [@col]: ONE bare item (its characters [ ] # ' are escaped with ', e.g. Table1[点数'[1']])
    return addItem(body.trim()) ? { table, specials, colFrom: cols[0] ?? null, colTo: cols[0] ?? null } : null;
  }
  // list of bracketed items separated by ',' and ':'
  let k = 0;
  while (k < body.length) {
    const ch = body[k];
    if (ch === '[') {
      let e = k + 1;
      for (; e < body.length; e++) {
        if (body[e] === "'") { e++; continue; }
        if (body[e] === ']') break;
      }
      if (e >= body.length) return null;
      if (!addItem(body.slice(k + 1, e))) return null;
      k = e + 1;
    } else if (ch === ':') { range = true; k++; }
    else if (ch === ',' || ch === ' ') k++;
    else {
      // a bare item between the brackets (XlsxWriter: [[#This Row],点数]): up to the next , or : outside brackets
      let e = k;
      for (; e < body.length; e++) {
        if (body[e] === "'") { e++; continue; }
        if (body[e] === ',' || body[e] === ':') break;
      }
      if (!addItem(body.slice(k, e).trim())) return null;
      k = e;
    }
  }
  if (cols.length > 2 || (cols.length === 2 && !range) || (range && cols.length !== 2)) return null;
  return { table, specials, colFrom: cols[0] ?? null, colTo: cols.length === 2 ? cols[1] : cols[0] ?? null };
}

interface Cx {
  sr: number;
  /** source column of the formula cell (to find the table an unqualified structured reference belongs to) */
  sc: number | undefined;
  tr: number;
  exact: boolean;
  widened: boolean;
  substituted: boolean;
  /** index of the token being rewritten (for the enclosing function names) */
  tok: number;
}

export class FormulaRemapper {
  /** every source data row (that holds data) is placed in the merged sheet */
  readonly allPlaced: boolean;
  /** the merged table has the same number of rows as the source table */
  readonly sameCount: boolean;
  /** every source row sits at the same distance from the first data row in the merged sheet (positions inside the table unchanged) */
  readonly layoutIdentical: boolean;
  /** source order == merged order */
  readonly orderPreserved: boolean;
  private readonly t = new Map<number, number>();
  private readonly runOf = new Map<number, number>();
  private readonly names = new Map<string, RemapName[]>();
  private readonly tables = new Map<string, RemapTable>();
  private toks: Token[] = [];
  private calls: CallInfo[] | null = null;
  private outsideCache = new Map<string, boolean>();

  constructor(private readonly s: RemapSpec) {
    const dup = new Set<number>();
    for (const [sr, mr] of s.rowPairs) {
      if (dup.has(sr)) continue;
      if (this.t.has(sr)) { if (this.t.get(sr) !== mr) { this.t.delete(sr); dup.add(sr); } continue; }
      this.t.set(sr, mr);
    }
    // the header row maps to the base header row (a pseudo row: header cells are written there)
    const withHeader: Array<[number, number]> = [...this.t];
    if (!this.t.has(s.region.headerRow) && !dup.has(s.region.headerRow)) withHeader.push([s.region.headerRow, s.target.headerRow]);
    withHeader.sort((a, b) => a[0] - b[0]);
    let run = 0, prev: [number, number] | null = null;
    for (const p of withHeader) {
      if (prev && !(prev[0] + 1 === p[0] && prev[1] + 1 === p[1])) run++;
      this.runOf.set(p[0], run);
      prev = p;
    }
    this.headerMapped = withHeader.some((p) => p[0] === s.region.headerRow);
    const rows = s.dataRows ?? range(s.region.firstDataRow, s.region.lastDataRow);
    this.allPlaced = rows.every((r) => this.t.has(r));
    this.sameCount = s.target.lastDataRow - s.target.firstDataRow === s.region.lastDataRow - s.region.firstDataRow;
    const placed = [...this.t].sort((a, b) => a[0] - b[0]);
    this.orderPreserved = placed.every((p, i) => i === 0 || placed[i - 1][1] < p[1]);
    this.layoutIdentical = this.allPlaced && this.sameCount && placed.every((p) => p[1] - s.target.firstDataRow === p[0] - s.region.firstDataRow);
    for (const n of s.names ?? []) {
      const k = n.name.toLowerCase();
      const l = this.names.get(k);
      if (l) l.push(n); else this.names.set(k, [n]);
    }
    for (const tb of s.tables ?? []) this.tables.set(tb.name.toLowerCase(), tb);
  }

  private readonly headerMapped: boolean;

  /** Rewrite `formula` (FILE form, no '=') of source row `srcRow` for the merged row `targetRow`. */
  remap(formula: string, srcRow: number, targetRow: number, srcCol?: number): RemapResult {
    const cx: Cx = { sr: srcRow, sc: srcCol, tr: targetRow, exact: true, widened: false, substituted: false, tok: 0 };
    try {
      this.toks = tokenize(formula);
      this.calls = null;
      const toks = this.toks;
      let out = '';
      for (let i = 0; i < toks.length; i++) {
        const t = toks[i];
        cx.tok = i;
        switch (t.type) {
          case 'unknown':
            return fail('unsupported', t.text);
          case 'error':
            if (t.prefix || /^#REF!/i.test(t.text)) return fail('unsupported', t.text);
            out += t.text;
            break;
          case 'func':
            this.checkFunc(t);
            out += t.text;
            break;
          case 'name':
            out += this.name(t, cx);
            break;
          case 'sref':
            out += this.sref(t, cx);
            break;
          case 'ref':
            out += this.ref(t, cx);
            break;
          case 'op':
            if (t.text === '#') return fail('array', 'spill reference'); // A1# : the spill master is imported as a value
            out += t.text;
            break;
          default:
            out += t.text;
        }
      }
      // Excel refuses formulas of more than 8,192 characters (the _xlfn. / _xlws. / _xlpm. prefixes of the file form do not count)
      if (out.length > 8192 && out.replace(/_xl(?:fn|ws|pm)\./g, '').length > 8192) return fail('unsupported', 'formula too long');
      return { ok: true, formula: out, exact: cx.exact, widened: cx.widened, substituted: cx.substituted };
    } catch (e) {
      if (e instanceof Fail) return { ok: false, reason: e.reason, ...(e.detail ? { detail: e.detail } : {}) };
      throw e;
    }
  }

  // ------------------------------------------------------------------------------------------------ functions
  private checkFunc(t: Token): void {
    if (t.prefix) fail('unsupported', t.text); // [1]!MyFunc( / Sheet1!Fn( : a function of another workbook / sheet
    if (/^_xludf\./i.test(t.text)) fail('unsupported', t.text);
    const n = funcName(t);
    if (POSITIONAL.has(n)) fail('positional', n);
    if (n === 'ANCHORARRAY') fail('array', n); // a spill reference: the spill master is imported as a value
    if (this.names.size && this.names.has(n.toLowerCase()) && !isBuiltinFunction(n)) fail('unresolvedName', n); // a named LAMBDA
  }

  /** the calls that enclose token `tok`, outermost first */
  private enclosingCalls(tok: number): CallInfo[] {
    if (!this.calls) this.calls = [...parseCalls(this.toks, matchBrackets(this.toks)).values()].filter((c) => c.close >= 0).sort((a, b) => a.open - b.open);
    return this.calls.filter((c) => c.open < tok && tok < c.close);
  }

  private enclosing(tok: number): string[] {
    return this.enclosingCalls(tok).map((c) => c.name);
  }

  // ------------------------------------------------------------------------------------------------ references
  private ref(t: Token, cx: Cx): string {
    const p = t.prefix;
    if (p) {
      if (p.book !== undefined) return fail('unsafe', t.text);
      if (p.sheet2 !== undefined) return fail('otherSheet', t.text);
      if (!p.sheet || !eqi(p.sheet, this.s.sheet)) return fail('otherSheet', t.text);
    }
    return renderArea(this.mapArea(t.area!, cx));
  }

  private rowOf(r: number): number | undefined {
    return this.t.get(r);
  }

  private inData(r: number): boolean {
    return r >= this.s.region.firstDataRow && r <= this.s.region.lastDataRow;
  }

  private outside(col: number, r1: number, r2: number): boolean {
    if (r1 > r2 || !this.s.hasContentOutside) return false;
    const key = `${col}:${r1}:${r2}`;
    let v = this.outsideCache.get(key);
    if (v === undefined) { v = this.s.hasContentOutside(col, r1, r2); this.outsideCache.set(key, v); }
    return v;
  }

  private mapCols(lo: number, hi: number, cx: Cx): { lo: number; hi: number; partner: boolean } {
    const cm = this.s.colMap;
    if (lo === hi) {
      const m = cm.get(lo);
      if (m !== undefined) return { lo: m, hi: m, partner: false };
      const p = this.s.partnerCols?.get(lo);
      if (p !== undefined) { cx.substituted = true; return { lo: p, hi: p, partner: true }; }
      return fail('notImported', `col ${lo}`);
    }
    const m0 = cm.get(lo);
    if (m0 === undefined) return fail('notImported', `col ${lo}`);
    for (let c = lo + 1; c <= hi; c++) if (cm.get(c) !== m0 + (c - lo)) return fail('notImported', `cols ${lo}:${hi}`);
    return { lo: m0, hi: m0 + (hi - lo), partner: false };
  }

  /**
   * Is the cached result of the source still the result when a range over the whole table is rewritten to the merged table?
   * Aggregates (SUM / AVERAGE / COUNT / RANK / LARGE / SUMIF …) need the same set of rows (every source row placed); functions whose
   * result is a POSITION in the range (MATCH on its own, INDEX with a literal position, TAKE …) also need the same layout, functions
   * that count rows (ROWS, COUNTBLANK) the same number of rows. MATCH whose position is consumed by an INDEX is value-based.
   */
  private wholeExact(cx: Cx): boolean {
    if (!this.allPlaced) return false;
    const enc = this.enclosingCalls(cx.tok);
    for (let k = enc.length - 1; k >= 0; k--) {
      const c = enc[k];
      if (!this.sameCount && COUNT_RESULT.has(c.name)) return false;
      if (this.layoutIdentical || !POSITION_RESULT.has(c.name)) continue;
      if (c.name === 'MATCH' || c.name === 'XMATCH') {
        if (enc.slice(0, k).some((a) => a.name === 'INDEX')) continue;
        return false;
      }
      if (c.name === 'INDEX') {
        if (c.args.slice(1).some((a) => this.isLiteralNumber(a))) return false;
        continue;
      }
      return false;
    }
    return true;
  }

  private isLiteralNumber(arg: [number, number]): boolean {
    const [s, e] = trimRange(this.toks, arg);
    return e - s === 1 && this.toks[s].type === 'number';
  }

  /** Map a source area to the merged sheet. Throws Fail. */
  private mapArea(a: Area, cx: Cx): Area {
    const s = this.s;
    if (a.kind === 'rows') return fail('externalCell', 'whole row');
    const cLo = Math.min(a.c1, a.c2), cHi = Math.max(a.c1, a.c2);
    // a reference entirely above / below / beside the data table reads a cell that is not part of it (a rate in a side cell, a footer total)
    if (s.region.firstCol != null && s.region.lastCol != null && (cHi < s.region.firstCol || cLo > s.region.lastCol)) return fail('externalCell', 'beside the table');
    if (a.kind !== 'cols' && (Math.max(a.r1, a.r2) < s.region.headerRow || Math.min(a.r1, a.r2) > s.region.lastDataRow)) return fail('externalCell', 'above / below the table');
    const cm = this.mapCols(cLo, cHi, cx);
    const colsRev = a.c1 > a.c2;
    const c1 = colsRev ? cm.hi : cm.lo, c2 = colsRev ? cm.lo : cm.hi;
    if (a.kind === 'cols') {
      if (!cm.partner) {
        for (let c = cLo; c <= cHi; c++) {
          if (this.outside(c, 1, s.region.headerRow - 1) || this.outside(c, s.region.lastDataRow + 1, MAX_ROW)) return fail('externalCell', 'cells above / below the table in a whole-column reference');
        }
      }
      this.noteWhole(cx);
      return { ...a, c1, c2 };
    }
    const rLo = Math.min(a.r1, a.r2), rHi = Math.max(a.r1, a.r2);
    const rowsRev = a.r1 > a.r2;
    const mk = (top: number, bottom: number): Area => ({ ...a, c1, c2, r1: rowsRev ? bottom : top, r2: rowsRev ? top : bottom });
    if (rLo === rHi) {
      let mr: number;
      if (rLo === cx.sr) mr = cx.tr;
      else if (rLo === s.region.headerRow && this.headerMapped) {
        mr = s.target.headerRow;
        // the header cell of the merged sheet holds the (possibly renamed) header text of an imported column; a key column's partner has its own
        const same = cm.partner ? false : s.headerSame ? (cLo === cHi ? s.headerSame(cLo) : this.allHeadersSame(cLo, cHi)) : true;
        if (!same) cx.exact = false;
      } else {
        const m = this.rowOf(rLo);
        if (m === undefined) return fail(this.inData(rLo) ? 'rowRelation' : 'externalCell', `row ${rLo}`);
        mr = m;
      }
      return mk(mr, mr);
    }
    // ---- several rows
    const cols: number[] = [];
    for (let c = cLo; c <= cHi && cols.length < 64; c++) cols.push(c);
    const noOutside = (r1: number, r2: number) => cm.partner || !cols.some((c) => this.outside(c, r1, r2));
    const topOk = rLo <= s.region.firstDataRow && (rLo >= s.region.headerRow || noOutside(rLo, s.region.headerRow - 1));
    const botOk = rHi >= s.region.lastDataRow && (rHi === s.region.lastDataRow || noOutside(s.region.lastDataRow + 1, rHi));
    if (topOk && botOk) {
      this.noteWhole(cx);
      return mk(rLo <= s.region.headerRow && this.headerMapped ? s.target.headerRow : s.target.firstDataRow, s.target.lastDataRow);
    }
    const r1 = this.runOf.get(rLo), r2 = this.runOf.get(rHi);
    const m1 = rLo === cx.sr ? cx.tr : this.rowOf(rLo) ?? (rLo === s.region.headerRow ? s.target.headerRow : undefined);
    const m2 = rHi === cx.sr ? cx.tr : this.rowOf(rHi) ?? (rHi === s.region.headerRow ? s.target.headerRow : undefined);
    if (r1 === undefined || r1 !== r2 || m1 === undefined || m2 === undefined || m2 - m1 !== rHi - rLo) {
      return fail(this.inData(rLo) && this.inData(rHi) ? 'rowRelation' : 'externalCell', `rows ${rLo}:${rHi}`);
    }
    return mk(m1, m2);
  }

  private allHeadersSame(lo: number, hi: number): boolean {
    for (let c = lo; c <= hi; c++) if (!this.s.headerSame!(c)) return false;
    return true;
  }

  private noteWhole(cx: Cx): void {
    if (!this.wholeExact(cx)) cx.exact = false;
    if (!this.allPlaced || !this.sameCount) cx.widened = true;
  }

  // ------------------------------------------------------------------------------------------------ names
  private lookupName(text: string, sheet: string | null): RemapName | undefined {
    const l = this.names.get(text.toLowerCase());
    if (!l) return undefined;
    return l.find((n) => n.scope != null && eqi(n.scope, sheet ?? this.s.sheet)) ?? l.find((n) => n.scope == null);
  }

  private name(t: Token, cx: Cx): string {
    const p = t.prefix;
    const raw = p ? t.text.slice(p.text.length) : t.text;
    if (!p && /^_xl(pm|op|eta)\./i.test(raw)) {
      if (/^_xleta\./i.test(raw) && POSITIONAL.has(raw.slice(7).replace(/^(?:_xlfn\.|_xlws\.)+/i, '').toUpperCase())) return fail('positional', raw);
      return t.text; // LET / LAMBDA parameter, eta-reduced function
    }
    if (p) {
      if (p.book !== undefined) return fail('unsafe', t.text);
      if (p.sheet2 !== undefined || !p.sheet || !eqi(p.sheet, this.s.sheet)) return fail('otherSheet', t.text);
    }
    const def = this.lookupName(raw, p ? p.sheet : null);
    if (!def) return fail('unresolvedName', raw);
    const toks = tokenize(def.refersTo.replace(/^\s*=/, '').trim()).filter((x) => x.type !== 'ws');
    if (toks.length === 1) {
      const x = toks[0];
      if (x.type === 'ref') {
        const xp = x.prefix;
        if (!xp) return fail('unresolvedName', raw);
        if (xp.book !== undefined) return fail('unsafe', raw);
        if (xp.sheet2 !== undefined || !xp.sheet || !eqi(xp.sheet, this.s.sheet)) return fail('otherSheet', raw);
        const a = x.area!;
        const abs = a.kind === 'cols' ? a.ac1 && a.ac2 : a.kind === 'cell' ? a.ac1 && a.ar1 : a.kind === 'area' ? a.ac1 && a.ar1 && a.ac2 && a.ar2 : false;
        if (!abs) return fail('unresolvedName', `${raw}: relative name`);
        return renderArea(this.mapArea(a, cx));
      }
      if (x.type === 'sref') return this.srefText(x.text, cx, raw);
      if (x.type === 'number' || x.type === 'string' || x.type === 'bool') return x.text;
      return fail('unresolvedName', raw);
    }
    // a signed number / percent literal: -0.5, 8%
    const lit = toks.map((x) => x.text).join('');
    if (toks.length > 1 && /^[+-]?\d+(?:\.\d+)?(?:[eE][+-]?\d+)?%?$/.test(lit)) return '(' + lit + ')';
    return fail('unresolvedName', raw);
  }

  // ------------------------------------------------------------------------------------------------ structured references
  private sref(t: Token, cx: Cx): string {
    const p = t.prefix;
    if (p) {
      if (p.book !== undefined) return fail('unsafe', t.text);
      if (p.sheet2 !== undefined || !p.sheet || !eqi(p.sheet, this.s.sheet)) return fail('otherSheet', t.text);
    }
    return this.srefText(p ? t.text.slice(p.text.length) : t.text, cx, t.text);
  }

  private hostTable(cx: Cx): RemapTable | undefined {
    if (cx.sc === undefined) return undefined;
    for (const t of this.tables.values()) {
      if (!eqi(t.sheet, this.s.sheet)) continue;
      const a = parseAreaText(t.ref.replace(/^.*!/, '').replace(/\$/g, ''));
      if (a && cx.sr >= Math.min(a.r1, a.r2) && cx.sr <= Math.max(a.r1, a.r2) && cx.sc >= Math.min(a.c1, a.c2) && cx.sc <= Math.max(a.c1, a.c2)) return t;
    }
    return undefined;
  }

  private srefText(text: string, cx: Cx, what: string): string {
    const sr = parseSref(text);
    if (!sr) return fail('unresolvedTable', what);
    // Excel always names the table in the file, but some writers (XlsxWriter) leave it out inside a table: it is the table of the formula's cell
    const tb = sr.table ? this.tables.get(sr.table.toLowerCase()) : this.hostTable(cx);
    if (!tb) return fail('unresolvedTable', what);
    if (!eqi(tb.sheet, this.s.sheet)) return fail('otherSheet', what);
    const ra = parseAreaText(tb.ref.replace(/^.*!/, '').replace(/\$/g, ''));
    if (!ra) return fail('unresolvedTable', what);
    const top = Math.min(ra.r1, ra.r2), bottom = Math.max(ra.r1, ra.r2), left = Math.min(ra.c1, ra.c2), right = Math.max(ra.c1, ra.c2);
    const dataTop = top + (tb.headerRow ? 1 : 0), dataBottom = bottom - (tb.totalsRow ? 1 : 0);
    const sp = new Set(sr.specials);
    if (sp.has('totals') || (sp.has('all') && tb.totalsRow)) return fail('unresolvedTable', `${what} (totals row)`);
    // columns
    let c1 = left, c2 = right, named = false;
    if (sr.colFrom != null) {
      const f = tb.columns.findIndex((n) => eqi(n, sr.colFrom)), g = tb.columns.findIndex((n) => eqi(n, sr.colTo));
      if (f < 0 || g < 0) return fail('unresolvedTable', `${what} (column)`);
      c1 = left + Math.min(f, g); c2 = left + Math.max(f, g);
      named = true;
    }
    if (c2 > right) return fail('unresolvedTable', what);
    let r1: number, r2: number, rel = false;
    const only = (k: string) => sp.size === 1 && sp.has(k);
    if (only('this row')) { r1 = r2 = cx.sr; rel = true; }
    else if (only('headers')) { if (!tb.headerRow) return fail('unresolvedTable', what); r1 = r2 = top; }
    else if (sp.size === 0 || only('data')) {
      r1 = dataTop; r2 = dataBottom;
      // the whole table / a column range as an operand outside a function is not a range the formula could use (@ intersection of a table)
      if (!named && !this.enclosing(cx.tok).some((f) => RANGE_FUNCTIONS.has(f))) return fail('unresolvedTable', `${what} (whole table outside a function)`);
    } else if ((sp.size === 2 && sp.has('headers') && sp.has('data')) || only('all')) {
      if (!tb.headerRow) return fail('unresolvedTable', what);
      r1 = top; r2 = dataBottom;
    } else return fail('unresolvedTable', what);
    if (r1! < 1 || r2! < 1 || (rel && (cx.sr < dataTop || cx.sr > dataBottom))) return fail('unresolvedTable', what);
    const single = r1! === r2! && c1 === c2;
    const area: Area = { kind: single ? 'cell' : 'area', c1, c2, r1: r1!, r2: r2!, ac1: !rel, ar1: !rel, ac2: !rel, ar2: !rel };
    return renderArea(this.mapArea(area, cx));
  }
}

function range(a: number, b: number): number[] {
  const out: number[] = [];
  for (let r = a; r <= b; r++) out.push(r);
  return out;
}
