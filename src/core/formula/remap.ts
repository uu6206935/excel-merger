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
 *
 * Round 9 (includeSourceSheets, RemapSpec.sheetCopies): the ORIGINAL sheets of the source are in the output, so what rules a-e cannot
 * keep can read the copies instead. Three modes, tried in this order (a formula that rules a-e keep is NEVER touched: same result as
 * round 8, `viaCopy` unset):
 *  h. 'hybrid': every reference is first tried by rules a-e (merged table); one that fails reads the COPY of its sheet at the ORIGINAL
 *     coordinates (`'<copy>'!C5`: the source row of the matched person — the copy keeps the original rows — so a same-row reference keeps
 *     its row offset from the source row, ranges stay as in the source). A reference to another sheet of the source always reads its copy.
 *  i. 'copy': every reference of the formula reads the copies = exactly the source formula with renamed sheet references (always the value
 *     the source has for that person). Used when a hybrid formula would be wrong: a range of the merged table and a range of a copy in the
 *     same formula are not row-aligned any more (SUMIF / SUMPRODUCT / INDEX-MATCH over both), reference operators (`:` / space / union)
 *     would join different sheets, or the formula uses ROW(ref) / COLUMN(ref) / OFFSET(ref, …) (host-independent only on the copies).
 *  j. A multi-cell reference that now lives on ANOTHER sheet is only kept where Excel does not use implicit intersection (a bare argument
 *     of a function that takes a range, anything inside SUMPRODUCT / LOOKUP, or an array formula): otherwise the formula keeps its
 *     round-8 outcome (value). 3-D references, external workbooks, INDIRECT / ADDRESS / CELL / FORMULATEXT / SHEET(S) / INFO and
 *     ROW() / COLUMN() without an argument stay values.
 */
import { MAX_ROW, parseAreaText, quoteSheetName, renderArea, tokenize, type Area, type Token } from './excelRef';
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
  /**
   * includeSourceSheets: lower-cased name of every sheet of the source that was copied into the output (its data sheet included) -> the name
   * of the copy. References the merged table cannot represent read these copies instead (see the file comment, rules h-j). Absent / empty =
   * the sheets of the source are not in the output: behaviour of round 8.
   */
  sheetCopies?: ReadonlyMap<string, string>;
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
  /** (only when true) at least one reference reads a COPY of a source sheet (RemapSpec.sheetCopies) */
  viaCopy?: boolean;
  /** (only with viaCopy) the names of the copies the formula reads, in order of first use */
  copySheets?: string[];
}

export interface RemapFail {
  ok: false;
  reason: RemapReason;
  detail?: string;
}

export type RemapResult = RemapOk | RemapFail;

/** thrown and caught inside this module only. Not an Error on purpose: an Error captures a stack trace (tens of microseconds), and a source with
 *  100,000 formulas that cannot be kept throws (and tries the copies) for every one of them. */
class Fail {
  constructor(readonly reason: RemapReason, readonly detail?: string) {}
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

/**
 * Which arguments of a function take a RANGE (never implicitly intersected with the formula's row / column): a reference that now lives
 * on another sheet is only safe there (rule j). 0-based argument index, number of arguments.
 */
const RANGE_PARAM: ReadonlyMap<string, (i: number, n: number) => boolean> = (() => {
  const m = new Map<string, (i: number, n: number) => boolean>();
  const put = (names: string, f: (i: number, n: number) => boolean) => { for (const k of names.split(/\s+/)) m.set(k, f); };
  put(`SUM AVERAGE AVERAGEA COUNT COUNTA COUNTBLANK MAX MIN MAXA MINA MEDIAN MODE MODE.SNGL MODE.MULT STDEV STDEV.S STDEV.P STDEVA STDEVP STDEVPA
    VAR VAR.S VAR.P VARA VARP VARPA PRODUCT SUMPRODUCT SUMSQ GEOMEAN HARMEAN DEVSQ AVEDEV KURT SKEW SKEW.P CORREL COVAR COVARIANCE.P COVARIANCE.S
    PEARSON RSQ SLOPE INTERCEPT STEYX SUMX2MY2 SUMX2PY2 SUMXMY2 CONCAT AND OR XOR ROW COLUMN ROWS COLUMNS AREAS`, () => true);
  put(`LARGE SMALL PERCENTILE PERCENTILE.INC PERCENTILE.EXC QUARTILE QUARTILE.INC QUARTILE.EXC PERCENTRANK PERCENTRANK.INC PERCENTRANK.EXC TRIMMEAN
    COUNTIF INDEX OFFSET UNIQUE SORT CHOOSEROWS CHOOSECOLS TAKE DROP TOCOL TOROW EXPAND WRAPROWS WRAPCOLS`, (i) => i === 0);
  put('FILTER', (i) => i <= 1);
  put('VLOOKUP HLOOKUP MATCH XMATCH RANK RANK.EQ RANK.AVG', (i) => i === 1);
  put('XLOOKUP', (i) => i === 1 || i === 2);
  put('LOOKUP', (i) => i >= 1);
  put('SUMIF AVERAGEIF', (i) => i === 0 || i === 2);
  put('COUNTIFS', (i) => i % 2 === 0);
  put('SUMIFS AVERAGEIFS MAXIFS MINIFS', (i) => i === 0 || i % 2 === 1);
  put('SUBTOTAL', (i) => i >= 1);
  put('AGGREGATE', (i) => i >= 2);
  put('TEXTJOIN', (i) => i >= 2);
  put('FORECAST FORECAST.LINEAR', (i) => i >= 1);
  return m;
})();
/** functions that evaluate everything inside their arguments as arrays (no implicit intersection even in an expression like (C2:C9>5)*D2:D9) */
const ARRAY_EVAL: ReadonlySet<string> = new Set(['SUMPRODUCT', 'LOOKUP', 'MMULT', 'MINVERSE', 'MDETERM', 'SUMX2MY2', 'SUMX2PY2', 'SUMXMY2']);
/** ROW / COLUMN / OFFSET are host-independent when they are given a reference and every reference reads the copies (mode 'copy') */
const COPY_POSITIONAL: ReadonlySet<string> = new Set(['ROW', 'COLUMN', 'OFFSET']);
/** the failures of the merged-table mapping that a copy of the source sheet can replace */
const NO_COPY_REASONS: ReadonlySet<RemapReason> = new Set(['unsafe', 'unsupported', 'array']);
/** functions whose arguments after the first are POSITIONS inside the first (a position that comes from a copy is in the source's layout) */
const POSITION_FUNCS: ReadonlySet<string> = new Set(['INDEX', 'CHOOSEROWS', 'TAKE', 'DROP']);
/** marker detail of the internal failure "range of the merged table and range of a copy in one formula" */
const ALIGN = 'alignment';

type Mode = 'merged' | 'hybrid' | 'copy';

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
  mode: Mode;
  /** the formula is an array / dynamic-array formula (no implicit intersection) */
  array: boolean;
  /** at least one reference reads a copy */
  viaCopy: boolean;
  copyNames: Set<string>;
  /** a multi-row range of the merged table / of a copy is in the formula (the two together are not row-aligned: mode 'copy') */
  mergedRange: boolean;
  copyRange: boolean;
  /** calls (index of the function token) whose range argument is in the merged table / that take a position from a copy: together they are not aligned */
  posMerged: Set<number>;
  posCopy: Set<number>;
}

/** does the area span more than one row (a whole-column reference included)? */
function multiRow(a: Area): boolean {
  return a.kind === 'cols' || (a.kind === 'area' && a.r1 !== a.r2);
}
function singleCell(a: Area): boolean {
  return a.kind === 'cell' || (a.kind === 'area' && a.r1 === a.r2 && a.c1 === a.c2);
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
  /** lower-cased sheet name of the source -> copy name (null: the source's sheets are not in the output) */
  private readonly copies: ReadonlyMap<string, string> | null;

  constructor(private readonly s: RemapSpec) {
    this.copies = s.sheetCopies && s.sheetCopies.size > 0 ? s.sheetCopies : null;
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

  /**
   * Rewrite `formula` (FILE form, no '=') of source row `srcRow` for the merged row `targetRow`.
   * `array`: the source cell holds an array / dynamic-array formula (no implicit intersection; only matters for the copies).
   * Rules a-e of the file comment decide first; only when they cannot keep the formula and the source's sheets are copied (spec.sheetCopies)
   * are the copies tried (rules h-j). The failure reported is always the one of the merged-table rules.
   */
  remap(formula: string, srcRow: number, targetRow: number, srcCol?: number, array = false): RemapResult {
    const first = this.attempt(formula, srcRow, targetRow, srcCol, 'merged', array);
    if (first.ok || !this.copies || NO_COPY_REASONS.has(first.reason)) return first;
    let second = this.attempt(formula, srcRow, targetRow, srcCol, this.needsCopyMode(this.toks) ? 'copy' : 'hybrid', array);
    if (!second.ok && second.detail === ALIGN) second = this.attempt(formula, srcRow, targetRow, srcCol, 'copy', array);
    return second.ok ? second : first;
  }

  private attempt(formula: string, srcRow: number, targetRow: number, srcCol: number | undefined, mode: Mode, array: boolean): RemapResult {
    const cx: Cx = {
      sr: srcRow, sc: srcCol, tr: targetRow, exact: true, widened: false, substituted: false, tok: 0,
      mode, array, viaCopy: false, copyNames: new Set(), mergedRange: false, copyRange: false, posMerged: new Set(), posCopy: new Set(),
    };
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
            this.checkFunc(t, cx);
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
      // a range of the merged table next to a range of a copy: the rows of the two are not aligned any more
      if (mode === 'hybrid' && ((cx.mergedRange && cx.copyRange) || [...cx.posCopy].some((c) => cx.posMerged.has(c)))) return fail('rowRelation', ALIGN);
      // Excel refuses formulas of more than 8,192 characters (the _xlfn. / _xlws. / _xlpm. prefixes of the file form do not count)
      if (out.length > 8192 && out.replace(/_xl(?:fn|ws|pm)\./g, '').length > 8192) return fail('unsupported', 'formula too long');
      return {
        ok: true, formula: out, exact: cx.exact, widened: cx.widened, substituted: cx.substituted,
        ...(cx.viaCopy ? { viaCopy: true, copySheets: [...cx.copyNames] } : {}),
      };
    } catch (e) {
      if (e instanceof Fail) return { ok: false, reason: e.reason, ...(e.detail ? { detail: e.detail } : {}) };
      throw e;
    }
  }

  /**
   * Must EVERY reference of the formula read the copies (mode 'copy')? Yes when it uses a function whose result depends on where a
   * reference is (ROW(ref) / COLUMN(ref) / OFFSET / the other positional functions: the latter fail there anyway) or a reference operator
   * (`:` between non-reference operands, the intersection space, the union comma): the operands must be on ONE sheet.
   */
  private needsCopyMode(toks: Token[]): boolean {
    const stack: boolean[] = []; // true = the '(' opens a function call
    let prev = -1;
    for (let i = 0; i < toks.length; i++) {
      const t = toks[i];
      if (t.type === 'func' && POSITIONAL.has(funcName(t))) return true;
      if (t.type === 'name' && /^_xleta\./i.test(t.text)) return true;
      if (t.type === 'op') {
        if (t.text === '(') stack.push(prev >= 0 && toks[prev].type === 'func' && prev === i - 1);
        else if (t.text === ')') stack.pop();
        else if (t.text === ':') return true;
        else if (t.text === ',' && stack.length && stack[stack.length - 1] === false) return true; // union (a, b)
      }
      if (t.type === 'ws' && prev >= 0 && i + 1 < toks.length) {
        const a = toks[prev], b = toks[i + 1];
        const left = a.type === 'ref' || a.type === 'name' || a.type === 'sref' || (a.type === 'op' && a.text === ')');
        const right = b.type === 'ref' || b.type === 'name' || b.type === 'sref' || (b.type === 'op' && b.text === '(');
        if (left && right) return true; // intersection
      }
      if (t.type !== 'ws') prev = i;
    }
    return false;
  }

  // ------------------------------------------------------------------------------------------------ functions
  private checkFunc(t: Token, cx: Cx): void {
    if (t.prefix) fail('unsupported', t.text); // [1]!MyFunc( / Sheet1!Fn( : a function of another workbook / sheet
    if (/^_xludf\./i.test(t.text)) fail('unsupported', t.text);
    const n = funcName(t);
    if (POSITIONAL.has(n) && !(cx.mode === 'copy' && this.positionalOnCopies(n, cx))) fail('positional', n);
    if (n === 'ANCHORARRAY') fail('array', n); // a spill reference: the spill master is imported as a value
    if (this.names.size && this.names.has(n.toLowerCase()) && !isBuiltinFunction(n)) fail('unresolvedName', n); // a named LAMBDA
  }

  /** ROW(ref) / COLUMN(ref) / OFFSET(ref, …): the same on the copies (they have the source's rows and columns) — with an argument only */
  private positionalOnCopies(n: string, cx: Cx): boolean {
    if (!COPY_POSITIONAL.has(n)) return false;
    const c = this.allCalls().find((x) => x.func === cx.tok);
    if (!c || !c.args.length) return false;
    const [s, e] = trimRange(this.toks, c.args[0]);
    return e > s;
  }

  private allCalls(): CallInfo[] {
    if (!this.calls) this.calls = [...parseCalls(this.toks, matchBrackets(this.toks)).values()].filter((c) => c.close >= 0).sort((a, b) => a.open - b.open);
    return this.calls;
  }

  /** the calls that enclose token `tok`, outermost first */
  private enclosingCalls(tok: number): CallInfo[] {
    return this.allCalls().filter((c) => c.open < tok && tok < c.close);
  }

  private enclosing(tok: number): string[] {
    return this.enclosingCalls(tok).map((c) => c.name);
  }

  // ------------------------------------------------------------------------------------------------ references
  private ref(t: Token, cx: Cx): string {
    const p = t.prefix;
    let sheet = this.s.sheet;
    if (p) {
      if (p.book !== undefined) return fail('unsafe', t.text);
      if (p.sheet2 !== undefined) return fail('otherSheet', t.text); // 3-D reference: stays a value
      if (!p.sheet) return fail('otherSheet', t.text);
      sheet = p.sheet;
      if (!eqi(sheet, this.s.sheet) && cx.mode === 'merged') return fail('otherSheet', t.text);
    }
    return this.route(t.area!, sheet, p ? t.text.slice(p.text.length) : t.text, cx, t.text);
  }

  /**
   * One reference (an area of `sheet` of the source; `body` = its text without sheet prefix) of the merged sheet: the merged table first (the
   * data sheet only), then — mode 'hybrid' / 'copy' — the copy of the sheet at the original coordinates. `what`: the text for the failure.
   */
  private route(a: Area, sheet: string, body: string, cx: Cx, what: string): string {
    const onData = eqi(sheet, this.s.sheet);
    if (cx.mode === 'merged') return onData ? renderArea(this.mapArea(a, cx)) : fail('otherSheet', what);
    let orig: Fail | undefined;
    if (onData && cx.mode === 'hybrid') {
      const r = this.tryMap(a, cx);
      if ('area' in r) {
        if (multiRow(r.area)) this.notePositions(cx, 'merged');
        return renderArea(r.area);
      }
      orig = r.fail;
    }
    return this.copyRef(a, sheet, body, cx, orig ?? new Fail('otherSheet', what));
  }

  /** mapArea that leaves `cx` as it was when the area cannot be mapped */
  private tryMap(a: Area, cx: Cx): { area: Area } | { fail: Fail } {
    const save = { exact: cx.exact, widened: cx.widened, substituted: cx.substituted, mergedRange: cx.mergedRange };
    try {
      return { area: this.mapArea(a, cx) };
    } catch (e) {
      if (!(e instanceof Fail)) throw e;
      Object.assign(cx, save);
      return { fail: e };
    }
  }

  /** the reference on the copy of `sheet`; throws `orig` when there is no copy or Excel would intersect the range implicitly there */
  private copyRef(a: Area, sheet: string, body: string, cx: Cx, orig: Fail): string {
    const copy = this.copies?.get(sheet.toLowerCase());
    if (copy === undefined || !this.rangeSafe(a, cx)) throw orig;
    cx.viaCopy = true;
    cx.copyNames.add(copy);
    if (multiRow(a)) cx.copyRange = true;
    this.notePositions(cx, 'copy');
    return quoteSheetName(copy) + '!' + body;
  }

  /**
   * INDEX(range, position): a range of the merged table and a position that comes from a copy (a cell of the source) are not aligned — the
   * position counts rows of the SOURCE layout. 'merged': the token is in the first argument of such a call (a merged multi-row range);
   * 'copy': the token is in a later argument (any reference that reads a copy).
   */
  private notePositions(cx: Cx, kind: 'merged' | 'copy'): void {
    for (const c of this.enclosingCalls(cx.tok)) {
      if (!POSITION_FUNCS.has(c.name)) continue;
      const k = c.args.findIndex(([s, e]) => cx.tok >= s && cx.tok < e);
      if (k === 0 && kind === 'merged') cx.posMerged.add(c.func);
      else if (k > 0 && kind === 'copy') cx.posCopy.add(c.func);
    }
  }

  /**
   * May the multi-cell reference at the current token live on ANOTHER sheet than the formula (rule j)? A single cell always may. A range
   * is intersected implicitly with the formula's row / column unless it is a bare argument of a function that takes a range there, sits
   * anywhere inside a function that evaluates its arguments as arrays (SUMPRODUCT …), or the formula is an array formula.
   */
  private rangeSafe(a: Area, cx: Cx): boolean {
    if (singleCell(a) || cx.array) return true;
    const enc = this.enclosingCalls(cx.tok);
    if (!enc.length) return false;
    if (enc.some((c) => ARRAY_EVAL.has(c.name))) return true;
    const c = enc[enc.length - 1];
    const k = c.args.findIndex(([s, e]) => cx.tok >= s && cx.tok < e);
    if (k < 0) return false;
    const [s, e] = trimRange(this.toks, c.args[k]);
    // the argument is the reference itself, or a reference expression (union / intersection / ranges built with ':' of references only)
    if (!(s === cx.tok && e === cx.tok + 1) && !this.onlyReferences(s, e)) return false;
    const takes = RANGE_PARAM.get(c.name);
    return !!takes && takes(k, c.args.length);
  }

  /** tokens [s, e) are references joined by the reference operators (space, ',' inside parentheses, ':') and nothing else */
  private onlyReferences(s: number, e: number): boolean {
    for (let i = s; i < e; i++) {
      const t = this.toks[i];
      if (t.type === 'ref' || t.type === 'name' || t.type === 'sref' || t.type === 'ws') continue;
      if (t.type === 'op' && (t.text === '(' || t.text === ')' || t.text === ',' || t.text === ':')) continue;
      return false;
    }
    return true;
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
    cx.mergedRange = true;
    return mk(m1, m2);
  }

  private allHeadersSame(lo: number, hi: number): boolean {
    for (let c = lo; c <= hi; c++) if (!this.s.headerSame!(c)) return false;
    return true;
  }

  private noteWhole(cx: Cx): void {
    cx.mergedRange = true;
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
    const all = tokenize(def.refersTo.replace(/^\s*=/, '').trim());
    const toks = all.filter((x) => x.type !== 'ws');
    if (toks.length === 1) {
      const x = toks[0];
      if (x.type === 'ref') {
        const xp = x.prefix;
        if (!xp) return fail('unresolvedName', raw);
        if (xp.book !== undefined) return fail('unsafe', raw);
        if (xp.sheet2 !== undefined || !xp.sheet) return fail('otherSheet', raw);
        if (!eqi(xp.sheet, this.s.sheet) && cx.mode === 'merged') return fail('otherSheet', raw);
        const a = x.area!;
        if (!absolute(a)) return fail('unresolvedName', `${raw}: relative name`);
        return this.route(a, xp.sheet, renderArea(a), cx, raw);
      }
      if (x.type === 'sref') return this.srefText(x.text, cx, raw);
      if (x.type === 'number' || x.type === 'string' || x.type === 'bool') return x.text;
      return fail('unresolvedName', raw);
    }
    // a signed number / percent literal: -0.5, 8%
    const lit = toks.map((x) => x.text).join('');
    if (toks.length > 1 && /^[+-]?\d+(?:\.\d+)?(?:[eE][+-]?\d+)?%?$/.test(lit)) return '(' + lit + ')';
    if (cx.mode !== 'merged' && toks.length > 1) return this.inlineExpression(all, cx, raw);
    return fail('unresolvedName', raw);
  }

  /**
   * A defined name whose formula is an expression over absolute single-cell references (`Rates!$B$1*1.1`, `IF(Rates!$B$2>0,Rates!$B$2,5)`):
   * inlined in parentheses with every reference reading the COPY of its sheet (only with copies; the formula of a name is evaluated on the
   * sheets it names, so this is the very same value as in the source). Anything else — relative references, ranges, other names, tables,
   * functions that depend on a position — cannot be inlined.
   */
  private inlineExpression(toks: Token[], cx: Cx, raw: string): string {
    let out = '';
    for (const x of toks) {
      switch (x.type) {
        case 'ws': case 'number': case 'string': case 'bool':
          out += x.text;
          break;
        case 'op':
          if (x.text === '#') return fail('unresolvedName', raw);
          out += x.text;
          break;
        case 'error':
          if (x.prefix || /^#REF!/i.test(x.text)) return fail('unresolvedName', raw);
          out += x.text;
          break;
        case 'func': {
          const n = funcName(x);
          if (x.prefix || /^_xludf\./i.test(x.text) || POSITIONAL.has(n) || n === 'ANCHORARRAY') return fail('unresolvedName', raw);
          out += x.text;
          break;
        }
        case 'ref': {
          const xp = x.prefix;
          if (!xp || !xp.sheet || xp.book !== undefined || xp.sheet2 !== undefined) return fail('unresolvedName', raw);
          const a = x.area!;
          if (!singleCell(a) || !absolute(a)) return fail('unresolvedName', raw);
          out += this.copyRef(a, xp.sheet, x.text.slice(xp.text.length), cx, new Fail('unresolvedName', raw));
          break;
        }
        default:
          return fail('unresolvedName', raw);
      }
    }
    return '(' + out.trim() + ')';
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
    const onData = eqi(tb.sheet, this.s.sheet);
    if (!onData && cx.mode === 'merged') return fail('otherSheet', what);
    const ra = parseAreaText(tb.ref.replace(/^.*!/, '').replace(/\$/g, ''));
    if (!ra) return fail('unresolvedTable', what);
    const top = Math.min(ra.r1, ra.r2), bottom = Math.max(ra.r1, ra.r2), left = Math.min(ra.c1, ra.c2), right = Math.max(ra.c1, ra.c2);
    const dataTop = top + (tb.headerRow ? 1 : 0), dataBottom = bottom - (tb.totalsRow ? 1 : 0);
    const sp = new Set(sr.specials);
    // the totals row is no merged data: only a copy can be read there
    const totals = sp.has('totals') || (sp.has('all') && tb.totalsRow);
    if (totals && (cx.mode === 'merged' || !tb.totalsRow)) return fail('unresolvedTable', `${what} (totals row)`);
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
    if (only('this row')) {
      // [@col] means "this row of the table the formula is in": a table of another sheet has no such row
      if (!onData) return fail('unresolvedTable', what);
      r1 = r2 = cx.sr; rel = true;
    }
    else if (only('headers')) { if (!tb.headerRow) return fail('unresolvedTable', what); r1 = r2 = top; }
    else if (only('totals')) { r1 = r2 = bottom; }
    else if (sp.size === 0 || only('data')) {
      r1 = dataTop; r2 = dataBottom;
      // the whole table / a column range as an operand outside a function is not a range the formula could use (@ intersection of a table)
      if (!named && !this.enclosing(cx.tok).some((f) => RANGE_FUNCTIONS.has(f))) return fail('unresolvedTable', `${what} (whole table outside a function)`);
    } else if (sp.size === 2 && sp.has('data') && sp.has('totals')) { r1 = dataTop; r2 = bottom; }
    else if ((sp.size === 2 && sp.has('headers') && sp.has('data')) || only('all')) {
      if (!tb.headerRow) return fail('unresolvedTable', what);
      r1 = top; r2 = only('all') && tb.totalsRow ? bottom : dataBottom;
    } else return fail('unresolvedTable', what);
    if (r1! < 1 || r2! < 1 || (rel && (cx.sr < dataTop || cx.sr > dataBottom))) return fail('unresolvedTable', what);
    const single = r1! === r2! && c1 === c2;
    const area: Area = { kind: single ? 'cell' : 'area', c1, c2, r1: r1!, r2: r2!, ac1: !rel, ar1: !rel, ac2: !rel, ar2: !rel };
    return this.route(area, tb.sheet, renderArea(area), cx, what);
  }
}

/** every part of the reference is absolute ($C$2, $C$2:$D$9, $C:$D) */
function absolute(a: Area): boolean {
  return a.kind === 'cols' ? a.ac1 && a.ac2 : a.kind === 'cell' ? a.ac1 && a.ar1 : a.kind === 'area' ? a.ac1 && a.ar1 && a.ac2 && a.ar2 : false;
}

function range(a: number, b: number): number[] {
  const out: number[] = [];
  for (let r = a; r <= b; r++) out.push(r);
  return out;
}
