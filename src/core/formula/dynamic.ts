/**
 * Dynamic-array detection (owner: formula engine F).
 *
 * A formula typed in Excel 365 (Range.Formula2 / the formula bar) is evaluated with dynamic-array semantics. Excel
 * stores it as a dynamic-array formula (<f t="array" ref="A1" ...> on a cell with cm="1") when its legacy
 * (pre-2019) evaluation could differ: it returns an array, or a multi-cell value flows into an operator or into a
 * parameter that expects a single value (where legacy Excel would silently intersect). A plain <f> without cm is
 * opened by Excel with legacy semantics and shown with '@' (=@A1:A3, =@SEQUENCE(3)) — a different result.
 *
 * isDynamicArrayFormula(fileFormula) predicts that flag for a FILE-form formula (toFileForm of what the user typed).
 * It is CONSERVATIVE: whenever it cannot prove that legacy evaluation is identical it answers true (unknown
 * functions, defined names without a resolver, OFFSET / INDIRECT, functions not in the range-parameter table ...).
 * A formula flagged dynamic that returns a single value behaves exactly like a normal one, so a false "true" is
 * harmless; a false "false" would change the result. Checked against every formula typed into real Excel in the
 * fixtures (tests/formula/dynamic.test.ts): no false "false".
 *
 * Use (JS engine): write a typed formula for which this returns true as
 *   <c r="B2" cm="1"><f t="array" ref="B2">...</f></c>   (plus the dynamic-array cell metadata part)
 * and the Excel engine writes it with Range.Formula2 — both then mean what the user typed.
 */
import { tokenize, parseStructuredRef, type Token } from './excelRef';
import { matchBrackets, parseCalls, trimRange, type CallInfo } from './calls';
import { isBuiltinFunction } from './forms';

type Shape = 'scalar' | 'multi';

/**
 * Functions Excel always marks dynamic (verified: IFERROR(F(1),0) typed through Formula2 gets cm="1" for exactly
 * these; tests/fixtures/formula/forms-excel-results.json, forms-new-excel-results.json) — they (can) return arrays.
 */
export const ARRAY_FUNCTIONS: ReadonlySet<string> = new Set(`BYCOL BYROW CHOOSECOLS CHOOSEROWS DROP EXPAND FIELDVALUE
FILTER FILTERXML FREQUENCY GROUPBY HSTACK IFS IMPORTCSV IMPORTTEXT INDIRECT JIS MAKEARRAY MAP MODE.MULT MUNIT
PIVOTBY RANDARRAY REDUCE REGEXEXTRACT SCAN SEQUENCE SORT SORTBY STOCKHISTORY SWITCH TAKE TEXTSPLIT TOCOL TOROW
TRANSPOSE TRIMRANGE UNIQUE VSTACK WRAPCOLS WRAPROWS COPILOT QUERYSTRING LONGTEXT PYTHON_STR PYTHON_TYPE PYTHON_TYPENAME
PY`.split(/\s+/).filter(Boolean));

const ALL = (): boolean => true;
const K0 = (k: number): boolean => k === 0;
const K1 = (k: number): boolean => k === 1;
const K01 = (k: number): boolean => k <= 1;
const K02 = (k: number): boolean => k === 0 || k === 2;
const IFS_RANGES = (k: number): boolean => k === 0 || k % 2 === 1; // SUMIFS(sum_range, crit_range1, crit1, ...)
/**
 * Parameters that take a range / array as a whole (no implicit intersection in legacy Excel) of functions that
 * return ONE value. A multi-cell argument anywhere else makes the formula dynamic.
 */
const RANGE_PARAMS: Readonly<Record<string, (k: number) => boolean>> = {
  SUM: ALL, SUMSQ: ALL, PRODUCT: ALL, AVERAGE: ALL, AVERAGEA: ALL, COUNT: ALL, COUNTA: ALL, MAX: ALL, MAXA: ALL, MIN: ALL,
  MINA: ALL, MEDIAN: ALL, MODE: ALL, 'MODE.SNGL': ALL, STDEV: ALL, 'STDEV.S': ALL, 'STDEV.P': ALL, STDEVA: ALL, STDEVP: ALL,
  STDEVPA: ALL, VAR: ALL, 'VAR.S': ALL, 'VAR.P': ALL, VARA: ALL, VARP: ALL, VARPA: ALL, AVEDEV: ALL, DEVSQ: ALL, GEOMEAN: ALL,
  HARMEAN: ALL, KURT: ALL, SKEW: ALL, 'SKEW.P': ALL, AND: ALL, OR: ALL, XOR: ALL, CONCAT: ALL, SUMPRODUCT: ALL, PERCENTOF: ALL,
  COUNTBLANK: K0, ROWS: K0, COLUMNS: K0, AREAS: K0,
  SUMIF: K02, COUNTIF: K0, AVERAGEIF: K02, SUMIFS: IFS_RANGES, AVERAGEIFS: IFS_RANGES, MAXIFS: IFS_RANGES, MINIFS: IFS_RANGES,
  COUNTIFS: (k) => k % 2 === 0,
  VLOOKUP: K1, HLOOKUP: K1, MATCH: K1, XMATCH: K1, LOOKUP: (k) => k >= 1, INDEX: K0, XLOOKUP: (k) => k === 1 || k === 2,
  LARGE: K0, SMALL: K0, RANK: K1, 'RANK.EQ': K1, 'RANK.AVG': K1, PERCENTILE: K0, 'PERCENTILE.INC': K0, 'PERCENTILE.EXC': K0,
  QUARTILE: K0, 'QUARTILE.INC': K0, 'QUARTILE.EXC': K0, PERCENTRANK: K0, 'PERCENTRANK.INC': K0, 'PERCENTRANK.EXC': K0,
  TRIMMEAN: K0, SUBTOTAL: (k) => k >= 1, AGGREGATE: (k) => k >= 2, TEXTJOIN: (k) => k >= 2,
  NPV: (k) => k >= 1, IRR: K0, MIRR: K0, XNPV: (k) => k === 1 || k === 2, XIRR: K01,
  CORREL: K01, COVAR: K01, 'COVARIANCE.P': K01, 'COVARIANCE.S': K01, PEARSON: K01, RSQ: K01, SLOPE: K01, INTERCEPT: K01,
  STEYX: K01, SUMX2MY2: K01, SUMX2PY2: K01, SUMXMY2: K01, FORECAST: (k) => k === 1 || k === 2, 'FORECAST.LINEAR': (k) => k === 1 || k === 2,
  DAVERAGE: K02, DCOUNT: K02, DCOUNTA: K02, DGET: K02, DMAX: K02, DMIN: K02, DPRODUCT: K02, DSTDEV: K02, DSTDEVP: K02,
  DSUM: K02, DVAR: K02, DVARP: K02, N: K0, MDETERM: K0, OFFSET: K0,
  // explicit intersection / spill / trim wrappers consume their argument as a whole
  SINGLE: ALL, ANCHORARRAY: ALL, _TRO_ALL: ALL, _TRO_LEADING: ALL, _TRO_TRAILING: ALL,
};

/** Options of isDynamicArrayFormula. */
export interface DynamicOptions {
  /**
   * refersTo (FILE form, with or without '=') of a defined name, or null when unknown. A name whose refersTo is a
   * single cell or a constant is a single value; without a resolver every defined name counts as a range.
   */
  resolveName?: (name: string) => string | null;
}

/** Would Excel store this FILE-form formula, typed in Excel 365, as a dynamic-array formula? (conservative) */
export function isDynamicArrayFormula(fileFormula: string, opts: DynamicOptions = {}): boolean {
  try {
    return new Analyzer(fileFormula.replace(/^=/, ''), opts, 0).run();
  } catch {
    return true; // pathological input (e.g. thousands of nested parentheses): stay on the safe side
  }
}

class Analyzer {
  private readonly toks: Token[];
  private readonly br: Int32Array;
  private readonly calls: Map<number, CallInfo>;
  private dyn = false;

  constructor(formula: string, private readonly opts: DynamicOptions, private readonly depth: number) {
    this.toks = tokenize(formula);
    this.br = matchBrackets(this.toks);
    this.calls = parseCalls(this.toks, this.br);
  }

  run(): boolean {
    if (this.toks.some((t) => t.type === 'unknown')) return true;
    if (this.expr(0, this.toks.length, new Map()) === 'multi') this.dyn = true;
    return this.dyn;
  }

  /** shape of the expression in tokens [s, e); sets `dyn` when a multi value meets an operator */
  private expr(s: number, e: number, scope: ReadonlyMap<string, Shape>): Shape {
    [s, e] = trimRange(this.toks, [s, e]);
    const shapes: Shape[] = [];
    let valueOp = false, rangeOp = false;
    for (let i = s; i < e;) {
      const t = this.toks[i];
      if (t.type === 'ws') { i++; continue; }
      if (t.type === 'op') {
        switch (t.text) {
          case '+': case '-': case '*': case '/': case '^': case '&': case '=': case '<>': case '<': case '>':
          case '<=': case '>=': case '%':
            valueOp = true; i++; continue;
          case ':': rangeOp = true; i++; continue;
          case '@': { // explicit intersection: one operand, single value
            const end = this.operandEnd(i + 1, e);
            if (end > i + 1) this.expr(i + 1, end, scope);
            shapes.push('scalar'); i = Math.max(end, i + 1); continue;
          }
          case '#': shapes[shapes.length - 1] = 'multi'; i++; continue; // X# spill reference
          case '(': {
            const close = this.br[i] > i ? Math.min(this.br[i], e) : e;
            shapes.push(this.group(i + 1, close, scope));
            i = close + 1; continue;
          }
          case '{': {
            const close = this.br[i] > i ? Math.min(this.br[i], e) : e;
            const inner = this.toks.slice(i + 1, close);
            shapes.push(inner.some((x) => x.type === 'op' && (x.text === ',' || x.text === ';')) ? 'multi' : 'scalar');
            i = close + 1; continue;
          }
          default: i++; continue;
        }
      }
      const r = this.operand(i, e, scope);
      shapes.push(r.shape);
      i = r.end;
    }
    if (valueOp) {
      if (shapes.includes('multi')) { this.dyn = true; return 'multi'; }
      return 'scalar';
    }
    if (rangeOp || shapes.length > 1) return 'multi'; // A1:INDEX(..), intersection "A1:B3 B2:C4"
    return shapes[0] ?? 'scalar';
  }

  /** a parenthesized group: an expression, or a union "(A1:B2,C3)" */
  private group(s: number, e: number, scope: ReadonlyMap<string, Shape>): Shape {
    const parts: Array<[number, number]> = [];
    let p = s;
    for (let k = s; k < e; k++) {
      const t = this.toks[k];
      if (t.type === 'op' && (t.text === '(' || t.text === '{') && this.br[k] > k) { k = this.br[k]; continue; }
      if (t.type === 'op' && t.text === ',') { parts.push([p, k]); p = k + 1; }
    }
    parts.push([p, e]);
    const shapes = parts.map(([a, b]) => this.expr(a, b, scope));
    return parts.length > 1 ? 'multi' : shapes[0];
  }

  /** end (exclusive) of the primary operand starting at i */
  private operandEnd(i: number, e: number): number {
    while (i < e && this.toks[i].type === 'ws') i++;
    if (i >= e) return i;
    const t = this.toks[i];
    let end = i + 1;
    if (t.type === 'func') { const c = this.calls.get(i); end = c && c.close >= 0 ? c.close + 1 : e; }
    else if (t.type === 'op' && (t.text === '(' || t.text === '{')) end = this.br[i] > i ? this.br[i] + 1 : e;
    if (end < e && this.toks[end].type === 'op' && this.toks[end].text === '#') end++;
    return Math.min(end, e);
  }

  private operand(i: number, e: number, scope: ReadonlyMap<string, Shape>): { end: number; shape: Shape } {
    const t = this.toks[i];
    switch (t.type) {
      case 'ref': return { end: i + 1, shape: t.area && t.area.kind === 'cell' && !t.area.trim ? 'scalar' : 'multi' };
      case 'sref': return { end: i + 1, shape: srefShape(t) };
      case 'name': return { end: i + 1, shape: this.nameShape(t, scope) };
      case 'func': return this.call(i, e, scope);
      default: return { end: i + 1, shape: 'scalar' }; // number, string, bool, error
    }
  }

  private nameShape(t: Token, scope: ReadonlyMap<string, Shape>): Shape {
    if (!t.prefix) {
      if (/^_xlpm\./i.test(t.text)) return scope.get(t.text.slice(6).toLowerCase()) ?? 'multi';
      if (/^_xleta\./i.test(t.text)) return 'scalar'; // a function used as a lambda value (only in LAMBDA positions)
      if (/^_xlop\./i.test(t.text)) return 'multi';
    }
    if (t.prefix?.book !== undefined) return 'multi';
    const name = t.text.slice(t.prefix?.text.length ?? 0);
    const refersTo = this.opts.resolveName?.(name);
    if (refersTo == null || this.depth > 8) return 'multi';
    const inner = new Analyzer(refersTo.replace(/^=/, ''), this.opts, this.depth + 1);
    if (inner.toks.some((x) => x.type === 'unknown')) return 'multi';
    const shape = inner.expr(0, inner.toks.length, new Map());
    if (inner.dyn) this.dyn = true;
    return shape;
  }

  private call(i: number, e: number, scope: ReadonlyMap<string, Shape>): { end: number; shape: Shape } {
    const t = this.toks[i];
    const c = this.calls.get(i);
    if (!c || c.close < 0 || c.close >= e) { this.dyn = true; return { end: e, shape: 'multi' }; }
    let end = c.close + 1;
    // LAMBDA(...)(args) / _xlpm.f(args)(...): calling a lambda
    if (end < e && this.toks[end].type === 'op' && this.toks[end].text === '(') {
      this.dyn = true;
      const close = this.br[end] > end ? this.br[end] : e - 1;
      this.group(end + 1, Math.min(close, e), scope);
      end = Math.min(close + 1, e);
    }
    const text = t.text.slice(t.prefix?.text.length ?? 0);
    const name = c.name;
    const args = c.args;
    const evalArgs = (isRange: (k: number) => boolean) => args.map((a, k) => {
      const sh = this.expr(a[0], a[1], scope);
      if (sh === 'multi' && !isRange(k)) this.dyn = true;
      return sh;
    });
    // lambda parameter called as a function, user-defined / unknown / other-book functions
    if (t.prefix || /^_xl(pm|udf)\./i.test(text) || !isBuiltinFunction(name)) {
      evalArgs(ALL);
      this.dyn = true;
      return { end, shape: 'multi' };
    }
    if (name === 'LET') return { end, shape: this.let(c, scope) };
    // a LAMBDA value (not called here): it only matters where it is called (-> dynamic, above) or passed to a
    // lambda helper (MAP, BYROW, ... -> ARRAY_FUNCTIONS); an uncalled LAMBDA is not flagged (verified)
    if (name === 'LAMBDA') return { end, shape: 'scalar' };
    if (ARRAY_FUNCTIONS.has(name)) {
      evalArgs(ALL);
      this.dyn = true;
      return { end, shape: 'multi' };
    }
    const sh = evalArgs(RANGE_PARAMS[name] ?? (() => false));
    switch (name) {
      case 'SINGLE': return { end, shape: 'scalar' };
      case 'ANCHORARRAY': return { end, shape: 'multi' };
      case '_TRO_ALL': case '_TRO_LEADING': case '_TRO_TRAILING': return { end, shape: sh[0] ?? 'multi' };
      case 'INDEX': return { end, shape: sh[0] === 'multi' ? this.indexShape(c) : 'scalar' };
      case 'XLOOKUP': return { end, shape: sh[2] === 'multi' && !this.oneDimensional(args[2]) ? 'multi' : 'scalar' };
      case 'OFFSET': { // OFFSET(base, rows, cols, [height], [width]): one cell for a one-cell base without a size
        const one = (k: number) => { const a = args[k]; if (!a) return true; const [s, e] = trimRange(this.toks, a); return e === s || (e - s === 1 && this.toks[s].text === '1'); };
        return { end, shape: sh[0] === 'scalar' && one(3) && one(4) ? 'scalar' : 'multi' };
      }
      default: return { end, shape: 'scalar' };
    }
  }

  /** LET(name1, value1, ..., calc): names take the shape of their value */
  private let(c: CallInfo, scope: ReadonlyMap<string, Shape>): Shape {
    const inner = new Map(scope);
    const n = c.args.length;
    let last: Shape = 'scalar';
    for (let k = 0; k < n; k++) {
      const [s, e] = trimRange(this.toks, c.args[k]);
      if (k < n - 1 && k % 2 === 0) continue; // declaration
      const sh = this.expr(s, e, inner);
      if (k < n - 1) {
        const [ds, de] = trimRange(this.toks, c.args[k - 1]);
        if (de - ds === 1) inner.set(this.toks[ds].text.replace(/^_xlpm\./i, '').toLowerCase(), sh);
      } else last = sh;
    }
    return last;
  }

  /** INDEX(array, row, [col]) returns one cell when the positions pick one (a literal 0 / a missing one on a 2-D
   * array returns a whole row / column) */
  private indexShape(c: CallInfo): Shape {
    const lit = (k: number): number | null => {
      const a = c.args[k];
      if (!a) return null;
      const [s, e] = trimRange(this.toks, a);
      if (e - s === 0) return 0; // INDEX(A1:B3,,2): omitted = 0
      return e - s === 1 && this.toks[s].type === 'number' ? Number(this.toks[s].text) : NaN; // NaN = computed
    };
    const row = lit(1), col = lit(2);
    if (row === 0 || col === 0) return 'multi';
    if (row === null) return 'multi';
    if (col === null && !this.oneDimensional(c.args[0])) return 'multi';
    return 'scalar';
  }

  /** a single A1 ref / structured ref that is one row or one column */
  private oneDimensional(arg: [number, number]): boolean {
    const [s, e] = trimRange(this.toks, arg);
    if (e - s !== 1) return false;
    const t = this.toks[s];
    if (t.type === 'ref' && t.area) {
      const a = t.area;
      if (a.kind === 'cell') return true;
      if (a.kind === 'cols') return a.c1 === a.c2;
      if (a.kind === 'rows') return a.r1 === a.r2;
      return a.c1 === a.c2 || a.r1 === a.r2;
    }
    if (t.type === 'sref') {
      const info = parseStructuredRef(t.text.slice(t.prefix?.text.length ?? 0));
      return !!info && info.columns.length === 1 && !t.text.includes(']:[');
    }
    return false;
  }
}

/** structured reference: one cell when it names one column in one row ([#This Row], [#Headers], [#Totals]) */
function srefShape(t: Token): Shape {
  const info = parseStructuredRef(t.text.slice(t.prefix?.text.length ?? 0));
  if (!info) return 'multi';
  const oneRow = info.thisRow || (info.specials.length > 0 && info.specials.every((x) => /^#(headers|totals)$/i.test(x)));
  const oneCol = info.columns.length === 1 && !t.text.includes(']:[');
  return oneRow && oneCol ? 'scalar' : 'multi';
}
