/**
 * Formula text forms (owner: formula engine F).
 *   FILE form: what xlsx stores (no '=', _xlfn./_xlws./_xlpm./_xlop. prefixes, _xlfn.ANCHORARRAY(X), _xlfn.SINGLE(x),
 *              Table1[[#This Row],[c]], Table1[]).
 *   UI form:   what Excel shows / the user types ('=' prefix, no prefixes, X#, @x, Table1[@c], [@c], Table1).
 *
 * Verified against real Excel 365 (tests/fixtures/formula/forms-excel-results.json: formulas typed through
 * Range.Formula2, saved, <f> read back from the XML; see tests/formula/forms.test.ts):
 *   - every function of FUTURE_FUNCTIONS is stored as _xlfn.NAME, FILTER / SORT as _xlfn._xlws.NAME;
 *   - LET / LAMBDA parameters are stored as _xlpm.name (declaration and every use in scope, also calls "f(2)"),
 *     an optional LAMBDA parameter "[b]" is declared as _xlop.b;
 *   - A1# -> _xlfn.ANCHORARRAY(A1); @x -> _xlfn.SINGLE(x) (parentheses of "@(...)" are kept inside);
 *   - trim ranges (forms-tro-excel-results.json / tro-ops-excel-results.json): X!B1:.B10 -> _xlfn._TRO_TRAILING(X!B1:B10),
 *     B1.:B10 -> _TRO_LEADING, B1.:.B10 -> _TRO_ALL, whole columns / rows too; Data!#REF!:. -> _TRO_TRAILING(Data!#REF!);
 *   - Tbl[@a] -> Tbl[[#This Row],[a]], Tbl[@[b c]] -> Tbl[[#This Row],[b c]], Tbl[@] -> Tbl[#This Row],
 *     [@a] inside the table -> Tbl[[#This Row],[a]], bare Tbl -> Tbl[];
 *   - eta-reduced lambdas: a bare function name in a LAMBDA position -> _xleta.NAME (GROUPBY(r,v,SUM),
 *     BYROW(r,STDEV.S) -> _xleta.STDEV.S without _xlfn., HSTACK(SUM,AVERAGE) inside GROUPBY, LET(f,SUM,f(x)));
 *   - Excel's spelling (tests/fixtures/formula/forms-extra-excel-results.json): function names (also unknown ones),
 *     A1 references, booleans and error codes upper-cased; LET parameters take their declared case; table / defined
 *     names / sheet prefixes take their defined case when FormOptions lists them; [#this row] -> [#This Row].
 *     Unknown functions and "_xludf.X" are stored without _xlfn.; FOOBAR( stays FOOBAR(.
 * Writing back through Excel COM: toUiForm(<f>) is what Range.Formula shows for a normal cell and what Range.Formula2
 * shows for an array / dynamic-array cell (verified on 35,000+ formulas) — write legacy cells with Range.Formula,
 * dynamic ones (cm="1") and newly typed formulas with Range.Formula2. Whether a typed formula must be stored as a
 * dynamic array: ./dynamic.ts isDynamicArrayFormula.
 * toUiForm is the exact inverse on everything Excel writes: toFileForm(toUiForm(f, o), o) === f (tested on every
 * Excel-written formula of all fixtures). Known, harmless difference (documented in tests/formula/DEVIATIONS.md
 * "Formula forms"): Excel drops a typed '@' where legacy implicit intersection applies anyway (=@A1:A3 is stored as
 * A1:A3) — toFileForm keeps _xlfn.SINGLE(A1:A3), which Excel reads back as the same =@A1:A3.
 */
import { tokenize, renderArea, quoteSheetName, trimSuffix, type Area, type Prefix, type Token } from './excelRef';
import { matchBrackets, parseCalls, trimRange, funcName } from './calls';

// ------------------------------------------------------------------------------------------------ function lists
/**
 * Functions stored with the "_xlfn." prefix (MS-XLSX 2.2.2 future functions, checked one by one in Excel 365 build
 * 20430 — all verified except COPILOT, QUERYSTRING, PQSOURCE, LONGTEXT, PYTHON_*, ANCHORARRAY which that build did not
 * know / would not accept typed; they follow the MS-XLSX list). Not prefixed (verified): ECMA.CEILING, ISO.CEILING,
 * NETWORKDAYS.INTL, WORKDAY.INTL and every pre-2010 function. IMPORTTEXT / IMPORTCSV (2025) verified separately
 * (forms-new-excel-results.json). _TRO_* are the trim-range operators (see TRO_FUNC).
 */
export const FUTURE_FUNCTIONS: ReadonlySet<string> = new Set(`ACOT ACOTH AGGREGATE ANCHORARRAY ARABIC ARRAYTOTEXT BASE BETA.DIST
BETA.INV BINOM.DIST BINOM.DIST.RANGE BINOM.INV BITAND BITLSHIFT BITOR BITRSHIFT BITXOR BYCOL BYROW CEILING.MATH
CEILING.PRECISE CHISQ.DIST CHISQ.DIST.RT CHISQ.INV CHISQ.INV.RT CHISQ.TEST CHOOSECOLS CHOOSEROWS COMBINA CONCAT
CONFIDENCE.NORM CONFIDENCE.T COPILOT COT COTH COVARIANCE.P COVARIANCE.S CSC CSCH DAYS DECIMAL DETECTLANGUAGE DROP
ENCODEURL ERF.PRECISE ERFC.PRECISE EXPAND EXPON.DIST F.DIST F.DIST.RT F.INV F.INV.RT F.TEST FIELDVALUE FILTERXML
FLOOR.MATH FLOOR.PRECISE FORECAST.ETS FORECAST.ETS.CONFINT FORECAST.ETS.SEASONALITY FORECAST.ETS.STAT
FORECAST.LINEAR FORMULATEXT GAMMA GAMMA.DIST GAMMA.INV GAMMALN.PRECISE GAUSS GROUPBY HSTACK HYPGEOM.DIST IFNA IFS
IMAGE IMCOSH IMCOT IMCSC IMCSCH IMSEC IMSECH IMSINH IMTAN ISFORMULA ISOMITTED ISOWEEKNUM LAMBDA LET LOGNORM.DIST
LOGNORM.INV LONGTEXT MAKEARRAY MAP MAXIFS MINIFS MODE.MULT MODE.SNGL MUNIT NEGBINOM.DIST NORM.DIST NORM.INV
NORM.S.DIST NORM.S.INV NUMBERVALUE PDURATION PERCENTILE.EXC PERCENTILE.INC PERCENTOF PERCENTRANK.EXC PERCENTRANK.INC
PERMUTATIONA PHI PIVOTBY POISSON.DIST PQSOURCE PYTHON_STR PYTHON_TYPE PYTHON_TYPENAME QUARTILE.EXC QUARTILE.INC
QUERYSTRING RANDARRAY RANK.AVG RANK.EQ REDUCE REGEXEXTRACT REGEXREPLACE REGEXTEST RRI SCAN SEC SECH SEQUENCE SHEET
SHEETS SINGLE SKEW.P SORTBY STDEV.P STDEV.S STOCKHISTORY SWITCH T.DIST T.DIST.2T T.DIST.RT T.INV T.INV.2T T.TEST TAKE
TEXTAFTER TEXTBEFORE TEXTJOIN TEXTSPLIT TOCOL TOROW TRANSLATE TRIMRANGE UNICHAR UNICODE UNIQUE VALUETOTEXT VAR.P VAR.S
VSTACK WEBSERVICE WEIBULL.DIST WRAPCOLS WRAPROWS XLOOKUP XMATCH XOR Z.TEST IMPORTCSV IMPORTTEXT _TRO_ALL _TRO_LEADING _TRO_TRAILING`.split(/\s+/).filter(Boolean));

/**
 * Trim-range operators (UI A1:.B9 / A1.:B9 / A1.:.B9) and the file-form functions Excel stores them as (verified:
 * tests/fixtures/formula/forms-tro-excel-results.json, e.g. SUM(B1:.B10) -> SUM(_xlfn._TRO_TRAILING(B1:B10))).
 */
const TRO_FUNC: Readonly<Record<'trailing' | 'leading' | 'all', string>> = { trailing: '_TRO_TRAILING', leading: '_TRO_LEADING', all: '_TRO_ALL' };
const TRO_KIND: Readonly<Record<string, 'trailing' | 'leading' | 'all'>> = { _TRO_TRAILING: 'trailing', _TRO_LEADING: 'leading', _TRO_ALL: 'all' };
const TRO_SEP: Readonly<Record<'trailing' | 'leading' | 'all', string>> = { trailing: ':.', leading: '.:', all: '.:.' };

/** Worksheet-only functions stored as "_xlfn._xlws.NAME" (verified: FILTER, SORT; PY per MS-XLSX). */
export const XLWS_FUNCTIONS: ReadonlySet<string> = new Set(['FILTER', 'SORT', 'PY']);

/**
 * Built-in worksheet functions stored WITHOUT a prefix (Excel 2007 set + the 2010 names Excel 365 stores bare, verified:
 * ECMA.CEILING, ISO.CEILING, NETWORKDAYS.INTL, WORKDAY.INTL). Used to recognise eta-reduced lambdas ("BYROW(r,SUM)").
 */
export const LEGACY_FUNCTIONS: ReadonlySet<string> = new Set(`ABS ACCRINT ACCRINTM ACOS ACOSH ADDRESS AMORDEGRC AMORLINC AND
AREAS ASC ASIN ASINH ATAN ATAN2 ATANH AVEDEV AVERAGE AVERAGEA AVERAGEIF AVERAGEIFS BAHTTEXT BESSELI BESSELJ BESSELK
BESSELY BETADIST BETAINV BIN2DEC BIN2HEX BIN2OCT BINOMDIST CALL CEILING CELL CHAR CHIDIST CHIINV CHITEST CHOOSE CLEAN CODE
COLUMN COLUMNS COMBIN COMPLEX CONCATENATE CONFIDENCE CONVERT CORREL COS COSH COUNT COUNTA COUNTBLANK COUNTIF COUNTIFS
COUPDAYBS COUPDAYS COUPDAYSNC COUPNCD COUPNUM COUPPCD COVAR CRITBINOM CUBEKPIMEMBER CUBEMEMBER CUBEMEMBERPROPERTY
CUBERANKEDMEMBER CUBESET CUBESETCOUNT CUBEVALUE CUMIPMT CUMPRINC DATE DATEDIF DATEVALUE DAVERAGE DAY DAYS360 DB DBCS
DCOUNT DCOUNTA DDB DEC2BIN DEC2HEX DEC2OCT DEGREES DELTA DEVSQ DGET DISC DMAX DMIN DOLLAR DOLLARDE DOLLARFR DPRODUCT
DSTDEV DSTDEVP DSUM DURATION DVAR DVARP ECMA.CEILING EDATE EFFECT EOMONTH ERF ERFC ERROR.TYPE EUROCONVERT EVEN EXACT EXP
EXPONDIST FACT FACTDOUBLE FDIST FIND FINDB FINV FISHER FISHERINV FIXED FLOOR FORECAST FREQUENCY FTEST FV FVSCHEDULE
GAMMADIST GAMMAINV GAMMALN GCD GEOMEAN GESTEP GETPIVOTDATA GROWTH HARMEAN HEX2BIN HEX2DEC HEX2OCT HLOOKUP HOUR HYPERLINK
HYPGEOMDIST IF IFERROR IMABS IMAGINARY IMARGUMENT IMCONJUGATE IMCOS IMDIV IMEXP IMLN IMLOG10 IMLOG2 IMPOWER IMPRODUCT
IMREAL IMSIN IMSQRT IMSUB IMSUM INDEX INDIRECT INFO INT INTERCEPT INTRATE IPMT IRR ISBLANK ISERR ISERROR ISEVEN ISLOGICAL
ISNA ISNONTEXT ISNUMBER ISO.CEILING ISODD ISPMT ISREF ISTEXT JIS KURT LARGE LCM LEFT LEFTB LEN LENB LINEST LN LOG LOG10
LOGEST LOGINV LOGNORMDIST LOOKUP LOWER MATCH MAX MAXA MDETERM MDURATION MEDIAN MID MIDB MIN MINA MINUTE MINVERSE MIRR
MMULT MOD MODE MONTH MROUND MULTINOMIAL N NA NEGBINOMDIST NETWORKDAYS NETWORKDAYS.INTL NOMINAL NORMDIST NORMINV NORMSDIST
NORMSINV NOT NOW NPER NPV OCT2BIN OCT2DEC OCT2HEX ODD ODDFPRICE ODDFYIELD ODDLPRICE ODDLYIELD OFFSET OR PEARSON
PERCENTILE PERCENTRANK PERMUT PHONETIC PI PMT POISSON POWER PPMT PRICE PRICEDISC PRICEMAT PROB PRODUCT PROPER PV QUARTILE
QUOTIENT RADIANS RAND RANDBETWEEN RANK RATE RECEIVED REGISTER.ID REPLACE REPLACEB REPT RIGHT RIGHTB ROMAN ROUND ROUNDDOWN
ROUNDUP ROW ROWS RSQ RTD SEARCH SEARCHB SECOND SERIESSUM SIGN SIN SINH SKEW SLN SLOPE SMALL SQRT SQRTPI STANDARDIZE
STDEV STDEVA STDEVP STDEVPA STEYX SUBSTITUTE SUBTOTAL SUM SUMIF SUMIFS SUMPRODUCT SUMSQ SUMX2MY2 SUMX2PY2 SUMXMY2 SYD T
TAN TANH TBILLEQ TBILLPRICE TBILLYIELD TDIST TEXT TIME TIMEVALUE TINV TODAY TRANSPOSE TREND TRIM TRIMMEAN TRUNC TTEST
TYPE UPPER USDOLLAR VALUE VAR VARA VARP VARPA VDB VLOOKUP WEEKDAY WEEKNUM WEIBULL WORKDAY WORKDAY.INTL XIRR XNPV YEAR
YEARFRAC YIELD YIELDDISC YIELDMAT ZTEST`.split(/\s+/).filter(Boolean));

/** true for any built-in worksheet function name (case-insensitive, without prefixes). */
export function isBuiltinFunction(name: string): boolean {
  const up = name.replace(PREFIX_ANY_RE, '').toUpperCase();
  return LEGACY_FUNCTIONS.has(up) || FUTURE_FUNCTIONS.has(up) || XLWS_FUNCTIONS.has(up);
}

/**
 * File-form spelling of a function name typed by the user, upper-cased like Excel does ("xlookup" ->
 * "_xlfn.XLOOKUP", "filter" -> "_xlfn._xlws.FILTER", "sum" -> "SUM", unknown "foobar" -> "FOOBAR"). A typed
 * "_xlfn." / "_xlws." prefix is normalised; "_xludf." (user-defined function marker) is kept as written.
 */
export function fileFunctionName(name: string): string {
  if (/^_xludf\./i.test(name)) return '_xludf.' + name.slice(7);
  const hadXlfn = /^_xlfn\./i.test(name);
  const up = name.replace(/^(?:_xlfn\.|_xlws\.)+/i, '').toUpperCase();
  if (XLWS_FUNCTIONS.has(up)) return '_xlfn._xlws.' + up;
  if (FUTURE_FUNCTIONS.has(up) || hadXlfn) return '_xlfn.' + up;
  return up;
}

/**
 * Argument positions (0-based; -1 = last argument) that take a LAMBDA, where a bare function name is an eta-reduced
 * lambda stored as "_xleta.NAME" (verified: GROUPBY(r,v,SUM) -> _xleta.SUM, also inside HSTACK(SUM,AVERAGE); BYROW,
 * BYCOL, MAP, REDUCE, SCAN, PIVOTBY; future functions stay without _xlfn.: _xleta.STDEV.S, _xleta.PERCENTOF).
 */
const ETA_ARGS: Readonly<Record<string, readonly number[]>> = {
  BYROW: [1], BYCOL: [1], MAP: [-1], REDUCE: [2], SCAN: [2], MAKEARRAY: [2], GROUPBY: [2], PIVOTBY: [3],
};
const PREFIX_ANY_RE = /^(?:_xlfn\.|_xlws\.|_xlpm\.|_xleta\.)+/i;

// ------------------------------------------------------------------------------------------------ options
export interface FormOptions {
  /**
   * Table (ListObject) that contains the formula cell. UI form omits the table name for references to it
   * ("[@単価]", "[単価]"); the file form always names the table. Without it, unqualified refs stay unqualified.
   */
  hostTable?: string;
  /**
   * Table names of the workbook: a bare table name "Tbl" (UI) is stored as "Tbl[]", and a table name typed in another
   * case is written with its defined case ("tblsales[a]" -> "TblSales[a]"), like Excel (toFileForm only).
   */
  tables?: readonly string[];
  /**
   * Defined names of the workbook (toFileForm only): written with their defined case ("nmrange" -> "nmRange") and never
   * mistaken for an eta-reduced function ("PRICE" as a name vs. the PRICE function).
   */
  names?: readonly string[];
  /** Sheet names of the workbook (toFileForm only): a sheet prefix typed in another case gets the real case (x!B1 -> X!B1). */
  sheets?: readonly string[];
}

const PREFIX_RE = /^(?:_xlfn\.|_xlws\.|_xlpm\.)+/i;
/** canonical spelling of structured-reference special items */
const SPECIALS: Readonly<Record<string, string>> = { 'this row': 'This Row', all: 'All', data: 'Data', headers: 'Headers', totals: 'Totals' };
/** single-token operands of '@' (Excel shows _xlfn.SINGLE(1) as @1) */
const PRIMARY: ReadonlySet<string> = new Set(['ref', 'name', 'sref', 'error', 'number', 'string', 'bool']);
const eqi = (a: string | null | undefined, b: string | null | undefined) => a != null && b != null && a.toLowerCase() === b.toLowerCase();

// ------------------------------------------------------------------------------------------------ structured refs
/** Split an sref token text into [prefix+table, body] where body is the text inside the outermost brackets. */
function splitSref(t: Token): { head: string; table: string; body: string } | null {
  const pre = t.prefix?.text ?? '';
  const rest = t.text.slice(pre.length);
  const lb = rest.indexOf('[');
  if (lb < 0 || !rest.endsWith(']')) return null;
  return { head: pre, table: rest.slice(0, lb), body: rest.slice(lb + 1, -1) };
}
/** Items of a bracketed list body "[a],[b]:[c]" -> ['[a]', ',', '[b]', ':', '[c]'] (null if not such a list). */
function srefItems(body: string): string[] | null {
  const out: string[] = [];
  let k = 0;
  while (k < body.length) {
    const ch = body[k];
    if (ch === '[') {
      let e = k + 1;
      for (; e < body.length; e++) { if (body[e] === "'") { e++; continue; } if (body[e] === ']') break; }
      if (e >= body.length) return null;
      out.push(body.slice(k, e + 1)); k = e + 1;
    } else if (ch === ',' || ch === ':') { out.push(ch); k++; }
    else if (ch === ' ') { k++; } // Excel accepts spaces between items; they are not kept in the file
    else return null;
  }
  return out;
}
/** Column names that Excel writes without the inner brackets in "[@name]" (no special characters). */
const SIMPLE_COL = /^[^\s,:.[\]#'"{}$^&*+=\-<>/@!%()~;?\\|`]+$/u;
const isThisRow = (item: string) => /^\[#this row\]$/i.test(item);

function srefToUi(t: Token, opts: FormOptions): string {
  const s = splitSref(t);
  if (!s) return t.text;
  let table = s.table, body = s.body;
  if (body === '' && table) return s.head + table; // Tbl[] -> Tbl
  if (/^#this row$/i.test(body)) body = '@';
  else {
    const items = srefItems(body);
    if (items && items.length >= 3 && isThisRow(items[0]) && items[1] === ',') {
      const rest = items.slice(2);
      if (rest.length === 1) { const name = rest[0].slice(1, -1); body = SIMPLE_COL.test(name) ? '@' + name : '@' + rest[0]; }
      else if (rest.length === 3 && rest[1] === ':' && rest[0].startsWith('[') && rest[2].startsWith('[')) body = '@' + rest.join('');
    }
  }
  if (opts.hostTable && !s.head && eqi(table, opts.hostTable)) table = '';
  return s.head + table + '[' + body + ']';
}

/** "[#this row]" -> "[#This Row]" etc. (only bracketed items that start with a bare '#'; "'#" is an escaped char) */
const canonSpecials = (body: string) =>
  body.replace(/\[#(this row|all|data|headers|totals)\]/gi, (_m, k: string) => '[#' + SPECIALS[k.toLowerCase()] + ']')
    .replace(/^#(this row|all|data|headers|totals)$/i, (_m, k: string) => '#' + SPECIALS[k.toLowerCase()]);

function srefToFile(t: Token, opts: FormOptions): string {
  const s = splitSref(t);
  if (!s) return t.text;
  let body = s.body;
  let table = s.table || (!s.head && opts.hostTable ? opts.hostTable : s.table);
  if (table && !s.head && opts.tables) { const real = opts.tables.find((x) => eqi(x, table)); if (real) table = real; }
  if (body === '@') body = '#This Row';
  else if (body.startsWith('@')) {
    const rest = body.slice(1).trim();
    body = rest.startsWith('[') ? '[#This Row],' + rest : '[#This Row],[' + rest + ']';
  }
  return s.head + table + '[' + canonSpecials(body) + ']';
}

// ------------------------------------------------------------------------------------------------ toUiForm
/** FILE form (no '=') -> UI form with leading '='. */
export function toUiForm(fileFormula: string, opts: FormOptions = {}): string {
  const toks = tokenize(fileFormula);
  const br = matchBrackets(toks);
  const calls = parseCalls(toks, br);
  const isPrimary = (s: number, e: number): boolean => {
    if (e - s === 1) return PRIMARY.has(toks[s].type);
    const t = toks[s];
    if (t.type === 'func') { const c = calls.get(s); return !!c && c.close === e - 1; }
    if (t.type === 'op' && t.text === '(') return br[s] === e - 1;
    return false;
  };
  const emit = (from: number, to: number): string => {
    let out = '';
    for (let i = from; i < to; i++) {
      const t = toks[i];
      if (t.type === 'func') {
        const c = calls.get(i);
        const name = funcName(t);
        const pre = t.prefix?.text ?? '';
        if (c && c.close >= 0 && c.close < to && c.args.length === 1 && /^_xlfn\./i.test(t.text.slice(pre.length)) && !pre) {
          const [s, e] = trimRange(toks, c.args[0]);
          if (name === 'ANCHORARRAY' && e - s === 1 && ['ref', 'name', 'error'].includes(toks[s].type)) {
            out += emit(s, e) + '#'; i = c.close; continue;
          }
          if (name === 'SINGLE' && e > s && isPrimary(s, e)) { out += '@' + emit(s, e); i = c.close; continue; }
          // _xlfn._TRO_TRAILING(B1:B10) -> B1:.B10 (only the exact form Excel writes: one range token, no blanks)
          const tro = TRO_KIND[name];
          const r = toks[c.args[0][0]];
          if (tro && c.args[0][1] - c.args[0][0] === 1 && r.type === 'ref' && r.area && r.area.kind !== 'cell' && !r.area.trim) {
            const pre2 = r.prefix?.text ?? '';
            out += pre2 + r.text.slice(pre2.length).replace(':', TRO_SEP[tro]); i = c.close; continue;
          }
          // a trimmed range that became #REF!: _xlfn._TRO_TRAILING(Data!#REF!) -> Data!#REF!:. (verified)
          if (tro && c.args[0][1] - c.args[0][0] === 1 && r.type === 'error' && !r.trim && /^#REF!$/i.test(r.text.slice(r.prefix?.text.length ?? 0))) {
            out += r.text + TRO_SEP[tro]; i = c.close; continue;
          }
        }
        out += pre + t.text.slice(pre.length).replace(PREFIX_RE, '');
        continue;
      }
      if (t.type === 'name' && !t.prefix) {
        if (/^_xlpm\./i.test(t.text)) { out += t.text.slice(6); continue; }
        if (/^_xlop\./i.test(t.text)) { out += '[' + t.text.slice(6) + ']'; continue; }
        if (/^_xleta\./i.test(t.text)) { out += t.text.slice(7).replace(PREFIX_RE, ''); continue; }
      }
      if (t.type === 'sref') { out += srefToUi(t, opts); continue; }
      out += t.text;
    }
    return out;
  };
  return '=' + emit(0, toks.length);
}

// ------------------------------------------------------------------------------------------------ toFileForm
/** UI form (with or without '=') -> FILE form (no '='): adds _xlfn. to future functions, X# -> ANCHORARRAY, ... */
export function toFileForm(uiFormula: string, opts: FormOptions = {}): string {
  const body = uiFormula.startsWith('=') ? uiFormula.slice(1) : uiFormula;
  const toks = tokenize(body);
  const br = matchBrackets(toks);
  const calls = parseCalls(toks, br);
  const byLower = (list: readonly string[] | undefined) => new Map((list ?? []).map((x) => [x.toLowerCase(), x] as const));
  const tables = byLower(opts.tables);
  const names = byLower(opts.names);
  const sheets = byLower(opts.sheets);
  const lower = (s: string) => s.toLowerCase();
  /** LET / LAMBDA parameters in scope: lower-case name -> declared spelling (uses take the declared case, like Excel) */
  type Scope = ReadonlyMap<string, string>;

  /** declaration of a LET / LAMBDA parameter: [name token | optional "[x]" sref] -> file text, or null */
  const declare = (s: number, e: number, optional: boolean): { text: string; name: string } | null => {
    const [a, b] = trimRange(toks, [s, e]);
    if (b - a !== 1) return null;
    const t = toks[a];
    if (t.prefix) return null;
    const lead = toks.slice(s, a).map((x) => x.text).join(''), tail = toks.slice(b, e).map((x) => x.text).join('');
    // Excel even accepts a cell-like LET name (LET(Z99,1,1) is stored as _xlfn.LET(_xlpm.Z99,1,1))
    if (t.type === 'name' || (t.type === 'ref' && t.area?.kind === 'cell' && !t.text.includes('$'))) {
      if (/^_xl(pm|op)\./i.test(t.text)) return { text: lead + t.text + tail, name: t.text.slice(6) };
      return { text: lead + '_xlpm.' + t.text + tail, name: t.text };
    }
    if (optional && t.type === 'sref' && /^\[[^[\]]+\]$/.test(t.text)) return { text: lead + '_xlop.' + t.text.slice(1, -1) + tail, name: t.text.slice(1, -1) };
    return null;
  };

  // eta-reduced lambdas: a bare built-in function name in a LAMBDA position (ETA_ARGS, also inside a HSTACK/VSTACK
  // there), or as a LET value whose name is later called ("LET(f,SUM,f(x))") -> "_xleta.NAME"
  const eta = new Set<number>();
  const single = (arg: [number, number]): number => { const [s, e] = trimRange(toks, arg); return e - s === 1 ? s : -1; };
  const etaCandidate = (i: number) => i >= 0 && toks[i].type === 'name' && !toks[i].prefix && isBuiltinFunction(toks[i].text) &&
    !/^_xl/i.test(toks[i].text) && !names.has(lower(toks[i].text)) && !tables.has(lower(toks[i].text));
  for (const c of calls.values()) {
    const spec = ETA_ARGS[c.name];
    if (spec) {
      for (const p of spec) {
        const arg = c.args[p < 0 ? c.args.length + p : p];
        if (!arg) continue;
        const k = single(arg);
        if (etaCandidate(k)) { eta.add(k); continue; }
        const [s, e] = trimRange(toks, arg);
        const inner = calls.get(s);
        if (inner && inner.close === e - 1 && (inner.name === 'HSTACK' || inner.name === 'VSTACK')) {
          for (const ia of inner.args) { const j = single(ia); if (etaCandidate(j)) eta.add(j); }
        }
      }
    }
    if (c.name === 'LET' && c.close >= 0) {
      for (let k = 0; k + 2 < c.args.length; k += 2) {
        const d = single(c.args[k]), v = single(c.args[k + 1]);
        if (d < 0 || !etaCandidate(v)) continue;
        const pname = lower(toks[d].text.replace(/^_xlpm\./i, ''));
        let called = false;
        for (let j = c.args[k + 2][0]; j < c.close && !called; j++) {
          called = toks[j].type === 'func' && !toks[j].prefix && lower(toks[j].text.replace(/^_xlpm\./i, '')) === pname;
        }
        if (called) eta.add(v);
      }
    }
  }

  /** Excel's spelling of a reference token: upper-case A1 text, sheet prefix in the sheet's real case */
  const prefixText = (p: Prefix | undefined): string => {
    if (!p) return '';
    if (p.sheet && p.book === undefined && p.sheet2 === undefined) {
      const real = sheets.get(lower(p.sheet));
      if (real && real !== p.sheet) return quoteSheetName(real) + '!';
    }
    return p.text;
  };
  const refText = (t: Token): string => {
    let bodyText = t.text.slice(t.prefix?.text.length ?? 0);
    if (/^[$A-Za-z0-9:.]+$/.test(bodyText) && /[a-z]/.test(bodyText)) bodyText = bodyText.toUpperCase();
    return prefixText(t.prefix) + bodyText;
  };
  const errorText = (t: Token): string => prefixText(t.prefix) + t.text.slice(t.prefix?.text.length ?? 0).toUpperCase();

  /** end (exclusive) of the primary operand starting at j, or -1 */
  const operandEnd = (j: number, to: number): number => {
    const t = toks[j];
    if (!t || j >= to) return -1;
    let e = -1;
    if (t.type === 'func') { const c = calls.get(j); e = c && c.close >= 0 ? c.close + 1 : -1; }
    else if (t.type === 'op' && t.text === '(') e = br[j] >= 0 ? br[j] + 1 : -1;
    else if (PRIMARY.has(t.type)) e = j + 1;
    if (e > 0 && e < to && toks[e].type === 'op' && toks[e].text === '#' && ['ref', 'name', 'error'].includes(t.type)) e++;
    return e > to ? -1 : e;
  };

  const emit = (from: number, to: number, scope: Scope): string => {
    let out = '';
    for (let i = from; i < to; i++) {
      const t = toks[i];
      // X#  ->  _xlfn.ANCHORARRAY(X)
      if ((t.type === 'ref' || t.type === 'name' || t.type === 'error') && i + 1 < to && toks[i + 1].type === 'op' && toks[i + 1].text === '#') {
        out += '_xlfn.ANCHORARRAY(' + emit(i, i + 1, scope) + ')'; i++; continue;
      }
      // @x  ->  _xlfn.SINGLE(x)
      if (t.type === 'op' && t.text === '@') {
        const e = operandEnd(i + 1, to);
        if (e > 0) { out += '_xlfn.SINGLE(' + emit(i + 1, e, scope) + ')'; i = e - 1; continue; }
        out += t.text; continue;
      }
      if (t.type === 'func') {
        const pre = t.prefix?.text ?? '';
        const text = t.text.slice(pre.length);
        const isPm = /^_xlpm\./i.test(text);
        const param = pre ? undefined : scope.get(lower(isPm ? text.slice(6) : text));
        const c = calls.get(i);
        // a LET / LAMBDA parameter called as a function ("f(2)") -> _xlpm.f; a function of another book / sheet
        // ("[1]!MyFn(") is kept as written; everything else gets Excel's spelling (upper case, _xlfn. prefixes)
        const fname = param !== undefined ? '_xlpm.' + param : isPm ? text : pre ? t.text : fileFunctionName(text);
        const bare = param !== undefined || isPm || pre ? '' : funcName(t);
        if (c && c.close >= 0 && c.close < to && (bare === 'LET' || bare === 'LAMBDA')) {
          const n = c.args.length;
          const inner = new Map(scope);
          const parts: string[] = [];
          let pending: string | null = null; // LET: a name is visible from the NEXT argument after its value
          for (let k = 0; k < n; k++) {
            const [s, e] = c.args[k];
            const isLast = k === n - 1;
            if (bare === 'LAMBDA' && !isLast) {
              const d = declare(s, e, true);
              if (d) { inner.set(lower(d.name), d.name); parts.push(d.text); } else parts.push(emit(s, e, inner));
            } else if (bare === 'LET' && !isLast && k % 2 === 0) {
              const d = declare(s, e, false);
              parts.push(d ? d.text : emit(s, e, inner));
              pending = d ? d.name : null;
            } else {
              parts.push(emit(s, e, inner));
              if (pending) { inner.set(lower(pending), pending); pending = null; }
            }
          }
          out += fname + '(' + parts.join(',') + ')';
          i = c.close;
          continue;
        }
        out += fname;
        continue;
      }
      if (t.type === 'name' && !t.prefix) {
        const lo = lower(t.text);
        const param = scope.get(lo);
        if (param !== undefined) { out += '_xlpm.' + param; continue; }
        if (eta.has(i)) { out += '_xleta.' + t.text.toUpperCase(); continue; }
        const table = tables.get(lo);
        if (table !== undefined) { out += table + '[]'; continue; }
        const name = names.get(lo);
        if (name !== undefined) { out += name; continue; }
        out += t.text; continue;
      }
      if (t.type === 'ref') {
        // a cell-like LET / LAMBDA parameter ("LET(z99,1,z99+1)")
        const param = !t.prefix && t.area?.kind === 'cell' && !t.text.includes('$') ? scope.get(lower(t.text)) : undefined;
        if (param !== undefined) { out += '_xlpm.' + param; continue; }
        // trim range B1:.B10 -> _xlfn._TRO_TRAILING(B1:B10) (the prefix goes inside: _TRO_TRAILING(X!B1:B10))
        if (t.area?.trim) {
          const text = refText(t), pl = text.length - (t.text.length - (t.prefix?.text.length ?? 0));
          out += '_xlfn.' + TRO_FUNC[t.area.trim] + '(' + text.slice(0, pl) + text.slice(pl).replace(/\.?:\.?/, ':') + ')';
          continue;
        }
        out += refText(t); continue;
      }
      if (t.type === 'error') {
        if (t.trim) { // Data!#REF!:. -> _xlfn._TRO_TRAILING(Data!#REF!)
          const e = errorText(t);
          out += '_xlfn.' + TRO_FUNC[t.trim] + '(' + e.slice(0, e.length - trimSuffix(t.trim).length) + ')';
          continue;
        }
        out += errorText(t); continue;
      }
      if (t.type === 'bool') { out += t.text.toUpperCase(); continue; }
      if (t.type === 'sref') { out += srefToFile(t, opts); continue; }
      out += t.text;
    }
    return out;
  };
  return emit(0, toks.length, new Map());
}

// ------------------------------------------------------------------------------------------------ input detection
/** true when the typed text is a formula ('=...' but not just '=' / '=' followed by blanks). */
export function isFormulaInput(text: string): boolean {
  return /^=[\s\S]*\S/.test(text);
}

// ------------------------------------------------------------------------------------------------ rewriteRefs
/** A reference token as seen by rewriteRefs callbacks. `sheet` = resolved sheet name (host sheet when unqualified). */
export interface RefInfo {
  sheet: string | null;
  prefix: Prefix | undefined;
  area: Area;
  /** true when the reference had an explicit sheet prefix */
  qualified: boolean;
}

export interface RewriteOptions {
  /**
   * Keep defined names (and bare table names) as they are. Default false: a formula that uses a defined name is
   * unmappable (the name belongs to the source workbook), so rewriteRefs returns null.
   */
  keepNames?: boolean;
  /** Keep structured references (Table[...]) as they are. Default false (null). */
  keepStructured?: boolean;
  /** Pass external ([1]Sheet!A1) and 3D (S1:S2!A1) refs to `map`. Default false (null). */
  allowExternal?: boolean;
  /** Allow INDIRECT / OFFSET-by-text style functions whose targets are text (INDIRECT). Default false (null). */
  allowIndirect?: boolean;
  /**
   * Allow functions whose result depends on WHERE a reference or the formula cell is (POSITIONAL_FUNCTIONS: ROW,
   * COLUMN, OFFSET, CELL, SHEET). Default false (null): moving a formula to another row / column would silently change
   * its value (=ROW()-1 numbering, =COLUMN(B7), OFFSET(B7,0,1) that lands on another imported column).
   */
  allowPositional?: boolean;
}

/** Functions whose value depends on the position of the formula cell or of a reference (see allowPositional). */
export const POSITIONAL_FUNCTIONS: ReadonlySet<string> = new Set(['ROW', 'COLUMN', 'OFFSET', 'CELL', 'SHEET']);

/**
 * Rewrite every reference token (ref tokens only; strings/functions untouched) through `map` (FILE form in,
 * FILE form out). `map` returns the new area and optionally a new prefix text ('' = unqualified), or null when the
 * reference cannot be kept -> rewriteRefs returns null for the whole formula.
 * Conservative by default (premerge use): also returns null when the formula contains a defined name, a structured
 * reference, an external or 3D reference, INDIRECT, a position-dependent function (ROW, COLUMN, OFFSET, CELL,
 * SHEET), or an unknown token — see RewriteOptions to relax.
 * LET/LAMBDA parameters (_xlpm./_xlop.) and eta-reduced functions (_xleta.) are not names. Errors with a sheet
 * prefix ("Src!#REF!") are unmappable.
 */
export function rewriteRefs(
  formula: string,
  hostSheet: string | null,
  map: (ref: RefInfo) => { area: Area; prefixText?: string } | null,
  opts: RewriteOptions = {},
): string | null {
  const toks = tokenize(formula);
  let out = '';
  for (const t of toks) {
    switch (t.type) {
      case 'unknown': return null;
      case 'name':
        if (/^_xl(pm|op|eta)\./i.test(t.text) && !t.prefix) {
          if (!opts.allowPositional && /^_xleta\./i.test(t.text) && POSITIONAL_FUNCTIONS.has(t.text.slice(7).toUpperCase())) return null;
          break;
        }
        if (!opts.keepNames) return null;
        break;
      case 'sref':
        if (!opts.keepStructured) return null;
        break;
      case 'error':
        if (t.prefix) return null;
        break;
      case 'func':
        if (!opts.allowIndirect && funcName(t) === 'INDIRECT') return null;
        if (!opts.allowPositional && POSITIONAL_FUNCTIONS.has(funcName(t))) return null;
        if (t.prefix && !opts.allowExternal) return null; // [1]!MyFunc( : function of another workbook
        break;
      case 'ref': {
        const p = t.prefix;
        if (p && (p.book !== undefined || p.sheet2 !== undefined) && !opts.allowExternal) return null;
        const info: RefInfo = { sheet: p ? p.sheet : hostSheet, prefix: p, area: t.area!, qualified: !!p };
        const r = map(info);
        if (!r) return null;
        const prefixText = r.prefixText ?? (p ? p.text : '');
        const same = r.area === t.area || renderArea(r.area) === renderArea(t.area!);
        out += same && prefixText === (p ? p.text : '') ? t.text : prefixText + renderArea(r.area);
        continue;
      }
    }
    out += t.text;
  }
  return out;
}
