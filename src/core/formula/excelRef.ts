/**
 * excelRef.ts — Excel formula tokenizer + A1 reference parser + structural reference adjuster.
 *
 * Design goals
 *  - Lossless: tokens.map(t => t.text).join('') === input, for ANY input (unknown chars become 'unknown' tokens).
 *  - Only reference tokens are ever rewritten; everything else (spacing, case, strings, function names,
 *    unknown/unsupported functions, LET/LAMBDA, structured refs, external refs) is copied verbatim.
 *  - Excel semantics for insert/delete/move of columns/rows, verified against TWO independent Excel 365 ground
 *    truths (tests/formula/oracle|exhaustive|cf.test.ts = R2, tests/formula/r2b.test.ts = R2b, ~26,000 checks).
 *    Every intentional difference is listed in tests/formula/DEVIATIONS.md (e.g. a move never creates #REF! for
 *    refs to the last columns of the sheet, which Excel's insert+cut+delete does).
 *  - Works on the xlsx storage form (no leading '=', _xlfn. prefixes, [1]Sheet!A1 external refs,
 *    Table1[[#This Row],[col]], _xlfn._TRO_TRAILING(A1:B9)) AND on UI form ('=' prefix, A1#, @, [@col],
 *    'C:\p\[B.xlsx]S'!A1, trim ranges A1:.B9 / A1.:B9 / A1.:.B9 = Area.trim, "Data!#REF!:." = Token.trim).
 *    Conversion between the two forms: ./forms.ts (toUiForm / toFileForm / rewriteRefs).
 *
 * API overview (all pure; 1-based rows/cols; StructOp = insert/delete/move of columns or rows on ONE sheet)
 *  Tokens / areas
 *    tokenize(f) -> Token[]                          lossless; ref/name/sref/func/... with parsed prefix + area
 *    parseAreaText / renderArea / parsePrefix / quoteSheetName / colToNum / numToCol / listRefs / relKey
 *    parseStructuredRef(text)                        table name, column names, #This Row etc. of a Table[...] ref
 *  Structural ops
 *    adjustFormula(f, op, ctx) -> AdjustResult      ctx = {hostSheet, targetSheet, isName?, tables?, hostCell?}
 *        .formula / .changed / .createdRefError (dry-run a delete: would it create #REF!?) / .unsafe / .warnings
 *    adjustFormulaSeq(f, ops, ctx)                   several ops (planMove output); tokenizes once, the context
 *        follows each op (advanceContext(ctx, op): table ranges / column lists and the host cell after the op)
 *    adjustArea(area, op) / adjustNameArea(area, op) / adjustSqref("A1:B2 C3", op) / mapIndex(i, op) / movePerm
 *    adjustTableRef("C3:H10", op)                    ListObject range (moves inside the table keep it)
 *    tableMoveKind(lo, hi, moveOp) -> 'reorder' | 'whole' | 'shift' | 'refused'   how Excel treats a column move around
 *        a table: a block from inside landing anywhere in [lo, hi + 1] (right after the last column included!) only
 *        reorders the table columns; a block covering the table moves it; refused: out of / into / partly over it
 *    adjustAnchoredRule(rule, op, sheet, {kind:'cf'|'dv', type, inheritFormat?, siblings?}) -> AnchoredRule[]
 *        CF / DV (sqref + formulas anchored at the sqref bounding box top-left, as in xlsx). `type` = the xlsx
 *        <cfRule type> / <dataValidation type>: which rules grow over inserted cells and spread onto moved blocks
 *        (see AnchoredRuleOptions); adjustAnchoredRulePerCell (alias); sqrefTopLeft
 *        OPTION SHAPE (callers pass the rule's own kind and type, see AnchoredRuleOptions for the rules):
 *          { kind?: 'cf' | 'dv'            // default 'cf'
 *            type?: string                 // FeatureRule.type = xlsx <cfRule type> / <dataValidation type>; unknown = spreads
 *            inheritFormat?: boolean       // default true; false = premerge "clean insert" (no growth, no spread)
 *            siblings?: string[][] }       // formulas of the other rules of the same <conditionalFormatting> block
 *          every CF type and every DV grows over cells inserted right after it; on a move cellIs / expression /
 *          containsText ... and every DV type spread onto the block landing right after / inside the area, while
 *          colorScale, dataBar, iconSet, top10, uniqueValues, duplicateValues, aboveAverage (NO_SPREAD_CF_TYPES) do not.
 *          Old option name `cfType` still works when `type` is absent.
 *    planMove(selectedCols, beforeCol) -> StructOp[] decomposition of a multi-selection move (both engines replay it)
 *    translateFormula(f, dRow, dCol)                 copy / fill / shared-formula expansion
 *    renameSheetInFormula(f, old, new)
 *  Helpers for the doc / UI (new in the product version)
 *    refsTouchingColumns(f, hostSheet, targetSheet, cols, {tables?, hostCell?}) -> TouchingRef[]
 *        references of a formula that point at any of `cols` on targetSheet (A1 refs incl. whole rows, and
 *        structured refs through the table layout). Use for "which formulas depend on the selected columns".
 *    lookupWarnings(f, op, ctx, {resolveName?}) -> LookupWarning[]  {fn, ref, message(日本語)} for VLOOKUP /
 *        HLOOKUP / INDEX / OFFSET / MATCH / XMATCH / CHOOSE / XLOOKUP whose range argument is changed inside by
 *        `op`: Excel adjusts the range but not numeric positions (exact check when the index is a literal number).
 *        Range arguments: A1 refs (whole rows/columns too), tables / structured refs (ctx.tables; columns reordered
 *        inside a table), defined names (resolveName(name) -> refersTo).
 *
 * Zero dependencies (besides ./calls.ts). Runs in browser and Node.
 */

import { parseCalls, trimRange } from './calls';

export const MAX_COL = 16384; // XFD
export const MAX_ROW = 1048576;

export type TokType =
  | 'ws'        // whitespace (may be the intersection operator — irrelevant for rewriting)
  | 'string'    // "..."  (with "" escapes)
  | 'number'
  | 'bool'      // TRUE / FALSE (not followed by '(')
  | 'error'     // #REF! #N/A ... (and prefix + error, e.g. Sheet1!#REF!)
  | 'func'      // function name (text excludes '('), incl. _xlfn./_xlws./_xludf. prefixes
  | 'ref'       // A1 reference: cell, area, whole cols, whole rows (optional sheet/book prefix)
  | 'name'      // defined name / LET-LAMBDA parameter (optional prefix)
  | 'sref'      // structured reference: Table1[...], [@col], [[#This Row],[col]]
  | 'op'        // + - * / ^ & = <> <= >= < > % : , ; ( ) { } @ #
  | 'unknown';  // anything we could not classify — callers should refuse to rewrite the formula

export interface Area {
  kind: 'cell' | 'area' | 'cols' | 'rows';
  // 1-based. For 'cols' r1/r2 are 0; for 'rows' c1/c2 are 0.
  c1: number; r1: number; c2: number; r2: number;
  ac1: boolean; ar1: boolean; ac2: boolean; ar2: boolean; // $ flags
  lower?: boolean; // original was written in lower case (kept only as info)
  /**
   * Trim-range operator of the UI form (Excel 365 "trim refs"): 'trailing' = A1:.B9, 'leading' = A1.:B9,
   * 'all' = A1.:.B9. The file form stores them as _xlfn._TRO_TRAILING(A1:B9) / _TRO_LEADING / _TRO_ALL (verified,
   * forms-tro-excel-results.json); structural ops adjust the range exactly like a plain range.
   */
  trim?: 'leading' | 'trailing' | 'all';
}

export interface Prefix {
  text: string;          // verbatim prefix incl. trailing '!'  ('' if none)
  sheet: string | null;  // unquoted sheet name (null if none / 3D / book-only)
  sheet2?: string;       // end sheet of a 3D ref (Sheet1:Sheet3!A1)
  book?: string;         // external workbook ('1' in xlsx form, 'Book.xlsx' or 'C:\p\[Book.xlsx]' in UI form)
}

export interface Token {
  type: TokType;
  text: string;
  start: number;
  end: number;
  prefix?: Prefix;       // for ref/name/error tokens written with a prefix
  area?: Area;           // for ref tokens
  /**
   * 'error' tokens only: a trimmed range that became #REF! keeps its trim marker in the UI form ("Data!#REF!:.",
   * "#REF!.:", "#REF!.:." = file _xlfn._TRO_TRAILING / _LEADING / _ALL(Data!#REF!), verified). The marker is part
   * of the token text.
   */
  trim?: Area['trim'];
}

// ---------------------------------------------------------------------------------------------
// character classes
// ---------------------------------------------------------------------------------------------
const reIdStart = /[\p{L}_\\]/u;
// Excel names may contain letters, digits, _, ., \, ? ; unquoted sheet names additionally contain
// many non-ASCII symbols (e.g. '・', '①'). We accept any non-ASCII char that is not whitespace and
// not a known full-width/ideographic punctuation Excel quotes. Being permissive here is safe:
// an unquoted run is only treated as a sheet prefix when followed by '!'.
const reIdChar = /[\p{L}\p{N}\p{M}_.\\?$\u00B7\u30FB\u2460-\u24FF]/u;
const isDigit = (c: string) => c >= '0' && c <= '9';
// ASCII part of reIdChar as a lookup table (the tokenizer's hot loop)
const ASCII_ID = new Uint8Array(128);
for (let c = 0; c < 128; c++) ASCII_ID[c] = reIdChar.test(String.fromCharCode(c)) ? 1 : 0;
/** char k may continue an identifier-ish run: reIdChar, or any non-ASCII char that is not blank / quote-only punct */
function isRunChar(s: string, k: number): boolean {
  const code = s.charCodeAt(k);
  if (code < 128) return ASCII_ID[code] === 1;
  const ch = s[k];
  return reIdChar.test(ch) || (!/\s/.test(ch) && !isQuotedOnlyPunct(ch));
}

const ERRORS = ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#GETTING_DATA',
  '#SPILL!', '#CALC!', '#FIELD!', '#BLOCKED!', '#CONNECT!', '#BUSY!', '#UNKNOWN!', '#PYTHON!', '#EXTERNAL!'];

export function colToNum(s: string): number {
  let n = 0;
  for (const ch of s.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}
export function numToCol(n: number): string {
  let s = '';
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

const reCell = /^(\$?)([A-Za-z]{1,3})(\$?)([0-9]{1,7})$/;
const reColOnly = /^(\$?)([A-Za-z]{1,3})$/;
const reRowOnly = /^(\$?)([0-9]{1,7})$/;

function parseCell(s: string): { c: number; r: number; ac: boolean; ar: boolean } | null {
  const m = reCell.exec(s);
  if (!m) return null;
  const c = colToNum(m[2]); const r = Number(m[4]);
  if (c < 1 || c > MAX_COL || r < 1 || r > MAX_ROW) return null;
  return { c, r, ac: m[1] === '$', ar: m[3] === '$' };
}
function parseCol(s: string): { c: number; ac: boolean } | null {
  const m = reColOnly.exec(s); if (!m) return null;
  const c = colToNum(m[2]); if (c > MAX_COL) return null;
  return { c, ac: m[1] === '$' };
}
function parseRow(s: string): { r: number; ar: boolean } | null {
  const m = reRowOnly.exec(s); if (!m) return null;
  const r = Number(m[2]); if (r < 1 || r > MAX_ROW) return null;
  return { r, ar: m[1] === '$' };
}

/**
 * Parse the part after an optional prefix: A1, A1:B2, A:B, 1:2 (exact match of the whole string); also the UI-form
 * trim ranges A1:.B2 / A1.:B2 / A1.:.B2 (Area.trim).
 */
export function parseAreaText(s: string): Area | null {
  const i = s.indexOf(':');
  if (i < 0) {
    const a = parseCell(s);
    return a ? { kind: 'cell', c1: a.c, r1: a.r, c2: a.c, r2: a.r, ac1: a.ac, ar1: a.ar, ac2: a.ac, ar2: a.ar } : null;
  }
  let L = s.slice(0, i), R = s.slice(i + 1);
  let trim: Area['trim'];
  const lead = L.endsWith('.'), trail = R.startsWith('.');
  if (lead || trail) {
    if (lead) L = L.slice(0, -1);
    if (trail) R = R.slice(1);
    trim = lead && trail ? 'all' : lead ? 'leading' : 'trailing';
  }
  let res: Area | null = null;
  const a = parseCell(L), b = a ? parseCell(R) : null;
  if (a && b) res = { kind: 'area', c1: a.c, r1: a.r, c2: b.c, r2: b.r, ac1: a.ac, ar1: a.ar, ac2: b.ac, ar2: b.ar };
  else {
    const ca = parseCol(L), cb = ca ? parseCol(R) : null;
    if (ca && cb) res = { kind: 'cols', c1: ca.c, r1: 0, c2: cb.c, r2: 0, ac1: ca.ac, ar1: false, ac2: cb.ac, ar2: false };
    else {
      const ra = parseRow(L), rb = ra ? parseRow(R) : null;
      if (ra && rb) res = { kind: 'rows', c1: 0, r1: ra.r, c2: 0, r2: rb.r, ac1: false, ar1: ra.ar, ac2: false, ar2: rb.ar };
    }
  }
  if (res && trim) res.trim = trim;
  return res;
}

/** Unquote/split a prefix ("'O''Brien'!", "[1]Sheet1!", "Sheet1:Sheet3!", "'C:\p\[B.xlsx]S'!", "[1]!") */
export function parsePrefix(text: string): Prefix {
  let body = text.slice(0, -1); // drop '!'
  if (body.startsWith("'") && body.endsWith("'")) body = body.slice(1, -1).replace(/''/g, "'");
  const p: Prefix = { text, sheet: null };
  const ext = /^(.*\[[^\]]*\])(.*)$/s.exec(body);
  if (ext) {
    p.book = ext[1].startsWith('[') && ext[1].endsWith(']') && !ext[1].slice(1, -1).includes('[') ? ext[1].slice(1, -1) : ext[1];
    body = ext[2];
    if (!body) return p; // [1]!Name
  }
  const colon = body.indexOf(':');
  if (colon > 0 && !p.book) { p.sheet = null; p.sheet2 = body.slice(colon + 1); (p as any).sheet1 = body.slice(0, colon); return p; }
  p.sheet = body;
  return p;
}

// ---------------------------------------------------------------------------------------------
// tokenizer
// ---------------------------------------------------------------------------------------------
export function tokenize(formula: string): Token[] {
  const s = formula;
  const n = s.length;
  const out: Token[] = [];
  let i = 0;
  const push = (type: TokType, start: number, end: number, extra?: Partial<Token>) => {
    out.push({ type, text: s.slice(start, end), start, end, ...extra });
    i = end;
  };

  /** read an identifier-ish run starting at j; returns end index */
  const readRun = (j: number) => {
    let k = j;
    while (k < n && isRunChar(s, k)) k++;
    return k;
  };

  /** After a prefix ending at j (index just after '!'), read the body: #REF!, area, cols, rows, name. */
  const readBody = (start: number, j: number, prefix: Prefix): boolean => {
    // #REF! etc. after prefix
    if (s[j] === '#') {
      const e = matchError(s, j);
      if (e) { const tm = refErrorTrim(s, e, j + e.length); push('error', start, j + e.length + (tm?.len ?? 0), tm ? { prefix, trim: tm.trim } : { prefix }); return true; }
      return false;
    }
    const m = matchArea(s, j);
    if (m) { push('ref', start, m.end, { prefix, area: m.area }); return true; }
    const k = readRun(j);
    if (k > j) {
      if (s[k] === '[') { const e = matchBrackets(s, k); if (e > 0) { push('sref', start, e, { prefix }); return true; } }
      if (s[k] === '(') { push('func', start, k, { prefix }); return true; }
      push('name', start, k, { prefix }); return true;
    }
    return false;
  };

  while (i < n) {
    const c = s[i];
    const start = i;
    // whitespace
    if (c === ' ' || c === '\t' || c === '\r' || c === '\n' || c === '\u3000') {
      let k = i + 1; while (k < n && /[ \t\r\n\u3000]/.test(s[k])) k++;
      push('ws', start, k); continue;
    }
    // string
    if (c === '"') {
      let k = i + 1;
      for (;;) {
        if (k >= n) { push('unknown', start, n); break; }           // unterminated
        if (s[k] === '"') { if (s[k + 1] === '"') { k += 2; continue; } push('string', start, k + 1); break; }
        k++;
      }
      continue;
    }
    // quoted prefix  '...'!
    if (c === "'") {
      let k = i + 1;
      for (;;) {
        if (k >= n) break;
        if (s[k] === "'") { if (s[k + 1] === "'") { k += 2; continue; } break; }
        k++;
      }
      if (k < n && s[k + 1] === '!') {
        const pText = s.slice(start, k + 2);
        if (readBody(start, k + 2, parsePrefix(pText))) continue;
      }
      push('unknown', start, Math.min(n, k + 1)); continue;
    }
    // errors / spill operator
    if (c === '#') {
      const e = matchError(s, i);
      if (e) { const tm = refErrorTrim(s, e, i + e.length); push('error', start, i + e.length + (tm?.len ?? 0), tm ? { trim: tm.trim } : undefined); continue; }
      push('op', start, i + 1); continue;
    }
    // '[' : external-book prefix ([1]Sheet1!A1, [1]!Name) or structured ref ([@col], [[#This Row],[c]])
    if (c === '[') {
      const ext = /^\[[^\[\]]+\]/.exec(s.slice(i));
      if (ext) {
        const k = i + ext[0].length;
        if (s[k] === '!') { if (readBody(start, k + 1, parsePrefix(s.slice(start, k + 1)))) continue; }
        const k2 = readRun(k);
        if (k2 > k && s[k2] === '!') { if (readBody(start, k2 + 1, parsePrefix(s.slice(start, k2 + 1)))) continue; }
      }
      const e = matchBrackets(s, i);
      if (e > 0) { push('sref', start, e); continue; }
      push('unknown', start, i + 1); continue;
    }
    // numbers and row ranges (1:3, $2:$10)
    if (isDigit(c) || (c === '.' && isDigit(s[i + 1] ?? '')) || (c === '$' && isDigit(s[i + 1] ?? ''))) {
      const m = matchArea(s, i);
      if (m && m.area.kind === 'rows') { push('ref', start, m.end, { area: m.area }); continue; }
      const num = /^(?:[0-9]+\.?[0-9]*|\.[0-9]+)(?:[eE][+-]?[0-9]+)?/.exec(s.slice(i));
      if (num && c !== '$') {
        // a digit-leading unquoted sheet name is illegal in Excel, so no prefix check needed
        push('number', start, i + num[0].length); continue;
      }
      push('unknown', start, i + 1); continue;
    }
    // identifiers: functions, sheet prefixes, cells, column ranges, names, bools, table names
    if (reIdStart.test(c) || c === '$' || (s.charCodeAt(i) > 0x7f && !isQuotedOnlyPunct(c))) {
      const k = readRun(i);
      const word = s.slice(i, k);
      // sheet prefix  Sheet1!  /  3D  Sheet1:Sheet3!
      if (s[k] === '!' && !word.includes('$')) {
        if (readBody(start, k + 1, parsePrefix(s.slice(start, k + 1)))) continue;
      }
      if (s[k] === ':' && !word.includes('$')) {
        const k2 = readRun(k + 1);
        if (k2 > k + 1 && s[k2] === '!' && !s.slice(k + 1, k2).includes('$')) {
          if (readBody(start, k2 + 1, parsePrefix(s.slice(start, k2 + 1)))) continue;
        }
      }
      if (s[k] === '(' && !word.includes('$')) { push('func', start, k); continue; }
      const m = matchArea(s, i);
      if (m) { push('ref', start, m.end, { area: m.area }); continue; }
      if (/^(TRUE|FALSE)$/i.test(word)) { push('bool', start, k); continue; }
      if (s[k] === '[') { const e = matchBrackets(s, k); if (e > 0) { push('sref', start, e); continue; } }
      if (word.includes('$')) { push('unknown', start, k); continue; }
      push('name', start, k); continue;
    }
    // operators
    const two = s.slice(i, i + 2);
    if (two === '<>' || two === '<=' || two === '>=') { push('op', start, i + 2); continue; }
    if ('+-*/^&=<>%:,;(){}@!'.includes(c)) { push('op', start, i + 1); continue; }
    push('unknown', start, i + 1);
  }
  return out;
}

// full-width / ideographic punctuation that forces quoting (so never part of an unquoted run)
function isQuotedOnlyPunct(ch: string): boolean {
  return /[\u3001\u3002\uFF08\uFF09\uFF0C\uFF01\uFF1A\uFF1B\uFF1F\uFF3B\uFF3D\u300C\u300D\u3010\u3011\uFF0B\uFF0D\uFF0A\uFF0F\uFF1D\uFF1C\uFF1E\uFF06\uFF05\uFF03\uFF20\uFF3E]/.test(ch);
}

/** Trim marker (":." / ".:" / ".:.") right after a #REF! error at k (UI form of a trimmed range that became #REF!). */
function refErrorTrim(s: string, err: string, k: number): { len: number; trim: NonNullable<Area['trim']> } | null {
  if (err.length !== 5 || err.toUpperCase() !== '#REF!') return null;
  let r: { len: number; trim: NonNullable<Area['trim']> } | null = null;
  if (s[k] === ':' && s[k + 1] === '.') r = { len: 2, trim: 'trailing' };
  else if (s[k] === '.' && s[k + 1] === ':') r = s[k + 2] === '.' ? { len: 3, trim: 'all' } : { len: 2, trim: 'leading' };
  // never swallow the start of another operand ("#REF!:.5", "#REF!.:A1" are not trim markers)
  if (r && k + r.len < s.length && (isRunChar(s, k + r.len) || isDigit(s[k + r.len]) || s[k + r.len] === '$')) return null;
  return r;
}
/** UI-form trim marker of a #REF! token ('' when not trimmed). */
export function trimSuffix(trim: Area['trim']): string {
  return trim === 'trailing' ? ':.' : trim === 'leading' ? '.:' : trim === 'all' ? '.:.' : '';
}

function matchError(s: string, i: number): string | null {
  const up = s.slice(i, i + 14).toUpperCase();
  for (const e of ERRORS) if (up.startsWith(e)) return s.slice(i, i + e.length);
  return null;
}

/** balanced [ ... ] with ' as escape char (structured refs); returns end index (exclusive) or -1 */
function matchBrackets(s: string, i: number): number {
  let depth = 0;
  for (let k = i; k < s.length; k++) {
    const ch = s[k];
    if (ch === "'") { k++; continue; }
    if (ch === '[') depth++;
    else if (ch === ']') { depth--; if (depth === 0) return k + 1; }
  }
  return -1;
}

// sticky: matched at lastIndex without slicing the formula
// ':' may carry the trim-range dots of the UI form (A1:.B2, A1.:B2, A1.:.B2, B:.B, 1.:3)
const reAreaHeadY = /(\$?[A-Za-z]{1,3}\$?[0-9]{1,7}(?:\.?:\.?\$?[A-Za-z]{1,3}\$?[0-9]{1,7})?|\$?[A-Za-z]{1,3}\.?:\.?\$?[A-Za-z]{1,3}|\$?[0-9]{1,7}\.?:\.?\$?[0-9]{1,7})/y;
/** Try to match an A1 area at i. Must not be followed by an identifier char / '(' / '!' / '[' (then it is a name/func/prefix). */
function matchArea(s: string, i: number): { end: number; area: Area } | null {
  reAreaHeadY.lastIndex = i;
  const m = reAreaHeadY.exec(s);
  if (!m) return null;
  let text = m[1];
  // A1:B2 where the second half is not a valid cell (e.g. A1:B2x) -> try just the first cell
  const tryText = (t: string) => {
    const end = i + t.length;
    const nx = s[end] ?? '';
    if (nx && (isRunChar(s, end) || nx === '(' || nx === '!' || nx === '[')) return null;
    const a = parseAreaText(t);
    return a ? { end, area: { ...a, lower: /[a-z]/.test(t) } } : null;
  };
  let r = tryText(text);
  if (!r && text.includes(':') && /^\$?[A-Za-z]/.test(text)) { const first = text.slice(0, text.indexOf(':')); if (parseCell(first)) r = tryText(first); }
  return r;
}

// ---------------------------------------------------------------------------------------------
// structural operations (Excel semantics)
// ---------------------------------------------------------------------------------------------
export type Axis = 'col' | 'row';
export type StructOp =
  | { type: 'insert'; axis: Axis; at: number; count: number }            // insert `count` before index `at` (1-based)
  | { type: 'delete'; axis: Axis; at: number; count: number }            // delete [at, at+count-1]
  | { type: 'move'; axis: Axis; from: number; count: number; to: number }; // Excel: Columns(from:from+count-1).Cut ; Columns(to).Insert
                                                                          // `to` = insert-before index in ORIGINAL coordinates

type Span = [number, number] | null; // null = #REF!

/** permutation of indices for a block move */
export function movePerm(from: number, count: number, to: number): (v: number) => number {
  const s1 = from, s2 = from + count - 1;
  if (to < s1) return (v) => (v >= s1 && v <= s2 ? to + (v - s1) : v >= to && v < s1 ? v + count : v);
  if (to > s2 + 1) return (v) => (v >= s1 && v <= s2 ? to - count + (v - s1) : v > s2 && v < to ? v - count : v);
  return (v) => v; // no-op move
}

function insertSpan(lo: number, hi: number, at: number, count: number, max: number, single: boolean): Span {
  // Excel: a range spanning the whole axis (1:1048576, A:XFD, A5:XFD5) never changes, even when inserting at 1 (R2b r04)
  if (!single && lo === 1 && hi === max) return [lo, hi];
  const l = lo >= at ? lo + count : lo;
  let h = hi >= at ? hi + count : hi;
  if (l > max) return null;
  if (h > max) { if (single) return null; h = max; }
  return [l, h];
}
function deleteSpan(lo: number, hi: number, at: number, count: number, max = Number.MAX_SAFE_INTEGER, single = false): Span {
  const s1 = at, s2 = at + count - 1;
  if (lo >= s1 && hi <= s2) return null;
  const l = lo < s1 ? lo : lo > s2 ? lo - count : s1;
  // Excel: a range (not a single cell) that reaches the last row/column stays anchored there:
  // $1:$1048576 minus a row is still $1:$1048576, and 1048576:1048576 minus row 5 is 1048575:1048576 (R2b x01..x07)
  const h = hi === max && !single ? max : hi > s2 ? hi - count : hi >= s1 ? s1 - 1 : hi;
  return [l, h];
}
/**
 * Excel "insert cut cells" (Cut columns/rows + Insert) semantics for a span [lo,hi].
 * Verified exhaustively against Excel 365 (128 block moves x 36 ranges, oracle/exhaustive-results.json):
 *
 *   p      = the index permutation caused by the move
 *   follow = [p(lo), p(hi)]                         (each endpoint follows its own cell)
 *   modelA = delete the block, then insert `count` at the destination (post-delete coordinates)
 *
 *   1. span entirely inside the moved block                          -> follow
 *   2. follow would be inverted (p(lo) > p(hi))                      -> modelA
 *   3. block entirely inside the span and an endpoint that lies in
 *      the block moves INWARD (left end moving right / right end
 *      moving left)                                                   -> modelA  (range keeps its extent)
 *   4. otherwise                                                      -> follow
 *
 * For spans disjoint from the block, follow == modelA (plain shift / expansion).
 */
function moveSpan(lo: number, hi: number, from: number, count: number, to: number): Span {
  const s1 = from, s2 = from + count - 1;
  if (to >= s1 && to <= s2 + 1) return [lo, hi]; // no-op
  const p = movePerm(from, count, to);
  const inB = (v: number) => v >= s1 && v <= s2;
  const modelA = (): Span => {
    const d = deleteSpan(lo, hi, s1, count)!;
    const at = to < s1 ? to : to - count;
    return insertSpan(d[0], d[1], at, count, Number.MAX_SAFE_INTEGER, false);
  };
  if (inB(lo) && inB(hi)) return [p(lo), p(hi)];
  const fl = p(lo), fh = p(hi);
  if (fl > fh) return modelA();
  const blockInside = lo <= s1 && s2 <= hi;
  const movingRight = to > s2 + 1;
  if (blockInside && ((inB(lo) && movingRight) || (inB(hi) && !movingRight))) return modelA();
  return [fl, fh];
}

function applySpan(lo: number, hi: number, op: StructOp, max: number, single: boolean): Span {
  switch (op.type) {
    case 'insert': return insertSpan(lo, hi, op.at, op.count, max, single);
    case 'delete': return deleteSpan(lo, hi, op.at, op.count, max, single);
    case 'move': return moveSpan(lo, hi, op.from, op.count, op.to);
  }
}

/** Adjust one area; returns null for #REF! */
export function adjustArea(a: Area, op: StructOp): Area | null {
  const isCol = op.axis === 'col';
  if (isCol && a.kind === 'rows') return a;   // whole-row refs unaffected by column ops
  if (!isCol && a.kind === 'cols') return a;
  const single = a.kind === 'cell';
  const lo = isCol ? a.c1 : a.r1, hi = isCol ? a.c2 : a.r2;
  // an area may be written "reversed" (B1:A1) — Excel keeps the written order; adjust by sorted span
  const rev = lo > hi;
  const r = applySpan(Math.min(lo, hi), Math.max(lo, hi), op, isCol ? MAX_COL : MAX_ROW, single);
  if (!r) return null;
  const [nl, nh] = rev ? [r[1], r[0]] : r;
  if (nl === lo && nh === hi) return a; // untouched: keep exactly as written
  const res: Area = isCol ? { ...a, c1: nl, c2: nh } : { ...a, r1: nl, r2: nh };
  return normalizeFull(res);
}

/** Excel renders an area spanning all columns as a row range (A1:XFD1 -> 1:1) and all rows as a column range. */
export function normalizeFull(a: Area): Area {
  if (a.kind !== 'area') return a;
  if (Math.min(a.c1, a.c2) === 1 && Math.max(a.c1, a.c2) === MAX_COL) return { ...a, kind: 'rows', c1: 0, c2: 0, ac1: false, ac2: false };
  if (Math.min(a.r1, a.r2) === 1 && Math.max(a.r1, a.r2) === MAX_ROW) return { ...a, kind: 'cols', r1: 0, r2: 0, ar1: false, ar2: false };
  return a;
}

export function renderArea(a: Area): string {
  const c = (n: number, abs: boolean) => (abs ? '$' : '') + numToCol(n);
  const r = (n: number, abs: boolean) => (abs ? '$' : '') + n;
  const sep = a.trim === undefined ? ':' : a.trim === 'trailing' ? ':.' : a.trim === 'leading' ? '.:' : '.:.';
  switch (a.kind) {
    case 'cell': return a.trim === undefined ? c(a.c1, a.ac1) + r(a.r1, a.ar1) : c(a.c1, a.ac1) + r(a.r1, a.ar1) + sep + c(a.c2, a.ac2) + r(a.r2, a.ar2);
    case 'area': return c(a.c1, a.ac1) + r(a.r1, a.ar1) + sep + c(a.c2, a.ac2) + r(a.r2, a.ar2);
    case 'cols': return c(a.c1, a.ac1) + sep + c(a.c2, a.ac2);
    case 'rows': return r(a.r1, a.ar1) + sep + r(a.r2, a.ar2);
  }
}

export interface AdjustContext {
  /** sheet that contains the formula (unprefixed refs resolve to it). null for workbook-level contexts w/o host */
  hostSheet: string | null;
  /** sheet on which the structural op is performed */
  targetSheet: string;
  /**
   * The formula is a defined name's refersTo. Relative parts of a name are offsets from the cell that uses the
   * name, so Excel does NOT adjust them (R2b: nm_rel =Data!C5 stays Data!C5 under every op); absolute parts are
   * adjusted normally. Leave unset for cell formulas, CF/DV formulas, chart series, etc.
   */
  isName?: boolean;
  /**
   * ListObjects of the workbook. Only used by DELETE ops: a structured reference that names a table column which
   * is deleted (or any reference to a table that is deleted entirely) becomes #REF!, exactly like Excel
   * (R2b t11/t12/t13/t20/t24). Structured refs never change on insert / move.
   */
  tables?: TableRef[];
  /**
   * Cell that holds the formula (on hostSheet). Only used to resolve UNQUALIFIED structured refs ("[@点数]",
   * UI form of a calculated column) to the table that contains the cell. File-form formulas always name the table.
   */
  hostCell?: { row: number; col: number };
}

/** Minimal table description for AdjustContext.tables (TableInfo + the sheet it lives on). */
export interface TableRef {
  name: string;
  sheet: string;
  /** A1 range incl. header (and totals) row, e.g. "C3:H10" */
  ref: string;
  /** column names, left to right */
  columns: string[];
}
export interface AdjustResult {
  formula: string;
  changed: boolean;
  /** true if some reference became #REF! */
  createdRefError: boolean;
  /**
   * true if the formula contains unknown tokens — caller should NOT trust the result. Only computed when the
   * formula can reference the target sheet (a formula on another sheet that never names it is returned unchanged).
   */
  unsafe: boolean;
  /** hints: INDIRECT/OFFSET/ADDRESS with text refs, VLOOKUP/HLOOKUP/INDEX/MATCH col indexes etc. */
  warnings: string[];
}

const sameSheet = (a: string | null, b: string | null) => a != null && b != null && a.toLowerCase() === b.toLowerCase();

/**
 * Adjust the area of a defined name: relative parts are offsets (never adjusted by Excel), absolute parts follow
 * the op. Returns the input object when nothing changes, null for #REF!.
 */
export function adjustNameArea(a: Area, op: StructOp): Area | null {
  const isCol = op.axis === 'col';
  if (isCol ? a.kind === 'rows' : a.kind === 'cols') return a;
  const abs1 = isCol ? a.ac1 : a.ar1, abs2 = isCol ? a.ac2 : a.ar2;
  if (!abs1 && !abs2) return a;
  const r = adjustArea(a, op);
  if (abs1 && abs2) return r;
  if (!r) return null;
  // mixed ($C5:E5 style): only the absolute endpoint follows the op
  const res: Area = { ...a };
  if (isCol) { if (abs1) res.c1 = r.c1; else res.c2 = r.c2; } else { if (abs1) res.r1 = r.r1; else res.r2 = r.r2; }
  return res.c1 === a.c1 && res.c2 === a.c2 && res.r1 === a.r1 && res.r2 === a.r2 ? a : normalizeFull(res);
}

/** Structured reference parts: table name (null for "[@col]" style) and every column name it mentions. */
export interface StructuredRefInfo {
  table: string | null;
  /** column names mentioned (endpoints of [a]:[b] ranges), unescaped */
  columns: string[];
  /** '#All' '#Data' '#Headers' '#Totals' '#This Row' (as written, without brackets) */
  specials: string[];
  /** true for [@...] / [#This Row] references */
  thisRow: boolean;
}

/**
 * Parse the text of an 'sref' token ("tbl[点数]", "tbl[[#This Row],[a b]]", "[@列]", "tbl[@[a]:[c]]", "tbl[]").
 * A sheet/book prefix (external refs) is skipped. Returns null when the text is not a structured reference.
 */
export function parseStructuredRef(text: string): StructuredRefInfo | null {
  const lb = findSrefOpen(text);
  if (lb < 0 || !text.endsWith(']')) return null;
  let table: string | null = text.slice(0, lb);
  const bang = table.lastIndexOf('!');
  if (bang >= 0) table = table.slice(bang + 1);
  if (!table) table = null;
  const inner = text.slice(lb + 1, -1);
  const info: StructuredRefInfo = { table, columns: [], specials: [], thisRow: false };
  const addItem = (raw: string) => {
    const r = raw;
    if (!r) return;
    // an escaped "'#" is part of a column name; only a bare leading '#' starts a special item
    if (r.startsWith('#')) { info.specials.push(r); if (/^#this\s+row$/i.test(r)) info.thisRow = true; }
    else info.columns.push(unescapeSref(r));
  };
  let body = inner;
  if (body.startsWith('@')) { info.thisRow = true; body = body.slice(1); }
  if (!body.trimStart().startsWith('[')) { addItem(body); return info; }
  // list of bracketed items separated by ',' / ':' / spaces
  for (let k = 0; k < body.length; k++) {
    if (body[k] !== '[') continue;
    let e = k + 1;
    for (; e < body.length; e++) { if (body[e] === "'") { e++; continue; } if (body[e] === ']') break; }
    addItem(body.slice(k + 1, e));
    k = e;
  }
  return info;
}
/** index of the '[' that opens the structured-reference body (after an optional prefix / table name), -1 if none */
function findSrefOpen(text: string): number {
  // external prefixes such as "[1]!" or "[Book.xlsx]Sheet!" start with '[' too: skip past the last '!' outside brackets
  let depth = 0, bang = -1;
  for (let k = 0; k < text.length; k++) {
    const ch = text[k];
    if (ch === "'") { k++; continue; }
    if (ch === '[') depth++; else if (ch === ']') depth--; else if (ch === '!' && depth === 0) bang = k;
  }
  return text.indexOf('[', bang + 1);
}
const unescapeSref = (s: string) => s.replace(/'(.)/gs, '$1');

/** Name of the table (ctx.tables, on ctx.hostSheet) whose range contains ctx.hostCell, or null. */
function tableAtHost(ctx: AdjustContext): string | null {
  const h = ctx.hostCell;
  if (!h || !ctx.tables) return null;
  for (const t of ctx.tables) {
    if (!sameSheet(t.sheet, ctx.hostSheet)) continue;
    const a = parseAreaText(t.ref.replace(/^.*!/, '').replace(/\$/g, ''));
    if (a && h.row >= Math.min(a.r1, a.r2) && h.row <= Math.max(a.r1, a.r2) && h.col >= Math.min(a.c1, a.c2) && h.col <= Math.max(a.c1, a.c2)) return t.name;
  }
  return null;
}

/** Table columns / whole table removed by a delete op (names lower-cased). null when the op does not touch the table. */
function deletedTableParts(t: TableRef, op: StructOp): { whole: boolean; cols: Set<string> } | null {
  if (op.type !== 'delete') return null;
  const a = parseAreaText(t.ref.replace(/^.*!/, ''));
  if (!a) return null;
  const s1 = op.at, s2 = op.at + op.count - 1;
  const lo = op.axis === 'col' ? Math.min(a.c1, a.c2) : Math.min(a.r1, a.r2);
  const hi = op.axis === 'col' ? Math.max(a.c1, a.c2) : Math.max(a.r1, a.r2);
  if (s2 < lo || s1 > hi) return null;
  const whole = s1 <= lo && s2 >= hi;
  const cols = new Set<string>();
  if (op.axis === 'col') for (let c = Math.max(lo, s1); c <= Math.min(hi, s2); c++) { const n = t.columns[c - lo]; if (n != null) cols.add(n.toLowerCase()); }
  return { whole, cols };
}

/**
 * fast path: a formula on another sheet can only reference the target sheet through its name (3D / external refs
 * are never adjusted); structured refs only matter for deletes with tables
 */
function cannotReferenceTarget(formula: string, ops: readonly StructOp[], ctx: AdjustContext): boolean {
  return !sameSheet(ctx.hostSheet, ctx.targetSheet) && !(ctx.tables?.length && ops.some((op) => op.type === 'delete')) &&
    !ctx.targetSheet.includes("'") && !formula.toLowerCase().includes(ctx.targetSheet.toLowerCase());
}

/** One op over a token list (tokens replaced in place: a broken reference becomes an 'error' token "Sheet!#REF!"). */
function adjustTokens(toks: Token[], op: StructOp, ctx: AdjustContext, res: AdjustResult): void {
  // tables on the target sheet touched by a delete (structured refs to deleted columns become #REF!)
  let hitTables: Map<string, { whole: boolean; cols: Set<string> }> | null = null;
  if (op.type === 'delete' && ctx.tables) {
    for (const t of ctx.tables) {
      if (!sameSheet(t.sheet, ctx.targetSheet)) continue;
      const d = deletedTableParts(t, op);
      if (d) (hitTables ??= new Map()).set(t.name.toLowerCase(), d);
    }
  }
  for (let i = 0; i < toks.length; i++) {
    const t = toks[i];
    if (t.type === 'unknown') res.unsafe = true;
    if (t.type === 'func') {
      const fn = t.text.replace(/^(_xlfn\.|_xlws\.)+/i, '').toUpperCase();
      const add = (w: string) => { if (!res.warnings.includes(w)) res.warnings.push(w); };
      if (fn === 'INDIRECT' || fn === 'ADDRESS') add(`${fn}: text references are not adjusted (same as Excel)`);
      if (fn === 'VLOOKUP' || fn === 'HLOOKUP') add(`${fn}: column/row index argument is a number and is not adjusted (same as Excel)`);
    }
    if (hitTables && (t.type === 'sref' || t.type === 'name') && !t.prefix?.book) {
      let broken = false;
      if (t.type === 'name') broken = !t.prefix && !!hitTables.get(t.text.toLowerCase())?.whole; // UI form: ROWS(tblMain)
      else {
        const s = parseStructuredRef(t.prefix ? t.text.slice(t.prefix.text.length) : t.text);
        const tname = s?.table ?? (s && !t.prefix ? tableAtHost(ctx) : null);
        const d = tname ? hitTables.get(tname.toLowerCase()) : undefined;
        broken = !!d && (d.whole || s!.columns.some((c) => d.cols.has(c.toLowerCase())));
      }
      if (broken) { toks[i] = { type: 'error', text: '#REF!', start: t.start, end: t.end }; res.changed = true; res.createdRefError = true; continue; }
    }
    if (t.type !== 'ref') continue;
    const p = t.prefix;
    if (p && (p.book || p.sheet2 !== undefined)) continue; // external / 3D refs: untouched (Excel: 3D not adjusted by single-sheet ops)
    const sheet = p ? p.sheet : ctx.hostSheet;
    if (!sameSheet(sheet, ctx.targetSheet)) continue;
    const na = ctx.isName ? adjustNameArea(t.area!, op) : adjustArea(t.area!, op);
    if (na === t.area) continue; // untouched: keep exactly as written (case, etc.)
    const prefixText = p ? p.text : ''; // prefix is copied verbatim (quoting/case preserved)
    if (!na) { // Excel keeps the prefix (Data!#REF!) and the trim marker of a UI-form trim range (Data!#REF!:.)
      const trim = t.area!.trim;
      toks[i] = { type: 'error', text: prefixText + '#REF!' + trimSuffix(trim), start: t.start, end: t.end, prefix: p, ...(trim ? { trim } : {}) };
      res.createdRefError = true; res.changed = true;
      continue;
    }
    const text = prefixText + renderArea(na);
    if (text !== t.text) res.changed = true;
    toks[i] = { ...t, text, area: na };
  }
}

/** Rewrite every reference in `formula` affected by `op` on `ctx.targetSheet`. */
export function adjustFormula(formula: string, op: StructOp, ctx: AdjustContext): AdjustResult {
  const res: AdjustResult = { formula, changed: false, createdRefError: false, unsafe: false, warnings: [] };
  if (cannotReferenceTarget(formula, [op], ctx)) return res;
  const toks = tokenize(formula);
  adjustTokens(toks, op, ctx, res);
  if (res.changed) res.formula = toks.map((t) => t.text).join('');
  return res;
}

/**
 * Apply a sequence of ops on the same sheet (e.g. a multi-selection move decomposed by planMove, or several deletes).
 * The formula is tokenized once. Between two ops the context follows the previous op (advanceContext): table ranges
 * / columns and the host cell move, so structured refs of a later delete resolve against the table as it is then.
 */
export function adjustFormulaSeq(formula: string, ops: StructOp[], ctx: AdjustContext): AdjustResult {
  const res: AdjustResult = { formula, changed: false, createdRefError: false, unsafe: false, warnings: [] };
  if (!ops.length || cannotReferenceTarget(formula, ops, ctx)) return res;
  const toks = tokenize(formula);
  let c = ctx;
  for (let k = 0; k < ops.length; k++) {
    adjustTokens(toks, ops[k], c, res);
    if (k + 1 < ops.length) c = advanceContext(c, ops[k]);
  }
  if (res.changed) res.formula = toks.map((t) => t.text).join('');
  return res;
}

/**
 * The AdjustContext for the NEXT op of a sequence: tables on the target sheet take their range / column list after
 * `op` (moved columns reordered, deleted ones removed, inserted ones get placeholder names that no reference can
 * match, a deleted table disappears) and the host cell follows `op` when it is on the target sheet (undefined when
 * deleted). Returns `ctx` itself when nothing changes.
 */
export function advanceContext(ctx: AdjustContext, op: StructOp): AdjustContext {
  let tables = ctx.tables;
  if (tables?.some((t) => sameSheet(t.sheet, ctx.targetSheet))) {
    tables = tables.flatMap((t) => {
      if (!sameSheet(t.sheet, ctx.targetSheet)) return [t];
      const ref = adjustTableRef(t.ref, op);
      if (ref == null) return [];
      if (op.axis === 'row') return ref === t.ref ? [t] : [{ ...t, ref }];
      const a = parseAreaText(t.ref.replace(/^.*!/, '').replace(/\$/g, ''));
      const na = parseAreaText(ref.replace(/^.*!/, '').replace(/\$/g, ''));
      if (!a || !na) return [{ ...t, ref }];
      const lo = Math.min(a.c1, a.c2), nlo = Math.min(na.c1, na.c2), width = Math.max(na.c1, na.c2) - nlo + 1;
      const cols: string[] = Array.from({ length: width }, (_, k) => `\u0000new${k}`);
      t.columns.forEach((name, k) => {
        const nc = mapIndex(lo + k, op);
        if (nc != null && nc >= nlo && nc < nlo + width) cols[nc - nlo] = name;
      });
      return [{ ...t, ref, columns: cols }];
    });
  }
  let hostCell = ctx.hostCell;
  if (hostCell && sameSheet(ctx.hostSheet, ctx.targetSheet)) {
    const v = mapIndex(op.axis === 'col' ? hostCell.col : hostCell.row, op);
    hostCell = v == null ? undefined : op.axis === 'col' ? { ...hostCell, col: v } : { ...hostCell, row: v };
  }
  return tables === ctx.tables && hostCell === ctx.hostCell ? ctx : { ...ctx, tables, hostCell };
}

/** Adjust an sqref-like list of plain areas ("B1:E10 G2" in xlsx; "B1:E10,G2" in COM). Used for CF / DV / merges. */
export function adjustSqref(sqref: string, op: StructOp, sep = ' '): string {
  return sqref.split(sep).map((part) => {
    const a = parseAreaText(part);
    if (!a) return part;
    const na = adjustArea(a, op);
    return na ? renderArea(na) : null;
  }).filter((x) => x != null).join(sep);
}

export type TableMoveKind = 'reorder' | 'whole' | 'shift' | 'refused';

/**
 * How Excel treats moving the block [from, from + count - 1] to before `to` (whole columns, native Insert Cut Cells)
 * when a table spans [lo, hi] on the same axis. Verified in Excel 365 (R2b t01 .. t22 and
 * tests/fixtures/formula/com/excel-table-move.ps1 -> tests/formula/tableMove.test.ts):
 *  - 'whole'    the block covers the whole table: the table moves with it (range like any range);
 *  - 'reorder'  the block lies inside the table (not all of it) and the destination is within [lo, hi + 1]: the table
 *               columns are reordered, the table range is unchanged. `to = hi + 1` (right after the last table column)
 *               IS a reorder: the block becomes the last columns of the table and does not leave it (Excel accepts it,
 *               probe in_to_I / in2_to_I / first_to_I);
 *  - 'shift'    the block is entirely outside the table and does not land strictly inside it (to <= lo or to > hi):
 *               the table is only shifted like a range (a block landing right after the table does not extend it,
 *               probe out_to_I);
 *  - 'refused'  Excel refuses: a block inside the table going further than `hi + 1` or before `lo` (t03), a block from
 *               outside landing strictly inside the table (to in (lo, hi]: t06, t07), a block partly over the table
 *               (t21). The doc / patcher refuse these (ARCHITECTURE decision 5).
 * A move that changes nothing (to within [from, from + count]) is reported as 'shift' (nothing happens).
 */
export function tableMoveKind(lo: number, hi: number, op: Extract<StructOp, { type: 'move' }>): TableMoveKind {
  const s = op.from, e = op.from + op.count - 1, to = op.to;
  if (to >= s && to <= e + 1) return 'shift';
  if (s <= lo && e >= hi) return 'whole';
  if (s >= lo && e <= hi) return to >= lo && to <= hi + 1 ? 'reorder' : 'refused';
  if (e < lo || s > hi) return to > lo && to <= hi ? 'refused' : 'shift';
  return 'refused';
}

/**
 * New range of a ListObject (table) after a structural op on its sheet; null when the table is deleted entirely.
 * Unlike a formula range, a table CONTAINS its columns: moving columns inside the table (tableMoveKind 'reorder':
 * block within the table, destination between its first column and the column right after its last) only reorders
 * the table columns and keeps the range (R2b t01, t02, t19, t22; Excel probe in_to_I). Moving the whole table or
 * columns elsewhere shifts it like a range. Inserting inside the table grows it; inserting right after it does not.
 * Excel refuses moves into / out of / partly over a table — callers must refuse those first (tableMoveKind
 * 'refused', ARCHITECTURE decision 5); for such ops the result here is only the range rule.
 */
export function adjustTableRef(ref: string, op: StructOp): string | null {
  const bang = ref.lastIndexOf('!');
  const prefix = bang >= 0 ? ref.slice(0, bang + 1) : '';
  const a = parseAreaText(ref.slice(prefix.length));
  if (!a) return ref;
  if (op.type === 'move') {
    const [lo, hi] = op.axis === 'col' ? [Math.min(a.c1, a.c2), Math.max(a.c1, a.c2)] : [Math.min(a.r1, a.r2), Math.max(a.r1, a.r2)];
    if (tableMoveKind(lo, hi, op) === 'reorder') return ref;
  }
  const r = adjustArea(a, op);
  return r ? (r === a ? ref : prefix + renderArea(r)) : null;
}

/**
 * Relative translation (copy / fill / shared-formula expansion): shifts relative parts by (dr, dc).
 * Absolute parts are kept. Out-of-sheet results become #REF! (Excel behaviour when copying).
 */
export function translateFormula(formula: string, dr: number, dc: number): string {
  if (dr === 0 && dc === 0) return formula;
  return tokenize(formula).map((t) => {
    if (t.type !== 'ref' || !t.area) return t.text;
    const a = { ...t.area };
    const mv = (v: number, abs: boolean, d: number) => (abs || v === 0 ? v : v + d);
    a.c1 = mv(a.c1, a.ac1, dc); a.c2 = mv(a.c2, a.ac2, dc); a.r1 = mv(a.r1, a.ar1, dr); a.r2 = mv(a.r2, a.ar2, dr);
    const bad = (a.kind !== 'rows' && (a.c1 < 1 || a.c2 < 1 || a.c1 > MAX_COL || a.c2 > MAX_COL)) ||
                (a.kind !== 'cols' && (a.r1 < 1 || a.r2 < 1 || a.r1 > MAX_ROW || a.r2 > MAX_ROW));
    return (t.prefix ? t.prefix.text : '') + (bad ? '#REF!' : renderArea(a));
  }).join('');
}

/**
 * Quote a sheet name for use in a formula prefix. Conservative: quotes whenever Excel might (observed in
 * Excel 365: spaces, '-', full-width punctuation, leading digit, names that look like A1/R1C1 refs such as
 * 'A1','TAX2025','R','C','RC','R1C1'). Quoting a name that did not need it is always accepted by Excel.
 */
export function quoteSheetName(name: string): string {
  const simple = /^[A-Za-z_\\\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}][A-Za-z0-9_.\\\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}ー]*$/u.test(name);
  const cellLike = /^\$?[A-Za-z]{1,3}\$?[0-9]+$/.test(name) || /^(R|C|RC|R-?\d*C-?\d*|R\d+|C\d+)$/i.test(name);
  const boolLike = /^(TRUE|FALSE)$/i.test(name);
  return simple && !cellLike && !boolLike && !name.includes('.') ? name : `'${name.replace(/'/g, "''")}'`;
}

/** Rename a sheet inside every prefix of a formula (keeps everything else verbatim; 3D ends handled). */
export function renameSheetInFormula(formula: string, oldName: string, newName: string): string {
  const eq = (a?: string | null) => a != null && a.toLowerCase() === oldName.toLowerCase();
  return tokenize(formula).map((t) => {
    const p = t.prefix;
    if (!p || p.book) return t.text;
    const body = t.text.slice(p.text.length);
    const s1 = (p as any).sheet1 as string | undefined;
    if (p.sheet2 !== undefined && (eq(s1) || eq(p.sheet2))) {
      const a = eq(s1) ? newName : s1!, b = eq(p.sheet2) ? newName : p.sheet2!;
      return `'${(a + ':' + b).replace(/'/g, "''")}'!` + body;
    }
    if (eq(p.sheet)) return quoteSheetName(newName) + '!' + body;
    return t.text;
  }).join('');
}

/** Collect the references of a formula (for dependency checks / warnings before deleting a column). */
export function listRefs(formula: string, hostSheet: string | null) {
  return tokenize(formula).filter((t) => t.type === 'ref').map((t) => ({
    text: t.text,
    sheet: t.prefix ? t.prefix.sheet : hostSheet,
    external: !!t.prefix?.book,
    threeD: t.prefix?.sheet2 !== undefined,
    area: t.area!,
  }));
}

// ---------------------------------------------------------------------------------------------
// dependents / lookup warnings (helpers for the doc & UI)
// ---------------------------------------------------------------------------------------------
/** A reference of a formula that touches some columns (refsTouchingColumns). */
export interface TouchingRef {
  /** token text as written ("Data!$C$2:$E$9", "tbl[点数]") */
  text: string;
  kind: 'ref' | 'sref';
  /** the A1 area for 'ref' tokens */
  area?: Area;
  /** touched columns of `cols` (sorted) */
  cols: number[];
}

/**
 * The references of `formula` (hosted on `hostSheet`) that point at any of the columns `cols` of `targetSheet`:
 * A1 refs whose column span contains one of them (whole-row refs contain every column) and, when `tables` is
 * given, structured references to columns of a table on targetSheet located in those columns ([#This Row],
 * [a]:[c] ranges and whole-table refs included). External and 3D refs are ignored (single-sheet ops never adjust
 * them). Empty array = the formula does not depend on those columns (defined names are not expanded — check the
 * names' own refersTo separately).
 */
export function refsTouchingColumns(
  formula: string,
  hostSheet: string | null,
  targetSheet: string,
  cols: Iterable<number>,
  opts: { tables?: TableRef[]; hostCell?: { row: number; col: number } } = {},
): TouchingRef[] {
  const list = [...new Set(cols)].sort((a, b) => a - b);
  if (!list.length) return [];
  const within = (lo: number, hi: number) => list.filter((c) => c >= lo && c <= hi);
  const out: TouchingRef[] = [];
  for (const t of tokenize(formula)) {
    if (t.type === 'ref') {
      const p = t.prefix;
      if (p && (p.book !== undefined || p.sheet2 !== undefined)) continue;
      if (!sameSheet(p ? p.sheet : hostSheet, targetSheet)) continue;
      const a = t.area!;
      const hit = a.kind === 'rows' ? list : within(Math.min(a.c1, a.c2), Math.max(a.c1, a.c2));
      if (hit.length) out.push({ text: t.text, kind: 'ref', area: a, cols: hit });
    } else if (t.type === 'sref' && opts.tables && !t.prefix?.book) {
      const s = parseStructuredRef(t.prefix ? t.text.slice(t.prefix.text.length) : t.text);
      if (!s) continue;
      const tname = s.table ?? (!t.prefix ? tableAtHost({ hostSheet, targetSheet, tables: opts.tables, hostCell: opts.hostCell }) : null);
      const tb = tname ? opts.tables.find((x) => x.name.toLowerCase() === tname.toLowerCase()) : undefined;
      if (!tb || !sameSheet(tb.sheet, targetSheet)) continue;
      const a = parseAreaText(tb.ref.replace(/^.*!/, '').replace(/\$/g, ''));
      if (!a) continue;
      const left = Math.min(a.c1, a.c2);
      const idx = s.columns.map((c) => tb.columns.findIndex((x) => x.toLowerCase() === c.toLowerCase())).filter((i) => i >= 0);
      let lo: number, hi: number;
      if (!s.columns.length) { lo = left; hi = Math.max(a.c1, a.c2); } // whole table / [#Data] / [#This Row]
      else if (!idx.length) continue;
      else { lo = left + Math.min(...idx); hi = left + Math.max(...idx); }
      const hit = within(lo, hi);
      if (hit.length) out.push({ text: t.text, kind: 'sref', cols: hit });
    }
  }
  return out;
}

/** A lookup-style function whose numeric position argument is not adjusted by Excel (lookupWarnings). */
export interface LookupWarning {
  /** function name (VLOOKUP, INDEX, ...) */
  fn: string;
  /** the range argument as written */
  ref: string;
  /** Japanese message for the UI */
  message: string;
}

/** argument positions (0-based) holding the ranges whose inner positions matter, per function */
const LOOKUP_RANGE_ARGS: Record<string, number[] | 'all'> = {
  VLOOKUP: [1], HLOOKUP: [1], INDEX: [0], MATCH: [1], XMATCH: [1], XLOOKUP: [1, 2], OFFSET: [0], CHOOSE: 'all',
};

/** Extra inputs of lookupWarnings. */
export interface LookupWarningOptions {
  /**
   * Resolves a defined name used as a range argument ("VLOOKUP(x,社員マスタ,3,0)") to its refersTo (FILE form, no
   * '='), or null when unknown. Without it, defined names are not checked.
   */
  resolveName?: (name: string) => string | null;
}

/** A range argument of a lookup function, reduced to its span on the op axis. */
interface LookupRange {
  text: string;
  lo: number;
  hi: number;
  /** span after the op (null: the range disappears -> #REF!, reported by the delete check instead) */
  after: Span;
  /** 1-D along the op axis (e.g. a single-row range for a column op): INDEX(r,n) counts along it */
  oneOther: boolean;
}

/**
 * Warnings for VLOOKUP / HLOOKUP / INDEX / OFFSET / MATCH / XMATCH / CHOOSE / XLOOKUP whose range argument is
 * affected by `op` inside (a column/row moved, inserted or deleted within the range): Excel adjusts the range but
 * NOT numeric position arguments, so the result can silently come from another column. With a literal index
 * (VLOOKUP(x,A:E,3,…), INDEX(r,n,k), OFFSET(r,0,2)) the check is exact: it warns only when the column the index
 * pointed at is no longer at that index. Otherwise it warns whenever the positions inside the range change.
 * Range arguments may be A1 refs, tables / structured refs (ctx.tables: "VLOOKUP(x,Table1,3,0)" warns when columns
 * are reordered INSIDE the table, which keeps the table range) and defined names (opts.resolveName).
 * Same sheet resolution as adjustFormula (ctx.hostSheet / ctx.targetSheet).
 */
export function lookupWarnings(formula: string, op: StructOp, ctx: AdjustContext, opts: LookupWarningOptions = {}): LookupWarning[] {
  const toks = tokenize(formula);
  if (!toks.some((t) => t.type === 'func')) return [];
  const calls = parseCalls(toks);
  const isCol = op.axis === 'col';
  const max = isCol ? MAX_COL : MAX_ROW;
  const axisWord = isCol ? '列' : '行';
  const out: LookupWarning[] = [];
  const seen = new Set<string>();
  const literal = (arg: [number, number] | undefined): number | null => {
    if (!arg) return null;
    const [s, e] = trimRange(toks, arg);
    return e - s === 1 && toks[s].type === 'number' ? Number(toks[s].text) : null;
  };
  /** does the op change what sits at 1-based position k of [lo, hi] (now at `after`)? (k = null: any position) */
  const changes = (lo: number, hi: number, after: Span, k: number | null): boolean => {
    if (!after) return false; // the whole range disappears: #REF!, reported by the delete check instead
    if (k != null) {
      if (!Number.isInteger(k) || k < 1) return false;
      const target = lo + k - 1;
      if (target > max) return false;
      const moved = mapIndex(target, op);
      return moved == null || moved !== after[0] + k - 1;
    }
    if (op.type === 'insert') return op.at > lo && op.at <= hi;
    if (op.type === 'delete') { const s2 = op.at + op.count - 1; return op.at <= hi && s2 >= lo; }
    const s1 = op.from, s2 = op.from + op.count - 1;
    if (op.to >= s1 && op.to <= s2 + 1) return false;
    if (lo >= s1 && hi <= s2) return false; // the whole range moves with the block
    return (s1 <= hi && s2 >= lo) || (op.to > lo && op.to <= hi);
  };

  /** A1 area on the target sheet -> LookupRange (null when the op cannot touch it) */
  const fromArea = (a: Area, text: string): LookupRange | null => {
    // whole rows for a column op (whole columns for a row op): the range spans the whole axis and never changes,
    // but the order of the columns inside it does (MATCH(x,1:1,0), VLOOKUP(x,1:5,3,0))
    if (isCol ? a.kind === 'rows' : a.kind === 'cols') return { text, lo: 1, hi: max, after: [1, max], oneOther: Math.min(isCol ? a.r1 : a.c1, isCol ? a.r2 : a.c2) === Math.max(isCol ? a.r1 : a.c1, isCol ? a.r2 : a.c2) };
    const lo = isCol ? Math.min(a.c1, a.c2) : Math.min(a.r1, a.r2);
    const hi = isCol ? Math.max(a.c1, a.c2) : Math.max(a.r1, a.r2);
    const oLo = isCol ? Math.min(a.r1, a.r2) : Math.min(a.c1, a.c2), oHi = isCol ? Math.max(a.r1, a.r2) : Math.max(a.c1, a.c2);
    const oneOther = a.kind !== (isCol ? 'cols' : 'rows') && oLo === oHi;
    return { text, lo, hi, after: applySpan(lo, hi, op, max, a.kind === 'cell'), oneOther };
  };
  const onTarget = (p: Prefix | undefined) =>
    !(p && (p.book !== undefined || p.sheet2 !== undefined)) && sameSheet(p ? p.sheet : ctx.hostSheet, ctx.targetSheet);
  const tableOf = (name: string | null): TableRef | undefined =>
    name ? ctx.tables?.find((x) => x.name.toLowerCase() === name.toLowerCase() && sameSheet(x.sheet, ctx.targetSheet)) : undefined;
  /** a whole table / a structured reference -> LookupRange (column ops only: tables are laid out in columns) */
  const fromTable = (tb: TableRef, columns: string[], text: string): LookupRange | null => {
    if (!isCol) return null;
    const ref = tb.ref.replace(/^.*!/, '').replace(/\$/g, '');
    const a = parseAreaText(ref);
    if (!a) return null;
    const left = Math.min(a.c1, a.c2), right = Math.max(a.c1, a.c2);
    if (!columns.length) {
      // whole table: moves inside the table keep its range (adjustTableRef) but reorder its columns
      const nr = adjustTableRef(ref, op);
      const na = nr ? parseAreaText(nr) : null;
      return { text, lo: left, hi: right, after: na ? [Math.min(na.c1, na.c2), Math.max(na.c1, na.c2)] : null, oneOther: false };
    }
    const idx = columns.map((c) => tb.columns.findIndex((x) => x.toLowerCase() === c.toLowerCase()));
    if (idx.some((i) => i < 0)) return null;
    const lo = left + Math.min(...idx), hi = left + Math.max(...idx);
    // [a]:[c] follows its end COLUMNS (by name), wherever they move
    const e1 = mapIndex(lo, op), e2 = mapIndex(hi, op);
    return { text, lo, hi, after: e1 == null || e2 == null ? null : [Math.min(e1, e2), Math.max(e1, e2)], oneOther: false };
  };
  /** the range arguments found in tokens [s, e) */
  const rangesIn = (s: number, e: number): LookupRange[] => {
    const res: LookupRange[] = [];
    for (let i = s; i < e; i++) {
      const t = toks[i];
      if (t.type === 'ref' && t.area) {
        if (!onTarget(t.prefix)) continue;
        const r = fromArea(t.area, t.text);
        if (r) res.push(r);
      } else if (t.type === 'sref' && ctx.tables && !t.prefix?.book) {
        const sr = parseStructuredRef(t.prefix ? t.text.slice(t.prefix.text.length) : t.text);
        if (!sr) continue;
        const tb = tableOf(sr.table ?? (!t.prefix ? tableAtHost(ctx) : null));
        const r = tb ? fromTable(tb, sr.columns, t.text) : null;
        if (r) res.push(r);
      } else if (t.type === 'name' && !t.prefix?.book) {
        const tb = !t.prefix ? tableOf(t.text) : undefined; // UI form: VLOOKUP(x,Table1,3,0)
        if (tb) { const r = fromTable(tb, [], t.text); if (r) res.push(r); continue; }
        const refersTo = opts.resolveName?.(t.text.slice(t.prefix?.text.length ?? 0));
        if (!refersTo) continue;
        for (const nt of tokenize(refersTo.replace(/^=/, ''))) {
          if (nt.type !== 'ref' || !nt.area || !onTarget(nt.prefix)) continue;
          const r = fromArea(nt.area, t.text);
          if (r) res.push(r);
        }
      }
    }
    return res;
  };

  for (const c of calls.values()) {
    const spec = LOOKUP_RANGE_ARGS[c.name];
    if (!spec) continue;
    const argIdx = spec === 'all' ? c.args.map((_, i) => i).slice(1) : spec.filter((i) => i < c.args.length);
    for (const ai of argIdx) {
      for (const g of rangesIn(c.args[ai][0], c.args[ai][1])) {
        const { lo, hi, after } = g;
        let k: number | null = null, exact = false, what = '位置';
        switch (c.name) {
          case 'VLOOKUP': if (!isCol) continue; k = literal(c.args[2]); exact = k != null; what = '列番号'; break;
          case 'HLOOKUP': if (isCol) continue; k = literal(c.args[2]); exact = k != null; what = '行番号'; break;
          case 'INDEX': {
            // INDEX(range, n) on a 1-D range: n counts along that range; otherwise row_num / column_num
            const one = lo !== hi && g.oneOther;
            const argK = one ? c.args[1] : isCol ? c.args[2] : c.args[1];
            if (!argK) { if (!one) continue; }
            k = literal(argK); exact = k != null; what = isCol ? '列番号' : '行番号';
            if (k === 0) continue; // whole row / column of the range
            break;
          }
          case 'OFFSET': {
            const d = literal(isCol ? c.args[2] : c.args[1]);
            if (d === 0 || (d == null && !(isCol ? c.args[2] : c.args[1]))) continue;
            what = isCol ? '列のずれ' : '行のずれ';
            if (d != null) { k = d + 1; exact = true; if (d < 0) continue; } // target relative to the base's first position
            else if (!changes(lo, max, applySpan(lo, max, op, max, false), null)) continue;
            break;
          }
          default: if (lo === hi) continue;
        }
        if (c.name !== 'OFFSET' && !exact && lo === hi) continue;
        let hit: boolean;
        if (c.name === 'OFFSET' && k != null) {
          // OFFSET returns (new position of the base) + d; the cell it pointed at may have moved elsewhere
          hit = !!after && changes(lo, lo + k - 1, [after[0], after[0]], k);
        } else if (c.name === 'OFFSET') hit = true; // computed offset: any change right of / below the base (checked above)
        else hit = changes(lo, hi, after, exact ? k : null);
        // VLOOKUP searches the FIRST column of its range: swapping the key column with another column of the range (A <-> B in
        // VLOOKUP(x,A:D,4,0)) leaves the returned column alone but makes every lookup search the wrong column (#N/A) — found by the
        // consistency oracle (tests/oracle/consistency), Excel does not adjust it either
        const searchColumn = !hit && c.name === 'VLOOKUP' && exact && changes(lo, hi, after, 1);
        if (!hit && !searchColumn) continue;
        const key = `${c.name}|${g.text}`;
        if (seen.has(key)) continue;
        seen.add(key);
        const message = searchColumn
          ? `${c.name} は範囲 ${g.text} の先頭の${axisWord}（検索する${axisWord}）で探します。その${axisWord}が移動・挿入・削除で別の${axisWord}になるため、見つからない（#N/A）か別の行の値を返す可能性があります（Excel でも自動では直りません）。`
          : ['MATCH', 'XMATCH', 'XLOOKUP', 'CHOOSE'].includes(c.name)
            ? `${c.name}: 範囲 ${g.text} の中で${axisWord}が移動・挿入・削除されます。結果の位置や形が変わる可能性があります。`
            : `${c.name} の${what}は数値のため自動では変わりません。範囲 ${g.text} の${axisWord}が移動・挿入・削除されるため、別の${axisWord}の値を返す可能性があります。`;
        out.push({ fn: c.name, ref: g.text, message });
      }
    }
  }
  return out;
}

// ---------------------------------------------------------------------------------------------
// Anchored rules: conditional formatting / data validation
// ---------------------------------------------------------------------------------------------
/**
 * Relative key of a formula hosted at (row, col): relative parts become offsets (R1C1-like),
 * absolute parts stay absolute. Two cells with equal keys can share one CF/DV rule.
 */
export function relKey(formula: string, row: number, col: number): string {
  return tokenize(formula).map((t) => {
    if (t.type !== 'ref' || !t.area) return t.text;
    const a = t.area;
    const c = (v: number, abs: boolean) => (v === 0 ? '' : abs ? `C${v}` : `C[${v - col}]`);
    const r = (v: number, abs: boolean) => (v === 0 ? '' : abs ? `R${v}` : `R[${v - row}]`);
    return (t.prefix ? t.prefix.text : '') + `{${a.kind}${a.trim ? '.' + a.trim : ''}:${r(a.r1, a.ar1)}${c(a.c1, a.ac1)}:${r(a.r2, a.ar2)}${c(a.c2, a.ac2)}}`;
  }).join('');
}

export interface AnchoredRule {
  /** areas of the sqref (1-based cell areas; whole columns are written C1:C1048576 in xlsx sqref) */
  sqref: Area[];
  /**
   * formulas relative to the top-left cell of the sqref BOUNDING BOX (min row, min col over all areas) — this is
   * how xlsx stores CF/DV formulas (CF: formula1/formula2, DV: formula1/formula2). Results use the same convention.
   */
  formulas: string[];
}

/**
 * OPTION SHAPE of adjustAnchoredRule / adjustAnchoredRulePerCell (round 2; callers: doc/struct.ts, patch/rules.ts).
 * Pass the rule's own kind and type — never pre-decide "leaks or not" yourself:
 *
 *   { kind?: 'cf' | 'dv', type?: string, inheritFormat?: boolean, siblings?: string[][] }
 *
 *   adjustAnchoredRule(r, op, sheet, { kind: fr.kind, type: fr.type })          // FeatureRule (doc)
 *   adjustAnchoredRule(r, op, sheet, { kind: 'cf', type: 'colorScale' })        // xlsx <cfRule type>
 *   adjustAnchoredRule(r, op, sheet, { kind: 'dv', type: 'list' })              // xlsx <dataValidation type>
 *   adjustAnchoredRule(r, op, sheet, { kind, type, inheritFormat: false })      // premerge "clean insert" of columns
 *   adjustAnchoredRule(r, op, sheet, { kind: 'cf', type, siblings })            // rule with other rules in its block
 *
 * What the options change (measured on rules of every type in Excel 365: tests/formula/cfSpread.test.ts,
 * cfProbe.test.ts; the patcher's goldens g-ins-D, g-ins-row14, t-mov-L-H, p-cf-mov):
 *  - insert right after an area (the area ends at X-1): the inserted cells copy the format of X-1 and join its rule,
 *    for EVERY CF type and for every data validation (inheritFormat: false turns this off). Inserting strictly inside
 *    an area always grows it (range semantics), whatever the options.
 *  - move ("spread"): the moved block lands at X; when X-1 belongs to an area of the rule the block joins that rule —
 *    unless inheritFormat is false, or the rule is a CF of a type in NO_SPREAD_CF_TYPES (colorScale, dataBar,
 *    iconSet, top10, uniqueValues, duplicateValues, aboveAverage), or a rule of the block (siblings) has, in the
 *    range object that covers the destination, a formula referring to the moved cells or to the destination (a cell
 *    referring to itself does not count, nor a run inside the block whose relative references stay inside it), or
 *    the range object itself contains the moved block (it is cut). cellIs, expression, containsText, containsBlanks,
 *    timePeriod ... and EVERY data validation type spread; with inheritFormat: false nothing spreads.
 *  - kind: 'cf' (default) or 'dv'. Both behave the same; the type list above applies to CF only.
 *  - type: FeatureRule.type (xlsx <cfRule type> / <dataValidation type>). Unknown / missing = spreads.
 *  - siblings: see the field. Omit it for a rule that is alone in its <conditionalFormatting> block.
 *
 * The clipboard-free sequence measured here is what the Excel runner executes for moves outside tables and what
 * replayLog implements (same log, same semantics). Excel's native Insert Cut Cells (used by the runner inside tables)
 * differs in one point: a cellIs / cellValue rule does NOT spread there (R2b); the round-1 option `cfType: 'cellValue'`
 * is kept for that comparison (tests/formula/r2b.test.ts).
 */
export interface AnchoredRuleOptions {
  /** 'cf' (default) conditional formatting, 'dv' data validation */
  kind?: 'cf' | 'dv';
  /** FeatureRule.type: the xlsx <cfRule type> ('cellIs', 'expression', 'colorScale', ...) / <dataValidation type> */
  type?: string;
  /** @deprecated round-1 name; used only when `type` is absent. 'cellValue' = does not spread onto a moved block. */
  cfType?: string;
  /**
   * Format inheritance (default true for CF and DV): growth over cells inserted right after an area and the spread
   * onto a moved block. Pass false for inserts that must not copy formats — the premerge's "clean insert" of
   * imported columns (ARCHITECTURE decision 1).
   */
  inheritFormat?: boolean;
  /**
   * Formulas (slots, same convention as AnchoredRule.formulas) of the OTHER rules that share this rule's sqref, i.e.
   * the other <cfRule>s of the same <conditionalFormatting> block. Excel keeps ONE range object per block: every rule
   * of a block is fragmented at the same places, and on a move the destination cells that joined the block by format
   * inheritance leave it again when ANY rule of the block refers to the moved block or to the destination (measured:
   * cfProbe.test.ts p5, the R2b / R2 blocks of 7-8 rules). Only matters for moves and splits; omit it (or pass []) for
   * a rule that is alone in its block. The result contains only this rule's formulas.
   */
  siblings?: string[][];
  /** @internal adjustAnchoredRule with siblings: the formulas of every rule of the block, own first */
  blockRules?: string[][];
}

/** CF rule types (xlsx <cfRule type>) Excel does not spread onto a moved block (statistics over the whole range). */
export const NO_SPREAD_CF_TYPES: ReadonlySet<string> = new Set([
  'colorScale', 'dataBar', 'iconSet', 'top10', 'uniqueValues', 'duplicateValues', 'aboveAverage',
]);

/** Cells inserted right after an area join its rule (format inheritance). */
export function ruleGrowsOnInsert(opts: AnchoredRuleOptions = {}): boolean {
  return opts.inheritFormat ?? true;
}

/**
 * The rule may spread onto a moved block that lands inside / right after one of its areas (the formula check is
 * done by adjustAnchoredRule). False for inheritFormat: false and for the CF types in NO_SPREAD_CF_TYPES.
 */
export function ruleSpreadsOnMove(opts: AnchoredRuleOptions = {}): boolean {
  if (opts.inheritFormat === false) return false;
  const isCf = (opts.kind ?? 'cf') === 'cf';
  if (opts.type != null) return !isCf || !NO_SPREAD_CF_TYPES.has(opts.type);
  if (opts.cfType === 'cellValue') return false; // round-1 switch
  return !isCf || !(opts.cfType != null && NO_SPREAD_CF_TYPES.has(opts.cfType));
}

/** Map a single index through an op (null = deleted). */
export function mapIndex(v: number, op: StructOp): number | null {
  const r = applySpan(v, v, op, op.axis === 'col' ? MAX_COL : MAX_ROW, true);
  return r ? r[0] : null;
}

/** [lo, hi] of an area on the op axis ('cols' areas span all rows, 'rows' areas all columns). */
function axisSpan(a: Area, isCol: boolean): [number, number] {
  if (isCol) return a.kind === 'rows' ? [1, MAX_COL] : [Math.min(a.c1, a.c2), Math.max(a.c1, a.c2)];
  return a.kind === 'cols' ? [1, MAX_ROW] : [Math.min(a.r1, a.r2), Math.max(a.r1, a.r2)];
}
/** Area with a new [lo, hi] on the op axis (other axis kept). */
function withAxisSpan(a: Area, isCol: boolean, lo: number, hi: number): Area {
  const na: Area = isCol ? { ...a, c1: lo, c2: hi } : { ...a, r1: lo, r2: hi };
  if (na.kind === 'cell' || na.kind === 'area') na.kind = na.c1 === na.c2 && na.r1 === na.r2 ? 'cell' : 'area';
  return na;
}
/** Top-left (row, col) of the bounding box of a list of areas (whole columns/rows start at 1). */
export function sqrefTopLeft(areas: Area[]): { row: number; col: number } {
  let row = Infinity, col = Infinity;
  for (const a of areas) {
    row = Math.min(row, axisSpan(a, false)[0]);
    col = Math.min(col, axisSpan(a, true)[0]);
  }
  return { row: row === Infinity ? 1 : row, col: col === Infinity ? 1 : col };
}

/**
 * A run of cells of one sqref area whose (per-cell) formulas share one relative (R1C1) form. `formulas` are
 * written for the cell `host` (= the first cell of the run).
 */
interface CellSeg {
  lo: number;
  hi: number;
  /** source area: its other-axis span is the band of the run */
  area: Area;
  host: [row: number, col: number];
  formulas: string[];
  key: string;
  /** area spanning the whole op axis (kept as written) */
  whole?: boolean;
}

const segKey = (formulas: string[], host: [number, number]) => formulas.map((f) => relKey(f, host[0], host[1])).join('\u0000');

/** Relative offsets (op axis, from the anchor) used by the formulas: a per-cell run is cut at `boundary - offset`. */
function relOffsets(formulas: string[], isCol: boolean, A0: { row: number; col: number }): Set<number> {
  const offs = new Set<number>([0]);
  for (const f of formulas) for (const t of tokenize(f)) if (t.type === 'ref' && t.area) {
    const a = t.area;
    if (isCol) { if (!a.ac1 && a.c1) offs.add(a.c1 - A0.col); if (!a.ac2 && a.c2) offs.add(a.c2 - A0.col); }
    else { if (!a.ar1 && a.r1) offs.add(a.r1 - A0.row); if (!a.ar2 && a.r2) offs.add(a.r2 - A0.row); }
  }
  return offs;
}

/** Per-area helpers: the cell at index v of the area's band, and the rule's formulas copied to that cell. */
function areaCells(rule: AnchoredRule, area: Area, isCol: boolean, A0: { row: number; col: number }) {
  const other = axisSpan(area, !isCol)[0];
  const cell = (v: number): [number, number] => (isCol ? [other, v] : [v, other]);
  const copyAt = (v: number) => { const [r, c] = cell(v); return rule.formulas.map((f) => translateFormula(f, r - A0.row, c - A0.col)); };
  return { cell, copyAt };
}

/**
 * Insert / delete on one rule, Excel's model (measured: cfSpread.test.ts, cf.test.ts, r2b.test.ts):
 *  - cells BEFORE the op point keep their place; each of them holds its own copy of the formula, adjusted like a
 *    normal formula hosted there (the rule splits where relative references cross the op point);
 *  - cells inserted inside an area (always) or right after it (`grow`, format inheritance) take the formula of the
 *    cell before them;
 *  - the cells of an area AFTER the op point shift and all take the formula of the first of them, adjusted as if
 *    hosted there (Excel does not split that part even when relative references cross the op point);
 *  - deleted cells drop out; an area spanning the whole op axis never changes (only absolute parts move when its
 *    first cell is inserted / deleted); an area reaching the last row/column stays anchored there on delete;
 *    inserts clamp at the sheet edge.
 * Cost O(#areas x #relative offsets), not O(#cells).
 */
function ruleSegments(rule: AnchoredRule, op: Exclude<StructOp, { type: 'move' }>, sheet: string, grow: boolean): CellSeg[] {
  const isCol = op.axis === 'col';
  const MAXA = isCol ? MAX_COL : MAX_ROW;
  const A0 = sqrefTopLeft(rule.sqref);
  const ctx: AdjustContext = { hostSheet: sheet, targetSheet: sheet };
  const offs = relOffsets(rule.formulas, isCol, A0);
  const n = op.count, at = op.at, ins = op.type === 'insert';
  const bounds = ins ? [at] : [at, at + n];
  const moveFrom = ins ? at : at + n; // first host that moves
  const out: CellSeg[] = [];
  for (const area of rule.sqref) {
    const [lo, hi] = axisSpan(area, isCol);
    const { cell, copyAt } = areaCells(rule, area, isCol, A0);
    if (lo === 1 && hi === MAXA) {
      const relOnly = at === 1;
      const fs = copyAt(lo).map((f) => adjustFormula(f, op, relOnly ? { ...ctx, isName: true } : ctx).formula);
      out.push({ lo, hi, area, host: cell(lo), formulas: fs, key: segKey(fs, cell(lo)), whole: true });
      continue;
    }
    // cells before the op point: per cell
    const lHi = Math.min(hi, at - 1);
    if (lo <= lHi) {
      const bp = new Set<number>([lo, lHi + 1]);
      for (const b of bounds) for (const o of offs) { const x = b - o; if (x > lo && x <= lHi) bp.add(x); }
      const pts = [...bp].sort((x, y) => x - y);
      for (let i = 0; i < pts.length - 1; i++) {
        const segLo = pts[i], segHi = pts[i + 1] - 1;
        const nh = ins && segHi === at - 1 && (at <= hi || grow) ? Math.min(at - 1 + n, MAXA) : segHi;
        const fs = copyAt(segLo).map((f) => adjustFormula(f, op, ctx).formula);
        out.push({ lo: segLo, hi: nh, area, host: cell(segLo), formulas: fs, key: segKey(fs, cell(segLo)) });
      }
    }
    // cells after the op point: one formula (the first such cell's) for the rest of the area
    const pl = Math.max(lo, moveFrom);
    if (pl <= hi) {
      const nl = ins ? pl + n : pl - n;
      const nh = ins ? Math.min(hi + n, MAXA) : hi === MAXA ? MAXA : hi - n;
      if (nl <= MAXA) {
        const fs = copyAt(pl).map((f) => adjustFormula(f, op, ctx).formula);
        out.push({ lo: nl, hi: nh, area, host: cell(nl), formulas: fs, key: segKey(fs, cell(nl)) });
      }
    }
  }
  return out;
}

/**
 * The paste step of an Excel move (R2b R3 step 2) for a span: cut [a, a+n-1] pasted onto [X, X+n-1] (disjoint).
 * A span inside the source travels; a single cell outside it stays, except a cell of the destination (pasted
 * over: #REF!, null); a span containing both the source and the destination stays; otherwise each end lying in the
 * source travels, and an inverted result leaves the span unchanged.
 */
function cutSpan(lo: number, hi: number, a: number, n: number, X: number, single: boolean): Span {
  const b = a + n - 1, d = X - a;
  if (lo >= a && hi <= b) return [lo + d, hi + d];
  if (single) return lo >= X && lo <= X + n - 1 ? null : [lo, hi];
  if (lo <= a && hi >= b && lo <= X && hi >= X + n - 1) return [lo, hi];
  const nl = lo >= a && lo <= b ? lo + d : lo, nh = hi >= a && hi <= b ? hi + d : hi;
  return nl > nh ? [lo, hi] : [nl, nh];
}

/**
 * References of `formula` (on `sheet`) through the paste step of a move (cutSpan). `absOnly`: only the parts written
 * absolute ($) on the op axis are adjusted (the formula of a range object Excel copies AFTER adjusting the original:
 * relative parts were translated to the copy's own anchor, not adjusted).
 */
function cutFormula(formula: string, sheet: string, isCol: boolean, a: number, n: number, X: number, absOnly = false): string {
  const toks = tokenize(formula);
  let changed = false;
  for (let i = 0; i < toks.length; i++) {
    const t = toks[i];
    if (t.type !== 'ref' || !t.area) continue;
    const p = t.prefix;
    if (p && (p.book || p.sheet2 !== undefined || !sameSheet(p.sheet, sheet))) continue;
    const ar = t.area;
    if (isCol ? ar.kind === 'rows' : ar.kind === 'cols') continue;
    const e1 = isCol ? ar.c1 : ar.r1, e2 = isCol ? ar.c2 : ar.r2;
    const abs1 = isCol ? ar.ac1 : ar.ar1, abs2 = isCol ? ar.ac2 : ar.ar2;
    if (absOnly && !abs1 && !abs2) continue;
    const pt = p ? p.text : '';
    if (absOnly && abs1 !== abs2) {
      // one end absolute: only that end travels with the block
      const mv = (e: number): number => { const q = cutSpan(e, e, a, n, X, true); return q ? q[0] : -1; };
      const m1 = abs1 ? mv(e1) : e1, m2 = abs2 ? mv(e2) : e2;
      if (m1 < 0 || m2 < 0) {
        toks[i] = { type: 'error', text: pt + '#REF!' + trimSuffix(ar.trim), start: t.start, end: t.end, prefix: p, ...(ar.trim ? { trim: ar.trim } : {}) };
        changed = true;
      } else if (m1 !== e1 || m2 !== e2) {
        const na: Area = normalizeFull(isCol ? { ...ar, c1: m1, c2: m2 } : { ...ar, r1: m1, r2: m2 });
        toks[i] = { ...t, text: pt + renderArea(na), area: na };
        changed = true;
      }
      continue;
    }
    const r = cutSpan(Math.min(e1, e2), Math.max(e1, e2), a, n, X, ar.kind === 'cell');
    if (!r) {
      toks[i] = { type: 'error', text: pt + '#REF!' + trimSuffix(ar.trim), start: t.start, end: t.end, prefix: p, ...(ar.trim ? { trim: ar.trim } : {}) };
      changed = true;
      continue;
    }
    const [nl, nh] = e1 > e2 ? [r[1], r[0]] : r;
    if (nl === e1 && nh === e2) continue;
    const na: Area = normalizeFull(isCol ? { ...ar, c1: nl, c2: nh } : { ...ar, r1: nl, r2: nh });
    toks[i] = { ...t, text: pt + renderArea(na), area: na };
    changed = true;
  }
  return changed ? toks.map((t) => t.text).join('') : formula;
}

/**
 * Paste step on one rule, area by area (an area = one range object of Excel, with ONE formula anchored at its first
 * cell; measured with cfProbe.test.ts p1 / p2 / p5 and cfSpread.test.ts):
 *  - the cells of the source block [a, a+n-1] travel to the destination [X, X+n-1], each with its own formula text
 *    (references into the block follow it, references to the destination die: cutFormula), so runs are split where a
 *    relative reference crosses a block boundary; the cells LEFT of the block stay in place and are split the same way
 *    (each holds its own adjusted formula);
 *  - the part right of the block is ONE piece: the formula of its first cell, untouched when the block does not start
 *    the area (a new range object copied from the original: a relative reference to the block dies in the delete step,
 *    an absolute one follows it: cutFormula absOnly), adjusted as a whole when the block started the area (the same
 *    range object, re-anchored); an area the block does not touch is adjusted as a whole as well;
 *  - the destination cells (which copied X-1's format in the insert step) keep the rule only when the area does not
 *    contain the block and `keepDest(key, cell, area)` allows it, with their formula unchanged.
 */
function cutSegments(rule: AnchoredRule, sheet: string, isCol: boolean, a: number, n: number, X: number, keepDest: (pos: number, area: Area) => boolean): CellSeg[] {
  const MAXA = isCol ? MAX_COL : MAX_ROW;
  const A0 = sqrefTopLeft(rule.sqref);
  const offs = relOffsets(rule.formulas, isCol, A0);
  const bounds = [a, a + n, X, X + n];
  const d = X - a;
  const out: CellSeg[] = [];
  for (const area of rule.sqref) {
    const [lo, hi] = axisSpan(area, isCol);
    const { cell, copyAt } = areaCells(rule, area, isCol, A0);
    const cut = (f: string) => cutFormula(f, sheet, isCol, a, n, X);
    if (lo === 1 && hi === MAXA) {
      const fs = copyAt(lo).map(cut);
      out.push({ lo, hi, area, host: cell(lo), formulas: fs, key: segKey(fs, cell(lo)), whole: true });
      continue;
    }
    const hasSrc = lo <= a + n - 1 && hi >= a;
    const bp = new Set<number>([lo, hi + 1]);
    for (const b of bounds) {
      if (b > lo && b <= hi) bp.add(b);
      // cells left of the cut block and the block itself: one formula per cell (split where a relative reference
      // crosses a block boundary)
      if (b === a || b === a + n) for (const o of offs) { const x = b - o; if (x > lo && x <= hi && x < a + n) bp.add(x); }
    }
    const pts = [...bp].sort((x, y) => x - y);
    // Pieces that are not cells of the block and not left of it:
    //  - the block started the area: the same range object, re-anchored at its first remaining cell (`base`): the
    //    formula of that cell adjusted as a whole (`anchorFs`), translated to later pieces;
    //  - an area the block does not touch: its first piece is adjusted as a whole (one formula for all its cells);
    //  - otherwise (the part right of the block, pieces after the destination): a NEW range object copied from the
    //    original anchor formula translated to the piece. A relative reference to the cut cells is not adjusted there
    //    and dies with the emptied source in the delete step; an absolute one follows the block (absOnly).
    const base = Math.max(lo, a + n);
    const anchorFs = hasSrc && lo >= a ? copyAt(base).map(cut) : [];
    const shiftTo = (fs: string[], v: number): string[] => {
      const [r0, c0] = cell(base), [r1, c1] = cell(v);
      return r0 === r1 && c0 === c1 ? fs : fs.map((f) => translateFormula(f, r1 - r0, c1 - c0));
    };
    for (let i = 0; i < pts.length - 1; i++) {
      const segLo = pts[i], segHi = pts[i + 1] - 1;
      const copy = copyAt(segLo);
      if (segLo >= X && segLo <= X + n - 1) {
        const key = segKey(copy, cell(segLo));
        if (!hasSrc && keepDest(segLo, area)) out.push({ lo: segLo, hi: segHi, area, host: cell(segLo), formulas: copy, key });
        continue;
      }
      const inSrc = segLo >= a && segLo <= a + n - 1;
      const nl = inSrc ? segLo + d : segLo, nh = inSrc ? segHi + d : segHi;
      let fs: string[];
      if (inSrc || segHi < a) fs = copy.map(cut);
      else if (hasSrc && lo >= a) fs = shiftTo(anchorFs, segLo);
      else if (!hasSrc && segLo === lo) fs = copy.map(cut);
      else fs = copy.map((f) => cutFormula(f, sheet, isCol, a, n, X, true));
      out.push({ lo: nl, hi: nh, area, host: cell(nl), formulas: fs, key: segKey(fs, cell(nl)) });
    }
  }
  return out;
}

/**
 * Does any cell of the run refer (on `sheet`) to a cell of `src` or of `dst` (op-axis intervals = whole columns /
 * rows)? A cell referring to itself does not count, nor does a run inside `src` whose relative references stay
 * inside `src`.
 */
function runTouches(q: CellSeg, src: [number, number], dst: [number, number], isCol: boolean, sheet: string): boolean {
  const K = q.hi - q.lo;
  const MAXA = isCol ? MAX_COL : MAX_ROW;
  const runInSrc = q.lo >= src[0] && q.hi <= src[1];
  const meets = (lo: number, hi: number, x: [number, number]) => lo <= x[1] && hi >= x[0];
  for (const f of q.formulas) for (const t of tokenize(f)) {
    if (t.type !== 'ref' || !t.area) continue;
    const pf = t.prefix;
    if (pf && (pf.book || pf.sheet2 !== undefined || !sameSheet(pf.sheet, sheet))) continue;
    const a = t.area;
    let lo: number, hi: number;
    if (isCol ? a.kind === 'rows' : a.kind === 'cols') { lo = 1; hi = MAXA; } // spans the whole op axis
    else {
      const e1 = isCol ? a.c1 : a.r1, e2 = isCol ? a.c2 : a.r2;
      const m1 = (isCol ? a.ac1 : a.ar1) ? 0 : K, m2 = (isCol ? a.ac2 : a.ar2) ? 0 : K; // relative ends follow the host
      lo = Math.min(e1, e2, e1 + m1, e2 + m2);
      hi = Math.max(e1, e2, e1 + m1, e2 + m2);
    }
    if (meets(lo, hi, src)) {
      // a run inside the block whose relative references stay inside it travels with them (nothing changes)
      const rel1 = !(isCol ? a.ac1 : a.ar1), rel2 = !(isCol ? a.ac2 : a.ar2);
      if (!(runInSrc && rel1 && rel2 && lo >= src[0] && hi <= src[1])) return true;
    } else if (meets(lo, hi, dst)) {
      const self = a.kind === 'cell' && !a.ac1 && !a.ar1 && a.r1 === q.host[0] && a.c1 === q.host[1];
      if (!self) return true;
    }
  }
  return false;
}

/** Runs -> rules: one rule per relative formula; runs of one band merged; formulas re-anchored at the bounding box. */
function buildRules(segs: CellSeg[], isCol: boolean): AnchoredRule[] {
  const groups = new Map<string, CellSeg[]>();
  for (const s of segs) { const g = groups.get(s.key); if (g) g.push(s); else groups.set(s.key, [s]); }
  const out: AnchoredRule[] = [];
  for (const g of groups.values()) {
    const areas: Area[] = [];
    const byBand = new Map<string, CellSeg[]>();
    for (const s of g) {
      if (s.whole) { areas.push(s.area); continue; }
      const b = axisSpan(s.area, !isCol).join('|');
      const l = byBand.get(b);
      if (l) l.push(s); else byBand.set(b, [s]);
    }
    for (const list of byBand.values()) {
      const area = list[0].area;
      list.sort((x, y) => x.lo - y.lo);
      let lo = list[0].lo, hi = list[0].hi;
      for (let i = 1; i < list.length; i++) {
        if (list[i].lo <= hi + 1) { hi = Math.max(hi, list[i].hi); continue; }
        areas.push(withAxisSpan(area, isCol, lo, hi)); lo = list[i].lo; hi = list[i].hi;
      }
      areas.push(withAxisSpan(area, isCol, lo, hi));
    }
    if (!areas.length) continue;
    areas.sort((x, y) => (isCol ? axisSpan(x, true)[0] - axisSpan(y, true)[0] : axisSpan(x, false)[0] - axisSpan(y, false)[0]));
    const tl = sqrefTopLeft(areas);
    const src = g.find((s) => s.host[0] === tl.row && s.host[1] === tl.col) ?? g[0];
    const [hr, hc] = src.host;
    const formulas = src.formulas.map((f) => translateFormula(f, tl.row - hr, tl.col - hc));
    if (formulas.some((f, i) => f.includes('#REF!') && !src.formulas[i].includes('#REF!'))) {
      // re-anchoring at the bounding box would push a relative reference off the sheet: one rule per area
      for (const a of areas) {
        const t = sqrefTopLeft([a]);
        out.push({ sqref: [a], formulas: src.formulas.map((f) => translateFormula(f, t.row - hr, t.col - hc)) });
      }
    } else out.push({ sqref: areas, formulas });
  }
  return out;
}

/**
 * A move is Excel's sequence (R2b R3, also what the Excel runner does): insert n blank cells at X (they copy
 * X-1's format), paste the block onto them, delete the emptied source. The destination keeps X-1's rule only when
 * the rule spreads (options) and no rule of the block (own formulas + opts.siblings, each on its own) has a run
 * covering the destination that refers to the moved cells or to the destination (runTouches).
 */
function moveRule(rule: AnchoredRule, op: Extract<StructOp, { type: 'move' }>, sheet: string, opts: AnchoredRuleOptions): AnchoredRule[] {
  const isCol = op.axis === 'col';
  const n = op.count, s2 = op.from + n - 1, X = op.to;
  const a = X > s2 ? op.from : op.from + n; // the block after the insert step
  const ins: StructOp = { type: 'insert', axis: op.axis, at: X, count: n };
  const segs1 = ruleSegments(rule, ins, sheet, true);
  const spreads = ruleSpreadsOnMove(opts);
  const src: [number, number] = [a, a + n - 1], dst: [number, number] = [X, X + n - 1];
  // Every rule of the block, each on its own: the destination cells that joined the block by format inheritance leave
  // it again when the run of ANY of these rules that covers them refers to the moved block or to the destination
  // (Excel keeps one range object per block), see AnchoredRuleOptions.siblings
  const alone = (opts.blockRules ?? [rule.formulas]).map((f) => ruleSegments({ sqref: rule.sqref, formulas: f }, ins, sheet, true));
  const band = (ar: Area) => axisSpan(ar, !isCol).join('|');
  const keepDest = (pos: number, area: Area): boolean => {
    if (!spreads) return false;
    const bd = band(area);
    return !alone.some((S) => S.some((r) => !r.whole && r.lo <= pos && r.hi >= pos && band(r.area) === bd && runTouches(r, src, dst, isCol, sheet)));
  };
  const rules1 = buildRules(segs1, isCol);
  const rules2 = buildRules(rules1.flatMap((r) => cutSegments(r, sheet, isCol, a, n, X, keepDest)), isCol);
  const del: StructOp = { type: 'delete', axis: op.axis, at: a, count: n };
  return buildRules(rules2.flatMap((r) => ruleSegments(r, del, sheet, false)), isCol);
}

/**
 * CF / DV rule (sqref + formulas anchored at the top-left of the sqref bounding box) under a structural op on its
 * own sheet, as Excel does it — verified against three independent Excel 365 ground truths (tests/formula/cf.test.ts,
 * r2b.test.ts, cfSpread.test.ts; tests/formula/DEVIATIONS.md lists what is left):
 *  - insert / delete: see ruleSegments (cells before the op point per cell, the rest of each area with the formula
 *    of its first shifted cell; format inheritance over cells inserted right after an area);
 *  - move: Excel's insert + paste + delete sequence (moveRule), including the "spread" of X-1's rule onto the block
 *    and the #REF! Excel writes for references to the pasted-over destination cells;
 *  - the result is regrouped into one rule per relative (R1C1) formula, re-anchored at each group's bounding-box
 *    top-left (several rules, or [] when every cell was deleted).
 * The option shape is documented at AnchoredRuleOptions.
 */
export function adjustAnchoredRule(rule: AnchoredRule, op: StructOp, sheet: string, opts: AnchoredRuleOptions = {}): AnchoredRule[] {
  if (!rule.sqref.length) return [rule];
  const sib = opts.siblings?.filter((f) => f.length);
  if (sib?.length) {
    // One range object per <conditionalFormatting> block: every rule of the block is fragmented at the same places and
    // a sibling's references take part in the move decisions. Process the block's formulas together, keep our slots.
    const own = rule.formulas.length;
    const res = adjustAnchoredRuleBlock({ sqref: rule.sqref, formulas: [...rule.formulas, ...sib.flat()] }, op, sheet, { ...opts, blockRules: [rule.formulas, ...sib] });
    return res.map((r) => ({ sqref: r.sqref, formulas: r.formulas.slice(0, own) }));
  }
  return adjustAnchoredRuleBlock(rule, op, sheet, opts);
}

/** adjustAnchoredRule for a rule whose formulas already include the block's (see AnchoredRuleOptions.siblings). */
function adjustAnchoredRuleBlock(rule: AnchoredRule, op: StructOp, sheet: string, opts: AnchoredRuleOptions): AnchoredRule[] {
  if (op.type === 'move') {
    if (op.to >= op.from && op.to <= op.from + op.count) return [rule]; // no-op move
    return moveRule(rule, op, sheet, opts);
  }
  return buildRules(ruleSegments(rule, op, sheet, ruleGrowsOnInsert(opts)), op.axis === 'col');
}

/** Same as adjustAnchoredRule (kept for round-1 callers; every op uses the per-cell model since round 2). */
export function adjustAnchoredRulePerCell(rule: AnchoredRule, op: StructOp, sheet: string, opts: AnchoredRuleOptions = {}): AnchoredRule[] {
  return adjustAnchoredRule(rule, op, sheet, opts);
}

// ---------------------------------------------------------------------------------------------
// multi-selection move planner
// ---------------------------------------------------------------------------------------------
/**
 * Decompose "move the (possibly non-contiguous) selected columns so that they end up, in their original
 * order, right before column `target`" into contiguous block moves (Excel can only move one contiguous
 * block at a time). Contiguous runs are moved as ONE block (same result as doing it by hand in Excel).
 * The same op list is replayed by the Excel-COM engine and by adjustFormulaSeq().
 */
export function planMove(selected: number[], target: number, axis: Axis = 'col'): StructOp[] {
  const sel = [...new Set(selected)].sort((a, b) => a - b);
  const runs: [number, number][] = [];
  for (const v of sel) { const last = runs[runs.length - 1]; if (last && v === last[1] + 1) last[1] = v; else runs.push([v, v]); }
  const ops: StructOp[] = [];
  const left: [number, number][] = [], right: [number, number][] = [];
  for (const [a, b] of runs) {
    if (b < target) left.push([a, b]);
    else if (a >= target) right.push([a, b]);
    else { left.push([a, target - 1]); right.push([target, b]); } // run straddles the drop position
  }
  let ins = target;
  for (let i = left.length - 1; i >= 0; i--) {
    const [a, b] = left[i]; const n = b - a + 1;
    if (b + 1 !== ins) ops.push({ type: 'move', axis, from: a, count: n, to: ins });
    ins -= n; // the moved run now occupies [ins-n, ins-1]; the next run goes right before it
  }
  let ins2 = target;
  for (const [a, b] of right) {
    const n = b - a + 1;
    if (a !== ins2) ops.push({ type: 'move', axis, from: a, count: n, to: ins2 });
    ins2 += n;
  }
  return ops;
}
