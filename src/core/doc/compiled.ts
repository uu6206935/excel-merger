/**
 * Compiled formulas: a formula split into literal parts and A1 reference tokens (from excelRef.tokenize), so
 * that structural ops can be applied without re-tokenizing. The output of an adjustment carries its own compiled
 * form, so a formula is tokenized at most once per editor session (util.ts keeps compiled forms aligned with the
 * immutable arrays that hold the formulas: a column's cells, doc.externalFormulas).
 *
 * Semantics are EXACTLY excelRef.adjustFormula (applied op by op, i.e. adjustFormulaSeq):
 *  - only 'ref' tokens are rewritten; external ([1]Sheet!A1) and 3D refs are never touched;
 *  - a ref is adjusted when its resolved sheet (prefix, else host sheet) equals the target sheet (case-insensitive);
 *  - a deleted ref becomes <prefix>#REF! (and is never adjusted again);
 *  - an untouched ref keeps its exact text; a changed one is rendered by renderArea.
 * tests/doc/compiled.test.ts checks this against adjustFormulaSeq on thousands of formulas.
 *
 * Extra (names only, R2b R4): references inside DEFINED NAMES that are fully relative (no '$' at all) are stored
 * as offsets by Excel and are NOT adjusted by structural ops (`skipRelative`).
 */
import { adjustArea, renderArea, tokenize, MAX_COL, MAX_ROW } from '../formula/excelRef';
import type { Area, StructOp } from '../formula/excelRef';

export interface CRef {
  /** verbatim token text (prefix + area) */
  text: string;
  /** verbatim prefix incl. '!' ('' when unqualified) */
  prefix: string;
  /** lower-cased sheet name of the prefix; null when unqualified (= host sheet) */
  sheet: string | null;
  /** external-book or 3D reference: never adjusted */
  fixed: boolean;
  area: Area;
}

export interface Compiled {
  /** parts[i] precedes refs[i]; parts[refs.length] is the tail. */
  parts: string[];
  refs: CRef[];
  /** tokenizer produced 'unknown' tokens */
  unsafe: boolean;
  /** contains structured references (Table1[...]) */
  sref: boolean;
  /**
   * contains functions that address a range's columns by NUMBER (VLOOKUP column index, INDEX/OFFSET/MATCH
   * positions, CHOOSECOLS ...). Excel never adjusts those numbers, so moving/deleting columns inside such a
   * range silently changes the result (warned about by the doc ops).
   */
  lookup: boolean;
}

// every function excelRef.lookupWarnings checks (plus column-number functions it may learn later)
const LOOKUP_FUNCS = new Set(['VLOOKUP', 'HLOOKUP', 'INDEX', 'OFFSET', 'MATCH', 'XMATCH', 'XLOOKUP', 'CHOOSE', 'CHOOSECOLS', 'TAKE', 'DROP']);

export function compile(text: string): Compiled {
  const toks = tokenize(text);
  const parts: string[] = [];
  const refs: CRef[] = [];
  let cur = '';
  let unsafe = false;
  let sref = false;
  let lookup = false;
  for (const t of toks) {
    if (t.type === 'ref' && t.area) {
      parts.push(cur);
      cur = '';
      const p = t.prefix;
      refs.push({
        text: t.text,
        prefix: p ? p.text : '',
        sheet: p && p.sheet != null ? p.sheet.toLowerCase() : null,
        fixed: !!(p && (p.book || p.sheet2 !== undefined)),
        area: t.area,
      });
      continue;
    }
    if (t.type === 'unknown') unsafe = true;
    else if (t.type === 'sref') sref = true;
    else if (t.type === 'func') {
      const fn = t.text.replace(/^(_xlfn\.|_xlws\.)+/i, '').toUpperCase();
      if (LOOKUP_FUNCS.has(fn)) lookup = true;
    }
    cur += t.text;
  }
  parts.push(cur);
  return { parts, refs, unsafe, sref, lookup };
}

export function render(c: Compiled): string {
  let s = c.parts[0];
  for (let i = 0; i < c.refs.length; i++) s += c.refs[i].text + c.parts[i + 1];
  return s;
}

const isFullyRelative = (a: Area): boolean => !a.ac1 && !a.ar1 && !a.ac2 && !a.ar2;

export interface AdjustOut {
  /** same object as the input when nothing changed */
  compiled: Compiled;
  text: string;
  changed: boolean;
  createdRefError: boolean;
}

/**
 * Per-batch memo of reference adjustments, keyed by the (immutable, often shared) CRef object: filled-down copies
 * of a formula share the CRef objects of their absolute references (compileWithHint), so e.g. the $A$2:$F$9999 of
 * 10,000 VLOOKUPs is adjusted once. Valid for one (ops, host sheet, target sheet, skipRelative) combination.
 */
export type AdjustMemo = Map<CRef, CRef | null>;
/** memo value: null = unchanged; DEAD_REF area = became #REF! */

/**
 * Apply `ops` (all on `targetLc`) to a compiled formula hosted on `hostLc` (lower-cased; null = no host sheet).
 * Returns null when no reference is affected.
 */
export function adjustCompiled(
  c: Compiled,
  ops: readonly StructOp[],
  hostLc: string | null,
  targetLc: string,
  skipRelative = false,
  memo?: AdjustMemo,
): AdjustOut | null {
  const refs = c.refs;
  const n = refs.length;
  let repl: Array<CRef | undefined> | null = null; // replaced refs by index (lazily allocated)
  let createdRefError = false;
  let dead = false;
  for (let i = 0; i < n; i++) {
    const r = refs[i];
    if (r.fixed) continue;
    const sh = r.sheet ?? hostLc;
    if (sh === null || sh !== targetLc) continue;
    if (skipRelative && isFullyRelative(r.area)) continue;
    let nr: CRef | null | undefined = memo?.get(r);
    if (nr === undefined) {
      let area: Area | null = r.area;
      let changed = false;
      for (const op of ops) {
        const na: Area | null = adjustArea(area, op);
        if (!na) {
          area = null;
          break;
        }
        if (na !== area) changed = true;
        area = na;
      }
      if (area && !changed) nr = null;
      else if (!area) nr = { text: r.prefix + '#REF!', prefix: r.prefix, sheet: r.sheet, fixed: r.fixed, area: DEAD };
      else nr = { text: r.prefix + renderArea(area), prefix: r.prefix, sheet: r.sheet, fixed: r.fixed, area };
      memo?.set(r, nr);
    }
    if (nr === null) continue;
    if (!repl) repl = new Array<CRef | undefined>(n);
    repl[i] = nr;
    if (nr.area === DEAD) {
      createdRefError = true;
      dead = true;
    }
  }
  if (!repl) return null;
  // one pass: output text, refs (dead refs are folded into the literal parts: they are error tokens now)
  const src = c.parts;
  let text = src[0];
  let changedText = false;
  const outRefs: CRef[] = new Array<CRef>(dead ? 0 : n);
  const parts: string[] = dead ? [src[0]] : src;
  for (let i = 0; i < n; i++) {
    const r = repl[i] ?? refs[i];
    if (r !== refs[i] && r.text !== refs[i].text) changedText = true;
    text += r.text + src[i + 1];
    if (!dead) outRefs[i] = r;
    else if (r.area === DEAD) parts[parts.length - 1] += r.text + src[i + 1];
    else {
      outRefs.push(r);
      parts.push(src[i + 1]);
    }
  }
  const compiled: Compiled = { parts, refs: outRefs, unsafe: c.unsafe, sref: c.sref, lookup: c.lookup };
  return { compiled, text, changed: changedText, createdRefError };
}

const DEAD: Area = { kind: 'cell', c1: 0, r1: 0, c2: 0, r2: 0, ac1: false, ar1: false, ac2: false, ar2: false };

/**
 * Compile `text` using the compiled form of a neighbouring formula (typically the previous row of the same
 * column, i.e. a filled-down copy): the neighbour is translated by (dr, dc) like excelRef.translateFormula and
 * rendered; when the result equals `text` its compiled form is reused without tokenizing.
 */
export function compileWithHint(text: string, hint: Compiled | null, dr: number, dc: number): Compiled {
  if (hint && (dr !== 0 || dc !== 0) && hint.refs.length > 0) {
    const refs: CRef[] = new Array(hint.refs.length);
    let s = hint.parts[0];
    let ok = true;
    for (let i = 0; i < hint.refs.length; i++) {
      const r = hint.refs[i];
      const a = r.area;
      const mv = (v: number, abs: boolean, d: number) => (abs || v === 0 ? v : v + d);
      const c1 = mv(a.c1, a.ac1, dc), c2 = mv(a.c2, a.ac2, dc), r1 = mv(a.r1, a.ar1, dr), r2 = mv(a.r2, a.ar2, dr);
      if (c1 === a.c1 && c2 === a.c2 && r1 === a.r1 && r2 === a.r2) {
        refs[i] = r;
      } else {
        if (
          (a.kind !== 'rows' && (c1 < 1 || c2 < 1 || c1 > MAX_COL || c2 > MAX_COL)) ||
          (a.kind !== 'cols' && (r1 < 1 || r2 < 1 || r1 > MAX_ROW || r2 > MAX_ROW))
        ) {
          ok = false;
          break;
        }
        const na: Area = { ...a, c1, c2, r1, r2 };
        refs[i] = { ...r, text: r.prefix + renderArea(na), area: na };
      }
      s += refs[i].text + hint.parts[i + 1];
      if (s.length > text.length) {
        ok = false;
        break;
      }
    }
    if (ok && s === text) return { parts: hint.parts, refs, unsafe: hint.unsafe, sref: hint.sref, lookup: hint.lookup };
  } else if (hint && dr === 0 && dc === 0 && render(hint) === text) {
    return hint;
  }
  return compile(text);
}
