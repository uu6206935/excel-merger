/**
 * Function-call structure over excelRef tokens (owner: formula engine F).
 * Pure helper shared by forms.ts (LET/LAMBDA scopes, _xlfn.SINGLE / ANCHORARRAY) and excelRef.lookupWarnings.
 * Works on any token list: unbalanced input simply yields calls without a closing parenthesis (close = -1).
 */
import type { Token } from './excelRef';

export interface CallInfo {
  /** index of the 'func' token */
  func: number;
  /** function name without _xlfn. / _xlws. / _xlpm. prefixes (and without a sheet/book prefix), upper case */
  name: string;
  /** index of the '(' token (always func + 1) */
  open: number;
  /** index of the matching ')' token, -1 when unbalanced */
  close: number;
  /** argument token ranges [start, end) (end exclusive), split at depth-0 ',' */
  args: Array<[number, number]>;
}

const PREFIX_RE = /^(?:_xlfn\.|_xlws\.|_xlpm\.)+/i;

/** Bare, upper-case function name of a 'func' token ("_xlfn._xlws.SORT" -> "SORT", "Sheet1!Fn" -> "FN"). */
export function funcName(t: Token): string {
  const text = t.prefix ? t.text.slice(t.prefix.text.length) : t.text;
  return text.replace(PREFIX_RE, '').toUpperCase();
}

/** Matching bracket index for every '(' / '{' op token (-1 when unbalanced). */
export function matchBrackets(toks: Token[]): Int32Array {
  const m = new Int32Array(toks.length).fill(-1);
  const stack: number[] = [];
  for (let i = 0; i < toks.length; i++) {
    const t = toks[i];
    if (t.type !== 'op') continue;
    if (t.text === '(' || t.text === '{') stack.push(i);
    else if (t.text === ')' || t.text === '}') {
      const want = t.text === ')' ? '(' : '{';
      // pop until the matching kind (tolerates garbage)
      for (let k = stack.length - 1; k >= 0; k--) {
        if (toks[stack[k]].text === want) { const o = stack[k]; stack.length = k; m[o] = i; m[i] = o; break; }
      }
    }
  }
  return m;
}

/** Every function call of the token list, keyed by the index of its 'func' token. */
export function parseCalls(toks: Token[], brackets: Int32Array = matchBrackets(toks)): Map<number, CallInfo> {
  const out = new Map<number, CallInfo>();
  for (let i = 0; i < toks.length - 1; i++) {
    if (toks[i].type !== 'func' || toks[i + 1].text !== '(') continue;
    const open = i + 1, close = brackets[open];
    const end = close < 0 ? toks.length : close;
    const args: Array<[number, number]> = [];
    let s = open + 1;
    for (let k = open + 1; k < end; k++) {
      const t = toks[k];
      if (t.type === 'op' && (t.text === '(' || t.text === '{') && brackets[k] > 0) { k = brackets[k]; continue; }
      if (t.type === 'op' && t.text === ',') { args.push([s, k]); s = k + 1; }
    }
    if (end > open + 1 || args.length) args.push([s, end]);
    out.set(i, { func: i, name: funcName(toks[i]), open, close, args });
  }
  return out;
}

/** [start, end) with leading / trailing whitespace tokens removed. */
export function trimRange(toks: Token[], [s, e]: [number, number]): [number, number] {
  while (s < e && toks[s].type === 'ws') s++;
  while (e > s && toks[e - 1].type === 'ws') e--;
  return [s, e];
}
