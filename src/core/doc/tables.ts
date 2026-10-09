/**
 * Excel tables (ListObjects) — only what the editor needs:
 *  - renaming a table column (editing a header cell renames the column; Excel rewrites structured references);
 *  - the names Excel gives to table columns created by a whole-column insert (measured with Excel 365 via COM:
 *    left neighbour "h1" -> "h2" when free, else "h12", "h13" ...; a left neighbour without a trailing number
 *    -> "列1", "列2" ... the first free one);
 *  - table geometry helpers.
 * Deleting table columns (structured refs -> #REF!) is done by excelRef.adjustFormula (AdjustContext.tables).
 * Moves and inserts never change structured references.
 */
import { tokenize, numToCol } from '../formula/excelRef';
import type { TableRef } from '../formula/excelRef';
import type { TableInfo } from '../types';
import { parseBox } from './util';

export interface ParsedSref {
  /** table name as written ('' for a bare [..] reference) */
  table: string;
  /** column names mentioned (unescaped), in order */
  columns: string[];
  /** offsets (in the token text) of each column-name item: [start, end) of the escaped name */
  spans: Array<[number, number]>;
}

const unescape = (s: string): string => s.replace(/'(.)/gs, '$1');
export const escapeSrefName = (s: string): string => s.replace(/['[\]#@]/g, "'$&");

/** Parse the text of an 'sref' token without prefix (Table1[..] or [..]). null = not understood. */
export function parseSref(text: string): ParsedSref | null {
  const b = text.indexOf('[');
  if (b < 0 || !text.endsWith(']')) return null;
  const table = text.slice(0, b);
  const columns: string[] = [];
  const spans: Array<[number, number]> = [];
  const innerStart = b + 1;
  const innerEnd = text.length - 1;
  let nested = false;
  for (let k = innerStart; k < innerEnd; k++) {
    if (text[k] === "'") {
      k++;
      continue;
    }
    if (text[k] === '[') {
      nested = true;
      break;
    }
  }
  const addItem = (s: number, e: number) => {
    let st = s;
    // '@' (this row, UI form) may precede a name: [@col]
    while (st < e && text[st] === '@') st++;
    const raw = text.slice(st, e);
    if (raw.trim() === '' || raw.startsWith('#')) return;
    columns.push(unescape(raw));
    spans.push([st, e]);
  };
  if (!nested) {
    addItem(innerStart, innerEnd);
    return { table, columns, spans };
  }
  let k = innerStart;
  while (k < innerEnd) {
    if (text[k] === '[') {
      let e = k + 1;
      while (e < innerEnd && text[e] !== ']') {
        if (text[e] === "'") e++;
        e++;
      }
      if (e >= innerEnd) return null;
      addItem(k + 1, e);
      k = e + 1;
      continue;
    }
    k++;
  }
  return { table, columns, spans };
}

/**
 * Rewrite every structured-reference token of `formula` with `fn` (return null to keep, a string to replace the
 * whole token text). Tokens with a prefix (external books) are left alone. Returns the same string when unchanged.
 */
export function mapSrefs(formula: string, fn: (p: ParsedSref, tokenText: string) => string | null): string {
  if (!formula.includes('[')) return formula;
  const toks = tokenize(formula);
  let changed = false;
  const out = toks.map((t) => {
    if (t.type !== 'sref' || t.prefix) return t.text;
    const p = parseSref(t.text);
    if (!p) return t.text;
    const r = fn(p, t.text);
    if (r === null || r === t.text) return t.text;
    changed = true;
    return r;
  });
  return changed ? out.join('') : formula;
}

/**
 * Rename column `oldName` of table `tableLc` (lower-cased name) to `newName` inside every structured reference.
 * `hostTable` = lower-cased name of the table containing the formula's cell (for bare [..] references), or null.
 */
export function renameSrefColumn(formula: string, tableLc: string, oldName: string, newName: string, hostTable: string | null): string {
  if (!formula.includes('[')) return formula;
  const oldLc = oldName.toLowerCase();
  return mapSrefs(formula, (p, text) => {
    const tname = p.table === '' ? hostTable : p.table.toLowerCase();
    if (tname !== tableLc) return null;
    let out = '';
    let last = 0;
    let hit = false;
    p.columns.forEach((c, i) => {
      if (c.toLowerCase() !== oldLc) return;
      const [s, e] = p.spans[i];
      out += text.slice(last, s) + escapeSrefName(newName);
      last = e;
      hit = true;
    });
    if (!hit) return null;
    return out + text.slice(last);
  });
}

/** Box of a table (1-based columns and rows). */
export function tableBox(t: TableInfo): { c1: number; c2: number; r1: number; r2: number } | null {
  return parseBox(t.ref);
}

export function boxRef(b: { c1: number; c2: number; r1: number; r2: number }): string {
  return `${numToCol(b.c1)}${b.r1}:${numToCol(b.c2)}${b.r2}`;
}

/** TableInfo[] of the merged sheet -> excelRef TableRef[] (for AdjustContext.tables). */
export function tableRefs(tables: TableInfo[], sheet: string): TableRef[] {
  return tables.map((t) => ({ name: t.name, sheet, ref: t.ref, columns: t.columns }));
}

/**
 * Names Excel gives to `count` table columns inserted right of the column named `leftName`
 * (null when there is no left table column). `existing` = current column names.
 */
export function excelNewColumnNames(existing: string[], leftName: string | null, count: number): string[] {
  const taken = new Set(existing.map((s) => s.toLowerCase()));
  const out: string[] = [];
  const take = (n: string) => {
    taken.add(n.toLowerCase());
    out.push(n);
  };
  const m = leftName != null ? /^(.*?)(\d+)$/su.exec(leftName) : null;
  for (let k = 0; k < count; k++) {
    if (m && leftName != null) {
      const next = m[1] + String(Number(m[2]) + 1);
      if (!taken.has(next.toLowerCase())) {
        take(next);
        continue;
      }
      for (let i = 2; ; i++) {
        const c = leftName + String(i);
        if (!taken.has(c.toLowerCase())) {
          take(c);
          break;
        }
      }
      continue;
    }
    for (let i = 1; ; i++) {
      const c = `列${i}`;
      if (!taken.has(c.toLowerCase())) {
        take(c);
        break;
      }
    }
  }
  return out;
}

/**
 * Header text typed by the user -> the unique column name Excel stores (duplicates get "2", "3" ... appended,
 * an empty header becomes 列N). `others` = the other columns' names.
 */
export function uniqueHeaderName(text: string, others: string[]): string {
  const taken = new Set(others.map((s) => s.toLowerCase()));
  if (text === '') return excelNewColumnNames(others, null, 1)[0];
  if (!taken.has(text.toLowerCase())) return text;
  for (let i = 2; ; i++) {
    const c = text + String(i);
    if (!taken.has(c.toLowerCase())) return c;
  }
}
