/**
 * Formula rewriting for the package patcher (owner: patcher P).
 * All reference math comes from formula/excelRef.ts (tokenize / adjustArea / parseStructuredRef ...). This file
 * only sequences SEVERAL structural ops over ONE tokenization (a replay of N ops costs one tokenize per formula)
 * and adds the package-level extras:
 *   - defined names: relative coordinates are NOT adjusted (same rule as excelRef adjustFormula `isName`),
 *   - premerge range extension (appended rows grow ranges that end at the old last data row),
 *   - structured references to deleted table columns / tables become #REF! (same rule as excelRef).
 * Semantics are EXACTLY excelRef.adjustFormulaSeq for cells and names (tests/xlsx/patch-formula.test.ts).
 */
import {
  adjustArea,
  adjustNameArea,
  parseStructuredRef,
  renderArea,
  tokenize,
  type Area,
  type StructOp,
  type Token,
} from '../../formula/excelRef';

/** A table touched by a DELETE op (excelRef deletedTableParts): column names lower-cased. */
export interface TableHit {
  whole: boolean;
  cols: Set<string>;
}

export interface OpEntry {
  /** sheet the structural op is performed on */
  sheet: string;
  op: StructOp;
  /** premerge row-append extension: areas with r1 <= firstDataRow and r2 == lastDataRow grow over the insert */
  ext?: { firstDataRow: number; lastDataRow: number };
  /** tables (lower-cased name) whose columns this DELETE op removes */
  tableHits?: Map<string, TableHit>;
}

export interface AdjustOut {
  text: string;
  changed: boolean;
  refError: boolean;
  unsafe: boolean;
}

export type AdjustMode = 'cell' | 'name';

const lower = (s: string | null | undefined) => (s == null ? null : s.toLowerCase());

/**
 * Row-append extension: an area spanning [<=first, last] grows by the inserted rows (premerge only), so that a total / count row BELOW
 * (or above) the table, or a formula that totals the whole data column, includes the appended rows.
 *
 * `hostRow`: the row of the formula's cell AFTER the insert, when the formula sits on the sheet of the op (undefined for formulas of other
 * sheets, names, charts, CF / DV rules ...). A formula INSIDE the data rows (as of before the insert) keeps a range that merely ends at
 * the old last data row when that end is its own row and relative: =SUM($D$2:D31) in row 31 is a running total, it must not start
 * summing the appended people (found by the consistency oracle, class append-range-running). A range that cannot be a running total —
 * its end lies below the formula's row, or its end row is absolute ($D$31) — covers the whole data column and grows like a total row.
 */
export function extendArea(orig: Area, adjusted: Area, e: OpEntry, hostRow?: number): Area {
  const op = e.op;
  if (!e.ext || op.type !== 'insert' || op.axis !== 'row') return adjusted;
  if (orig.kind !== 'area' && orig.kind !== 'rows') return adjusted;
  const lo = Math.min(orig.r1, orig.r2), hi = Math.max(orig.r1, orig.r2);
  if (!(lo <= e.ext.firstDataRow && hi === e.ext.lastDataRow && hi < op.at)) return adjusted;
  if (hostRow !== undefined) {
    // the formula's row before the insert (rows from `at` on moved down by `count`)
    const before = hostRow >= op.at + op.count ? hostRow - op.count : hostRow;
    if (before >= e.ext.firstDataRow && before <= e.ext.lastDataRow && before === hi) {
      const endAbsolute = orig.r1 <= orig.r2 ? orig.ar2 : orig.ar1;
      if (!endAbsolute) return adjusted; // a running total
    }
  }
  const g = { ...adjusted };
  if (g.r1 > g.r2) g.r1 = hi + op.count; else g.r2 = hi + op.count;
  return g;
}

/** One op applied to one area (extension + name rules), null = #REF!. `hostRow`: see extendArea. */
export function adjustAreaEntry(a: Area, e: OpEntry, mode: AdjustMode, hostRow?: number): Area | null {
  const na = mode === 'name' ? adjustNameArea(a, e.op) : adjustArea(a, e.op);
  if (!na) return null;
  if (e.ext) {
    if (mode === 'name') {
      const abs = e.op.axis === 'row' ? a.ar1 && a.ar2 : a.ac1 && a.ac2;
      if (!abs) return na;
    }
    const g = extendArea(a, na, e, hostRow);
    if (g !== na) return g;
  }
  return na;
}

function srefBroken(t: Token, hits: Map<string, TableHit>): boolean {
  if (t.prefix?.book) return false;
  if (t.type === 'name') return !t.prefix && !!hits.get(t.text.toLowerCase())?.whole;
  const s = parseStructuredRef(t.prefix ? t.text.slice(t.prefix.text.length) : t.text);
  const d = s?.table ? hits.get(s.table.toLowerCase()) : undefined;
  return !!d && (d.whole || s!.columns.some((c) => d.cols.has(c.toLowerCase())));
}

/**
 * Apply ops[from..] to a formula (tokenized once).
 * `host` = sheet containing the formula (unqualified refs resolve to it); null for names / charts.
 * `skipSheet`: entries on this sheet are ignored (already applied by an anchored per-op pass).
 * `hostRow`: the row of the cell that holds the formula (cell formulas only; see extendArea).
 */
export function adjustFormulaOps(
  formula: string,
  ops: readonly OpEntry[],
  host: string | null,
  mode: AdjustMode = 'cell',
  from = 0,
  skipSheet: string | null = null,
  hostRow?: number,
): AdjustOut {
  const out: AdjustOut = { text: formula, changed: false, refError: false, unsafe: false };
  if (from >= ops.length || !formula) return out;
  const skipL = lower(skipSheet);
  let anyTable = false;
  for (let i = from; i < ops.length; i++) if (ops[i].tableHits && ops[i].tableHits!.size) { anyTable = true; break; }
  const toks = tokenize(formula);
  let res = '';
  const hostL = lower(host);
  for (const t of toks) {
    if (t.type === 'unknown') out.unsafe = true;
    if (anyTable && (t.type === 'sref' || t.type === 'name')) {
      let broken = false;
      for (let i = from; i < ops.length && !broken; i++) {
        const e = ops[i];
        if (!e.tableHits || !e.tableHits.size) continue;
        if (skipL != null && e.sheet.toLowerCase() === skipL) continue;
        if (srefBroken(t, e.tableHits)) broken = true;
      }
      if (broken) { res += '#REF!'; out.changed = true; out.refError = true; continue; }
      res += t.text;
      continue;
    }
    if (t.type !== 'ref' || !t.area) { res += t.text; continue; }
    const p = t.prefix;
    if (p && (p.book || p.sheet2 !== undefined)) { res += t.text; continue; }
    const sheetL = p ? lower(p.sheet) : hostL;
    if (sheetL == null) { res += t.text; continue; }
    let a: Area | null = t.area;
    let touched = false;
    for (let i = from; i < ops.length; i++) {
      const e = ops[i];
      if (e.sheet.toLowerCase() !== sheetL) continue;
      if (skipL != null && sheetL === skipL) continue;
      const na = adjustAreaEntry(a!, e, mode, hostRow !== undefined && sheetL === hostL ? hostRow : undefined);
      if (!na) { a = null; break; }
      if (na !== a) { touched = true; a = na; }
    }
    const prefixText = p ? p.text : '';
    let nt: string;
    if (!a) { nt = prefixText + '#REF!'; out.refError = true; }
    else nt = touched ? prefixText + renderArea(a) : t.text;
    if (nt !== t.text) out.changed = true;
    res += nt;
  }
  out.text = out.changed ? res : formula;
  return out;
}

// ---------------------------------------------------------------------------------------------------
// table column rename (a new header typed over a table header cell)
// ---------------------------------------------------------------------------------------------------

/** Escape a column name inside a structured reference item: ' before [ ] # ' (Excel's escape character). */
export function escapeSrefItem(name: string): string {
  return name.replace(/['[\]#]/g, "'$&");
}

/**
 * Rename column `oldName` of table `tableLc` (lower-cased) to `newName` in every structured reference of a
 * formula (FILE form). `hostTableLc`: the table the formula lives in, for references without a table name.
 * External / prefixed references are left alone. Returns the same string when nothing changed.
 */
export function renameTableColumnInFormula(formula: string, tableLc: string, oldName: string, newName: string, hostTableLc: string | null): string {
  if (!formula || formula.indexOf('[') < 0) return formula;
  const oldLc = oldName.toLowerCase();
  let changed = false;
  let out = '';
  for (const t of tokenize(formula)) {
    if (t.type !== 'sref' || t.prefix) { out += t.text; continue; }
    const r = renameInSref(t.text, tableLc, oldLc, newName, hostTableLc);
    if (r === null || r === t.text) { out += t.text; continue; }
    out += r;
    changed = true;
  }
  return changed ? out : formula;
}

function renameInSref(text: string, tableLc: string, oldLc: string, newName: string, hostTableLc: string | null): string | null {
  const lb = text.indexOf('[');
  if (lb < 0 || !text.endsWith(']')) return null;
  const table = text.slice(0, lb);
  const tname = table ? table.toLowerCase() : hostTableLc;
  if (tname !== tableLc) return null;
  const inner = text.slice(lb + 1, -1);
  const unesc2 = (s: string) => s.replace(/'(.)/gs, '$1');
  // simple form Table[name] / Table[@name] (no nested brackets)
  let nested = false;
  for (let k = 0; k < inner.length; k++) {
    if (inner[k] === "'") { k++; continue; }
    if (inner[k] === '[') { nested = true; break; }
  }
  if (!nested) {
    const at = inner.startsWith('@') ? '@' : '';
    const raw = inner.slice(at.length);
    if (!raw || raw.startsWith('#') || unesc2(raw).toLowerCase() !== oldLc) return null;
    // the FILE form keeps the single-bracket item whatever the characters (golden h-special: tblMain[a-b'#c])
    return `${table}[${at}${escapeSrefItem(newName)}]`;
  }
  // list form Table[[#This Row],[a]:[b]] / Table[@[a]]
  let res = '';
  let hit = false;
  let k = 0;
  while (k < inner.length) {
    const ch = inner[k];
    if (ch !== '[') { res += ch; k++; continue; }
    let e = k + 1;
    for (; e < inner.length; e++) { if (inner[e] === "'") { e++; continue; } if (inner[e] === ']') break; }
    const raw = inner.slice(k + 1, e);
    if (raw && !raw.startsWith('#') && unesc2(raw).toLowerCase() === oldLc) {
      res += `[${escapeSrefItem(newName)}]`;
      hit = true;
    } else res += inner.slice(k, e + 1);
    k = e + 1;
  }
  return hit ? `${table}[${res}]` : null;
}

/** Sqref text "A1:B2 C3" -> areas (whole-column sqrefs stay as areas). Invalid parts are dropped. */
export function parseSqref(sqref: string): Area[] {
  const out: Area[] = [];
  for (const part of sqref.trim().split(/\s+/)) {
    if (!part) continue;
    const t = tokenize(part);
    if (t.length === 1 && t[0].type === 'ref' && t[0].area && !t[0].prefix) out.push(t[0].area);
  }
  return out;
}

export function renderSqref(areas: Area[]): string {
  return areas.map((a) => renderArea(stripAbs(a))).join(' ');
}

/** sqref / ref attributes never carry '$'. */
export function stripAbs(a: Area): Area {
  if (!a.ac1 && !a.ac2 && !a.ar1 && !a.ar2) return a;
  return { ...a, ac1: false, ac2: false, ar1: false, ar2: false };
}

/** Adjust a plain area (sqref / ref attribute semantics) through entries of one sheet. null = deleted. */
export function adjustPlainArea(a: Area, ops: readonly OpEntry[], sheet: string, from = 0): Area | null {
  const sl = sheet.toLowerCase();
  let cur: Area | null = a;
  for (let i = from; i < ops.length && cur; i++) {
    if (ops[i].sheet.toLowerCase() !== sl) continue;
    cur = adjustAreaEntry(cur, ops[i], 'cell');
  }
  return cur;
}
