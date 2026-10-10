/**
 * Formulas of a sheet that is copied into another workbook (addsheets.ts): every reference to a sheet of the SAME source file is
 * re-pointed at the renamed copy, everything that would reach outside the copies (external workbooks, 3D references, sheets that are not
 * copied, tables / names that are not copied, call-out functions) is reported so that the caller can keep the cached value instead.
 *
 * Rewriting works on tokens (formula/excelRef.ts tokenize) — never by regular expressions on formula text. A formula is only tokenized
 * when it can contain something to rewrite (a "!" prefix, a "[" of an external book / structured reference, a defined name of the
 * source, INDIRECT / HYPERLINK strings); the other 99 % of the formulas of a big sheet pass through untouched.
 */
import { isUnsafeFormula } from '../../export/comJob';
import { parseStructuredRef, quoteSheetName, tokenize } from '../../formula/excelRef';

export type CopyFailReason =
  /** DDE, WEBSERVICE, a path / URL reference ... (isUnsafeFormula) */
  | 'unsafe'
  /** a reference into another workbook ([1]Sheet1!A1) */
  | 'external'
  /** Sheet1:Sheet3!A1 */
  | 'threeD'
  /** a sheet that is not among the copied sheets (chart sheet, deleted sheet) */
  | 'missingSheet'
  /** a table (structured reference) that is not copied */
  | 'table'
  /** a defined name that cannot be copied */
  | 'name'
  /** a formula the tokenizer cannot read */
  | 'syntax';

export type Rewrite = { ok: true; text: string; changed: boolean; names: string[] } | { ok: false; reason: CopyFailReason; detail?: string };

/** the defined names as seen from one host sheet (a sheet-scoped name hides a workbook name of the same text) */
export interface NameView {
  /** the (lower-cased) name is defined for this sheet */
  has(lower: string): boolean;
  /** it is defined but cannot be copied (external, ...) */
  bad(lower: string): boolean;
  /** case-insensitive RegExp of all the names (a cheap "does this formula mention a name" test); null = no names */
  pattern: RegExp | null;
}

export const NO_NAMES: NameView = { has: () => false, bad: () => false, pattern: null };

export interface CopierOptions {
  /** lower-cased name of a source sheet -> its name in the output (the worksheets that are copied) */
  sheets: ReadonlyMap<string, string>;
  /** lower-cased table name of the source -> the name of its copy; tables that are not in this map are not copied */
  tables?: ReadonlyMap<string, string>;
  /** lower-cased names of ALL tables of the source (a bare table name or a Table[...] reference of an uncopied table makes a formula a value) */
  allTables?: ReadonlySet<string>;
  /** lower-cased names of the source sheets whose tables are copied: there the table-less form `[@col]` stays valid */
  sheetsWithTables?: ReadonlySet<string>;
}

const NEEDS = /[!\[]|INDIRECT|HYPERLINK/i;
const DANGER_QUICK = /[|\[\]\\]|webservice|image|rtd|call|register|exec/i;
const NON_ASCII = /[^\x00-\x7f]/;
const INDIRECT_FN = /^(?:_xlfn\.)*(?:INDIRECT|HYPERLINK)$/i;
const EMPTY: string[] = [];

/**
 * isUnsafeFormula with a cheap pre-test: the formula has none of the characters / words the predicate looks for (and no character that NFKC
 * folding could turn into one), so the answer is "safe" without normalizing and scanning it.
 */
export function maybeUnsafe(f: string): boolean {
  if (DANGER_QUICK.test(f)) return true;
  if (f.indexOf("'") >= 0 && f.indexOf('/') >= 0) return true;
  return NON_ASCII.test(f);
}

export function unsafeFormula(f: string): boolean {
  return maybeUnsafe(f) && isUnsafeFormula(f);
}

const escapeRe = (s: string): string => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

/**
 * The RegExp of a list of names (null when empty): the name as a whole word (not inside a longer identifier, SUBTOTAL does not mention a
 * name "total"). A cheap pre-test only: the tokens decide.
 */
export function namePattern(names: Iterable<string>): RegExp | null {
  const list = [...new Set(names)].filter((n) => n !== '').sort((a, b) => b.length - a.length).map(escapeRe);
  return list.length ? new RegExp(`(?<![\\p{L}\\p{N}_.\\\\])(?:${list.join('|')})(?![\\p{L}\\p{N}_.])`, 'iu') : null;
}

export class FormulaCopier {
  constructor(readonly opts: CopierOptions) {}

  /**
   * `hostSheet`: the lower-cased source name of the sheet the formula lives on (null: a workbook-level name). `names`: the defined names
   * of that sheet. A successful result carries the lower-cased defined names the formula uses (they have to exist in the copy).
   */
  rewrite(f: string, hostSheet: string | null, names: NameView = NO_NAMES): Rewrite {
    if (unsafeFormula(f)) return { ok: false, reason: 'unsafe' };
    if (!NEEDS.test(f) && !(names.pattern && names.pattern.test(f)) && !(this.opts.allTables?.size && f.indexOf('[') >= 0)) {
      return { ok: true, text: f, changed: false, names: EMPTY };
    }
    const toks = tokenize(f);
    const parts: string[] = new Array(toks.length);
    const used = new Set<string>();
    let changed = false;
    let indirect = false;
    for (let i = 0; i < toks.length; i++) {
      const t = toks[i];
      parts[i] = t.text;
      if (t.type === 'unknown') return { ok: false, reason: 'syntax' };
      const p = t.prefix;
      if (p) {
        if (p.book !== undefined) return { ok: false, reason: 'external' };
        if (p.sheet2 !== undefined) return { ok: false, reason: 'threeD' };
        if (p.sheet != null) {
          const nn = this.opts.sheets.get(p.sheet.toLowerCase());
          if (nn === undefined) return { ok: false, reason: 'missingSheet', detail: p.sheet };
          if (t.type === 'sref') {
            // Sheet1!Table1[col]: a structured reference never carries a sheet in files written by Excel
            return { ok: false, reason: 'table', detail: t.text };
          }
          parts[i] = quoteSheetName(nn) + '!' + t.text.slice(p.text.length);
          if (parts[i] !== t.text) changed = true;
        }
        if (t.type === 'func' && INDIRECT_FN.test(t.text.slice(p.text.length))) indirect = true;
        continue;
      }
      if (t.type === 'name') {
        const l = t.text.toLowerCase();
        if (this.opts.allTables?.has(l)) {
          const nt = this.opts.tables?.get(l);
          if (nt === undefined) return { ok: false, reason: 'table', detail: t.text };
          if (nt !== t.text) { parts[i] = nt; changed = true; }
        } else if (names.has(l)) {
          if (names.bad(l)) return { ok: false, reason: 'name', detail: t.text };
          used.add(l);
        }
      } else if (t.type === 'func') {
        if (INDIRECT_FN.test(t.text)) indirect = true;
        const l = t.text.replace(/^(?:_xlfn\.)+/i, '').toLowerCase();
        if (names.has(l)) {
          if (names.bad(l)) return { ok: false, reason: 'name', detail: t.text };
          used.add(l);
        }
      } else if (t.type === 'sref') {
        const info = parseStructuredRef(t.text);
        if (!info) return { ok: false, reason: 'syntax' };
        if (info.table) {
          const nt = this.opts.tables?.get(info.table.toLowerCase());
          if (nt === undefined) return { ok: false, reason: 'table', detail: info.table };
          if (nt !== info.table) { parts[i] = nt + t.text.slice(info.table.length); changed = true; }
        } else if (!(hostSheet && this.opts.sheetsWithTables?.has(hostSheet))) {
          return { ok: false, reason: 'table', detail: t.text };
        }
      }
    }
    if (indirect) {
      // INDIRECT("Sheet2!A1") / HYPERLINK("#Sheet2!A1"): a text that is exactly one reference to a sheet of the source follows its copy
      for (let i = 0; i < toks.length; i++) {
        if (toks[i].type !== 'string') continue;
        const r = this.rewriteRefString(toks[i].text);
        if (r !== null && r !== parts[i]) { parts[i] = r; changed = true; }
      }
    }
    return { ok: true, text: changed ? parts.join('') : f, changed, names: used.size ? [...used] : EMPTY };
  }

  /** a string literal ("Sheet2!A1" or "#Sheet2!A1") whose whole content is one reference to a copied sheet, re-pointed; else null */
  private rewriteRefString(lit: string): string | null {
    if (lit.length < 4) return null;
    const body = lit.slice(1, -1).replace(/""/g, '"');
    const hash = body.startsWith('#') ? '#' : '';
    const rest = body.slice(hash.length);
    if (rest.indexOf('!') < 0) return null;
    const toks = tokenize(rest);
    if (toks.length !== 1) return null;
    const t = toks[0];
    const p = t.prefix;
    if (!p || p.book !== undefined || p.sheet2 !== undefined || p.sheet == null || (t.type !== 'ref' && t.type !== 'name')) return null;
    const nn = this.opts.sheets.get(p.sheet.toLowerCase());
    if (nn === undefined) return null;
    const next = hash + quoteSheetName(nn) + '!' + t.text.slice(p.text.length);
    return '"' + next.replace(/"/g, '""') + '"';
  }
}

/** Japanese words for the reasons (warning texts) */
export const FAIL_JA: Record<CopyFailReason, string> = {
  unsafe: '他のブック・外部を参照または呼び出す数式',
  external: '他のブックを参照している数式',
  threeD: '複数シートにまたがる参照（3D 参照）の数式',
  missingSheet: 'コピーしないシート（グラフシートなど）を参照している数式',
  table: 'テーブル参照の数式',
  name: 'コピーできない名前を使っている数式',
  syntax: '解釈できない書き方の数式',
};
