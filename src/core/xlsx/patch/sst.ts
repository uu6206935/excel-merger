/**
 * Shared strings table (owner: patcher P): read plain text of existing items, append new strings (deduplicated
 * against existing plain items) and rich / phonetic items (deduplicated by their XML). The original <si> items are
 * never rewritten.
 */
import { normalizeEol } from '../xml';
import { decodeXEsc, escText, getAttr, setAttr, unesc } from './xml';

export class Sst {
  private items: string[] = [];
  /** inner XML of the original items (for copying rich text runs / phonetic properties) */
  private raws: string[] = [];
  private index: Map<string, number> | null = null;
  private richIndex: Map<string, number> | null = null;
  /** appended items: plain text, or the inner XML of a rich item */
  private added: Array<{ text: string; xml?: string }> = [];
  private addedRefs = 0;
  private xml: string;
  private baseCount = 0;

  constructor(xml: string) {
    this.xml = xml;
    const open = /<sst\b[^>]*>/.exec(xml);
    if (open) this.baseCount = Number(getAttr(open[0], 'count') ?? 0) || 0;
    // parse items: <si>...</si> (plain <t> or rich runs <r><t>)
    // (`<si />` with a space is what .NET writers emit for an empty item: it still counts as an item)
    const re = /<si\s*\/>|<si(?:\s[^>]*)?>([\s\S]*?)<\/si>/g;
    let m: RegExpExecArray | null;
    while ((m = re.exec(xml))) {
      this.items.push(m[1] === undefined ? '' : siText(m[1]));
      this.raws.push(m[1] ?? '');
    }
  }

  get size(): number {
    return this.items.length + this.added.length;
  }

  text(i: number): string {
    if (i < this.items.length) return this.items[i] ?? '';
    return this.added[i - this.items.length]?.text ?? '';
  }

  /** inner XML of item i when it carries rich text runs or phonetic properties, else null */
  richXml(i: number): string | null {
    const x = i < this.raws.length ? this.raws[i] : this.added[i - this.items.length]?.xml;
    return x && isRichOrPhonetic(x) ? x : null;
  }

  /** index of `s` (appended when not present) */
  add(s: string): number {
    if (!this.index) {
      this.index = new Map();
      this.items.forEach((t, i) => { if (!this.index!.has(t) && !isRichOrPhonetic(this.raws[i])) this.index!.set(t, i); });
    }
    this.addedRefs++;
    const ex = this.index.get(s);
    if (ex !== undefined) return ex;
    const i = this.items.length + this.added.length;
    this.added.push({ text: s });
    this.index.set(s, i);
    return i;
  }

  /** index of a rich / phonetic item given by its inner XML (deduplicated by XML) */
  addRich(innerXml: string): number {
    if (!this.richIndex) {
      this.richIndex = new Map();
      this.raws.forEach((x, i) => { if (!this.richIndex!.has(x)) this.richIndex!.set(x, i); });
    }
    this.addedRefs++;
    const ex = this.richIndex.get(innerXml);
    if (ex !== undefined) return ex;
    const i = this.items.length + this.added.length;
    this.added.push({ text: siText(innerXml), xml: innerXml });
    this.richIndex.set(innerXml, i);
    return i;
  }

  get changed(): boolean {
    return this.added.length > 0 || this.addedRefs > 0;
  }

  serialize(): string {
    if (!this.changed) return this.xml;
    const total = this.items.length + this.added.length;
    const add = this.added.map((it) => {
      if (it.xml !== undefined) return `<si>${it.xml}</si>`;
      const s = it.text;
      const sp = /^\s|\s$|\n/.test(s) ? ' xml:space="preserve"' : '';
      return `<si><t${sp}>${escText(s)}</t></si>`;
    }).join('');
    let x = this.xml;
    x = x.replace(/<sst\b[^>]*>/, (open) => {
      // setAttr keeps an attribute where it is (either quote style) or appends it
      const cnt = String(Math.max(this.baseCount + this.addedRefs, total));
      return setAttr(setAttr(open, 'count', cnt), 'uniqueCount', String(total));
    });
    if (/<sst\b[^>]*\/>\s*$/.test(x)) x = x.replace(/<sst\b([^>]*?)\s*\/>\s*$/, `<sst$1>${add}</sst>`);
    else x = x.replace(/<\/sst>\s*$/, add + '</sst>');
    return x;
  }
}

/** true when an <si> / <is> body has formatting runs (<r>) or phonetic data (<rPh>, <phoneticPr>) */
export function isRichOrPhonetic(inner: string | undefined): boolean {
  return !!inner && (inner.indexOf('<r>') >= 0 || inner.indexOf('<r ') >= 0 || inner.indexOf('<rPh') >= 0 || inner.indexOf('<phoneticPr') >= 0);
}

/** Plain text of an <si> / <is> body: concatenation of <t> texts, phonetic runs (<rPh>) excluded. */
export function siText(inner: string): string {
  const body = inner.replace(/<rPh\b[\s\S]*?<\/rPh>/g, '').replace(/<phoneticPr\b[^>]*\/>/g, '');
  let out = '';
  const re = /<t(?:\s[^>]*)?>([\s\S]*?)<\/t>|<t(?:\s[^>]*)?\/>/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(body))) if (m[1] !== undefined) out += m[1];
  // XML end-of-line handling first (a literal CR LF reads as LF; `_x000D_` and `&#13;` stay a real CR)
  return decodeXEsc(unesc(normalizeEol(out)));
}
