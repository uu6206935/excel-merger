/**
 * Shared strings / inline strings scanner (owner: reader R). String scanning, no DOM; handles prefixed
 * (Strict / <x:sst>) documents, rich text runs (<r><rPr>…<t>) and ignores phonetic runs (<rPh>, ふりがな).
 */
import type { XStyle } from '../types';
import { DEFAULT_PALETTE, type Palette } from './theme';
import { fontStyle, parseFont } from './styles';
import { decodeXmlText, parseXml, rootOf } from './xml';

export type RichRuns = Array<{ text: string; style?: XStyle }>;

export interface SharedStrings {
  texts: string[];
  /** sst index -> runs, only for rich strings (the array object is shared by every cell using the string) */
  rich: Map<number, RichRuns>;
}

export const EMPTY_SST: SharedStrings = { texts: [], rich: new Map() };

/** Element-name prefix of the document's root element ("x:" for <x:sst>, '' for the default namespace). */
export function rootPrefix(xml: string): string {
  let i = 0;
  for (;;) {
    const lt = xml.indexOf('<', i);
    if (lt < 0) return '';
    const c = xml.charCodeAt(lt + 1);
    if (c === 63 /* ? */ || c === 33 /* ! */) {
      i = lt + 2;
      continue;
    }
    let k = lt + 1;
    while (k < xml.length) {
      const ch = xml.charCodeAt(k);
      if (ch === 32 || ch === 62 || ch === 47 || ch === 9 || ch === 10 || ch === 13) break;
      k++;
    }
    const qn = xml.slice(lt + 1, k);
    const colon = qn.indexOf(':');
    return colon < 0 ? '' : qn.slice(0, colon + 1);
  }
}

function escapeRe(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/** Scanner for the content of an <si> / <is> element with a given element prefix. */
export class StringItemScanner {
  private readonly p: string;
  private readonly reRPh: RegExp;
  private readonly reT: RegExp;
  private readonly reRun: RegExp;
  private readonly reRPr: RegExp;
  private readonly runOpen1: string;
  private readonly runOpen2: string;
  private readonly pal: Palette;
  private readonly styleCache = new Map<string, XStyle | undefined>();

  constructor(prefix: string, pal: Palette = DEFAULT_PALETTE) {
    this.p = prefix;
    this.pal = pal;
    const p = escapeRe(prefix);
    this.reRPh = new RegExp(`<${p}rPh[\\s>][\\s\\S]*?</${p}rPh>`, 'g');
    // attribute runs exclude '/' so that a self-closing <t/> never matches the content alternative
    this.reT = new RegExp(`<${p}t(?:\\s[^>/]*)?(?:/>|>([\\s\\S]*?)</${p}t>)`, 'g');
    this.reRun = new RegExp(`<${p}r(?:\\s[^>/]*)?>([\\s\\S]*?)</${p}r>`, 'g');
    this.reRPr = new RegExp(`<${p}rPr(?:\\s[^>/]*)?(?:/>|>[\\s\\S]*?</${p}rPr>)`);
    this.runOpen1 = `<${prefix}r>`;
    this.runOpen2 = `<${prefix}r `;
  }

  /** Plain text of the item (runs concatenated, phonetic runs dropped). */
  private texts(s: string): string {
    // fast path: <t>text</t> only
    const p = this.p;
    const open = `<${p}t>`;
    const close = `</${p}t>`;
    if (s.startsWith(open) && s.endsWith(close) && s.indexOf('<', open.length) === s.length - close.length) {
      return decodeXmlText(s.slice(open.length, s.length - close.length));
    }
    let out = '';
    this.reT.lastIndex = 0;
    let m: RegExpExecArray | null;
    while ((m = this.reT.exec(s))) if (m[1]) out += decodeXmlText(m[1]);
    return out;
  }

  /** Scan the inner XML of an <si>/<is>: plain text plus runs when the item is rich text. */
  scan(inner: string): { text: string; rich?: RichRuns } {
    let s = inner;
    if (s.indexOf(`<${this.p}rPh`) >= 0) s = s.replace(this.reRPh, '');
    if (s.indexOf(this.runOpen1) < 0 && s.indexOf(this.runOpen2) < 0) return { text: this.texts(s) };
    const runs: RichRuns = [];
    let text = '';
    this.reRun.lastIndex = 0;
    let m: RegExpExecArray | null;
    while ((m = this.reRun.exec(s))) {
      const body = m[1];
      const rPr = this.reRPr.exec(body);
      const t = this.texts(rPr ? body.slice(0, rPr.index) + body.slice(rPr.index + rPr[0].length) : body);
      text += t;
      const run: { text: string; style?: XStyle } = { text: t };
      if (rPr) {
        const st = this.runStyle(rPr[0]);
        if (st) run.style = st;
      }
      runs.push(run);
    }
    return { text, rich: runs };
  }

  private runStyle(rPrXml: string): XStyle | undefined {
    if (this.styleCache.has(rPrXml)) return this.styleCache.get(rPrXml);
    let st: XStyle | undefined;
    try {
      const el = rootOf(parseXml(rPrXml));
      if (el) {
        st = fontStyle(parseFont(el), this.pal);
        if (Object.keys(st).length === 0) st = undefined;
      }
    } catch {
      st = undefined;
    }
    this.styleCache.set(rPrXml, st);
    return st;
  }
}

/** Parse xl/sharedStrings.xml. Never throws; a missing / broken part yields an empty table. */
export function parseSharedStrings(xml: string | null | undefined, pal: Palette = DEFAULT_PALETTE): SharedStrings {
  if (!xml) return { texts: [], rich: new Map() };
  const p = rootPrefix(xml);
  const scanner = new StringItemScanner(p, pal);
  const texts: string[] = [];
  const rich = new Map<number, RichRuns>();
  const open = `<${p}si`;
  const close = `</${p}si>`;
  let pos = 0;
  for (;;) {
    const lt = xml.indexOf(open, pos);
    if (lt < 0) break;
    const after = xml.charCodeAt(lt + open.length);
    if (after !== 62 /* > */ && after !== 32 && after !== 47 /* / */ && after !== 9 && after !== 10 && after !== 13) {
      pos = lt + open.length; // e.g. <sit…> — not ours
      continue;
    }
    const gt = xml.indexOf('>', lt);
    if (gt < 0) break;
    if (xml.charCodeAt(gt - 1) === 47) {
      texts.push('');
      pos = gt + 1;
      continue;
    }
    const end = xml.indexOf(close, gt);
    const inner = xml.slice(gt + 1, end < 0 ? xml.length : end);
    const r = scanner.scan(inner);
    if (r.rich) rich.set(texts.length, r.rich);
    texts.push(r.text);
    if (end < 0) break;
    pos = end + close.length;
  }
  return { texts, rich };
}
