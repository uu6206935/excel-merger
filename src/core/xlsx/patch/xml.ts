/**
 * Minimal XML string helpers for the package patcher (owner: patcher P).
 * The patcher never re-serializes a whole part through a DOM: it rewrites only the nodes / attributes that
 * carry coordinates or references and copies every other byte verbatim.
 */

import { encodeOoxmlEscapes } from '../xml';

const ENT: Record<string, string> = { lt: '<', gt: '>', quot: '"', apos: "'", amp: '&' };

/**
 * Decode XML entities (named + numeric). A character reference that is not a character (&#99999999; beyond U+10FFFF,
 * a lone surrogate) is kept verbatim instead of throwing.
 */
export function unesc(s: string): string {
  // a writer that wraps text in <![CDATA[ ]]> (or leaves comments inside a <t> / <f>): the section is literal text
  if (s.indexOf('<') >= 0 && (s.indexOf('<![CDATA[') >= 0 || s.indexOf('<!--') >= 0 || s.indexOf('<?') >= 0)) return unescMarkup(s);
  if (s.indexOf('&') < 0) return s;
  return s.replace(/&(#x[0-9a-fA-F]+|#[0-9]+|lt|gt|quot|apos|amp);/g, (m, e: string) => {
    if (e[0] === '#') {
      const code = e[1] === 'x' || e[1] === 'X' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
      return Number.isFinite(code) && code >= 0 && code <= 0x10ffff && !(code >= 0xd800 && code <= 0xdfff) ? String.fromCodePoint(code) : m;
    }
    return ENT[e] ?? m;
  });
}

/** Text with CDATA sections (literal), comments and processing instructions (dropped) and entity references (decoded). */
function unescMarkup(s: string): string {
  let out = '';
  let i = 0;
  while (i < s.length) {
    const lt = s.indexOf('<', i);
    if (lt < 0) { out += unesc(s.slice(i)); break; }
    out += unesc(s.slice(i, lt));
    if (s.startsWith('<![CDATA[', lt)) {
      const e = s.indexOf(']]>', lt + 9);
      out += s.slice(lt + 9, e < 0 ? s.length : e);
      i = e < 0 ? s.length : e + 3;
    } else if (s.startsWith('<!--', lt)) {
      const e = s.indexOf('-->', lt + 4);
      i = e < 0 ? s.length : e + 3;
    } else if (s.startsWith('<?', lt)) {
      const e = s.indexOf('?>', lt + 2);
      i = e < 0 ? s.length : e + 2;
    } else {
      out += '<';
      i = lt + 1;
    }
  }
  return out;
}

/** Escape text content. */
export function esc(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/** Escape an attribute value (double quoted). */
export function escAttr(s: string): string {
  return esc(s).replace(/"/g, '&quot;').replace(/\r/g, '&#13;').replace(/\n/g, '&#10;').replace(/\t/g, '&#9;');
}

/**
 * Escape text for <t>/<v> content: also encodes characters that XML 1.0 cannot carry the way Excel does
 * (`_xHHHH_`): control characters, U+FFFE / U+FFFF, lone surrogates and \u2014 important \u2014 CR (a literal CR would be
 * read back as LF by every XML parser; Excel writes `_x000D_`). A literal "_x0041_" becomes "_x005F_x0041_".
 * Tab and LF stay literal. (Reader R's encodeOoxmlEscapes does the work.)
 */
export function escText(s: string): string {
  return esc(encodeOoxmlEscapes(s));
}

/** Decode Excel's _xHHHH_ escapes in shared-string / inline text. */
export function decodeXEsc(s: string): string {
  if (s.indexOf('_x') < 0) return s;
  return s.replace(/_x([0-9A-Fa-f]{4})_/g, (_m, h: string) => String.fromCharCode(parseInt(h, 16)));
}

/** Read one attribute from an attribute string / start tag. Returns the UNESCAPED value or undefined. */
const getRe = new Map<string, RegExp>();
const setRe = new Map<string, RegExp>();

export function getAttr(tagOrAttrs: string, name: string): string | undefined {
  if (tagOrAttrs.indexOf(name) < 0) return undefined;
  let re = getRe.get(name);
  if (!re) {
    re = new RegExp('(?:^|\\s)' + name.replace(/[.:]/g, (c) => '\\' + c) + '\\s*=\\s*("([^"]*)"|\'([^\']*)\')');
    getRe.set(name, re);
  }
  const m = re.exec(tagOrAttrs);
  if (!m) return undefined;
  return unesc(m[2] ?? m[3] ?? '');
}

/**
 * Set (or add / remove with value null) an attribute inside a start tag or attribute string.
 * When the attribute exists its position is kept; otherwise it is appended before the closing '>' or '/>'
 * (or at the end of a bare attribute string).
 */
export function setAttr(tagOrAttrs: string, name: string, value: string | null): string {
  let re = setRe.get(name);
  if (!re) {
    re = new RegExp('(\\s)' + name.replace(/[.:]/g, (c) => '\\' + c) + '\\s*=\\s*("[^"]*"|\'[^\']*\')');
    setRe.set(name, re);
  }
  const m = tagOrAttrs.indexOf(name) < 0 ? null : re.exec(tagOrAttrs);
  if (m) {
    if (value === null) return tagOrAttrs.slice(0, m.index) + tagOrAttrs.slice(m.index + m[0].length);
    return tagOrAttrs.slice(0, m.index) + m[1] + name + '="' + escAttr(value) + '"' + tagOrAttrs.slice(m.index + m[0].length);
  }
  if (value === null) return tagOrAttrs;
  const add = ' ' + name + '="' + escAttr(value) + '"';
  const t = tagOrAttrs;
  if (t.endsWith('/>')) return t.slice(0, -2).replace(/\s+$/, '') + add + '/>';
  if (t.endsWith('>')) return t.slice(0, -1).replace(/\s+$/, '') + add + '>';
  return t + add;
}

/** Find the first element <tag ...>...</tag> or <tag .../> ; returns indices and pieces. */
export interface ElemMatch {
  start: number;
  end: number;
  /** full start tag text, e.g. '<cols>' or '<pane xSplit="1"/>' */
  open: string;
  /** inner xml ('' for self-closing) */
  inner: string;
  selfClosing: boolean;
  text: string;
}

function tagRe(tag: string): string {
  return tag.replace(/[.:]/g, (c) => '\\' + c);
}

/** All (non-nested) elements with the given qualified tag name, in document order. */
export function findAll(xml: string, tag: string, from = 0, to = xml.length): ElemMatch[] {
  const out: ElemMatch[] = [];
  const t = tagRe(tag);
  const re = new RegExp('<' + t + '(?=[\\s/>])[^>]*>', 'g');
  re.lastIndex = from;
  const close = '</' + tag + '>';
  let m: RegExpExecArray | null;
  while ((m = re.exec(xml)) && m.index < to) {
    const open = m[0];
    const start = m.index;
    if (open.endsWith('/>')) {
      out.push({ start, end: start + open.length, open, inner: '', selfClosing: true, text: open });
      continue;
    }
    // handle nesting of the same tag name
    let depth = 1;
    let pos = start + open.length;
    const reAny = new RegExp('<(/?)' + t + '(?=[\\s/>])[^>]*>', 'g');
    reAny.lastIndex = pos;
    let e: RegExpExecArray | null;
    let endInner = -1;
    while ((e = reAny.exec(xml))) {
      if (e[1] === '/') {
        depth--;
        if (depth === 0) { endInner = e.index; break; }
      } else if (!e[0].endsWith('/>')) depth++;
    }
    if (endInner < 0) break; // malformed: stop
    const end = endInner + close.length;
    out.push({ start, end, open, inner: xml.slice(start + open.length, endInner), selfClosing: false, text: xml.slice(start, end) });
    re.lastIndex = end;
    pos = end;
  }
  return out;
}

export function findFirst(xml: string, tag: string, from = 0): ElemMatch | null {
  const all = findAll(xml, tag, from);
  return all.length ? all[0] : null;
}

/** Replace matches (from findAll) with new texts; `repl` returns the replacement text. */
export function replaceElems(xml: string, list: ElemMatch[], repl: (m: ElemMatch, i: number) => string): string {
  if (!list.length) return xml;
  let out = '';
  let pos = 0;
  list.forEach((m, i) => {
    out += xml.slice(pos, m.start) + repl(m, i);
    pos = m.end;
  });
  return out + xml.slice(pos);
}

/** Replace every element <tag> (non-nested scan) via callback. */
export function mapElems(xml: string, tag: string, repl: (m: ElemMatch) => string): string {
  return replaceElems(xml, findAll(xml, tag), repl);
}

/** Text content of a simple element <tag>text</tag> (unescaped), or undefined. */
export function elemText(xml: string, tag: string): string | undefined {
  const m = findFirst(xml, tag);
  return m ? unesc(m.inner) : undefined;
}

const MAIN_NS = ['http://schemas.openxmlformats.org/spreadsheetml/2006/main', 'http://purl.oclc.org/ooxml/spreadsheetml/main'];

/**
 * A part whose SpreadsheetML elements carry a namespace prefix (<x:worksheet xmlns:x="…/main"><x:sheetData><x:row>…,
 * written by some generators) is rewritten to the default-namespace form the patcher's string scanners understand
 * (<worksheet xmlns="…/main"><sheetData><row>…): the prefix is dropped from every element of that namespace. Returns
 *  - the text itself when the root element is not prefixed with the main namespace,
 *  - the rewritten text, or
 *  - null when the prefix cannot be removed safely (another default namespace on the root, the prefix declared again
 *    further down): the caller reports a PatchError instead of patching text it cannot read.
 */
export function stripMainPrefix(xml: string): string | null {
  const head = xml.length > 16384 ? xml.slice(0, 16384) : xml;
  const m = /^(?:\s|<\?[\s\S]*?\?>|<!--[\s\S]*?-->)*<([A-Za-z_][\w.-]*):[A-Za-z_][\w.-]*(\s[^>]*)?>/.exec(head);
  if (!m) return xml;
  const p = m[1];
  const attrs = m[2] ?? '';
  const pe = p.replace(/[.\-]/g, '\\$&');
  const decl = new RegExp('\\sxmlns:' + pe + '\\s*=\\s*"([^"]*)"').exec(attrs);
  if (!decl || !MAIN_NS.includes(decl[1])) return xml;
  const def = /\sxmlns\s*=\s*"([^"]*)"/.exec(attrs);
  if (def && def[1] !== decl[1]) return null;
  if (xml.split('xmlns:' + p + '=').length > 2) return null;
  let out = xml.replace(new RegExp('<(/?)' + pe + ':', 'g'), '<$1');
  out = out.replace(decl[0], def ? '' : ` xmlns="${decl[1]}"`);
  return out;
}

/**
 * The element prefix (with its ':' — '' for the default namespace) an XML document binds to namespace `uri`, or `fallback` when
 * the document never declares that namespace. Generators differ: Excel / XlsxWriter write xdr: c: v: x:, openpyxl writes drawings
 * and charts in the default namespace and VML as ns0: ns1: ns2:, other libraries pick any prefix.
 */
export function nsPrefix(xml: string, uri: string, fallback: string): string {
  const i = xml.indexOf(uri);
  if (i < 0) return fallback;
  const start = xml.lastIndexOf('xmlns', i);
  if (start < 0) return fallback;
  const m = /^xmlns(?::([A-Za-z_][\w.-]*))?\s*=\s*["']$/.exec(xml.slice(start, i));
  if (!m) return fallback;
  return m[1] ? m[1] + ':' : '';
}

/** RegExp source of a literal string */
export function reEsc(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\-]/g, '\\$&');
}

const REL_NS = ['http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'http://purl.oclc.org/ooxml/officeDocument/relationships'];

/**
 * A part whose relationship attributes carry a prefix other than `r` (<sheet xmlns:rel="…/relationships" rel:id="rId1"/>,
 * <drawing ns1:id=…/>: generators that let the XML library pick prefixes) is rewritten to the `r:` spelling the patcher's
 * scanners look for (getAttr(el, 'r:id')): the attributes `p:name=` inside start tags become `r:name=` and, when `r` is not
 * declared, the declaration follows. Returns the text unchanged when the root declares no other relationship prefix, and
 * null when `r` is already bound to another namespace (renaming would be ambiguous; the caller keeps the text as it is).
 */
export function normalizeRelPrefix(xml: string): string | null {
  const head = xml.length > 8192 ? xml.slice(0, 8192) : xml;
  const decls = [...head.matchAll(/\sxmlns:([A-Za-z_][\w.-]*)\s*=\s*"([^"]*)"/g)];
  const other = decls.filter((d) => REL_NS.includes(d[2]) && d[1] !== 'r').map((d) => d[1]);
  if (!other.length) return xml;
  const rDecl = decls.find((d) => d[1] === 'r');
  if (rDecl && !REL_NS.includes(rDecl[2])) return null;
  let out = xml;
  for (const p of other) {
    const pe = p.replace(/[.\-]/g, '\\$&');
    out = out.replace(/<(?![?!/])[^<>]*>/g, (tag) => (tag.indexOf(p + ':') < 0 ? tag : tag.replace(new RegExp('(\\s)' + pe + ':([A-Za-z_][\\w.-]*)(\\s*=)', 'g'), '$1r:$2$3')));
  }
  if (!rDecl) {
    const first = other[0].replace(/[.\-]/g, '\\$&');
    out = out.replace(new RegExp('\\sxmlns:' + first + '\\s*='), ' xmlns:r=');
  }
  return out;
}

/**
 * Attribute values written with single quotes (<c r='A1' s='3'>, valid XML that some generators emit) are rewritten with
 * double quotes: every scanner of the patcher reads name="value". Only start tags are touched (a double-quoted value that
 * happens to contain `k='v'` is skipped as a unit); the XML declaration and text content keep their quotes. Text that has
 * no single-quoted attribute comes back unchanged (same string).
 */
export function normalizeAttrQuotes(xml: string): string {
  if (!/=\s*'/.test(xml)) return xml;
  return xml.replace(/<(?![?!/])[^<>]*>/g, (tag) => (tag.indexOf("'") < 0
    ? tag
    : tag.replace(/(\s[^\s=/<>"']+\s*=\s*)(?:"[^"]*"|'([^']*)')/g, (m, head: string, v: string | undefined) => (v === undefined ? m : `${head}"${v.replace(/"/g, '&quot;')}"`))));
}

/** Decode a part's bytes (UTF-8 or UTF-16 with BOM). */
export function decodeBytes(b: Uint8Array): { text: string; utf16: boolean } {
  if (b.length >= 2 && b[0] === 0xff && b[1] === 0xfe) {
    return { text: new TextDecoder('utf-16le').decode(b.subarray(2)), utf16: true };
  }
  if (b.length >= 2 && b[0] === 0xfe && b[1] === 0xff) {
    return { text: new TextDecoder('utf-16be').decode(b.subarray(2)), utf16: true };
  }
  let s = new TextDecoder('utf-8').decode(b);
  if (s.charCodeAt(0) === 0xfeff) s = s.slice(1);
  return { text: s, utf16: false };
}
