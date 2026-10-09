/**
 * Minimal, dependency-free XML helpers for OOXML parts (owner: reader R).
 *
 * - No DOMParser / XMLSerializer: runs in the browser (main thread or worker) and in Node.
 * - `parseXml` builds a small element tree; use it for SMALL parts only (workbook.xml, rels,
 *   [Content_Types].xml, styles.xml, theme1.xml, tables, the non-<sheetData> part of a sheet).
 *   Large parts (sheetData, sharedStrings) are scanned by dedicated string scanners (sheet.ts, strings.ts).
 * - Namespace prefixes are kept in `name` and stripped in `local`, so Strict OOXML / prefixed
 *   documents (<x:worksheet xmlns:x="...">) are handled by matching on `local`.
 *
 * Exported for reuse by other xlsx modules (the patcher may import these):
 *   decodeEntities, decodeOoxmlEscapes, decodeXmlText, escapeXmlText, escapeXmlAttr, encodeOoxmlEscapes,
 *   parseXml, kid, kids, attr, relId, descendants, innerText, localName.
 */

// ---------------------------------------------------------------------------------------------
// Entities and _xHHHH_ escapes
// ---------------------------------------------------------------------------------------------

const NAMED: Record<string, string> = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };

/** Decode the 5 predefined XML entities and numeric character references (&#n; &#xh;). Unknown entities are kept verbatim. */
export function decodeEntities(s: string): string {
  if (s.indexOf('&') < 0) return s;
  return s.replace(/&(#[xX][0-9a-fA-F]+|#[0-9]+|amp|lt|gt|quot|apos);/g, (m, e: string) => {
    if (e.charCodeAt(0) === 35 /* # */) {
      const code = e[1] === 'x' || e[1] === 'X' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
      if (!Number.isFinite(code) || code < 0 || code > 0x10ffff) return m;
      try {
        return String.fromCodePoint(code);
      } catch {
        return m;
      }
    }
    return NAMED[e] ?? m;
  });
}

/**
 * Decode OOXML ST_Xstring escapes: `_xHHHH_` is the UTF-16 code unit HHHH (Excel writes control characters such as
 * CR as `_x000D_`, and a literal "_x0041_" as "_x005F_x0041_"). Apply AFTER entity decoding.
 */
export function decodeOoxmlEscapes(s: string): string {
  if (s.indexOf('_x') < 0) return s;
  return s.replace(/_x([0-9A-Fa-f]{4})_/g, (_m, h: string) => String.fromCharCode(parseInt(h, 16)));
}

/** XML 1.0 end-of-line handling (§2.11): literal CR LF and lone CR become LF (character references are kept). */
export function normalizeEol(s: string): string {
  return s.indexOf('\r') < 0 ? s : s.replace(/\r\n?/g, '\n');
}

/** XML attribute-value normalisation (§3.3.3) for CDATA attributes: literal CR/LF/TAB become spaces. */
function normalizeAttr(s: string): string {
  return /[\r\n\t]/.test(s) ? s.replace(/\r\n|[\r\n\t]/g, ' ') : s;
}

/**
 * Raw XML text content of a string-bearing element (<t>, <v> of t="str", <f>) → value:
 * end-of-line normalisation, entities, then _xHHHH_ escapes (so a real CR written as _x000D_ survives).
 */
export function decodeXmlText(s: string): string {
  // markup inside the text: CDATA sections are literal text, comments / processing instructions are not text at all
  if (s.indexOf('<') >= 0 && (s.indexOf('<![CDATA[') >= 0 || s.indexOf('<!--') >= 0 || s.indexOf('<?') >= 0)) return decodeOoxmlEscapes(decodeMarkupText(s));
  return decodeOoxmlEscapes(decodeEntities(normalizeEol(s)));
}

function decodeMarkupText(s: string): string {
  let out = '';
  let i = 0;
  while (i < s.length) {
    const lt = s.indexOf('<', i);
    if (lt < 0) { out += decodeEntities(normalizeEol(s.slice(i))); break; }
    out += decodeEntities(normalizeEol(s.slice(i, lt)));
    if (s.startsWith('<![CDATA[', lt)) {
      const e = s.indexOf(']]>', lt + 9);
      out += normalizeEol(s.slice(lt + 9, e < 0 ? s.length : e));
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

/** Escape text for element content. */
export function escapeXmlText(s: string): string {
  return s.replace(/[&<>]/g, (c) => (c === '&' ? '&amp;' : c === '<' ? '&lt;' : '&gt;'));
}

/** Escape text for a double-quoted attribute value (also newlines/tabs so they survive attribute normalisation). */
export function escapeXmlAttr(s: string): string {
  return s.replace(/[&<>"\n\r\t]/g, (c) =>
    c === '&' ? '&amp;' : c === '<' ? '&lt;' : c === '>' ? '&gt;' : c === '"' ? '&quot;' : c === '\n' ? '&#10;' : c === '\r' ? '&#13;' : '&#9;',
  );
}

/**
 * Encode a string value for OOXML (inverse of decodeOoxmlEscapes): characters that XML 1.0 cannot carry
 * (C0 controls except TAB/LF, CR, U+FFFE/U+FFFF, lone surrogates) become `_xHHHH_`, and an existing literal
 * `_xHHHH_` gets its underscore escaped (`_x005F_`). The result still needs escapeXmlText.
 */
export function encodeOoxmlEscapes(s: string): string {
  let out = '';
  let last = 0;
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i);
    let rep: string | null = null;
    if ((c < 0x20 && c !== 0x09 && c !== 0x0a) || c === 0xfffe || c === 0xffff) rep = hex4(c);
    else if (c >= 0xd800 && c <= 0xdbff) {
      const d = s.charCodeAt(i + 1);
      if (d >= 0xdc00 && d <= 0xdfff) {
        i++;
        continue;
      }
      rep = hex4(c);
    } else if (c >= 0xdc00 && c <= 0xdfff) rep = hex4(c);
    else if (c === 0x5f /* _ */ && /^_x[0-9A-Fa-f]{4}_/.test(s.slice(i, i + 7))) rep = '_x005F_';
    if (rep !== null) {
      out += s.slice(last, i) + rep;
      last = i + 1;
    }
  }
  return last === 0 ? s : out + s.slice(last);
}

function hex4(c: number): string {
  return '_x' + c.toString(16).toUpperCase().padStart(4, '0') + '_';
}

// ---------------------------------------------------------------------------------------------
// Small tree parser
// ---------------------------------------------------------------------------------------------

export interface XEl {
  /** qualified name as written ("x:c", "a:srgbClr") */
  name: string;
  /** local name (prefix stripped) */
  local: string;
  /** attributes, qualified names as written, values entity-decoded */
  attrs: Record<string, string>;
  children: XEl[];
  /** concatenated DIRECT text content (entity-decoded, CDATA included); _xHHHH_ is NOT decoded */
  text: string;
  /** offsets of the element in the parsed string: xml.slice(start, end) is its raw XML (start tag .. end tag) */
  start: number;
  end: number;
}

export function localName(qname: string): string {
  const i = qname.indexOf(':');
  return i < 0 ? qname : qname.slice(i + 1);
}

function isSpace(c: number): boolean {
  return c === 32 || c === 9 || c === 10 || c === 13;
}

/**
 * Parse an XML document into a tree. Returns a synthetic '#document' element whose children hold the root.
 * Tolerant: never throws on malformed input (unclosed tags are closed at EOF, stray end tags ignored).
 * DOCTYPE declarations are skipped (no entity expansion — safe against billion-laughs / XXE).
 */
export function parseXml(xml: string): XEl {
  const n = xml.length;
  const doc: XEl = { name: '#document', local: '#document', attrs: {}, children: [], text: '', start: 0, end: n };
  const stack: XEl[] = [doc];
  let i = 0;
  if (xml.charCodeAt(0) === 0xfeff) i = 1;
  while (i < n) {
    const lt = xml.indexOf('<', i);
    const top = stack[stack.length - 1];
    if (lt < 0) {
      if (stack.length > 1) top.text += decodeEntities(normalizeEol(xml.slice(i)));
      break;
    }
    if (lt > i && stack.length > 1) top.text += decodeEntities(normalizeEol(xml.slice(i, lt)));
    const c1 = xml.charCodeAt(lt + 1);
    if (c1 === 33 /* ! */) {
      if (xml.startsWith('<!--', lt)) {
        const e = xml.indexOf('-->', lt + 4);
        i = e < 0 ? n : e + 3;
      } else if (xml.startsWith('<![CDATA[', lt)) {
        const e = xml.indexOf(']]>', lt + 9);
        if (stack.length > 1) top.text += normalizeEol(xml.slice(lt + 9, e < 0 ? n : e));
        i = e < 0 ? n : e + 3;
      } else {
        // <!DOCTYPE ... [ internal subset ] >
        let j = lt + 2;
        let depth = 0;
        for (; j < n; j++) {
          const ch = xml.charCodeAt(j);
          if (ch === 91) depth++;
          else if (ch === 93) depth--;
          else if (ch === 62 && depth <= 0) break;
        }
        i = j + 1;
      }
      continue;
    }
    if (c1 === 63 /* ? */) {
      const e = xml.indexOf('?>', lt + 2);
      i = e < 0 ? n : e + 2;
      continue;
    }
    if (c1 === 47 /* / */) {
      const e = xml.indexOf('>', lt + 2);
      const qn = xml.slice(lt + 2, e < 0 ? n : e).trim();
      // pop to the matching element (tolerant)
      const endPos = e < 0 ? n : e + 1;
      for (let k = stack.length - 1; k >= 1; k--) {
        if (stack[k].name === qn) {
          for (let m = k; m < stack.length; m++) stack[m].end = endPos;
          stack.length = k;
          break;
        }
      }
      i = endPos;
      continue;
    }
    // start tag: find '>' outside quotes
    let j = lt + 1;
    let q = 0;
    for (; j < n; j++) {
      const ch = xml.charCodeAt(j);
      if (q) {
        if (ch === q) q = 0;
      } else if (ch === 34 || ch === 39) q = ch;
      else if (ch === 62) break;
    }
    const selfClose = xml.charCodeAt(j - 1) === 47;
    const innerEnd = selfClose ? j - 1 : j;
    let k = lt + 1;
    while (k < innerEnd && !isSpace(xml.charCodeAt(k))) k++;
    const qname = xml.slice(lt + 1, k);
    const el: XEl = { name: qname, local: localName(qname), attrs: parseAttrString(xml, k, innerEnd), children: [], text: '', start: lt, end: Math.min(n, j + 1) };
    top.children.push(el);
    if (!selfClose) stack.push(el);
    i = j + 1;
  }
  // elements left open at EOF end there
  for (let m = 1; m < stack.length; m++) stack[m].end = n;
  return doc;
}

/** Prefix ("x:") of a qualified element name, '' when unprefixed. */
export function prefixOf(qname: string): string {
  const i = qname.indexOf(':');
  return i < 0 ? '' : qname.slice(0, i + 1);
}

/** Parse `name="value"` pairs in xml[from, to). Values are entity-decoded. */
export function parseAttrString(xml: string, from: number, to: number): Record<string, string> {
  const out: Record<string, string> = {};
  let i = from;
  while (i < to) {
    while (i < to && isSpace(xml.charCodeAt(i))) i++;
    if (i >= to) break;
    const ns = i;
    while (i < to) {
      const c = xml.charCodeAt(i);
      if (c === 61 /* = */ || isSpace(c)) break;
      i++;
    }
    const name = xml.slice(ns, i);
    while (i < to && isSpace(xml.charCodeAt(i))) i++;
    if (xml.charCodeAt(i) !== 61) {
      // attribute without value (not XML, be tolerant)
      if (name) out[name] = '';
      continue;
    }
    i++;
    while (i < to && isSpace(xml.charCodeAt(i))) i++;
    const q = xml.charCodeAt(i);
    if (q !== 34 && q !== 39) {
      const vs = i;
      while (i < to && !isSpace(xml.charCodeAt(i))) i++;
      out[name] = decodeEntities(normalizeAttr(xml.slice(vs, i)));
      continue;
    }
    const vs = i + 1;
    let ve = xml.indexOf(String.fromCharCode(q), vs);
    if (ve < 0 || ve > to) ve = to;
    out[name] = decodeEntities(normalizeAttr(xml.slice(vs, ve)));
    i = ve + 1;
  }
  return out;
}

/** The document's root element (first element child of the '#document' node), or undefined. */
export function rootOf(doc: XEl): XEl | undefined {
  return doc.children[0];
}

/** First child with this local name. */
export function kid(el: XEl | undefined, local: string): XEl | undefined {
  if (!el) return undefined;
  for (const c of el.children) if (c.local === local) return c;
  return undefined;
}

/** All children with this local name. */
export function kids(el: XEl | undefined, local: string): XEl[] {
  if (!el) return [];
  const out: XEl[] = [];
  for (const c of el.children) if (c.local === local) out.push(c);
  return out;
}

/** All descendants (depth-first, document order) with this local name. */
export function descendants(el: XEl | undefined, local: string, out: XEl[] = []): XEl[] {
  if (!el) return out;
  for (const c of el.children) {
    if (c.local === local) out.push(c);
    descendants(c, local, out);
  }
  return out;
}

/**
 * Attribute by name: exact qualified match first ("r:id", "name"), then — only when `name` has no prefix —
 * an unprefixed match is required (a prefixed attribute is NOT returned for an unprefixed name, so `id`
 * never picks up `r:id`). Use relId() for relationship ids.
 */
export function attr(el: XEl | undefined, name: string): string | undefined {
  if (!el) return undefined;
  return el.attrs[name];
}

/** Relationship id attribute (r:id with any prefix bound to the relationships namespace; usually "r:id"). */
export function relId(el: XEl | undefined): string | undefined {
  if (!el) return undefined;
  const direct = el.attrs['r:id'];
  if (direct !== undefined) return direct;
  for (const k of Object.keys(el.attrs)) if (k.endsWith(':id')) return el.attrs[k];
  return undefined;
}

/** Concatenated text of the element and all its descendants (document order). */
export function innerText(el: XEl | undefined): string {
  if (!el) return '';
  if (el.children.length === 0) return el.text;
  let s = el.text;
  for (const c of el.children) s += innerText(c);
  return s;
}

/** Parse a boolean OOXML attribute ('1'/'true' → true, '0'/'false' → false, absent → dflt). */
export function boolAttr(v: string | undefined, dflt = false): boolean {
  if (v === undefined) return dflt;
  return v === '1' || v === 'true' || v === 'on';
}

/** Parse a numeric attribute; undefined when absent or not a finite number. */
export function numAttr(v: string | undefined): number | undefined {
  if (v === undefined || v === '') return undefined;
  const n = Number(v);
  return Number.isFinite(n) ? n : undefined;
}
