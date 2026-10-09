/**
 * Output safety net of the package patcher (owner: patcher P): a fast well-formedness check for the XML parts the
 * patcher wrote. The patcher edits XML as text; a hostile or damaged input (truncated sheet, raw '&', unbalanced tags,
 * control characters, a namespace prefix that lost its declaration ...) must end as a PatchError, never as a package
 * that Excel "repairs" or refuses. Linear time, no DOM.
 *
 * Checked: one root element, balanced tags with matching names, attributes written as name="value" (no '<' inside,
 * no duplicates on elements other than <c> / <row>), entity references (only the five predefined ones and numeric
 * ones), characters allowed by XML 1.0 (no control characters, no lone surrogates), namespace prefixes declared in
 * scope, no DOCTYPE. Comments, CDATA sections and processing instructions are skipped.
 */

// eslint-disable-next-line no-control-regex
const BAD_CHAR = /[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]|[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/;
// eslint-disable-next-line no-control-regex
const CHAR_CANDIDATE = /[\u0000-\u0008\u000B\u000C\u000E-\u001F\uD800-\uDFFF￾￿]/;
const BAD_AMP =/&(?!(?:amp|lt|gt|quot|apos|#[0-9]+|#[xX][0-9a-fA-F]+);)/;
const TAG = /<(\/?)([^\s/>=<"'!?&]+)([^>]*?)(\/?)>/y;
const ATTRS_OK = /^(?:\s+[^\s=<>"'/&]+\s*=\s*(?:"[^"<]*"|'[^'<]*'))*\s*$/;
const ATTR_NAME = /\s([^\s=]+)\s*=\s*(?:"[^"]*"|'[^']*')/g;
const ATTR_ANY = /\s([^\s=]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;

/** null = well-formed, else a short description of the first problem found. */
export function xmlProblem(xml: string): string | null {
  // the exact test has look-arounds and is slow on a 50 MB sheet: the cheap class test rules out the (usual) clean text first
  const bad = CHAR_CANDIDATE.test(xml) ? BAD_CHAR.exec(xml) : null;
  if (bad) return `使用できない文字（U+${bad[0].charCodeAt(0).toString(16).toUpperCase().padStart(4, '0')}）`;
  const n = xml.length;
  const stack: string[] = [];
  /** prefixes declared by each open element (null = none) */
  const decls: Array<string[] | null> = [];
  const active = new Map<string, number>([['xml', 1], ['xmlns', 1]]);
  let roots = 0;
  let i = 0;
  let nextAmp = -2;
  while (i < n) {
    const lt = xml.indexOf('<', i);
    const textEnd = lt < 0 ? n : lt;
    if (textEnd > i) {
      if (!stack.length) {
        if (/\S/.test(xml.slice(i, textEnd))) return '文書の外側に文字があります';
      } else {
        if (nextAmp !== -1 && nextAmp < i) nextAmp = xml.indexOf('&', i);
        if (nextAmp >= 0 && nextAmp < textEnd && BAD_AMP.test(xml.slice(i, textEnd))) return `不正な「&」（${around(xml, nextAmp)}）`;
      }
    }
    if (lt < 0) break;
    const c1 = xml.charCodeAt(lt + 1);
    if (c1 === 33 /* ! */) {
      if (xml.startsWith('<!--', lt)) {
        const e = xml.indexOf('-->', lt + 4);
        if (e < 0) return 'コメントが閉じていません';
        i = e + 3;
        continue;
      }
      if (xml.startsWith('<![CDATA[', lt)) {
        const e = xml.indexOf(']]>', lt + 9);
        if (e < 0) return 'CDATA が閉じていません';
        if (!stack.length) return '文書の外側に CDATA があります';
        i = e + 3;
        continue;
      }
      return `使用できない宣言（${around(xml, lt)}）`;
    }
    if (c1 === 63 /* ? */) {
      const e = xml.indexOf('?>', lt + 2);
      if (e < 0) return '処理命令が閉じていません';
      if (stack.length || roots) return '処理命令の位置が不正です';
      i = e + 2;
      continue;
    }
    TAG.lastIndex = lt;
    const m = TAG.exec(xml);
    if (!m) return `タグが正しくありません（${around(xml, lt)}）`;
    i = lt + m[0].length;
    const name = m[2];
    const attrs = m[3];
    if (m[1] === '/') {
      if (attrs.trim() !== '' || m[4]) return `終了タグが正しくありません（</${name}>）`;
      const top = stack.pop();
      if (top === undefined) return `対応する開始タグがありません（</${name}>）`;
      if (top !== name) return `タグが対応していません（<${top}> が </${name}> で閉じています）`;
      const d = decls.pop();
      if (d) for (const p of d) active.set(p, (active.get(p) ?? 1) - 1);
      continue;
    }
    // start tag (or empty element)
    if (!stack.length && ++roots > 1) return '文書に複数のルート要素があります';
    if (attrs.length) {
      if (name === 'c' || name === 'row') {
        // the bulk of a sheet: attribute syntax only (duplicates would come from the input, see SheetModel)
        if (!ATTRS_OK.test(attrs)) return `属性が正しくありません（<${name}${attrs.slice(0, 60)}>）`;
      } else {
        if (!ATTRS_OK.test(attrs)) return `属性が正しくありません（<${name}${attrs.slice(0, 60)}>）`;
        const seen = new Set<string>();
        for (const a of attrs.matchAll(ATTR_NAME)) {
          if (seen.has(a[1])) return `属性が重複しています（<${name}> の ${a[1]}）`;
          seen.add(a[1]);
        }
      }
      if (attrs.indexOf('&') >= 0) {
        for (const a of attrs.matchAll(ATTR_ANY)) {
          const v = a[2] ?? a[3] ?? '';
          if (v.indexOf('&') >= 0 && BAD_AMP.test(v)) return `属性値の「&」が不正です（<${name}> の ${a[1]}）`;
        }
      }
    }
    let declared: string[] | null = null;
    if (attrs.indexOf('xmlns:') >= 0) {
      for (const a of attrs.matchAll(/\sxmlns:([^\s=]+)\s*=/g)) {
        (declared ??= []).push(a[1]);
        active.set(a[1], (active.get(a[1]) ?? 0) + 1);
      }
    }
    let bad2: string | null = null;
    const col = name.indexOf(':');
    if (col > 0 && !(active.get(name.slice(0, col)))) bad2 = name.slice(0, col);
    if (!bad2 && attrs.indexOf(':') >= 0) {
      for (const a of attrs.matchAll(ATTR_NAME)) {
        const k = a[1].indexOf(':');
        if (k > 0 && !active.get(a[1].slice(0, k))) { bad2 = a[1].slice(0, k); break; }
      }
    }
    if (bad2) return `名前空間の接頭辞「${bad2}」が宣言されていません（<${name}>）`;
    if (m[4] === '/') {
      if (declared) for (const p of declared) active.set(p, (active.get(p) ?? 1) - 1);
    } else {
      stack.push(name);
      decls.push(declared);
    }
  }
  if (stack.length) return `タグが閉じていません（<${stack[stack.length - 1]}>）`;
  if (!roots) return '要素がありません';
  return null;
}

function around(xml: string, at: number): string {
  return xml.slice(Math.max(0, at - 10), at + 40).replace(/\s+/g, ' ');
}
