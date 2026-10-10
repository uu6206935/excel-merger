// Package-level ("XML") diff of two .xlsx files (owner: T) — the second view of tools/compare.mjs, e.g. the
// JS-engine output vs the Excel-engine output of the same op log. Both packages are parsed into a NORMALIZED model
// so that only real differences remain:
//   - cell styles are compared as RESOLVED styles (number format code, font, fill, border, alignment, protection,
//     named cell style), never by xf index; CF/DXF styles likewise
//   - strings are compared as text, never by sharedStrings index; inline strings == shared strings
//   - shared formulas (<f t="shared" si="…"/>) are expanded before comparing
//   - parts are reached through relationships (part file names / rIds do not matter)
// Known-benign differences are COUNTED in `filtered` (see FILTER_LABELS) instead of being reported.
// Pure JS (jszip + a tiny XML tokenizer); no Excel needed. Unit-tested in tests/qa/package-diff.test.ts.
import fs from 'node:fs';
import JSZip from 'jszip';
import { cellAddr, numToCol, parseCell, parseRange, shiftFormula } from './compare-core.mjs';

// ---------------------------------------------------------------------------------------------------------------
// XML
// ---------------------------------------------------------------------------------------------------------------

const ENT = { lt: '<', gt: '>', amp: '&', quot: '"', apos: "'" };

export function decodeXml(s) {
  if (s.indexOf('&') < 0) return s;
  return s.replace(/&(#x[0-9a-fA-F]+|#\d+|lt|gt|amp|quot|apos);/g, (_m, e) =>
    e[0] === '#' ? String.fromCodePoint(e[1] === 'x' || e[1] === 'X' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10)) : ENT[e],
  );
}

// comment | processing instruction | CDATA (1) | doctype | end tag (2) | start tag (3 name, 4 attrs, 5 self-closing) | text (6)
const TOKEN_RE =
  /<!--[\s\S]*?-->|<\?[\s\S]*?\?>|<!\[CDATA\[([\s\S]*?)\]\]>|<!DOCTYPE[^>]*>|<\/([^\s>]+)\s*>|<([^\s/>!?]+)((?:\s+[^\s=/>]+\s*=\s*(?:"[^"]*"|'[^']*'))*)\s*(\/?)>|([^<]+)/g;
const ATTR_RE = /([^\s=/>]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;

/** Minimal XML parser for machine-written OOXML parts: returns the document element {name, local, attrs, children, text}. */
export function parseXml(text) {
  const root = { name: '#root', local: '#root', attrs: {}, children: [], text: '' };
  const stack = [root];
  const re = new RegExp(TOKEN_RE.source, 'g');
  let m;
  while ((m = re.exec(text))) {
    const top = stack[stack.length - 1];
    if (m[1] !== undefined) top.text += m[1];
    else if (m[2] !== undefined) {
      if (stack.length > 1) stack.pop();
    } else if (m[3] !== undefined) {
      const attrs = {};
      if (m[4]) {
        const ar = new RegExp(ATTR_RE.source, 'g');
        let a;
        while ((a = ar.exec(m[4]))) attrs[a[1]] = decodeXml(a[2] ?? a[3] ?? '');
      }
      const name = m[3];
      const i = name.indexOf(':');
      const node = { name, local: i >= 0 ? name.slice(i + 1) : name, attrs, children: [], text: '' };
      top.children.push(node);
      if (!m[5]) stack.push(node);
    } else if (m[6] !== undefined && stack.length > 1) top.text += decodeXml(m[6]);
  }
  return root.children[0] ?? null;
}

const kids = (n, local) => (n ? n.children.filter((c) => c.local === local) : []);
const kid = (n, local) => (n ? n.children.find((c) => c.local === local) ?? null : null);
/** attribute by exact name, else by local name (namespace prefixes may differ between writers) */
function att(n, name) {
  if (!n) return undefined;
  if (name in n.attrs) return n.attrs[name];
  for (const k of Object.keys(n.attrs)) if (k.slice(k.indexOf(':') + 1) === name) return n.attrs[k];
  return undefined;
}
function textOf(n, skip = null) {
  if (!n) return '';
  let s = n.text;
  for (const c of n.children) if (!skip || !skip.has(c.local)) s += textOf(c, skip);
  return s;
}
function* walk(n) {
  if (!n) return;
  yield n;
  for (const c of n.children) yield* walk(c);
}
const isTrue = (v) => v === '1' || v === 'true';

/** attributes that are revision ids / xml noise and never compared */
const NOISE_ATTR = /^(xmlns(:.*)?|mc:Ignorable|xr\d*:uid|x14ac:dyDescent|xr:uid|uid|id|x14:id)$/;
/** stable text form of an element tree (sorted attributes, noise attributes dropped) */
function ser(n, skipLocals = null) {
  if (!n) return '';
  const attrs = Object.keys(n.attrs)
    .filter((k) => !NOISE_ATTR.test(k))
    .sort()
    .map((k) => `${k.slice(k.indexOf(':') + 1)}=${n.attrs[k]}`)
    .join(',');
  const inner = n.children.filter((c) => !skipLocals || !skipLocals.has(c.local)).map((c) => ser(c, skipLocals)).join('');
  const t = n.text.trim();
  return `${n.local}${attrs ? `[${attrs}]` : ''}${t ? `"${t}"` : ''}${inner ? `{${inner}}` : ''}`;
}

// ---------------------------------------------------------------------------------------------------------------
// Package + relationships
// ---------------------------------------------------------------------------------------------------------------

function dirOf(p) {
  const i = p.lastIndexOf('/');
  return i < 0 ? '' : p.slice(0, i + 1);
}
function resolvePath(base, target) {
  if (target.startsWith('/')) return target.slice(1);
  const parts = (dirOf(base) + target).split('/');
  const out = [];
  for (const p of parts) {
    if (p === '..') out.pop();
    else if (p !== '.' && p !== '') out.push(p);
  }
  return out.join('/');
}
function relsPath(part) {
  return `${dirOf(part)}_rels/${part.slice(dirOf(part).length)}.rels`;
}

async function openPackage(input) {
  const buf = typeof input === 'string' ? fs.readFileSync(input) : input;
  const zip = await JSZip.loadAsync(buf);
  const files = Object.keys(zip.files).filter((p) => !zip.files[p].dir);
  const lower = new Map(files.map((f) => [f.toLowerCase(), f]));
  const xmlCache = new Map();
  const relCache = new Map();
  const find = (p) => (zip.file(p) ? p : lower.get(p.toLowerCase()) ?? null);
  const read = async (p) => {
    const f = p && find(p);
    return f ? zip.file(f).async('string') : null;
  };
  const xml = async (p) => {
    if (!p) return null;
    if (!xmlCache.has(p)) {
      const t = await read(p);
      xmlCache.set(p, t == null ? null : parseXml(t));
    }
    return xmlCache.get(p);
  };
  /** Map rId -> {type (last URI segment), target (resolved part path or external URL), external} */
  const rels = async (part) => {
    if (relCache.has(part)) return relCache.get(part);
    const map = new Map();
    const doc = await xml(relsPath(part));
    for (const r of kids(doc, 'Relationship')) {
      const type = String(r.attrs.Type ?? '').split('/').pop();
      const external = r.attrs.TargetMode === 'External';
      map.set(r.attrs.Id, { type, external, target: external ? r.attrs.Target : resolvePath(part, r.attrs.Target ?? '') });
    }
    relCache.set(part, map);
    return map;
  };
  const relOfType = async (part, type) => [...(await rels(part)).values()].filter((r) => r.type === type);
  return { zip, files, read, xml, rels, relOfType };
}

// ---------------------------------------------------------------------------------------------------------------
// Styles
// ---------------------------------------------------------------------------------------------------------------

/** built-in number formats (ja-JP codes) — only ids whose code is certain; others stay '#<id>' */
const BUILTIN_FMT = {
  0: 'General',
  1: '0',
  2: '0.00',
  3: '#,##0',
  4: '#,##0.00',
  9: '0%',
  10: '0.00%',
  11: '0.00E+00',
  12: '# ?/?',
  13: '# ??/??',
  14: 'yyyy/m/d',
  15: 'd-mmm-yy',
  16: 'd-mmm',
  17: 'mmm-yy',
  18: 'h:mm AM/PM',
  19: 'h:mm:ss AM/PM',
  20: 'h:mm',
  21: 'h:mm:ss',
  22: 'yyyy/m/d h:mm',
  37: '#,##0 ;(#,##0)',
  38: '#,##0 ;[Red](#,##0)',
  39: '#,##0.00;(#,##0.00)',
  40: '#,##0.00;[Red](#,##0.00)',
  45: 'mm:ss',
  46: '[h]:mm:ss',
  47: 'mm:ss.0',
  48: '##0.0E+0',
  49: '@',
};

function colorSig(n, { autoIsEmpty = false } = {}) {
  if (!n) return '';
  const a = n.attrs;
  let s = '';
  if (isTrue(a.auto)) s = 'auto';
  else if (a.rgb != null) {
    let rgb = String(a.rgb).toUpperCase();
    if (rgb.length === 6) rgb = 'FF' + rgb;
    s = '#' + rgb;
  } else if (a.theme != null) s = 't' + a.theme;
  else if (a.indexed != null) s = 'i' + a.indexed;
  if (autoIsEmpty && (s === 'auto' || s === 'i64')) s = '';
  if (s && a.tint != null && Number(a.tint) !== 0) s += `~${Number(a.tint).toFixed(6)}`;
  return s;
}

function boolEl(n) {
  if (!n) return false;
  const v = n.attrs.val;
  return v == null || isTrue(v);
}

function fontSig(f) {
  if (!f) return '';
  const p = [];
  const name = kid(f, 'name')?.attrs.val ?? kid(f, 'rFont')?.attrs.val;
  if (name) p.push(`name=${name}`);
  const sz = kid(f, 'sz')?.attrs.val;
  if (sz != null) p.push(`sz=${Number(sz)}`);
  if (boolEl(kid(f, 'b'))) p.push('b');
  if (boolEl(kid(f, 'i'))) p.push('i');
  const u = kid(f, 'u');
  if (u && u.attrs.val !== 'none') p.push(`u=${u.attrs.val ?? 'single'}`);
  if (boolEl(kid(f, 'strike'))) p.push('strike');
  if (boolEl(kid(f, 'outline'))) p.push('outline');
  if (boolEl(kid(f, 'shadow'))) p.push('shadow');
  const va = kid(f, 'vertAlign')?.attrs.val;
  if (va && va !== 'baseline') p.push(`va=${va}`);
  const c = colorSig(kid(f, 'color'));
  if (c) p.push(`color=${c}`);
  return p.join(',');
}

function fillSig(fl, dxf = false) {
  if (!fl) return 'none';
  const pf = kid(fl, 'patternFill');
  if (pf) {
    const pt = pf.attrs.patternType ?? (dxf ? 'solid' : 'none');
    if (pt === 'none') return 'none';
    const fg = colorSig(kid(pf, 'fgColor'));
    // for a solid cell fill the bgColor is meaningless (Excel writes indexed=64); DXF fills use bgColor as the colour
    const bg = pt === 'solid' && !dxf ? '' : colorSig(kid(pf, 'bgColor'));
    return `${pt}:${fg}${bg ? '/' + bg : ''}`;
  }
  const gf = kid(fl, 'gradientFill');
  if (gf) return `gradient:${ser(gf)}`;
  return 'none';
}

function borderSig(b) {
  if (!b) return '';
  const p = [];
  const sideName = { start: 'left', end: 'right' };
  for (const c of b.children) {
    const st = c.attrs.style;
    if (!st || st === 'none') continue;
    const col = colorSig(kid(c, 'color'), { autoIsEmpty: true });
    p.push(`${sideName[c.local] ?? c.local}=${st}${col ? ':' + col : ''}`);
  }
  if (isTrue(b.attrs.diagonalUp)) p.push('dUp');
  if (isTrue(b.attrs.diagonalDown)) p.push('dDown');
  if (isTrue(b.attrs.outline)) p.push('outline');
  return p.sort().join(',');
}

function alignSig(a) {
  if (!a) return '';
  const at = a.attrs;
  const p = [];
  if (at.horizontal && at.horizontal !== 'general') p.push(`h=${at.horizontal}`);
  if (at.vertical && at.vertical !== 'bottom') p.push(`v=${at.vertical}`);
  if (isTrue(at.wrapText)) p.push('wrap');
  if (at.indent && at.indent !== '0') p.push(`ind=${at.indent}`);
  if (at.relativeIndent && at.relativeIndent !== '0') p.push(`rind=${at.relativeIndent}`);
  if (at.textRotation && at.textRotation !== '0') p.push(`rot=${at.textRotation}`);
  if (isTrue(at.shrinkToFit)) p.push('shrink');
  if (at.readingOrder && at.readingOrder !== '0') p.push(`ro=${at.readingOrder}`);
  if (isTrue(at.justifyLastLine)) p.push('jll');
  return p.join(',');
}

function protSig(pr) {
  if (!pr) return '';
  const p = [];
  if (pr.attrs.locked != null && !isTrue(pr.attrs.locked)) p.push('unlocked');
  if (isTrue(pr.attrs.hidden)) p.push('hidden');
  return p.join(',');
}

function buildStyles(doc) {
  const numFmts = new Map();
  for (const n of kids(kid(doc, 'numFmts'), 'numFmt')) numFmts.set(Number(n.attrs.numFmtId), n.attrs.formatCode ?? '');
  const fmt = (id) => {
    const n = Number(id ?? 0);
    // (the format name "General" is case-insensitive: LibreOffice writes GENERAL, Excel General)
    if (numFmts.has(n)) return /^general$/i.test(numFmts.get(n)) ? 'General' : numFmts.get(n);
    return BUILTIN_FMT[n] ?? `#${n}`;
  };
  const fonts = kids(kid(doc, 'fonts'), 'font');
  const fills = kids(kid(doc, 'fills'), 'fill');
  const borders = kids(kid(doc, 'borders'), 'border');
  const xfs = kids(kid(doc, 'cellXfs'), 'xf');
  const styleNames = new Map();
  // (builtinId 0 is the built-in style "Normal": a Japanese Excel saves it under its localized name 「標準」 — the same style)
  for (const cs of kids(kid(doc, 'cellStyles'), 'cellStyle')) if (!styleNames.has(Number(cs.attrs.xfId))) styleNames.set(Number(cs.attrs.xfId), cs.attrs.builtinId === '0' ? 'Normal' : cs.attrs.name ?? '');
  const dxfs = kids(kid(doc, 'dxfs'), 'dxf');
  const xfCache = new Map();
  const xfSig = (i) => {
    const idx = Number(i ?? 0);
    if (xfCache.has(idx)) return xfCache.get(idx);
    const xf = xfs[idx] ?? xfs[0];
    let sig = 'default';
    if (xf) {
      const a = xf.attrs;
      sig = [
        `fmt=${fmt(a.numFmtId)}`,
        `font=${fontSig(fonts[Number(a.fontId ?? 0)])}`,
        `fill=${fillSig(fills[Number(a.fillId ?? 0)])}`,
        `border=${borderSig(borders[Number(a.borderId ?? 0)])}`,
        `align=${alignSig(kid(xf, 'alignment'))}`,
        `prot=${protSig(kid(xf, 'protection'))}`,
        isTrue(a.quotePrefix) ? 'quotePrefix=あり' : '',
        `style=${styleNames.get(Number(a.xfId ?? 0)) ?? ''}`,
      ].join('|');
    }
    xfCache.set(idx, sig);
    return sig;
  };
  const dxfSigNode = (d) => {
    if (!d) return '';
    const nf = kid(d, 'numFmt');
    return [
      nf ? `fmt=${nf.attrs.formatCode ?? fmt(nf.attrs.numFmtId)}` : '',
      kid(d, 'font') ? `font=${fontSig(kid(d, 'font'))}` : '',
      kid(d, 'fill') ? `fill=${fillSig(kid(d, 'fill'), true)}` : '',
      kid(d, 'border') ? `border=${borderSig(kid(d, 'border'))}` : '',
      kid(d, 'alignment') ? `align=${alignSig(kid(d, 'alignment'))}` : '',
    ]
      .filter(Boolean)
      .join('|');
  };
  const dxfSig = (i) => (i == null ? '' : dxfSigNode(dxfs[Number(i)]));
  return { xfSig, dxfSig, dxfSigNode, fmt };
}

// ---------------------------------------------------------------------------------------------------------------
// Model
// ---------------------------------------------------------------------------------------------------------------

function siText(si) {
  if (!si) return '';
  const t = kid(si, 't');
  if (t && !kids(si, 'r').length) return textOf(t);
  let s = t ? textOf(t) : '';
  for (const r of kids(si, 'r')) s += textOf(kid(r, 't'));
  return s;
}

function normSqref(s) {
  return String(s ?? '')
    .trim()
    .split(/\s+/)
    .filter(Boolean)
    .sort()
    .join(' ');
}

function cfRuleSig(rule, styles) {
  const a = rule.attrs;
  const flags = ['operator', 'aboveAverage', 'percent', 'bottom', 'rank', 'text', 'timePeriod', 'equalAverage', 'stdDev']
    .filter((k) => a[k] != null)
    .map((k) => `${k}=${a[k]}`);
  if (isTrue(a.stopIfTrue)) flags.push('stop');
  const formulas = [...kids(rule, 'formula'), ...kids(rule, 'f')].map((f) => textOf(f));
  const extra = ['colorScale', 'dataBar', 'iconSet']
    .map((k) => kid(rule, k))
    .filter(Boolean)
    .map((n) => ser(n, new Set(['extLst'])));
  const dxf = a.dxfId != null ? styles.dxfSig(a.dxfId) : kid(rule, 'dxf') ? styles.dxfSigNode(kid(rule, 'dxf')) : '';
  return `${a.type}|${flags.join(',')}|f=${formulas.join(' ; ')}|${extra.join('')}|dxf=${dxf}`;
}

function dvSig(dv) {
  const a = dv.attrs;
  const keys = ['type', 'operator', 'allowBlank', 'showDropDown', 'showInputMessage', 'showErrorMessage', 'errorStyle', 'imeMode', 'errorTitle', 'error', 'promptTitle', 'prompt'];
  const attrs = keys.filter((k) => a[k] != null).map((k) => `${k}=${a[k]}`);
  const f = (name) => {
    const n = kid(dv, name);
    if (!n) return '';
    return kid(n, 'f') ? textOf(kid(n, 'f')) : textOf(n);
  };
  const sq = a.sqref ?? textOf(kid(dv, 'sqref'));
  return `${normSqref(sq)}|${attrs.join(',')}|f1=${f('formula1')}|f2=${f('formula2')}`;
}

async function chartSig(pkg, part) {
  const doc = await pkg.xml(part);
  if (!doc) return `chart(missing ${part})`;
  const chart = kid(doc, 'chart');
  const title = kid(chart, 'title');
  const titleText = title ? (kid(kid(title, 'tx'), 'strRef') ? 'ref:' + textOf(kid(kid(kid(title, 'tx'), 'strRef'), 'f')) : textOf(title, new Set(['numCache', 'strCache']))) : '';
  const plot = kid(chart, 'plotArea');
  const types = [];
  for (const t of plot?.children ?? []) {
    if (!/Chart$/.test(t.local)) continue;
    const series = kids(t, 'ser').map((s) => {
      const fOf = (name) => {
        const n = kid(s, name);
        if (!n) return '';
        // (a series reference is absolute whether it is written with $ or without: Excel adds them)
        for (const d of walk(n)) if (d.local === 'f') return textOf(d).replace(/\$/g, '');
        const v = kid(n, 'v');
        return v ? `"${textOf(v)}"` : '';
      };
      return ['tx', 'cat', 'val', 'xVal', 'yVal', 'bubbleSize']
        .map((k) => [k, fOf(k)])
        .filter(([, v]) => v)
        .map(([k, v]) => `${k}=${v}`)
        .join(',');
    });
    const dir = kid(t, 'barDir')?.attrs.val;
    const grouping = kid(t, 'grouping')?.attrs.val;
    types.push(`${t.local}${dir ? ':' + dir : ''}${grouping ? ':' + grouping : ''}[${series.join(' ; ')}]`);
  }
  return `title=${titleText}|${types.join(' + ')}`;
}

async function drawingSigs(pkg, part) {
  const doc = await pkg.xml(part);
  if (!doc) return [];
  const rels = await pkg.rels(part);
  const out = [];
  for (const an of doc.children) {
    if (!/Anchor$/.test(an.local)) continue;
    const pos = (n) => (n ? `${kid(n, 'col') ? textOf(kid(n, 'col')) : '?'},${kid(n, 'row') ? textOf(kid(n, 'row')) : '?'}` : '');
    const offs = ['from', 'to'].map((k) => kid(an, k)).filter(Boolean).map((n) => `${textOf(kid(n, 'colOff'))}/${textOf(kid(n, 'rowOff'))}`).join(';');
    let what = '';
    for (const d of walk(an)) {
      if (d.local === 'chart' && att(d, 'id')) {
        const r = rels.get(att(d, 'id'));
        what = r ? await chartSig(pkg, r.target) : 'chart(?)';
        break;
      }
      if (d.local === 'pic') {
        what = `pic:${att(kid(kid(d, 'nvPicPr'), 'cNvPr'), 'name') ?? ''}`;
        break;
      }
      if (d.local === 'sp') {
        what = `shape:${att(kid(kid(d, 'nvSpPr'), 'cNvPr'), 'name') ?? ''}:${textOf(kid(d, 'txBody')).trim()}`;
        break;
      }
      if (d.local === 'grpSp') {
        what = 'group';
        break;
      }
    }
    const anchor = an.local === 'absoluteAnchor' ? 'abs' : `${an.local.replace(/Anchor$/, '')}(${pos(kid(an, 'from'))}${kid(an, 'to') ? '-' + pos(kid(an, 'to')) : ''})${an.attrs.editAs ? ':' + an.attrs.editAs : ''}`;
    out.push({ sig: `${anchor} ${what}`, offs });
  }
  return out;
}

async function pivotSig(pkg, part) {
  const doc = await pkg.xml(part);
  if (!doc) return null;
  let cacheFields = [];
  let source = '';
  for (const r of await pkg.relOfType(part, 'pivotCacheDefinition')) {
    const cd = await pkg.xml(r.target);
    const ws = kid(kid(cd, 'cacheSource'), 'worksheetSource');
    source = ws ? `${ws.attrs.sheet ?? ''}!${ws.attrs.ref ?? ''}${ws.attrs.name ? ' name=' + ws.attrs.name : ''}` : ser(kid(cd, 'cacheSource'));
    cacheFields = kids(kid(cd, 'cacheFields'), 'cacheField').map((f) => f.attrs.name ?? '');
  }
  const fname = (x) => (Number(x) === -2 ? '(値)' : cacheFields[Number(x)] ?? `#${x}`);
  const rowF = kids(kid(doc, 'rowFields'), 'field').map((f) => fname(f.attrs.x));
  const colF = kids(kid(doc, 'colFields'), 'field').map((f) => fname(f.attrs.x));
  const dataF = kids(kid(doc, 'dataFields'), 'dataField').map((f) => `${f.attrs.name ?? ''}(${fname(f.attrs.fld)}:${f.attrs.subtotal ?? 'sum'})`);
  const pageF = kids(kid(doc, 'pageFields'), 'pageField').map((f) => fname(f.attrs.fld));
  return {
    name: doc.attrs.name ?? '',
    sig: `${doc.attrs.name ?? ''}|src=${source}|fields=${cacheFields.join(',')}|rows=${rowF.join(',')}|cols=${colF.join(',')}|data=${dataF.join(',')}|page=${pageF.join(',')}`,
    location: kid(doc, 'location')?.attrs.ref ?? '',
  };
}

/**
 * A cell of type t="d" (an ISO 8601 date / date-time text, what Strict Open XML writes for a date) as the serial number Excel reads it as; null when the text is not one.
 * Dates before 1900-03-01 (the 1900 leap-year bug) are not converted.
 */
export function isoDateSerial(text, date1904 = false) {
  const m = /^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2})(?:\.(\d+))?)?)?Z?$/.exec(String(text).trim());
  if (!m) return null;
  const day = Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3])) / 86400000;
  const epoch = date1904 ? Date.UTC(1904, 0, 1) / 86400000 : Date.UTC(1899, 11, 30) / 86400000;
  const serial = day - epoch;
  if (!date1904 && serial < 61) return null;
  const frac = (Number(m[4] ?? 0) * 3600 + Number(m[5] ?? 0) * 60 + Number(m[6] ?? 0) + Number('0.' + (m[7] ?? '0'))) / 86400;
  return serial + frac;
}

async function parseSheet(pkg, part, name, styles, sst, date1904 = false) {
  const doc = await pkg.xml(part);
  const rels = await pkg.rels(part);
  const sh = { name, part, cells: new Map(), rows: new Map(), cols: [], extras: {} };
  if (!doc) return sh;
  sh.tabColor = colorSig(kid(kid(doc, 'sheetPr'), 'tabColor'));
  sh.dimension = kid(doc, 'dimension')?.attrs.ref ?? '';
  const fp = kid(doc, 'sheetFormatPr');
  sh.defaultColWidth = fp?.attrs.defaultColWidth != null ? Number(fp.attrs.defaultColWidth) : null;
  // (values equal to the schema default are the same as no value: baseColWidth 8, zeroHeight false; true / 1 and false / 0 are the same)
  const b01 = (v, d) => (v == null ? d : v === 'true' || v === '1' ? '1' : '0');
  sh.formatPr = fp ? `baseColWidth=${fp.attrs.baseColWidth && fp.attrs.baseColWidth !== '8' ? fp.attrs.baseColWidth : ''},defaultColWidth=${fp.attrs.defaultColWidth ?? ''},zeroHeight=${b01(fp.attrs.zeroHeight, '0') === '1' ? '1' : ''}` : '';
  if (sh.formatPr === 'baseColWidth=,defaultColWidth=,zeroHeight=') sh.formatPr = ''; // an element with nothing in it = no element
  // a sheet whose default row height is custom (Google Sheets: customHeight="1" defaultRowHeight="15.75") has that height on every row without its own;
  // Excel writes it out on each <row>
  sh.defaultRowHeight = fp?.attrs.defaultRowHeight != null && isTrue(fp.attrs.customHeight) ? Number(fp.attrs.defaultRowHeight) : null;

  // sheet view: the pane split is real, selection / scroll position is not
  const sv = kid(kid(doc, 'sheetViews'), 'sheetView');
  const pane = kid(sv, 'pane');
  const va = sv?.attrs ?? {};
  sh.view = [
    // (numbers: Google Sheets writes xSplit="2.0", Excel "2")
    pane && pane.attrs.state !== 'split' ? `pane=${Number(pane.attrs.xSplit ?? 0)},${Number(pane.attrs.ySplit ?? 0)},${pane.attrs.state ?? 'split'}` : pane ? `split=${Number(pane.attrs.xSplit ?? 0)},${Number(pane.attrs.ySplit ?? 0)}` : 'pane=none',
    `zoom=${va.zoomScale ?? '100'}`,
    `grid=${b01(va.showGridLines, '1')}`,
    `rtl=${b01(va.rightToLeft, '0')}`,
    `zeros=${b01(va.showZeros, '1')}`,
    `headers=${b01(va.showRowColHeaders, '1')}`,
    `view=${va.view ?? 'normal'}`,
  ].join('|');
  sh.selection = `${va.tabSelected ?? ''}|${va.topLeftCell ?? ''}|${pane?.attrs.topLeftCell ?? ''}|${pane?.attrs.activePane ?? ''}|${kids(sv, 'selection').map((s) => ser(s)).join('')}`;

  // columns as segments [min, max, props]
  for (const c of kids(kid(doc, 'cols'), 'col')) {
    sh.cols.push({
      min: Number(c.attrs.min),
      max: Number(c.attrs.max),
      width: c.attrs.width != null ? Number(c.attrs.width) : null,
      custom: isTrue(c.attrs.customWidth),
      hidden: isTrue(c.attrs.hidden),
      lvl: Number(c.attrs.outlineLevel ?? 0),
      s: c.attrs.style != null ? Number(c.attrs.style) : null,
    });
  }

  // cells
  const masters = new Map();
  const pending = [];
  for (const row of kids(kid(doc, 'sheetData'), 'row')) {
    const r = Number(row.attrs.r);
    sh.rows.set(r, {
      ht: row.attrs.ht != null ? Number(row.attrs.ht) : null,
      custom: isTrue(row.attrs.customHeight),
      hidden: isTrue(row.attrs.hidden),
      lvl: Number(row.attrs.outlineLevel ?? 0),
      s: isTrue(row.attrs.customFormat) && row.attrs.s != null ? Number(row.attrs.s) : null,
    });
    let col = 0;
    for (const c of row.children) {
      if (c.local !== 'c') continue;
      const p = c.attrs.r ? parseCell(c.attrs.r) : { row: r, col: col + 1 };
      col = p.col;
      const addr = cellAddr(p.row, p.col);
      const t = c.attrs.t ?? 'n';
      const vn = kid(c, 'v');
      const vt = vn ? textOf(vn) : null;
      let type = 'z';
      let value = null;
      let sstIdx;
      if (t === 's' && vt != null) {
        type = 's';
        sstIdx = Number(vt);
        value = sst[sstIdx] ?? '';
      } else if (t === 'inlineStr') {
        type = 's';
        value = siText(kid(c, 'is'));
      } else if (t === 'str' && vt != null) {
        type = 's';
        value = vt;
      } else if (t === 'b' && vt != null) {
        type = 'b';
        value = isTrue(vt);
      } else if (t === 'e' && vt != null) {
        type = 'e';
        value = vt;
      } else if (t === 'd' && vt != null) {
        // (Strict Open XML stores a date as ISO text: Excel reads it as the date's serial number, and so does the Excel engine's output)
        const serial = isoDateSerial(vt, date1904);
        type = serial == null ? 'd' : 'n';
        value = serial == null ? vt : serial;
      } else if (vt != null && vt !== '') {
        type = 'n';
        value = Number(vt);
      }
      const cell = { row: p.row, col: p.col, s: c.attrs.s != null ? Number(c.attrs.s) : 0, type, value, sstIdx };
      if (c.attrs.cm != null) cell.dyn = true;
      if (c.attrs.vm != null) cell.vm = true;
      const fn = kid(c, 'f');
      if (fn) {
        const ft = fn.attrs.t ?? 'normal';
        const text = textOf(fn);
        cell.fType = ft === 'shared' ? 'normal' : ft;
        if (ft === 'array' || ft === 'dataTable') cell.ref = fn.attrs.ref ?? addr;
        if (ft === 'dataTable') cell.f = `TABLE(${fn.attrs.r1 ?? ''},${fn.attrs.r2 ?? ''})${ser(fn)}`;
        else if (ft === 'shared') {
          cell.shared = true;
          if (text) {
            cell.f = text;
            masters.set(fn.attrs.si, { text, row: p.row, col: p.col });
          } else pending.push([cell, fn.attrs.si]);
        } else cell.f = text;
      }
      sh.cells.set(addr, cell);
    }
  }
  for (const [cell, si] of pending) {
    const m = masters.get(si);
    cell.f = m ? shiftFormula(m.text, cell.row - m.row, cell.col - m.col) : `(shared si=${si} without master)`;
  }

  // features
  sh.merges = kids(kid(doc, 'mergeCells'), 'mergeCell')
    .map((m) => m.attrs.ref)
    .sort();
  const af = kid(doc, 'autoFilter');
  sh.autoFilter = af ? `${af.attrs.ref ?? ''}${af.children.length ? ' ' + af.children.map((c) => ser(c)).join('') : ''}` : '';

  const cf = [];
  const cfRaw = [];
  const addCf = (block, rules, sqref) => {
    for (const rule of rules) {
      const prio = Number(rule.attrs.priority ?? 0);
      cf.push({ prio, sig: `${normSqref(sqref)}|${cfRuleSig(rule, styles)}` });
      cfRaw.push(`${prio}:${rule.attrs.dxfId ?? ''}`);
    }
    void block;
  };
  for (const block of kids(doc, 'conditionalFormatting')) addCf(block, kids(block, 'cfRule'), block.attrs.sqref);
  const dv = [];
  for (const d of kids(kid(doc, 'dataValidations'), 'dataValidation')) dv.push(dvSig(d));
  const sparklines = [];
  for (const ext of kids(kid(doc, 'extLst'), 'ext')) {
    for (const cfs of kids(ext, 'conditionalFormattings'))
      for (const block of kids(cfs, 'conditionalFormatting')) addCf(block, kids(block, 'cfRule'), textOf(kid(block, 'sqref')));
    for (const dvs of kids(ext, 'dataValidations')) for (const d of kids(dvs, 'dataValidation')) dv.push(dvSig(d));
    for (const sgs of kids(ext, 'sparklineGroups'))
      for (const g of kids(sgs, 'sparklineGroup'))
        sparklines.push(
          kids(kid(g, 'sparklines'), 'sparkline')
            .map((s) => `${textOf(kid(s, 'f'))}->${textOf(kid(s, 'sqref'))}`)
            .join(','),
        );
  }
  cf.sort((a, b) => a.prio - b.prio);
  sh.cf = cf.map((c) => c.sig);
  sh.cfRaw = cfRaw.join(',');
  sh.dv = dv.sort();
  sh.sparklines = sparklines.sort();

  sh.hyperlinks = kids(kid(doc, 'hyperlinks'), 'hyperlink')
    .map((h) => {
      const rid = att(h, 'id');
      const target = rid ? rels.get(rid)?.target ?? '?' : '';
      return `${h.attrs.ref}|${target}|${h.attrs.location ?? ''}|${h.attrs.display ?? ''}|${h.attrs.tooltip ?? ''}`;
    })
    .sort();
  const prot = kid(doc, 'sheetProtection');
  // (an attribute that equals its schema default is the same as no attribute: openpyxl writes them all, Excel only what differs)
  const PROT_DEFAULT = { sheet: '0', objects: '0', scenarios: '0', formatCells: '1', formatColumns: '1', formatRows: '1', insertColumns: '1', insertRows: '1', insertHyperlinks: '1', deleteColumns: '1', deleteRows: '1', selectLockedCells: '0', sort: '1', autoFilter: '1', pivotTables: '1', selectUnlockedCells: '0' };
  const boolStr = (v) => (v === 'true' ? '1' : v === 'false' ? '0' : v);
  sh.protection = prot
    ? Object.keys(prot.attrs)
        .filter((k) => !/hashValue|saltValue|spinCount|algorithmName|password/.test(k))
        .filter((k) => !(k in PROT_DEFAULT && boolStr(prot.attrs[k]) === PROT_DEFAULT[k] && k !== 'sheet'))
        .sort()
        .map((k) => `${k}=${boolStr(prot.attrs[k])}`)
        .join(',')
    : '';
  sh.breaks = ['rowBreaks', 'colBreaks']
    .map((k) => `${k}:${kids(kid(doc, k), 'brk').map((b) => b.attrs.id).join(',')}`)
    .join('|');

  // related parts
  sh.tables = [];
  sh.drawings = [];
  sh.drawingOffsets = [];
  sh.comments = [];
  sh.pivots = [];
  for (const r of rels.values()) {
    if (r.external) continue;
    if (r.type === 'table') {
      const t = await pkg.xml(r.target);
      if (!t) continue;
      const cols = kids(kid(t, 'tableColumns'), 'tableColumn').map(
        (c) => `${c.attrs.name}${kid(c, 'calculatedColumnFormula') ? '=' + textOf(kid(c, 'calculatedColumnFormula')) : ''}${c.attrs.totalsRowFunction ? ':' + c.attrs.totalsRowFunction : ''}`,
      );
      const si = kid(t, 'tableStyleInfo');
      sh.tables.push(
        `${t.attrs.name}|${t.attrs.displayName}|${t.attrs.ref}|hdr=${t.attrs.headerRowCount ?? '1'}|tot=${t.attrs.totalsRowCount ?? '0'}|af=${kid(t, 'autoFilter')?.attrs.ref ?? ''}|${cols.join(',')}|${si ? ser(si) : ''}`,
      );
    } else if (r.type === 'drawing') {
      for (const d of await drawingSigs(pkg, r.target)) {
        sh.drawings.push(d.sig);
        sh.drawingOffsets.push(d.offs);
      }
    } else if (r.type === 'comments') {
      const c = await pkg.xml(r.target);
      for (const cm of kids(kid(c, 'commentList'), 'comment')) sh.comments.push(`${cm.attrs.ref}: ${textOf(kid(cm, 'text'), new Set(['rPh', 'phoneticPr']))}`);
    } else if (r.type === 'threadedComment') {
      const c = await pkg.xml(r.target);
      for (const cm of kids(c, 'threadedComment')) sh.comments.push(`${cm.attrs.ref} (スレッド): ${textOf(kid(cm, 'text'))}`);
    } else if (r.type === 'pivotTable') {
      const p = await pivotSig(pkg, r.target);
      if (p) sh.pivots.push(p);
    }
  }
  sh.tables.sort();
  sh.drawings.sort();
  sh.comments.sort();
  sh.pivots.sort((a, b) => (a.name < b.name ? -1 : a.name > b.name ? 1 : 0));
  return sh;
}

/** Parses an .xlsx / .xlsm (path or bytes) into the normalized model compared by comparePackageModels. */
export async function loadPackageModel(input) {
  const pkg = await openPackage(input);
  const root = await pkg.xml('_rels/.rels');
  const wbRel = kids(root, 'Relationship').find((r) => String(r.attrs.Type ?? '').endsWith('/officeDocument'));
  const wbPart = wbRel ? resolvePath('', String(wbRel.attrs.Target).replace(/^\//, '')) : 'xl/workbook.xml';
  const wb = await pkg.xml(wbPart);
  if (!wb) throw new Error(`workbook part not found (${wbPart})`);
  const wbRels = await pkg.rels(wbPart);
  const relTarget = (type) => [...wbRels.values()].find((r) => r.type === type)?.target ?? null;
  const styles = buildStyles(await pkg.xml(relTarget('styles')));
  const sstDoc = await pkg.xml(relTarget('sharedStrings'));
  const sst = kids(sstDoc, 'si').map(siText);

  const sheetEls = kids(kid(wb, 'sheets'), 'sheet');
  const sheets = [];
  for (const s of sheetEls) {
    const r = wbRels.get(att(s, 'id'));
    sheets.push({ name: s.attrs.name, state: s.attrs.state ?? 'visible', part: r?.target ?? null, kind: r?.type ?? '?' });
  }
  const names = kids(kid(wb, 'definedNames'), 'definedName').map((n) => ({
    name: n.attrs.name,
    scope: n.attrs.localSheetId != null ? sheets[Number(n.attrs.localSheetId)]?.name ?? `#${n.attrs.localSheetId}` : null,
    text: textOf(n),
    hidden: isTrue(n.attrs.hidden),
  }));
  const model = {
    files: pkg.files,
    sheets,
    names,
    date1904: isTrue(kid(wb, 'workbookPr')?.attrs.date1904),
    workbookMisc: [ser(kid(wb, 'calcPr')), ser(kid(wb, 'fileVersion')), ser(kid(wb, 'bookViews')), ser(kid(wb, 'pivotCaches'))].join('|'),
    styles,
    sheetModels: new Map(),
    docProps: [(await pkg.read('docProps/core.xml')) ?? '', (await pkg.read('docProps/app.xml')) ?? ''].join('\n'),
  };
  for (const s of sheets) {
    if (s.kind !== 'worksheet' || !s.part) continue;
    model.sheetModels.set(s.name, await parseSheet(pkg, s.part, s.name, styles, sst, model.date1904));
  }
  return model;
}

// ---------------------------------------------------------------------------------------------------------------
// Comparison
// ---------------------------------------------------------------------------------------------------------------

/** What the package view filters as known-benign (key -> Japanese description). */
export const FILTER_LABELS = {
  calcChain: 'xl/calcChain.xml の有無・内容（計算チェーン。Excel は開くときに作り直す／内蔵エンジンは削除する）',
  docProps: 'docProps/core.xml・app.xml（作成者・更新日時・アプリ情報）',
  xfIndex: 'セルのスタイル番号 s= の違い（解決した書式＝表示形式・フォント・塗り・罫線・配置・保護・セルスタイル名は同じ）',
  sstIndex: '共有文字列 (sharedStrings) の番号・並びの違い（文字列は同じ）',
  sharedFormula: '共有数式 (t="shared") と個別の数式の書き方の違い（展開した数式は同じ）',
  cachedValues: '数式セルのキャッシュ値 <v>（Excel は開くときに再計算する。再計算後の値は Excel ビューで比較）',
  spill: '動的配列のスピル範囲 (ref) とスピル先セルの値（Excel が再計算で決める。Excel ビューの spills で比較）',
  emptyCells: '値のない空セル <c/> の有無（見た目の書式は同じ）',
  dimension: '<dimension ref>（使用範囲の目安。Excel が保存時に計算し直す）',
  rowAttrs: '行の自動の高さ（customHeight なし）・spans・x14ac:dyDescent',
  colSplit: '<col> 要素の分け方・まとめ方、customWidth / bestFit 属性、書式なし列の既定幅（列ごとの幅・非表示・書式で比較）',
  selection: 'シートの選択セル・アクティブセル・スクロール位置・選択中のシート (selection / activeCell / topLeftCell / tabSelected)',
  cfPriority: '条件付き書式の priority 番号・dxfId・要素の並び（ルールの優先順と書式の中身で比較）',
  workbookMisc: 'workbook.xml の calcPr（fullCalcOnLoad 等）・fileVersion・bookViews・absPath・revisionPtr・pivotCache の cacheId',
  filterDatabase: '名前 _xlnm._FilterDatabase（オートフィルターから Excel が自動で作る。autoFilter で比較）',
  pivotLayout: 'ピボットテーブルの表示位置 location（Excel が更新時に作り直す。元データ・フィールドは比較）',
  anchorOffsets: 'グラフ・図形のセル内オフセット colOff / rowOff（アンカーのセル位置は比較）',
  xmlNoise: 'XML の属性の順序・名前空間の宣言と接頭辞・mc:Ignorable・xr:uid 等のリビジョン ID・空白（常に無視）',
  styleNoise: 'フォントの family / charset / scheme 属性、xf の applyXxx フラグ、使われていないスタイル（常に無視）',
  otherParts: 'パーツのファイル名・rId、テーマ、プリンター設定、VML、ピボットキャッシュのレコード、グラフの値キャッシュ（常に無視）',
};
const ALWAYS = ['xmlNoise', 'styleNoise', 'otherParts'];

const MAX_EX = 8;

function sameValue(a, b) {
  const ta = a?.type ?? 'z';
  const tb = b?.type ?? 'z';
  const va = a?.value ?? null;
  const vb = b?.value ?? null;
  if (ta === 'z' && tb === 'z') return true;
  if (ta === 'n' && tb === 'n') return va === vb || Math.abs(va - vb) <= 1e-12 * Math.max(1, Math.abs(va));
  return ta === tb && va === vb;
}
const showVal = (c) => (!c || c.type === 'z' ? '(空)' : JSON.stringify(c.value));
const fKind = (c) => (!c?.f ? '' : c.dyn ? 'dyn' : c.fType === 'array' ? `array ${c.ref}` : c.fType === 'dataTable' ? `dataTable ${c.ref}` : 'normal');

const STYLE_PARTS = ['fmt', 'font', 'fill', 'border', 'align', 'prot', 'quotePrefix', 'style'];
/** 'font: name=游ゴシック,sz=11 -> name=游ゴシック,sz=11,b' for the parts of two xf signatures that differ */
function styleDiff(a, b) {
  const pa = a.split('|');
  const pb = b.split('|');
  const out = [];
  for (let i = 0; i < Math.max(pa.length, pb.length); i++) {
    if (pa[i] === pb[i]) continue;
    const label = STYLE_PARTS[i] ?? String(i);
    const strip = (s) => (s ?? '').replace(new RegExp(`^${label}=`), '') || 'なし';
    out.push(`${label}: ${strip(pa[i])} -> ${strip(pb[i])}`);
  }
  return out.join(', ');
}

function listDiff(a, b) {
  const missing = a.filter((x) => !b.includes(x));
  const extra = b.filter((x) => !a.includes(x));
  if (!missing.length && !extra.length) return 'order: ' + a.map((x, i) => (x === b[i] ? '=' : `${i + 1}`)).join('');
  return `${missing.length ? `a のみ: ${missing.slice(0, MAX_EX).join(' ; ')}` : ''}${missing.length && extra.length ? '  ' : ''}${extra.length ? `b のみ: ${extra.slice(0, MAX_EX).join(' ; ')}` : ''}`;
}

/** column segments -> props of column c */
function colAt(sh, c) {
  for (const seg of sh.cols) if (c >= seg.min && c <= seg.max) return seg;
  return null;
}

/**
 * Compares two package models. Returns {rows, filtered}: rows = {group, sheet, ok, detail} for the real
 * comparison + {filtered: true} rows for every known-benign category that occurred; filtered = [{key, label, count,
 * examples}].
 */
export function comparePackageModels(A, B, opts = {}) {
  const rows = [];
  const push = (group, sheet, ok, detail = '') => rows.push({ group, sheet, ok, detail });
  const counts = new Map();
  const filt = (key, n = 1, example = null) => {
    const e = counts.get(key) ?? { count: 0, examples: [] };
    e.count += n;
    if (example && e.examples.length < 3) e.examples.push(example);
    counts.set(key, e);
  };

  // parts that only one package has (by kind; file names are not compared)
  const kinds = (m) => {
    const k = {};
    for (const f of m.files) {
      let kind = null;
      if (/^xl\/vbaProject\.bin$/i.test(f)) kind = 'VBA マクロ';
      else if (/^xl\/externalLinks\/externalLink\d*\.xml$/i.test(f)) kind = '外部リンク';
      else if (/^xl\/connections\.xml$/i.test(f)) kind = 'データ接続';
      else if (/^xl\/queryTables\//i.test(f)) kind = 'クエリテーブル';
      else if (/^xl\/slicers\//i.test(f)) kind = 'スライサー';
      else if (/^xl\/timelines\//i.test(f)) kind = 'タイムライン';
      else if (/^xl\/model\//i.test(f)) kind = 'データモデル';
      else if (/^xl\/richData\//i.test(f)) kind = 'セル内の画像・リッチデータ';
      else if (/^xl\/media\//i.test(f)) kind = '画像';
      else if (/^xl\/charts\/chart\d*\.xml$/i.test(f)) kind = 'グラフ';
      else if (/^xl\/pivotTables\/pivotTable\d*\.xml$/i.test(f)) kind = 'ピボットテーブル';
      if (kind) k[kind] = (k[kind] ?? 0) + 1;
    }
    return k;
  };
  const ka = kinds(A);
  const kb = kinds(B);
  const kd = [...new Set([...Object.keys(ka), ...Object.keys(kb)])].filter((k) => (ka[k] ?? 0) !== (kb[k] ?? 0)).map((k) => `${k} ${ka[k] ?? 0} -> ${kb[k] ?? 0}`);
  push('parts', '', kd.length === 0, kd.join(' ; '));
  const hasCalc = (m) => m.files.some((f) => /^xl\/calcChain\.xml$/i.test(f));
  if (hasCalc(A) !== hasCalc(B)) filt('calcChain', 1, `a: ${hasCalc(A) ? 'あり' : 'なし'} / b: ${hasCalc(B) ? 'あり' : 'なし'}`);
  if (A.docProps !== B.docProps) filt('docProps');
  if (A.workbookMisc !== B.workbookMisc) filt('workbookMisc');

  // workbook
  const sl = (m) => m.sheets.map((s) => `${s.name}${s.state !== 'visible' ? `(${s.state})` : ''}`);
  const sa = sl(A);
  const sb = sl(B);
  push('sheets', '', JSON.stringify(sa) === JSON.stringify(sb), JSON.stringify(sa) === JSON.stringify(sb) ? '' : `${sa.join(',')} -> ${sb.join(',')}`);
  if (A.date1904 !== B.date1904) push('date1904', '', false, `${A.date1904} -> ${B.date1904}`);
  const nl = (m, fd) =>
    m.names
      .filter((n) => fd === (n.name === '_xlnm._FilterDatabase'))
      .map((n) => `${n.scope ? n.scope + '!' : ''}${n.name}=${n.text}${n.hidden ? ' (hidden)' : ''}`)
      .sort();
  const na = nl(A, false);
  const nb = nl(B, false);
  push('names', '', JSON.stringify(na) === JSON.stringify(nb), JSON.stringify(na) === JSON.stringify(nb) ? '' : listDiff(na, nb));
  if (JSON.stringify(nl(A, true)) !== JSON.stringify(nl(B, true))) filt('filterDatabase', 1, `${nl(A, true).join(',')} -> ${nl(B, true).join(',')}`);

  for (const [name, x] of A.sheetModels) {
    const y = B.sheetModels.get(name);
    if (!y) {
      push('sheet', name, false, 'b にありません');
      continue;
    }
    compareSheet(name, x, y, A.styles, B.styles, push, filt, opts);
  }
  for (const name of B.sheetModels.keys()) if (!A.sheetModels.has(name)) push('sheet', name, false, 'a にありません');

  const filtered = [];
  for (const [key, label] of Object.entries(FILTER_LABELS)) {
    const e = counts.get(key);
    if (e) filtered.push({ key, label, count: e.count, examples: e.examples });
    else if (ALWAYS.includes(key)) filtered.push({ key, label, count: null, examples: [] });
  }
  for (const f of filtered)
    rows.push({
      group: f.key,
      sheet: '',
      ok: true,
      filtered: true,
      detail: `${f.label}${f.count != null ? `: ${f.count} 件` : ''}${f.examples.length ? `（例: ${f.examples.join(' ; ')}）` : ''}`,
    });
  return { rows, filtered };
}

function compareSheet(name, x, y, stA, stB, push, filt, opts) {
  // ---- cells
  const spillBoxes = [];
  for (const m of [x, y]) for (const c of m.cells.values()) if (c.dyn && c.ref) { const b = parseRange(c.ref); if (b) spillBoxes.push({ b, master: cellAddr(c.row, c.col) }); }
  const inSpill = (addr, p) => spillBoxes.some((s) => s.master !== addr && p.row >= s.b.r1 && p.row <= s.b.r2 && p.col >= s.b.c1 && p.col <= s.b.c2);
  const effStyle = (m, st, cell, p) => {
    if (cell) return st.xfSig(cell.s ?? 0);
    const row = m.rows.get(p.row);
    if (row && row.s != null) return st.xfSig(row.s);
    const col = colAt(m, p.col);
    if (col && col.s != null) return st.xfSig(col.s);
    return st.xfSig(0);
  };
  const fDiff = [];
  const vDiff = [];
  const sDiff = [];
  let fTotal = 0;
  let fSame = 0;
  const addrs = new Set([...x.cells.keys(), ...y.cells.keys()]);
  for (const addr of addrs) {
    const a = x.cells.get(addr);
    const b = y.cells.get(addr);
    const p = a ?? b;
    const pos = { row: p.row, col: p.col };
    const fa = a?.f ?? null;
    const fb = b?.f ?? null;
    if (fa != null || fb != null) {
      fTotal++;
      if (fa !== fb) fDiff.push(`${addr}: ${fa == null ? `(数式なし ${showVal(a)})` : '=' + fa} -> ${fb == null ? `(数式なし ${showVal(b)})` : '=' + fb}`);
      else {
        const ka = fKind(a);
        const kb = fKind(b);
        if (ka !== kb && !(ka === 'dyn' && kb === 'dyn')) fDiff.push(`${addr}: =${fa} [${ka} -> ${kb}]`);
        else {
          fSame++;
          if (!!a.shared !== !!b.shared) filt('sharedFormula', 1, addr);
          if (a.dyn && b.dyn && a.ref !== b.ref) filt('spill', 1, `${addr} ${a.ref} -> ${b.ref}`);
        }
      }
      if (fa != null && fb != null && !sameValue(a, b)) filt('cachedValues', 1, `${addr} ${showVal(a)} -> ${showVal(b)}`);
    } else if (!sameValue(a, b)) {
      if (inSpill(addr, pos)) filt('spill', 1, `${addr} ${showVal(a)} -> ${showVal(b)}`);
      else vDiff.push(`${addr}: ${showVal(a)} -> ${showVal(b)}`);
    } else if (a?.sstIdx != null && b?.sstIdx != null && a.sstIdx !== b.sstIdx) filt('sstIndex');
    const sa = effStyle(x, stA, a, pos);
    const sb = effStyle(y, stB, b, pos);
    if (sa !== sb) sDiff.push(`${addr}: ${styleDiff(sa, sb)}`);
    else {
      if (a && b && (a.s ?? 0) !== (b.s ?? 0)) filt('xfIndex');
      if (!a !== !b && (a ?? b).type === 'z' && !(a ?? b).f) filt('emptyCells', 1, addr);
    }
  }
  push('formulas', name, fDiff.length === 0, `${fSame}/${fTotal}${fDiff.length ? ` ; ${fDiff.length} 件: ${fDiff.slice(0, MAX_EX).join(' ; ')}` : ''}`);
  push('values', name, vDiff.length === 0, vDiff.length ? `${vDiff.length} 件: ${vDiff.slice(0, MAX_EX).join(' ; ')}` : '');
  push('styles', name, sDiff.length === 0, sDiff.length ? `${sDiff.length} セル: ${sDiff.slice(0, MAX_EX).join(' ; ')}` : '');
  if (x.dimension !== y.dimension) filt('dimension', 1, `${name} ${x.dimension} -> ${y.dimension}`);

  // ---- columns (per column, over the union of segment boundaries)
  const cuts = new Set([1]);
  for (const m of [x, y]) for (const s of m.cols) {
    cuts.add(s.min);
    cuts.add(s.max + 1);
  }
  const bounds = [...cuts].filter((c) => c <= 16385).sort((p, q) => p - q);
  const cDiff = [];
  let splitNoise = false;
  const segKey = (m) => m.cols.map((s) => `${s.min}-${s.max}`).join(',');
  if (segKey(x) !== segKey(y)) splitNoise = true;
  for (let i = 0; i < bounds.length - 1; i++) {
    const c1 = bounds[i];
    const c2 = bounds[i + 1] - 1;
    const a = colAt(x, c1);
    const b = colAt(y, c1);
    const where = c1 === c2 ? numToCol(c1) : `${numToCol(c1)}:${numToCol(c2)}`;
    const wa = a?.width ?? null;
    const wb = b?.width ?? null;
    if (wa != null && wb != null) {
      if (Math.abs(wa - wb) > 0.01) cDiff.push(`${where} 幅 ${wa} -> ${wb}`);
    } else if (wa !== wb) {
      const w = wa ?? wb;
      const custom = (a ?? b).custom;
      const def = (wa == null ? x : y).defaultColWidth;
      if (custom && !(def != null && Math.abs(def - w) <= 0.01)) cDiff.push(`${where} 幅 ${wa ?? '既定'} -> ${wb ?? '既定'}`);
      else splitNoise = true;
    }
    if (!!a?.hidden !== !!b?.hidden) cDiff.push(`${where} 非表示 ${!!a?.hidden} -> ${!!b?.hidden}`);
    if ((a?.lvl ?? 0) !== (b?.lvl ?? 0)) cDiff.push(`${where} アウトライン ${a?.lvl ?? 0} -> ${b?.lvl ?? 0}`);
    const csa = stA.xfSig(a?.s ?? 0);
    const csb = stB.xfSig(b?.s ?? 0);
    if (csa !== csb) cDiff.push(`${where} 列の書式 ${styleDiff(csa, csb)}`);
    if (!!a?.custom !== !!b?.custom) splitNoise = true;
  }
  if (splitNoise) filt('colSplit', 1, name);
  push('cols', name, cDiff.length === 0, cDiff.slice(0, MAX_EX).join(' ; '));

  // ---- rows
  const rDiff = [];
  let rowNoise = 0;
  for (const r of new Set([...x.rows.keys(), ...y.rows.keys()])) {
    let a = x.rows.get(r);
    let b = y.rows.get(r);
    // (a row without its own height has the custom default height of its sheet, if there is one)
    const withDefault = (row, m) => (row?.ht == null && m.defaultRowHeight != null ? { ...(row ?? {}), ht: m.defaultRowHeight, custom: true } : row);
    a = withDefault(a, x);
    b = withDefault(b, y);
    if (a?.custom || b?.custom) {
      if (a?.ht == null || b?.ht == null || Math.abs(a.ht - b.ht) > 0.05) rDiff.push(`${r} 行の高さ ${a?.ht ?? '自動'} -> ${b?.ht ?? '自動'}`);
    } else if ((a?.ht ?? null) !== (b?.ht ?? null)) rowNoise++;
    if (!!a?.hidden !== !!b?.hidden) rDiff.push(`${r} 行 非表示 ${!!a?.hidden} -> ${!!b?.hidden}`);
    if ((a?.lvl ?? 0) !== (b?.lvl ?? 0)) rDiff.push(`${r} 行 アウトライン ${a?.lvl ?? 0} -> ${b?.lvl ?? 0}`);
    const rsa = a?.s != null ? stA.xfSig(a.s) : '';
    const rsb = b?.s != null ? stB.xfSig(b.s) : '';
    if (rsa !== rsb) rDiff.push(`${r} 行の書式 ${styleDiff(rsa || 'なし', rsb || 'なし')}`);
  }
  if (rowNoise) filt('rowAttrs', rowNoise);
  push('rows', name, rDiff.length === 0, `${rDiff.length ? rDiff.length + ' 件: ' : ''}${rDiff.slice(0, MAX_EX).join(' ; ')}`);

  // ---- simple features
  const simple = (group, va, vb) => {
    const sa = JSON.stringify(va);
    const sb = JSON.stringify(vb);
    push(group, name, sa === sb, sa === sb ? '' : Array.isArray(va) && Array.isArray(vb) ? listDiff(va, vb) : `${sa} -> ${sb}`);
  };
  simple('merges', x.merges, y.merges);
  simple('view', x.view, y.view);
  if (x.selection !== y.selection) filt('selection', 1, name);
  simple('autoFilter', x.autoFilter, y.autoFilter);
  simple('cf', x.cf, y.cf);
  if (JSON.stringify(x.cf) === JSON.stringify(y.cf) && x.cfRaw !== y.cfRaw) filt('cfPriority', 1, `${name} ${x.cfRaw} -> ${y.cfRaw}`);
  simple('dv', x.dv, y.dv);
  simple('sparklines', x.sparklines, y.sparklines);
  simple('hyperlinks', x.hyperlinks, y.hyperlinks);
  simple('tables', x.tables, y.tables);
  simple('drawings', x.drawings, y.drawings);
  if (JSON.stringify(x.drawings) === JSON.stringify(y.drawings) && JSON.stringify(x.drawingOffsets) !== JSON.stringify(y.drawingOffsets)) filt('anchorOffsets', 1, name);
  simple('comments', x.comments, y.comments);
  simple(
    'pivots',
    x.pivots.map((p) => p.sig),
    y.pivots.map((p) => p.sig),
  );
  const la = x.pivots.map((p) => `${p.name}@${p.location}`);
  const lb = y.pivots.map((p) => `${p.name}@${p.location}`);
  if (JSON.stringify(la) !== JSON.stringify(lb)) {
    if (opts.pivotLayout) push('pivotLayout', name, false, `${la.join(',')} -> ${lb.join(',')}`);
    else filt('pivotLayout', 1, `${la.join(',')} -> ${lb.join(',')}`);
  }
  simple('protection', x.protection, y.protection);
  simple('breaks', x.breaks, y.breaks);
  simple('sheetProps', `tab=${x.tabColor}|${x.formatPr}`, `tab=${y.tabColor}|${y.formatPr}`);
}

/** Loads and compares two packages (paths or bytes). */
export async function comparePackages(a, b, opts = {}) {
  const [A, B] = await Promise.all([loadPackageModel(a), loadPackageModel(b)]);
  return comparePackageModels(A, B, opts);
}

