#!/usr/bin/env node
/**
 * Corpus of workbooks that Excel did NOT write, part 2 (round 5, module H1): ExcelJS / SheetJS / xlsx-populate output and
 * hand-made packages that copy what other producers write (Google Sheets, LibreOffice Calc, Apache POI SXSSF, the Open XML
 * SDK with `x:` prefixed elements, Mac Excel 2011 (1904 dates), Excel 2007 (ja), a CSV-to-xlsx ERP converter, Strict OOXML,
 * a package without sharedStrings / styles / theme / dimension, shared formulas, cells without `r`, backslash zip names,
 * an odd relationship prefix, absolute targets, single-quoted attributes, UTF-16 parts, ...).
 *
 *   node make-corpus.mjs <outdir>
 *
 * ExcelJS and xlsx-populate are NOT dependencies of the app: point CORPUS_MODULES at a folder that has them installed
 * (`npm i exceljs xlsx-populate` in a scratch folder; default ../../../../.tmp/h1/npm). Without them only the hand-made files
 * and the SheetJS ones (SheetJS `xlsx` is a real dependency) are written. The openpyxl / XlsxWriter / pandas half is
 * make-corpus.py. Data is deterministic (the same 氏名 / 学籍番号 as make-corpus.py, so partner.xlsx matches them).
 */
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';
import * as XLSX from 'xlsx';

const here = path.dirname(fileURLToPath(import.meta.url));
const OUT = path.resolve(process.argv[2] ?? path.join(here, 'corpus'));
fs.mkdirSync(OUT, { recursive: true });
const MODS = process.env.CORPUS_MODULES ?? path.resolve(here, '../../../../.tmp/h1/npm');
const req = createRequire(path.join(MODS, 'x.js'));
const tryRequire = (n) => { try { return req(n); } catch { return null; } };

// ------------------------------------------------------------------------------------------------ data
const SURN = ['佐藤', '鈴木', '高橋', '田中', '伊藤', '渡辺', '山本', '中村', '小林', '加藤', '吉田', '山田', '佐々木', '山口', '松本', '井上', '木村', '林', '斎藤', '清水'];
const GIVEN = ['太郎', '花子', '一郎', '美咲', '健太', '結衣', '大輔', '陽子', '翔', '愛', '拓也', '直美', '誠', '由美', '亮', '恵'];
const UNIV = ['慶應義塾大学', '早稲田大学', '東京大学', '一橋大学', '上智大学', '明治大学', '青山学院大学', '立教大学'];
const FAC = ['総合政策学部', '環境情報学部', '経済学部', '法学部', '商学部', '理工学部'];
const person = (i) => SURN[i % SURN.length] + GIVEN[(i * 7 + 3) % GIVEN.length];
const sid = (i) => 'S' + String(2024000 + i).padStart(7, '0');
const N = 30;
const R = Array.from({ length: N }, (_, k) => {
  const i = k + 1;
  return { id: sid(i), name: person(i), univ: UNIV[i % UNIV.length], fac: FAC[(i * 5) % FAC.length], score: 50 + ((i * 37) % 51), day: (i * 11) % 200, memo: i % 4 ? '' : `メモ${i}` };
});
const HEAD = ['学籍番号', '氏名', '大学', '学部', '点数', '入学日', '備考'];
/** Excel serial (1900 system) of 2024-04-01 + day */
const serial = (day, d1904 = false) => 45383 + day - (d1904 ? 1462 : 0);

const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const escA = (s) => esc(s).replace(/"/g, '&quot;');
const col = (n) => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
const XMLDECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const NSR = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CT = {
  wb: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
  ws: 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
  styles: 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml',
  sst: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml',
  theme: 'application/vnd.openxmlformats-officedocument.theme+xml',
  core: 'application/vnd.openxmlformats-package.core-properties+xml',
  app: 'application/vnd.openxmlformats-officedocument.extended-properties+xml',
  drawing: 'application/vnd.openxmlformats-officedocument.drawing+xml',
};

// ------------------------------------------------------------------------------------------------ writer helpers
async function save(name, files, opts = {}) {
  const zip = new JSZip();
  for (const [p, content] of Object.entries(files)) {
    if (content == null) continue;
    zip.file(p, content, opts.fileOpts?.(p) ?? {});
  }
  const bytes = await zip.generateAsync({
    type: 'uint8array', compression: opts.store ? 'STORE' : 'DEFLATE', streamFiles: !!opts.streamFiles, comment: opts.comment, platform: 'UNIX',
  });
  fs.writeFileSync(path.join(OUT, name), bytes);
}

/** [Content_Types].xml with Overrides for the given parts (Default rels / xml like Excel's) */
function contentTypes(parts, extraDefaults = '') {
  return XMLDECL + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
    + extraDefaults + parts.map(([p, t]) => `<Override PartName="/${p}" ContentType="${t}"/>`).join('') + '</Types>';
}
const rels = (list, ns = 'http://schemas.openxmlformats.org/package/2006/relationships') =>
  XMLDECL + `<Relationships xmlns="${ns}">` + list.map(([id, type, target, ext]) => `<Relationship Id="${id}" Type="${type}" Target="${target}"${ext ? ' TargetMode="External"' : ''}/>`).join('') + '</Relationships>';

class Sst {
  constructor() { this.list = []; this.map = new Map(); this.count = 0; }
  idx(s) { this.count++; let i = this.map.get(s); if (i === undefined) { i = this.list.length; this.list.push(s); this.map.set(s, i); } return i; }
  xml(p = '', preserve = true) {
    const si = this.list.map((s) => `<${p}si><${p}t${preserve && /^\s|\s$|\n/.test(s) ? ' xml:space="preserve"' : ''}>${esc(s)}</${p}t></${p}si>`).join('');
    return XMLDECL + `<${p}sst xmlns${p ? ':' + p.slice(0, -1) : ''}="${NS}" count="${this.count}" uniqueCount="${this.list.length}">${si}</${p}sst>`;
  }
}

/**
 * Serialize rows of cell specs into the inside of <sheetData>. A cell spec is undefined/null (skipped), a string, a number,
 * {n: number, s?} number, {d: serial, s} date, {f, v, s?, t?} formula with cached value, {b: bool}, {e: '#N/A'}, {s:<style>, v}.
 * d = { p: element prefix, inline: bool, sst: Sst, noCellR, noRowR, q: attribute quote, cellT: write t="n" explicitly,
 *       rowAttr(r): extra row attributes, cellAttr?(cell): extra cell attributes, num: number formatter, boolTrue: '1'|'true' }
 */
function sheetData(d, rows, top = 1, left = 1) {
  const p = d.p ?? '';
  const q = d.q ?? '"';
  const a = (k, v) => ` ${k}=${q}${v}${q}`;
  const num = d.num ?? String;
  let out = '';
  rows.forEach((row, ri) => {
    const r = top + ri;
    let cells = '';
    (row ?? []).forEach((spec, ci) => {
      if (spec === undefined || spec === null) return;
      const c = left + ci;
      const ref = col(c) + r;
      const sp = typeof spec === 'object' ? spec : { v: spec };
      let attrs = d.noCellR ? '' : a('r', ref);
      if (sp.s) attrs += a('s', sp.s);
      let inner = '';
      if (sp.f !== undefined) {
        const fattrs = sp.fattrs ?? '';
        const v = sp.v;
        const ft = typeof v === 'string' ? (/^#/.test(v) ? 'e' : 'str') : typeof v === 'boolean' ? 'b' : null;
        if (ft) attrs += a('t', ft);
        else if (d.cellT) attrs += a('t', 'n');
        inner = `<${p}f${fattrs}>${esc(sp.f)}</${p}f>` + (v === undefined || v === null ? '' : `<${p}v>${typeof v === 'boolean' ? (v ? 1 : 0) : esc(typeof v === 'number' ? num(v) : v)}</${p}v>`);
        if (sp.fempty) inner = `<${p}f${fattrs}/>` + `<${p}v>${num(v)}</${p}v>`;
      } else if (sp.d !== undefined) {
        if (d.cellT) attrs += a('t', 'n');
        inner = `<${p}v>${num(sp.d)}</${p}v>`;
      } else if (sp.iso !== undefined) {
        attrs += a('t', 'd');
        inner = `<${p}v>${sp.iso}</${p}v>`;
      } else if (sp.n !== undefined) {
        if (d.cellT) attrs += a('t', 'n');
        inner = `<${p}v>${num(sp.n)}</${p}v>`;
      } else if (sp.b !== undefined) {
        attrs += a('t', 'b');
        inner = `<${p}v>${sp.b ? (d.boolTrue ?? '1') : '0'}</${p}v>`;
      } else if (sp.e !== undefined) {
        attrs += a('t', 'e');
        inner = `<${p}v>${esc(sp.e)}</${p}v>`;
      } else if (typeof sp.v === 'number') {
        if (d.cellT) attrs += a('t', 'n');
        inner = `<${p}v>${num(sp.v)}</${p}v>`;
      } else if (sp.v !== undefined && sp.v !== '') {
        const text = String(sp.v);
        const preserve = /^\s|\s$|\n/.test(text) ? ' xml:space="preserve"' : '';
        if (d.inline) {
          attrs += a('t', 'inlineStr');
          inner = `<${p}is><${p}t${preserve}>${esc(text)}</${p}t></${p}is>`;
        } else {
          attrs += a('t', 's');
          inner = `<${p}v>${d.sst.idx(text)}</${p}v>`;
        }
      }
      cells += inner === '' ? `<${p}c${attrs}/>` : `<${p}c${attrs}>${inner}</${p}c>`;
    });
    const rattr = (d.noRowR ? '' : a('r', r)) + (d.rowAttr ? d.rowAttr(r) : '');
    out += `<${p}row${rattr}>${cells}</${p}row>`;
  });
  return out;
}

/** the standard roster rows: header, N data rows, blank row, 合計 row. style ids: {h, date, int, dec} */
function rosterRows(st, o = {}) {
  const d1904 = !!o.d1904;
  const rows = [HEAD.map((h) => ({ v: h, s: st.h }))];
  const bodyDate = o.dateAsText ? (day) => ({ v: `2024/${4 + Math.floor(day / 30)}/${1 + (day % 28)}`, s: st.text }) : (day) => (o.iso ? { iso: new Date(Date.UTC(2024, 3, 1 + day)).toISOString().slice(0, 10), s: st.date } : { d: serial(day, d1904), s: st.date });
  R.forEach((r, k) => {
    const row = r.score;
    rows.push([
      { v: r.id, s: st.text }, { v: r.name, s: st.text }, { v: r.univ, s: st.text }, { v: r.fac, s: st.text },
      o.scoreAsText ? { v: String(row), s: st.text } : { n: row, s: st.int }, bodyDate(r.day), r.memo ? { v: r.memo, s: st.text } : null,
      o.noFormula ? null : { f: `E${k + 2}*1.1`, v: row * 1.1, s: st.dec },
    ]);
  });
  if (o.total !== false) {
    rows.push([]);
    const last = N + 1;
    const sum = R.reduce((a, r) => a + r.score, 0);
    rows.push([{ v: '合計' }, null, null, null, { f: `SUM(E2:E${last})`, v: sum, s: st.int }, null, null, o.noFormula ? null : { f: `SUM(H2:H${last})`, v: sum * 1.1, s: st.dec }]);
  }
  return rows;
}

/** a complete hand-made package. `sheets`: [{name, xml (worksheet part text), state?, extraRels?, extraParts?}] */
async function pkg(name, o) {
  const sheets = o.sheets;
  const files = {};
  const parts = [];
  const wbRels = [];
  let rid = o.firstRid ?? 1;
  const sheetEls = [];
  sheets.forEach((s, i) => {
    const target = s.part ?? `worksheets/sheet${i + 1}.xml`;
    const id = `rId${rid++}`;
    files[`xl/${target}`.replace(/^xl\/\//, 'xl/')] = s.xml;
    parts.push([`xl/${target}`, CT.ws]);
    wbRels.push([id, `${REL}/worksheet`, o.absTargets ? `/xl/${target}` : target]);
    sheetEls.push(`<sheet name="${escA(s.name)}" sheetId="${s.sheetId ?? i + 1}"${s.state ? ` state="${s.state}"` : ''} ${o.rPrefix ?? 'r'}:id="${id}"/>`);
    for (const [p, c] of Object.entries(s.extraParts ?? {})) files[p] = c;
    for (const p of s.extraCts ?? []) parts.push(p);
  });
  if (o.styles !== null) {
    files['xl/styles.xml'] = o.styles;
    parts.push(['xl/styles.xml', CT.styles]);
    wbRels.push([`rId${rid++}`, `${REL}/styles`, o.absTargets ? '/xl/styles.xml' : 'styles.xml']);
  }
  if (o.sstXml) {
    files['xl/sharedStrings.xml'] = o.sstXml;
    parts.push(['xl/sharedStrings.xml', CT.sst]);
    wbRels.push([`rId${rid++}`, `${REL}/sharedStrings`, o.absTargets ? '/xl/sharedStrings.xml' : 'sharedStrings.xml']);
  }
  if (o.theme) {
    files['xl/theme/theme1.xml'] = o.theme;
    parts.push(['xl/theme/theme1.xml', CT.theme]);
    wbRels.push([`rId${rid++}`, `${REL}/theme`, 'theme/theme1.xml']);
  }
  const wbPath = o.wbPath ?? 'xl/workbook.xml';
  const nsDecl = o.wbNs ?? `xmlns="${NS}" xmlns:r="${NSR}"`;
  const relPrefixDecl = o.rPrefix && o.rPrefix !== 'r' ? ` xmlns:${o.rPrefix}="${NSR}"` : '';
  files[wbPath] = o.workbookXml ? o.workbookXml(sheetEls) : XMLDECL + `<workbook ${nsDecl}${relPrefixDecl}>${o.wbPr ?? '<workbookPr/>'}${o.bookViews ?? ''}<sheets>${sheetEls.join('')}</sheets>${o.definedNames ?? ''}${o.calcPr ?? ''}</workbook>`;
  parts.push([wbPath, o.wbCt ?? CT.wb]);
  files[wbPath.replace(/([^/]+)$/, '_rels/$1.rels')] = rels(wbRels);
  const root = [['rId1', `${REL}/officeDocument`, wbPath]];
  if (o.docProps) {
    files['docProps/core.xml'] = o.docProps.core;
    files['docProps/app.xml'] = o.docProps.app;
    parts.push(['docProps/core.xml', CT.core], ['docProps/app.xml', CT.app]);
    root.push(['rId2', 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties', 'docProps/core.xml'], ['rId3', `${REL}/extended-properties`, 'docProps/app.xml']);
  }
  files['_rels/.rels'] = rels(root);
  files['[Content_Types].xml'] = contentTypes(parts);
  for (const [p, c] of Object.entries(o.extraParts ?? {})) files[p] = c;
  if (o.post) o.post(files);
  await save(name, files, o.saveOpts);
}

const THEME = XMLDECL + '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="1F497D"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2><a:accent1><a:srgbClr val="4F81BD"/></a:accent1><a:accent2><a:srgbClr val="C0504D"/></a:accent2><a:accent3><a:srgbClr val="9BBB59"/></a:accent3><a:accent4><a:srgbClr val="8064A2"/></a:accent4><a:accent5><a:srgbClr val="4BACC6"/></a:accent5><a:accent6><a:srgbClr val="F79646"/></a:accent6><a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Cambria"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="ＭＳ Ｐゴシック"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="ＭＳ Ｐゴシック"/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="25400"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="38100"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>';

/** a plain Excel-like styles part. cellXfs: 0 default, 1 header (bold + fill + border), 2 date, 3 #,##0, 4 0.0, 5 text @ */
function basicStyles(o = {}) {
  const p = o.p ?? '';
  const font = o.font ?? 'Calibri';
  const dateFmt = o.dateFmt ?? 'yyyy/m/d';
  return XMLDECL + `<${p}styleSheet xmlns${p ? ':' + p.slice(0, -1) : ''}="${NS}">`
    + `<${p}numFmts count="1"><${p}numFmt numFmtId="164" formatCode="${escA(dateFmt)}"/></${p}numFmts>`
    + `<${p}fonts count="2"><${p}font><${p}sz val="11"/><${p}color theme="1"/><${p}name val="${font}"/><${p}family val="2"/></${p}font><${p}font><${p}b/><${p}sz val="11"/><${p}color theme="1"/><${p}name val="${font}"/><${p}family val="2"/></${p}font></${p}fonts>`
    + `<${p}fills count="3"><${p}fill><${p}patternFill patternType="none"/></${p}fill><${p}fill><${p}patternFill patternType="gray125"/></${p}fill><${p}fill><${p}patternFill patternType="solid"><${p}fgColor rgb="FFDDEBF7"/><${p}bgColor indexed="64"/></${p}patternFill></${p}fill></${p}fills>`
    + `<${p}borders count="2"><${p}border><${p}left/><${p}right/><${p}top/><${p}bottom/><${p}diagonal/></${p}border><${p}border><${p}left style="thin"><${p}color auto="1"/></${p}left><${p}right style="thin"><${p}color auto="1"/></${p}right><${p}top style="thin"><${p}color auto="1"/></${p}top><${p}bottom style="thin"><${p}color auto="1"/></${p}bottom><${p}diagonal/></${p}border></${p}borders>`
    + `<${p}cellStyleXfs count="1"><${p}xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></${p}cellStyleXfs>`
    + `<${p}cellXfs count="9"><${p}xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>`
    + `<${p}xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>`
    + `<${p}xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>`
    + `<${p}xf numFmtId="3" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>`
    + `<${p}xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"><${p}alignment horizontal="left"/></${p}xf>`
    + `<${p}xf numFmtId="49" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>`
    // 6: h:mm (time of day), 7: [h]:mm (duration), 8: yyyy/m/d h:mm (date + time)
    + `<${p}xf numFmtId="20" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>`
    + `<${p}xf numFmtId="46" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>`
    + `<${p}xf numFmtId="22" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/></${p}cellXfs>`
    + `<${p}cellStyles count="1"><${p}cellStyle name="Normal" xfId="0" builtinId="0"/></${p}cellStyles><${p}dxfs count="0"/><${p}tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/></${p}styleSheet>`;
}
// 0 default, 1 header, 2 date, 3 int, 4 decimal (numFmtId 0 + left alignment -> shown as General), 5 text
const ST = { h: 1, date: 2, int: 3, dec: 0, text: 5 };

const WS_OPEN = `<worksheet xmlns="${NS}" xmlns:r="${NSR}">`;
const widthCols = (ws) => '<cols>' + ws.map(([a, b, w]) => `<col min="${a}" max="${b}" width="${w}" customWidth="1"/>`).join('') + '</cols>';

// ================================================================================================ hand-made flavours
async function handmade() {
  // ---- hc-01 Google Sheets ------------------------------------------------------------------------------------
  {
    const sst = new Sst();
    const d = { sst, num: (n) => (Number.isInteger(n) ? n.toFixed(1) : String(n)) };
    const rows = rosterRows({ h: 1, date: 2, int: 3, dec: 3, text: 4 });
    // Google pads the sheet to 1000 rows of empty rows
    const total = rows.length;
    let body = sheetData(d, rows);
    for (let r = total + 1; r <= 1000; r++) body += `<row r="${r}"/>`;
    const xml = XMLDECL + `<worksheet xmlns="${NS}" xmlns:r="${NSR}"><sheetPr><outlinePr summaryBelow="0" summaryRight="0"/></sheetPr><sheetViews><sheetView workbookViewId="0"><pane xSplit="2.0" ySplit="1.0" topLeftCell="C2" activePane="bottomRight" state="frozen"/><selection activeCell="C2" sqref="C2"/><selection pane="topRight" activeCell="C1" sqref="C1"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/><selection pane="bottomRight" activeCell="C2" sqref="C2"/></sheetView></sheetViews><sheetFormatPr customHeight="1" defaultColWidth="14.43" defaultRowHeight="15.75"/><cols><col customWidth="1" min="1" max="1" width="18.0"/><col customWidth="1" min="2" max="2" width="21.0"/><col customWidth="1" min="3" max="8" width="16.29"/></cols><sheetData>${body}</sheetData><autoFilter ref="A1:H31"/><conditionalFormatting sqref="E2:E31"><cfRule type="colorScale" priority="1"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFFFFFFF"/><color rgb="FF57BB8A"/></colorScale></cfRule></conditionalFormatting><dataValidations count="1"><dataValidation type="list" allowBlank="1" showDropDown="1" sqref="C2:C31"><formula1>"${UNIV.slice(0, 3).join(',')}"</formula1></dataValidation></dataValidations><drawing r:id="rId1"/></worksheet>`;
    const gs = XMLDECL + `<styleSheet xmlns="${NS}"><numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy/mm/dd"/></numFmts><fonts count="3"><font><sz val="10.0"/><color rgb="FF000000"/><name val="Arial"/></font><font><b/><sz val="10.0"/><color rgb="FF000000"/><name val="Arial"/></font><font><sz val="10.0"/><color rgb="FF000000"/><name val="Arial"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="lightGray"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFD9EAD3"/><bgColor rgb="FFD9EAD3"/></patternFill></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf borderId="0" fillId="0" fontId="0" numFmtId="0" applyAlignment="1" applyFont="1"/></cellStyleXfs><cellXfs count="5"><xf borderId="0" fillId="0" fontId="0" numFmtId="0" xfId="0" applyAlignment="1" applyFont="1"><alignment readingOrder="0" shrinkToFit="0" vertical="bottom" wrapText="0"/></xf><xf borderId="0" fillId="2" fontId="1" numFmtId="0" xfId="0" applyAlignment="1" applyFill="1" applyFont="1"><alignment readingOrder="0"/></xf><xf borderId="0" fillId="0" fontId="2" numFmtId="164" xfId="0" applyAlignment="1" applyFont="1" applyNumberFormat="1"><alignment horizontal="right" readingOrder="0"/></xf><xf borderId="0" fillId="0" fontId="0" numFmtId="0" xfId="0" applyAlignment="1" applyFont="1"><alignment horizontal="right" readingOrder="0"/></xf><xf borderId="0" fillId="0" fontId="0" numFmtId="0" xfId="0" applyAlignment="1" applyFont="1"><alignment readingOrder="0"/></xf></cellXfs><cellStyles count="1"><cellStyle xfId="0" name="Normal" builtinId="0"/></cellStyles><dxfs count="0"/></styleSheet>`;
    await pkg('hc-01-google-sheets.xlsx', {
      sheets: [{ name: '名簿', sheetId: 1, xml, extraParts: { 'xl/drawings/drawing1.xml': XMLDECL + '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"/>', 'xl/worksheets/_rels/sheet1.xml.rels': rels([['rId1', `${REL}/drawing`, '../drawings/drawing1.xml']]) }, extraCts: [['xl/drawings/drawing1.xml', CT.drawing]] }],
      styles: gs, sstXml: sst.xml(), firstRid: 4, wbPr: '<workbookPr/>', calcPr: '<calcPr/>',
    });
  }

  // ---- hc-02 LibreOffice Calc -------------------------------------------------------------------------------------
  {
    const sst = new Sst();
    const d = { sst, cellT: true, rowAttr: () => ' customFormat="false" ht="12.8" hidden="false" customHeight="false" outlineLevel="0" collapsed="false"', boolTrue: '1' };
    const rows = rosterRows({ h: 2, date: 3, int: 4, dec: 5, text: 1 });
    rows.forEach((r, i) => { if (i === 4) r.fattrs = ''; });
    const lo = XMLDECL + `<worksheet xmlns="${NS}" xmlns:r="${NSR}"><sheetPr filterMode="false"><pageSetUpPr fitToPage="false"/></sheetPr><dimension ref="A1:H33"/><sheetViews><sheetView showFormulas="false" showGridLines="true" showRowColHeaders="true" showZeros="true" rightToLeft="false" tabSelected="true" showOutlineSymbols="true" defaultGridColor="true" view="normal" topLeftCell="A1" colorId="64" zoomScale="100" zoomScaleNormal="100" zoomScalePageLayoutView="100" workbookViewId="0"><pane xSplit="0" ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="topLeft" activeCell="A1" activeCellId="0" sqref="A1"/><selection pane="bottomLeft" activeCell="A2" activeCellId="0" sqref="A2"/></sheetView></sheetViews><sheetFormatPr defaultColWidth="11.53" defaultRowHeight="12.8" zeroHeight="false" outlineLevelRow="0" outlineLevelCol="0"></sheetFormatPr><cols><col collapsed="false" customWidth="true" hidden="false" outlineLevel="0" max="1" min="1" style="0" width="14.2"/><col collapsed="false" customWidth="true" hidden="false" outlineLevel="0" max="2" min="2" style="0" width="16.5"/><col collapsed="false" customWidth="false" hidden="false" outlineLevel="0" max="1025" min="3" style="0" width="11.53"/></cols><sheetData>${sheetData(d, rows)}</sheetData><autoFilter ref="A1:H31"/><printOptions headings="false" gridLines="false" gridLinesSet="true" horizontalCentered="false" verticalCentered="false"/><pageMargins left="0.7875" right="0.7875" top="1.05277777777778" bottom="1.05277777777778" header="0.7875" footer="0.7875"/><pageSetup paperSize="9" scale="100" fitToWidth="1" fitToHeight="1" pageOrder="downThenOver" orientation="portrait" blackAndWhite="false" draft="false" cellComments="none" firstPageNumber="1" useFirstPageNumber="true" horizontalDpi="300" verticalDpi="300" copies="1"/><headerFooter differentFirst="false" differentOddEven="false"><oddHeader>&amp;C&amp;"Times New Roman,Regular"&amp;12&amp;A</oddHeader><oddFooter>&amp;C&amp;"Times New Roman,Regular"&amp;12Page &amp;P</oddFooter></headerFooter></worksheet>`;
    const fmt = (id, code) => `<numFmt numFmtId="${id}" formatCode="${escA(code)}"/>`;
    const styles = XMLDECL + `<styleSheet xmlns="${NS}"><numFmts count="4">${fmt(164, 'GENERAL')}${fmt(165, 'YYYY/MM/DD')}${fmt(166, '#,##0')}${fmt(167, '0.0')}</numFmts><fonts count="2"><font><sz val="10"/><name val="Arial"/><family val="2"/></font><font><b val="true"/><sz val="10"/><name val="Arial"/><family val="2"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFCCCCCC"/><bgColor rgb="FF000000"/></patternFill></fill></fills><borders count="2"><border diagonalUp="false" diagonalDown="false"><left/><right/><top/><bottom/><diagonal/></border><border diagonalUp="false" diagonalDown="false"><left style="thin"><color rgb="FF000000"/></left><right style="thin"><color rgb="FF000000"/></right><top style="thin"><color rgb="FF000000"/></top><bottom style="thin"><color rgb="FF000000"/></bottom><diagonal/></border></borders><cellStyleXfs count="2"><xf numFmtId="164" fontId="0" fillId="0" borderId="0" applyFont="true" applyBorder="true" applyAlignment="true" applyProtection="true"><alignment horizontal="general" vertical="bottom" textRotation="0" wrapText="false" indent="0" shrinkToFit="false"/><protection locked="true" hidden="false"/></xf><xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyFont="true" applyBorder="true" applyAlignment="true" applyProtection="true"><alignment horizontal="general" vertical="bottom" textRotation="0" wrapText="false" indent="0" shrinkToFit="false"/><protection locked="true" hidden="false"/></xf></cellStyleXfs><cellXfs count="6"><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyFont="false" applyBorder="false" applyAlignment="false" applyProtection="false"><alignment horizontal="general" vertical="bottom" textRotation="0" wrapText="false" indent="0" shrinkToFit="false"/><protection locked="true" hidden="false"/></xf><xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyFont="false" applyBorder="true" applyAlignment="false" applyProtection="false"><alignment horizontal="general" vertical="bottom" textRotation="0" wrapText="false" indent="0" shrinkToFit="false"/><protection locked="true" hidden="false"/></xf><xf numFmtId="164" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="true" applyBorder="true" applyAlignment="false" applyProtection="false"><alignment horizontal="center" vertical="bottom" textRotation="0" wrapText="false" indent="0" shrinkToFit="false"/><protection locked="true" hidden="false"/></xf><xf numFmtId="165" fontId="0" fillId="0" borderId="1" xfId="0" applyFont="false" applyBorder="true" applyAlignment="false" applyProtection="false"><alignment horizontal="general" vertical="bottom" textRotation="0" wrapText="false" indent="0" shrinkToFit="false"/><protection locked="true" hidden="false"/></xf><xf numFmtId="166" fontId="0" fillId="0" borderId="1" xfId="0" applyFont="false" applyBorder="true" applyAlignment="false" applyProtection="false"><alignment horizontal="general" vertical="bottom" textRotation="0" wrapText="false" indent="0" shrinkToFit="false"/><protection locked="true" hidden="false"/></xf><xf numFmtId="167" fontId="0" fillId="0" borderId="1" xfId="0" applyFont="false" applyBorder="true" applyAlignment="false" applyProtection="false"><alignment horizontal="general" vertical="bottom" textRotation="0" wrapText="false" indent="0" shrinkToFit="false"/><protection locked="true" hidden="false"/></xf></cellXfs><cellStyles count="1"><cellStyle name="Excel Built-in Normal" xfId="0" builtinId="0"/></cellStyles><colors><indexedColors><rgbColor rgb="FF000000"/><rgbColor rgb="FFFFFFFF"/><rgbColor rgb="FFFF0000"/><rgbColor rgb="FF00FF00"/><rgbColor rgb="FF0000FF"/><rgbColor rgb="FFFFFF00"/><rgbColor rgb="FFFF00FF"/><rgbColor rgb="FF00FFFF"/><rgbColor rgb="FF800000"/><rgbColor rgb="FF008000"/><rgbColor rgb="FF000080"/><rgbColor rgb="FF808000"/><rgbColor rgb="FF800080"/><rgbColor rgb="FF008080"/><rgbColor rgb="FFC0C0C0"/><rgbColor rgb="FF808080"/><rgbColor rgb="FF9999FF"/><rgbColor rgb="FF993366"/><rgbColor rgb="FFFFFFCC"/><rgbColor rgb="FFCCFFFF"/><rgbColor rgb="FF660066"/><rgbColor rgb="FFFF8080"/><rgbColor rgb="FF0066CC"/><rgbColor rgb="FFCCCCFF"/><rgbColor rgb="FF000080"/><rgbColor rgb="FFFF00FF"/><rgbColor rgb="FFFFFF00"/><rgbColor rgb="FF00FFFF"/><rgbColor rgb="FF800080"/><rgbColor rgb="FF800000"/><rgbColor rgb="FF008080"/><rgbColor rgb="FF0000FF"/><rgbColor rgb="FF00CCFF"/><rgbColor rgb="FFCCFFFF"/><rgbColor rgb="FFCCFFCC"/><rgbColor rgb="FFFFFF99"/><rgbColor rgb="FF99CCFF"/><rgbColor rgb="FFFF99CC"/><rgbColor rgb="FFCC99FF"/><rgbColor rgb="FFFFCC99"/><rgbColor rgb="FF3366FF"/><rgbColor rgb="FF33CCCC"/><rgbColor rgb="FF99CC00"/><rgbColor rgb="FFFFCC00"/><rgbColor rgb="FFFF9900"/><rgbColor rgb="FFFF6600"/><rgbColor rgb="FF666699"/><rgbColor rgb="FF969696"/><rgbColor rgb="FF003366"/><rgbColor rgb="FF339966"/><rgbColor rgb="FF003300"/><rgbColor rgb="FF333300"/><rgbColor rgb="FF993300"/><rgbColor rgb="FF993366"/><rgbColor rgb="FF333399"/><rgbColor rgb="FF333333"/></indexedColors></colors></styleSheet>`;
    await pkg('hc-02-libreoffice.xlsx', {
      sheets: [{ name: '名簿', xml: lo }], styles, sstXml: sst.xml(undefined, false), firstRid: 2,
      wbPr: '<workbookPr backupFile="false" showObjectPlaceholders="false" date1904="false" dateCompatibility="true" showBorderUnselectedTables="true" filterPrivacy="false" promptedSolutions="false" showInkAnnotation="true" generator="LibreOffice" hidePivotFieldList="false" allowRefreshQuery="false" publishItems="false" checkCompatibility="false" autoCompressPictures="true" refreshAllConnections="false" defaultThemeVersion="124226"/>',
      bookViews: '<bookViews><workbookView showHorizontalScroll="true" showVerticalScroll="true" showSheetTabs="true" xWindow="0" yWindow="0" windowWidth="16384" windowHeight="8192" tabRatio="500" firstSheet="0" activeTab="0"/></bookViews>',
      definedNames: '<definedNames><definedName function="false" hidden="true" localSheetId="0" name="_xlnm._FilterDatabase" vbProcedure="false">名簿!$A$1:$H$31</definedName><definedName function="false" hidden="false" name="点数" vbProcedure="false">名簿!$E$2:$E$31</definedName></definedNames>',
      calcPr: '<calcPr iterateCount="100" refMode="A1" iterate="false" iterateDelta="0.001"/>',
      docProps: { core: XMLDECL + '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dcterms:created xsi:type="dcterms:W3CDTF">2024-10-01T09:00:00Z</dcterms:created><cp:revision>0</cp:revision></cp:coreProperties>', app: XMLDECL + '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><TotalTime>0</TotalTime><Application>LibreOffice/7.6.2.1$Windows_X86_64 LibreOffice_project/56f7684011345957bbf33a7ee678afaf4d2ba333</Application></Properties>' },
    });
  }

  // ---- hc-03 Apache POI SXSSF: inline strings, no sharedStrings, no theme, dimension A1, customWidth="true" -----------------
  {
    const d = { inline: true, num: (n) => (Number.isInteger(n) ? n.toFixed(1) : String(n)), noRowR: false };
    const rows = rosterRows({ h: 1, date: 2, int: 3, dec: 3, text: 0 });
    const xml = XMLDECL + `<worksheet xmlns="${NS}"><dimension ref="A1"/><sheetViews><sheetView workbookViewId="0"><pane ySplit="1.0" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="15.0"/><cols><col min="1" max="1" width="14.7109375" customWidth="true"/><col min="2" max="2" width="18.7109375" customWidth="true"/></cols><sheetData>${sheetData(d, rows)}</sheetData><autoFilter ref="A1:H31"/><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>`;
    const styles = XMLDECL + `<styleSheet xmlns="${NS}"><numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy/mm/dd"/></numFmts><fonts count="2"><font><sz val="11.0"/><color indexed="8"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font><font><b/><sz val="11.0"/><color indexed="8"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="darkGray"/></fill><fill><patternFill patternType="solid"><fgColor indexed="22"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="4"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="true" applyFill="true"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="true"/><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/></styleSheet>`;
    await pkg('hc-03-poi-sxssf-inline.xlsx', {
      sheets: [{ name: '名簿', xml }], styles, theme: null, firstRid: 1, wbPr: '<workbookPr date1904="false"/>', bookViews: '<bookViews><workbookView activeTab="0"/></bookViews>',
      docProps: { core: XMLDECL + '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:creator>Apache POI</dc:creator></cp:coreProperties>', app: XMLDECL + '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Apache POI</Application></Properties>' },
    });
  }

  // ---- hc-04 Open XML SDK: everything in x: prefixed elements (worksheet, workbook, styles, sharedStrings) --------------
  {
    const sst = new Sst();
    const d = { p: 'x:', sst };
    const rows = rosterRows({ h: 1, date: 2, int: 3, dec: 3, text: 5 });
    const xml = XMLDECL.replace('UTF-8', 'utf-8') + `<x:worksheet xmlns:x="${NS}" xmlns:r="${NSR}"><x:sheetViews><x:sheetView workbookViewId="0"/></x:sheetViews><x:sheetFormatPr defaultRowHeight="15"/><x:cols><x:col min="1" max="2" width="16" customWidth="1"/></x:cols><x:sheetData>${sheetData(d, rows)}</x:sheetData><x:autoFilter ref="A1:H31"/><x:mergeCells count="1"><x:mergeCell ref="J1:K2"/></x:mergeCells><x:pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></x:worksheet>`;
    const workbookXml = (els) => XMLDECL.replace('UTF-8', 'utf-8') + `<x:workbook xmlns:x="${NS}" xmlns:r="${NSR}"><x:sheets>${els.map((e) => e.replace('<sheet', '<x:sheet')).join('')}</x:sheets></x:workbook>`;
    await pkg('hc-04-openxmlsdk-prefixed.xlsx', { sheets: [{ name: '名簿', xml }], styles: basicStyles({ p: 'x:' }), sstXml: sst.xml('x:'), workbookXml });
  }

  // ---- hc-05 Mac Excel 2011: 1904 dates, baseColWidth 10, Mac page setup DPI, mx: extensions ---------------------------------
  {
    const sst = new Sst();
    const d = { sst, rowAttr: () => ' spans="1:11" x14ac:dyDescent="0"' };
    const rows = rosterRows({ h: 1, date: 2, int: 3, dec: 3, text: 0, }, { d1904: true });
    // I: time of day, J: duration, K: date + time (1904 serials: a time < 1 and a duration are NOT shifted between date systems, a date + time is)
    rows[0].push({ v: '点数x1.1', s: 1 }, { v: '始業', s: 1 }, { v: '所要', s: 1 }, { v: '受付日時', s: 1 });
    R.forEach((r, k) => rows[k + 1].push({ d: 0.25 + (k % 8) / 24, s: 6 }, { d: 1.5 + k / 24, s: 7 }, { d: serial(r.day, true) + 0.375, s: 8 }));
    const xml = XMLDECL + `<worksheet xmlns="${NS}" xmlns:r="${NSR}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac"><dimension ref="A1:H33"/><sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews><sheetFormatPr baseColWidth="10" defaultRowHeight="15" x14ac:dyDescent="0"/><cols><col min="1" max="1" width="13" customWidth="1"/><col min="2" max="4" width="18" customWidth="1"/></cols><sheetData>${sheetData(d, rows)}</sheetData><phoneticPr fontId="1" type="noConversion"/><pageMargins left="0.75" right="0.75" top="1" bottom="1" header="0.5" footer="0.5"/><pageSetup paperSize="9" orientation="portrait" horizontalDpi="4294967292" verticalDpi="4294967292"/><extLst><ext uri="{64002731-A6B0-56B0-2670-7721B7C09600}" xmlns:mx="http://schemas.microsoft.com/office/mac/excel/2008/main"><mx:PLV Mode="0" OnePage="0" WScale="0"/></ext></extLst></worksheet>`;
    await pkg('hc-05-mac-excel-2011-1904.xlsx', {
      sheets: [{ name: '名簿', xml }], styles: basicStyles({ font: 'ＭＳ Ｐゴシック' }), sstXml: sst.xml(), theme: THEME,
      wbPr: '<workbookPr date1904="1" showInkAnnotation="0" autoCompressPictures="0"/>',
      bookViews: '<bookViews><workbookView xWindow="0" yWindow="0" windowWidth="25600" windowHeight="16060" tabRatio="500"/></bookViews>',
      calcPr: '<calcPr calcId="0" concurrentCalc="0"/>',
      wbNs: `xmlns="${NS}" xmlns:r="${NSR}"`,
      docProps: { core: XMLDECL + '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:creator>User</dc:creator></cp:coreProperties>', app: XMLDECL + '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Microsoft Macintosh Excel</Application><AppVersion>14.0300</AppVersion></Properties>' },
    });
  }

  // ---- hc-06 Excel 2007 (ja): Japanese built-in number formats, era dates, MS P Gothic, no x14ac ---------------------------
  {
    const sst = new Sst();
    const d = { sst, rowAttr: () => ' spans="1:9"' };
    const builtin = [27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 50, 51, 52, 53, 54, 55, 56, 57, 58];
    const head = ['学籍番号', '氏名', ...builtin.map((i) => `fmt${i}`)];
    const rows = [head.map((h) => ({ v: h, s: 1 }))];
    R.slice(0, 20).forEach((r) => {
      rows.push([{ v: r.id, s: 5 }, { v: r.name, s: 5 }, ...builtin.map((_, bi) => ({ d: serial(r.day) + (bi === 5 ? 0.5 : 0), s: 6 + bi }))]);
    });
    const xml = XMLDECL + `<worksheet xmlns="${NS}" xmlns:r="${NSR}"><dimension ref="A1:U21"/><sheetViews><sheetView tabSelected="1" workbookViewId="0"><selection activeCell="A1" sqref="A1"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="13.5"/><cols><col min="1" max="2" width="13" customWidth="1"/></cols><sheetData>${sheetData(d, rows)}</sheetData><phoneticPr fontId="1"/><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><pageSetup paperSize="9" orientation="portrait"/></worksheet>`;
    const xfs = builtin.map((id) => `<xf numFmtId="${id}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"><alignment vertical="center"/></xf>`).join('');
    const styles = XMLDECL + `<styleSheet xmlns="${NS}"><numFmts count="1"><numFmt numFmtId="176" formatCode="[$-411]ggge&quot;年&quot;m&quot;月&quot;d&quot;日&quot;;@"/></numFmts><fonts count="2"><font><sz val="11"/><name val="ＭＳ Ｐゴシック"/><family val="3"/><charset val="128"/></font><font><b/><sz val="11"/><name val="ＭＳ Ｐゴシック"/><family val="3"/><charset val="128"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor indexed="22"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"><alignment vertical="center"/></xf></cellStyleXfs><cellXfs count="${6 + builtin.length}"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"><alignment vertical="center"/></xf><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"><alignment vertical="center"/></xf><xf numFmtId="176" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"><alignment vertical="center"/></xf><xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"><alignment vertical="center"/></xf><xf numFmtId="3" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"><alignment vertical="center"/></xf><xf numFmtId="49" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"><alignment vertical="center"/></xf>${xfs}</cellXfs><cellStyles count="1"><cellStyle name="標準" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/></styleSheet>`;
    await pkg('hc-06-excel2007-ja-numfmts.xlsx', {
      sheets: [{ name: '名簿', xml }],
      styles, sstXml: sst.xml(), theme: THEME, wbPr: '<workbookPr defaultThemeVersion="124226"/>',
      bookViews: '<bookViews><workbookView xWindow="120" yWindow="30" windowWidth="18780" windowHeight="8115"/></bookViews>', calcPr: '<calcPr calcId="124519"/>',
      extraParts: { },
      docProps: { core: XMLDECL + '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:creator>事務</dc:creator></cp:coreProperties>', app: XMLDECL + '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Microsoft Excel</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop><AppVersion>12.0000</AppVersion></Properties>' },
    });
  }

  // ---- hc-07 ERP / CSV-to-xlsx converter: everything text, numbers and dates as text, header with line breaks, ﾊﾝｶｸ sheet name --------------
  {
    const d = { inline: true };
    const rows = [['学籍\n番号', '氏　名', '大学名', '点数(点)', '入学日', '電話番号', 'ﾌﾘｶﾞﾅ'].map((h) => ({ v: h, s: 1 }))];
    R.forEach((r, i) => rows.push([{ v: r.id, s: 0 }, { v: r.name, s: 0 }, { v: r.univ, s: 0 }, { v: String(r.score), s: 0 }, { v: `2024/04/${String(1 + (r.day % 28)).padStart(2, '0')}`, s: 0 }, { v: `0${3 + (i % 6)}-1234-${String(1000 + i)}`, s: 0 }, { v: ' ﾊﾝｶｸ ' + (i % 3), s: 0 }]));
    const xml = XMLDECL + `<worksheet xmlns="${NS}"><sheetFormatPr defaultRowHeight="13.5"/><cols><col min="1" max="1" width="255" customWidth="1"/><col min="2" max="7" width="40" customWidth="1"/></cols><sheetData>${sheetData(d, rows)}</sheetData></worksheet>`;
    await pkg('hc-07-erp-all-text.xlsx', { sheets: [{ name: 'ﾃﾞｰﾀ', xml }], styles: basicStyles({ font: 'ＭＳ ゴシック' }), wbPr: '' });
  }

  // ---- hc-08 Strict OOXML (purl.oclc.org namespaces, ISO dates t="d") ----------------------------------------------------------
  {
    const sst = new Sst();
    const d = { sst };
    const rows = rosterRows({ h: 1, date: 2, int: 3, dec: 3, text: 5 }, { iso: true });
    const NSS = 'http://purl.oclc.org/ooxml/spreadsheetml/main';
    const NSRS = 'http://purl.oclc.org/ooxml/officeDocument/relationships';
    const xml = XMLDECL + `<worksheet xmlns="${NSS}" xmlns:r="${NSRS}"><dimension ref="A1:H33"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData>${sheetData(d, rows)}</sheetData><autoFilter ref="A1:H31"/></worksheet>`;
    const strictStyles = basicStyles().replace(NS, NSS);
    await pkg('hc-08-strict-ooxml.xlsx', {
      sheets: [{ name: '名簿', xml }], styles: strictStyles, sstXml: sst.xml().replace(NS, NSS),
      wbNs: `xmlns="${NSS}" xmlns:r="${NSRS}"`, wbPr: '<workbookPr dateCompatibility="0" defaultThemeVersion="124226"/>',
      post: (files) => {
        for (const p of ['xl/_rels/workbook.xml.rels', '_rels/.rels']) files[p] = files[p].replaceAll(REL, NSRS);
      },
    });
  }

  // ---- hc-09 minimal: no sharedStrings, no styles, no theme, no dimension, no docProps, inline strings -------------------------------
  {
    const d = { inline: true };
    const rows = rosterRows({ h: 0, date: 0, int: 0, dec: 0, text: 0 }, { total: false });
    const xml = `<worksheet xmlns="${NS}"><sheetData>${sheetData(d, rows.map((r) => r.map((c) => (c && c.d !== undefined ? { v: c.d } : c))))}</sheetData></worksheet>`;
    await pkg('hc-09-minimal-no-parts.xlsx', { sheets: [{ name: 'Sheet1', xml }], styles: null, wbPr: '' });
  }

  // ---- hc-10 shared formulas (Excel style si / ref) + shared formula without master text + array formula -------------------------------
  {
    const sst = new Sst();
    const d = { sst };
    const rows = rosterRows({ h: 1, date: 2, int: 3, dec: 3, text: 5 }, { total: false });
    rows.forEach((r, i) => {
      if (i >= 1 && i <= N) r[7] = i === 1 ? { f: 'E2*1.1', v: R[0].score * 1.1, s: 3, fattrs: ` t="shared" ref="H2:H${N + 1}" si="0"` } : { f: '', fempty: true, v: R[i - 1].score * 1.1, s: 3, fattrs: ' t="shared" si="0"' };
      if (i >= 1 && i <= N) r[8] = i === 1 ? { f: 'IF(E2>=80,"優","可")&"/"&$B2', v: '可/x', s: 5, fattrs: ` t="shared" ref="I2:I${N + 1}" si="1"` } : { f: '', fempty: true, v: '優/x', s: 5, fattrs: ' t="shared" si="1"' };
    });
    rows.push([]);
    rows.push([{ v: '合計' }, null, null, null, { f: `SUM(E2:E${N + 1})`, v: 0, s: 3 }, null, null, { f: `SUM(H2:H${N + 1})`, v: 0, s: 3 }]);
    const xml = XMLDECL + `<worksheet xmlns="${NS}" xmlns:r="${NSR}"><dimension ref="A1:I33"/><sheetData>${sheetData(d, rows)}</sheetData></worksheet>`;
    await pkg('hc-10-shared-formulas.xlsx', { sheets: [{ name: '名簿', xml }], styles: basicStyles(), sstXml: sst.xml(), theme: THEME });
  }

  // ---- hc-11 cells and rows WITHOUT r attributes (some generators omit them) --------------------------------------------------------------
  {
    const sst = new Sst();
    const d = { sst, noCellR: true, noRowR: true };
    const rows = rosterRows({ h: 1, date: 2, int: 3, dec: 3, text: 5 }, { total: false, noFormula: true }).map((r) => r.map((c) => c ?? { v: '' }));
    const xml = XMLDECL + `<worksheet xmlns="${NS}"><sheetData>${sheetData(d, rows)}</sheetData></worksheet>`;
    await pkg('hc-11-no-r-attributes.xlsx', { sheets: [{ name: '名簿', xml }], styles: basicStyles(), sstXml: sst.xml() });
  }

  // ---- hc-12 zip entries with BACKSLASH names (old .NET Framework ZipFile) + STORE + data descriptors ----------------------------------------
  {
    const sst = new Sst();
    const rows = rosterRows({ h: 1, date: 2, int: 3, dec: 3, text: 5 }, { noFormula: true });
    const xml = XMLDECL + `<worksheet xmlns="${NS}" xmlns:r="${NSR}"><sheetData>${sheetData({ sst }, rows)}</sheetData></worksheet>`;
    await pkg('hc-12-zip-backslash-stored.xlsx', {
      sheets: [{ name: '名簿', xml }], styles: basicStyles(), sstXml: sst.xml(),
      saveOpts: { store: true, streamFiles: true, comment: 'made by an old zip writer' },
      post: (files) => {
        for (const p of Object.keys(files)) if (/^xl\//.test(p)) { files[p.replaceAll('/', '\\')] = files[p]; delete files[p]; }
      },
    });
  }

  // ---- hc-13 relationship namespace with an odd prefix (rel:id) + workbook with ns-prefixed default + spaces around = ---------------------------
  {
    const sst = new Sst();
    const rows = rosterRows({ h: 1, date: 2, int: 3, dec: 3, text: 5 }, { noFormula: true });
    const xml = XMLDECL + `<worksheet xmlns="${NS}" xmlns:rel="${NSR}"><sheetData>${sheetData({ sst }, rows)}</sheetData></worksheet>`;
    await pkg('hc-13-odd-rel-prefix.xlsx', { sheets: [{ name: '名簿', xml }], styles: basicStyles(), sstXml: sst.xml(), rPrefix: 'rel', wbNs: `xmlns="${NS}"` });
  }

  // ---- hc-14 absolute part targets (/xl/worksheets/sheet1.xml) and a workbook part that is not xl/workbook.xml ---------------------------------
  {
    const sst = new Sst();
    const rows = rosterRows({ h: 1, date: 2, int: 3, dec: 3, text: 5 }, { noFormula: true });
    const xml = XMLDECL + `<worksheet xmlns="${NS}" xmlns:r="${NSR}"><sheetData>${sheetData({ sst }, rows)}</sheetData></worksheet>`;
    await pkg('hc-14-absolute-targets-wb2.xlsx', { sheets: [{ name: '名簿', xml }], styles: basicStyles(), sstXml: sst.xml(), absTargets: true, wbPath: 'xl/workbook2.xml' });
  }

  // ---- hc-15 single-quoted attributes everywhere ------------------------------------------------------------------------------------------------
  {
    const sst = new Sst();
    const rows = rosterRows({ h: 1, date: 2, int: 3, dec: 3, text: 5 });
    const xml = `<?xml version='1.0' encoding='UTF-8' standalone='yes'?>\n<worksheet xmlns='${NS}' xmlns:r='${NSR}'><sheetViews><sheetView workbookViewId='0'/></sheetViews><cols><col min='1' max='2' width='16' customWidth='1'/></cols><sheetData>${sheetData({ sst, q: "'" }, rows)}</sheetData><autoFilter ref='A1:H31'/></worksheet>`;
    await pkg('hc-15-single-quote-attrs.xlsx', {
      sheets: [{ name: '名簿', xml }], styles: basicStyles(), sstXml: sst.xml(),
      post: (files) => { files['xl/workbook.xml'] = files['xl/workbook.xml'].replace(/="([^"]*)"/g, "='$1'"); },
    });
  }

  // ---- hc-16 UTF-16 (BOM) XML parts ------------------------------------------------------------------------------------------------------------------
  {
    const sst = new Sst();
    const rows = rosterRows({ h: 1, date: 2, int: 3, dec: 3, text: 5 }, { noFormula: true });
    const xml = XMLDECL.replace('UTF-8', 'UTF-16') + `<worksheet xmlns="${NS}" xmlns:r="${NSR}"><sheetData>${sheetData({ sst }, rows)}</sheetData></worksheet>`;
    const u16 = (s) => { const b = Buffer.from('﻿' + s, 'utf16le'); return new Uint8Array(b); };
    await pkg('hc-16-utf16-parts.xlsx', {
      sheets: [{ name: '名簿', xml: u16(xml) }], styles: basicStyles(), sstXml: sst.xml(),
      post: (files) => { files['xl/sharedStrings.xml'] = u16(files['xl/sharedStrings.xml'].replace('encoding="UTF-8"', 'encoding="UTF-16"')); },
    });
  }

  // ---- hc-17 sheet names with apostrophes / spaces / & / 31 chars, hidden FIRST sheet, veryHidden sheet, formulas across sheets -------------------------------
  {
    const sst = new Sst();
    const d = { sst };
    const longName = "O'Brien & 名簿 (2024年度 前期 成績一覧 試験)".slice(0, 31);
    const q = "'" + longName.replace(/'/g, "''") + "'";
    const rows = rosterRows({ h: 1, date: 2, int: 3, dec: 3, text: 5 }, { total: false });
    const main = XMLDECL + `<worksheet xmlns="${NS}" xmlns:r="${NSR}"><sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews><sheetData>${sheetData(d, rows)}</sheetData><autoFilter ref="A1:H31"/></worksheet>`;
    const master = XMLDECL + `<worksheet xmlns="${NS}" xmlns:r="${NSR}"><sheetData>${sheetData(d, [[{ v: '大学' }], ...UNIV.map((u) => [{ v: u }])])}</sheetData></worksheet>`;
    const summ = XMLDECL + `<worksheet xmlns="${NS}" xmlns:r="${NSR}"><sheetData>${sheetData(d, [[{ v: '平均' }, { f: `AVERAGE(${q}!E2:E31)`, v: 75 }], [{ v: '人数' }, { f: `COUNTA(${q}!A2:A31)`, v: 30 }], [{ v: '最高' }, { f: `MAX(${q}!$E$2:$E$31)`, v: 100 }]])}</sheetData></worksheet>`;
    await pkg('hc-17-sheetnames-hidden-first.xlsx', {
      sheets: [{ name: 'マスタ', xml: master, state: 'hidden' }, { name: longName, xml: main }, { name: '集計 2024', xml: summ }, { name: '裏', xml: master, state: 'veryHidden' }],
      styles: basicStyles(), sstXml: sst.xml(), theme: THEME,
      bookViews: '<bookViews><workbookView activeTab="1"/></bookViews>',
      definedNames: `<definedNames><definedName name="大学一覧">マスタ!$A$2:$A$9</definedName><definedName name="_xlnm._FilterDatabase" localSheetId="1" hidden="1">${esc(q)}!$A$1:$H$31</definedName></definedNames>`,
    });
  }

  // ---- hc-18 numbers in odd (but valid) lexical forms + whitespace in <v> + bool as true / 1 + error cells + t="str" constants + rich inline strings --------------
  // hc-22: the same with lexical forms Excel itself rejects (+5, 0x10, NaN, 1,234, full-width digits, TRUE in a boolean cell): it asks to repair such
  // a file, the app must still read / merge it without failing (the invalid text is ignored / kept as text)
  for (const invalid of [false, true]) {
    const lines = ['<row r="1"><c r="A1" t="inlineStr"><is><t>学籍番号</t></is></c><c r="B1" t="inlineStr"><is><t>氏名</t></is></c><c r="C1" t="inlineStr"><is><t>値</t></is></c><c r="D1" t="inlineStr"><is><t>メモ</t></is></c></row>'];
    const odd = invalid ? ['+5', '0x10', 'NaN', '1,234', '１２', '1E2', '.5', '5.', '0012', ' 7 ', '-0', '1e-3'] : ['1E2', '.5', '5.', '0012', ' 7 ', '-0', '1e-3', '1e2', '1.0E+2', '00', '0.1e1', '-.5'];
    R.slice(0, 12).forEach((r, i) => {
      const n = i + 2;
      const bool = invalid ? '<v>TRUE</v>' : i % 2 ? '<v>true</v>' : '<v>1</v>';
      lines.push(`<row r="${n}"><c r="A${n}" t="inlineStr"><is><t>${r.id}</t></is></c><c r="B${n}" t="str"><v>${r.name}</v></c><c r="C${n}"><v>${odd[i]}</v></c><c r="D${n}" t="${i % 3 === 0 ? 'b' : i % 3 === 1 ? 'e' : 'inlineStr'}">${i % 3 === 0 ? bool : i % 3 === 1 ? '<v>#DIV/0!</v>' : '<is><r><t>リッチ</t></r><r><rPr><b/></rPr><t xml:space="preserve"> 太字 </t></r></is>'}</c></row>`);
    });
    await pkg(invalid ? 'hc-22-invalid-lexical-values.xlsx' : 'hc-18-odd-lexical-values.xlsx', { sheets: [{ name: '名簿', xml: XMLDECL + `<worksheet xmlns="${NS}"><sheetData>${lines.join('')}</sheetData></worksheet>` }], styles: basicStyles() });
  }

  // ---- hc-19 two-row header with merged group cells + title band whose style is on the top-left only + frozen panes at the data ------------------------------------
  {
    const sst = new Sst();
    const d = { sst };
    const grp = [{ v: '基本情報', s: 1 }, { s: 1 }, { s: 1 }, { s: 1 }, { v: '成績', s: 1 }, { s: 1 }, { v: '備考', s: 1 }];
    const rows = [[{ v: '2024年度 学生名簿', s: 1 }], [], grp, ...rosterRows({ h: 1, date: 2, int: 3, dec: 3, text: 5 }, { total: false, noFormula: true })];
    const xml = XMLDECL + `<worksheet xmlns="${NS}" xmlns:r="${NSR}"><sheetViews><sheetView workbookViewId="0"><pane xSplit="1" ySplit="4" topLeftCell="B5" activePane="bottomRight" state="frozen"/></sheetView></sheetViews><sheetData>${sheetData(d, rows)}</sheetData><mergeCells count="4"><mergeCell ref="A1:G1"/><mergeCell ref="A3:D3"/><mergeCell ref="E3:F3"/><mergeCell ref="G3:G3"/></mergeCells></worksheet>`.replace('<mergeCell ref="G3:G3"/>', '').replace('count="4"', 'count="3"');
    await pkg('hc-19-merged-two-row-header.xlsx', { sheets: [{ name: '名簿', xml }], styles: basicStyles(), sstXml: sst.xml() });
  }

  // ---- hc-20 sparse huge formatted tail: styled empty cells thousands of rows / columns away, row styles, column styles ----------------------------------------------
  {
    const sst = new Sst();
    const d = { sst };
    const rows = rosterRows({ h: 1, date: 2, int: 3, dec: 3, text: 5 }, { total: false });
    const tail = '<row r="5000" s="3" customFormat="1"><c r="A5000" s="3"/><c r="Z5000" s="2"/></row><row r="1048576"><c r="XFD1048576" s="1"/></row>';
    const xml = XMLDECL + `<worksheet xmlns="${NS}" xmlns:r="${NSR}"><dimension ref="A1:XFD1048576"/><cols><col min="1" max="1" width="14" style="5" customWidth="1"/><col min="9" max="16384" width="9" style="3"/></cols><sheetData>${sheetData(d, rows)}${tail}</sheetData></worksheet>`;
    await pkg('hc-20-sparse-formatted-tail.xlsx', { sheets: [{ name: '名簿', xml }], styles: basicStyles(), sstXml: sst.xml() });
  }

  // ---- hc-21 ClosedXML / EPPlus style: sheetPr outline, tabColor, defined names with comments, row heights, rich strings in sharedStrings -----------------------------
  {
    const sst = new Sst();
    const d = { sst, rowAttr: (r) => (r === 1 ? ' ht="30" customHeight="1"' : '') };
    const rows = rosterRows({ h: 1, date: 2, int: 3, dec: 3, text: 5 });
    // a rich text shared string (two runs) used by the memo of row 2
    rows[1][6] = { v: 'RICH' };
    const body = sheetData(d, rows);
    let x = sst.xml();
    x = x.replace('<si><t>RICH</t></si>', '<si><r><rPr><b/><sz val="11"/><color rgb="FFFF0000"/><rFont val="Calibri"/><family val="2"/></rPr><t>重要</t></r><r><rPr><sz val="11"/><color theme="1"/><rFont val="Calibri"/><family val="2"/></rPr><t xml:space="preserve"> 連絡済</t></r><rPh sb="0" eb="2"><t>ジュウヨウ</t></rPh><phoneticPr fontId="1"/></si>');
    const xml = XMLDECL + `<worksheet xmlns="${NS}" xmlns:r="${NSR}"><sheetPr><tabColor rgb="FF00B050"/><outlinePr summaryBelow="1" summaryRight="1"/></sheetPr><dimension ref="A1:H33"/><sheetViews><sheetView tabSelected="1" workbookViewId="0"><selection activeCell="A1" sqref="A1"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="15" outlineLevelRow="1"/><cols><col min="1" max="1" width="14.7109375" bestFit="1" customWidth="1"/><col min="2" max="2" width="10.7109375" bestFit="1" customWidth="1"/></cols><sheetData>${body}</sheetData><autoFilter ref="A1:H31"/><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>`;
    await pkg('hc-21-closedxml-rich-strings.xlsx', { sheets: [{ name: '名簿', xml }], styles: basicStyles(), sstXml: x, theme: THEME, definedNames: '<definedNames><definedName name="点数" comment="点数の列">名簿!$E$2:$E$31</definedName></definedNames>' });
  }

  // ---- hc-23 part names in a different case than the relationships / content types say (OPC part names are case-insensitive) ---------------------------------------
  {
    const sst = new Sst();
    const rows = rosterRows({ h: 1, date: 2, int: 3, dec: 3, text: 5 }, { noFormula: true });
    const xml = XMLDECL + `<worksheet xmlns="${NS}" xmlns:r="${NSR}"><sheetData>${sheetData({ sst }, rows)}</sheetData><autoFilter ref="A1:H31"/></worksheet>`;
    await pkg('hc-23-part-name-case.xlsx', {
      sheets: [{ name: '名簿', xml }], styles: basicStyles(), sstXml: sst.xml(),
      post: (files) => {
        const ren = (from, to) => { files[to] = files[from]; delete files[from]; };
        ren('xl/worksheets/sheet1.xml', 'XL/Worksheets/Sheet1.xml');
        ren('xl/sharedStrings.xml', 'xl/SharedStrings.XML');
        ren('xl/styles.xml', 'xl/Styles.xml');
      },
    });
  }

  // ---- hc-24 Japanese part names (raw UTF-8 entry names; relationship targets percent-encoded / raw) -------------------------------------------------------------
  {
    const sst = new Sst();
    const rows = rosterRows({ h: 1, date: 2, int: 3, dec: 3, text: 5 }, { noFormula: true });
    const xml = XMLDECL + `<worksheet xmlns="${NS}" xmlns:r="${NSR}"><sheetData>${sheetData({ sst }, rows)}</sheetData></worksheet>`;
    await pkg('hc-24-japanese-part-names.xlsx', {
      sheets: [{ name: '名簿', xml, part: 'worksheets/名簿 1.xml' }], styles: basicStyles(), sstXml: sst.xml(),
      post: (files) => {
        // the workbook rels point at the part percent-encoded; the zip entry name is raw
        files['xl/_rels/workbook.xml.rels'] = files['xl/_rels/workbook.xml.rels'].replace('worksheets/名簿 1.xml', encodeURI('worksheets/名簿 1.xml'));
      },
    });
  }

  // ---- hc-25 sheets whose part order / rId order differ from the tab order, sheetId not increasing, a chartsheet-less workbook with a first sheet that is NOT the data ---------
  {
    const sst = new Sst();
    const d = { sst };
    const rows = rosterRows({ h: 1, date: 2, int: 3, dec: 3, text: 5 }, { total: false });
    const data = XMLDECL + `<worksheet xmlns="${NS}" xmlns:r="${NSR}"><sheetData>${sheetData(d, rows)}</sheetData><autoFilter ref="A1:H31"/></worksheet>`;
    const notes = XMLDECL + `<worksheet xmlns="${NS}" xmlns:r="${NSR}"><sheetData>${sheetData(d, [[{ v: '注意事項' }], [{ v: 'このブックは名簿を含みます。' }], [{ v: '更新日' }, { d: serial(100), s: 2 }]])}</sheetData></worksheet>`;
    await pkg('hc-25-sheet-order-ids.xlsx', {
      sheets: [{ name: '説明', xml: notes, part: 'worksheets/sheet9.xml', sheetId: 7 }, { name: '名簿', xml: data, part: 'worksheets/sheet2.xml', sheetId: 3 }],
      styles: basicStyles(), sstXml: sst.xml(), firstRid: 10,
    });
  }

  // ---- hc-26 CDATA sections, comments and processing instructions inside <sheetData>, numeric character references, _xHHHH_ escapes, line breaks --------------------
  {
    const rowsX = ['<row r="1"><c r="A1" t="inlineStr"><is><t>学籍番号</t></is></c><c r="B1" t="inlineStr"><is><t>氏名</t></is></c><c r="C1" t="inlineStr"><is><t>メモ</t></is></c><c r="D1" t="inlineStr"><is><t>点数</t></is></c></row>'];
    R.slice(0, 14).forEach((r, i) => {
      const n = i + 2;
      const memo = [
        '<![CDATA[a < b & c > d]]>', '&#x3042;&#12356;&#x1F600;', 'line1&#10;line2', 'cr_x000D_lf&#10;', '<![CDATA[  padded  ]]>', 'tab&#9;tab', '&lt;tag&gt; &amp; &quot;q&quot; &apos;a&apos;', '_x0041_ literal', '<![CDATA[]]>', '漢字<![CDATA[ & ]]>混在', '𠮷野家', 'a<!-- inner comment -->b', '&#xA;', 'x',
      ][i];
      const isT = memo.includes('padded') || memo.includes('  ') ? ' xml:space="preserve"' : '';
      rowsX.push(`<!-- row ${n} --><row r="${n}"><c r="A${n}" t="inlineStr"><is><t><![CDATA[${r.id}]]></t></is></c><c r="B${n}" t="inlineStr"><is><t>${r.name}</t></is></c><c r="C${n}" t="inlineStr"><is><t${isT}>${memo}</t></is></c><?pi data?><c r="D${n}"><v>${r.score}</v></c></row>`);
    });
    await pkg('hc-26-cdata-comments-entities.xlsx', { sheets: [{ name: '名簿', xml: XMLDECL + `<worksheet xmlns="${NS}"><!-- sheet comment --><sheetData>${rowsX.join('')}</sheetData></worksheet>` }], styles: basicStyles() });
  }

  // ---- hc-27 custom indexed colour palette (<colors><indexedColors>) used by fills / fonts through indexed="n" ---------------------------------------------------
  {
    const sst = new Sst();
    const d = { sst };
    const rows = rosterRows({ h: 1, date: 2, int: 3, dec: 3, text: 5 }, { total: false, noFormula: true });
    const pal = Array.from({ length: 56 }, (_, i) => `<rgbColor rgb="FF${(0x101010 * (i % 16) + 0x0a0a0a).toString(16).padStart(6, '0').slice(-6).toUpperCase()}"/>`);
    pal[22] = '<rgbColor rgb="FF00AA00"/>'; // index 22 (light gray in the default palette) is green here
    pal[10] = '<rgbColor rgb="FFAA0000"/>';
    let styles = basicStyles();
    styles = styles.replace('<fgColor rgb="FFDDEBF7"/><bgColor indexed="64"/>', '<fgColor indexed="22"/><bgColor indexed="64"/>')
      .replace('<color theme="1"/><name val="Calibri"/><family val="2"/></font><font><b/><sz val="11"/><color theme="1"/>', '<color indexed="10"/><name val="Calibri"/><family val="2"/></font><font><b/><sz val="11"/><color indexed="10"/>')
      .replace('</styleSheet>', `<colors><indexedColors>${pal.join('')}</indexedColors></colors></styleSheet>`);
    const xml = XMLDECL + `<worksheet xmlns="${NS}" xmlns:r="${NSR}"><sheetData>${sheetData(d, rows)}</sheetData></worksheet>`;
    await pkg('hc-27-custom-indexed-palette.xlsx', { sheets: [{ name: '名簿', xml }], styles, sstXml: sst.xml(), theme: THEME });
  }

  // ---- hc-28 a ListObject with awkward column names (line break, brackets, #, ', &, parentheses, full-width space) and a calculated column -----------------------------
  {
    const sst = new Sst();
    const d = { sst };
    const names = ['学籍番号', '氏名', '点数(点)', '備考\n(任意)', '率[%]', '#件数', "It's", 'A&B', '加点　表'];
    const xname = (s) => esc(s).replace(/"/g, '&quot;').replace(/\n/g, '_x000a_');
    const rows = [names.map((h) => ({ v: h, s: 1 }))];
    R.slice(0, 10).forEach((r, i) => rows.push([{ v: r.id }, { v: r.name }, { n: r.score }, { v: i % 2 ? 'a\nb' : '' }, { n: 0.5 }, { n: i }, { v: 'x' }, { v: 'y' }, { f: 'Tbl特殊[[#This Row],[点数(点)]]+1', v: r.score + 1 }]));
    const table = XMLDECL + `<table xmlns="${NS}" id="1" name="Tbl特殊" displayName="Tbl特殊" ref="A1:I11" totalsRowShown="0"><autoFilter ref="A1:I11"/><tableColumns count="9">${names.map((h, i) => `<tableColumn id="${i + 1}" name="${xname(h)}"${i === 8 ? '><calculatedColumnFormula>Tbl特殊[[#This Row],[点数(点)]]+1</calculatedColumnFormula></tableColumn>' : '/>'}`).join('')}</tableColumns><tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>`;
    const xml = XMLDECL + `<worksheet xmlns="${NS}" xmlns:r="${NSR}"><dimension ref="A1:I11"/><sheetData>${sheetData(d, rows)}</sheetData><tableParts count="1"><tablePart r:id="rId1"/></tableParts></worksheet>`;
    await pkg('hc-28-table-awkward-names.xlsx', {
      sheets: [{ name: '名簿', xml, extraParts: { 'xl/worksheets/_rels/sheet1.xml.rels': rels([['rId1', `${REL}/table`, '../tables/table1.xml']]), 'xl/tables/table1.xml': table }, extraCts: [['xl/tables/table1.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml']] }],
      styles: basicStyles(), sstXml: sst.xml(), theme: THEME,
    });
  }

  // ---- hc-29 formulas: cross-sheet, defined names, whole-column refs, array (CSE) formula, string results, error-free fallbacks -------------------------------------------
  {
    const sst = new Sst();
    const d = { sst };
    const rows = rosterRows({ h: 1, date: 2, int: 3, dec: 3, text: 5 }, { total: false, noFormula: true });
    for (let i = 1; i <= 12; i++) {
      rows[i][7] = { f: `VLOOKUP(A${i + 1},マスタ!$A$2:$B$9,2,FALSE)`, v: '慶應', s: 5 };
      rows[i][8] = { f: `点数列*0+E${i + 1}`, v: R[i - 1].score, s: 3 };
      rows[i][9] = { f: `IFERROR(1/(E${i + 1}-${R[i - 1].score}),"zero")`, v: 'zero', s: 5 };
      rows[i][10] = { f: `SUMPRODUCT((A$2:A$31=A${i + 1})*(E$2:E$31))`, v: R[i - 1].score, s: 3 };
    }
    rows[13] = [{ v: '集計' }, null, null, null, { f: 'SUM(E:E)', v: 0, s: 3 }, null, null, { f: 'COUNTA(A:A)', v: 31, s: 3 }];
    rows[14] = [{ v: '配列' }, null, null, null, { f: 'SUM(E2:E31*2)', v: 0, s: 3, fattrs: ' t="array" ref="E15"' }];
    const master = [[{ v: '大学' }, { v: '略称' }], ...UNIV.map((u, i) => [{ v: u }, { v: u.slice(0, 2) + i }])];
    const s1 = XMLDECL + `<worksheet xmlns="${NS}" xmlns:r="${NSR}"><sheetData>${sheetData(d, rows)}</sheetData></worksheet>`;
    const s2 = XMLDECL + `<worksheet xmlns="${NS}" xmlns:r="${NSR}"><sheetData>${sheetData(d, master)}</sheetData></worksheet>`;
    const s3 = XMLDECL + `<worksheet xmlns="${NS}" xmlns:r="${NSR}"><sheetData>${sheetData(d, [[{ v: '平均' }, { f: 'AVERAGE(名簿!E2:E31)', v: 75 }], [{ v: '最大' }, { f: 'MAX(点数列)', v: 100 }], [{ v: '人数' }, { f: 'COUNTA(名簿!$A$2:$A$31)', v: 30 }], [{ v: '名簿先頭' }, { f: '名簿!A2', v: R[0].id }]])}</sheetData></worksheet>`;
    await pkg('hc-29-formula-variety.xlsx', {
      sheets: [{ name: '名簿', xml: s1 }, { name: 'マスタ', xml: s2 }, { name: '集計', xml: s3 }],
      styles: basicStyles(), sstXml: sst.xml(), theme: THEME,
      definedNames: '<definedNames><definedName name="点数列">名簿!$E$2:$E$31</definedName><definedName name="マスタ範囲">マスタ!$A$2:$B$9</definedName></definedNames>',
    });
  }

  // ---- hc-30 LibreOffice-style hidden rows / columns / outline ("true" / "false" booleans), an AutoFilter with an active filter and the rows it hides ----------------------
  {
    const sst = new Sst();
    const hidden = new Set([4, 6, 9]);
    const d = { sst, cellT: true, rowAttr: (r) => ` customFormat="false" ht="12.8" hidden="${hidden.has(r) ? 'true' : 'false'}" customHeight="false" outlineLevel="${r >= 12 && r <= 14 ? 1 : 0}" collapsed="false"` };
    const rows = rosterRows({ h: 1, date: 2, int: 3, dec: 3, text: 5 }, { total: false, noFormula: true });
    const cols = '<cols><col collapsed="false" customWidth="true" hidden="false" outlineLevel="0" max="1" min="1" style="0" width="14"/><col collapsed="false" customWidth="true" hidden="true" outlineLevel="0" max="4" min="4" style="0" width="14"/><col collapsed="false" customWidth="true" hidden="false" outlineLevel="1" max="6" min="5" style="0" width="11"/><col collapsed="false" customWidth="false" hidden="false" outlineLevel="0" max="1025" min="7" style="0" width="11.53"/></cols>';
    const xml = XMLDECL + `<worksheet xmlns="${NS}" xmlns:r="${NSR}"><sheetPr filterMode="true"><outlinePr summaryBelow="false" summaryRight="false"/></sheetPr><dimension ref="A1:G31"/><sheetViews><sheetView showGridLines="true" tabSelected="true" workbookViewId="0"><pane xSplit="2" ySplit="1" topLeftCell="C2" activePane="bottomRight" state="frozen"/></sheetView></sheetViews><sheetFormatPr defaultColWidth="11.53" defaultRowHeight="12.8" outlineLevelRow="1" outlineLevelCol="1"/>${cols}<sheetData>${sheetData(d, rows)}</sheetData><autoFilter ref="A1:G31"><filterColumn colId="2"><filters><filter val="早稲田大学"/><filter val="東京大学"/></filters></filterColumn></autoFilter></worksheet>`;
    await pkg('hc-30-libreoffice-hidden-filter.xlsx', {
      sheets: [{ name: '名簿', xml }], styles: basicStyles(), sstXml: sst.xml(), theme: THEME,
      definedNames: '<definedNames><definedName function="false" hidden="true" localSheetId="0" name="_xlnm._FilterDatabase" vbProcedure="false">名簿!$A$1:$G$31</definedName></definedNames>',
    });
  }

  // ---- hc-31 a macro-enabled workbook (.xlsm: vbaProject part, macroEnabled content type) --------------------------------------------------------------------
  {
    const sst = new Sst();
    const rows = rosterRows({ h: 1, date: 2, int: 3, dec: 3, text: 5 }, { noFormula: true });
    const xml = XMLDECL + `<worksheet xmlns="${NS}" xmlns:r="${NSR}"><sheetPr codeName="Sheet1"/><sheetData>${sheetData({ sst }, rows)}</sheetData></worksheet>`;
    await pkg('hc-31-macro-enabled.xlsm', {
      sheets: [{ name: '名簿', xml }], styles: basicStyles(), sstXml: sst.xml(), theme: THEME,
      wbCt: 'application/vnd.ms-excel.sheet.macroEnabled.main+xml', wbPr: '<workbookPr codeName="ThisWorkbook"/>',
      extraParts: { 'xl/vbaProject.bin': new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, 0, 0, 0, 0]) },
      post: (files) => {
        files['xl/_rels/workbook.xml.rels'] = files['xl/_rels/workbook.xml.rels'].replace('</Relationships>', '<Relationship Id="rId90" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/></Relationships>');
        files['[Content_Types].xml'] = files['[Content_Types].xml'].replace('<Default Extension="xml"', '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/><Default Extension="xml"');
      },
    });
  }

  // ---- hc-32 sheet names that look like references / numbers / booleans, used in formulas of other sheets -----------------------------------------------------------
  {
    const sst = new Sst();
    const d = { sst };
    const rows = rosterRows({ h: 1, date: 2, int: 3, dec: 3, text: 5 }, { total: false, noFormula: true });
    const main = XMLDECL + `<worksheet xmlns="${NS}" xmlns:r="${NSR}"><sheetData>${sheetData(d, rows)}</sheetData></worksheet>`;
    const other = (n) => XMLDECL + `<worksheet xmlns="${NS}" xmlns:r="${NSR}"><sheetData>${sheetData(d, [[{ v: n }, { f: "'2024'!E2+A1!E2", v: 0 }, { f: "SUM('R1C1'!E2:E5)+SUM('TRUE'!E2:E3)", v: 0 }, { f: "COUNTA('名簿 (旧)'!A:A)", v: 0 }]])}</sheetData></worksheet>`;
    await pkg('hc-32-sheetnames-tricky.xlsx', {
      sheets: [{ name: '2024', xml: main }, { name: 'A1', xml: main }, { name: 'R1C1', xml: main }, { name: 'TRUE', xml: main }, { name: '名簿 (旧)', xml: main }, { name: '集計', xml: other('集計') }],
      styles: basicStyles(), sstXml: sst.xml(), theme: THEME,
    });
  }

  // ---- hc-33 odd formula texts (3-D refs, intersections, unions, spill refs, names with dots, whole rows, strings with quotes / A1 inside) ------------------------------
  {
    const sst = new Sst();
    const d = { sst };
    const rows = rosterRows({ h: 1, date: 2, int: 3, dec: 3, text: 5 }, { total: false, noFormula: true });
    const fs = [
      'SUM(名簿!A1:名簿!A3)', 'SUM(E2:E5 E3:E6)', 'SUM((E2,E4,E6))', 'SUM(2:2)', 'SUM($E:$F)', '"a""b" & "E2"', 'IF(E2>0,"E2","F2")', 'OFFSET(E2,1,1)', 'INDIRECT("E"&ROW())', 'SUM(E2:INDEX(E:E,5))',
      'SUMPRODUCT(E2:E5*{1;2;3;4})', '_xlfn.LET(_xlpm.x,E2,_xlpm.x+1)', 'my.name+1', 'SUM(名簿!$E$2:$E$31)/COUNT(名簿!$E$2:$E$31)', 'E2&CHAR(10)&F2', 'SUM( E2 : E5 )', 'IFERROR(VLOOKUP(A2,$A$2:$E$31,5,0),0)', 'TEXT(F2,"yyyy/m/d")', 'E2+#REF!',
    ];
    fs.forEach((f, i) => { rows[i + 1][8] = { f, v: 0, s: 3 }; });
    const xml = XMLDECL + `<worksheet xmlns="${NS}" xmlns:r="${NSR}"><sheetData>${sheetData(d, rows)}</sheetData></worksheet>`;
    await pkg('hc-33-odd-formulas.xlsx', {
      sheets: [{ name: '名簿', xml }], styles: basicStyles(), sstXml: sst.xml(), theme: THEME,
      definedNames: '<definedNames><definedName name="my.name">名簿!$E$2</definedName></definedNames>',
    });
  }

  // ---- hc-34 formulas WITHOUT cached values (openpyxl writes none) that cannot be kept live when imported: lookups / absolute refs / other sheets ---------------------------
  {
    const sst = new Sst();
    const d = { sst };
    const rows = rosterRows({ h: 1, date: 2, int: 3, dec: 3, text: 5 }, { total: false, noFormula: true });
    rows[0][7] = { v: '順位', s: 1 };
    rows[0][8] = { v: '比率', s: 1 };
    rows[0][9] = { v: '参照', s: 1 };
    for (let i = 1; i <= N; i++) {
      rows[i][7] = { f: `RANK(E${i + 1},$E$2:$E$31)` };
      rows[i][8] = { f: `E${i + 1}/SUM($E$2:$E$31)` };
      rows[i][9] = { f: `INDEX(マスタ!$A$2:$A$9,MOD(ROW(),8)+1)` };
    }
    const m = XMLDECL + `<worksheet xmlns="${NS}" xmlns:r="${NSR}"><sheetData>${sheetData(d, [[{ v: '大学' }], ...UNIV.map((u) => [{ v: u }])])}</sheetData></worksheet>`;
    const xml = XMLDECL + `<worksheet xmlns="${NS}" xmlns:r="${NSR}"><sheetData>${sheetData(d, rows)}</sheetData></worksheet>`;
    await pkg('hc-34-formulas-without-cache.xlsx', { sheets: [{ name: '名簿', xml }, { name: 'マスタ', xml: m }], styles: basicStyles(), sstXml: sst.xml(), theme: THEME });
  }
  // ---- hc-35 names (dynamic OFFSET range, external [1] ref, #REF!, sheet-scoped, relative) + conditional formats of many kinds + data validations, written by a LibreOffice-like tool -------
  {
    const sst = new Sst();
    const d = { sst };
    const rows = rosterRows({ h: 1, date: 2, int: 3, dec: 3, text: 5 }, { total: false, noFormula: true });
    const master = [[{ v: '大学' }], ...UNIV.map((u) => [{ v: u }])];
    const dxfs = '<dxfs count="2"><dxf><font><color rgb="FF9C0006"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf><dxf><fill><patternFill><bgColor rgb="FFC6EFCE"/></patternFill></fill></dxf></dxfs>';
    const styles = basicStyles().replace('<dxfs count="0"/>', dxfs);
    const cf = [
      '<conditionalFormatting sqref="A2:A31 C2:C31"><cfRule type="containsText" dxfId="0" priority="1" operator="containsText" text="学"><formula>NOT(ISERROR(SEARCH("学",A2)))</formula></cfRule></conditionalFormatting>',
      '<conditionalFormatting sqref="E2:E31"><cfRule type="top10" dxfId="1" priority="2" rank="3"/><cfRule type="dataBar" priority="3"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar></cfRule></conditionalFormatting>',
      '<conditionalFormatting sqref="B2:B31"><cfRule type="duplicateValues" dxfId="0" priority="4"/></conditionalFormatting>',
      '<conditionalFormatting sqref="A2:G31"><cfRule type="expression" dxfId="1" priority="5"><formula>AND($E2&gt;=90,$C2=$C$2)</formula></cfRule></conditionalFormatting>',
      '<conditionalFormatting sqref="F2:F31"><cfRule type="iconSet" priority="6"><iconSet iconSet="3Arrows"><cfvo type="percent" val="0"/><cfvo type="percent" val="33"/><cfvo type="percent" val="67"/></iconSet></cfRule></conditionalFormatting>',
    ].join('');
    const dv = '<dataValidations count="3"><dataValidation type="list" allowBlank="1" showErrorMessage="1" sqref="C2:C31"><formula1>マスタ!$A$2:$A$9</formula1></dataValidation><dataValidation type="whole" operator="between" allowBlank="1" showErrorMessage="1" sqref="E2:E31"><formula1>0</formula1><formula2>$E$33</formula2></dataValidation><dataValidation type="list" allowBlank="1" sqref="D2:D31"><formula1>"総合政策学部,環境情報学部,看護医療学部"</formula1></dataValidation></dataValidations>';
    rows[32] = [{ v: '上限' }, null, null, null, { n: 100, s: 3 }];
    const xml = XMLDECL + `<worksheet xmlns="${NS}" xmlns:r="${NSR}"><sheetData>${sheetData(d, rows)}</sheetData><autoFilter ref="A1:G31"/>${cf}${dv}</worksheet>`;
    const m = XMLDECL + `<worksheet xmlns="${NS}" xmlns:r="${NSR}"><sheetData>${sheetData(d, master)}</sheetData></worksheet>`;
    await pkg('hc-35-names-cf-dv.xlsx', {
      sheets: [{ name: '名簿', xml }, { name: 'マスタ', xml: m }], styles, sstXml: sst.xml(), theme: THEME,
      definedNames: '<definedNames><definedName name="氏名リスト">OFFSET(名簿!$B$2,0,0,COUNTA(名簿!$B$2:$B$31),1)</definedName><definedName name="壊れた名前">#REF!</definedName><definedName name="局所" localSheetId="0">名簿!$E$2:$E$31</definedName><definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">名簿!$A$1:$G$31</definedName><definedName name="_xlnm.Print_Titles" localSheetId="0">名簿!$A:$B,名簿!$1:$1</definedName><definedName name="_xlnm.Print_Area" localSheetId="0">名簿!$A$1:$H$33</definedName></definedNames>',
    });
  }
  // ---- hc-36 Office 2013+ theme with THEME-coloured header / cells (fgColor theme=4 tint): imported into a base with the 2007 theme the colours must not change -------------------
  {
    const sst = new Sst();
    const d = { sst };
    const rows = rosterRows({ h: 1, date: 2, int: 3, dec: 3, text: 5 }, { total: false, noFormula: true });
    const theme2013 = THEME.replace('4F81BD', '4472C4').replace('C0504D', 'ED7D31').replace('9BBB59', 'A5A5A5').replace('8064A2', 'FFC000').replace('4BACC6', '5B9BD5').replace('F79646', '70AD47').replace('<a:dk2><a:srgbClr val="1F497D"/>', '<a:dk2><a:srgbClr val="44546A"/>').replace('<a:lt2><a:srgbClr val="EEECE1"/>', '<a:lt2><a:srgbClr val="E7E6E6"/>');
    let styles = basicStyles();
    // header fill: accent1 (theme 4) lighter 60 %, font colour theme 1; number cells: accent2 (theme 5) darker 25 %
    styles = styles.replace('<fgColor rgb="FFDDEBF7"/><bgColor indexed="64"/>', '<fgColor theme="4" tint="0.59999389629810485"/><bgColor indexed="64"/>')
      .replace('</fills>', '<fill><patternFill patternType="solid"><fgColor theme="5" tint="-0.249977111117893"/><bgColor indexed="64"/></patternFill></fill></fills>').replace('<fills count="3">', '<fills count="4">')
      .replace('<xf numFmtId="3" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>', '<xf numFmtId="3" fontId="0" fillId="3" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1"/>');
    const xml = XMLDECL + `<worksheet xmlns="${NS}" xmlns:r="${NSR}"><sheetData>${sheetData(d, rows)}</sheetData></worksheet>`;
    await pkg('hc-36-theme2013-colours.xlsx', { sheets: [{ name: '名簿', xml }], styles, sstXml: sst.xml(), theme: theme2013 });
  }

  // ---- hc-37 external workbook links: formulas like [1]Sheet1!B2 while the book also has a sheet called Sheet1 (they must not follow column moves of the LOCAL sheet) -------------------
  {
    const sst = new Sst();
    const d = { sst };
    const rows = rosterRows({ h: 1, date: 2, int: 3, dec: 3, text: 5 }, { total: false, noFormula: true });
    for (let i = 1; i <= 10; i++) rows[i][7] = { f: `[1]Sheet1!E${i + 1}+Sheet1!E${i + 1}`, v: R[i - 1].score * 2, s: 3 };
    const main = XMLDECL + `<worksheet xmlns="${NS}" xmlns:r="${NSR}"><sheetData>${sheetData(d, rows)}</sheetData></worksheet>`;
    const ext = XMLDECL + `<externalLink xmlns="${NS}"><externalBook xmlns:r="${NSR}" r:id="rId1"><sheetNames><sheetName val="Sheet1"/></sheetNames><sheetDataSet><sheetData sheetId="0"><row r="2"><cell r="E2"><v>1</v></cell></row></sheetData></sheetDataSet></externalBook></externalLink>`;
    await pkg('hc-37-external-link.xlsx', {
      sheets: [{ name: 'Sheet1', xml: main }],
      styles: basicStyles(), sstXml: sst.xml(), theme: THEME,
      extraParts: { 'xl/externalLinks/externalLink1.xml': ext, 'xl/externalLinks/_rels/externalLink1.xml.rels': rels([['rId1', `${REL}/externalLinkPath`, 'file:///C:/old/other.xlsx', true]]) },
      post: (files) => {
        files['xl/_rels/workbook.xml.rels'] = files['xl/_rels/workbook.xml.rels'].replace('</Relationships>', `<Relationship Id="rId91" Type="${REL}/externalLink" Target="externalLinks/externalLink1.xml"/></Relationships>`);
        files['[Content_Types].xml'] = files['[Content_Types].xml'].replace('</Types>', '<Override PartName="/xl/externalLinks/externalLink1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/></Types>');
        files['xl/workbook.xml'] = files['xl/workbook.xml'].replace('</sheets>', '</sheets><externalReferences><externalReference r:id="rId91"/></externalReferences>');
      },
    });
  }
}

// ================================================================================================ library flavours
async function libraries() {
  // ---- SheetJS ----------------------------------------------------------------------------------------------------
  {
    const aoa = [HEAD, ...R.map((r) => [r.id, r.name, r.univ, r.fac, r.score, new Date(Date.UTC(2024, 3, 1 + r.day)), r.memo])];
    const ws = XLSX.utils.aoa_to_sheet(aoa, { cellDates: true });
    ws['!cols'] = [{ wch: 12 }, { wch: 14 }, { wch: 18 }];
    ws['!autofilter'] = { ref: `A1:G${N + 1}` };
    ws['!merges'] = [{ s: { r: 0, c: 8 }, e: { r: 0, c: 9 } }];
    ws['!rows'] = [{ hpt: 24 }, , , { hidden: true }];
    XLSX.utils.sheet_add_aoa(ws, [[{ t: 'n', f: `SUM(E2:E${N + 1})`, v: R.reduce((a, r) => a + r.score, 0) }]], { origin: `E${N + 3}` });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '名簿');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['大学'], ...UNIV.map((u) => [u])]), 'マスタ');
    wb.Workbook = { Sheets: [{ Hidden: 0 }, { Hidden: 1 }], Names: [{ Name: '点数', Ref: `名簿!$E$2:$E$${N + 1}` }] };
    fs.writeFileSync(path.join(OUT, 'js-01-sheetjs-basic.xlsx'), XLSX.write(wb, { bookType: 'xlsx', type: 'buffer' }));
    const wb2 = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb2, XLSX.utils.aoa_to_sheet(aoa, { cellDates: true }), '名簿');
    wb2.Workbook = { WBProps: { date1904: true } };
    fs.writeFileSync(path.join(OUT, 'js-02-sheetjs-date1904.xlsx'), XLSX.write(wb2, { bookType: 'xlsx', type: 'buffer', cellDates: true }));
    // numbers stored as text + header on row 3
    const wb3 = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb3, XLSX.utils.aoa_to_sheet([['名簿（エクスポート）'], [], HEAD, ...R.map((r) => [r.id, r.name, r.univ, r.fac, String(r.score), `2024/04/${1 + (r.day % 28)}`, r.memo])]), 'data');
    fs.writeFileSync(path.join(OUT, 'js-03-sheetjs-text-numbers.xlsx'), XLSX.write(wb3, { bookType: 'xlsx', type: 'buffer' }));
  }

  // ---- ExcelJS ----------------------------------------------------------------------------------------------------
  const ExcelJS = tryRequire('exceljs');
  if (ExcelJS) {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('名簿', { views: [{ state: 'frozen', xSplit: 1, ySplit: 3 }] });
    ws.mergeCells('A1:H1');
    ws.getCell('A1').value = '学生名簿（ExcelJS）';
    ws.getCell('A1').font = { bold: true, size: 14, name: 'Meiryo' };
    ws.getCell('A1').alignment = { horizontal: 'center' };
    ws.getRow(3).values = [...HEAD, '加点'];
    ws.getRow(3).font = { bold: true };
    ws.getRow(3).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFE699' } };
    R.forEach((r, k) => {
      const row = ws.getRow(4 + k);
      row.values = [r.id, r.name, r.univ, r.fac, r.score, new Date(Date.UTC(2024, 3, 1 + r.day)), r.memo || null, { formula: `E${4 + k}+5`, result: r.score + 5 }];
      row.getCell(6).numFmt = 'yyyy/mm/dd';
      row.getCell(5).numFmt = '#,##0';
    });
    ws.getCell(`E${N + 5}`).value = { formula: `SUM(E4:E${N + 3})`, result: R.reduce((a, r) => a + r.score, 0) };
    ws.autoFilter = `A3:I${N + 3}`;
    ws.columns.forEach((c, i) => { c.width = [12, 14, 18, 16, 8, 12, 10, 8][i] ?? 10; });
    ws.addConditionalFormatting({ ref: `E4:E${N + 3}`, rules: [{ type: 'cellIs', operator: 'lessThan', formulae: [60], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFFC7CE' } } } }] });
    ws.dataValidations.add(`C4:C${N + 3}`, { type: 'list', allowBlank: true, formulae: ['"慶應義塾大学,早稲田大学"'] });
    wb.addWorksheet('集計').getCell('A1').value = { formula: `AVERAGE('名簿'!E4:E${N + 3})`, result: 75 };
    await wb.xlsx.writeFile(path.join(OUT, 'js-04-exceljs-basic.xlsx'));

    // table
    const wb2 = new ExcelJS.Workbook();
    const ws2 = wb2.addWorksheet('名簿');
    ws2.addTable({
      name: 'tblRoster', ref: 'B3', headerRow: true, totalsRow: true, style: { theme: 'TableStyleMedium2', showRowStripes: true },
      columns: [{ name: '学籍番号', totalsRowLabel: '合計' }, { name: '氏名' }, { name: '大学' }, { name: '点数', totalsRowFunction: 'sum' }],
      rows: R.map((r) => [r.id, r.name, r.univ, r.score]),
    });
    ws2.getCell('B1').value = 'ExcelJS table';
    await wb2.xlsx.writeFile(path.join(OUT, 'js-05-exceljs-table.xlsx'));

    // streaming writer: inline strings (no shared strings), styles optional
    const swb = new ExcelJS.stream.xlsx.WorkbookWriter({ filename: path.join(OUT, 'js-06-exceljs-stream.xlsx'), useStyles: true, useSharedStrings: false });
    const sws = swb.addWorksheet('名簿');
    sws.columns = HEAD.map((h, i) => ({ header: h, key: 'c' + i, width: 14 }));
    R.forEach((r) => sws.addRow([r.id, r.name, r.univ, r.fac, r.score, new Date(Date.UTC(2024, 3, 1 + r.day)), r.memo]).commit());
    sws.commit();
    await swb.commit();

    // streaming with shared strings + rich text + a hyperlink
    const swb2 = new ExcelJS.stream.xlsx.WorkbookWriter({ filename: path.join(OUT, 'js-07-exceljs-stream-sst.xlsx'), useStyles: true, useSharedStrings: true });
    const sws2 = swb2.addWorksheet('名簿', { views: [{ state: 'frozen', ySplit: 1 }] });
    sws2.columns = HEAD.map((h, i) => ({ header: h, key: 'c' + i, width: 14 }));
    R.forEach((r, i) => {
      const row = sws2.addRow([r.id, r.name, { richText: [{ font: { bold: true, color: { argb: 'FFFF0000' } }, text: r.univ.slice(0, 2) }, { text: r.univ.slice(2) }] }, r.fac, r.score, new Date(Date.UTC(2024, 3, 1 + r.day)), i === 0 ? { text: 'リンク', hyperlink: 'https://example.com/?a=1&b=2' } : r.memo]);
      row.commit();
    });
    sws2.commit();
    await swb2.commit();
  }

  // ---- excel4node, write-excel-file, xlsx-js-style (three more node writers, each with its own XML habits) ---------------------------
  const E4N = tryRequire('excel4node');
  if (E4N) {
    const wb = new E4N.Workbook();
    const ws = wb.addWorksheet('名簿');
    const head = wb.createStyle({ font: { bold: true, color: 'FF0000' }, fill: { type: 'pattern', patternType: 'solid', fgColor: 'DDEBF7' } });
    const dfmt = wb.createStyle({ numberFormat: 'yyyy/mm/dd' });
    HEAD.forEach((h, j) => ws.cell(1, j + 1).string(h).style(head));
    R.forEach((r, i) => {
      ws.cell(i + 2, 1).string(r.id);
      ws.cell(i + 2, 2).string(r.name);
      ws.cell(i + 2, 3).string(r.univ);
      ws.cell(i + 2, 4).string(r.fac);
      ws.cell(i + 2, 5).number(r.score);
      ws.cell(i + 2, 6).date(new Date(Date.UTC(2024, 3, 1 + r.day))).style(dfmt);
      if (r.memo) ws.cell(i + 2, 7).string(r.memo);
      ws.cell(i + 2, 8).formula(`E${i + 2}*1.1`);
    });
    ws.column(1).setWidth(14);
    ws.row(1).freeze();
    await new Promise((res, rej) => wb.write(path.join(OUT, 'js-09-excel4node.xlsx'), (e) => (e ? rej(e) : res())));
  }
  const WEF = tryRequire('write-excel-file/node');
  if (WEF) {
    const writeXlsxFile = WEF.default ?? WEF;
    const rowsW = [HEAD.map((v) => ({ value: v, fontWeight: 'bold' })), ...R.map((r) => [{ value: r.id, type: String }, { value: r.name, type: String }, { value: r.univ, type: String }, { value: r.fac, type: String }, { value: r.score, type: Number }, { value: new Date(Date.UTC(2024, 3, 1 + r.day)), type: Date, format: 'yyyy/mm/dd' }, { value: r.memo || undefined, type: String }])];
    try {
      const made = await writeXlsxFile(rowsW, { sheet: '名簿' }); // v4: { toBuffer, toStream, toFile }
      fs.writeFileSync(path.join(OUT, 'js-10-write-excel-file.xlsx'), Buffer.from(await made.toBuffer()));
    } catch (e) { console.log('write-excel-file skipped:', e.message); }
  }
  const XJS = tryRequire('xlsx-js-style');
  if (XJS) {
    const aoa = [HEAD.map((h) => ({ v: h, t: 's', s: { font: { bold: true }, fill: { fgColor: { rgb: 'FFDDEBF7' } }, border: { bottom: { style: 'thin', color: { rgb: 'FF000000' } } } } })), ...R.map((r) => [r.id, r.name, r.univ, r.fac, r.score, { v: serial(r.day), t: 'n', z: 'yyyy/m/d' }, r.memo])];
    const wsx = XJS.utils.aoa_to_sheet(aoa);
    wsx['!cols'] = [{ wch: 12 }, { wch: 14 }];
    const wbx = XJS.utils.book_new();
    XJS.utils.book_append_sheet(wbx, wsx, '名簿');
    fs.writeFileSync(path.join(OUT, 'js-11-xlsx-js-style.xlsx'), XJS.write(wbx, { bookType: 'xlsx', type: 'buffer' }));
  }

  // ---- excel-export (a 2012 writer: x: prefixed worksheet, package/services/metadata/core-properties, xl/theme/theme.xml) (xlsx-writestream was tried too: its zip has sizes of 0 in the central directory and Excel itself refuses the file)
  const EE = tryRequire('excel-export');
  if (EE) {
    const conf = { cols: HEAD.map((h, i) => ({ caption: h, type: i === 4 ? 'number' : 'string' })), rows: R.map((r) => [r.id, r.name, r.univ, r.fac, r.score, `2024/04/${String(1 + (r.day % 28)).padStart(2, '0')}`, r.memo]) };
    fs.writeFileSync(path.join(OUT, 'js-12-excel-export.xlsx'), EE.execute(conf), 'binary');
  }
  // ---- officegen (templates of an old generator) and Go excelize through its WebAssembly port (tabSelected="true", <dimension></dimension>, no declaration standalone)
  const OG = tryRequire('officegen');
  if (OG) {
    const x = OG('xlsx');
    const sh = x.makeNewSheet();
    sh.name = '名簿';
    sh.data[0] = HEAD;
    R.forEach((r, i) => { sh.data[i + 1] = [r.id, r.name, r.univ, r.fac, r.score, `2024/04/${String(1 + (r.day % 28)).padStart(2, '0')}`, r.memo]; });
    await new Promise((res, rej) => { const out = fs.createWriteStream(path.join(OUT, 'js-13-officegen.xlsx')); out.on('close', res); out.on('error', rej); x.generate(out); });
  }
  const EXZ = tryRequire('excelize-wasm');
  if (EXZ) {
    const excelize = await EXZ.init(path.join(MODS, 'node_modules/excelize-wasm/excelize.wasm.gz'));
    const f = excelize.NewFile();
    f.SetSheetName('Sheet1', '名簿');
    const S = '名簿';
    HEAD.forEach((h, j) => f.SetCellValue(S, `${col(j + 1)}1`, h));
    const date = f.NewStyle({ CustomNumFmt: 'yyyy/m/d' });
    const bold = f.NewStyle({ Font: { Bold: true }, Fill: { Type: 'pattern', Color: ['#DDEBF7'], Pattern: 1 } });
    R.forEach((r, i) => {
      const n = i + 2;
      [r.id, r.name, r.univ, r.fac, r.score].forEach((v, j) => f.SetCellValue(S, `${col(j + 1)}${n}`, v));
      f.SetCellValue(S, `F${n}`, new Date(Date.UTC(2024, 3, 1 + r.day)));
      if (r.memo) f.SetCellValue(S, `G${n}`, r.memo);
      f.SetCellFormula(S, `H${n}`, `E${n}*1.1`);
    });
    if (date.style !== undefined) f.SetCellStyle(S, 'F2', `F${N + 1}`, date.style);
    if (bold.style !== undefined) f.SetCellStyle(S, 'A1', 'H1', bold.style);
    f.SetColWidth(S, 'A', 'B', 14);
    f.SetPanes(S, { Freeze: true, Split: false, XSplit: 0, YSplit: 1, TopLeftCell: 'A2', ActivePane: 'bottomLeft' });
    f.AutoFilter(S, `A1:H${N + 1}`, []);
    f.SetDefinedName({ Name: '点数', RefersTo: `${S}!$E$2:$E$${N + 1}`, Scope: 'Workbook' });
    const { buffer, error } = f.WriteToBuffer();
    if (error) console.log('excelize skipped:', error);
    else fs.writeFileSync(path.join(OUT, 'js-14-excelize-go.xlsx'), buffer);
  }

  // ---- xlsx-populate ----------------------------------------------------------------------------------------------
  const XP = tryRequire('xlsx-populate');
  if (XP) {
    const wb = await XP.fromBlankAsync();
    const sh = wb.sheet(0).name('名簿');
    sh.cell('A1').value([HEAD, ...R.map((r) => [r.id, r.name, r.univ, r.fac, r.score, new Date(Date.UTC(2024, 3, 1 + r.day)), r.memo || undefined])]);
    sh.range('A1:G1').style({ bold: true, fill: 'DDEBF7' });
    sh.column('F').style('numberFormat', 'yyyy/mm/dd').width(14);
    sh.column('A').width(14);
    sh.cell(`E${N + 3}`).formula(`SUM(E2:E${N + 1})`);
    sh.freezePanes(1, 1);
    sh.autoFilter(sh.range('A1:G31'));
    sh.range('I1:J1').merged(true);
    await wb.toFileAsync(path.join(OUT, 'js-08-xlsx-populate.xlsx'));
  }
}

await handmade();
await libraries();
console.log('written:', fs.readdirSync(OUT).filter((f) => f.endsWith('.xlsx')).length, 'xlsx files in', OUT);
