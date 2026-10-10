/**
 * A flexible hand-made .xlsx builder for the sheet-copier tests: whole worksheet parts, styles, theme, shared strings, defined names, the date
 * system, extra parts with relationships, workbook-level extras. (dropsheets-helpers' pkgOf has a fixed theme and no date-system switch.)
 */
import JSZip from 'jszip';

export const MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
export const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
export const WS_NS = `xmlns="${MAIN_NS}" xmlns:r="${R_NS}"`;
export const X14_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main';
export const XM_NS = 'http://schemas.microsoft.com/office/excel/2006/main';

export interface RelDef {
  id: string;
  type: string;
  target: string;
  external?: boolean;
}

export interface SheetSpec {
  name: string;
  state?: 'hidden' | 'veryHidden';
  /** inner xml of <sheetData> */
  data?: string;
  /** elements between <sheetFormatPr/> and <sheetData> (<cols>) */
  head?: string;
  /** the <sheetPr> element */
  pr?: string;
  /** the <sheetViews> inner xml (default one view) */
  views?: string;
  /** everything after </sheetData> */
  tail?: string;
  rels?: RelDef[];
  /** the whole worksheet part (overrides everything above) */
  xml?: string;
  /** extra root attributes / namespace declarations */
  rootAttrs?: string;
}

export interface PkgSpec {
  sheets: SheetSpec[];
  /** shared string items: plain texts, or { xml } for an item with runs */
  sst?: Array<string | { xml: string }>;
  styles?: string;
  theme?: string;
  names?: string;
  date1904?: boolean;
  wbExtra?: string;
  bookViews?: string;
  /** extra parts: path -> { xml, ct } */
  parts?: Array<{ path: string; xml: string; ct?: string }>;
  /** relationships of any part */
  partRels?: Record<string, RelDef[]>;
  wbRels?: RelDef[];
  metadata?: string;
  app?: boolean;
  /** file kind: xlsm */
  macro?: boolean;
}

const relType = (t: string) => (t.includes('://') ? t : `${R_NS}/${t}`);
export const relsXml = (rels: RelDef[]): string => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.map((r) => `<Relationship Id="${r.id}" Type="${relType(r.type)}" Target="${r.target.replace(/&/g, '&amp;')}"${r.external ? ' TargetMode="External"' : ''}/>`).join('')}</Relationships>`;
export const escX = (s: string): string => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

/** the default style sheet: font 0 游ゴシック 11, fills none / gray125, one border, xf 0 */
export const PLAIN_STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<styleSheet xmlns="${MAIN_NS}"><fonts count="1"><font><sz val="11"/><color theme="1"/><name val="游ゴシック"/><family val="3"/><charset val="128"/><scheme val="minor"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="標準" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/></styleSheet>`;

/** an Office theme part with the given accent colours (dk1 lt1 dk2 lt2 accent1..6 hlink folHlink) and font names */
export function themeXml(colors: string[], major = '游ゴシック Light', minor = '游ゴシック', name = 'Office Theme'): string {
  const [dk1, lt1, dk2, lt2, a1, a2, a3, a4, a5, a6, hl, fh] = colors;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="${name}"><a:themeElements><a:clrScheme name="${name}"><a:dk1><a:srgbClr val="${dk1}"/></a:dk1><a:lt1><a:srgbClr val="${lt1}"/></a:lt1><a:dk2><a:srgbClr val="${dk2}"/></a:dk2><a:lt2><a:srgbClr val="${lt2}"/></a:lt2><a:accent1><a:srgbClr val="${a1}"/></a:accent1><a:accent2><a:srgbClr val="${a2}"/></a:accent2><a:accent3><a:srgbClr val="${a3}"/></a:accent3><a:accent4><a:srgbClr val="${a4}"/></a:accent4><a:accent5><a:srgbClr val="${a5}"/></a:accent5><a:accent6><a:srgbClr val="${a6}"/></a:accent6><a:hlink><a:srgbClr val="${hl}"/></a:hlink><a:folHlink><a:srgbClr val="${fh}"/></a:folHlink></a:clrScheme><a:fontScheme name="${name}"><a:majorFont><a:latin typeface="${major}"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="${minor}"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="${name}"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>`;
}

export const OFFICE_COLORS = ['000000', 'FFFFFF', '44546A', 'E7E6E6', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '0563C1', '954F72'];
export const OTHER_COLORS = ['000000', 'FFFFFF', '1F497D', 'EEECE1', 'C0504D', '9BBB59', '8064A2', '4BACC6', 'F79646', '2C4D75', '0000FF', '800080'];

export async function buildPackage(spec: PkgSpec): Promise<Uint8Array> {
  const zip = new JSZip();
  const overrides: string[] = [];
  const wbRels: RelDef[] = [];
  const sheetEls: string[] = [];
  spec.sheets.forEach((s, i) => {
    const p = `xl/worksheets/sheet${i + 1}.xml`;
    const xml = s.xml ?? `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet ${WS_NS}${s.rootAttrs ? ' ' + s.rootAttrs : ''}>${s.pr ?? ''}<dimension ref="A1"/><sheetViews>${s.views ?? `<sheetView${i === 0 ? ' tabSelected="1"' : ''} workbookViewId="0"/>`}</sheetViews><sheetFormatPr defaultRowHeight="18.75"/>${s.head ?? ''}<sheetData>${s.data ?? ''}</sheetData>${s.tail ?? ''}</worksheet>`;
    zip.file(p, xml);
    overrides.push(`<Override PartName="/${p}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`);
    wbRels.push({ id: `rId${i + 1}`, type: 'worksheet', target: `worksheets/sheet${i + 1}.xml` });
    if (s.rels?.length) zip.file(`xl/worksheets/_rels/sheet${i + 1}.xml.rels`, relsXml(s.rels));
    sheetEls.push(`<sheet name="${escX(s.name).replace(/"/g, '&quot;')}" sheetId="${i + 1}"${s.state ? ` state="${s.state}"` : ''} r:id="rId${i + 1}"/>`);
  });
  const n = spec.sheets.length;
  let k = n;
  wbRels.push({ id: `rId${++k}`, type: 'styles', target: 'styles.xml' });
  zip.file('xl/styles.xml', spec.styles ?? PLAIN_STYLES);
  overrides.push('<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>');
  wbRels.push({ id: `rId${++k}`, type: 'theme', target: 'theme/theme1.xml' });
  zip.file('xl/theme/theme1.xml', spec.theme ?? themeXml(OFFICE_COLORS));
  overrides.push('<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>');
  if (spec.sst) {
    wbRels.push({ id: `rId${++k}`, type: 'sharedStrings', target: 'sharedStrings.xml' });
    zip.file('xl/sharedStrings.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<sst ${WS_NS} count="${spec.sst.length}" uniqueCount="${spec.sst.length}">${spec.sst.map((t) => (typeof t === 'string' ? `<si><t xml:space="preserve">${escX(t)}</t></si>` : `<si>${t.xml}</si>`)).join('')}</sst>`);
    overrides.push('<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>');
  }
  if (spec.metadata) {
    wbRels.push({ id: `rId${++k}`, type: 'sheetMetadata', target: 'metadata.xml' });
    zip.file('xl/metadata.xml', spec.metadata);
    overrides.push('<Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/>');
  }
  for (const r of spec.wbRels ?? []) wbRels.push(r);
  for (const p of spec.parts ?? []) {
    zip.file(p.path, p.xml);
    if (p.ct) overrides.push(`<Override PartName="/${p.path}" ContentType="${p.ct}"/>`);
  }
  for (const [owner, rels] of Object.entries(spec.partRels ?? {})) {
    const i = owner.lastIndexOf('/');
    zip.file(`${owner.slice(0, i + 1)}_rels/${owner.slice(i + 1)}.rels`, relsXml(rels));
  }
  zip.file('xl/workbook.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook ${WS_NS}><workbookPr${spec.date1904 ? ' date1904="1"' : ''}/>${spec.bookViews ?? '<bookViews><workbookView/></bookViews>'}<sheets>${sheetEls.join('')}</sheets>${spec.names ? `<definedNames>${spec.names}</definedNames>` : ''}<calcPr calcId="191029"/>${spec.wbExtra ?? ''}</workbook>`);
  zip.file('xl/_rels/workbook.xml.rels', relsXml(wbRels));
  const root: RelDef[] = [{ id: 'rId1', type: 'officeDocument', target: 'xl/workbook.xml' }];
  if (spec.app) {
    zip.file('docProps/app.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Microsoft Excel</Application><HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>ワークシート</vt:lpstr></vt:variant><vt:variant><vt:i4>${n}</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size="${n}" baseType="lpstr">${spec.sheets.map((s) => `<vt:lpstr>${escX(s.name)}</vt:lpstr>`).join('')}</vt:vector></TitlesOfParts></Properties>`);
    overrides.push('<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>');
    root.push({ id: 'rId2', type: 'extended-properties', target: 'docProps/app.xml' });
  }
  zip.file('_rels/.rels', relsXml(root));
  const mainCt = spec.macro ? 'application/vnd.ms-excel.sheet.macroEnabled.main+xml' : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml';
  zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/><Override PartName="/xl/workbook.xml" ContentType="${mainCt}"/>${overrides.join('')}</Types>`);
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

// ---------------------------------------------------------------------------------------------------------------------
// cell helpers
// ---------------------------------------------------------------------------------------------------------------------

export const cn = (n: number): string => {
  let s = '';
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
};
export const num = (ref: string, v: number, s = 0) => `<c r="${ref}"${s ? ` s="${s}"` : ''}><v>${v}</v></c>`;
export const sst = (ref: string, idx: number, s = 0) => `<c r="${ref}"${s ? ` s="${s}"` : ''} t="s"><v>${idx}</v></c>`;
export const inl = (ref: string, text: string, s = 0) => `<c r="${ref}"${s ? ` s="${s}"` : ''} t="inlineStr"><is><t xml:space="preserve">${escX(text)}</t></is></c>`;
export const boolc = (ref: string, v: boolean, s = 0) => `<c r="${ref}"${s ? ` s="${s}"` : ''} t="b"><v>${v ? 1 : 0}</v></c>`;
export const errc = (ref: string, e: string) => `<c r="${ref}" t="e"><v>${e}</v></c>`;
/** a formula cell with its cached result (type n number, s string, b boolean, e error, '' none) */
export function fml(ref: string, f: string, v: string | number | boolean | null = null, type: 'n' | 's' | 'b' | 'e' | '' = 'n', s = 0, fAttrs = '', cAttrs = ''): string {
  const t = type === 's' ? ' t="str"' : type === 'b' ? ' t="b"' : type === 'e' ? ' t="e"' : '';
  const body = v === null || type === '' ? '' : `<v>${type === 'b' ? (v ? 1 : 0) : escX(String(v))}</v>`;
  return `<c r="${ref}"${s ? ` s="${s}"` : ''}${t}${cAttrs}><f${fAttrs ? ' ' + fAttrs : ''}>${escX(f)}</f>${body}</c>`;
}
export const row = (r: number, cells: string, attrs = ''): string => `<row r="${r}"${attrs ? ' ' + attrs : ''}>${cells}</row>`;

/** a grid of values: numbers / strings (inline) / null → rows */
export function grid(rows: Array<Array<number | string | boolean | null | { f: string; v?: number | string }>>, startRow = 1): string {
  return rows.map((r, i) => row(startRow + i, r.map((v, j) => {
    const ref = cn(j + 1) + (startRow + i);
    if (v === null || v === undefined) return '';
    if (typeof v === 'number') return num(ref, v);
    if (typeof v === 'boolean') return boolc(ref, v);
    if (typeof v === 'string') return inl(ref, v);
    return fml(ref, v.f, v.v ?? null, typeof v.v === 'string' ? 's' : 'n');
  }).join(''))).join('');
}
