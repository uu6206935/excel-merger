/**
 * Minimal .xlsx package builder for patcher unit tests (hand-written sheet XML, shared strings, styles).
 */
import JSZip from 'jszip';

export interface SheetSpec {
  name: string;
  /** inner xml of <sheetData> */
  data: string;
  /** xml between <sheetData> and </worksheet> (mergeCells, CF, DV ...) */
  tail?: string;
  /** xml before <sheetData> after <sheetViews> (e.g. <cols>) */
  head?: string;
  /** extra parts related to this sheet: rel type suffix -> [path, xml, contentType] */
  rels?: Array<{ type: string; path: string; xml: string; contentType: string; id: string }>;
}

const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

export const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet ${NS}><numFmts count="1"><numFmt numFmtId="164" formatCode="0.0&quot;pt&quot;"/></numFmts><fonts count="2"><font><sz val="11"/><name val="游ゴシック"/></font><font><b/><sz val="11"/><name val="游ゴシック"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color indexed="64"/></left><right style="thin"><color indexed="64"/></right><top style="thin"><color indexed="64"/></top><bottom style="thin"><color indexed="64"/></bottom><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="4"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/><xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/><xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs><cellStyles count="1"><cellStyle name="標準" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/></styleSheet>`;

export async function buildXlsx(sheets: SheetSpec[], opts: { sst?: string[]; names?: string; calcChain?: boolean; styles?: string } = {}): Promise<Uint8Array> {
  const zip = new JSZip();
  const overrides: string[] = [];
  const wbRels: string[] = [];
  sheets.forEach((s, i) => {
    const p = `xl/worksheets/sheet${i + 1}.xml`;
    zip.file(p, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet ${NS}><dimension ref="A1"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="18.75"/>${s.head ?? ''}<sheetData>${s.data}</sheetData>${s.tail ?? ''}</worksheet>`);
    overrides.push(`<Override PartName="/${p}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`);
    wbRels.push(`<Relationship Id="rId${i + 1}" Type="${REL}/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`);
    if (s.rels?.length) {
      zip.file(`xl/worksheets/_rels/sheet${i + 1}.xml.rels`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${s.rels.map((r) => `<Relationship Id="${r.id}" Type="${REL}/${r.type}" Target="../${r.path.replace(/^xl\//, '')}"/>`).join('')}</Relationships>`);
      for (const r of s.rels) {
        zip.file(r.path, r.xml);
        overrides.push(`<Override PartName="/${r.path}" ContentType="${r.contentType}"/>`);
      }
    }
  });
  const n = sheets.length;
  wbRels.push(`<Relationship Id="rId${n + 1}" Type="${REL}/styles" Target="styles.xml"/>`);
  zip.file('xl/styles.xml', opts.styles ?? STYLES);
  overrides.push('<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>');
  if (opts.sst) {
    wbRels.push(`<Relationship Id="rId${n + 2}" Type="${REL}/sharedStrings" Target="sharedStrings.xml"/>`);
    zip.file('xl/sharedStrings.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<sst ${NS} count="${opts.sst.length}" uniqueCount="${opts.sst.length}">${opts.sst.map((t) => `<si><t>${t}</t></si>`).join('')}</sst>`);
    overrides.push('<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>');
  }
  if (opts.calcChain) {
    wbRels.push(`<Relationship Id="rId${n + 3}" Type="${REL}/calcChain" Target="calcChain.xml"/>`);
    zip.file('xl/calcChain.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<calcChain ${NS}><c r="A1" i="1"/></calcChain>`);
    overrides.push('<Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>');
  }
  zip.file('xl/workbook.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook ${NS}><sheets>${sheets.map((s, i) => `<sheet name="${s.name}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('')}</sheets>${opts.names ? `<definedNames>${opts.names}</definedNames>` : ''}<calcPr calcId="191029"/></workbook>`);
  zip.file('xl/_rels/workbook.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${wbRels.join('')}</Relationships>`);
  zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
  zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>${overrides.join('')}</Types>`);
  return zip.generateAsync({ type: 'uint8array' });
}

/** read one part of a package as text */
export async function part(bytes: Uint8Array, path: string): Promise<string | null> {
  const zip = await JSZip.loadAsync(bytes);
  const f = zip.file(path);
  return f ? f.async('string') : null;
}

export async function partBytes(bytes: Uint8Array, path: string): Promise<Uint8Array | null> {
  const zip = await JSZip.loadAsync(bytes);
  const f = zip.file(path);
  return f ? f.async('uint8array') : null;
}

export async function listParts(bytes: Uint8Array): Promise<string[]> {
  const zip = await JSZip.loadAsync(bytes);
  return Object.keys(zip.files).filter((p) => !zip.files[p].dir).sort();
}
