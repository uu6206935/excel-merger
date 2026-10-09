/**
 * Test helpers for the reader tests (owner: reader R). Fixtures in this folder:
 *   torture.xlsx            R1 torture workbook (Excel-authored)          strict.xlsx      the same saved as Strict OOXML
 *   theme-probe.xlsx        torture + "Probe" sheet: theme colours x tints / indexed colours (excel-colors.json)
 *   apply-probe.xlsx        hand-written cellXfs with applyX="0" vs a parent style (gen/make-apply-probe.mjs)
 *   excel-probes.json       what Excel reports via COM (com/record-probes.ps1): Text, Value2, NumberFormat and
 *                           font / fill / alignment / borders for torture, apply-probe, interview/*, storage-test
 *   legacy.xls, date1904.xlsx, password.xlsx, password.xls, strict.xlsx  (com/make-fixtures.ps1)
 *   excel-fixtures.json     Excel's Text / Value2 for legacy.xls and date1904.xlsx
 *   legacy.ods, legacy.xlsb, legacy2003.xls (XML Spreadsheet 2003), legacy-excel.csv (Shift_JIS), legacy-utf8.csv,
 *   legacy-unicode.txt (UTF-16LE TSV)  legacy.xls re-saved by Excel (com/make-ods-xlsb.ps1)
 *   torture.xls, date1904.xls  Excel's BIFF8 saves of torture.xlsx / date1904.xlsx; xls-variants.json = what Excel shows
 *                           for them (com/make-xls-variants.ps1)
 *   corrupt/*.xls           byte-corrupted legacy.xls copies (fuzzing) on which plain SheetJS loops forever
 *   sjis.csv, utf8bom.csv, plain.tsv  (gen/make-csv.mjs); html-sjis.xls Shift_JIS HTML table (gen/make-html.mjs)
 *   generated/*.xlsx        packages the reader generated for the legacy inputs, opened by Excel without repair
 *                           (com/verify-open.ps1 → generated/verify-open.json)
 *   interview/*.xlsx        the reference app's 4 inputs; storage-test.xlsx / shared-test.xlsx (R2 oracle);
 *   sharedtest.xlsx (R1 shared formula variants); perf_10000x40.xlsx (R1 perf file)
 */
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';
import type { Cell, SheetGrid } from '../../../src/core/types';

export const FIXTURES = dirname(fileURLToPath(import.meta.url));

export function fx(name: string): Uint8Array {
  return new Uint8Array(readFileSync(join(FIXTURES, name)));
}

export function fxJson<T = unknown>(name: string): T {
  return JSON.parse(readFileSync(join(FIXTURES, name), 'utf8').replace(/^﻿/, '')) as T;
}

export function rc(addr: string): [number, number] {
  const m = /^\$?([A-Z]+)\$?(\d+)$/.exec(addr);
  if (!m) throw new Error('bad address ' + addr);
  let c = 0;
  for (const ch of m[1]) c = c * 26 + ch.charCodeAt(0) - 64;
  return [Number(m[2]), c];
}

export function cellAt(sheet: SheetGrid, addr: string): Cell | undefined {
  const [r, c] = rc(addr);
  return sheet.rows[r - 1]?.[c - 1];
}

/** Excel COM error integers (Value2 of an error cell) → error text. */
export function comError(v: number): string {
  const code = v + 2146826288 + 2000;
  const map: Record<number, string> = {
    2000: '#NULL!', 2007: '#DIV/0!', 2015: '#VALUE!', 2023: '#REF!', 2029: '#NAME?', 2036: '#NUM!', 2042: '#N/A', 2043: '#GETTING_DATA',
    2045: '#SPILL!', 2046: '#CONNECT!', 2047: '#BLOCKED!', 2048: '#UNKNOWN!', 2049: '#FIELD!', 2050: '#CALC!',
  };
  return map[code] ?? `#ERR${code}`;
}

export interface ProbeBorder {
  lineStyle: number;
  weight: number;
  color: string | null;
}

export interface ProbeCell {
  addr: string;
  text: string;
  vt: 'n' | 's' | 'b' | 'e' | 'null';
  v: number | string | boolean | null;
  nf: string;
  font?: { name: string; size: number; bold: boolean; italic: boolean; underline: number; strike: boolean; color: string | null };
  fill?: { pattern: number; color: string | null };
  hAlign?: number;
  vAlign?: number;
  wrap?: boolean;
  indent?: number;
  borders?: { left: ProbeBorder | null; top: ProbeBorder | null; bottom: ProbeBorder | null; right: ProbeBorder | null };
}

export interface ProbeFile {
  file: string;
  date1904: boolean;
  sheets: Array<{ name: string; usedRange: string; cells: ProbeCell[] }>;
}

let probeCache: { excelVersion: string; files: ProbeFile[] } | undefined;
export function probes(file: string): ProbeFile {
  const all = (probeCache ??= fxJson<{ excelVersion: string; files: ProbeFile[] }>('excel-probes.json'));
  const f = all.files.find((x) => x.file === file);
  if (!f) throw new Error('no probes for ' + file);
  return f;
}

/** COM (LineStyle, Weight) → OOXML border style. */
export function comBorderStyle(b: ProbeBorder): string {
  const w = b.weight;
  switch (b.lineStyle) {
    case 1:
      return w === 1 ? 'hair' : w === -4138 ? 'medium' : w === 4 ? 'thick' : 'thin';
    case -4115:
      return w === -4138 ? 'mediumDashed' : 'dashed';
    case -4118:
      return 'dotted';
    case -4119:
      return 'double';
    case 4:
      return w === -4138 ? 'mediumDashDot' : 'dashDot';
    case 5:
      return w === -4138 ? 'mediumDashDotDot' : 'dashDotDot';
    case 13:
      return 'slantDashDot';
    default:
      return `ls${b.lineStyle}/w${w}`;
  }
}

export const COM_HALIGN: Record<number, string | undefined> = {
  1: undefined, [-4131]: 'left', [-4108]: 'center', [-4152]: 'right', 5: 'fill', [-4130]: 'justify', 7: 'centerContinuous', [-4117]: 'distributed',
};
export const COM_VALIGN: Record<number, string | undefined> = { [-4160]: 'top', [-4108]: 'middle', [-4107]: undefined, [-4130]: 'top', [-4117]: 'middle' };

// ---------------------------------------------------------------------------------------------
// In-memory package builder for edge cases
// ---------------------------------------------------------------------------------------------

const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const RNS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

export interface PkgSheet {
  name: string;
  /** full worksheet XML; when it does not start with '<?xml' or '<', it is wrapped as <sheetData> content */
  xml: string;
  state?: 'hidden' | 'veryHidden';
  kind?: 'worksheet' | 'chartsheet';
  rels?: string;
  extraParts?: Record<string, string>;
}

export async function makePackage(o: {
  sheets: PkgSheet[];
  styles?: string | null;
  sst?: string | null;
  workbookExtra?: string;
  workbookPr?: string;
  extraParts?: Record<string, string>;
}): Promise<Uint8Array> {
  const zip = new JSZip();
  const overrides: string[] = [];
  const rels: string[] = [];
  const sheetEls: string[] = [];
  o.sheets.forEach((s, i) => {
    const kind = s.kind ?? 'worksheet';
    const folder = kind === 'worksheet' ? 'worksheets' : 'chartsheets';
    const path = `xl/${folder}/sheet${i + 1}.xml`;
    const xml = s.xml.trimStart().startsWith('<') && /^<\?xml|^<[A-Za-z:]*(worksheet|chartsheet)/.test(s.xml.trimStart())
      ? s.xml
      : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="${NS}" xmlns:r="${RNS}"><sheetData>${s.xml}</sheetData></worksheet>`;
    zip.file(path, xml);
    if (s.rels) zip.file(`xl/${folder}/_rels/sheet${i + 1}.xml.rels`, s.rels);
    for (const [p, x] of Object.entries(s.extraParts ?? {})) zip.file(p, x);
    overrides.push(`<Override PartName="/${path}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.${kind}+xml"/>`);
    rels.push(`<Relationship Id="rId${i + 1}" Type="${RNS}/${kind}" Target="${folder}/sheet${i + 1}.xml"/>`);
    sheetEls.push(`<sheet name="${s.name}" sheetId="${i + 1}"${s.state ? ` state="${s.state}"` : ''} r:id="rId${i + 1}"/>`);
  });
  const n = o.sheets.length;
  if (o.styles !== null) {
    zip.file('xl/styles.xml', o.styles ?? `<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="${NS}"><fonts count="1"><font><sz val="11"/><name val="游ゴシック"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="20" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs></styleSheet>`);
    rels.push(`<Relationship Id="rId${n + 1}" Type="${RNS}/styles" Target="styles.xml"/>`);
  }
  if (o.sst) {
    zip.file('xl/sharedStrings.xml', o.sst);
    rels.push(`<Relationship Id="rId${n + 2}" Type="${RNS}/sharedStrings" Target="sharedStrings.xml"/>`);
  }
  for (const [p, x] of Object.entries(o.extraParts ?? {})) zip.file(p, x);
  zip.file(
    'xl/workbook.xml',
    `<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="${NS}" xmlns:r="${RNS}">${o.workbookPr ?? '<workbookPr/>'}<sheets>${sheetEls.join('')}</sheets>${o.workbookExtra ?? ''}</workbook>`,
  );
  zip.file('xl/_rels/workbook.xml.rels', `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join('')}</Relationships>`);
  zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${RNS}/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>${overrides.join('')}</Types>`,
  );
  return zip.generateAsync({ type: 'uint8array' });
}
