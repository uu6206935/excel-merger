/**
 * Builds the RAW round 3 golden fixture (hand-written package; Excel normalizes it into tests/fixtures/golden/r3-misc.xlsx).
 * Run: npx tsx tests/golden/tools/make-r3-fixture.ts <raw.xlsx>
 *
 *  Band  merged areas whose style sits on the top-left cell only (openpyxl style): a band A1:F1 with a thin box on A1, a
 *        three-row area A3:C5 (box on A3) with cells of its own (date / number format), a theme-filled band E7:G7, an
 *        area whose top-left cell has no look (A9:C9, B9 has its own fill), an area with a border of its own on C11
 *        (tests/doc/invariant.test.ts reads this sheet: keep its layout)
 *  Q     quote-prefix cells (B1:B3, E1) and text-format cells (C1, D1) for the CellWrite.quote cases
 *  Tb    table TblQ (A1:E4: ID, Q1, Q2, Q3, 合計) for the naming of inserted table columns
 *  Outline  outlines as Excel stores them (A1:F1 left / top / bottom on A1, A3:C5 left / top) and a theme-filled band
 *        with a medium red underline on its top-left cell (A7:C7)
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { buildXlsx } from '../build';

const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

export const R3_STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet ${NS}><numFmts count="1"><numFmt numFmtId="164" formatCode="0.0&quot;pt&quot;"/></numFmts>` +
  '<fonts count="2"><font><sz val="11"/><name val="游ゴシック"/></font><font><b/><sz val="11"/><name val="游ゴシック"/></font></fonts>' +
  '<fills count="5"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>' +
  '<fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/><bgColor indexed="64"/></patternFill></fill>' +
  '<fill><patternFill patternType="solid"><fgColor theme="4" tint="0.39997558519241921"/><bgColor indexed="64"/></patternFill></fill>' +
  '<fill><patternFill patternType="solid"><fgColor rgb="FFDDEBF7"/><bgColor indexed="64"/></patternFill></fill></fills>' +
  '<borders count="5"><border><left/><right/><top/><bottom/><diagonal/></border>' +
  '<border><left style="thin"><color indexed="64"/></left><right style="thin"><color indexed="64"/></right><top style="thin"><color indexed="64"/></top><bottom style="thin"><color indexed="64"/></bottom><diagonal/></border>' +
  '<border><left/><right/><top/><bottom style="medium"><color rgb="FFFF0000"/></bottom><diagonal/></border>' +
  '<border><left style="thin"><color indexed="64"/></left><right/><top style="thin"><color indexed="64"/></top><bottom style="thin"><color indexed="64"/></bottom><diagonal/></border>' + // 3 left / top / bottom
  '<border><left style="thin"><color indexed="64"/></left><right/><top style="thin"><color indexed="64"/></top><bottom/><diagonal/></border></borders>' + // 4 left / top
  '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
  '<cellXfs count="13">' +
  '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' + // 0
  '<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>' + // 1 bold, yellow, thin box
  '<xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>' + // 2 0.0"pt" + thin box
  '<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' + // 3 date
  '<xf numFmtId="0" fontId="1" fillId="3" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>' + // 4 theme fill, bold, centred
  '<xf numFmtId="0" fontId="0" fillId="0" borderId="2" xfId="0" applyBorder="1"/>' + // 5 medium red bottom
  '<xf numFmtId="0" fontId="0" fillId="4" borderId="0" xfId="0" applyFill="1"/>' + // 6 light blue fill
  '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" quotePrefix="1"/>' + // 7 quote prefix
  '<xf numFmtId="49" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' + // 8 text format
  '<xf numFmtId="49" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" quotePrefix="1"/>' + // 9 text format + quote prefix
  '<xf numFmtId="0" fontId="1" fillId="2" borderId="3" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>' + // 10 bold, yellow, left / top / bottom (an outline as Excel stores it)
  '<xf numFmtId="0" fontId="1" fillId="2" borderId="4" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>' + // 11 bold, yellow, left / top
  '<xf numFmtId="0" fontId="1" fillId="3" borderId="2" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>' + // 12 theme fill, medium red underline, centred
  '</cellXfs><cellStyles count="1"><cellStyle name="標準" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/>' +
  '<tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/></styleSheet>';

const SST = ['売上集計', '領域', 'テーマ帯', 'plain', 'ID', 'Q1', 'Q2', 'Q3', '合計', '00123', '=text', '00456', '007'];
const str = (ref: string, i: number, s = 0) => `<c r="${ref}"${s ? ` s="${s}"` : ''} t="s"><v>${i}</v></c>`;
const num = (ref: string, v: number, s = 0) => `<c r="${ref}"${s ? ` s="${s}"` : ''}><v>${v}</v></c>`;
const blank = (ref: string, s: number) => `<c r="${ref}" s="${s}"/>`;

const TABLE_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="TblQ" displayName="TblQ" ref="A1:E4" totalsRowShown="0"><autoFilter ref="A1:E4"/>` +
  '<tableColumns count="5"><tableColumn id="1" name="ID"/><tableColumn id="2" name="Q1"/><tableColumn id="3" name="Q2"/><tableColumn id="4" name="Q3"/><tableColumn id="5" name="合計"/></tableColumns>' +
  '<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>';

export async function buildR3Fixture(): Promise<Uint8Array> {
  return buildXlsx([
    {
      name: 'Band',
      data:
        `<row r="1">${str('A1', 0, 1)}</row>` +
        `<row r="3">${str('A3', 1, 1)}</row>` +
        `<row r="4">${num('B4', 46295, 3)}</row>` +
        `<row r="5">${num('C5', 12.5, 2)}</row>` +
        `<row r="7">${str('E7', 2, 4)}</row>` +
        `<row r="9">${str('A9', 3)}${blank('B9', 6)}</row>` +
        `<row r="11">${str('A11', 0, 1)}${blank('C11', 5)}</row>`,
      tail: '<mergeCells count="5"><mergeCell ref="A1:F1"/><mergeCell ref="A3:C5"/><mergeCell ref="E7:G7"/><mergeCell ref="A9:C9"/><mergeCell ref="A11:C11"/></mergeCells>',
    },
    {
      name: 'Q',
      data:
        `<row r="1">${str('B1', 9, 7)}${blank('C1', 8)}${blank('D1', 8)}${str('E1', 12, 9)}</row>` +
        `<row r="2">${str('B2', 10, 7)}</row>` +
        `<row r="3">${str('B3', 11, 7)}</row>` +
        `<row r="5">${blank('A5', 1)}</row>`,
    },
    {
      name: 'Tb',
      data:
        `<row r="1">${str('A1', 4)}${str('B1', 5)}${str('C1', 6)}${str('D1', 7)}${str('E1', 8)}</row>` +
        [2, 3, 4].map((r) => `<row r="${r}">${num('A' + r, r - 1)}${num('B' + r, 10 * r)}${num('C' + r, 20 * r)}${num('D' + r, 30 * r)}${num('E' + r, 60 * r)}</row>`).join(''),
      tail: '<tableParts count="1"><tablePart r:id="rId1"/></tableParts>',
      rels: [{ type: 'table', path: 'xl/tables/table1.xml', xml: TABLE_XML, contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml', id: 'rId1' }],
    },
    {
      // (a separate sheet: the layout of Band is used by tests/doc/invariant.test.ts as is)
      name: 'Outline',
      data:
        `<row r="1">${str('A1', 0, 10)}</row>` +
        `<row r="3">${str('A3', 1, 11)}</row>` +
        `<row r="7">${str('A7', 2, 12)}</row>`,
      tail: '<mergeCells count="3"><mergeCell ref="A1:F1"/><mergeCell ref="A3:C5"/><mergeCell ref="A7:C7"/></mergeCells>',
    },
  ], { sst: SST, styles: R3_STYLES });
}

// npx tsx tests/golden/tools/make-r3-fixture.ts <raw.xlsx>
// The hand-written package is only the RAW fixture: com/probe-r3.ps1 opens it in Excel and saves the normalized
// tests/fixtures/golden/r3-misc.xlsx (Excel's own XML: fonts with family / charset / scheme, dimensions ...), which is what
// the golden cases replay and what Excel starts from.
if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const out = path.resolve(process.argv[2] ?? 'r3-misc-raw.xlsx');
  fs.writeFileSync(out, await buildR3Fixture());
  console.log('wrote', out);
}
