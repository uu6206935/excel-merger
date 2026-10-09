/**
 * Round 3 fixtures and logs for the real-Excel parity test (tests/com/parity.test.ts).
 *
 * The workbook is built as a PACKAGE by hand (openpyxl style), not by Excel: Excel's own Range.Merge copies the top-left
 * format to every cell of the area, so "only the top-left cell carries the fill" (what openpyxl / many generators write, and
 * what LoggedOp 'unmerge' must repair) cannot be authored through Excel. Sheet 名簿 is compared with the JS engine (replayLog);
 * sheet 防御 is only for the runner's own safety net (plain strings Excel converts anyway) and is not part of the log.
 */
import JSZip from 'jszip';
import type { LoggedOp } from '../../src/core/types';

export const R3_SHEET = '名簿';
export const R3_DEF = '防御';

const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const RNS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** shared strings of the fixture (index = position) */
const SST = ['2026年度 社員名簿', '部署別メモ', 'ID', '氏名', '電話', '備考', '日付', '金額', '山田', '佐藤', '鈴木', '高橋', '営業', '開発', '総務', '人事', '区分'];
const si = (text: string) => SST.indexOf(text);

const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="${NS}">
<fonts count="3">
<font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/></font>
<font><b/><sz val="14"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/></font>
<font><b/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/></font>
</fonts>
<fills count="4">
<fill><patternFill patternType="none"/></fill>
<fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor theme="5" tint="0.39997558519241921"/><bgColor indexed="64"/></patternFill></fill>
</fills>
<borders count="3">
<border><left/><right/><top/><bottom/><diagonal/></border>
<border><left style="thin"><color auto="1"/></left><right style="thin"><color auto="1"/></right><top style="thin"><color auto="1"/></top><bottom style="thin"><color auto="1"/></bottom><diagonal/></border>
<border><left style="medium"><color rgb="FFC00000"/></left><right style="medium"><color rgb="FFC00000"/></right><top style="medium"><color rgb="FFC00000"/></top><bottom style="medium"><color rgb="FFC00000"/></bottom><diagonal/></border>
</borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="6">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
<xf numFmtId="0" fontId="0" fillId="3" borderId="2" xfId="0" applyFill="1" applyBorder="1"/>
<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="49" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>
</cellXfs>
<cellStyles count="1"><cellStyle name="標準" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`;
/** xf indexes of STYLES */
const XF = { title: 1, band: 2, date: 3, text: 4, head: 5 };

const s = (ref: string, xf: number, text: string) => `<c r="${ref}" s="${xf}" t="s"><v>${si(text)}</v></c>`;
const n = (ref: string, value: number, xf = 0) => `<c r="${ref}"${xf ? ` s="${xf}"` : ''}><v>${value}</v></c>`;
const blank = (ref: string, xf: number) => `<c r="${ref}" s="${xf}"/>`;

function sheet1(): string {
  const rows: string[] = [];
  // row 1: the merged title A1:F1 (fill + borders on the top-left cell ONLY) and the two-row band H1:J2 (theme fill)
  rows.push(`<row r="1">${s('A1', XF.title, SST[0])}${s('H1', XF.band, SST[1])}</row>`);
  rows.push(`<row r="3">${['ID', '氏名', '電話', '備考', '日付', '金額'].map((h, i) => s(`${'ABCDEF'[i]}3`, XF.head, h)).join('')}</row>`);
  const names = ['山田', '佐藤', '鈴木', '高橋'];
  for (let i = 0; i < 4; i++) {
    const r = 4 + i;
    // C4..C7: empty General cells (the writes go there); E: date format; C8 / C9: text format; F5: text format
    rows.push(`<row r="${r}">${n(`A${r}`, 100 + i)}${s(`B${r}`, 0, names[i])}${s(`D${r}`, 0, i % 2 ? '営業' : '開発')}${n(`E${r}`, 46000 + i * 31, XF.date)}${n(`F${r}`, 1000 * (i + 1))}</row>`);
  }
  rows.push(`<row r="8">${blank('C8', 0)}${n('E8', 46200, XF.date)}</row>`);
  rows.push(`<row r="9">${blank('C9', XF.text)}${blank('E9', XF.text)}</row>`);
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="${NS}" xmlns:r="${RNS}"><dimension ref="A1:J9"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><cols><col min="1" max="1" width="10" customWidth="1"/><col min="2" max="6" width="14" customWidth="1"/></cols><sheetData>${rows.join('')}</sheetData><mergeCells count="2"><mergeCell ref="A1:F1"/><mergeCell ref="H1:J2"/></mergeCells></worksheet>`;
}

/** 防御: column A holds cells of DIFFERENT number formats (General / date / text / General ...) so the runner must remember them per cell */
function sheet2(): string {
  const rows: string[] = [];
  const styles = [0, XF.date, XF.text, 0, XF.date, 0, 0, 0, 0, 0, 0, 0];
  for (let i = 0; i < styles.length; i++) rows.push(`<row r="${i + 1}">${blank(`A${i + 1}`, styles[i])}</row>`);
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="${NS}" xmlns:r="${RNS}"><dimension ref="A1:A${styles.length}"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData>${rows.join('')}</sheetData></worksheet>`;
}

export async function buildR3Fixture(): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>`,
  );
  zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
  zip.file(
    'xl/workbook.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="${NS}" xmlns:r="${RNS}"><bookViews><workbookView/></bookViews><sheets><sheet name="${R3_SHEET}" sheetId="1" r:id="rId1"/><sheet name="${R3_DEF}" sheetId="2" r:id="rId2"/></sheets><calcPr calcId="191029"/></workbook>`,
  );
  zip.file(
    'xl/_rels/workbook.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>`,
  );
  zip.file('xl/styles.xml', STYLES);
  zip.file('xl/sharedStrings.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="${NS}" count="${SST.length}" uniqueCount="${SST.length}">${SST.map((t) => `<si><t>${t}</t></si>`).join('')}</sst>`);
  zip.file('xl/worksheets/sheet1.xml', sheet1());
  zip.file('xl/worksheets/sheet2.xml', sheet2());
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

/** The user ops of the parity scenario (sheet 名簿). Positions: C = 電話 (col 3), E = 日付 (col 5). */
export function buildR3Log(): LoggedOp[] {
  return [
    // the merged title and the two-row band become plain cells; only their top-left cell carries fill + borders
    { kind: 'unmerge', sheet: R3_SHEET, refs: ['A1:F1'], centerAcross: true },
    { kind: 'unmerge', sheet: R3_SHEET, refs: ['H1:J2'], centerAcross: false },
    {
      kind: 'cells',
      sheet: R3_SHEET,
      reason: 'edit',
      writes: [
        { row: 4, col: 3, cell: { t: 's', v: '03-1234-5678' } }, // a phone number: Excel keeps it as text, NO quotePrefix
        { row: 5, col: 3, cell: { t: 's', v: '001', quote: true } },
        { row: 6, col: 3, cell: { t: 's', v: '1/2', quote: true } }, // would be a date
        { row: 7, col: 3, cell: { t: 's', v: '=text', quote: true } }, // would be a formula
        { row: 8, col: 3, cell: { t: 's', v: '007', numFmt: '@' } }, // text format, no quotePrefix
        { row: 9, col: 3, cell: { t: 's', v: '005', quote: true, numFmt: '@' } }, // both: quotePrefix exactly
        { row: 4, col: 5, cell: { t: 'n', v: 5, numFmt: 'General' } }, // date-formatted cell -> General
        { row: 9, col: 5, cell: { t: 's', v: 'abc', numFmt: 'General' } }, // text-formatted cell -> General
        { row: 6, col: 5, cell: { t: 'd', v: 46100, numFmt: 'yyyy/m/d' } }, // control
        { row: 4, col: 4, cell: { t: 's', v: 'メモ' } },
      ],
    },
  ];
}

/**
 * Raw ops for sheet 防御 (NOT in the log, so not compared with the JS engine): plain strings, i.e. without "quote", that Excel
 * converts when it parses them. The runner must notice it (read back), rewrite them as text WITH the apostrophe, restore the number
 * format the cell had (Excel changes a General cell to a date format for "1/2") and warn TEXT_COERCED.
 * Column A: rows 1..12, number formats General / date / text / General / date / General ...
 */
export const DEF_TEXTS = ['TRUE', '1/2', 'abc', '=(', '2026-01-02', '007', '=abc', '03-1234-5678', 'a\r\nb', "'x", ' 12', 'メモ'];
export function defensiveOps(): Array<Record<string, unknown>> {
  return [{ op: 'setCells', sheet: R3_DEF, at: 'A1', rows: DEF_TEXTS.map((t) => [t]) }];
}
