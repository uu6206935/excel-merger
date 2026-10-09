/**
 * ColProps.customWidth: the width was set by the USER = <col customWidth="1"> WITHOUT bestFit. An auto width — no
 * customWidth, or an auto-fitted column (bestFit="1", which Excel writes together with customWidth="1") — is what
 * Excel widens when a wider value is typed; the doc and the patcher both key on this flag. It is per column after
 * span expansion and stays undefined (never false) for auto widths.
 *
 * The other formats carry the same information and read the same way: BIFF8 ColInfo (fUserSet 0x2 without fBestFit
 * 0x4), XLSB BrtColInfo (same flags), ODS (style:use-optimal-column-width), XML 2003 (ss:AutoFitWidth="0").
 */
import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { readBook } from '../../src/core/xlsx/read';
import { prepareOds, scanBiff } from '../../src/core/xlsx/legacy';
import type { SheetGrid } from '../../src/core/types';
import { fx, makePackage } from '../fixtures/xlsx/helpers';

const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

async function colPropsOf(cols: string): Promise<SheetGrid['colProps']> {
  const pkg = await makePackage({
    sheets: [{ name: 'S', xml: `<?xml version="1.0"?><worksheet xmlns="${NS}"><cols>${cols}</cols><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>` }],
  });
  return (await readBook(pkg, 'cols.xlsx')).sheets[0].colProps;
}

describe('ColProps.customWidth', () => {
  it('torture.xlsx: set per column after span expansion; <col width="9" outlineLevel="1"/> has none', async () => {
    const book = await readBook(fx('torture.xlsx'), 'torture.xlsx');
    const d = book.sheets.find((s) => s.name === 'Data')!;
    // <col min="4" max="5" customWidth="1"> → both columns carry the flag
    for (const i of [0, 1, 2, 3, 4, 5, 6, 7, 8, 9]) expect(d.colProps[i]?.customWidth, `col ${i + 1}`).toBe(true);
    expect(d.colProps[10]).toEqual({ width: 0, hidden: true, customWidth: true });
    // <col min="12" max="13" width="9" outlineLevel="1"/>: auto width, the key is absent (not `false`)
    for (const i of [11, 12]) {
      expect(d.colProps[i]).toEqual({ width: 9, outlineLevel: 1 });
      expect(d.colProps[i]!.customWidth).toBeUndefined();
      expect('customWidth' in d.colProps[i]!).toBe(false);
    }
    // columns without a <col> element have no props at all
    expect(d.colProps[13]).toBeUndefined();
    // the other sheets: customWidth alone = user width; customWidth together with bestFit = an auto-fitted column
    // (auto width, the width itself is still reported); no <cols> at all
    const summary = book.sheets.find((s) => s.name === '集計 シート')!;
    expect(summary.colProps.map((p) => p?.customWidth)).toEqual([true, true]);
    const pivot = book.sheets.find((s) => s.name === 'Pivot')!;
    expect(pivot.colProps).toEqual([{ width: 11.25 }, { width: 11.875 }]); // bestFit="1" customWidth="1"
    expect(pivot.colProps.map((p) => p?.customWidth)).toEqual([undefined, undefined]);
    expect(pivot.colProps.every((p) => !!p && !('customWidth' in p))).toBe(true);
    expect(book.sheets.find((s) => s.name === 'Tbl')!.colProps).toEqual([]);
  });

  it('attribute spellings: only "1" / "true" set it; "0" / "false" / absent leave it undefined', async () => {
    const cp = await colPropsOf(
      '<col min="1" max="1" width="10" customWidth="1"/>' +
        '<col min="2" max="2" width="10" customWidth="true"/>' +
        '<col min="3" max="3" width="10" customWidth="0"/>' +
        '<col min="4" max="4" width="10" customWidth="false"/>' +
        '<col min="5" max="5" width="10"/>',
    );
    expect(cp.map((p) => p?.customWidth)).toEqual([true, true, undefined, undefined, undefined]);
    expect(cp.map((p) => p?.width)).toEqual([10, 10, 10, 10, 10]);
  });

  it('bestFit ends the user width: customWidth + bestFit ("1" / "true") is an auto-fitted column, read as auto width', async () => {
    const cp = await colPropsOf(
      '<col min="1" max="1" width="11.25" bestFit="1" customWidth="1"/>' + // what Excel writes after an auto-fit
        '<col min="2" max="2" width="11.25" customWidth="1" bestFit="1"/>' + // attribute order is irrelevant
        '<col min="3" max="3" width="11.25" bestFit="true" customWidth="true"/>' +
        '<col min="4" max="4" width="11.25" bestFit="1"/>' + // bestFit alone: auto
        '<col min="5" max="5" width="11.25" bestFit="0" customWidth="1"/>' + // bestFit off: the user width stands
        '<col min="6" max="6" width="11.25" bestFit="false" customWidth="true"/>' +
        '<col min="7" max="7" width="11.25" bestFit="0"/>' +
        '<col min="8" max="8" width="11.25" customWidth="1" bestFit=""/>', // empty = not true
    );
    expect(cp.map((p) => p?.customWidth)).toEqual([undefined, undefined, undefined, undefined, true, true, undefined, true]);
    // the width itself is kept for every column, and the flag key is absent (not false) for auto widths
    expect(cp.map((p) => p?.width)).toEqual([11.25, 11.25, 11.25, 11.25, 11.25, 11.25, 11.25, 11.25]);
    for (const i of [0, 1, 2, 3, 6]) expect('customWidth' in cp[i]!, `col ${i + 1}`).toBe(false);
  });

  it('the patcher\'s spellings: a non-xsd:boolean "on" is neither customWidth nor bestFit', async () => {
    const cp = await colPropsOf(
      '<col min="1" max="1" width="10" customWidth="on"/>' + // not a valid xsd:boolean: no user width
        '<col min="2" max="2" width="10" customWidth="1" bestFit="on"/>', // not bestFit either: still a user width
    );
    expect(cp.map((p) => p?.customWidth)).toEqual([undefined, true]);
  });

  it('a span is expanded to every column it covers, sharing one flag; neighbours are independent', async () => {
    const cp = await colPropsOf(
      '<col min="2" max="4" width="12" customWidth="1"/><col min="5" max="6" width="9"/><col min="7" max="7" width="15" customWidth="1" hidden="1"/>',
    );
    expect(cp[0]).toBeUndefined();
    expect([1, 2, 3].map((i) => cp[i]?.customWidth)).toEqual([true, true, true]);
    expect([4, 5].map((i) => cp[i]?.customWidth)).toEqual([undefined, undefined]);
    expect(cp[6]).toEqual({ width: 15, hidden: true, customWidth: true });
    expect(cp[7]).toBeUndefined();
  });

  it('an auto-fitted span is auto for every column it covers; a user span next to it keeps its flag', async () => {
    const cp = await colPropsOf(
      '<col min="1" max="3" width="14" bestFit="1" customWidth="1"/><col min="4" max="5" width="14" customWidth="1"/><col min="6" max="6" width="0" hidden="1" bestFit="1" customWidth="1"/>',
    );
    expect([0, 1, 2].map((i) => cp[i]?.customWidth)).toEqual([undefined, undefined, undefined]);
    expect([3, 4].map((i) => cp[i]?.customWidth)).toEqual([true, true]);
    expect(cp[5]).toEqual({ width: 0, hidden: true });
  });

  it('customWidth without a width attribute still reports the flag (width stays undefined); with bestFit it does not', async () => {
    const cp = await colPropsOf('<col min="1" max="1" customWidth="1" style="2"/><col min="2" max="2" customWidth="1" bestFit="1" style="2"/>');
    expect(cp[0]).toEqual({ customWidth: true, s: 2 });
    expect(cp[1]).toEqual({ s: 2 });
  });

  it('legacy .xls / .ods / .xlsb / XML 2003: a user width is customWidth, an auto-fitted column (BIFF 0x6 = userSet + bestFit) and untouched columns have none', async () => {
    // the same Excel workbook in four formats: column B is sized by the user, column C was auto-fitted
    for (const f of ['legacy.xls', 'legacy.ods', 'legacy.xlsb', 'legacy2003.xls']) {
      const book = await readBook(fx(f), f);
      const s = book.sheets[0];
      expect(s.colProps[0], f).toBeUndefined();
      expect(s.colProps[1]?.customWidth, `${f} col B`).toBe(true);
      expect(s.colProps[2]?.width, `${f} col C width`).toBeGreaterThan(11);
      expect(s.colProps[2]?.customWidth, `${f} col C`).toBeUndefined();
      expect('customWidth' in s.colProps[2]!, `${f} col C key`).toBe(false);
      expect(s.colProps[1]?.width, f).toBeGreaterThan(14);
      // the generated replacement package reads back the same flags
      const again = await readBook(book.bytes, 'generated.xlsx');
      expect(again.sheets[0].colProps, f).toEqual(s.colProps);
    }
  });

  it('date1904.xls / .xlsx: the auto-fitted date column reads as auto width in both twins', async () => {
    const xls = await readBook(fx('date1904.xls'), 'date1904.xls');
    const xlsx = await readBook(fx('date1904.xlsx'), 'date1904.xlsx');
    expect(xlsx.sheets[0].colProps[1]?.customWidth).toBeUndefined();
    expect(xls.sheets[0].colProps[1]?.customWidth).toBeUndefined();
    expect(xls.sheets[0].colProps[1]?.width).toBeCloseTo(xlsx.sheets[0].colProps[1]!.width!, 3);
  });

  it('torture.xls: BIFF ColInfo fUserSet → customWidth; fBestFit (0x6, the Pivot sheet) and an outline-only ColInfo are auto widths, like torture.xlsx', async () => {
    const xls = await readBook(fx('torture.xls'), 'torture.xls');
    const xlsx = await readBook(fx('torture.xlsx'), 'torture.xlsx');
    const dx = xls.sheets.find((s) => s.name === 'Data')!;
    const dt = xlsx.sheets.find((s) => s.name === 'Data')!;
    // Excel's BIFF8 save of torture.xlsx: COLINFO flags 0x0002 (fUserSet) for the set widths, 0x0100 (outline
    // level 1 only, no fUserSet) for columns L:M
    for (const i of [0, 1, 2, 3, 4, 5, 6, 7, 8, 9]) expect(dx.colProps[i]?.customWidth, `xls col ${i + 1}`).toBe(true);
    for (const i of [11, 12]) {
      expect(dx.colProps[i]?.width).toBe(9);
      expect(dx.colProps[i]?.customWidth, `xls col ${i + 1}`).toBeUndefined();
    }
    // the flag agrees with the xlsx twin column by column (the hidden width-0 column K is not carried by the
    // values-only conversion of .xls)
    for (let i = 0; i < 13; i++) {
      if (i === 10) continue;
      expect(dx.colProps[i]?.customWidth, `col ${i + 1}`).toBe(dt.colProps[i]?.customWidth);
    }
    // a sheet whose ColInfo records all carry fUserSet
    const sx = xls.sheets.find((s) => s.name === '集計 シート')!;
    expect(sx.colProps.map((p) => p?.customWidth)).toEqual([true, true]);
    // the Pivot sheet: COLINFO flags 0x0006 (fUserSet + fBestFit) = <col bestFit="1" customWidth="1"> in the xlsx twin
    const px = xls.sheets.find((s) => s.name === 'Pivot')!;
    const pt = xlsx.sheets.find((s) => s.name === 'Pivot')!;
    expect(px.colProps.map((p) => p?.width)).toEqual([11.25, 11.875]);
    expect(px.colProps.map((p) => p?.customWidth)).toEqual([undefined, undefined]);
    expect(px.colProps.map((p) => p?.customWidth)).toEqual(pt.colProps.map((p) => p?.customWidth));
  });
});

// ---------------------------------------------------------------------------------------------
// Synthetic inputs for the flag decoding of each legacy format
// ---------------------------------------------------------------------------------------------

function rec(type: number, payload: number[]): number[] {
  return [type & 0xff, type >> 8, payload.length & 0xff, payload.length >> 8, ...payload];
}
const le16 = (n: number) => [n & 0xff, (n >> 8) & 0xff];
const le32 = (n: number) => [n & 0xff, (n >> 8) & 0xff, (n >> 16) & 0xff, (n >>> 24) & 0xff];

/** BIFF8 workbook stream: one sheet whose substream holds COLINFO records [first, last, width in 1/256 chars, grbit]. */
function biffWithColInfo(cols: Array<[number, number, number, number]>): Uint8Array {
  const bof = (dt: number) => rec(0x0809, [...le16(0x0600), ...le16(dt), ...le16(0x0dbb), ...le16(0x07cc), ...le32(0), ...le32(0x06)]);
  const boundSheet = (pos: number) => rec(0x0085, [...le32(pos), 0, 0, 1, 0, 0x53]); // name "S", compressed
  const globalsLen = bof(0x0005).length + boundSheet(0).length + rec(0x000a, []).length;
  const sheet = [...bof(0x0010), ...cols.flatMap(([a, b, w, g]) => rec(0x007d, [...le16(a), ...le16(b), ...le16(w), ...le16(15), ...le16(g), ...le16(0)])), ...rec(0x000a, [])];
  return new Uint8Array([...bof(0x0005), ...boundSheet(globalsLen), ...rec(0x000a, []), ...sheet]);
}

describe('BIFF ColInfo flags (scanBiff)', () => {
  it('fUserSet (0x2) without fBestFit (0x4) is a user width; every other combination is an auto width', () => {
    const scan = scanBiff(
      biffWithColInfo([
        [0, 0, 2304, 0x0000], // A: no flags
        [1, 1, 2304, 0x0002], // B: fUserSet
        [2, 2, 2304, 0x0004], // C: fBestFit alone
        [3, 3, 2304, 0x0006], // D: fUserSet + fBestFit = an auto-fitted column (Excel's own save)
        [4, 4, 2304, 0x0100], // E: outline level only
        [5, 5, 2304, 0x0106], // F: outline + userSet + bestFit
        [6, 6, 2304, 0x0102], // G: outline + userSet
        [7, 7, 2304, 0x000a], // H: hidden-ish bit 3 (fPhonetic) + userSet
        [8, 9, 2304, 0x0006], // I:J span, auto-fitted
      ]),
    );
    expect(scan.sheets).toHaveLength(1);
    const s = scan.sheets[0];
    expect(s.name).toBe('S');
    expect(s.widths.slice(0, 10)).toEqual(new Array(10).fill(9));
    expect([...s.autoWidths].sort((a, b) => a - b)).toEqual([0, 2, 3, 4, 5, 8, 9]);
  });

  it('a later ColInfo over the same columns wins (user set after auto, and back)', () => {
    const scan = scanBiff(
      biffWithColInfo([
        [0, 2, 2304, 0x0006], // A:C auto-fitted …
        [1, 1, 3072, 0x0002], // … then B sized by the user
        [2, 2, 3072, 0x0006], // … C auto-fitted again
        [3, 3, 2304, 0x0002],
        [3, 3, 2304, 0x0006], // D: user, then auto-fitted
      ]),
    );
    const s = scan.sheets[0];
    expect([...s.autoWidths].sort((a, b) => a - b)).toEqual([0, 2, 3]);
    expect(s.widths.slice(0, 4)).toEqual([9, 12, 12, 9]);
  });

  it('truncated / garbage streams never throw', () => {
    const good = biffWithColInfo([[0, 0, 2304, 0x0006]]);
    for (let cut = 0; cut < good.length; cut++) expect(() => scanBiff(good.slice(0, cut)), `cut ${cut}`).not.toThrow();
  });
});

function xlsbRec(id: number, payload: number[]): number[] {
  const hdr = id < 0x80 ? [id] : [(id & 0x7f) | 0x80, id >> 7];
  const size: number[] = [];
  let n = payload.length;
  do {
    const b = n & 0x7f;
    n >>= 7;
    size.push(n ? b | 0x80 : b);
  } while (n);
  return [...hdr, ...size, ...payload];
}
const wide = (s: string) => [...le32(s.length), ...[...s].flatMap((c) => le16(c.charCodeAt(0)))];
const bundleSh = (relId: string, name: string) => xlsbRec(156, [...le32(0), ...le32(1), ...wide(relId), ...wide(name)]);
const colInfo = (first: number, last: number, grbit: number) => xlsbRec(60, [...le32(first), ...le32(last), ...le32(2304), ...le32(0), ...le16(grbit)]);

async function xlsbZip(sheets: Array<{ name: string; relId: string; path: string; bin: number[] }>, wbBin?: number[]): Promise<Uint8Array> {
  const z = new JSZip();
  z.file('xl/workbook.bin', new Uint8Array(wbBin ?? sheets.flatMap((s) => bundleSh(s.relId, s.name))));
  z.file(
    'xl/_rels/workbook.bin.rels',
    '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      sheets.map((s) => `<Relationship Id="${s.relId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="${s.path}"/>`).join('') +
      '</Relationships>',
  );
  for (const s of sheets) z.file('xl/' + s.path.replace(/^\/xl\//, ''), new Uint8Array(s.bin));
  return z.generateAsync({ type: 'uint8array' });
}

describe('XLSB BrtColInfo flags (prepareOds)', () => {
  it('per sheet NAME through workbook.bin.rels (not by part number); same rule as BIFF', async () => {
    const bytes = await xlsbZip([
      // "Data" is the FIRST sheet but lives in sheet2.bin; "売上" in sheet1.bin
      { name: 'Data', relId: 'rId1', path: 'worksheets/sheet2.bin', bin: [...colInfo(0, 0, 0x0002), ...colInfo(1, 1, 0x0006), ...colInfo(2, 3, 0x0104), ...xlsbRec(145, [])] },
      { name: '売上', relId: 'rId2', path: '/xl/worksheets/sheet1.bin', bin: [...colInfo(0, 1, 0x0006), ...colInfo(1, 1, 0x0002), ...xlsbRec(145, [])] },
      { name: 'Empty', relId: 'rId3', path: 'worksheets/sheet3.bin', bin: [...colInfo(0, 0, 0x0002)] },
    ]);
    const p = await prepareOds(bytes);
    expect(p.isOds).toBe(false);
    expect([...p.autoWidthsByName.keys()].sort()).toEqual(['Data', '売上'].sort());
    expect([...p.autoWidthsByName.get('Data')!].sort((a, b) => a - b)).toEqual([1, 2, 3]);
    expect([...p.autoWidthsByName.get('売上')!]).toEqual([0]); // A auto-fitted, B set by the user afterwards
  });

  it('column infos after BrtBeginSheetData are not read; a sheet with only user widths has no entry', async () => {
    const bytes = await xlsbZip([
      { name: 'A', relId: 'rId1', path: 'worksheets/sheet1.bin', bin: [...colInfo(0, 0, 0x0002), ...xlsbRec(145, []), ...colInfo(1, 1, 0x0006)] },
    ]);
    expect((await prepareOds(bytes)).autoWidthsByName.size).toBe(0);
  });

  it('broken workbook.bin / missing parts / truncated records never throw', async () => {
    const good = await xlsbZip([{ name: 'A', relId: 'rId1', path: 'worksheets/sheet1.bin', bin: [...colInfo(0, 0, 0x0006)] }]);
    expect((await prepareOds(good)).autoWidthsByName.get('A')).toEqual(new Set([0]));
    const garbage = await xlsbZip([{ name: 'A', relId: 'rId1', path: 'worksheets/sheet1.bin', bin: [0x3c, 0x12, 1, 2, 3] }], [0x9c, 0x01, 0xff, 0xff, 0xff, 0xff, 0x7f]);
    expect((await prepareOds(garbage)).autoWidthsByName.size).toBe(0);
    const missing = await xlsbZip([{ name: 'A', relId: 'rId9', path: 'worksheets/sheet1.bin', bin: [...colInfo(0, 0, 0x0006)] }], bundleSh('rId1', 'A'));
    expect((await prepareOds(missing)).autoWidthsByName.size).toBe(0);
    const noParts = new JSZip();
    noParts.file('xl/other.bin', new Uint8Array([1, 2, 3]));
    expect((await prepareOds(await noParts.generateAsync({ type: 'uint8array' }))).autoWidthsByName.size).toBe(0);
  });
});

describe('ODS use-optimal-column-width (prepareOds)', () => {
  const content =
    '<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0">' +
    '<office:automatic-styles>' +
    '<style:style style:name="co1" style:family="table-column"><style:table-column-properties fo:break-before="auto" style:column-width="1.905cm"/></style:style>' +
    '<style:style style:name="co2" style:family="table-column"><style:table-column-properties fo:break-before="auto" style:column-width="3.095625cm"/></style:style>' +
    '<style:style style:name="co3" style:family="table-column"><style:table-column-properties fo:break-before="auto" style:column-width="2.4077cm" style:use-optimal-column-width="true"/></style:style>' +
    '<style:style style:name="co4" style:family="table-column"><style:table-column-properties fo:break-before="auto" style:column-width="4cm" style:use-optimal-column-width="false"/></style:style>' +
    '</office:automatic-styles><office:body><office:spreadsheet>' +
    '<table:table table:name="T1"><table:table-column table:style-name="co1"/><table:table-column table:style-name="co2"/><table:table-column table:style-name="co3" table:number-columns-repeated="2"/><table:table-column table:style-name="co4"/><table:table-column table:style-name="co1" table:number-columns-repeated="16379"/>' +
    '<table:table-row><table:table-cell office:value-type="string"><text:p xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">x</text:p></table:table-cell></table:table-row></table:table>' +
    '<table:table table:name="T2"><table:table-column table:style-name="co1" table:number-columns-repeated="16384"/>' +
    '<table:table-row><table:table-cell office:value-type="string"/></table:table-row></table:table>' +
    '</office:spreadsheet></office:body></office:document-content>';

  it('optimal-width columns that differ from the default are auto; explicit widths are not; per table', async () => {
    const z = new JSZip();
    z.file('mimetype', 'application/vnd.oasis.opendocument.spreadsheet');
    z.file('content.xml', content);
    const p = await prepareOds(await z.generateAsync({ type: 'uint8array' }));
    expect(p.isOds).toBe(true);
    // B (co2) explicit, C:D (co3) optimal, E (co4 use-optimal-column-width="false") explicit; A and F.. = default co1
    expect(p.colWidths[0].flatMap((w, i) => (w === undefined ? [] : [i]))).toEqual([1, 2, 3, 4]);
    expect(p.autoWidths[0]).toEqual(new Set([2, 3]));
    expect(p.autoWidths[1]).toBeUndefined();
    expect(p.colWidths[1]).toEqual([]);
  });
});

describe('XML 2003 ss:AutoFitWidth', () => {
  const xml = (cols: string, second = '') =>
    '<?xml version="1.0"?>\n<?mso-application progid="Excel.Sheet"?>\n' +
    '<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">' +
    `<Worksheet ss:Name="Data"><Table>${cols}<Row><Cell><Data ss:Type="String">a</Data></Cell></Row></Table></Worksheet>${second}</Workbook>`;

  it('ss:AutoFitWidth="0" = set by the user; a <Column> without it (or with "1") is an auto width; spans follow their column', async () => {
    const book = await readBook(
      new TextEncoder().encode(
        xml(
          '<Column ss:AutoFitWidth="0" ss:Width="90"/>' +
            '<Column ss:Width="70"/>' +
            '<Column ss:Index="5" ss:AutoFitWidth="0" ss:Width="80" ss:Span="1"/>' +
            '<Column ss:Index="8" ss:AutoFitWidth="1" ss:Width="60"/>',
          '<Worksheet ss:Name="Two"><Table><Column ss:Width="50" ss:Span="1"/><Row><Cell><Data ss:Type="String">b</Data></Cell></Row></Table></Worksheet>',
        ),
      ),
      'x.xls',
    );
    expect(book.sheets.map((s) => s.name)).toEqual(['Data', 'Two']);
    expect(book.sheets[0].colProps.map((p) => p?.customWidth)).toEqual([true, undefined, undefined, undefined, true, true, undefined, undefined]);
    expect(book.sheets[0].colProps[1]?.width).toBeGreaterThan(10); // the auto width is still carried
    expect(book.sheets[1].colProps.map((p) => p?.customWidth)).toEqual([undefined, undefined]);
    expect(book.sheets[1].colProps.map((p) => p?.width !== undefined)).toEqual([true, true]);
  });
});
