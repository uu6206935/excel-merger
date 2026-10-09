/**
 * Robustness and edge cases: encrypted / corrupt / empty inputs (Japanese messages), 1904 date system (vs Excel),
 * prefixed namespaces, sparse far-away formatted cells, inline / rich / phonetic strings, t="d", shared formula
 * corner cases, data tables, chart sheets, missing parts.
 */
import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { readBook } from '../../src/core/xlsx/read';
import { cellText } from '../../src/core/xlsx/cellText';
import { COL_SLACK, ROW_SLACK, isoToSerial } from '../../src/core/xlsx/sheet';
import { MSG_CORRUPT, MSG_EMPTY, MSG_PASSWORD, ReadBookError } from '../../src/core/xlsx/zip';
import { decodeXmlText, encodeOoxmlEscapes, parseXml } from '../../src/core/xlsx/xml';
import type { SourceBook } from '../../src/core/types';
import { cellAt, fx, fxJson, makePackage, type ProbeCell } from '../fixtures/xlsx/helpers';

async function readErr(bytes: Uint8Array, name: string): Promise<ReadBookError> {
  try {
    await readBook(bytes, name);
  } catch (e) {
    expect(e).toBeInstanceOf(ReadBookError);
    return e as ReadBookError;
  }
  throw new Error('expected readBook to throw');
}

describe('bad input never crashes: Japanese error messages', () => {
  it('password-protected xlsx (CFB EncryptedPackage made by Excel)', async () => {
    const e = await readErr(fx('password.xlsx'), 'password.xlsx');
    expect(e.code).toBe('password');
    expect(e.message).toBe(MSG_PASSWORD);
    expect(e.message).toBe('パスワード付きのファイルは読み込めません。Excelでパスワードを解除して保存し直してください。');
  });

  it('truncated / corrupt zip', async () => {
    const good = fx('torture.xlsx');
    for (const bytes of [good.slice(0, 2000), good.slice(0, good.length - 100)]) {
      const e = await readErr(bytes, 'broken.xlsx');
      expect(e.message).toBe(MSG_CORRUPT);
      expect(e.message).toBe('Excelファイルとして読み込めませんでした（ファイルが壊れているか、Excel形式ではありません）');
    }
  });

  it('random binary, zip without a workbook, empty file', async () => {
    const rnd = new Uint8Array(4096);
    for (let i = 0; i < rnd.length; i++) rnd[i] = (i * 7919 + 13) % 256;
    rnd[5] = 0;
    expect((await readErr(rnd, 'x.xlsx')).message).toBe(MSG_CORRUPT);
    const zip = new JSZip();
    zip.file('hello.txt', 'hi');
    expect((await readErr(await zip.generateAsync({ type: 'uint8array' }), 'x.xlsx')).message).toBe(MSG_CORRUPT);
    const e = await readErr(new Uint8Array(0), 'x.xlsx');
    expect([e.code, e.message]).toEqual(['empty', MSG_EMPTY]);
  });

  it('workbook part that is not XML → corrupt; sheet part garbage → empty sheet, no crash', async () => {
    const zip = new JSZip();
    zip.file('xl/workbook.xml', 'not xml at all');
    expect((await readErr(await zip.generateAsync({ type: 'uint8array' }), 'x.xlsx')).message).toBe(MSG_CORRUPT);
    const pkg = await makePackage({ sheets: [{ name: 'S', xml: '<<<>>> garbage <c r="A1"><v>1' }] });
    const book = await readBook(pkg, 'g.xlsx');
    expect(book.sheets[0].name).toBe('S');
  });
});

describe('1904 date system (vs Excel)', () => {
  it('date cells are converted to 1900 serials; times and plain numbers keep their value', async () => {
    const book = await readBook(fx('date1904.xlsx'), 'date1904.xlsx');
    expect(book.date1904).toBe(true);
    const s = book.sheets[0];
    const excel = fxJson<{ date1904: ProbeCell[] }>('excel-fixtures.json').date1904;
    const byAddr = Object.fromEntries(excel.map((c) => [c.addr, c]));
    expect(byAddr.A1.v).toBe(44853); // what the 1904 file stores
    expect(cellAt(s, 'A1')).toMatchObject({ t: 'd', v: 44853 + 1462 }); // = 2026/10/20 in the 1900 system
    expect(cellText(book, cellAt(s, 'A1'))).toBe('2026/10/20');
    expect(cellText(book, cellAt(s, 'B1'))).toBe(byAddr.B1.text); // 2026年10月21日
    expect(cellAt(s, 'A2')).toMatchObject({ t: 'd', v: byAddr.A2.v }); // 12:30 stays 0.52…
    expect(cellText(book, cellAt(s, 'A2'))).toBe('12:30');
    expect(cellAt(s, 'A3')).toMatchObject({ t: 'n', v: 1234 });
    expect(cellAt(s, 'A4')).toMatchObject({ t: 's', v: '1900/1/1' });
  });
});

describe('package / sheet XML edge cases', () => {
  const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

  it('prefixed (x:) namespace documents', async () => {
    const pkg = await makePackage({
      sheets: [
        {
          name: 'P',
          xml: `<?xml version="1.0"?><x:worksheet xmlns:x="${NS}"><x:sheetData><x:row r="1"><x:c r="A1" t="s"><x:v>0</x:v></x:c><x:c r="B1" t="inlineStr"><x:is><x:t>インライン</x:t></x:is></x:c><x:c r="C1"><x:f>1+1</x:f><x:v>2</x:v></x:c></x:row></x:sheetData><x:mergeCells><x:mergeCell ref="A2:B2"/></x:mergeCells></x:worksheet>`,
        },
      ],
      sst: `<x:sst xmlns:x="${NS}"><x:si><x:t>共有</x:t></x:si></x:sst>`,
    });
    const s = (await readBook(pkg, 'p.xlsx')).sheets[0];
    expect(s.rows[0]).toEqual([{ t: 's', v: '共有' }, { t: 's', v: 'インライン' }, { t: 'n', v: 2, f: '1+1' }]);
    expect(s.features.merges).toEqual([{ r1: 2, c1: 1, r2: 2, c2: 2 }]);
  });

  it('a stray formatted cell at XFD1048576 does not allocate a huge grid', async () => {
    const pkg = await makePackage({
      sheets: [{ name: 'S', xml: '<row r="1"><c r="A1"><v>1</v></c><c r="B1" s="1"/></row><row r="3"><c r="C3" s="1"/></row><row r="1048576" hidden="1"><c r="XFD1048576" s="1"/></row>' }],
    });
    const s = (await readBook(pkg, 's.xlsx')).sheets[0];
    expect([s.maxRow, s.maxCol]).toEqual([3, 3]); // formatted cells near the data count, the far one does not
    expect(s.rows.length).toBeLessThanOrEqual(3);
    expect(s.rowProps.length).toBeLessThanOrEqual(1 + ROW_SLACK);
    expect(s.rows[2]![2]).toEqual({ t: 'z', v: null, s: 1 });
    expect(COL_SLACK).toBeGreaterThan(0);
  });

  it('a real value far away is kept (sparse, not dense)', async () => {
    const pkg = await makePackage({ sheets: [{ name: 'S', xml: '<row r="1"><c r="A1"><v>1</v></c></row><row r="200000"><c r="ZZ200000" t="s"><v>0</v></c></row>' }], sst: `<sst xmlns="${NS}"><si><t>far</t></si></sst>` });
    const s = (await readBook(pkg, 's.xlsx')).sheets[0];
    expect([s.maxRow, s.maxCol]).toEqual([200000, 702]);
    expect(s.rows[199999]![701]).toEqual({ t: 's', v: 'far' });
    expect(Object.keys(s.rows).length).toBe(2);
  });

  it('rows / cells without r attributes follow on from the previous one', async () => {
    const pkg = await makePackage({ sheets: [{ name: 'S', xml: '<row><c><v>1</v></c><c><v>2</v></c></row><row><c r="C2"><v>3</v></c><c><v>4</v></c></row>' }] });
    const s = (await readBook(pkg, 's.xlsx')).sheets[0];
    expect(s.rows[0]).toEqual([{ t: 'n', v: 1 }, { t: 'n', v: 2 }]);
    expect(s.rows[1]![2]).toEqual({ t: 'n', v: 3 });
    expect(s.rows[1]![3]).toEqual({ t: 'n', v: 4 });
  });

  it('value types: b, e, str, inlineStr (rich + phonetic), d (ISO), n with date style, missing <v>', async () => {
    const pkg = await makePackage({
      sheets: [
        {
          name: 'S',
          xml:
            '<row r="1"><c r="A1" t="b"><v>1</v></c><c r="B1" t="e"><v>#N/A</v></c><c r="C1" t="str"><f>"a"&amp;"b"</f><v>ab</v></c>' +
            '<c r="D1" t="inlineStr"><is><r><t>太字</t></r><r><rPr><b/><color rgb="FFFF0000"/></rPr><t xml:space="preserve"> 赤</t></r><rPh sb="0" eb="2"><t>フトジ</t></rPh></is></c>' +
            '<c r="E1" t="d"><v>2026-10-20T12:00:00Z</v></c><c r="F1" s="1"><v>46315</v></c><c r="G1" s="2"><v>0.5</v></c><c r="H1"><f>NOW()</f></c><c r="I1" t="s"/></row>',
        },
      ],
    });
    const book = await readBook(pkg, 't.xlsx');
    const r = book.sheets[0].rows[0]!;
    expect(r[0]).toEqual({ t: 'b', v: true });
    expect(r[1]).toEqual({ t: 'e', v: '#N/A' });
    expect(r[2]).toEqual({ t: 's', v: 'ab', f: '"a"&"b"' });
    expect(r[3]).toEqual({ t: 's', v: '太字 赤', rich: [{ text: '太字' }, { text: ' 赤', style: { bold: true, color: '#FF0000' } }] });
    expect(r[4]).toEqual({ t: 'd', v: 46315.5 });
    expect(r[5]).toEqual({ t: 'd', v: 46315, s: 1 });
    expect(cellText(book, r[5])).toBe('2026/10/20');
    expect(r[6]).toEqual({ t: 'd', v: 0.5, s: 2 });
    expect(cellText(book, r[6])).toBe('12:00');
    expect(r[7]).toEqual({ t: 'z', v: null, f: 'NOW()' });
    expect(r[8]).toBeUndefined();
  });

  it('shared strings: entities, _xHHHH_ escapes, CR LF normalisation, phonetic runs ignored, rich runs', async () => {
    const sst =
      `<sst xmlns="${NS}"><si><t>A &amp; B &lt;C&gt; &#x1F600; &#12354;</t></si><si><t>line1_x000D_&#10;line2</t></si><si><t>_x005F_x0041_</t></si>` +
      `<si><t>山田</t><rPh sb="0" eb="2"><t>ヤマダ</t></rPh><phoneticPr fontId="1"/></si><si/><si><t xml:space="preserve">  lead</t></si></sst>`;
    const pkg = await makePackage({ sheets: [{ name: 'S', xml: '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" t="s"><v>2</v></c><c r="D1" t="s"><v>3</v></c><c r="E1" t="s"><v>4</v></c><c r="F1" t="s"><v>5</v></c><c r="G1" t="s"><v>99</v></c></row>' }], sst });
    const book = await readBook(pkg, 's.xlsx');
    const r = book.sheets[0].rows[0]!;
    expect(r.slice(0, 6).map((c) => c?.v)).toEqual(['A & B <C> 😀 あ', 'line1\r\nline2', '_x0041_', '山田', '', '  lead']);
    expect(r[6]).toEqual({ t: 's', v: '' }); // index out of range
    expect(book.warnings.some((w) => w.includes('読み取れない値'))).toBe(true);
  });

  it('shared formulas: 2-D expansion, absolute parts kept, missing master → value only + warning', async () => {
    const pkg = await makePackage({
      sheets: [
        {
          name: 'S',
          xml:
            '<row r="2"><c r="B2"><f t="shared" ref="B2:C3" si="0">A2*$A$1+SUM(A$1:A2)</f><v>1</v></c><c r="C2"><f t="shared" si="0"/><v>2</v></c></row>' +
            '<row r="3"><c r="B3"><f t="shared" si="0"/><v>3</v></c><c r="C3"><f t="shared" si="0"/><v>4</v></c><c r="D3"><f t="shared" si="7"/><v>5</v></c></row>',
        },
      ],
    });
    const book = await readBook(pkg, 's.xlsx');
    const s = book.sheets[0];
    expect(cellAt(s, 'B2')!.f).toBe('A2*$A$1+SUM(A$1:A2)');
    expect(cellAt(s, 'C2')!.f).toBe('B2*$A$1+SUM(B$1:B2)');
    expect(cellAt(s, 'B3')!.f).toBe('A3*$A$1+SUM(A$1:A3)');
    expect(cellAt(s, 'C3')!.f).toBe('B3*$A$1+SUM(B$1:B3)');
    expect(cellAt(s, 'D3')).toEqual({ t: 'n', v: 5 });
    expect(book.warnings.some((w) => w.includes('共有数式'))).toBe(true);
  });

  it('data table (What-If) formulas keep values with a warning; CSE array without cm is not dynamic', async () => {
    const pkg = await makePackage({
      sheets: [{ name: 'S', xml: '<row r="1"><c r="A1"><f t="dataTable" ref="A1:B2" dt2D="1" dtr="1" r1="C1" r2="D1"/><v>7</v></c><c r="C1"><f t="array" ref="C1:C2">A1:A2*2</f><v>14</v></c></row>' }],
    });
    const book = await readBook(pkg, 's.xlsx');
    const s = book.sheets[0];
    expect(cellAt(s, 'A1')).toEqual({ t: 'n', v: 7 });
    expect(cellAt(s, 'C1')).toEqual({ t: 'n', v: 14, f: 'A1:A2*2', arr: 'C1:C2' });
    expect(s.features.arrays).toEqual([{ r1: 1, c1: 3, r2: 2, c2: 3 }]);
    expect(book.warnings.some((w) => w.includes('データテーブル'))).toBe(true);
  });

  it('bad style index, missing styles and sharedStrings parts, empty sheet, sheet without sheetData', async () => {
    const pkg = await makePackage({
      styles: null,
      sheets: [
        { name: 'A', xml: '<row r="1"><c r="A1" s="42"><v>1</v></c></row>' },
        { name: 'Empty', xml: '' },
        { name: 'NoData', xml: `<?xml version="1.0"?><worksheet xmlns="${NS}"><dimension ref="A1"/></worksheet>` },
      ],
    });
    const book = await readBook(pkg, 's.xlsx');
    expect(book.styles.length).toBe(1);
    expect(book.sheets[0].rows[0]).toEqual([{ t: 'n', v: 1 }]);
    expect(book.sheets.slice(1).map((s) => [s.name, s.maxRow, s.maxCol, s.rows.length])).toEqual([
      ['Empty', 0, 0, 0],
      ['NoData', 0, 0, 0],
    ]);
    expect(book.warnings.some((w) => w.includes('書式の番号'))).toBe(true);
  });

  it('chart sheets become empty grids (sheet order and localSheetId stay right)', async () => {
    const pkg = await makePackage({
      sheets: [
        { name: 'Chart1', kind: 'chartsheet', xml: `<?xml version="1.0"?><chartsheet xmlns="${NS}"/>` },
        { name: 'Data', xml: '<row r="1"><c r="A1"><v>1</v></c></row>' },
      ],
      workbookExtra: '<definedNames><definedName name="X" localSheetId="1">Data!$A$1</definedName><definedName name="Bad" localSheetId="9">1</definedName></definedNames>',
    });
    const book = await readBook(pkg, 'c.xlsx');
    expect(book.sheets.map((s) => [s.name, s.maxRow])).toEqual([
      ['Chart1', 0],
      ['Data', 1],
    ]);
    expect(book.names).toEqual([
      { name: 'X', scope: 'Data', refersTo: 'Data!$A$1' },
      { name: 'Bad', scope: null, refersTo: '1' },
    ]);
    expect(book.warnings.some((w) => w.includes('グラフシート'))).toBe(true);
  });

  it('frozen panes, hyperlinks (rels target), autofilter, cols spanning to XFD, custom row props', async () => {
    const pkg = await makePackage({
      sheets: [
        {
          name: 'S',
          xml:
            `<?xml version="1.0"?><worksheet xmlns="${NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
            '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" state="frozen"/></sheetView></sheetViews>' +
            '<sheetFormatPr defaultColWidth="10.5" defaultRowHeight="15"/><cols><col min="1" max="2" width="20" customWidth="1" style="1"/><col min="5" max="16384" hidden="1" width="0"/></cols>' +
            '<sheetData><row r="1" ht="30" customHeight="1" s="1" customFormat="1"><c r="A1"><v>1</v></c></row><row r="2" hidden="1" outlineLevel="2"><c r="A2"><v>2</v></c></row></sheetData>' +
            '<autoFilter ref="A1:B2"/><hyperlinks><hyperlink ref="A1" r:id="rId1"/><hyperlink ref="A2" location="S!B1"/></hyperlinks></worksheet>',
          rels: '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.jp/a?b=1&amp;c=2" TargetMode="External"/></Relationships>',
        },
      ],
    });
    const s = (await readBook(pkg, 'f.xlsx')).sheets[0];
    expect(s.features.freeze).toEqual({ cols: 0, rows: 1 });
    expect(s.features.autoFilter).toBe('A1:B2');
    expect(s.features.hyperlinks).toEqual([
      { ref: 'A1', target: 'https://example.jp/a?b=1&c=2' },
      { ref: 'A2', location: 'S!B1' },
    ]);
    expect(s.colProps[0]).toEqual({ width: 20, customWidth: true, s: 1 });
    expect(s.colProps[1]).toBe(s.colProps[0]);
    expect(s.colProps[2]).toBeUndefined();
    expect(s.colProps[16383]).toEqual({ width: 0, hidden: true });
    expect(s.rowProps[0]).toEqual({ height: 30, s: 1 });
    expect(s.rowProps[1]).toEqual({ hidden: true, outlineLevel: 2 });
    expect([s.defaultColWidth, s.defaultRowHeight]).toEqual([10.5, 15]);
  });

  it('x14 CF/DV inside mc:AlternateContent-free extLst, cfvo formulas, sparkline groups counted', async () => {
    const pkg = await makePackage({
      sheets: [
        {
          name: 'S',
          xml:
            `<?xml version="1.0"?><worksheet xmlns="${NS}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">` +
            '<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>' +
            '<conditionalFormatting sqref="A1:A9 C1:C9"><cfRule type="colorScale" priority="1"><colorScale><cfvo type="num" val="$B$1"/><cfvo type="max"/><color rgb="FF000000"/><color rgb="FFFFFFFF"/></colorScale></cfRule><cfRule type="expression" priority="2"><formula>AND(A1&gt;0,A1&lt;&gt;"")</formula></cfRule></conditionalFormatting>' +
            '<dataValidations count="1"><dataValidation type="list" sqref="B1:B9"><formula1>"はい,いいえ"</formula1></dataValidation></dataValidations>' +
            '<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}"><x14:conditionalFormattings><x14:conditionalFormatting><x14:cfRule type="dataBar" id="{1}"><x14:dataBar><x14:cfvo type="formula"><xm:f>$D$1</xm:f></x14:cfvo><x14:cfvo type="max"/></x14:dataBar></x14:cfRule><xm:sqref>D1:D9</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext>' +
            '<ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}"><x14:sparklineGroups><x14:sparklineGroup><x14:sparklines><x14:sparkline><xm:f>S!A1:A9</xm:f><xm:sqref>E1</xm:sqref></x14:sparkline></x14:sparklines></x14:sparklineGroup></x14:sparklineGroups></ext></extLst></worksheet>',
        },
      ],
    });
    const book = await readBook(pkg, 'x.xlsx');
    expect(book.sheets[0].features.rules).toEqual([
      { sqref: 'A1:A9 C1:C9', type: 'colorScale', formulas: ['$B$1'], kind: 'cf' },
      { sqref: 'A1:A9 C1:C9', type: 'expression', formulas: ['AND(A1>0,A1<>"")'], kind: 'cf' },
      { sqref: 'B1:B9', type: 'list', formulas: ['"はい,いいえ"'], kind: 'dv' },
      { sqref: 'D1:D9', type: 'dataBar', formulas: ['$D$1'], kind: 'cf', ext: true },
    ]);
    expect(book.features.sparklineGroups).toBe(1);
  });

  it('progress callback is called with Japanese messages; id is generated when missing', async () => {
    const msgs: string[] = [];
    const book: SourceBook = await readBook(fx('torture.xlsx'), 'torture.xlsx', { onProgress: (p) => msgs.push(p.message) });
    expect(book.id).toMatch(/^book-/);
    expect(msgs[0]).toBe('ファイルを展開中…');
    expect(msgs.some((m) => m.startsWith('シートを読み込み中'))).toBe(true);
    expect(msgs[msgs.length - 1]).toBe('読み込み完了');
  });
});

describe('xml helpers', () => {
  it('entities, _xHHHH_ and EOL normalisation; encode is the inverse of decode', () => {
    expect(decodeXmlText('a&amp;b&#10;c\r\nd\re')).toBe('a&b\nc\nd\ne');
    expect(decodeXmlText('_x000D__x0009_')).toBe('\r\t');
    const s = 'x\u0001y_x0041_z\r';
    expect(decodeXmlText(encodeOoxmlEscapes(s))).toBe(s);
    const doc = parseXml('<a x="1&#10;2" y="p\nq"><b>t<![CDATA[<raw>]]></b><!-- c --><c/></a>');
    const a = doc.children[0];
    expect(a.attrs).toEqual({ x: '1\n2', y: 'p q' });
    expect(a.children[0].text).toBe('t<raw>');
    expect(a.children.map((c) => c.local)).toEqual(['b', 'c']);
  });

  it('isoToSerial', () => {
    expect(isoToSerial('2026-10-20')).toBe(46315);
    expect(isoToSerial('1900-03-01')).toBe(61);
    expect(isoToSerial('1900-01-01')).toBe(1);
    expect(isoToSerial('12:00:00')).toBe(0.5);
    expect(isoToSerial('garbage')).toBeUndefined();
  });
});
