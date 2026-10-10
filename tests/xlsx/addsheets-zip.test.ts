/**
 * Output safety net and speed of the package writer (patch/wellformed.ts, patch/nativezip.ts, patch/pkg.ts generate):
 *  - xmlProblem keeps finding broken XML after its one-match fast path for ordinary <c> elements;
 *  - big parts are deflated by the native codec and handed to JSZip as finished entries: the package reads back byte for byte (CRC checked),
 *    the entries are real DEFLATE entries, the order of the parts is the order of the writes, small parts take the usual route;
 *  - the JavaScript CRC-32 (the browser fallback) equals Node's.
 */
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { Engine } from '../../src/core/xlsx/patch/engine';
import { compressParts, deflateTemplate, jsCrc32, nativeCodec, nativeInflateEntry } from '../../src/core/xlsx/patch/nativezip';
import { Pkg } from '../../src/core/xlsx/patch/pkg';
import { xmlProblem } from '../../src/core/xlsx/patch/wellformed';
import { buildPackage, cn, num, row, sst } from './addsheets-build';

const ws = (body: string, attrs = '') => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"${attrs}><sheetData>${body}</sheetData></worksheet>`;

describe('xmlProblem: the fast path for ordinary cells never hides a problem', () => {
  it('ordinary sheets are fine (cells of every usual shape, entity references, prefixed attributes, single quotes)', () => {
    const body = row(1, `<c r="A1" s="1" t="s"><v>5</v></c><c r="B1"/><c r="C1" t="b"><v>1</v></c><c r="D1"><f>SUM(A2:B2)&amp;"x"&lt;</f><v>3</v></c><c r="E1" t="str"><f t="shared" ref="E1:E4" si="0">A1</f><v>a &amp; b</v></c><c r="F1" t="inlineStr"><is><t xml:space="preserve"> x </t></is></c><c r="G1" x14ac:dyDescent="1"><v>1</v></c><c r='H1'><v>2</v></c>`)
      + `<row r="2"><c r="A2" cm="1"><f t="array" ref="A2">SEQUENCE(2)</f><v>1</v></c><c r="B2" t="e"><v>#N/A</v></c></row>`;
    expect(xmlProblem(ws(body, ' xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"'))).toBeNull();
  });

  it.each([
    ['an unclosed cell', ws('<row r="1"><c r="A1"><v>1</v></row>')],
    ['an unclosed value', ws('<row r="1"><c r="A1"><v>1</c></row>')],
    ['a mismatched end tag', ws('<row r="1"><c r="A1"><f>A2</v></c></row>')],
    ['a bare ampersand in a value', ws('<row r="1"><c r="A1" t="str"><v>a & b</v></c></row>')],
    ['a bare ampersand in a formula', ws('<row r="1"><c r="A1"><f>A2&B2</f><v>1</v></c></row>')],
    ['an unknown entity', ws('<row r="1"><c r="A1" t="str"><v>a &foo; b</v></c></row>')],
    ['a control character', ws(`<row r="1"><c r="A1" t="str"><v>a${String.fromCharCode(3)}b</v></c></row>`)],
    ['an attribute with an ampersand', ws('<row r="1"><c r="A1" s="a&b"><v>1</v></c></row>')],
    ['a prefix that is not declared', ws('<row r="1"><c r="A1" foo:bar="1"><v>1</v></c></row>')],
    ['a text outside the document', ws('<row r="1"><c r="A1"><v>1</v></c></row>') + 'stray'],
    ['a second root', ws('') + '<worksheet/>'],
    ['a cell without its end', ws('<row r="1"><c r="A1"><v>1</v>')],
  ])('%s is reported', (_name, xml) => {
    expect(xmlProblem(xml)).not.toBeNull();
  });
});

describe('natively compressed parts', () => {
  const big = (n: number) => ws(Array.from({ length: n }, (_x, r) => row(r + 1, Array.from({ length: 30 }, (_y, c) => num(`${cn(c + 1)}${r + 1}`, (r * 31 + c) / 7, (c % 5) + 1)).join(''))).join(''));

  it('a package with parts above and below the threshold reads back exactly (CRC checked), small parts included', async () => {
    const bytes = await buildPackage({ sheets: [{ name: 'S', data: row(1, num('A1', 1)) }], sst: ['a'] });
    const eng = await Engine.open(bytes);
    const text = big(6000); // ~ 2 MB
    expect(text.length).toBeGreaterThan(1 << 20);
    eng.pkg.write('xl/worksheets/sheet9.xml', text);
    eng.pkg.write('xl/worksheets/sheet10.xml', big(10)); // small: the usual route
    const out = await eng.pkg.generate();
    const zip = await JSZip.loadAsync(out, { checkCRC32: true });
    expect(await zip.file('xl/worksheets/sheet9.xml')!.async('string')).toBe(text);
    expect(await zip.file('xl/worksheets/sheet10.xml')!.async('string')).toBe(big(10));
    const info = (zip.file('xl/worksheets/sheet9.xml') as unknown as { _data: { compression: { magic: string }; compressedSize: number; uncompressedSize: number } })._data;
    expect(info.compression.magic).toBe('\x08\x00'); // DEFLATE
    expect(info.compressedSize).toBeLessThan(info.uncompressedSize / 3);
    expect(info.uncompressedSize).toBe(new TextEncoder().encode(text).length);
    // the untouched parts are untouched
    expect(await zip.file('xl/sharedStrings.xml')!.async('string')).toContain('<si>');
  });

  it('compressParts: entries come back in the order of the parts, the JavaScript CRC-32 equals the native one, and without a deflated template nothing is done', async () => {
    const bytes = await buildPackage({ sheets: [{ name: 'S', data: sst('A1', 0) }], sst: ['a'] });
    const zip = await JSZip.loadAsync(bytes);
    expect(deflateTemplate(zip)).not.toBeNull();
    const parts = [{ path: 'p/a.xml', text: big(40) }, { path: 'p/b.xml', text: 'b'.repeat(100) }, { path: 'p/c.xml', text: big(5) }];
    const ready = await compressParts(zip, parts);
    expect(ready!.map((r) => r.path)).toEqual(['p/a.xml', 'p/b.xml', 'p/c.xml']);
    expect(await compressParts(new JSZip().file('x.txt', 'plain'), parts)).toBeNull();
    const c = nativeCodec();
    expect(c).not.toBeNull();
    const data = new TextEncoder().encode(big(30));
    expect(jsCrc32(data)).toBe(c!.crc32(data));
    expect(jsCrc32(new Uint8Array(0))).toBe(0);
    expect(jsCrc32(new TextEncoder().encode('123456789'))).toBe(0xcbf43926);
  });
});

describe('natively inflated parts', () => {
  const big = (n: number) => ws(Array.from({ length: n }, (_x, r) => row(r + 1, Array.from({ length: 30 }, (_y, c) => num(`${cn(c + 1)}${r + 1}`, (r * 31 + c) / 7, (c % 5) + 1)).join(''))).join(''));
  const packageWith = async (text: string) => {
    const zip = new JSZip();
    zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>');
    zip.file('xl/workbook.xml', '<workbook/>');
    zip.file('xl/worksheets/sheet1.xml', text);
    zip.file('xl/worksheets/sheet2.xml', ws(row(1, num('A1', 1))));
    return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  };

  it('a big part is inflated by zlib and reads back exactly; a small one is left to JSZip', async () => {
    const text = big(3000); // ~ 1 MB
    expect(text.length).toBeGreaterThan(1 << 18);
    const pkg = await Pkg.load(await packageWith(text));
    const f = pkg.zip.file('xl/worksheets/sheet1.xml')!;
    const raw = await nativeInflateEntry(f);
    expect(raw).not.toBeNull();
    expect(new TextDecoder().decode(raw!)).toBe(text);
    expect(await nativeInflateEntry(pkg.zip.file('xl/worksheets/sheet2.xml')!)).toBeNull();
    expect(await pkg.read('xl/worksheets/sheet1.xml')).toBe(text);
    expect(await pkg.read('xl/worksheets/sheet2.xml')).toContain('<c r="A1">');
  });

  it('a size limit below the declared size, a lying directory entry and damaged data are never inflated natively (the reader reports them as before)', async () => {
    const text = big(3000);
    const pkg = await Pkg.load(await packageWith(text));
    const f = pkg.zip.file('xl/worksheets/sheet1.xml')!;
    expect(await nativeInflateEntry(f, 1000)).toBeNull();
    const d = (f as unknown as { _data: { uncompressedSize: number; compressedContent: Uint8Array } })._data;
    const real = d.uncompressedSize;
    d.uncompressedSize = real - 10; // declares less than the stream holds: zlib stops at the cap
    expect(await nativeInflateEntry(f)).toBeNull();
    d.uncompressedSize = real + 10; // declares more than the stream holds
    expect(await nativeInflateEntry(f)).toBeNull();
    d.uncompressedSize = real;
    const keep = d.compressedContent;
    d.compressedContent = keep.slice(0, keep.length >> 1); // truncated deflate data
    expect(await nativeInflateEntry(f)).toBeNull();
    d.compressedContent = keep;
    expect(await nativeInflateEntry(f)).not.toBeNull();
  });
});
