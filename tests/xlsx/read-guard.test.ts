/**
 * Reader round 2: FeatureRule.type (CF / DV incl. x14), SourceBook.encoding for text inputs, SourceFormat 'xlsb',
 * and the robustness guards — zip bombs (declared total > 1.5 GB, overlapping entries, ZIP64 sizes, entries that
 * inflate beyond their declared size), the shared-strings / sheet part size checks and the text size check.
 *
 * Bomb fixtures are built in memory: a deflate stream of N MiB of zeros is N copies of one full-flushed 1 MiB
 * segment (~1 KB each, byte aligned, no back references across segments) plus an empty final block, so a
 * "1.6 GB" part costs ~1.6 MB here and the guards are checked without anything big being inflated.
 */
import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { constants, crc32 as nodeCrc32, deflateRawSync, inflateRawSync } from 'node:zlib';
import Encoding from 'encoding-japanese';
import { readBook } from '../../src/core/xlsx/read';
import { decodeText, restoreZip } from '../../src/core/xlsx/legacy';
import {
  MAX_PART_BYTES,
  MAX_UNCOMPRESSED_TOTAL,
  MSG_CORRUPT,
  Package,
  ReadBookError,
  assertExpandable,
  crc32,
  entrySize,
  loadZip,
  msgSharedStringsTooLarge,
  msgTextTooLarge,
  msgTooLarge,
  readEntryBytes,
  sizeText,
  zipDirectoryInfo,
} from '../../src/core/xlsx/zip';
import { cellAt, fx, makePackage } from '../fixtures/xlsx/helpers';

const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const MiB = 1 << 20;

async function readErr(p: Promise<unknown>): Promise<ReadBookError> {
  try {
    await p;
  } catch (e) {
    expect(e).toBeInstanceOf(ReadBookError);
    return e as ReadBookError;
  }
  throw new Error('expected a ReadBookError');
}

// ---------------------------------------------------------------------------------------------
// Hand-made zips
// ---------------------------------------------------------------------------------------------

const SEG = new Uint8Array(deflateRawSync(Buffer.alloc(MiB), { finishFlush: constants.Z_FULL_FLUSH }));
const FINAL = new Uint8Array([0x03, 0x00]);

/** Raw deflate stream that inflates to `mib` MiB of zeros. */
function zerosDeflate(mib: number): Uint8Array {
  const out = new Uint8Array(SEG.length * mib + FINAL.length);
  for (let i = 0; i < mib; i++) out.set(SEG, i * SEG.length);
  out.set(FINAL, SEG.length * mib);
  return out;
}

interface ZEntry {
  name: string;
  /** stored bytes (method 0) or a raw deflate stream (method 8) */
  data: Uint8Array;
  method?: 0 | 8;
  /** declared uncompressed size (default: data length for stored entries) */
  usize?: number;
  crc?: number;
  /** write usize as 0xFFFFFFFF + a ZIP64 extra field holding it (8 bytes) */
  zip64?: boolean;
  /** extra central-directory records pointing at the same local entry (overlapping-entry zip bomb) */
  copies?: number;
}

function buildZip(entries: ZEntry[]): Uint8Array {
  const enc = new TextEncoder();
  const chunks: Uint8Array[] = [];
  const central: Uint8Array[] = [];
  let off = 0;
  let records = 0;
  for (const e of entries) {
    const name = enc.encode(e.name);
    const method = e.method ?? 0;
    const usize = e.usize ?? e.data.length;
    const crc = e.crc ?? (method === 0 ? crc32(e.data) : 0);
    const extra = e.zip64 ? new Uint8Array(12) : new Uint8Array(0);
    if (e.zip64) {
      const dv = new DataView(extra.buffer);
      dv.setUint16(0, 1, true);
      dv.setUint16(2, 8, true);
      dv.setUint32(4, usize % 0x100000000, true);
      dv.setUint32(8, Math.floor(usize / 0x100000000), true);
    }
    const usize32 = e.zip64 ? 0xffffffff : usize;
    const local = new Uint8Array(30 + name.length);
    const lv = new DataView(local.buffer);
    lv.setUint32(0, 0x04034b50, true);
    lv.setUint16(4, 20, true);
    lv.setUint16(6, 0x0800, true);
    lv.setUint16(8, method, true);
    lv.setUint16(12, 0x21, true);
    lv.setUint32(14, crc >>> 0, true);
    lv.setUint32(18, e.data.length, true);
    lv.setUint32(22, usize32 >>> 0, true);
    lv.setUint16(26, name.length, true);
    local.set(name, 30);
    const localOff = off;
    chunks.push(local, e.data);
    off += local.length + e.data.length;
    for (let k = 0; k <= (e.copies ?? 0); k++) {
      const c = new Uint8Array(46 + name.length + extra.length);
      const cv = new DataView(c.buffer);
      cv.setUint32(0, 0x02014b50, true);
      cv.setUint16(4, 20, true);
      cv.setUint16(6, 20, true);
      cv.setUint16(8, 0x0800, true);
      cv.setUint16(10, method, true);
      cv.setUint16(14, 0x21, true);
      cv.setUint32(16, crc >>> 0, true);
      cv.setUint32(20, e.data.length, true);
      cv.setUint32(24, usize32 >>> 0, true);
      cv.setUint16(28, name.length, true);
      cv.setUint16(30, extra.length, true);
      cv.setUint32(42, localOff, true);
      c.set(name, 46);
      c.set(extra, 46 + name.length);
      central.push(c);
      records++;
    }
  }
  const cdOff = off;
  let cdSize = 0;
  for (const c of central) cdSize += c.length;
  const eocd = new Uint8Array(22);
  const ev = new DataView(eocd.buffer);
  ev.setUint32(0, 0x06054b50, true);
  ev.setUint16(8, records, true);
  ev.setUint16(10, records, true);
  ev.setUint32(12, cdSize, true);
  ev.setUint32(16, cdOff, true);
  const all = [...chunks, ...central, eocd];
  const out = new Uint8Array(all.reduce((s, c) => s + c.length, 0));
  let p = 0;
  for (const c of all) {
    out.set(c, p);
    p += c.length;
  }
  return out;
}

/** Re-pack a package (all parts stored) and add / replace entries. */
async function repack(pkg: Uint8Array, replace: ZEntry[]): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(pkg);
  const byName = new Map(replace.map((e) => [e.name, e] as const));
  const entries: ZEntry[] = [];
  for (const name of Object.keys(zip.files)) {
    if (zip.files[name].dir) continue;
    entries.push(byName.get(name) ?? { name, data: await zip.files[name].async('uint8array') });
    byName.delete(name);
  }
  return buildZip([...entries, ...byName.values()]);
}

const SIMPLE_SHEET = '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><v>42</v></c></row>';
const SIMPLE_SST = `<sst xmlns="${NS}"><si><t>学生ID</t></si></sst>`;

// ---------------------------------------------------------------------------------------------
// 1. FeatureRule.type
// ---------------------------------------------------------------------------------------------

describe('FeatureRule.type from <cfRule type> / <dataValidation type>', () => {
  it('torture.xlsx: every CF / DV rule (main + x14) carries its type; Strict twin the same (its own order)', async () => {
    for (const file of ['torture.xlsx', 'strict.xlsx']) {
      const book = await readBook(fx(file), file);
      const d = book.sheets.find((s) => s.name === 'Data')!;
      const got = d.features.rules.map((r) => [r.kind, r.type, r.sqref, !!r.ext]);
      // Excel's Strict save writes the rules in another order
      if (file === 'strict.xlsx') got.sort((a, b) => String(a).localeCompare(String(b)));
      const want = [
        ['cf', 'cellIs', 'E4:E23', false],
        ['cf', 'expression', 'B4:B23', false],
        ['cf', 'colorScale', 'G4:G23', false],
        ['cf', 'dataBar', 'H4:H23', false],
        ['cf', 'iconSet', 'F4:F23', false],
        ['dv', 'whole', 'G4:G23', false],
        ['cf', 'dataBar', 'H4:H23', true],
        ['dv', 'list', 'C4:C23', true],
      ];
      if (file === 'strict.xlsx') want.sort((a, b) => String(a).localeCompare(String(b)));
      expect(got, file).toEqual(want);
      // no rule without a type anywhere in the book
      for (const s of book.sheets) for (const r of s.features.rules) expect(r.type, `${file} ${s.name} ${r.sqref}`).toBeTruthy();
    }
  });

  it('all CF types incl. x14-only ones; DV without type = "none" (schema default); x14 DV type', async () => {
    const x14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main';
    const xm = 'http://schemas.microsoft.com/office/excel/2006/main';
    const rules =
      '<conditionalFormatting sqref="A1:A9"><cfRule type="top10" priority="1" rank="3" dxfId="0"/></conditionalFormatting>' +
      '<conditionalFormatting sqref="B1:B9"><cfRule type="duplicateValues" priority="2" dxfId="0"/><cfRule type="uniqueValues" priority="3" dxfId="0"/></conditionalFormatting>' +
      '<conditionalFormatting sqref="C1:C9"><cfRule type="aboveAverage" priority="4" dxfId="0"/><cfRule type="containsText" priority="5" operator="containsText" text="済"><formula>NOT(ISERROR(SEARCH("済",C1)))</formula></cfRule></conditionalFormatting>' +
      '<conditionalFormatting sqref="D1:D9"><cfRule type="timePeriod" priority="6" timePeriod="today"><formula>FLOOR(D1,1)=TODAY()</formula></cfRule><cfRule type="containsBlanks" priority="7"><formula>LEN(TRIM(D1))=0</formula></cfRule></conditionalFormatting>' +
      '<dataValidations count="2"><dataValidation allowBlank="1" showInputMessage="1" promptTitle="入力" prompt="氏名を入力" sqref="E1:E9"/><dataValidation type="textLength" operator="lessThanOrEqual" sqref="F1:F9"><formula1>10</formula1></dataValidation></dataValidations>' +
      `<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}"><x14:conditionalFormattings>` +
      '<x14:conditionalFormatting><x14:cfRule type="expression" priority="8" id="{A}"><xm:f>Lists!$A$1&lt;&gt;G1</xm:f><x14:dxf/></x14:cfRule><xm:sqref>G1:G9</xm:sqref></x14:conditionalFormatting>' +
      '<x14:conditionalFormatting><x14:cfRule type="iconSet" priority="9" id="{B}"><x14:iconSet iconSet="3Triangles"><x14:cfvo type="percent"><xm:f>0</xm:f></x14:cfvo><x14:cfvo type="num"><xm:f>Lists!$B$1</xm:f></x14:cfvo><x14:cfvo type="percent"><xm:f>67</xm:f></x14:cfvo></x14:iconSet></x14:cfRule><xm:sqref>H1:H9</xm:sqref></x14:conditionalFormatting>' +
      '</x14:conditionalFormattings></ext><ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}"><x14:dataValidations count="2">' +
      '<x14:dataValidation type="custom" allowBlank="1"><x14:formula1><xm:f>COUNTIF(Lists!$A:$A,I1)&gt;0</xm:f></x14:formula1><xm:sqref>I1:I9</xm:sqref></x14:dataValidation>' +
      '<x14:dataValidation allowBlank="1"><xm:sqref>J1:J9</xm:sqref></x14:dataValidation>' +
      '</x14:dataValidations></ext></extLst>';
    const pkg = await makePackage({
      sheets: [
        { name: 'S', xml: `<?xml version="1.0"?><worksheet xmlns="${NS}" xmlns:x14="${x14}" xmlns:xm="${xm}"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>${rules}</worksheet>` },
        { name: 'Lists', xml: '<row r="1"><c r="A1"><v>1</v></c></row>' },
      ],
    });
    const book = await readBook(pkg, 'rules.xlsx');
    expect(book.sheets[0].features.rules).toEqual([
      { sqref: 'A1:A9', type: 'top10', formulas: [], kind: 'cf' },
      { sqref: 'B1:B9', type: 'duplicateValues', formulas: [], kind: 'cf' },
      { sqref: 'B1:B9', type: 'uniqueValues', formulas: [], kind: 'cf' },
      { sqref: 'C1:C9', type: 'aboveAverage', formulas: [], kind: 'cf' },
      { sqref: 'C1:C9', type: 'containsText', formulas: ['NOT(ISERROR(SEARCH("済",C1)))'], kind: 'cf' },
      { sqref: 'D1:D9', type: 'timePeriod', formulas: ['FLOOR(D1,1)=TODAY()'], kind: 'cf' },
      { sqref: 'D1:D9', type: 'containsBlanks', formulas: ['LEN(TRIM(D1))=0'], kind: 'cf' },
      { sqref: 'E1:E9', type: 'none', formulas: [], kind: 'dv' },
      { sqref: 'F1:F9', type: 'textLength', formulas: ['10'], kind: 'dv' },
      { sqref: 'G1:G9', type: 'expression', formulas: ['Lists!$A$1<>G1'], kind: 'cf', ext: true },
      { sqref: 'H1:H9', type: 'iconSet', formulas: ['Lists!$B$1'], kind: 'cf', ext: true },
      { sqref: 'I1:I9', type: 'custom', formulas: ['COUNTIF(Lists!$A:$A,I1)>0'], kind: 'dv', ext: true },
      { sqref: 'J1:J9', type: 'none', formulas: [], kind: 'dv', ext: true },
    ]);
  });

  it('a cfRule without a type attribute keeps type undefined (never invented)', async () => {
    const pkg = await makePackage({ sheets: [{ name: 'S', xml: `<?xml version="1.0"?><worksheet xmlns="${NS}"><sheetData/><conditionalFormatting sqref="A1"><cfRule priority="1"><formula>A1&gt;0</formula></cfRule></conditionalFormatting></worksheet>` }] });
    const r = (await readBook(pkg, 'x.xlsx')).sheets[0].features.rules;
    expect(r).toEqual([{ sqref: 'A1', formulas: ['A1>0'], kind: 'cf' }]);
    expect('type' in r[0]).toBe(false);
  });
});

// ---------------------------------------------------------------------------------------------
// 2. SourceBook.encoding / 3. SourceFormat 'xlsb'
// ---------------------------------------------------------------------------------------------

describe('SourceBook.encoding for text inputs', () => {
  it.each([
    ['sjis.csv', 'Shift_JIS', 'Shift_JIS'],
    ['legacy-excel.csv', 'Shift_JIS', 'Shift_JIS'],
    ['utf8bom.csv', 'UTF-8', 'UTF-8'],
    ['legacy-utf8.csv', 'UTF-8', 'UTF-8'],
    ['plain.tsv', 'UTF-8', 'UTF-8'],
    ['legacy-unicode.txt', 'UTF-16LE', 'UTF-16'],
  ])('%s → encoding %s; warnings[0] unchanged (文字コード: %s)', async (file, encoding, label) => {
    const book = await readBook(fx(file), file);
    expect(book.encoding).toBe(encoding);
    expect(book.format).toBe('csv');
    const kind = file.endsWith('.tsv') ? 'TSV' : 'CSV';
    expect(book.warnings[0]).toBe(`${kind}ファイルを読み込みました（文字コード: ${label}）。先頭が0のコード（00123など）や日付は文字のまま扱います。`);
  });

  it('EUC-JP, ISO-2022-JP (7-bit JIS), UTF-16BE with BOM', async () => {
    const text = 'ID,氏名\r\n00123,髙橋 一郎\r\n';
    const bytesOf = (to: Encoding.Encoding) => new Uint8Array(Encoding.convert(Encoding.stringToCode(text.replace('髙', '高')), { to, from: 'UNICODE' }));
    const euc = await readBook(bytesOf('EUCJP'), 'euc.csv');
    expect([euc.encoding, euc.sheets[0].rows[1]![1]?.v]).toEqual(['EUC-JP', '高橋 一郎']);
    const jisBytes = bytesOf('JIS');
    expect(jisBytes.every((b) => b < 0x80)).toBe(true); // valid UTF-8 as well: recognised by its escape sequences
    expect(decodeText(jisBytes).encoding).toBe('JIS');
    const jis = await readBook(jisBytes, 'jis.csv');
    expect([jis.encoding, jis.sheets[0].rows[1]![0]?.v, jis.sheets[0].rows[1]![1]?.v]).toEqual(['ISO-2022-JP', '00123', '高橋 一郎']);
    expect(jis.warnings[0]).toContain('（文字コード: JIS）');
    const be = new Uint8Array(2 + text.length * 2);
    be.set([0xfe, 0xff]);
    for (let i = 0; i < text.length; i++) [be[2 + i * 2], be[3 + i * 2]] = [text.charCodeAt(i) >> 8, text.charCodeAt(i) & 0xff];
    const b16 = await readBook(be, 'be.txt');
    expect([b16.encoding, b16.sheets[0].rows[1]![1]?.v]).toEqual(['UTF-16BE', '髙橋 一郎']);
    // ASCII text with an ESC that is not a JIS escape stays UTF-8
    expect(decodeText(new TextEncoder().encode('a,\u001b[31mred\n')).encoding).toBe('UTF8');
  });

  it('HTML "xls" reports its encoding too; xlsx / xls / ods / xlsb inputs have none', async () => {
    expect((await readBook(fx('html-sjis.xls'), 'html-sjis.xls')).encoding).toBe('Shift_JIS');
    for (const f of ['torture.xlsx', 'legacy.xls', 'legacy.ods', 'legacy.xlsb']) expect((await readBook(fx(f), f)).encoding, f).toBeUndefined();
  });
});

describe("SourceFormat 'xlsb'", () => {
  it('legacy.xlsb → format xlsb, values fidelity, generated package, xlsb warning', async () => {
    const book = await readBook(fx('legacy.xlsb'), 'legacy.xlsb');
    expect([book.format, book.fidelity]).toEqual(['xlsb', 'values']);
    expect(book.warnings[0]).toBe('xlsb形式のファイルは値（と表示形式）だけを読み込みました。数式・色・罫線などは引き継がれません。');
    expect(book.size).toBe(fx('legacy.xlsb').length);
    expect(book.bytes[0]).toBe(0x50); // a generated OOXML package, not the binary workbook
    const again = await readBook(book.bytes, 'legacy_xlsb.xlsx');
    expect(again.format).toBe('xlsx');
    // a .xlsb is recognised by its content, whatever its name
    expect((await readBook(fx('legacy.xlsb'), 'renamed.xlsx')).format).toBe('xlsb');
  });
});

// ---------------------------------------------------------------------------------------------
// 4. Zip-bomb guard and part size checks
// ---------------------------------------------------------------------------------------------

describe('zip central directory, sizes, CRC', () => {
  it('zipDirectoryInfo agrees with JSZip on every real fixture', async () => {
    for (const f of ['torture.xlsx', 'strict.xlsx', 'perf_10000x40.xlsx', 'legacy.xlsb', 'legacy.ods', 'interview/学生回答.xlsx', 'theme-probe.xlsx']) {
      const bytes = fx(f);
      const zip = await JSZip.loadAsync(bytes);
      let total = 0;
      for (const name of Object.keys(zip.files)) total += entrySize(zip.files[name]) ?? 0; // folders: no size
      const info = zipDirectoryInfo(bytes)!;
      expect(info, f).toBeDefined();
      expect(info.uncompressedTotal, f).toBe(total);
      expect(info.entries, f).toBe(Object.keys(zip.files).length);
    }
    expect(zipDirectoryInfo(new Uint8Array(100))).toBeUndefined();
    const truncated = fx('torture.xlsx').slice(0, 5000); // no end-of-central-directory record
    expect(zipDirectoryInfo(truncated)).toBeUndefined();
  });

  it('crc32 matches zlib; sizeText', () => {
    const s = new TextEncoder().encode('123456789');
    expect(crc32(s)).toBe(0xcbf43926);
    const big = fx('torture.xlsx');
    expect(crc32(big)).toBe(nodeCrc32(big) >>> 0);
    expect(crc32(new Uint8Array(0))).toBe(0);
    expect(sizeText(1.8 * 1024 ** 3)).toBe('約 1.8 GB');
    expect(sizeText(620 * MiB)).toBe('約 620 MB');
    expect(sizeText(10)).toBe('約 1 MB');
    expect(sizeText(Infinity)).toBe('2 GB 以上');
  });

  it('the zero-stream builder really inflates to N MiB (sanity of the fixtures below)', () => {
    const z = inflateRawSync(zerosDeflate(3));
    expect(z.length).toBe(3 * MiB);
    expect(z.every((b) => b === 0)).toBe(true);
    expect(SEG.length).toBeLessThan(4096);
  });
});

describe('zip-bomb guard: refused before anything is inflated', () => {
  it('a package whose parts declare > 1.5 GB in total → Japanese tooLarge error (honest 1.6 GB part)', async () => {
    const base = await makePackage({ sheets: [{ name: 'S', xml: SIMPLE_SHEET }], sst: SIMPLE_SST });
    const bomb = await repack(base, [{ name: 'xl/media/image1.png', data: zerosDeflate(1640), method: 8, usize: 1640 * MiB }]);
    expect(bomb.length).toBeLessThan(3 * MiB);
    const t0 = performance.now();
    const e = await readErr(readBook(bomb, 'bomb.xlsx'));
    expect(performance.now() - t0).toBeLessThan(1000);
    expect(e.code).toBe('tooLarge');
    expect(e.message).toBe(msgTooLarge(zipDirectoryInfo(bomb)!.uncompressedTotal));
    expect(e.message).toMatch(/^ファイルを展開すると大きすぎるため読み込めません（展開後 約 1\.6 GB、上限 1\.5 GB）。/);
    // just under the limit is accepted (the unused part is never inflated)
    const ok = await repack(base, [{ name: 'xl/media/image1.png', data: zerosDeflate(1400), method: 8, usize: 1400 * MiB }]);
    const book = await readBook(ok, 'ok.xlsx');
    expect(cellAt(book.sheets[0], 'A1')?.v).toBe('学生ID');
    expect(book.features.images).toBe(1);
  });

  it('overlapping entries (many directory records → one deflate stream) count once per record', async () => {
    const base = await makePackage({ sheets: [{ name: 'S', xml: SIMPLE_SHEET }] });
    const bomb = await repack(base, [{ name: 'xl/media/big.bin', data: zerosDeflate(100), method: 8, usize: 100 * MiB, copies: 19 }]);
    expect(zipDirectoryInfo(bomb)!.uncompressedTotal).toBeGreaterThan(MAX_UNCOMPRESSED_TOTAL);
    // JSZip keys entries by their local name: its own view sees the stream once (100 MB)
    const js = await JSZip.loadAsync(bomb);
    expect(entrySize(js.files['xl/media/big.bin'])).toBe(100 * MiB);
    expect((await readErr(readBook(bomb, 'overlap.xlsx'))).code).toBe('tooLarge');
  });

  it('more than 65,535 entries (16-bit entry count wraps in the EOCD): every directory record is counted', async () => {
    const entries: ZEntry[] = [];
    for (let i = 0; i < 70000; i++) entries.push({ name: `p/${i}.bin`, data: FINAL, method: 8, usize: 24000 });
    const zip = buildZip(entries);
    expect(new DataView(zip.buffer).getUint16(zip.length - 22 + 10, true)).toBe(70000 - 65536);
    const info = zipDirectoryInfo(zip)!;
    expect(info.entries).toBe(70000);
    expect(info.uncompressedTotal).toBe(70000 * 24000);
    expect((await readErr(readBook(zip, 'many.xlsx'))).code).toBe('tooLarge');
  });

  it('ZIP64 sizes are read with full precision (JSZip wraps them to 32 bits)', async () => {
    const base = await makePackage({ sheets: [{ name: 'S', xml: SIMPLE_SHEET }] });
    const five = 5 * 1024 ** 3 + 4096; // low 32 bits = 1 GiB + 4 KiB: what a 32-bit reader sees
    const bomb = await repack(base, [{ name: 'xl/media/huge.bin', data: zerosDeflate(4), method: 8, usize: five, zip64: true }]);
    expect(zipDirectoryInfo(bomb)!.uncompressedTotal).toBeGreaterThanOrEqual(five);
    const e = await readErr(readBook(bomb, 'zip64.xlsx'));
    expect(e.code).toBe('tooLarge');
    expect(e.message).toContain('展開後 約 5.0 GB');
    expect(() => assertExpandable(bomb)).toThrow(ReadBookError);
  });

  it('a 32-bit size >= 2 GB (negative in JSZip) is refused; .ods / .xlsb paths (restoreZip) are guarded too', async () => {
    const base = await makePackage({ sheets: [{ name: 'S', xml: SIMPLE_SHEET }] });
    const bomb = await repack(base, [{ name: 'xl/media/x.bin', data: zerosDeflate(2), method: 8, usize: 0xfff00000 }]);
    expect((await readErr(readBook(bomb, 'neg.xlsx'))).code).toBe('tooLarge');
    const js = await JSZip.loadAsync(bomb);
    expect(entrySize(js.files['xl/media/x.bin'])).toBeLessThan(0);
    expect(() => assertExpandable(new Uint8Array(0), js)).toThrow(msgTooLarge(Infinity));
    // restoreZip (the SheetJS pre-pass for .ods / .xlsb) refuses the same bombs itself
    const ods = buildZip([
      { name: 'mimetype', data: new TextEncoder().encode('application/vnd.oasis.opendocument.spreadsheet') },
      { name: 'content.xml', data: zerosDeflate(1600), method: 8, usize: 1600 * MiB },
    ]);
    expect((await readErr(restoreZip(ods))).code).toBe('tooLarge');
    expect((await readErr(readBook(ods, 'bomb.ods'))).code).toBe('tooLarge');
  });
});

describe('entries that inflate beyond their declared size are stopped early', () => {
  // 1 GiB of zeros in ~1 MB; declared as 1000 bytes. Inflating it all would take seconds and 1 GB of memory.
  const LIAR = zerosDeflate(1024);

  it('readEntryBytes rejects right after the declared size, fast', async () => {
    const zip = await JSZip.loadAsync(buildZip([{ name: 'a.xml', data: LIAR, method: 8, usize: 1000 }]));
    const t0 = performance.now();
    await expect(readEntryBytes(zip.files['a.xml'])).rejects.toThrow(/declared size/);
    expect(performance.now() - t0).toBeLessThan(1000);
  });

  it('a lying sharedStrings part → corrupt file (not silently empty strings); a lying sheet → empty sheet + warning', async () => {
    const base = await makePackage({ sheets: [{ name: 'S', xml: SIMPLE_SHEET }, { name: 'T', xml: SIMPLE_SHEET }], sst: SIMPLE_SST });
    const badSst = await repack(base, [{ name: 'xl/sharedStrings.xml', data: LIAR, method: 8, usize: 1000 }]);
    const e = await readErr(readBook(badSst, 'liar.xlsx'));
    expect([e.code, e.message]).toEqual(['corrupt', MSG_CORRUPT]);
    const badSheet = await repack(base, [{ name: 'xl/worksheets/sheet2.xml', data: LIAR, method: 8, usize: 1000 }]);
    const book = await readBook(badSheet, 'liar2.xlsx');
    expect(cellAt(book.sheets[0], 'A1')?.v).toBe('学生ID');
    expect(book.sheets[1].maxRow).toBe(0);
    expect(book.warnings).toContain('読み込めなかったシートが 1 個あります（空のシートとして扱います）。');
  });

  it('declared size larger than the real data, bad deflate data, CRC mismatch (restoreZip checks CRC)', async () => {
    const data = new TextEncoder().encode('<a/>');
    const zip = await JSZip.loadAsync(
      buildZip([
        { name: 'short.xml', data: new Uint8Array(deflateRawSync(data)), method: 8, usize: 10 },
        { name: 'garbage.xml', data: new Uint8Array([0xff, 0xff, 0xff, 0xff]), method: 8, usize: 4 },
        { name: 'ok.xml', data: new Uint8Array(deflateRawSync(data)), method: 8, usize: data.length, crc: crc32(data) },
        { name: 'badcrc.xml', data, usize: data.length, crc: 12345 },
      ]),
    );
    await expect(readEntryBytes(zip.files['short.xml'])).rejects.toThrow(/size mismatch/);
    await expect(readEntryBytes(zip.files['garbage.xml'])).rejects.toThrow();
    expect(await readEntryBytes(zip.files['ok.xml'], true)).toEqual(data);
    expect(await readEntryBytes(zip.files['badcrc.xml'])).toEqual(data);
    await expect(readEntryBytes(zip.files['badcrc.xml'], true)).rejects.toThrow(/CRC/);
    const pkg = new Package(zip);
    expect(await pkg.load('short.xml')).toEqual({ ok: false, problem: 'corrupt' });
    expect(await pkg.load('nope.xml')).toEqual({ ok: false, problem: 'missing' });
    expect(await pkg.load('OK.XML')).toEqual({ ok: true, text: '<a/>' });
    expect(await pkg.text('short.xml')).toBeNull();
    expect(pkg.size('ok.xml')).toBe(4);
    const ods = buildZip([{ name: 'content.xml', data, usize: data.length, crc: 999 }]);
    expect((await readErr(restoreZip(ods))).message).toBe(MSG_CORRUPT);
  });
});

describe('part size checks (V8 string limit): sharedStrings, sheets, text inputs', () => {
  it('sharedStrings declared beyond MAX_PART_BYTES → Japanese tooLarge error before inflating', async () => {
    const base = await makePackage({ sheets: [{ name: 'S', xml: SIMPLE_SHEET }], sst: SIMPLE_SST });
    const mib = Math.ceil(MAX_PART_BYTES / MiB) + 8; // 520 MiB: under the 1.5 GB total, over the part limit
    const big = await repack(base, [{ name: 'xl/sharedStrings.xml', data: zerosDeflate(mib), method: 8, usize: mib * MiB }]);
    const t0 = performance.now();
    const e = await readErr(readBook(big, 'sst.xlsx'));
    expect(performance.now() - t0).toBeLessThan(1000);
    expect(e.code).toBe('tooLarge');
    expect(e.message).toBe(msgSharedStringsTooLarge(mib * MiB));
    expect(e.message).toBe('文字データ（共有文字列）が大きすぎるため読み込めません（展開後 約 520 MB、上限 約 512 MB）。不要なシートや行を削除するか、ファイルを分けて保存し直してください。');
  });

  it('a sheet part beyond MAX_PART_BYTES → that sheet empty with a named warning; the others are read', async () => {
    const base = await makePackage({ sheets: [{ name: 'S', xml: SIMPLE_SHEET }, { name: '巨大', xml: SIMPLE_SHEET }], sst: SIMPLE_SST });
    const big = await repack(base, [{ name: 'xl/worksheets/sheet2.xml', data: zerosDeflate(600), method: 8, usize: 600 * MiB }]);
    const book = await readBook(big, 'sheet.xlsx');
    expect(cellAt(book.sheets[0], 'B1')?.v).toBe(42);
    expect(book.sheets[1]).toMatchObject({ name: '巨大', maxRow: 0, maxCol: 0 });
    expect(book.warnings).toEqual(['シート「巨大」は大きすぎるため読み込めませんでした（展開後 約 600 MB）。空のシートとして扱います。']);
  });

  it('text input beyond MAX_PART_BYTES → tooLarge (checked on the byte length, nothing decoded)', async () => {
    const bytes = new TextEncoder().encode('ID,氏名\n1,a\n');
    Object.defineProperty(bytes, 'length', { value: MAX_PART_BYTES + 1 });
    const e = await readErr(readBook(bytes, 'huge.csv'));
    expect([e.code, e.message]).toEqual(['tooLarge', msgTextTooLarge(MAX_PART_BYTES + 1)]);
    expect(e.message).toBe('ファイルが大きすぎるため読み込めません（約 512 MB、テキスト形式の上限 約 512 MB）。ファイルを分けて保存し直してください。');
  });

  it('normal packages still load through the guarded path (loadZip → Package.load)', async () => {
    const zip = await loadZip(fx('torture.xlsx'));
    const pkg = new Package(zip);
    const wb = await pkg.load('xl/workbook.xml');
    expect(wb.ok && wb.text.includes('<sheets>')).toBe(true);
    expect(MAX_UNCOMPRESSED_TOTAL).toBe(1610612736);
  });
});

// ---------------------------------------------------------------------------------------------
// Round 3: the API the patcher's package wrapper (patch/pkg.ts) calls — documented at the top of zip.ts
// ---------------------------------------------------------------------------------------------

describe('patcher API: loadZip + readEntryBytes(file, checkCrc, maxBytes)', () => {
  it('loadZip refuses a bomb before inflating anything (tooLarge, Japanese message), a broken zip is corrupt', async () => {
    const base = await makePackage({ sheets: [{ name: 'S', xml: SIMPLE_SHEET }] });
    const bomb = await repack(base, [{ name: 'xl/media/image1.png', data: zerosDeflate(1640), method: 8, usize: 1640 * MiB }]);
    const t0 = performance.now();
    const e = await readErr(loadZip(bomb));
    expect(performance.now() - t0).toBeLessThan(1000);
    expect([e.code, e.message]).toEqual(['tooLarge', msgTooLarge(zipDirectoryInfo(bomb)!.uncompressedTotal)]);
    // what a wrapper has to do to report it in its own error type
    expect(e).toBeInstanceOf(ReadBookError);
    expect(e.message).toMatch(/^ファイルを展開すると大きすぎるため読み込めません/);
    const broken = await readErr(loadZip(base.slice(0, 50)));
    expect([broken.code, broken.message]).toEqual(['corrupt', MSG_CORRUPT]);
    // overlapping entries and ZIP64 sizes are refused the same way
    const overlap = await repack(base, [{ name: 'xl/media/big.bin', data: zerosDeflate(100), method: 8, usize: 100 * MiB, copies: 19 }]);
    expect((await readErr(loadZip(overlap))).code).toBe('tooLarge');
  });

  it('assertExpandable(bytes) before / assertExpandable(bytes, zip) after a plain JSZip.loadAsync gives the same verdict', async () => {
    const base = await makePackage({ sheets: [{ name: 'S', xml: SIMPLE_SHEET }] });
    const bomb = await repack(base, [{ name: 'xl/media/x.bin', data: zerosDeflate(2), method: 8, usize: 0xfff00000 }]);
    expect(() => assertExpandable(bomb)).toThrow(ReadBookError);
    const js = await JSZip.loadAsync(bomb);
    expect(() => assertExpandable(new Uint8Array(0), js)).toThrow(msgTooLarge(Infinity));
    const ok = await JSZip.loadAsync(base);
    expect(() => assertExpandable(base, ok)).not.toThrow();
  });

  it('maxBytes: an entry that declares more is rejected before inflating, fast; at the limit it is read', async () => {
    const big = await JSZip.loadAsync(buildZip([{ name: 'xl/sharedStrings.xml', data: zerosDeflate(600), method: 8, usize: 600 * MiB }]));
    const t0 = performance.now();
    await expect(readEntryBytes(big.files['xl/sharedStrings.xml'], false, MAX_PART_BYTES)).rejects.toThrow(/larger than the allowed/);
    expect(performance.now() - t0).toBeLessThan(1000);
    const data = new TextEncoder().encode('<sst/>');
    const z = await JSZip.loadAsync(buildZip([{ name: 'a.xml', data: new Uint8Array(deflateRawSync(data)), method: 8, usize: data.length, crc: crc32(data) }]));
    expect(await readEntryBytes(z.files['a.xml'], true, data.length)).toEqual(data);
    await expect(readEntryBytes(z.files['a.xml'], true, data.length - 1)).rejects.toThrow(/larger than the allowed/);
    // without maxBytes nothing changes
    expect(await readEntryBytes(z.files['a.xml'], true)).toEqual(data);
  });

  it('a lying entry is stopped at its declared size with or without maxBytes', async () => {
    const zip = await JSZip.loadAsync(buildZip([{ name: 'a.xml', data: zerosDeflate(1024), method: 8, usize: 1000 }]));
    const t0 = performance.now();
    await expect(readEntryBytes(zip.files['a.xml'], false, MAX_PART_BYTES)).rejects.toThrow(/declared size/);
    await expect(readEntryBytes(zip.files['a.xml'])).rejects.toThrow(/declared size/);
    expect(performance.now() - t0).toBeLessThan(2000);
  });

  it('entries JSZip holds in memory (zip.file(name, string | Uint8Array)) read back; maxBytes still caps them', async () => {
    const zip = await loadZip(fx('torture.xlsx'));
    const text = '<?xml version="1.0"?><a>日本語 ≤ ✓</a>';
    zip.file('xl/new-text.xml', text);
    const raw = new Uint8Array([0, 1, 2, 3, 250, 251, 252]);
    zip.file('xl/new-bin.bin', raw);
    expect(entrySize(zip.files['xl/new-text.xml'])).toBeUndefined();
    const u8 = await readEntryBytes(zip.files['xl/new-text.xml'], false, MAX_PART_BYTES);
    expect(new TextDecoder().decode(u8)).toBe(text);
    expect(await readEntryBytes(zip.files['xl/new-bin.bin'])).toEqual(raw);
    await expect(readEntryBytes(zip.files['xl/new-bin.bin'], false, 3)).rejects.toThrow(/larger than the allowed/);
    // an untouched part of the same zip is still read through the guard and equals JSZip's own inflate
    const wb = await readEntryBytes(zip.files['xl/workbook.xml'], true, MAX_PART_BYTES);
    expect(wb).toEqual(await zip.files['xl/workbook.xml'].async('uint8array'));
  });

  it('a patcher-style wrapper loop over a real package reads every part through the guard (CRC checked)', async () => {
    const bytes = fx('torture.xlsx');
    const zip = await loadZip(bytes);
    let n = 0;
    for (const name of Object.keys(zip.files)) {
      if (zip.files[name].dir) continue;
      const guarded = await readEntryBytes(zip.files[name], true, MAX_PART_BYTES);
      expect(guarded.length, name).toBe(entrySize(zip.files[name]));
      n++;
    }
    expect(n).toBeGreaterThan(10);
  });
});
