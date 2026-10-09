/**
 * Legacy inputs through SheetJS → generated minimal .xlsx → package reader:
 *  - legacy.xls made by Excel (FileFormat 56) vs Excel's own Text / Value2 (excel-fixtures.json); torture.xls and
 *    date1904.xls (Excel's BIFF8 saves) vs Excel (xls-variants.json); .ods / .xlsb / XML 2003 / CSV / UTF-8 CSV /
 *    Unicode text saves of legacy.xls made by Excel
 *  - corrupt .xls on which plain SheetJS loops forever (sanitizeBiff8 regression, checked in a child process)
 *  - Shift_JIS / UTF-8 BOM CSV and TSV (gen/make-csv.mjs): '00123' stays text, CP932 characters survive;
 *    own CSV parser rules; Shift_JIS HTML "xls" (gen/make-html.mjs)
 *  - the generated packages were opened by Excel WITHOUT repair (com/verify-open.ps1 → generated/verify-open.json);
 *    the reader must still generate the same XML parts as those verified files.
 */
import { describe, expect, it } from 'vitest';
import { spawnSync } from 'node:child_process';
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import JSZip from 'jszip';
import * as XLSX from 'xlsx';
import Encoding from 'encoding-japanese';
import { readBook } from '../../src/core/xlsx/read';
import { cellText } from '../../src/core/xlsx/cellText';
import { biffWorkbookStream, classifyText, csvCell, declaredCharset, decodeText, detectDelimiter, parseDelimited, sanitizeBiff8 } from '../../src/core/xlsx/legacy';
import { MSG_CORRUPT, MSG_PASSWORD, ReadBookError } from '../../src/core/xlsx/zip';
import type { SourceBook } from '../../src/core/types';
import { FIXTURES, cellAt, comError, fx, fxJson, type ProbeCell } from '../fixtures/xlsx/helpers';

const CSV_ROWS = [
  ['学生ID', '氏名', '大学', '金額', 'メモ', '点数'],
  ['00123', '髙橋　一郎', '慶應義塾大学', '1,234', '①要確認', '85'],
  ['00456', '山田～太郎', '早稲田大学', '-5', '"引用"を含む', '72.5'],
  ['1001', '佐藤 花子', '東京大学', '1234567890123456', '改行を\n含む', '0'],
  ['', '鈴木－次郎', '', '3.14', 'a,b', '2026/10/20'],
];

async function partsOf(bytes: Uint8Array): Promise<Record<string, string>> {
  const zip = await JSZip.loadAsync(bytes);
  const out: Record<string, string> = {};
  for (const name of Object.keys(zip.files).sort()) if (!zip.files[name].dir) out[name] = await zip.file(name)!.async('string');
  return out;
}

describe('legacy.xls (BIFF8 made by Excel)', () => {
  it('values, types and display text equal Excel; fidelity values; hidden sheet; merges; widths', async () => {
    const book = await readBook(fx('legacy.xls'), 'legacy.xls');
    expect(book.format).toBe('xls');
    expect(book.fidelity).toBe('values');
    expect(book.size).toBe(fx('legacy.xls').length);
    expect(book.warnings[0]).toContain('xls形式');
    expect(book.sheets.map((s) => [s.name, s.state])).toEqual([
      ['Data', 'visible'],
      ['隠し', 'hidden'],
    ]);
    const s = book.sheets[0];
    const expected = fxJson<{ legacyXls: ProbeCell[] }>('excel-fixtures.json').legacyXls;
    for (const pc of expected) {
      const c = cellAt(s, pc.addr);
      expect(c, pc.addr).toBeDefined();
      if (pc.vt === 'n') expect(c!.v).toBeCloseTo(pc.v as number, 9);
      else if (pc.vt === 'e') expect(c!.v).toBe(comError(pc.v as number));
      else expect(c!.v, pc.addr).toBe(pc.v);
      expect(cellText(book, c), pc.addr).toBe(pc.text);
      expect(c!.f, pc.addr).toBeUndefined(); // values only
    }
    expect(cellAt(s, 'C2')).toMatchObject({ t: 'd', v: 46315 });
    expect(cellAt(s, 'E2')).toMatchObject({ t: 's', v: '00120' });
    expect(s.features.merges).toEqual([{ r1: 6, c1: 1, r2: 6, c2: 3 }]);
    expect(s.colProps[1]?.width).toBeCloseTo(14.625, 2);
    // the replacement bytes are a real xlsx package that reads back identically
    const again = await readBook(book.bytes, 'legacy.xlsx');
    expect(again.fidelity).toBe('full');
    expect(again.sheets[0].rows).toEqual(s.rows);
  });

  it('password-protected .xls → Japanese password message', async () => {
    await expect(readBook(fx('password.xls'), 'password.xls')).rejects.toThrow(MSG_PASSWORD);
  });
});

describe('Excel-made .xls variants of torture.xlsx / date1904.xlsx (com/make-xls-variants.ps1 → xls-variants.json)', () => {
  interface VarSheet {
    name: string;
    visible: number;
    cells: ProbeCell[];
  }
  const variants = fxJson<Record<'torture' | 'date1904', { date1904: boolean; sheets: VarSheet[] }>>('xls-variants.json');

  it('torture.xls: every value and display text equals Excel (formula results incl. dates / errors / strings)', async () => {
    const book = await readBook(fx('torture.xls'), 'torture.xls');
    expect(book.fidelity).toBe('values');
    expect(book.sheets.map((s) => [s.name, s.state])).toEqual(
      variants.torture.sheets.map((s) => [s.name, s.visible === -1 ? 'visible' : s.visible === 0 ? 'hidden' : 'veryHidden']),
    );
    const bad: string[] = [];
    let checked = 0;
    for (const vs of variants.torture.sheets) {
      const s = book.sheets.find((x) => x.name === vs.name)!;
      for (const pc of vs.cells) {
        const c = cellAt(s, pc.addr);
        const at = `${vs.name}!${pc.addr}`;
        checked++;
        if (!c) {
          bad.push(`${at} missing`);
          continue;
        }
        const ev = pc.vt === 'e' ? comError(pc.v as number) : pc.v;
        const same = pc.vt === 'n' ? typeof c.v === 'number' && Math.abs(c.v - (ev as number)) < 1e-9 : c.v === ev;
        if (!same) bad.push(`${at} ${c.t}:${JSON.stringify(c.v)} vs ${JSON.stringify(ev)}`);
        if (!/^#+$/.test(pc.text) && cellText(book, c) !== pc.text) bad.push(`${at} text ${JSON.stringify(cellText(book, c))} vs ${JSON.stringify(pc.text)}`);
      }
    }
    expect(bad).toEqual([]);
    expect(checked).toBeGreaterThan(400);
  });

  it('date1904.xls: dates converted to 1900 serials (Excel reports 1904 serials), formats of formula cells kept', async () => {
    const book = await readBook(fx('date1904.xls'), 'date1904.xls');
    expect(variants.date1904.date1904).toBe(true);
    expect(book.date1904).toBe(false); // the generated package is a 1900 book
    const s = book.sheets[0];
    const byAddr = Object.fromEntries(variants.date1904.sheets[0].cells.map((c) => [c.addr, c]));
    expect(cellAt(s, 'A1')).toMatchObject({ t: 'd', v: (byAddr.A1.v as number) + 1462 });
    expect(cellAt(s, 'B1')).toMatchObject({ t: 'd', v: (byAddr.B1.v as number) + 1462 }); // =A1+1
    expect(cellAt(s, 'A2')).toMatchObject({ t: 'd', v: byAddr.A2.v }); // 12:30 keeps its value
    expect(cellAt(s, 'A3')).toMatchObject({ t: 'n', v: 1234 });
    for (const a of ['A1', 'A2', 'A3', 'A4', 'B1']) if (!/^#+$/.test(byAddr[a].text)) expect(cellText(book, cellAt(s, a)), a).toBe(byAddr[a].text);
    expect(cellText(book, cellAt(s, 'A1'))).toBe('2026/10/20'); // Excel: ####### (column too narrow)
    expect(cellText(book, cellAt(s, 'B1'))).toBe('2026年10月21日');
  });
});

describe('corrupt .xls files that make SheetJS loop forever (found by fuzzing, see sanitizeBiff8)', () => {
  const script = join(FIXTURES, 'gen', 'sheetjs-read.mjs');

  it('sanitizeBiff8: only value records stay, formulas lose their tokens, BoundSheet offsets point at the moved BOFs', () => {
    const stream = biffWorkbookStream(fx('torture.xls'))!;
    const clean = sanitizeBiff8(stream)!;
    expect(clean.length).toBeLessThan(stream.length);
    const u16 = (p: number) => clean[p] | (clean[p + 1] << 8);
    const types = new Set<number>();
    const bofs = new Set<number>();
    const boundSheets: number[] = [];
    let formulas = 0;
    for (let p = 0; p + 4 <= clean.length; ) {
      const type = u16(p);
      const len = u16(p + 2);
      types.add(type);
      if (type === 0x0809) bofs.add(p);
      if (type === 0x0085) boundSheets.push(clean[p + 4] | (clean[p + 5] << 8) | (clean[p + 6] << 16) | (clean[p + 7] << 24));
      if (type === 0x0006) {
        formulas++;
        expect(len).toBe(22);
        expect(u16(p + 4 + 20)).toBe(0); // cce
      }
      p += 4 + len;
    }
    expect(formulas).toBeGreaterThan(20);
    for (const t of types) expect([0x0809, 0x000a, 0x0042, 0x0022, 0x041e, 0x00e0, 0x00fc, 0x003c, 0x0085, 0x0200, 0x0203, 0x027e, 0x00bd, 0x00fd, 0x0204, 0x00d6, 0x0205, 0x0006, 0x0207, 0x00e5, 0x007d]).toContain(t);
    expect(boundSheets.length).toBe(5);
    for (const off of boundSheets) expect(bofs.has(off)).toBe(true);
    // SheetJS reads the same values from the sanitized stream as from the original file
    const a = XLSX.read(fx('torture.xls'), { type: 'array', cellStyles: false });
    const b = XLSX.read(clean, { type: 'array', cellStyles: false });
    expect(b.SheetNames).toEqual(a.SheetNames);
    for (const n of a.SheetNames) {
      const va = Object.fromEntries(Object.entries(a.Sheets[n]).filter(([k]) => k[0] !== '!').map(([k, c]) => [k, [c.t, c.v]]));
      const vb = Object.fromEntries(Object.entries(b.Sheets[n]).filter(([k]) => k[0] !== '!').map(([k, c]) => [k, [c.t, c.v]]));
      expect(vb, n).toEqual(va);
    }
    expect(sanitizeBiff8(new Uint8Array(16))).toBeUndefined();
  });

  it.each(['theme-zip.xls', 'zeroed-records.xls', 'style-to-formula.xls'])(
    '%s: SheetJS finishes on what readBook gives it (child process, 20 s limit), readBook resolves or throws a Japanese error',
    async (file) => {
      const bytes = fx('corrupt/' + file);
      const clean = sanitizeBiff8(biffWorkbookStream(bytes)!)!;
      const dir = mkdtempSync(join(tmpdir(), 'rdr-'));
      try {
        const p = join(dir, 'clean.bin');
        writeFileSync(p, clean);
        const r = spawnSync(process.execPath, [script, p, JSON.stringify({ cellNF: true, cellStyles: false })], { timeout: 20000, encoding: 'utf8' });
        expect(r.error, 'SheetJS did not finish').toBeUndefined();
        expect(r.status).toBe(0);
      } finally {
        rmSync(dir, { recursive: true, force: true });
      }
      // only reached when SheetJS finished above, so this cannot hang the test run
      try {
        const book = await readBook(bytes, file);
        expect(book.fidelity).toBe('values');
      } catch (e) {
        expect(e).toBeInstanceOf(ReadBookError);
      }
    },
    60000,
  );
});

describe('the same workbook saved by Excel as .ods / .xlsb / XML Spreadsheet 2003 (com/make-ods-xlsb.ps1)', () => {
  const expected = fxJson<{ legacyXls: ProbeCell[] }>('excel-fixtures.json').legacyXls;

  it.each([
    ['legacy.ods', 'ods', 'ODS形式', [14.625, 11.375]],
    ['legacy.xlsb', 'xlsb', 'xlsb形式', [14.625, 11.375]],
    // SheetJS converts XML 2003 point widths with its own digit width: within one pixel
    ['legacy2003.xls', 'xls', 'XMLスプレッドシート形式', [14.5, 11.33]],
  ] as const)('%s: values, types, display text, hidden sheet, merges, widths equal the Excel-made .xls', async (file, format, label, widths) => {
    const book = await readBook(fx(file), file);
    expect(book.format).toBe(format);
    expect(book.fidelity).toBe('values');
    expect(book.warnings[0]).toContain(label);
    expect(book.sheets.map((s) => [s.name, s.state])).toEqual([
      ['Data', 'visible'],
      ['隠し', 'hidden'],
    ]);
    const s = book.sheets[0];
    for (const pc of expected) {
      const c = cellAt(s, pc.addr);
      expect(c, pc.addr).toBeDefined();
      if (pc.vt === 'n') expect(c!.v).toBeCloseTo(pc.v as number, 9);
      else if (pc.vt === 'e') expect(c!.v).toBe(comError(pc.v as number));
      else expect(c!.v, pc.addr).toBe(pc.v);
      expect(cellText(book, c), `${file} ${pc.addr}`).toBe(pc.text);
    }
    expect(cellAt(s, 'C2')).toMatchObject({ t: 'd', v: 46315 });
    expect(cellAt(s, 'E2')).toEqual({ t: 's', v: '00120' });
    expect(s.features.merges).toEqual([{ r1: 6, c1: 1, r2: 6, c2: 3 }]);
    expect(s.colProps[0]?.width).toBeUndefined();
    expect(s.colProps[1]?.width).toBeCloseTo(widths[0], 1);
    expect(s.colProps[2]?.width).toBeCloseTo(widths[1], 1);
    expect(book.sheets[1].colProps).toEqual([]);
    expect(cellAt(book.sheets[1], 'A1')).toEqual({ t: 's', v: 'secret' });
  });
});

describe('HTML / XML "xls" files exported by business systems', () => {
  it('Shift_JIS HTML table (meta charset): decoded, read like CSV (00123 stays text), sheet named after the file', async () => {
    const book = await readBook(fx('html-sjis.xls'), 'html-sjis.xls');
    expect(book.format).toBe('xls');
    expect(book.fidelity).toBe('values');
    expect(book.warnings[0]).toContain('HTML形式');
    const s = book.sheets[0];
    expect(s.name).toBe('html-sjis');
    expect([s.maxRow, s.maxCol]).toEqual([5, 6]);
    expect(s.rows[0]!.map((c) => c?.v)).toEqual(['学生ID', '氏名', '大学', '金額', 'メモ', '日付']);
    expect(cellAt(s, 'A2')).toEqual({ t: 's', v: '00123' });
    expect(cellAt(s, 'B2')).toEqual({ t: 's', v: '髙橋　一郎' });
    expect(cellAt(s, 'B3')).toEqual({ t: 's', v: '山田～太郎' });
    expect(cellAt(s, 'E2')).toEqual({ t: 's', v: '①要確認' });
    expect(cellAt(s, 'E3')).toEqual({ t: 's', v: 'A&B <注>' });
    expect(cellAt(s, 'E4')).toEqual({ t: 's', v: '改行を\n含む' });
    expect(cellAt(s, 'A4')).toEqual({ t: 'n', v: 1001 });
    expect(cellAt(s, 'D2')).toMatchObject({ t: 'n', v: 1234 });
    expect(cellText(book, cellAt(s, 'D2'))).toBe('1,234');
    // dates stay text (no time-zone shifted serials like SheetJS' own HTML date guessing)
    expect(cellAt(s, 'F2')).toEqual({ t: 's', v: '2026/10/20' });
    expect(cellAt(s, 'A5')).toEqual({ t: 's', v: '鈴木－次郎（合計行）' });
    expect(s.features.merges).toEqual([{ r1: 5, c1: 1, r2: 5, c2: 3 }]);
  });

  it('UTF-8 HTML without a declaration, .htm name → format other', async () => {
    const html = '<table><tr><td>ID</td><td>氏名</td></tr><tr><td>007</td><td>髙橋</td></tr></table>';
    const book = await readBook(new TextEncoder().encode(html), 'export.htm');
    expect(book.format).toBe('other');
    expect(book.sheets[0].rows[1]).toEqual([{ t: 's', v: '007' }, { t: 's', v: '髙橋' }]);
  });

  it('XML that is not a spreadsheet → "not an Excel file"', async () => {
    const xml = '<?xml version="1.0" encoding="UTF-8"?><config><item>1</item></config>';
    await expect(readBook(new TextEncoder().encode(xml), 'config.xls')).rejects.toThrow(MSG_CORRUPT);
  });

  it('classifyText / declaredCharset', () => {
    expect(classifyText('ID,氏名\n1,a')).toBe('delimited');
    expect(classifyText('﻿<html><body><table>')).toBe('html');
    expect(classifyText('  <table><tr><td>1</td></tr></table>')).toBe('html');
    expect(classifyText('<?xml version="1.0"?>\n<?mso-application progid="Excel.Sheet"?>\n<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet">')).toBe('xml2003');
    expect(classifyText('<?xml version="1.0"?><root/>')).toBe('xml');
    const enc = (s: string) => new TextEncoder().encode(s);
    expect(declaredCharset(enc('<html><head><meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS"></head>'))).toBe('Shift_JIS');
    expect(declaredCharset(enc('<meta charset="euc-jp">'))).toBe('euc-jp');
    expect(declaredCharset(enc('<?xml version="1.0" encoding="windows-31j"?><a/>'))).toBe('windows-31j');
    expect(declaredCharset(enc('ID,氏名'))).toBeUndefined();
  });

  it('decodeText honours a declared EUC-JP charset; a stale UTF-8 declaration on Shift_JIS bytes falls back to detection', () => {
    const body = '<meta charset="euc-jp"><table><tr><td>高橋 山田</td></tr></table>';
    const euc = new Uint8Array(Encoding.convert(Encoding.stringToCode(body), { to: 'EUCJP', from: 'UNICODE' }));
    expect(decodeText(euc)).toEqual({ text: body, encoding: 'EUCJP' });
    const stale = '<meta charset="utf-8"><table><tr><td>山田 太郎</td></tr></table>';
    const sjis = new Uint8Array(Encoding.convert(Encoding.stringToCode(stale), { to: 'SJIS', from: 'UNICODE' }));
    expect(decodeText(sjis)).toEqual({ text: stale, encoding: 'SJIS' });
  });
});

describe('CSV / TSV', () => {
  function expectRows(book: SourceBook) {
    const s = book.sheets[0];
    expect([s.maxRow, s.maxCol]).toEqual([5, 6]);
    CSV_ROWS.forEach((row, ri) =>
      row.forEach((text, ci) => {
        const c = s.rows[ri]?.[ci];
        if (text === '') expect(c, `r${ri + 1}c${ci + 1}`).toBeUndefined();
        else expect(cellText(book, c), `r${ri + 1}c${ci + 1}`).toBe(text);
      }),
    );
    // leading zeros / 16-digit codes / dates stay text; plain numbers become numbers
    expect(s.rows[1]![0]).toEqual({ t: 's', v: '00123' });
    expect(s.rows[3]![3]).toEqual({ t: 's', v: '1234567890123456' });
    expect(s.rows[4]![5]).toEqual({ t: 's', v: '2026/10/20' });
    expect(s.rows[3]![0]).toEqual({ t: 'n', v: 1001 });
    expect(s.rows[2]![5]).toEqual({ t: 'n', v: 72.5 });
    expect(s.rows[1]![3]).toMatchObject({ t: 'n', v: 1234 }); // "1,234" → 1234 shown as 1,234
    expect(book.format).toBe('csv');
    expect(book.fidelity).toBe('values');
  }

  it('Shift_JIS (CP932) CSV: detected and decoded (髙 ① ～ －), quoted commas / quotes / line breaks', async () => {
    const book = await readBook(fx('sjis.csv'), 'sjis.csv');
    expectRows(book);
    expect(book.sheets[0].name).toBe('sjis');
    expect(book.warnings[0]).toBe('CSVファイルを読み込みました（文字コード: Shift_JIS）。先頭が0のコード（00123など）や日付は文字のまま扱います。');
  });

  it('UTF-8 BOM CSV and TSV read the same', async () => {
    expectRows(await readBook(fx('utf8bom.csv'), 'utf8bom.csv'));
    expectRows(await readBook(fx('plain.tsv'), 'plain.tsv'));
  });

  it('decodeText: BOMs, UTF-8 without BOM, Shift_JIS', () => {
    expect(decodeText(new Uint8Array([0xef, 0xbb, 0xbf, 0x41])).text).toBe('A');
    expect(decodeText(new TextEncoder().encode('氏名,ID')).encoding).toBe('UTF8');
    expect(decodeText(fx('sjis.csv')).encoding).toBe('SJIS');
    expect(decodeText(new Uint8Array([0xff, 0xfe, 0x41, 0x00])).text).toBe('A');
  });

  it('csvCell: only unambiguous numbers are converted', () => {
    expect(csvCell('00123')).toEqual({ t: 's', v: '00123' });
    expect(csvCell('0')).toEqual({ t: 'n', v: 0 });
    expect(csvCell('0.5')).toEqual({ t: 'n', v: 0.5 });
    expect(csvCell('-12')).toEqual({ t: 'n', v: -12 });
    expect(csvCell('+12')).toEqual({ t: 's', v: '+12' });
    expect(csvCell('1e5')).toEqual({ t: 's', v: '1e5' });
    expect(csvCell('12,345,678')).toEqual({ t: 'n', v: 12345678, fmt: '#,##0' });
    expect(csvCell('1,23')).toEqual({ t: 's', v: '1,23' });
    expect(csvCell('10%')).toEqual({ t: 's', v: '10%' });
    expect(csvCell('1234567890123456')).toEqual({ t: 's', v: '1234567890123456' });
    expect(csvCell('1,234.5')).toEqual({ t: 'n', v: 1234.5, fmt: '#,##0.0' });
    expect(csvCell('-1,234.50')).toEqual({ t: 'n', v: -1234.5, fmt: '#,##0.00' });
    expect(csvCell('TRUE')).toEqual({ t: 'b', v: true });
    expect(csvCell('false')).toEqual({ t: 'b', v: false });
    expect(csvCell('#N/A')).toEqual({ t: 'e', v: '#N/A' });
    expect(csvCell('#div/0!')).toEqual({ t: 'e', v: '#DIV/0!' });
    expect(csvCell('#TAG')).toEqual({ t: 's', v: '#TAG' });
  });

  it.each([
    ['legacy-excel.csv', 'SJIS'],
    ['legacy-utf8.csv', 'UTF8'],
    ['legacy-unicode.txt', 'UTF16LE'],
  ])('%s (saved by Excel from legacy.xls): values and display text equal the source workbook', async (file, encoding) => {
    expect(decodeText(fx(file)).encoding).toBe(encoding);
    const book = await readBook(fx(file), file);
    expect(book.format).toBe('csv');
    const s = book.sheets[0];
    expect(s.name).toBe(file.replace(/\.[^.]+$/, ''));
    const expected = fxJson<{ legacyXls: ProbeCell[] }>('excel-fixtures.json').legacyXls;
    for (const pc of expected) {
      const c = cellAt(s, pc.addr);
      expect(c, `${file} ${pc.addr}`).toBeDefined();
      if (pc.addr[0] === 'C' && pc.addr !== 'C1') {
        expect(c).toEqual({ t: 's', v: pc.text }); // dates stay text ('2026/10/20') by design
        continue;
      }
      if (pc.vt === 'n') expect(c!.v, `${file} ${pc.addr}`).toBeCloseTo(pc.v as number, 9);
      else if (pc.vt === 'e') expect(c!.v).toBe(comError(pc.v as number));
      else expect(c!.v, `${file} ${pc.addr}`).toBe(pc.v);
      expect(cellText(book, c), `${file} ${pc.addr}`).toBe(pc.text);
    }
  });

  it('parseDelimited: Excel quoting rules, line endings, empty lines / fields', () => {
    const rows = (text: string, name = 'a.csv') => Array.from(parseDelimited(text, name).rows, (r) => (r ? Array.from(r, (c) => (c ? c.v : null)) : null));
    expect(rows('a,b\r\n1,"x,y"\r\n')).toEqual([['a', 'b'], [1, 'x,y']]);
    expect(rows('"say ""hi""",b')).toEqual([['say "hi"', 'b']]);
    expect(rows('"line1\r\nline2",z\rnext')).toEqual([['line1\nline2', 'z'], ['next']]); // CR LF in quotes → LF; lone CR ends a record
    expect(rows('"ab"cd,e')).toEqual([['abcd', 'e']]); // text after the closing quote is appended (Excel)
    expect(rows('"unterminated,x\ny')).toEqual([['unterminated,x\ny']]);
    expect(rows('a,,c,\n\n\nd')).toEqual([['a', null, 'c'], null, null, ['d']]);
    expect(rows('x"y,1')).toEqual([['x"y', 1]]);
    expect(rows('')).toEqual([]);
    expect(parseDelimited('a,b', 'data.csv').name).toBe('data');
  });

  it('delimiters: sep= line, TSV, TXT (tab by default), semicolon / tab CSV; content never sniffed as SYLK / XML', async () => {
    const rows = (text: string, name: string) => Array.from(parseDelimited(text, name).rows, (r) => (r ? Array.from(r, (c) => (c ? c.v : null)) : null));
    expect(rows('sep=;\nID;氏名\n00123;山田', 'x.csv')).toEqual([['ID', '氏名'], ['00123', '山田']]);
    expect(rows('ID;氏名\n1;山田', 'x.csv')).toEqual([['ID', '氏名'], [1, '山田']]); // "ID;" is the SYLK magic for SheetJS
    expect(rows('ID\t氏名\n1\t山田', 'x.csv')).toEqual([['ID', '氏名'], [1, '山田']]);
    expect(rows('a,b\tc', 'x.txt')).toEqual([['a,b', 'c']]);
    expect(rows('a,b', 'x.txt')).toEqual([['a', 'b']]);
    expect(rows('a,b\tc', 'x.tsv')).toEqual([['a,b', 'c']]);
    expect(detectDelimiter('"a;b",c', 'x.csv')).toBe(',');
    const book = await readBook(new TextEncoder().encode('ID;氏名;点数\n00123;山田;80\n'), 'semi.csv');
    expect(book.sheets[0].rows[1]).toEqual([{ t: 's', v: '00123' }, { t: 's', v: '山田' }, { t: 'n', v: 80 }]);
    const lt = await readBook(new TextEncoder().encode('<注>,値\n1,2\n'), 'lt.csv');
    expect(lt.sheets[0].rows[0]).toEqual([{ t: 's', v: '<注>' }, { t: 's', v: '値' }]);
  });

  it('a CSV mis-named .xlsx is still read (as CSV)', async () => {
    const book = await readBook(new TextEncoder().encode('ID,氏名\n1,山田\n'), 'data.xlsx');
    expect(book.fidelity).toBe('values');
    expect(book.sheets[0].rows[1]).toEqual([{ t: 'n', v: 1 }, { t: 's', v: '山田' }]);
  });
});

describe('generated packages are the ones Excel opened without repair', () => {
  const verified = fxJson<Array<{ file: string; opened: boolean; error: string | null; cells: ProbeCell[] }>>('generated/verify-open.json');

  it('Excel opened every generated package normally', () => {
    expect(verified.map((v) => [v.file, v.opened, v.error]).sort()).toEqual([
      ['date1904_xls.xlsx', true, null],
      ['html-sjis_xls.xlsx', true, null],
      ['legacy2003_xls.xlsx', true, null],
      ['legacy_ods.xlsx', true, null],
      ['legacy_xls.xlsx', true, null],
      ['legacy_xlsb.xlsx', true, null],
      ['plain_tsv.xlsx', true, null],
      ['sjis_csv.xlsx', true, null],
      ['torture_xls.xlsx', true, null],
      ['utf8bom_csv.xlsx', true, null],
    ]);
  });

  it.each([
    ['legacy.xls', 'legacy_xls.xlsx'],
    ['legacy.ods', 'legacy_ods.xlsx'],
    ['legacy.xlsb', 'legacy_xlsb.xlsx'],
    ['legacy2003.xls', 'legacy2003_xls.xlsx'],
    ['html-sjis.xls', 'html-sjis_xls.xlsx'],
    ['torture.xls', 'torture_xls.xlsx'],
    ['date1904.xls', 'date1904_xls.xlsx'],
    ['sjis.csv', 'sjis_csv.xlsx'],
    ['utf8bom.csv', 'utf8bom_csv.xlsx'],
    ['plain.tsv', 'plain_tsv.xlsx'],
  ])('%s → same XML parts as the verified %s; Excel shows the same text', async (input, gen) => {
    const book = await readBook(fx(input), input);
    expect(await partsOf(book.bytes)).toEqual(await partsOf(fx('generated/' + gen)));
    const v = verified.find((x) => x.file === gen)!;
    for (const pc of v.cells) {
      if (/^#+$/.test(pc.text)) continue;
      expect(cellText(book, cellAt(book.sheets[0], pc.addr)), `${gen} ${pc.addr}`).toBe(pc.text);
    }
  });
});
