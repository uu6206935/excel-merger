/**
 * Real inputs: the reference app's 4 interview workbooks, the R2 oracle storage / shared-formula workbooks and the
 * R1 shared-formula variants. Values and display text vs Excel (excel-probes.json), formulas vs SheetJS.
 */
import { describe, expect, it } from 'vitest';
import * as XLSX from 'xlsx';
import { readBook } from '../../src/core/xlsx/read';
import { cellText } from '../../src/core/xlsx/cellText';
import type { SourceBook } from '../../src/core/types';
import { cellAt, comError, fx, fxJson, probes } from '../fixtures/xlsx/helpers';

const INTERVIEW: Array<[string, string]> = [
  ['interview/人事回答.xlsx', 'interview1.xlsx'],
  ['interview/学生-職場マッチング.xlsx', 'interview2.xlsx'],
  ['interview/学生回答.xlsx', 'interview3.xlsx'],
  ['interview/職場回答.xlsx', 'interview4.xlsx'],
];

function compareWithExcel(book: SourceBook, probeFile: string): { checked: number; bad: string[] } {
  const bad: string[] = [];
  let checked = 0;
  const pf = probes(probeFile);
  expect(book.sheets.map((s) => s.name)).toEqual(pf.sheets.map((s) => s.name));
  for (const ps of pf.sheets) {
    const s = book.sheets.find((x) => x.name === ps.name)!;
    for (const pc of ps.cells) {
      if (pc.vt === 'null') continue;
      const c = cellAt(s, pc.addr);
      const at = `${ps.name}!${pc.addr}`;
      checked++;
      if (!c) {
        bad.push(`${at} missing`);
        continue;
      }
      if (pc.vt === 'n') {
        if (!(c.t === 'n' || c.t === 'd') || Math.abs((c.v as number) - (pc.v as number)) > 1e-9) bad.push(`${at} value ${c.t}:${c.v} vs ${pc.v}`);
      } else if (pc.vt === 'e') {
        if (c.t !== 'e' || c.v !== comError(pc.v as number)) bad.push(`${at} error ${c.v}`);
      } else if (pc.vt === 'b') {
        if (c.t !== 'b' || c.v !== pc.v) bad.push(`${at} bool ${c.v}`);
      } else if (c.t !== 's' || c.v !== pc.v) bad.push(`${at} string ${JSON.stringify(c.v)} vs ${JSON.stringify(pc.v)}`);
      if (!/^#+$/.test(pc.text)) {
        const t = cellText(book, c);
        if (t !== pc.text) bad.push(`${at} text ${JSON.stringify(t)} vs ${JSON.stringify(pc.text)}`);
      }
    }
  }
  return { checked, bad };
}

/**
 * Formula text vs SheetJS. `skip`: cells where SheetJS is known to be wrong (it does not shift whole-column /
 * whole-row references when expanding shared formulas); those are asserted against Excel separately.
 */
function compareFormulasWithSheetJs(book: SourceBook, bytes: Uint8Array, skip: string[] = []): number {
  const wb = XLSX.read(bytes, { type: 'array', cellFormula: true, xlfn: true });
  let n = 0;
  for (const s of book.sheets) {
    const ws = wb.Sheets[s.name];
    if (!ws) continue;
    for (const key of Object.keys(ws)) {
      if (key.startsWith('!') || skip.includes(`${s.name}!${key}`)) continue;
      const sj = ws[key] as XLSX.CellObject;
      if (!sj.f) continue;
      expect(cellAt(s, key)?.f, `${s.name}!${key}`).toBe(sj.f);
      n++;
    }
  }
  return n;
}

/** 'a.b.c'!A1 and a.b.c!A1 are the same reference (Excel's UI omits optional quotes the file keeps). */
function unquoteSheets(f: string): string {
  return f.replace(/'((?:[^']|'')+)'!/g, (_m, n: string) => (/^[^\s'()!,;:[\]{}"+\-*/^&=<>%]+$/.test(n.replace(/''/g, "'")) ? n + '!' : `'${n}'!`));
}

describe('interview inputs (reference app) vs Excel', () => {
  it.each(INTERVIEW)('%s: values and display text equal Excel', async (path, probeName) => {
    const bytes = fx(path);
    const book = await readBook(bytes, path.split('/').pop()!);
    expect(book.fidelity).toBe('full');
    const { checked, bad } = compareWithExcel(book, probeName);
    expect(bad).toEqual([]);
    expect(checked).toBeGreaterThan(50);
    compareFormulasWithSheetJs(book, bytes);
    // extents cover the data Excel reports
    for (const s of book.sheets) {
      const ps = probes(probeName).sheets.find((x) => x.name === s.name)!;
      for (const pc of ps.cells) {
        const m = /^([A-Z]+)(\d+)$/.exec(pc.addr)!;
        expect(Number(m[2])).toBeLessThanOrEqual(s.maxRow);
      }
    }
  });
});

describe('R2 oracle storage-test.xlsx (odd sheet names, table formulas, LET/LAMBDA, 3D refs)', () => {
  it('sheet names, values vs Excel, formulas vs SheetJS', async () => {
    const bytes = fx('storage-test.xlsx');
    const book = await readBook(bytes, 'storage-test.xlsx');
    const { bad } = compareWithExcel(book, 'storage-test.xlsx');
    expect(bad).toEqual([]);
    expect(compareFormulasWithSheetJs(book, bytes)).toBeGreaterThan(40);
  });

  it('cross-sheet formulas are stored exactly as Excel shows them (minus "=")', async () => {
    const book = await readBook(fx('storage-test.xlsx'), 'storage-test.xlsx');
    const quoting = fxJson<{ quoting: Record<string, string> }>('storage-results.json').quoting;
    const f = book.sheets.find((s) => s.name === 'F')!;
    const formulas = new Set(f.rows.map((r) => r?.[0]?.f).filter((x): x is string => !!x).map(unquoteSheets));
    for (const [name, ui] of Object.entries(quoting)) {
      expect(book.sheets.some((s) => s.name === name), name).toBe(true);
      expect(formulas.has(unquoteSheets(ui.slice(1))), ui).toBe(true);
    }
    const t = book.sheets.find((s) => s.name === 'T')!;
    expect(cellAt(t, 'E6')!.f).toBe('_xlfn.LET(_xlpm.x,B2,_xlpm.y,B3,_xlpm.x+_xlpm.y)');
    expect(cellAt(t, 'E20')!.f).toBe("SUM('Sheet1:TAX2025'!A1)");
    expect(t.features.tables.map((x) => x.name)).toEqual(['Table1']);
    expect(book.names).toEqual([
      { name: 'TAX_RATE', scope: null, refersTo: '0.1' },
      { name: '単価', scope: null, refersTo: 'T!$B$2' },
    ]);
  });
});

describe('shared formulas (FillDown / AutoFill / Range.Formula / R1C1 / Copy variants)', () => {
  it('sharedtest.xlsx: every expanded formula equals SheetJS', async () => {
    const bytes = fx('sharedtest.xlsx');
    const book = await readBook(bytes, 'sharedtest.xlsx');
    expect(compareFormulasWithSheetJs(book, bytes)).toBeGreaterThan(10);
    expect(book.warnings).toEqual([]);
  });

  it('shared-test.xlsx: expanded formulas equal SheetJS, and whole-column refs shift like Excel FillRight', async () => {
    const bytes = fx('shared-test.xlsx');
    const book = await readBook(bytes, 'shared-test.xlsx');
    expect(compareFormulasWithSheetJs(book, bytes, ['Data!H2'])).toBeGreaterThan(10);
    // r2/oracle/excel-shared.ps1: F2 = SUM(A:A)+COUNTIF(C2:C9,">0"), then Range('F2:H2').FillRight()
    const d = book.sheets.find((s) => s.name === 'Data')!;
    expect(cellAt(d, 'F2')!.f).toBe('SUM(A:A)+COUNTIF(C2:C9,">0")');
    expect(cellAt(d, 'G2')!.f).toBe('SUM(B:B)+COUNTIF(D2:D9,">0")');
    expect(cellAt(d, 'H2')!.f).toBe('SUM(C:C)+COUNTIF(E2:E9,">0")'); // SheetJS wrongly keeps B:B here
    expect(book.warnings).toEqual([]);
  });
});
