/**
 * buildPremerged tests on Excel-authored fixtures (tests/golden/com/make-fixtures.ps1; scenario in
 * tests/golden/premerge-scenario.ts). The outputs of the same scenario were opened in Excel without repair and
 * recalculated (tests/golden/RESULTS.md).
 */
import fs from 'node:fs';
import path from 'node:path';
import { describe, expect, it } from 'vitest';
import { buildPremerged } from '../../src/core/xlsx/premerge';
import { readBook } from '../../src/core/xlsx/read';
import { replayLog } from '../../src/core/xlsx/replay';
import { buildXlsx, part } from '../golden/build';
import { GOLDEN_FIX, insideTableScenario, scenario, source } from '../golden/premerge-scenario';
import { snapshot, styleSigs } from '../golden/snapshot';

const JP = /[ぁ-んァ-ン一-龥]/;

describe('buildPremerged: plain range base (AutoFilter)', () => {
  it('appends rows, extends ranges, inserts imported columns with source styles and remapped formulas', async () => {
    const r = await buildPremerged(scenario('plain'));
    expect(r.layout).toEqual({
      sheet: '学生',
      insertAt: 5,
      imported: [
        { col: 5, sourceId: 'src', sourceCol: 3, header: '志望' },
        { col: 6, sourceId: 'src', sourceCol: 4, header: '評価' },
        { col: 7, sourceId: 'src', sourceCol: 5, header: '備考' },
        { col: 8, sourceId: 'src', sourceCol: 6, header: '点数（pm-src）' },
        { col: 9, sourceId: 'src', sourceCol: 7, header: 'Calc' },
        { col: 10, sourceId: 'src', sourceCol: 8, header: 'Total' },
        { col: 11, sourceId: 'src', sourceCol: 9, header: '氏名（pm-src）' },
      ],
      appendAt: 7,
      appended: [{ row: 7, sourceId: 'src', sourceRow: 4 }],
    });
    expect(r.report).toMatchObject({ mergedRows: 6, matchedRows: 4, baseOnlyRows: 1, appendedRows: 1, importedColumns: 7, formulasKept: 11, formulasToValues: 0 });
    // round 8: the 5 formulas that used to become values (Total = SUM($F$2:$F$6) over the whole source table, per row) are kept too
    expect(r.report.formulaStats).toEqual([{ sourceId: 'src', label: 'pm-src', kept: 11, toValues: 0, empty: 0, reasons: {} }]);
    const codes = r.report.items.map((i) => i.code);
    expect(codes).toEqual(expect.arrayContaining(['appendedRows', 'headerRenamed', 'filterExtended', 'formulasKept', 'formulaRangesWidened']));
    expect(r.report.items.every((i) => JP.test(i.message))).toBe(true);

    const snap = await snapshot(r.bytes);
    const s = snap.sheets.find((x) => x.name === '学生')!;
    // 合計 row moved below the appended row and includes it (range extension)
    expect(s.formulas.D9).toBe('SUM(D2:D7)');
    expect(s.formulas.B9).toBe('COUNTA(B2:B7)');
    expect(s.values.C9).toBe('s:合計');
    // appended row: key cells from keyFill, imported values from its source row
    expect(s.values.A7).toBe('s:S009');
    expect(s.values.B7).toBe('s:学生Z');
    expect(s.values.E7).toBe('s:C');
    // matched rows: values by key (S001 <- source row 3, S003 <- source row 2)
    expect(s.values.E2).toBe('s:B');
    expect(s.values.F2).toBe('n:3.25');
    expect(s.values.E4).toBe('s:A');
    // unmatched S004: no values
    for (const c of 'EFGHIJK') expect(s.values[`${c}5`]).toBeUndefined();
    // formulas: same-row refs remapped (F*2+D -> H*2+F), no-ref formula kept, column total -> cached value
    expect(s.formulas.I2).toBe('H2*2+F2');
    expect(s.formulas.I7).toBe('H7*2+F7');
    expect(s.formulas.G2).toBe('"line1"&CHAR(10)&"line2"');
    // Total = SUM($F$2:$F$6) over the whole source table (all 5 source rows are placed: 4 matched + 1 appended) -> the whole merged
    // table H2:H7; the source's result 362 stays (the same rows are summed, the unmatched base row holds nothing)
    expect(s.formulas.J2).toBe('SUM($H$2:$H$7)');
    expect((await readBook(r.bytes, 'm.xlsx')).sheets.find((x) => x.name === '学生')!.rows[1]![9]).toMatchObject({ f: 'SUM($H$2:$H$7)', v: 362 });
    // header cells
    expect(s.values.H1).toBe('s:点数（pm-src）');
    expect(s.values.K1).toBe('s:氏名（pm-src）');
    // autofilter over imported columns and the appended row; _FilterDatabase follows
    expect(s.autoFilter).toBe('A1:K7');
    expect(snap.names).toEqual(expect.arrayContaining(['Scores = 学生!$D$2:$D$7', '学生!_xlnm._FilterDatabase = 学生!$A$1:$K$7']));
    // other sheet + chart follow the extension
    const rep = snap.sheets.find((x) => x.name === 'Report')!;
    expect(rep.formulas.A1).toBe('AVERAGE(学生!D2:D7)');
    expect(snap.charts[0]).toContain('学生!$D$2:$D$7');
    // CF grows over the appended row
    expect(s.cf.some((c) => c.startsWith('D7:'))).toBe(true);
    // widths from the source columns
    expect(s.cols.E).toBe('w=18.625');
    expect(s.cols.G).toBe('w=22.625');
  });

  it('writes source styles (merged into the base styles.xml) and styles unmatched rows like the column', async () => {
    const r = await buildPremerged(scenario('plain'));
    const s = (await snapshot(r.bytes)).sheets[0];
    const srcSnap = (await snapshot(new Uint8Array(fs.readFileSync(path.join(GOLDEN_FIX, 'pm-src.xlsx'))))).sheets[0];
    // 志望 (source C) -> E : italic red font + yellow fill, identical signature
    expect(s.styles.E2).toBe(srcSnap.styles.C3);
    expect(s.styles.E5).toBe(srcSnap.styles.C3); // unmatched row: most common style of the column
    // 評価 number format 0.00"点" kept (custom id renumbered into the base numFmts)
    expect(s.styles.F2).toContain('nf=0.00"点"');
    // header style (bold + fill) from the source header
    expect(s.styles.E1).toBe(srcSnap.styles.C1);
    // base styles.xml keeps its old records (existing s indexes keep their meaning)
    const baseSigs = styleSigs((await part(new Uint8Array(fs.readFileSync(path.join(GOLDEN_FIX, 'pm-base-plain.xlsx'))), 'xl/styles.xml'))!);
    const outSigs = styleSigs((await part(r.bytes, 'xl/styles.xml'))!);
    expect(outSigs.slice(0, baseSigs.length)).toEqual(baseSigs);
    // deduplicated: importing again does not add xfs
    const again = await buildPremerged({ ...scenario('plain'), base: source('base', 'x.xlsx', scenario('plain').base.region, 'x', 'full', r.bytes) });
    const outSigs2 = styleSigs((await part(again.bytes, 'xl/styles.xml'))!);
    expect(new Set(outSigs2).size).toBe(new Set(outSigs).size);
    // appended row copies the last data row's format
    expect(s.styles.D7).toBe(s.styles.D6);
    expect(s.styles.A7).toBe(s.styles.A6);
  });

  it('formulaPolicy values converts every imported formula', async () => {
    const r = await buildPremerged(scenario('plain', { formulaPolicy: 'values' }));
    expect(r.report.formulasKept).toBe(0);
    expect(r.report.formulasToValues).toBe(11);
    const s = (await snapshot(r.bytes)).sheets[0];
    expect(s.formulas.I2).toBeUndefined();
    expect(s.values.I2).toBe('n:185.25');
  });

  it('headerConflict keep leaves headers as they are', async () => {
    const r = await buildPremerged(scenario('plain', { headerConflict: 'keep' }));
    expect(r.layout.imported.map((i) => i.header)).toEqual(['志望', '評価', '備考', '点数', 'Calc', 'Total', '氏名']);
  });

  it('extendTable false keeps the AutoFilter as it was (rows and columns)', async () => {
    const r = await buildPremerged(scenario('plain', { extendTable: false }));
    const s = (await snapshot(r.bytes)).sheets[0];
    expect(s.autoFilter).toBe('A1:D6');
    expect(s.formulas.D9).toBe('SUM(D2:D7)'); // formula ranges still grow
  });
});

describe('buildPremerged: ListObject base', () => {
  it('extends the table over the imported columns and the appended rows (unique column names = header cells)', async () => {
    const r = await buildPremerged(scenario('table'));
    const snap = await snapshot(r.bytes);
    const s = snap.sheets[0];
    expect(s.tables).toEqual(['tblStudents ref=A1:K7 hdr=1 tot=0 af=A1:K7 cols=[ID|氏名|大学|点数|志望|評価|備考|点数（pm-src）|Calc|Total|氏名（pm-src）] calc=[]']);
    expect(r.report.items.some((i) => i.code === 'tableExtended')).toBe(true);
    expect(s.formulas.D9).toBe('SUM(D2:D7)');
    // header cells equal the table column names (else Excel repairs the table)
    expect(['E1', 'F1', 'G1', 'H1', 'I1', 'J1', 'K1'].map((a) => s.values[a])).toEqual(['s:志望', 's:評価', 's:備考', 's:点数（pm-src）', 's:Calc', 's:Total', 's:氏名（pm-src）']);
  });

  it('makes table column names unique when headerConflict is keep', async () => {
    const r = await buildPremerged(scenario('table', { headerConflict: 'keep' }));
    const s = (await snapshot(r.bytes)).sheets[0];
    expect(s.tables[0]).toContain('cols=[ID|氏名|大学|点数|志望|評価|備考|点数2|Calc|Total|氏名2]');
    expect(s.values.H1).toBe('s:点数2');
    expect(r.layout.imported[3].header).toBe('点数2');
  });

  it('a region narrower than its table: imported columns land inside the table and are named after their headers', async () => {
    const r = await buildPremerged(insideTableScenario());
    const s = (await snapshot(r.bytes)).sheets.find((x) => x.name === '名簿')!;
    expect(s.tables[0]).toContain('ref=A3:I13');
    expect(s.tables[0]).toContain('cols=[ID|氏名|大学|点数1|点数2|志望|Calc|合計|列1]');
    expect(s.values.F3).toBe('s:志望');
    expect(s.values.G3).toBe('s:Calc');
    // Calc (=F*2+D in the source) refers to columns that were not imported -> value
    expect(s.formulas.G4).toBeUndefined();
    expect(r.report.formulasToValues).toBe(5);
  });

  it('the premerged package replays like any other (move an imported column into the table)', async () => {
    const r = await buildPremerged(scenario('table'));
    const out = await replayLog(r.bytes, [{ kind: 'struct', sheet: '学生', op: { type: 'move', axis: 'col', from: 9, count: 1, to: 2 } }]);
    const s = (await snapshot(out)).sheets[0];
    expect(s.values.B1).toBe('s:Calc');
    expect(s.formulas.B2).toBe('I2*2+G2');
    expect(s.tables[0]).toContain('cols=[ID|Calc|氏名|');
  });
});

describe('buildPremerged: other inputs', () => {
  it("'values' fidelity sources are imported without styles (General)", async () => {
    const vbytes = await buildXlsx([{ name: 'Sheet1', data: '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row><row r="2"><c r="A2" t="s"><v>2</v></c><c r="B2" s="1"><v>7</v></c></row>' }], { sst: ['ID', 'メモ', 'S002'] });
    const input = scenario('plain');
    input.sources = [source('v', 'v.csv', { sheet: 'Sheet1', headerRow: 1, firstDataRow: 2, lastDataRow: 2, firstCol: 1, lastCol: 2 }, 'v', 'values', vbytes)];
    input.rows = input.rows.filter((x) => x.baseRow != null).map((x) => ({ baseRow: x.baseRow, matches: { v: x.baseRow === 3 ? { row: 2, via: 'ID', flags: [] } : null } }));
    input.imports = [{ sourceId: 'v', col: 2 }];
    input.keyFill = {};
    const r = await buildPremerged(input);
    const s = (await snapshot(r.bytes)).sheets[0];
    expect(s.values.E3).toBe('n:7');
    expect(s.styles.E3).toBeUndefined();
    expect(s.values.E1).toBe('s:メモ');
    expect(r.report.items.some((i) => i.code === 'valuesOnly')).toBe(true);
  });

  it('converts theme colours to RGB when the source theme differs from the base theme', async () => {
    const input = scenario('plain');
    // torture.xlsx has a custom theme (accent1 = CC9933); Data!A1 is filled with a tinted theme colour
    input.sources = [source('t', 'torture.xlsx', { sheet: 'Data', headerRow: 3, firstDataRow: 4, lastDataRow: 23, firstCol: 1, lastCol: 13 }, 'torture')];
    input.rows = input.rows.filter((x) => x.baseRow != null).map((x) => ({ baseRow: x.baseRow, matches: { t: { row: 4, via: 'ID', flags: [] } } }));
    input.imports = [{ sourceId: 't', col: 1 }];
    input.keyFill = {};
    const r = await buildPremerged(input);
    const s = (await snapshot(r.bytes)).sheets[0];
    expect(s.styles.E1).toMatch(/fgColor rgb="FF[0-9A-F]{6}"/); // header fill (theme 5 + tint) converted
    expect(s.styles.E1).not.toMatch(/theme=/);
    expect(s.values.E2).toBe('n:1001');
  });

  it('rejects bad input with Japanese messages', async () => {
    const bad = scenario('plain');
    bad.base = { ...bad.base, region: { ...bad.base.region, sheet: 'ない' } };
    await expect(buildPremerged(bad)).rejects.toThrow(/シート「ない」/);
    const bad2 = scenario('plain');
    bad2.imports = [{ sourceId: 'zzz', col: 1 }];
    await expect(buildPremerged(bad2)).rejects.toThrow(JP);
    const bad3 = scenario('plain');
    bad3.base = { ...bad3.base, region: { ...bad3.base.region, lastCol: 16384 } };
    await expect(buildPremerged(bad3)).rejects.toThrow(/16,384/);
    const bad4 = scenario('plain');
    bad4.base = { ...bad4.base, book: { ...bad4.base.book, bytes: new Uint8Array([80, 75, 3, 4, 0]) } };
    await expect(buildPremerged(bad4)).rejects.toThrow(JP);
  });

  it('with nothing to import or append still produces a clean package', async () => {
    const input = scenario('plain');
    input.imports = [];
    input.rows = input.rows.filter((x) => x.baseRow != null);
    const r = await buildPremerged(input);
    expect(r.layout.imported).toEqual([]);
    expect(r.layout.appendAt).toBeNull();
    const snap = await snapshot(r.bytes);
    expect(snap.calcChain).toBe(false);
    expect(snap.fullCalcOnLoad).toBe(true);
  });

  it('scales: 10,000 base rows x 40 imported columns in reasonable time', async () => {
    const rows = 10000, cols = 40;
    const hdr = Array.from({ length: cols + 1 }, (_x, c) => `<c r="${colL(c + 1)}1" t="inlineStr"><is><t>H${c}</t></is></c>`).join('');
    const data = (n: number, withId: boolean) => {
      let x = `<row r="1">${hdr}</row>`;
      for (let r = 2; r <= n + 1; r++) {
        x += `<row r="${r}"><c r="A${r}"><v>${r}</v></c>`;
        if (withId) for (let c = 2; c <= cols + 1; c++) x += `<c r="${colL(c)}${r}" s="1"><v>${r * c}</v></c>`;
        x += '</row>';
      }
      return x;
    };
    const baseBytes = await buildXlsx([{ name: 'B', data: data(rows, false) }]);
    const srcBytes = await buildXlsx([{ name: 'S', data: data(rows, true) }]);
    const t0 = Date.now();
    const r = await buildPremerged({
      base: source('b', 'b.xlsx', { sheet: 'B', headerRow: 1, firstDataRow: 2, lastDataRow: rows + 1, firstCol: 1, lastCol: 1 }, 'b', 'full', baseBytes),
      sources: [source('s', 's.xlsx', { sheet: 'S', headerRow: 1, firstDataRow: 2, lastDataRow: rows + 1, firstCol: 1, lastCol: cols + 1 }, 's', 'full', srcBytes)],
      rows: Array.from({ length: rows }, (_x, i) => ({ baseRow: i + 2, matches: { s: { row: rows + 1 - i, via: 'ID', flags: [] } } })),
      imports: Array.from({ length: cols }, (_x, i) => ({ sourceId: 's', col: i + 2 })),
      keyFill: {},
      options: { formulaPolicy: 'keep-if-possible', extendTable: true, headerConflict: 'suffix' },
    });
    const ms = Date.now() - t0;
    expect(r.report.importedColumns).toBe(40);
    expect(ms).toBeLessThan(20000);
  }, 60000);
});

function colL(n: number): string {
  let s = '';
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

describe('date systems', () => {
  it('recognises date / time number formats', async () => {
    const { isDateCode } = await import('../../src/core/xlsx/patch/styles');
    for (const c of ['yyyy/m/d', 'm"月"d"日"(aaa)', '[$-411]ggge"年"m"月"d"日"', 'h:mm:ss', '[h]:mm', 'yyyy\-mm\-dd', 'mm:ss.0']) expect(isDateCode(c), c).toBe(true);
    for (const c of ['General', '0.00E+00', '#,##0 "個"', '[$¥-411]#,##0', '0.0%', '#,##0_ ;[Red]\-#,##0 ', '"day"0', '@']) expect(isDateCode(c), c).toBe(false);
  });

  it('moves date serials of a 1904-date source into the 1900 base date system', async () => {
    const input = scenario('plain');
    const bytes = await buildXlsx([{ name: 'D', data: '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" t="s"><v>2</v></c></row><row r="2"><c r="A2"><v>1</v></c><c r="B2" s="3"><v>44000</v></c><c r="C2"><v>44000</v></c></row>' }], { sst: ['ID', '日付', '数'] });
    const JSZip = (await import('jszip')).default;
    const z = await JSZip.loadAsync(bytes);
    z.file('xl/workbook.xml', (await z.file('xl/workbook.xml')!.async('string')).replace('<sheets>', '<workbookPr date1904="1"/><sheets>'));
    const b1904 = await z.generateAsync({ type: 'uint8array' });
    input.sources = [source('d', 'd.xlsx', { sheet: 'D', headerRow: 1, firstDataRow: 2, lastDataRow: 2, firstCol: 1, lastCol: 3 }, 'd', 'full', b1904)];
    input.rows = input.rows.filter((x) => x.baseRow != null).map((x) => ({ baseRow: x.baseRow, matches: { d: x.baseRow === 2 ? { row: 2, via: 'ID', flags: [] } : null } }));
    input.imports = [{ sourceId: 'd', col: 2 }, { sourceId: 'd', col: 3 }];
    input.keyFill = {};
    const s = (await snapshot((await buildPremerged(input)).bytes)).sheets[0];
    expect(s.values.E2).toBe('n:45462'); // date: + 1462 days
    expect(s.values.F2).toBe('n:44000'); // plain number untouched
  });
});

describe('style merge', () => {
  it('renumbers colliding custom number formats and deduplicates by format code', async () => {
    // both packages use numFmt 164; the source's 164 has a different code than the base's
    const JSZip = (await import('jszip')).default;
    const base = await buildXlsx([{ name: 'B', data: '<row r="1"><c r="A1" t="s"><v>0</v></c></row><row r="2"><c r="A2" s="2"><v>1</v></c></row>' }], { sst: ['ID'] });
    const srcRaw = await buildXlsx([{ name: 'S', data: '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" t="s"><v>2</v></c></row><row r="2"><c r="A2"><v>1</v></c><c r="B2" s="2"><v>5</v></c><c r="C2" s="4"><v>6</v></c></row>' }], { sst: ['ID', 'x', 'y'] });
    const z = await JSZip.loadAsync(srcRaw);
    const sx = (await z.file('xl/styles.xml')!.async('string'))
      .replace('<numFmts count="1"><numFmt numFmtId="164" formatCode="0.0&quot;pt&quot;"/></numFmts>', '<numFmts count="2"><numFmt numFmtId="164" formatCode="#,##0&quot;円&quot;"/><numFmt numFmtId="165" formatCode="0.0&quot;pt&quot;"/></numFmts>')
      .replace('<cellXfs count="4">', '<cellXfs count="5">')
      .replace('</cellXfs>', '<xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>');
    z.file('xl/styles.xml', sx);
    const src = await z.generateAsync({ type: 'uint8array' });
    const r = await buildPremerged({
      base: source('b', 'b.xlsx', { sheet: 'B', headerRow: 1, firstDataRow: 2, lastDataRow: 2, firstCol: 1, lastCol: 1 }, 'b', 'full', base),
      sources: [source('s', 's.xlsx', { sheet: 'S', headerRow: 1, firstDataRow: 2, lastDataRow: 2, firstCol: 1, lastCol: 3 }, 's', 'full', src)],
      rows: [{ baseRow: 2, matches: { s: { row: 2, via: 'ID', flags: [] } } }],
      imports: [{ sourceId: 's', col: 2 }, { sourceId: 's', col: 3 }],
      keyFill: {},
      options: { formulaPolicy: 'keep-if-possible', extendTable: true, headerConflict: 'suffix' },
    });
    const s = (await snapshot(r.bytes)).sheets[0];
    expect(s.styles.A2).toContain('nf=0.0"pt"'); // base cell unchanged
    expect(s.styles.B2).toContain('nf=#,##0"円"'); // source 164 -> new id in the base
    expect(s.styles.C2).toContain('nf=0.0"pt"'); // source 165 -> the base's existing 164 (same code)
    const styles = (await part(r.bytes, 'xl/styles.xml'))!;
    expect(styles).toContain('<numFmt numFmtId="165" formatCode="#,##0&quot;円&quot;"/>');
    expect((styles.match(/formatCode="0\.0&quot;pt&quot;"/g) ?? []).length).toBe(1);
  });
});
