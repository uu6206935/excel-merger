/**
 * Reader vs Excel on the R1 torture workbook (Excel-authored; charts, pivot, tables, CF/DV incl. x14, shared and
 * dynamic-array formulas, rich text, names with sheet scope …) and its Strict-OOXML twin.
 * Ground truth: excel-probes.json (recorded once via COM, com/record-probes.ps1) and SheetJS for formula text.
 */
import { beforeAll, describe, expect, it } from 'vitest';
import * as XLSX from 'xlsx';
import { readBook } from '../../src/core/xlsx/read';
import { cellText } from '../../src/core/xlsx/cellText';
import type { Cell, SheetGrid, SourceBook, XBorderSide, XStyle } from '../../src/core/types';
import { COM_HALIGN, COM_VALIGN, cellAt, comBorderStyle, comError, fx, probes, rc, type ProbeBorder } from '../fixtures/xlsx/helpers';

let book: SourceBook;
let strict: SourceBook;
const sheet = (b: SourceBook, name: string): SheetGrid => {
  const s = b.sheets.find((x) => x.name === name);
  if (!s) throw new Error('missing sheet ' + name);
  return s;
};

beforeAll(async () => {
  book = await readBook(fx('torture.xlsx'), 'torture.xlsx', { id: 't1' });
  strict = await readBook(fx('strict.xlsx'), 'strict.xlsx');
});

describe('torture.xlsx: workbook structure', () => {
  it('sheets, states, extents, format', () => {
    expect(book.id).toBe('t1');
    expect(book.format).toBe('xlsx');
    expect(book.fidelity).toBe('full');
    expect(book.date1904).toBe(false);
    expect(book.sheets.map((s) => [s.name, s.state])).toEqual([
      ['Data', 'visible'],
      ['集計 シート', 'visible'],
      ['Tbl', 'visible'],
      ['Pivot', 'visible'],
      ['Lists', 'hidden'],
    ]);
    const d = sheet(book, 'Data');
    expect([d.maxRow, d.maxCol]).toEqual([23, 20]);
    expect(book.warnings).toEqual([]);
  });

  it('BookFeatureCounts from package parts', () => {
    expect(book.features).toEqual({
      charts: 1, pivotTables: 1, vba: false, drawings: 1, slicers: 0, sparklineGroups: 0, externalLinks: 0, threadedComments: 1, formControls: 0, tables: 1, images: 1,
    });
  });

  it('defined names with sheet scope (localSheetId → sheet name)', () => {
    expect(book.names).toEqual([
      { name: '_xlnm._FilterDatabase', scope: 'Data', refersTo: 'Data!$A$3:$J$23', hidden: true },
      { name: 'AmountCol', scope: null, refersTo: 'Data!$E:$E' },
      { name: 'LocalRange', scope: 'Data', refersTo: 'Data!$A$4:$A$23' },
      { name: '_xlnm.Print_Area', scope: 'Data', refersTo: 'Data!$A$1:$J$23' },
      { name: '_xlnm.Print_Titles', scope: 'Data', refersTo: 'Data!$3:$3' },
      { name: 'TaxRate', scope: null, refersTo: "'集計 シート'!$B$2" },
    ]);
  });
});

describe('torture.xlsx: values and display text vs Excel', () => {
  it('every cell value equals Excel Value2 (all sheets)', () => {
    const pf = probes('torture.xlsx');
    let checked = 0;
    for (const ps of pf.sheets) {
      const s = sheet(book, ps.name);
      for (const pc of ps.cells) {
        const c = cellAt(s, pc.addr);
        const where = `${ps.name}!${pc.addr}`;
        if (pc.vt === 'null') {
          expect(c?.v ?? null, where).toBeNull();
          continue;
        }
        expect(c, where).toBeDefined();
        if (pc.vt === 'n') {
          expect(['n', 'd'], where).toContain(c!.t);
          expect(c!.v as number, where).toBeCloseTo(pc.v as number, 9);
        } else if (pc.vt === 's') {
          expect(c!.t, where).toBe('s');
          expect(c!.v, where).toBe(pc.v);
        } else if (pc.vt === 'b') {
          expect([c!.t, c!.v], where).toEqual(['b', pc.v]);
        } else {
          expect([c!.t, c!.v], where).toEqual(['e', comError(pc.v as number)]);
        }
        checked++;
      }
    }
    expect(checked).toBeGreaterThan(300);
  });

  it('dates are typed d (xf number format is a date format), numbers n', () => {
    const d = sheet(book, 'Data');
    expect(cellAt(d, 'D4')).toMatchObject({ t: 'd', v: 46000 });
    expect(cellAt(d, 'D5')).toMatchObject({ t: 'd', v: 46003 }); // m"月"d"日"\(aaa\)
    expect(cellAt(d, 'E4')).toMatchObject({ t: 'n', v: 12000 });
    expect(cellAt(d, 'F4')).toMatchObject({ t: 'n', v: 0 });
  });

  it('cellText equals Excel .Text for every cell (except #### column overflow)', () => {
    const pf = probes('torture.xlsx');
    const bad: string[] = [];
    let checked = 0;
    for (const ps of pf.sheets) {
      const s = sheet(book, ps.name);
      for (const pc of ps.cells) {
        if (pc.vt === 'null' || /^#+$/.test(pc.text)) continue;
        const got = cellText(book, cellAt(s, pc.addr));
        checked++;
        if (got !== pc.text) bad.push(`${ps.name}!${pc.addr}: got ${JSON.stringify(got)} excel ${JSON.stringify(pc.text)}`);
      }
    }
    expect(bad).toEqual([]);
    expect(checked).toBeGreaterThan(300);
  });

  it('rich text: runs with their own fonts, plain text in v', () => {
    const c = cellAt(sheet(book, 'Data'), 'J6')!;
    expect(c.v).toBe('重要: 要確認のメモ');
    expect(c.rich).toEqual([{ text: '重要:' }, { text: ' 要確認のメモ', style: { size: 11, font: '游ゴシック', color: '#000000' } }]);
  });

  it('line breaks: literal CR LF in the XML reads as LF (XML end-of-line handling), like Excel', () => {
    expect(cellAt(sheet(book, 'Data'), 'J7')!.v).toBe('改行を含む\nセル');
  });
});

describe('torture.xlsx: formulas', () => {
  it('formula text equals SheetJS for every formula cell (shared formulas expanded)', () => {
    const wb = XLSX.read(fx('torture.xlsx'), { type: 'array', cellFormula: true, cellNF: false, xlfn: true });
    let checked = 0;
    for (const s of book.sheets) {
      const ws = wb.Sheets[s.name];
      for (const key of Object.keys(ws)) {
        if (key.startsWith('!')) continue;
        const sj = ws[key] as XLSX.CellObject;
        if (!sj.f) continue;
        const mine = cellAt(s, key);
        expect(mine?.f, `${s.name}!${key}`).toBe(sj.f);
        checked++;
      }
      // and no extra formulas on our side
      s.rows.forEach((row, ri) =>
        row?.forEach((c, ci) => {
          if (c?.f) expect((ws[XLSX.utils.encode_cell({ r: ri, c: ci })] as XLSX.CellObject | undefined)?.f, `${s.name} r${ri + 1}c${ci + 1}`).toBe(c.f);
        }),
      );
    }
    expect(checked).toBeGreaterThan(100);
  });

  it('shared formulas: S4:S23 (range assignment) and T5:T23 (FillDown) expanded with translateFormula', () => {
    const d = sheet(book, 'Data');
    for (let r = 4; r <= 23; r++) expect(cellAt(d, `S${r}`)!.f).toBe(`E${r}*G${r}`);
    expect(cellAt(d, 'T4')!.f).toBe('$E4+H4');
    for (let r = 5; r <= 23; r++) expect(cellAt(d, `T${r}`)!.f).toBe(`$E${r}+H${r}`);
  });

  it('array (CSE) and dynamic-array masters get arr / dyn; spilled cells are plain values', () => {
    const d = sheet(book, 'Data');
    expect(cellAt(d, 'N4')).toMatchObject({ f: 'SUM(E4:E23*G4:G23)', arr: 'N4', t: 'n', v: 46133461 });
    expect(cellAt(d, 'N4')!.dyn).toBeUndefined();
    expect(cellAt(d, 'P4')).toMatchObject({ f: '_xlfn._xlws.SORT(_xlfn.UNIQUE(C4:C23))', arr: 'P4:P8', dyn: true, t: 's', v: 'SFC' });
    expect(cellAt(d, 'P5')).toEqual({ t: 's', v: '京大' });
    expect(d.features.arrays).toEqual([
      { r1: 4, c1: 14, r2: 4, c2: 14 },
      { r1: 4, c1: 16, r2: 8, c2: 16 },
    ]);
  });

  it('ANCHORARRAY (P4#), names, structured refs, cross-sheet refs kept in FILE form', () => {
    const s = sheet(book, '集計 シート');
    expect(cellAt(s, 'B8')).toMatchObject({ f: 'ROWS(_xlfn.ANCHORARRAY(Data!P4))', v: 5 });
    expect(cellAt(s, 'B6')).toMatchObject({ f: 'TaxRate*100', v: 10 });
    const d = sheet(book, 'Data');
    expect(cellAt(d, 'R6')!.f).toBe('SUM(SalesTbl[金額])');
    expect(cellAt(d, 'O4')!.f).toBe("'集計 シート'!$B$2*E4");
    expect(cellAt(d, 'Q4')).toMatchObject({ f: '1/0', t: 'e', v: '#DIV/0!' });
    expect(cellAt(d, 'Q5')).toMatchObject({ f: 'NA()', t: 'e', v: '#N/A' });
  });
});

describe('torture.xlsx: sheet features', () => {
  it('merges, freeze panes, autofilter, hyperlinks', () => {
    const d = sheet(book, 'Data');
    expect(d.features.merges).toEqual([{ r1: 1, c1: 1, r2: 1, c2: 6 }]);
    expect(d.features.freeze).toEqual({ cols: 2, rows: 3 });
    expect(d.features.autoFilter).toBe('A3:J23');
    expect(d.features.hyperlinks).toEqual([
      { ref: 'J4', target: 'https://example.com/ja' },
      { ref: 'J5', location: "'集計 シート'!A1" },
    ]);
  });

  it('conditional formatting and data validation incl. x14 extLst; type = <cfRule type> / <dataValidation type>', () => {
    const d = sheet(book, 'Data');
    expect(d.features.rules).toEqual([
      { sqref: 'E4:E23', type: 'cellIs', formulas: ['50000'], kind: 'cf' },
      { sqref: 'B4:B23', type: 'expression', formulas: ['$F4>0.5'], kind: 'cf' },
      { sqref: 'G4:G23', type: 'colorScale', formulas: [], kind: 'cf' },
      { sqref: 'H4:H23', type: 'dataBar', formulas: [], kind: 'cf' },
      { sqref: 'F4:F23', type: 'iconSet', formulas: [], kind: 'cf' },
      { sqref: 'G4:G23', type: 'whole', formulas: ['1', '100'], kind: 'dv' },
      { sqref: 'H4:H23', type: 'dataBar', formulas: [], kind: 'cf', ext: true }, // x14 counterpart of the main dataBar
      { sqref: 'C4:C23', type: 'list', formulas: ['Lists!$A$1:$A$5'], kind: 'dv', ext: true },
    ]);
  });

  it('tables via sheet rels (headerRowCount absent = header row)', () => {
    expect(sheet(book, 'Tbl').features.tables).toEqual([{ name: 'SalesTbl', ref: 'A1:D12', headerRow: true, totalsRow: true, columns: ['商品', '数量', '単価', '金額'] }]);
    expect(sheet(book, 'Data').features.tables).toEqual([]);
  });

  it('column spans → per-column props; rows; sheetFormatPr defaults', () => {
    const d = sheet(book, 'Data');
    expect(d.colProps[0]).toEqual({ width: 8.625, customWidth: true });
    expect(d.colProps[3]).toEqual({ width: 12.625, customWidth: true });
    expect(d.colProps[4]).toEqual({ width: 12.625, customWidth: true }); // <col min=4 max=5>
    expect(d.colProps[10]).toEqual({ width: 0, hidden: true, customWidth: true });
    // <col min=12 max=13 width="9" outlineLevel="1"/> has no customWidth: an auto width
    expect(d.colProps[11]).toEqual({ width: 9, outlineLevel: 1 });
    expect(d.colProps[12]).toEqual({ width: 9, outlineLevel: 1 });
    expect(d.colProps[13]).toBeUndefined();
    expect(d.rowProps[0]).toEqual({ height: 25.5 });
    expect(d.rowProps[2]).toEqual({ height: 30 });
    expect(d.rowProps[1]).toBeUndefined();
    expect(d.defaultRowHeight).toBe(18.75);
    expect(d.defaultColWidth).toBeUndefined();
  });
});

describe('torture.xlsx: display styles vs Excel (font, fill, alignment, borders, number format)', () => {
  const sides = ['left', 'top', 'bottom', 'right'] as const;
  const opposite = { left: 'right', right: 'left', top: 'bottom', bottom: 'top' } as const;
  const delta = { left: [0, -1], right: [0, 1], top: [-1, 0], bottom: [1, 0] } as const;

  function styleOf(b: SourceBook, s: SheetGrid, r: number, c: number): XStyle {
    const cell: Cell | undefined = s.rows[r - 1]?.[c - 1];
    return b.styles[cell?.s ?? 0] ?? {};
  }
  const sideKey = (x: XBorderSide | undefined) => (x ? `${x.style}/${(x.color ?? '#000000').slice(1)}` : 'none');
  const comKey = (x: ProbeBorder | null) => (x ? `${comBorderStyle(x)}/${x.color ?? '000000'}` : 'none');

  function checkFile(b: SourceBook, file: string, sheets: string[]): { checked: number; bad: string[] } {
    const bad: string[] = [];
    let checked = 0;
    for (const ps of probes(file).sheets.filter((x) => sheets.includes(x.name))) {
      const s = sheet(b, ps.name);
      for (const pc of ps.cells) {
        if (!pc.font) continue;
        const [r, c] = rc(pc.addr);
        const st = styleOf(b, s, r, c);
        const at = `${ps.name}!${pc.addr}`;
        const cell = s.rows[r - 1]?.[c - 1];
        checked++;
        const f = pc.font;
        if (st.font !== f.name) bad.push(`${at} font ${st.font} vs ${f.name}`);
        if (st.size !== f.size) bad.push(`${at} size ${st.size} vs ${f.size}`);
        if (!cell?.rich) {
          if (!!st.bold !== f.bold) bad.push(`${at} bold`);
          if (f.color !== null && (st.color ?? '#000000') !== '#' + f.color) bad.push(`${at} color ${st.color} vs ${f.color}`);
        }
        if (!!st.italic !== f.italic) bad.push(`${at} italic`);
        if (!!st.underline !== (f.underline !== -4142)) bad.push(`${at} underline`);
        if (!!st.strike !== f.strike) bad.push(`${at} strike`);
        const fill = pc.fill!.pattern === -4142 ? undefined : '#' + pc.fill!.color;
        if (st.fill !== fill) bad.push(`${at} fill ${st.fill} vs ${fill}`);
        if (st.hAlign !== COM_HALIGN[pc.hAlign!]) bad.push(`${at} hAlign ${st.hAlign} vs ${pc.hAlign}`);
        if (st.vAlign !== COM_VALIGN[pc.vAlign!]) bad.push(`${at} vAlign ${st.vAlign} vs ${pc.vAlign}`);
        if (!!st.wrap !== pc.wrap) bad.push(`${at} wrap`);
        if ((st.indent ?? 0) !== pc.indent) bad.push(`${at} indent`);
        // Excel's NumberFormat shows the code without redundant backslash escapes (#,##0\ "個" → #,##0 "個")
        const nf = pc.nf === 'G/標準' ? undefined : pc.nf;
        const unescaped = st.numFmt?.replace(/\\(.)/g, '$1');
        if (st.numFmt !== nf && unescaped !== nf) bad.push(`${at} numFmt ${st.numFmt} vs ${nf}`);
        // Borders: Excel reports the edge as drawn, which may come from the neighbouring cell's opposite side.
        for (const side of sides) {
          const mine = st.border?.[side];
          const [dr, dc] = delta[side];
          const nb = r + dr >= 1 && c + dc >= 1 ? styleOf(b, s, r + dr, c + dc).border?.[opposite[side]] : undefined;
          const excel = comKey(pc.borders![side]);
          const allowed = new Set([sideKey(mine), sideKey(nb)]);
          if (!mine && !nb) allowed.add('none');
          if (!allowed.has(excel)) bad.push(`${at} border ${side}: excel ${excel}, mine ${sideKey(mine)}, neighbour ${sideKey(nb)}`);
        }
      }
    }
    return { checked, bad };
  }

  it('torture Data / Tbl / 集計 シート: every probed cell', () => {
    const { checked, bad } = checkFile(book, 'torture.xlsx', ['Data', 'Tbl', '集計 シート']);
    expect(bad).toEqual([]);
    expect(checked).toBeGreaterThan(500);
  });

  it('title and header styles (theme colours with tint, medium red bottom border)', () => {
    const d = sheet(book, 'Data');
    expect(book.styles[cellAt(d, 'A1')!.s!]).toMatchObject({ bold: true, size: 16, fill: '#E0C285', hAlign: 'center', vAlign: 'middle' });
    expect(book.styles[cellAt(d, 'A3')!.s!]).toMatchObject({
      bold: true,
      color: '#FFFFFF',
      fill: '#997426',
      border: { bottom: { style: 'medium', color: '#FF0000' }, top: { style: 'thin', color: '#000000' } },
    });
    expect(book.styles[cellAt(d, 'J4')!.s!]).toMatchObject({ underline: true, color: '#467886', wrap: true, vAlign: 'top' });
    expect(book.styles[0]).toEqual({ size: 11, font: '游ゴシック', color: '#000000', vAlign: 'middle' });
  });
});

describe('strict.xlsx (Strict OOXML namespaces) reads exactly like the transitional original', () => {
  it('same sheets, cells, styles, names, features', () => {
    expect(strict.sheets.map((s) => [s.name, s.state, s.maxRow, s.maxCol])).toEqual(book.sheets.map((s) => [s.name, s.state, s.maxRow, s.maxCol]));
    for (const s of book.sheets) {
      const t = sheet(strict, s.name);
      for (let r = 1; r <= s.maxRow; r++) {
        for (let c = 1; c <= s.maxCol; c++) {
          const a = s.rows[r - 1]?.[c - 1];
          const b = t.rows[r - 1]?.[c - 1];
          const where = `${s.name} r${r}c${c}`;
          expect(b?.v ?? null, where).toEqual(a?.v ?? null);
          expect(b?.t, where).toBe(a?.t);
          expect(b?.f, where).toBe(a?.f);
          expect(b?.arr, where).toBe(a?.arr);
          expect(strict.styles[b?.s ?? 0], where).toEqual(book.styles[a?.s ?? 0]);
        }
      }
      expect(t.features.merges).toEqual(s.features.merges);
      expect(t.features.tables).toEqual(s.features.tables);
      expect(t.features.freeze).toEqual(s.features.freeze);
      expect([...t.features.rules].sort((x, y) => x.sqref.localeCompare(y.sqref) || x.kind.localeCompare(y.kind))).toEqual(
        [...s.features.rules].sort((x, y) => x.sqref.localeCompare(y.sqref) || x.kind.localeCompare(y.kind)),
      );
    }
    expect(strict.names).toEqual(book.names);
    expect(strict.features).toEqual(book.features);
  });
});
