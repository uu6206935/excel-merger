/**
 * Single-sheet output (PremergeOptions.dropOtherSheets) feature by feature, on hand-made packages: which formulas / names / validations /
 * conditional formats / links / charts / sparklines / pivot tables / slicers / workbook parts go, what stays, and what the report says.
 * The packages are built with tests/xlsx/dropsheets-helpers.ts pkgOf; the base sheet is always `主`, the sheets to drop are `他` and `他 2`.
 */
import { describe, expect, it } from 'vitest';
import { buildPremerged, readBook } from '../../src/core/xlsx';
import { NEW_METADATA_XML } from '../../src/core/xlsx/patch/metadata';
import type { TableRegion } from '../../src/core/types';
import {
  CT, bareBase, baseOptions, files, fml, lintSingleSheet, lintText, num, pkgOf, premergeOne, row, str,
  type PkgDef, type RelDef, type SheetDef,
} from './dropsheets-helpers';

const MAIN = '主';
const EXT = '他';
const EXT2 = '他 2';
const region = (lastDataRow: number, lastCol = 2, sheet = MAIN): TableRegion => ({ sheet, headerRow: 1, firstDataRow: 2, lastDataRow, firstCol: 1, lastCol });

/** the sheet to drop: A1:A5 = a..e, B1:E1 = 1..4 */
const other = (): SheetDef => ({
  name: EXT,
  data: [row(1, `${str('A1', 'a')}${num('B1', 1)}${num('C1', 2)}${num('D1', 3)}${num('E1', 4)}`), ...['b', 'c', 'd', 'e'].map((t, i) => row(i + 2, str(`A${i + 2}`, t)))].join(''),
});
const other2 = (): SheetDef => ({ name: EXT2, data: row(1, `${num('A1', 10)}${num('B1', 20)}`) + row(2, `${num('A2', 30)}${num('B2', 40)}`) });

/** a base sheet `主`: header ID / 値, data rows 2..9 */
const mainRows = (extra: (r: number) => string = () => ''): string =>
  row(1, `${str('A1', 'ID')}${str('B1', '値')}`) + Array.from({ length: 8 }, (_x, i) => row(i + 2, `${num(`A${i + 2}`, i + 1)}${num(`B${i + 2}`, (i + 1) * 10)}${extra(i + 2)}`)).join('');

const premerge = async (def: PkgDef, lastDataRow = 9, lastCol = 2, extra: Parameters<typeof premergeOne>[1] = { dropOtherSheets: true }) => {
  const bytes = await pkgOf(def);
  const res = await premergeOne(bareBase(bytes, region(lastDataRow, lastCol)), extra);
  const out = await readBook(res.bytes, 'out.xlsx', { id: 'out' });
  const f = await files(res.bytes);
  return { res, out, f, sheet: out.sheets[0], xml: f.get('xl/worksheets/sheet1.xml') ?? '', inBytes: bytes };
};

const cell = (sheet: { rows: Array<Array<{ f?: string; v: unknown; t: string; arr?: string; dyn?: boolean; s?: number } | undefined> | undefined> }, ref: string) => {
  const m = /^([A-Z]+)(\d+)$/.exec(ref)!;
  return sheet.rows[Number(m[2]) - 1]?.[m[1].split('').reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1];
};
const clean = async (res: { bytes: Uint8Array }, removed: string[] = [EXT, EXT2]) => expect(lintText(await lintSingleSheet(res.bytes, removed))).toBe('');

describe('formulas that read a removed sheet become their cached values', () => {
  const names = '<definedName name="Nx_1">他!$A$1</definedName><definedName name="Nx_2">Nx_1*2</definedName><definedName name="Nx_3">主!$B$2</definedName>';
  const def = (): PkgDef => ({
    sheets: [
      {
        name: MAIN,
        data: mainRows((r) => ({
          2: fml('C2', '他!A1', 5, 'n', ' s="2"'),
          3: fml('C3', "SUM('他 2'!A1:B2)", 100),
          4: fml('C4', 'SUM(他!A:A)', 7),
          5: fml('C5', 'SUM(主:他!B2)', 3),
          6: fml('C6', 'B2+1', 11),
          7: fml('C7', '主!B2*2', 20),
          8: fml('C8', 'Nx_2+1', 3),
          9: fml('C9', 'Nx_3*2', 20),
        } as Record<number, string>)[r] ?? ''),
      },
      other(),
      other2(),
    ],
    names,
  });

  it('direct, quoted, whole-column, 3D, through a chain of names; the rest stays live', async () => {
    const { res, sheet } = await premerge(def());
    await clean(res);
    expect(res.report.formulasDetached).toBe(5);
    expect(res.report.items.find((i) => i.code === 'detachedFormulas')).toMatchObject({ severity: 'info', message: '他のシートを参照していた数式 5 個を値にしました', rows: [2, 3, 4, 5, 8] });
    for (const ref of ['C2', 'C3', 'C4', 'C5', 'C8']) expect(cell(sheet, ref)?.f, ref).toBeUndefined();
    expect([2, 3, 4, 5, 8].map((r) => cell(sheet, `C${r}`)?.v)).toEqual([5, 100, 7, 3, 3]);
    expect(cell(sheet, 'C6')?.f).toBe('B2+1');
    expect(cell(sheet, 'C7')?.f).toBe('主!B2*2');
    expect(cell(sheet, 'C9')?.f).toBe('Nx_3*2');
  });

  it('keeps the style of the cell and the cached value types: string, boolean, error, no value at all', async () => {
    const d: PkgDef = {
      sheets: [
        {
          name: MAIN,
          data: mainRows((r) => ({
            2: fml('C2', '他!A1', 'hello', 's', ' s="1"'),
            3: fml('C3', '他!A2', true, 'b'),
            4: fml('C4', '他!A3', '#N/A', 'e'),
            5: fml('C5', '他!A4', null, 'n', ' s="1"'),
            6: fml('C6', '他!A5', 'a&b <c>', 's'),
          } as Record<number, string>)[r] ?? ''),
        },
        other(),
      ],
      sst: [],
    };
    const { sheet, xml, res } = await premerge(d);
    await clean(res, [EXT]);
    expect(cell(sheet, 'C2')).toMatchObject({ t: 's', v: 'hello', s: 1 });
    expect(cell(sheet, 'C3')).toMatchObject({ t: 'b', v: true });
    expect(cell(sheet, 'C4')).toMatchObject({ t: 'e', v: '#N/A' });
    expect(cell(sheet, 'C5')?.f).toBeUndefined();
    expect(cell(sheet, 'C5')?.v ?? null).toBeNull();
    expect(cell(sheet, 'C5')?.s).toBe(1);
    expect(cell(sheet, 'C6')).toMatchObject({ t: 's', v: 'a&b <c>' });
    expect(xml).not.toContain('t="str"');
    expect(res.report.formulasDetached).toBe(5);
  });

  it('INDIRECT with a literal that names a removed sheet; INDIRECT of the merged sheet and external references stay', async () => {
    const d: PkgDef = {
      sheets: [
        {
          name: MAIN,
          data: mainRows((r) => ({
            2: fml('C2', 'INDIRECT("他!A1")', 'a', 's'),
            3: fml('C3', 'INDIRECT("B2")', 10),
            4: fml('C4', "INDIRECT(\"'他 2'!A1\")", 10),
            5: fml('C5', '[1]外部!A1', 1),
            6: fml('C6', 'INDIRECT("Nx_1")', 5),
          } as Record<number, string>)[r] ?? ''),
        },
        other(),
        other2(),
      ],
      names: '<definedName name="Nx_1">他!$A$1</definedName>',
    };
    const { sheet, res } = await premerge(d);
    expect(res.report.formulasDetached).toBe(3);
    expect(cell(sheet, 'C3')?.f).toBe('INDIRECT("B2")');
    expect(cell(sheet, 'C5')?.f).toBe('[1]外部!A1');
    for (const ref of ['C2', 'C4', 'C6']) expect(cell(sheet, ref)?.f, ref).toBeUndefined();
  });

  it('shared formulas: master and children all read the removed sheet, all become values', async () => {
    const d: PkgDef = {
      sheets: [
        {
          name: MAIN,
          data: row(1, `${str('A1', 'ID')}${str('B1', '値')}`)
            + row(2, `${num('A2', 1)}<c r="B2"><f t="shared" ref="B2:B4" si="0">他!B1*2</f><v>2</v></c>`)
            + row(3, `${num('A3', 2)}<c r="B3"><f t="shared" si="0"/><v>4</v></c>`)
            + row(4, `${num('A4', 3)}<c r="B4"><f t="shared" si="0"/><v>6</v></c>`)
            + row(5, `${num('A5', 4)}<c r="B5"><f t="shared" ref="B5:B6" si="1">A5*2</f><v>8</v></c>`)
            + row(6, `${num('A6', 5)}<c r="B6"><f t="shared" si="1"/><v>10</v></c>`),
        },
        other(),
      ],
    };
    const { sheet, res } = await premerge(d, 6);
    expect(res.report.formulasDetached).toBe(3);
    expect([2, 3, 4].map((r) => [cell(sheet, `B${r}`)?.f, cell(sheet, `B${r}`)?.v])).toEqual([[undefined, 2], [undefined, 4], [undefined, 6]]);
    expect([5, 6].map((r) => cell(sheet, `B${r}`)?.f)).toEqual(['A5*2', 'A6*2']);
  });

  it('structured references to a table of a removed sheet (also bare) are detached, to a table of the merged sheet stay', async () => {
    const tbl = (id: number, name: string, ref: string, cols: string[]) =>
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<table ${NS_MAIN} id="${id}" name="${name}" displayName="${name}" ref="${ref}" totalsRowShown="0"><autoFilter ref="${ref}"/><tableColumns count="${cols.length}">${cols.map((c, i) => `<tableColumn id="${i + 1}" name="${c}"/>`).join('')}</tableColumns><tableStyleInfo name="TableStyleMedium2" showRowStripes="1"/></table>`;
    const d: PkgDef = {
      sheets: [
        {
          name: MAIN,
          data: row(1, `${str('A1', 'ID')}${str('B1', '値')}`)
            + row(2, `${num('A2', 1)}${num('B2', 10)}${fml('C2', 'SUM(Rem[Qty])', 6)}${fml('C2'.replace('C', 'D'), 'SUM(Loc[値])', 10)}`)
            + row(3, `${num('A3', 2)}${num('B3', 20)}${fml('C3', 'ROWS(Rem)', 3)}${fml('D3', 'Loc[[#This Row],[値]]*2', 40)}`),
          tail: '<tableParts count="1"><tablePart r:id="rId1"/></tableParts>',
          rels: [{ id: 'rId1', type: 'table', target: '../tables/table1.xml' }],
        },
        {
          name: EXT,
          data: row(1, `${str('A1', 'Qty')}`) + row(2, num('A2', 1)) + row(3, num('A3', 2)) + row(4, num('A4', 3)),
          tail: '<tableParts count="1"><tablePart r:id="rId1"/></tableParts>',
          rels: [{ id: 'rId1', type: 'table', target: '../tables/table2.xml' }],
        },
      ],
      parts: [
        { path: 'xl/tables/table1.xml', xml: tbl(1, 'Loc', 'A1:B3', ['ID', '値']), ct: CT.table },
        { path: 'xl/tables/table2.xml', xml: tbl(2, 'Rem', 'A1:A4', ['Qty']), ct: CT.table },
      ],
    };
    const { sheet, res, f } = await premerge(d, 3);
    await clean(res, [EXT]);
    expect(res.report.formulasDetached).toBe(2);
    expect(cell(sheet, 'C2')?.f).toBeUndefined();
    expect(cell(sheet, 'C3')?.f).toBeUndefined();
    expect(cell(sheet, 'D2')?.f).toBe('SUM(Loc[値])');
    expect(cell(sheet, 'D3')?.f).toBe('Loc[[#This Row],[値]]*2');
    expect([...f.keys()].filter((k) => k.startsWith('xl/tables/'))).toEqual(['xl/tables/table1.xml']);
  });

  it('array formulas: a CSE array that reads a removed sheet becomes values as a whole; a spill reference to a detached spill follows', async () => {
    const d: PkgDef = {
      sheets: [
        {
          name: MAIN,
          data: row(1, `${str('A1', 'ID')}${str('B1', '値')}`)
            + row(2, `${num('A2', 1)}${num('B2', 10)}<c r="C2"><f t="array" ref="C2:C4">他!B1:D1*2</f><v>2</v></c><c r="D2" cm="1"><f t="array" ref="D2:D4">_xlfn._xlws.SORT(他!B1:D1)</f><v>1</v></c>${fml('E2', '_xlfn.ANCHORARRAY(D2)', 1)}<c r="F2" cm="1"><f t="array" ref="F2:F3">_xlfn._xlws.SORT(B2:B3)</f><v>10</v></c>`)
            + row(3, `${num('A3', 2)}${num('B3', 20)}${num('C3', 4)}${num('D3', 2)}<c r="F3"><v>20</v></c>`)
            + row(4, `${num('A4', 3)}${num('B4', 30)}${num('C4', 6)}${num('D4', 3)}`),
        },
        other(),
      ],
      parts: [{ path: 'xl/metadata.xml', xml: NEW_METADATA_XML, ct: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml' }],
      wbRels: [{ id: 'rId20', type: 'sheetMetadata', target: 'metadata.xml' }],
    };
    const { sheet, res, f, xml } = await premerge(d, 4);
    await clean(res, [EXT]);
    expect(res.report.formulasDetached).toBe(3); // C2 (CSE), D2 (spill), E2 (D2#)
    expect(sheet.features.arrays.length).toBe(1); // the spill F2:F3 only
    expect([cell(sheet, 'C2'), cell(sheet, 'C3'), cell(sheet, 'C4')].map((c) => [c?.f, c?.v])).toEqual([[undefined, 2], [undefined, 4], [undefined, 6]]);
    expect([2, 3, 4].map((r) => cell(sheet, `D${r}`)?.f ?? null)).toEqual([null, null, null]);
    expect(cell(sheet, 'E2')?.f).toBeUndefined();
    expect(cell(sheet, 'F2')).toMatchObject({ dyn: true, arr: 'F2:F3' });
    expect(xml).not.toContain('cm="1"><f t="array" ref="D2');
    expect(f.has('xl/metadata.xml')).toBe(true); // F2 still is a dynamic array
  });

  it('the dynamic-array metadata part goes when no cell uses it any more', async () => {
    const d: PkgDef = {
      sheets: [
        { name: MAIN, data: row(1, `${str('A1', 'ID')}${str('B1', '値')}`) + row(2, `${num('A2', 1)}${num('B2', 10)}<c r="C2" cm="1"><f t="array" ref="C2:C3">_xlfn._xlws.SORT(他!B1:C1)</f><v>1</v></c>`) + row(3, `${num('A3', 2)}${num('B3', 20)}<c r="C3"><v>2</v></c>`) },
        other(),
      ],
      parts: [{ path: 'xl/metadata.xml', xml: NEW_METADATA_XML, ct: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml' }],
      wbRels: [{ id: 'rId20', type: 'sheetMetadata', target: 'metadata.xml' }],
    };
    const { f, res } = await premerge(d, 3);
    await clean(res, [EXT]);
    expect(f.has('xl/metadata.xml')).toBe(false);
    expect(f.get('xl/_rels/workbook.xml.rels')).not.toContain('sheetMetadata');
  });

  it('a name that is defined on the merged sheet and shadows a bad global name decides', async () => {
    const d: PkgDef = {
      sheets: [
        { name: MAIN, data: mainRows((r) => (r === 2 ? fml('C2', 'Nx_1+1', 2) : r === 3 ? fml('C3', 'Nx_2+1', 3) : '')) },
        other(),
      ],
      // Nx_1 is global and bad, but the sheet-scoped Nx_1 points at 主; Nx_2 is global and bad
      names: '<definedName name="Nx_1">他!$A$1</definedName><definedName name="Nx_1" localSheetId="0">主!$B$2</definedName><definedName name="Nx_2">他!$A$2</definedName>',
    };
    const { sheet, res, f } = await premerge(d);
    expect(res.report.formulasDetached).toBe(1);
    expect(cell(sheet, 'C2')?.f).toBe('Nx_1+1');
    expect(cell(sheet, 'C3')?.f).toBeUndefined();
    expect(f.get('xl/workbook.xml')).toContain('<definedName name="Nx_1" localSheetId="0">主!$B$2</definedName>');
    expect(f.get('xl/workbook.xml')).not.toContain('他');
  });

  it('a formula that the tokenizer cannot read but that names a removed sheet is detached too', async () => {
    const d: PkgDef = {
      sheets: [{ name: MAIN, data: mainRows((r) => (r === 2 ? fml('C2', '他!A1+"x', 5) : '')) }, other()],
    };
    const { res } = await premerge(d);
    expect(res.report.formulasDetached).toBe(1);
  });
});

const NS_MAIN = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';

describe('defined names', () => {
  it('names scoped to a removed sheet or reading one go; names of the merged sheet are re-indexed to 0, built-in names keep their scope', async () => {
    // order: 他 (0), 主 (1), 他 2 (2): the merged sheet is the middle one
    const d: PkgDef = {
      sheets: [other(), { name: MAIN, data: mainRows() }, other2()],
      names: [
        '<definedName name="Chain1">他!$A$1</definedName>',
        '<definedName name="Chain2">Chain1*2</definedName>',
        '<definedName name="Good">主!$B$2:$B$5</definedName>',
        '<definedName name="OnRemoved" localSheetId="0">他!$A$2</definedName>',
        '<definedName name="OnRemoved2" localSheetId="2">主!$A$2</definedName>',
        '<definedName name="Mine" localSheetId="1">主!$B$3</definedName>',
        '<definedName name="MineBad" localSheetId="1">Chain2+1</definedName>',
        '<definedName name="_xlnm.Print_Area" localSheetId="0">他!$A$1:$A$5</definedName>',
        '<definedName name="_xlnm.Print_Area" localSheetId="1">主!$A$1:$B$9</definedName>',
        '<definedName name="_xlnm.Print_Titles" localSheetId="1">主!$1:$1</definedName>',
        '<definedName name="_xlnm._FilterDatabase" localSheetId="1" hidden="1">主!$A$1:$B$9</definedName>',
        '<definedName name="_xlnm._FilterDatabase" localSheetId="2" hidden="1">\'他 2\'!$A$1:$B$2</definedName>',
        '<definedName name="WholeBook">SUM(\'他 2\'!A1:B2,主!B2)</definedName>',
        '<definedName name="Const">42</definedName>',
        '<definedName name="Err">#REF!</definedName>',
        '<definedName name="Ext">[1]Sheet1!$A$1</definedName>',
      ].join(''),
      bookViews: '<bookViews><workbookView activeTab="2" firstSheet="1"/></bookViews>',
    };
    const bytes = await pkgOf(d);
    const res = await premergeOne(bareBase(bytes, region(9)));
    await clean(res);
    const wb = (await files(res.bytes)).get('xl/workbook.xml')!;
    const names = [...wb.matchAll(/<definedName\b([^>]*)>([^<]*)</g)].map((m) => `${/ name="([^"]*)"/.exec(m[1])![1]}|${/localSheetId="(\d+)"/.exec(m[1])?.[1] ?? ''}|${m[2]}`).sort();
    expect(names).toEqual([
      'Const||42',
      'Err||#REF!',
      'Ext||[1]Sheet1!$A$1',
      'Good||主!$B$2:$B$5',
      'Mine|0|主!$B$3',
      '_xlnm.Print_Area|0|主!$A$1:$B$9',
      '_xlnm.Print_Titles|0|主!$1:$1',
      '_xlnm._FilterDatabase|0|主!$A$1:$B$9',
    ].sort());
    expect(wb).toContain('<bookViews><workbookView/></bookViews>');
    expect(wb).toContain('<sheets><sheet name="主" sheetId="2" r:id="rId2"/></sheets>');
    expect(res.report.droppedSheets).toEqual([EXT, EXT2]);
  });

  it('a name that is removed with no names left leaves no empty <definedNames>', async () => {
    const d: PkgDef = { sheets: [{ name: MAIN, data: mainRows() }, other()], names: '<definedName name="X">他!$A$1</definedName>' };
    const { f } = await premerge(d);
    expect(f.get('xl/workbook.xml')).not.toContain('definedNames');
  });
});

describe('data validations', () => {
  const dv = (type: string, sqref: string, f1: string, extra = '') => `<dataValidation type="${type}" allowBlank="1" showErrorMessage="1" sqref="${sqref}"${extra}><formula1>${f1}</formula1></dataValidation>`;
  const x14dv = (type: string, sqref: string, f1: string) => `<x14:dataValidation type="${type}" allowBlank="1" showInputMessage="1" showErrorMessage="1"><x14:formula1><xm:f>${f1}</xm:f></x14:formula1><xm:sqref>${sqref}</xm:sqref></x14:dataValidation>`;
  const ext = (inner: string) => `<extLst><ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:dataValidations count="1" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">${inner}</x14:dataValidations></ext></extLst>`;
  const rules = (sheet: { features: { rules: Array<{ kind: string; type?: string; sqref: string; formulas: string[]; ext?: boolean }> } }) =>
    sheet.features.rules.filter((r) => r.kind === 'dv').map((r) => `${r.sqref}|${r.type}|${r.ext ? 'x14' : ''}|${r.formulas.join(',')}`).sort();

  it('list from a column / row / name / formatted cells becomes a literal list; x14 lists move into the classic container', async () => {
    const o: SheetDef = {
      name: EXT,
      data: row(1, `${str('A1', 'a')}${num('B1', 1)}${num('C1', 2)}${num('D1', 3)}<c r="F1" s="2"><v>1.5</v></c><c r="G1" s="3"><v>44927</v></c>`) + ['b', 'c', 'd', 'e'].map((t, i) => row(i + 2, str(`A${i + 2}`, t))).join('')
        + row(8, `<c r="A8"><f>1+1</f><v>2</v></c>`),
    };
    const d: PkgDef = {
      sheets: [
        {
          name: MAIN,
          data: mainRows(),
          tail: `<dataValidations count="5">${dv('list', 'B2:B9', '他!$A$1:$A$5')}${dv('list', 'C2:C9', 'LIST')}${dv('list', 'D2:D9', '他!$B$1:$D$1')}${dv('list', 'E2:E9', '他!$F$1:$G$1')}${dv('list', 'A2:A9', '"x,y,z"')}</dataValidations>`
            + ext(x14dv('list', 'F2:F9', '他!$A$1:$A$4')),
        },
        o,
      ],
      names: '<definedName name="LIST">他!$A$1:$A$5</definedName>',
    };
    const { sheet, res, xml } = await premerge(d);
    await clean(res, [EXT]);
    expect(res.report.items.find((i) => i.code === 'dvToList')!.message).toBe('入力規則 5 個を固定のリストにしました');
    expect(rules(sheet)).toEqual([
      'A2:A9|list||"x,y,z"',
      'B2:B9|list||"a,b,c,d,e"',
      'C2:C9|list||"a,b,c,d,e"',
      'D2:D9|list||"1,2,3"',
      'E2:E9|list||"1.5pt,2023/1/1"',
      'F2:F9|list||"a,b,c,d"',
    ].sort());
    expect(xml).toContain('<dataValidations count="6">');
    expect(xml).not.toContain('x14:');
    expect(xml).not.toContain('<extLst');
    // the attributes of the x14 validation moved with it
    expect(xml).toMatch(/<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="F2:F9">/);
  });

  it('a source that cannot be a literal list drops the validation: commas, quotes, more than 255 characters, 2D range, calculated source, other types', async () => {
    const long = Array.from({ length: 30 }, (_x, i) => `item-number-${i}`);
    const o: SheetDef = {
      name: EXT,
      data: row(1, `${str('A1', 'a,b')}${str('B1', 'x')}${str('C1', 'y')}`) + row(2, `${str('A2', 'c')}${str('B2', 'z')}${str('C2', 'w')}`)
        + long.map((t, i) => row(i + 10, str(`E${i + 10}`, t))).join('') + row(50, `${str('G50', 'say "hi"')}${str('G51', 'ok')}`),
    };
    const d: PkgDef = {
      sheets: [
        {
          name: MAIN,
          data: mainRows(),
          tail: `<dataValidations count="9">${dv('list', 'A2:A9', '他!$A$1:$A$2')}${dv('list', 'B2:B9', '他!$E$10:$E$39')}${dv('list', 'C2:C9', '他!$B$1:$C$2')}${dv('list', 'D2:D9', 'OFFSET(他!$A$1,0,0,2,1)')}`
            + `${dv('whole', 'E2:E9', '他!$B$1').replace('</dataValidation>', '<formula2>100</formula2></dataValidation>')}${dv('custom', 'F2:F9', '他!$B$1>0')}${dv('list', 'G2:G9', '他!$G$50:$G$51')}${dv('list', 'H2:H9', '他!$Z$1:$Z$3')}${dv('list', 'I2:I9', '主!$B$2:$B$5')}</dataValidations>`,
        },
        o,
      ],
    };
    const { sheet, res } = await premerge(d);
    await clean(res, [EXT]);
    expect(res.report.items.find((i) => i.code === 'dvDropped')!.message).toBe('他のシートを参照していた入力規則 8 個を外しました');
    expect(res.report.items.some((i) => i.code === 'dvToList')).toBe(false);
    expect(rules(sheet)).toEqual(['I2:I9|list||主!$B$2:$B$5']);
  });

  it('x14-only validations: a literal list creates the classic container where the schema wants it, others are removed with the empty extLst', async () => {
    const o: SheetDef = { name: EXT, data: row(1, `${str('A1', 'p')}`) + row(2, str('A2', 'q')) };
    const d: PkgDef = {
      sheets: [
        {
          name: MAIN,
          data: mainRows(),
          tail: `<hyperlinks><hyperlink ref="A2" location="主!B2"/></hyperlinks><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>${ext(x14dv('list', 'C2:C9', '他!$A$1:$A$2') + x14dv('whole', 'D2:D9', '他!$A$1'))}`,
        },
        o,
      ],
    };
    const { sheet, xml, res } = await premerge(d);
    await clean(res, [EXT]);
    expect(rules(sheet)).toEqual(['C2:C9|list||"p,q"']);
    // CT_Worksheet order: dataValidations before hyperlinks
    expect(xml.indexOf('<dataValidations')).toBeLessThan(xml.indexOf('<hyperlinks'));
    expect(xml).not.toContain('<extLst');
    expect(res.report.items.find((i) => i.code === 'dvDropped')!.message).toContain('1 個');
  });

  it('a list that reads a table column of a removed sheet through a name', async () => {
    const o: SheetDef = { name: EXT, data: row(1, str('A1', 'Cat')) + row(2, str('A2', 'x')) + row(3, str('A3', 'y')) };
    const d: PkgDef = {
      sheets: [{ name: MAIN, data: mainRows(), tail: `<dataValidations count="1">${dv('list', 'B2:B9', 'Cats')}</dataValidations>` }, o],
      names: '<definedName name="Cats">他!$A$2:$A$3</definedName>',
    };
    const { sheet } = await premerge(d);
    expect(rules(sheet)).toEqual(['B2:B9|list||"x,y"']);
  });
});

describe('conditional formatting', () => {
  const X14 = 'xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"';
  const GUID1 = '{11111111-1111-1111-1111-111111111111}';
  const GUID2 = '{22222222-2222-2222-2222-222222222222}';
  const bar = (guid: string) => `<cfRule type="dataBar" priority="PRI"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar><extLst><ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}" ${X14}><x14:id>${guid}</x14:id></ext></extLst></cfRule>`;

  it('rules that read a removed sheet (formula, cfvo, name, x14) are dropped, their linked x14 / classic partner with them; the rest keeps its priority order', async () => {
    const tail = [
      '<conditionalFormatting sqref="B2:B9"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>他!$A$1</formula></cfRule></conditionalFormatting>',
      '<conditionalFormatting sqref="C2:C9"><cfRule type="expression" dxfId="0" priority="2"><formula>$B2&gt;1</formula></cfRule><cfRule type="expression" dxfId="0" priority="3"><formula>Nx_1&gt;1</formula></cfRule></conditionalFormatting>',
      '<conditionalFormatting sqref="D2:D9"><cfRule type="colorScale" priority="4"><colorScale><cfvo type="min"/><cfvo type="formula" val="他!$A$1"/><color rgb="FFF8696B"/><color rgb="FF63BE7B"/></colorScale></cfRule></conditionalFormatting>',
      `<conditionalFormatting sqref="E2:E9">${bar(GUID1).replace('PRI', '5')}</conditionalFormatting>`,
      `<conditionalFormatting sqref="G2:G9">${bar(GUID2).replace('PRI', '6')}</conditionalFormatting>`,
      '<conditionalFormatting sqref="H2:H9"><cfRule type="top10" dxfId="0" priority="7" rank="3"/></conditionalFormatting>',
      `<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" ${X14}><x14:conditionalFormattings>`
        + `<x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x14:cfRule type="dataBar" id="${GUID1}"><x14:dataBar minLength="0" maxLength="100"><x14:cfvo type="num"><xm:f>他!$A$1</xm:f></x14:cfvo><x14:cfvo type="max"/></x14:dataBar></x14:cfRule><xm:sqref>E2:E9</xm:sqref></x14:conditionalFormatting>`
        + `<x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x14:cfRule type="dataBar" id="${GUID2}"><x14:dataBar minLength="0" maxLength="100"><x14:cfvo type="min"/><x14:cfvo type="max"/></x14:dataBar></x14:cfRule><xm:sqref>G2:G9</xm:sqref></x14:conditionalFormatting>`
        + `<x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x14:cfRule type="iconSet" priority="8" id="{33333333-3333-3333-3333-333333333333}"><x14:iconSet iconSet="3Arrows"><x14:cfvo type="percent"><xm:f>0</xm:f></x14:cfvo><x14:cfvo type="num"><xm:f>他!$A$2</xm:f></x14:cfvo><x14:cfvo type="percent"><xm:f>67</xm:f></x14:cfvo></x14:iconSet></x14:cfRule><xm:sqref>I2:I9</xm:sqref></x14:conditionalFormatting>`
        + '</x14:conditionalFormattings></ext></extLst>',
    ].join('');
    const d: PkgDef = { sheets: [{ name: MAIN, data: mainRows(), tail }, other()], names: '<definedName name="Nx_1">他!$A$1</definedName>' };
    const { sheet, xml, res } = await premerge(d);
    await clean(res, [EXT]);
    // classic: cellIs, Nx_1 rule, colorScale, dataBar(GUID1) ; x14: iconSet  (the x14 dataBar of GUID1 is the partner of the classic one)
    expect(res.report.items.find((i) => i.code === 'cfDropped')!.message).toBe('条件付き書式 5 個を外しました');
    expect(sheet.features.rules.filter((r) => r.kind === 'cf').map((r) => `${r.sqref}|${r.type}|${r.ext ? 'x14' : ''}|${r.formulas.join(',')}`).sort()).toEqual([
      'C2:C9|expression||$B2>1', 'G2:G9|dataBar||', 'G2:G9|dataBar|x14|', 'H2:H9|top10||',
    ].sort());
    expect(xml).not.toContain('他');
    expect(xml).not.toContain(GUID1);
    expect([...xml.matchAll(/priority="(\d+)"/g)].map((m) => Number(m[1]))).toEqual([1, 2, 3]);
  });

  it('every rule gone: no empty containers, no empty extLst', async () => {
    const tail = `<conditionalFormatting sqref="B2:B9"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>他!$A$1</formula></cfRule></conditionalFormatting><extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" ${X14}><x14:conditionalFormattings><x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x14:cfRule type="dataBar" id="${GUID1}"><x14:dataBar minLength="0" maxLength="100"><x14:cfvo type="num"><xm:f>他!$A$1</xm:f></x14:cfvo><x14:cfvo type="max"/></x14:dataBar></x14:cfRule><xm:sqref>E2:E9</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst>`;
    const { xml, res } = await premerge({ sheets: [{ name: MAIN, data: mainRows(), tail }, other()] });
    expect(res.report.items.find((i) => i.code === 'cfDropped')!.message).toBe('条件付き書式 2 個を外しました');
    expect(xml).not.toMatch(/conditionalFormatting|<extLst|<ext /);
  });
});

describe('hyperlinks, sparklines, links of shapes', () => {
  it('a link to a removed sheet (or to a name that points at one) becomes a plain cell; links to the merged sheet and external links stay', async () => {
    const d: PkgDef = {
      sheets: [
        {
          name: MAIN,
          data: mainRows(),
          tail: '<hyperlinks><hyperlink ref="A2" location="他!A1" display="x"/><hyperlink ref="A3" location="主!B2"/><hyperlink ref="A4" r:id="rId1"/><hyperlink ref="A5" location="Nx_1"/><hyperlink ref="A6" location="\'他 2\'!A1:B2"/></hyperlinks>',
          rels: [{ id: 'rId1', type: 'hyperlink', target: 'https://example.com/', external: true }],
        },
        other(),
        other2(),
      ],
      names: '<definedName name="Nx_1">他!$A$1</definedName>',
    };
    const { sheet, res } = await premerge(d);
    await clean(res);
    expect(res.report.items.find((i) => i.code === 'linksDropped')!.message).toBe('他のシートへのリンク 3 個を外しました');
    expect(sheet.features.hyperlinks).toEqual([{ ref: 'A3', location: '主!B2' }, { ref: 'A4', target: 'https://example.com/' }]);
  });

  it('all links gone: no empty <hyperlinks>', async () => {
    const { xml } = await premerge({ sheets: [{ name: MAIN, data: mainRows(), tail: '<hyperlinks><hyperlink ref="A2" location="他!A1"/></hyperlinks>' }, other()] });
    expect(xml).not.toContain('hyperlinks');
  });

  it('sparklines that read a removed sheet go with their empty groups', async () => {
    const g = (inner: string) => `<x14:sparklineGroup displayEmptyCellsAs="gap"><x14:colorSeries rgb="FF376092"/><x14:sparklines>${inner}</x14:sparklines></x14:sparklineGroup>`;
    const sp = (f: string, sq: string) => `<x14:sparkline><xm:f>${f}</xm:f><xm:sqref>${sq}</xm:sqref></x14:sparkline>`;
    const tail = `<extLst><ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">${g(sp('他!A1:D1', 'E2') + sp('主!B2:B5', 'E3'))}${g(sp('他!B1:E1', 'E4'))}</x14:sparklineGroups></ext></extLst>`;
    const { xml, res } = await premerge({ sheets: [{ name: MAIN, data: mainRows(), tail }, other()] });
    expect(res.report.items.find((i) => i.code === 'sparklinesDropped')!.message).toBe('他のシートのデータを使うスパークライン 2 個を外しました');
    expect((xml.match(/<x14:sparkline>/g) ?? []).length).toBe(1);
    expect((xml.match(/<x14:sparklineGroup\b/g) ?? []).length).toBe(1);
    expect(xml).not.toContain('他');
  });
});

describe('drawings: charts, pictures, shapes', () => {
  const NSD = 'xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"';
  const pt = (c: number, r: number) => `<xdr:col>${c}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${r}</xdr:row><xdr:rowOff>0</xdr:rowOff>`;
  const anchor = (n: number, inner: string) => `<xdr:twoCellAnchor><xdr:from>${pt(5 + n, 1)}</xdr:from><xdr:to>${pt(8 + n, 8)}</xdr:to>${inner}<xdr:clientData/></xdr:twoCellAnchor>`;
  const chartFrame = (id: number, rid: string) => `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="Chart ${id}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="${rid}"/></a:graphicData></a:graphic></xdr:graphicFrame>`;
  const pic = (id: number, rid: string) => `<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${id}" name="Picture ${id}"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr/></xdr:pic>`;
  const box = (id: number, link: string) => `<xdr:sp macro="" textlink="${link}"><xdr:nvSpPr><xdr:cNvPr id="${id}" name="TextBox ${id}"/><xdr:cNvSpPr txBox="1"/></xdr:nvSpPr><xdr:spPr/><xdr:txBody><a:bodyPr/><a:p><a:r><a:t>t</a:t></a:r></a:p></xdr:txBody></xdr:sp>`;
  const chart = (...fs: string[]) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:chart><c:plotArea><c:layout/><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:ser><c:idx val="0"/><c:order val="0"/>${fs.map((f, i) => (i === 0 ? `<c:tx><c:strRef><c:f>${f}</c:f></c:strRef></c:tx>` : `<c:val><c:numRef><c:f>${f}</c:f></c:numRef></c:val>`)).join('')}</c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>`;
  const PNG = new Uint8Array(Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==', 'base64'));

  const build = (wsDr: string, charts: Record<string, string>, rels: RelDef[]): PkgDef => ({
    sheets: [
      { name: MAIN, data: mainRows(), tail: '<drawing r:id="rId1"/>', rels: [{ id: 'rId1', type: 'drawing', target: '../drawings/drawing1.xml' }] },
      other(),
    ],
    parts: [
      { path: 'xl/drawings/drawing1.xml', xml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr ${NSD}>${wsDr}</xdr:wsDr>`, ct: CT.drawing },
      ...(rels.some((r) => r.type === 'image') ? [{ path: 'xl/media/image1.png', data: PNG }] : []),
      ...Object.entries(charts).map(([p, xml]) => ({ path: p, xml, ct: CT.chart })),
    ],
    partRels: { 'xl/drawings/drawing1.xml': rels },
  });

  it('a chart that reads a removed sheet (also through a name / a multi-area range) is removed with its part; own-sheet charts, pictures and shapes stay', async () => {
    const d = build(
      anchor(0, chartFrame(2, 'rId1')) + anchor(1, chartFrame(3, 'rId2')) + anchor(2, pic(4, 'rId5')) + anchor(3, chartFrame(5, 'rId3')) + anchor(4, chartFrame(6, 'rId4')) + anchor(5, box(7, '他!$A$1')) + anchor(6, box(8, '主!$A$1')),
      {
        'xl/charts/chart1.xml': chart('主!$B$1', '主!$B$2:$B$5'),
        'xl/charts/chart2.xml': chart('主!$B$1', '他!$A$1:$A$5'),
        'xl/charts/chart3.xml': chart('主!$B$1', '(主!$A$2,他!$B$1)'),
        'xl/charts/chart4.xml': chart('主!$B$1', 'Nx_1'),
      },
      [
        { id: 'rId1', type: 'chart', target: '../charts/chart1.xml' }, { id: 'rId2', type: 'chart', target: '../charts/chart2.xml' },
        { id: 'rId3', type: 'chart', target: '../charts/chart3.xml' }, { id: 'rId4', type: 'chart', target: '../charts/chart4.xml' },
        { id: 'rId5', type: 'image', target: '../media/image1.png' },
      ],
    );
    d.names = '<definedName name="Nx_1">他!$A$1:$A$5</definedName>';
    const { res, f, out } = await premerge(d);
    await clean(res, [EXT]);
    expect(res.report.items.find((i) => i.code === 'chartsDropped')!.message).toBe('他のシートのデータを使うグラフ 3 個は含めません');
    expect([...f.keys()].filter((k) => k.startsWith('xl/charts/')).sort()).toEqual(['xl/charts/chart1.xml']);
    const dr = f.get('xl/drawings/drawing1.xml')!;
    expect((dr.match(/<xdr:twoCellAnchor>/g) ?? []).length).toBe(4); // chart 1, the picture, two text boxes
    expect(dr).toContain('r:id="rId1"');
    expect(dr).toContain('r:embed="rId5"');
    expect(dr).toContain('textlink=""');
    expect(dr).toContain('textlink="主!$A$1"');
    expect(dr).not.toContain('他');
    const rels = f.get('xl/drawings/_rels/drawing1.xml.rels')!;
    expect(rels).toContain('chart1.xml');
    expect(rels).toContain('image1.png');
    expect(rels).not.toMatch(/chart[234]\.xml/);
    expect(f.get('[Content_Types].xml')).not.toMatch(/chart[234]\.xml/);
    expect(out.features).toMatchObject({ charts: 1, images: 1 });
  });

  it('an mc:AlternateContent wrapper around the anchor of a removed chart goes with it', async () => {
    const wrapped = `<mc:AlternateContent><mc:Choice Requires="c14">${anchor(1, chartFrame(3, 'rId2'))}</mc:Choice><mc:Fallback>${anchor(1, box(9, ''))}</mc:Fallback></mc:AlternateContent>`;
    const d = build(
      anchor(0, chartFrame(2, 'rId1')) + wrapped,
      { 'xl/charts/chart1.xml': chart('主!$B$1', '主!$B$2:$B$5'), 'xl/charts/chart2.xml': chart('主!$B$1', '他!$A$1:$A$5') },
      [{ id: 'rId1', type: 'chart', target: '../charts/chart1.xml' }, { id: 'rId2', type: 'chart', target: '../charts/chart2.xml' }],
    );
    const { res, f } = await premerge(d);
    await clean(res, [EXT]);
    const dr = f.get('xl/drawings/drawing1.xml')!;
    expect(dr).not.toContain('AlternateContent');
    expect((dr.match(/<xdr:twoCellAnchor>/g) ?? []).length).toBe(1);
  });

  it('every chart gone: the drawing stays (empty) and valid', async () => {
    const d = build(anchor(0, chartFrame(2, 'rId1')), { 'xl/charts/chart1.xml': chart('他!$A$1', '他!$A$2:$A$5') }, [{ id: 'rId1', type: 'chart', target: '../charts/chart1.xml' }]);
    const { res, f } = await premerge(d);
    await clean(res, [EXT]);
    expect(f.get('xl/drawings/drawing1.xml')).toContain('</xdr:wsDr>');
    expect([...f.keys()].some((k) => k.startsWith('xl/charts/'))).toBe(false);
  });
});

describe('workbook parts: chart / dialog sheets, calcChain, pivot tables and caches, slicers, document properties', () => {
  const pivotCache = (sheet: string) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<pivotCacheDefinition ${NS_MAIN} xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1" refreshedBy="x" refreshedDate="45000" createdVersion="6" refreshedVersion="6" recordCount="2"><cacheSource type="worksheet"><worksheetSource ref="A1:B3" sheet="${sheet}"/></cacheSource><cacheFields count="2"><cacheField name="ID" numFmtId="0"><sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1" containsInteger="1" minValue="1" maxValue="2"/></cacheField><cacheField name="値" numFmtId="0"><sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1" containsInteger="1" minValue="10" maxValue="20"/></cacheField></cacheFields></pivotCacheDefinition>`;
  const pivotRecords = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<pivotCacheRecords ${NS_MAIN} count="2"><r><n v="1"/><n v="10"/></r><r><n v="2"/><n v="20"/></r></pivotCacheRecords>`;
  const pivotTable = (name: string, cacheId: number, ref: string) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<pivotTableDefinition ${NS_MAIN} name="${name}" cacheId="${cacheId}" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1" dataCaption="値" updatedVersion="6" minRefreshableVersion="3" createdVersion="6" indent="0" outline="1" outlineData="1"><location ref="${ref}" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/><pivotFields count="2"><pivotField axis="axisRow" showAll="0"><items count="1"><item t="default"/></items></pivotField><pivotField dataField="1" showAll="0"/></pivotFields><dataFields count="1"><dataField name="合計" fld="1" baseField="0" baseItem="0"/></dataFields></pivotTableDefinition>`;

  it('chart sheets and dialog sheets are removed like worksheets, with the parts only they use; calcChain goes', async () => {
    const d: PkgDef = {
      sheets: [
        { name: MAIN, data: mainRows() },
        { name: 'グラフ1', kind: 'chartsheet', xml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<chartsheet ${NS_MAIN} xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetViews><sheetView workbookViewId="0"/></sheetViews><drawing r:id="rId1"/></chartsheet>`, rels: [{ id: 'rId1', type: 'drawing', target: '../drawings/drawing1.xml' }] },
        { name: 'ダイアログ', kind: 'dialogsheet', xml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<dialogsheet ${NS_MAIN}><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="13.5"/><sheetData/></dialogsheet>` },
        other(),
      ],
      parts: [
        { path: 'xl/drawings/drawing1.xml', xml: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"/>', ct: CT.drawing },
      ],
      calcChain: true,
    };
    const { res, f } = await premerge(d);
    await clean(res, ['グラフ1', 'ダイアログ', EXT]);
    expect(res.report.droppedSheets).toEqual(['グラフ1', 'ダイアログ', EXT]);
    expect([...f.keys()].filter((k) => /chartsheets|dialogsheets|drawings|calcChain/.test(k))).toEqual([]);
  });

  it('pivot tables on removed sheets go with their cache definition, records and the <pivotCaches> entry; a pivot on the merged sheet keeps its cache', async () => {
    const d: PkgDef = {
      sheets: [
        { name: MAIN, data: mainRows(), rels: [{ id: 'rId1', type: 'pivotTable', target: '../pivotTables/pivotTable2.xml' }] },
        { name: EXT, data: other().data, rels: [{ id: 'rId1', type: 'pivotTable', target: '../pivotTables/pivotTable1.xml' }] },
      ],
      parts: [
        { path: 'xl/pivotTables/pivotTable1.xml', xml: pivotTable('PT1', 1, 'A3:B6'), ct: CT.pivotTable },
        { path: 'xl/pivotTables/pivotTable2.xml', xml: pivotTable('PT2', 2, 'F2:G5'), ct: CT.pivotTable },
        { path: 'xl/pivotCache/pivotCacheDefinition1.xml', xml: pivotCache('他'), ct: CT.pivotCacheDef },
        { path: 'xl/pivotCache/pivotCacheRecords1.xml', xml: pivotRecords, ct: CT.pivotCacheRec },
        { path: 'xl/pivotCache/pivotCacheDefinition2.xml', xml: pivotCache(MAIN), ct: CT.pivotCacheDef },
        { path: 'xl/pivotCache/pivotCacheRecords2.xml', xml: pivotRecords, ct: CT.pivotCacheRec },
      ],
      partRels: {
        'xl/pivotTables/pivotTable1.xml': [{ id: 'rId1', type: 'pivotCacheDefinition', target: '../pivotCache/pivotCacheDefinition1.xml' }],
        'xl/pivotTables/pivotTable2.xml': [{ id: 'rId1', type: 'pivotCacheDefinition', target: '../pivotCache/pivotCacheDefinition2.xml' }],
        'xl/pivotCache/pivotCacheDefinition1.xml': [{ id: 'rId1', type: 'pivotCacheRecords', target: 'pivotCacheRecords1.xml' }],
        'xl/pivotCache/pivotCacheDefinition2.xml': [{ id: 'rId1', type: 'pivotCacheRecords', target: 'pivotCacheRecords2.xml' }],
      },
      wbRels: [{ id: 'rId11', type: 'pivotCacheDefinition', target: 'pivotCache/pivotCacheDefinition1.xml' }, { id: 'rId12', type: 'pivotCacheDefinition', target: 'pivotCache/pivotCacheDefinition2.xml' }],
      wbExtra: '<pivotCaches><pivotCache cacheId="1" r:id="rId11"/><pivotCache cacheId="2" r:id="rId12"/></pivotCaches>',
    };
    const { res, f, out } = await premerge(d);
    await clean(res, [EXT]);
    const keys = [...f.keys()].filter((k) => /pivot/i.test(k)).sort();
    expect(keys).toEqual(['xl/pivotCache/_rels/pivotCacheDefinition2.xml.rels', 'xl/pivotCache/pivotCacheDefinition2.xml', 'xl/pivotCache/pivotCacheRecords2.xml', 'xl/pivotTables/_rels/pivotTable2.xml.rels', 'xl/pivotTables/pivotTable2.xml', 'xl/worksheets/_rels/sheet1.xml.rels'].filter((k) => !/worksheets/.test(k)).sort());
    const wb = f.get('xl/workbook.xml')!;
    expect(wb).toContain('<pivotCaches><pivotCache cacheId="2" r:id="rId12"/></pivotCaches>');
    expect(out.features.pivotTables).toBe(1);
    expect(res.report.items.some((i) => i.code === 'pivotsToValues')).toBe(false);
  });

  it('a pivot table on the merged sheet that is fed by a removed sheet stays as plain values (no pivot, no cache)', async () => {
    const d: PkgDef = {
      sheets: [
        { name: MAIN, data: mainRows(), head: '', rels: [{ id: 'rId1', type: 'pivotTable', target: '../pivotTables/pivotTable1.xml' }] },
        other(),
      ],
      parts: [
        { path: 'xl/pivotTables/pivotTable1.xml', xml: pivotTable('PT1', 1, 'F2:G5'), ct: CT.pivotTable },
        { path: 'xl/pivotCache/pivotCacheDefinition1.xml', xml: pivotCache('他'), ct: CT.pivotCacheDef },
        { path: 'xl/pivotCache/pivotCacheRecords1.xml', xml: pivotRecords, ct: CT.pivotCacheRec },
      ],
      partRels: {
        'xl/pivotTables/pivotTable1.xml': [{ id: 'rId1', type: 'pivotCacheDefinition', target: '../pivotCache/pivotCacheDefinition1.xml' }],
        'xl/pivotCache/pivotCacheDefinition1.xml': [{ id: 'rId1', type: 'pivotCacheRecords', target: 'pivotCacheRecords1.xml' }],
      },
      wbRels: [{ id: 'rId11', type: 'pivotCacheDefinition', target: 'pivotCache/pivotCacheDefinition1.xml' }],
      wbExtra: '<pivotCaches><pivotCache cacheId="1" r:id="rId11"/></pivotCaches>',
    };
    const { res, f } = await premerge(d);
    await clean(res, [EXT]);
    expect([...f.keys()].filter((k) => /pivot/i.test(k))).toEqual([]);
    expect(f.get('xl/workbook.xml')).not.toContain('pivotCaches');
    expect(res.report.items.find((i) => i.code === 'pivotsToValues')!.message).toBe('他のシートのデータを使うピボットテーブル 1 個を値にしました');
  });

  it('slicer and timeline caches of pivot tables on removed sheets go with their names; slicers on the merged sheet that were bound to them go too', async () => {
    const X14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main';
    const cache = (name: string, tab: number, pivot: string) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<slicerCacheDefinition xmlns="${X14}" name="${name}" sourceName="値"><pivotTables><pivotTable tabId="${tab}" name="${pivot}"/></pivotTables><data><tabular pivotCacheId="1"><items count="1"><i x="0" s="1"/></items></tabular></data></slicerCacheDefinition>`;
    const d: PkgDef = {
      sheets: [
        { name: MAIN, data: mainRows(), tail: '<drawing r:id="rId1"/><extLst><ext uri="{A8765BA9-456A-4dab-B4F3-ACF838C121DE}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:slicerList><x14:slicer r:id="rId2"/></x14:slicerList></ext></extLst>', rels: [{ id: 'rId1', type: 'drawing', target: '../drawings/drawing1.xml' }, { id: 'rId2', type: 'http://schemas.microsoft.com/office/2007/relationships/slicer', target: '../slicers/slicer1.xml' }] },
        { name: EXT, data: other().data, rels: [{ id: 'rId1', type: 'pivotTable', target: '../pivotTables/pivotTable1.xml' }] },
      ],
      parts: [
        { path: 'xl/pivotTables/pivotTable1.xml', xml: pivotTable('PT1', 1, 'A3:B6'), ct: CT.pivotTable },
        { path: 'xl/pivotCache/pivotCacheDefinition1.xml', xml: pivotCache('他'), ct: CT.pivotCacheDef },
        { path: 'xl/pivotCache/pivotCacheRecords1.xml', xml: pivotRecords, ct: CT.pivotCacheRec },
        { path: 'xl/slicerCaches/slicerCache1.xml', xml: cache('Slicer_値', 2, 'PT1'), ct: CT.slicerCache },
        { path: 'xl/slicers/slicer1.xml', xml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<slicers xmlns="${X14}"><slicer name="値" cache="Slicer_値" caption="値" rowHeight="241300"/></slicers>`, ct: CT.slicer },
        {
          path: 'xl/drawings/drawing1.xml',
          xml: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
            + '<mc:AlternateContent><mc:Choice xmlns:sle15="http://schemas.microsoft.com/office/drawing/2012/slicer" Requires="sle15"><xdr:twoCellAnchor><xdr:from><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>7</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>8</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="値"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/drawing/2010/slicer"><sle:slicer xmlns:sle="http://schemas.microsoft.com/office/drawing/2010/slicer" name="値"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor></mc:Choice><mc:Fallback><xdr:twoCellAnchor><xdr:from><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>7</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>8</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="0" name=""/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr/></xdr:sp><xdr:clientData/></xdr:twoCellAnchor></mc:Fallback></mc:AlternateContent></xdr:wsDr>',
          ct: CT.drawing,
        },
      ],
      partRels: {
        'xl/pivotTables/pivotTable1.xml': [{ id: 'rId1', type: 'pivotCacheDefinition', target: '../pivotCache/pivotCacheDefinition1.xml' }],
        'xl/pivotCache/pivotCacheDefinition1.xml': [{ id: 'rId1', type: 'pivotCacheRecords', target: 'pivotCacheRecords1.xml' }],
      },
      wbRels: [{ id: 'rId11', type: 'pivotCacheDefinition', target: 'pivotCache/pivotCacheDefinition1.xml' }, { id: 'rId12', type: 'http://schemas.microsoft.com/office/2007/relationships/slicerCache', target: 'slicerCaches/slicerCache1.xml' }],
      wbExtra: '<pivotCaches><pivotCache cacheId="1" r:id="rId11"/></pivotCaches><extLst><ext uri="{BBE1A952-AA13-448e-AADC-164F8A28A991}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:slicerCaches><x14:slicerCache r:id="rId12"/></x14:slicerCaches></ext></extLst>',
      names: '<definedName name="Slicer_値">#N/A</definedName><definedName name="Keep">主!$A$1</definedName>',
    };
    const { res, f, xml } = await premerge(d);
    await clean(res, [EXT]);
    expect([...f.keys()].filter((k) => /slicer|pivot|drawing/i.test(k))).toEqual(['xl/drawings/drawing1.xml']);
    const wb = f.get('xl/workbook.xml')!;
    expect(wb).not.toMatch(/slicer|Slicer|extLst/);
    expect(wb).toContain('<definedName name="Keep">主!$A$1</definedName>');
    expect(xml).not.toMatch(/slicer|extLst/);
    expect(xml).toContain('<drawing r:id="rId1"/>');
    expect(f.get('xl/drawings/drawing1.xml')).not.toMatch(/AlternateContent|graphicFrame/);
    expect(res.report.items.find((i) => i.code === 'slicersDropped')!.message).toContain('1 個');
  });

  it('threaded comments of removed sheets: the list of people goes when no threaded comment is left; notes of the merged sheet stay', async () => {
    const d: PkgDef = {
      sheets: [
        { name: MAIN, data: mainRows(), tail: '<legacyDrawing r:id="rId1"/>', rels: [{ id: 'rId1', type: 'vmlDrawing', target: '../drawings/vmlDrawing1.vml' }, { id: 'rId2', type: 'comments', target: '../comments1.xml' }] },
        { name: EXT, data: other().data, rels: [{ id: 'rId1', type: 'http://schemas.microsoft.com/office/2017/10/relationships/threadedComment', target: '../threadedComments/threadedComment1.xml' }, { id: 'rId2', type: 'comments', target: '../comments2.xml' }] },
      ],
      parts: [
        { path: 'xl/drawings/vmlDrawing1.vml', xml: '<xml xmlns:v="urn:schemas-microsoft-com:vml"/>' },
        { path: 'xl/comments1.xml', xml: `<comments ${NS_MAIN}><authors><author>a</author></authors><commentList><comment ref="A2" authorId="0"><text><r><t>n</t></r></text></comment></commentList></comments>`, ct: CT.comments },
        { path: 'xl/comments2.xml', xml: `<comments ${NS_MAIN}><authors><author>a</author></authors><commentList><comment ref="A2" authorId="0"><text><r><t>n</t></r></text></comment></commentList></comments>`, ct: CT.comments },
        { path: 'xl/threadedComments/threadedComment1.xml', xml: '<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"/>', ct: 'application/vnd.ms-excel.threadedcomments+xml' },
        { path: 'xl/persons/person.xml', xml: '<personList xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"/>', ct: 'application/vnd.ms-excel.person+xml' },
      ],
      wbRels: [{ id: 'rId11', type: 'http://schemas.microsoft.com/office/2017/10/relationships/person', target: 'persons/person.xml' }],
    };
    const { res, f } = await premerge(d);
    await clean(res, [EXT]);
    expect([...f.keys()].filter((k) => /comments|person|threaded|vml/i.test(k)).sort()).toEqual(['xl/comments1.xml', 'xl/drawings/vmlDrawing1.vml']);
    expect(f.get('xl/_rels/workbook.xml.rels')).not.toContain('person');
  });

  it('docProps/app.xml: sheets and named ranges are listed for the one sheet; an unexpected layout drops the lists', async () => {
    const vt = 'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"';
    const app = (inner: string) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" ${vt}><Application>Microsoft Excel</Application>${inner}<AppVersion>16.0300</AppVersion></Properties>`;
    const names = '<definedName name="Good">主!$A$1</definedName><definedName name="Bad">他!$A$1</definedName><definedName name="Loc" localSheetId="0">主!$B$2</definedName><definedName name="_xlnm.Print_Area" localSheetId="0">主!$A$1:$B$9</definedName><definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">主!$A$1:$B$9</definedName>';
    const ok = app('<HeadingPairs><vt:vector size="6" baseType="variant"><vt:variant><vt:lpstr>ワークシート</vt:lpstr></vt:variant><vt:variant><vt:i4>2</vt:i4></vt:variant><vt:variant><vt:lpstr>グラフ</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant><vt:variant><vt:lpstr>名前付き一覧</vt:lpstr></vt:variant><vt:variant><vt:i4>4</vt:i4></vt:variant></vt:vector></HeadingPairs>'
      + '<TitlesOfParts><vt:vector size="7" baseType="lpstr"><vt:lpstr>主</vt:lpstr><vt:lpstr>他</vt:lpstr><vt:lpstr>グラフ1</vt:lpstr><vt:lpstr>Bad</vt:lpstr><vt:lpstr>Good</vt:lpstr><vt:lpstr>主!Loc</vt:lpstr><vt:lpstr>主!Print_Area</vt:lpstr></vt:vector></TitlesOfParts>');
    const a = await premerge({ sheets: [{ name: MAIN, data: mainRows() }, other()], names, app: ok });
    await clean(a.res, [EXT]);
    const text = a.f.get('docProps/app.xml')!;
    expect(text).toContain('<HeadingPairs><vt:vector size="4" baseType="variant"><vt:variant><vt:lpstr>ワークシート</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant><vt:variant><vt:lpstr>名前付き一覧</vt:lpstr></vt:variant><vt:variant><vt:i4>3</vt:i4></vt:variant></vt:vector></HeadingPairs>');
    expect(text).toContain('<TitlesOfParts><vt:vector size="4" baseType="lpstr"><vt:lpstr>主</vt:lpstr><vt:lpstr>Good</vt:lpstr><vt:lpstr>主!Loc</vt:lpstr><vt:lpstr>主!Print_Area</vt:lpstr></vt:vector></TitlesOfParts>');
    expect(text).toContain('<Application>Microsoft Excel</Application>');
    // mismatching counts: the lists are removed (optional elements)
    const odd = app('<HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>ワークシート</vt:lpstr></vt:variant><vt:variant><vt:i4>5</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size="1" baseType="lpstr"><vt:lpstr>主</vt:lpstr></vt:vector></TitlesOfParts>');
    const b = await premerge({ sheets: [{ name: MAIN, data: mainRows() }, other()], app: odd });
    expect(b.f.get('docProps/app.xml')).not.toMatch(/HeadingPairs|TitlesOfParts/);
    expect(b.f.get('docProps/app.xml')).toContain('<Application>Microsoft Excel</Application>');
  });

  it('the merged sheet is visible and selected even when it was hidden or not the active tab', async () => {
    const { f, res } = await premerge({ sheets: [other(), { name: MAIN, state: 'hidden', data: mainRows() }], bookViews: '<bookViews><workbookView activeTab="0" firstSheet="0"/></bookViews>' });
    await clean(res, [EXT]);
    expect(f.get('xl/workbook.xml')).toContain('<sheets><sheet name="主" sheetId="2" r:id="rId2"/></sheets>');
    expect(f.get('xl/worksheets/sheet2.xml')).toContain('tabSelected="1"');
  });
});

describe('inside the premerge: imports, appended rows, tables and totals after dropping', () => {
  it('formulas below the table that read a removed sheet are detached and reported with their final rows when rows are appended', async () => {
    const src = await pkgOf({ sheets: [{ name: 'S', data: row(1, `${str('A1', 'ID')}${str('B1', 'メモ')}`) + row(2, `${num('A2', 6)}${str('B2', 'm6')}`) }] });
    const baseBytes = await pkgOf({
      sheets: [
        {
          name: MAIN,
          data: row(1, `${str('A1', 'ID')}${str('B1', '値')}`) + [2, 3].map((r) => row(r, `${num(`A${r}`, r - 1)}${num(`B${r}`, r * 10)}`)).join('')
            + row(5, `${str('A5', '合計')}${fml('B5', 'SUM(B2:B3)', 50)}${fml('C5', '他!A1', 7)}`) + row(6, fml('B6', 'SUM(他!B1:E1)', 10)),
        },
        other(),
      ],
    });
    const baseSrc = bareBase(baseBytes, { sheet: MAIN, headerRow: 1, firstDataRow: 2, lastDataRow: 3, firstCol: 1, lastCol: 2 });
    const res = await buildPremerged({
      base: baseSrc,
      sources: [bareBase(src, { sheet: 'S', headerRow: 1, firstDataRow: 2, lastDataRow: 2, firstCol: 1, lastCol: 2 }, 's')].map((k) => ({ ...k, id: 's' })),
      rows: [{ baseRow: 2, matches: {} }, { baseRow: 3, matches: {} }, { baseRow: null, appendedFrom: { sourceId: 's', row: 2 }, matches: { s: { row: 2, via: 'ID', flags: [] } } }],
      imports: [{ sourceId: 's', col: 2 }],
      keyFill: { s: [{ baseCol: 1, sourceCol: 1 }] },
      options: baseOptions({ dropOtherSheets: true }),
    });
    await clean(res, [EXT]);
    expect(res.report).toMatchObject({ droppedSheets: [EXT], formulasDetached: 2, appendedRows: 1, importedColumns: 1 });
    expect(res.report.items.find((i) => i.code === 'detachedFormulas')!.rows).toEqual([6, 7]); // were rows 5 and 6; one row was inserted above them
    const s = (await readBook(res.bytes, 'o.xlsx')).sheets[0];
    // the imported column went in at C: the old C5 is D6 now
    expect(cell(s, 'D6')).toMatchObject({ v: 7 });
    expect(cell(s, 'D6')?.f).toBeUndefined();
    expect(cell(s, 'B7')?.v).toBe(10);
    expect(cell(s, 'B7')?.f).toBeUndefined();
    expect(cell(s, 'B6')?.f).toBe('SUM(B2:B4)'); // 合計 grew over the appended row
    expect(cell(s, 'C4')?.v).toBe('m6');
  });

  it('a table calculated column that reads a removed name loses its formula: appended rows get no live formula pointing at a removed sheet', async () => {
    const tableXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<table ${NS_MAIN} id="1" name="T1" displayName="T1" ref="A1:C3" totalsRowShown="0"><autoFilter ref="A1:C3"/><tableColumns count="3"><tableColumn id="1" name="ID"/><tableColumn id="2" name="値"/><tableColumn id="3" name="税込"><calculatedColumnFormula>T1[[#This Row],[値]]*Rate</calculatedColumnFormula></tableColumn></tableColumns><tableStyleInfo name="TableStyleMedium2" showRowStripes="1"/></table>`;
    const baseBytes = await pkgOf({
      sheets: [
        {
          name: MAIN,
          data: row(1, `${str('A1', 'ID')}${str('B1', '値')}${str('C1', '税込')}`)
            + row(2, `${num('A2', 1)}${num('B2', 10)}${fml('C2', 'T1[[#This Row],[値]]*Rate', 11)}`) + row(3, `${num('A3', 2)}${num('B3', 20)}${fml('C3', 'T1[[#This Row],[値]]*Rate', 22)}`),
          tail: '<tableParts count="1"><tablePart r:id="rId1"/></tableParts>',
          rels: [{ id: 'rId1', type: 'table', target: '../tables/table1.xml' }],
        },
        { name: EXT, data: row(1, num('A1', 1.1)) },
      ],
      parts: [{ path: 'xl/tables/table1.xml', xml: tableXml, ct: CT.table }],
      names: '<definedName name="Rate">他!$A$1</definedName>',
    });
    const srcBytes = await pkgOf({ sheets: [{ name: 'S', data: row(1, `${str('A1', 'ID')}${str('B1', 'x')}`) + row(2, `${num('A2', 5)}${num('B2', 50)}`) }] });
    const res = await buildPremerged({
      base: bareBase(baseBytes, { sheet: MAIN, headerRow: 1, firstDataRow: 2, lastDataRow: 3, firstCol: 1, lastCol: 3 }),
      sources: [{ ...bareBase(srcBytes, { sheet: 'S', headerRow: 1, firstDataRow: 2, lastDataRow: 2, firstCol: 1, lastCol: 2 }, 's'), id: 's' }],
      rows: [{ baseRow: 2, matches: {} }, { baseRow: 3, matches: {} }, { baseRow: null, appendedFrom: { sourceId: 's', row: 2 }, matches: { s: { row: 2, via: 'ID', flags: [] } } }],
      imports: [{ sourceId: 's', col: 2 }],
      keyFill: { s: [{ baseCol: 1, sourceCol: 1 }] },
      options: baseOptions({ dropOtherSheets: true }),
    });
    await clean(res, [EXT]);
    expect(res.report.formulasDetached).toBe(2);
    const f = await files(res.bytes);
    expect(f.get('xl/tables/table1.xml')).not.toContain('calculatedColumnFormula');
    expect(f.get('xl/tables/table1.xml')).toContain('ref="A1:D4"');
    const s = (await readBook(res.bytes, 'o.xlsx')).sheets[0];
    expect(s.features.tables.map((t) => t.ref)).toEqual(['A1:D4']);
    expect([2, 3, 4].map((r) => cell(s, `C${r}`)?.f)).toEqual([undefined, undefined, undefined]);
    expect(cell(s, 'C2')?.v).toBe(11);
  });

  it('a workbook with only one sheet: nothing to drop, droppedSheets is empty, the output equals the default one', async () => {
    const d: PkgDef = { sheets: [{ name: MAIN, data: mainRows((r) => (r === 2 ? fml('C2', 'B2+1', 11) : '')) }], names: '<definedName name="X">主!$A$1</definedName>' };
    const bytes = await pkgOf(d);
    const on = await premergeOne(bareBase(bytes, region(9)), { dropOtherSheets: true });
    const off = await premergeOne(bareBase(bytes, region(9)), {});
    expect(on.report.droppedSheets).toEqual([]);
    expect(on.report.formulasDetached).toBe(0);
    expect(on.report.items).toEqual([]);
    expect([...(await files(on.bytes))]).toEqual([...(await files(off.bytes))]);
    expect(off.report.droppedSheets).toBeUndefined();
  });

  it('bigger sheets: 20,000 formula rows with names and a long list on a removed sheet in a few seconds', async () => {
    const n = 20000;
    const data = row(1, `${str('A1', 'ID')}${str('B1', '値')}`) + Array.from({ length: n }, (_x, i) => row(i + 2, `${num(`A${i + 2}`, i)}${num(`B${i + 2}`, i * 2)}${fml(`C${i + 2}`, `B${i + 2}*Rate`, i)}${fml(`D${i + 2}`, `B${i + 2}+1`, i)}${fml(`E${i + 2}`, `VLOOKUP(A${i + 2},他!$A:$B,2,FALSE)`, 1)}`)).join('');
    const rem = Array.from({ length: 5000 }, (_x, i) => row(i + 1, `${num(`A${i + 1}`, i)}${num(`B${i + 1}`, i)}`)).join('');
    const bytes = await pkgOf({ sheets: [{ name: MAIN, data }, { name: EXT, data: rem }], names: '<definedName name="Rate">他!$A$1</definedName>' });
    const t0 = Date.now();
    const res = await premergeOne(bareBase(bytes, region(n + 1, 2)), { dropOtherSheets: true });
    expect(res.report.formulasDetached).toBe(n * 2);
    expect(Date.now() - t0).toBeLessThan(15000);
  }, 60000);
});

describe('robustness', () => {
  it('a removed sheet whose XML is damaged does not matter unless a list validation has to read it (then the validation is dropped)', async () => {
    const d: PkgDef = {
      sheets: [
        { name: MAIN, data: mainRows(), tail: '<dataValidations count="1"><dataValidation type="list" sqref="B2:B9"><formula1>他!$A$1:$A$3</formula1></dataValidation></dataValidations>' },
        { name: EXT, xml: '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>x</t></is></c></row>' },
      ],
    };
    const { res, sheet } = await premerge(d);
    await clean(res, [EXT]);
    expect(sheet.features.rules.filter((r) => r.kind === 'dv')).toEqual([]);
    expect(res.report.items.some((i) => i.code === 'dvDropped')).toBe(true);
  });

  it('a removed sheet that is missing from the package (dangling relationship) is dropped from the workbook all the same', async () => {
    const bytes = await pkgOf({ sheets: [{ name: MAIN, data: mainRows() }, other()] });
    const JSZip = (await import('jszip')).default;
    const z = await JSZip.loadAsync(bytes);
    z.remove('xl/worksheets/sheet2.xml');
    const broken = await z.generateAsync({ type: 'uint8array' });
    const res = await premergeOne(bareBase(broken, region(9)), { dropOtherSheets: true });
    expect(res.report.droppedSheets).toEqual([EXT]);
    expect((await files(res.bytes)).get('xl/workbook.xml')).not.toContain('他');
  });
});
