/**
 * What the core copy leaves out (phase 1: drawings / charts / pictures / shapes, comments and form controls, tables, pivot tables,
 * ActiveX / OLE controls, slicers, backgrounds, chart sheets): removed cleanly (no dangling r:id, no orphan part, no leftover relationship)
 * and named in ONE 'sheetCopyLimits' warning per sheet; sheets that cannot be copied at all (chart sheets) are reported by name.
 * The sources are hand-made packages (pkgOf of dropsheets-helpers), the base a one-sheet workbook.
 */
import { describe, expect, it } from 'vitest';
import { readBook } from '../../src/core/xlsx';
import { CT, fml, num, pkgOf, row, str, type PkgDef } from './dropsheets-helpers';
import { insertWorksheetElement } from '../../src/core/xlsx/patch/addsheets';
import { lintPackage, lintText, premergeCopies, readFiles, sheetsOf, type FileIn } from './addsheets-helpers';

const BASE_REGION = { sheet: '主', headerRow: 1, firstDataRow: 2, lastDataRow: 3, firstCol: 1, lastCol: 2 };
const base = async (): Promise<FileIn> => ({
  bytes: await pkgOf({ sheets: [{ name: '主', data: row(1, str('A1', 'ID') + str('B1', '値')) + row(2, str('A2', 'x') + num('B2', 1)) + row(3, str('A3', 'y') + num('B3', 2)) }] }),
  fileName: '主.xlsx',
  region: BASE_REGION,
});

const DRAWING = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
<xdr:twoCellAnchor><xdr:from><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>8</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>10</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId1"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>
<xdr:twoCellAnchor><xdr:from><xdr:col>9</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>10</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="3" name="Picture 2"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:twoCellAnchor>
<xdr:twoCellAnchor><xdr:from><xdr:col>9</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>5</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>10</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>7</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="4" name="TextBox 3"/><xdr:cNvSpPr txBox="1"/></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr><xdr:txBody><a:bodyPr/><a:p><a:r><a:t>memo</a:t></a:r></a:p></xdr:txBody></xdr:sp><xdr:clientData/></xdr:twoCellAnchor>
</xdr:wsDr>`;

const CHART = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea><c:barChart><c:ser><c:idx val="0"/><c:order val="0"/><c:val><c:numRef><c:f>データ!$B$2:$B$3</c:f></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>`;
const COMMENTS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors><author>me</author></authors><commentList><comment ref="A2" authorId="0"><text><r><t>hello</t></r></text></comment><comment ref="B3" authorId="0"><text><r><t>again</t></r></text></comment></commentList></comments>`;
const VML = `<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"><v:shape id="_x0000_s1025" type="#_x0000_t202"><x:ClientData ObjectType="Note"><x:Row>1</x:Row><x:Column>0</x:Column></x:ClientData></v:shape></xml>`;
const TABLE = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="tblData" displayName="tblData" ref="A1:B3"><autoFilter ref="A1:B3"/><tableColumns count="2"><tableColumn id="1" name="ID"/><tableColumn id="2" name="値"/></tableColumns><tableStyleInfo name="TableStyleMedium2" showRowStripes="1"/></table>`;

/** a source with a busy sheet `データ` (drawing with chart + picture + text box, comments, a table, a form control) and a plain sheet */
async function objectsSource(): Promise<FileIn> {
  const def: PkgDef = {
    sheets: [
      {
        name: 'データ',
        data: row(1, str('A1', 'ID') + str('B1', '値') + str('C1', '計')) + row(2, str('A2', 'a') + num('B2', 1) + fml('C2', 'SUM(tblData[値])', 3)) + row(3, str('A3', 'b') + num('B3', 2)),
        tail: '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><drawing r:id="rId1"/><legacyDrawing r:id="rId2"/><tableParts count="1"><tablePart r:id="rId4"/></tableParts>',
        rels: [
          { id: 'rId1', type: 'drawing', target: '../drawings/drawing1.xml' },
          { id: 'rId2', type: 'vmlDrawing', target: '../drawings/vmlDrawing1.vml' },
          { id: 'rId3', type: 'comments', target: '../comments1.xml' },
          { id: 'rId4', type: 'table', target: '../tables/table1.xml' },
        ],
      },
      { name: '別', data: row(1, num('A1', 1) + fml('B1', 'データ!B2*2', 2)) },
    ],
    parts: [
      { path: 'xl/drawings/drawing1.xml', xml: DRAWING, ct: CT.drawing },
      { path: 'xl/charts/chart1.xml', xml: CHART, ct: CT.chart },
      { path: 'xl/media/image1.png', xml: 'PNG' },
      { path: 'xl/comments1.xml', xml: COMMENTS, ct: CT.comments },
      { path: 'xl/drawings/vmlDrawing1.vml', xml: VML },
      { path: 'xl/tables/table1.xml', xml: TABLE, ct: CT.table },
    ],
    partRels: {
      'xl/drawings/drawing1.xml': [{ id: 'rId1', type: 'chart', target: '../charts/chart1.xml' }, { id: 'rId2', type: 'image', target: '../media/image1.png' }],
    },
  };
  return { bytes: await pkgOf(def), fileName: 'オブジェクト.xlsx', sheet: 'データ' };
}

describe('objects of a source sheet (round 9 phase 2: copied by the object phase, addsheets-objects.ts)', () => {
  // updated for A2: phase 1 left these out and named them in a warning; the object phase copies them (tests/xlsx/addsheets-objects-*.test.ts go into the details)
  it('drawing / chart / picture / text box / comments / table are copied with the sheet: parts, relationships, content types; no warning; the structured reference stays a formula', async () => {
    const res = await premergeCopies(await base(), [await objectsSource()], { dropOtherSheets: true });
    const f = await readFiles(res.bytes);
    const sh = sheetsOf(f).find((s) => s.name === 'オブジェクト_データ')!;
    const x = f.get(sh.part)!;
    expect(x).toMatch(/<drawing r:id=/);
    expect(x).toMatch(/<legacyDrawing r:id=/);
    expect(x).toMatch(/<tableParts count="1">/);
    const relsPath = sh.part.replace('worksheets/', 'worksheets/_rels/') + '.rels';
    expect(f.get(relsPath)).toMatch(/\/drawing"/);
    for (const kind of ['chart', 'drawing', 'comments', 'tables/table', 'vmlDrawing']) expect([...f.keys()].some((p) => p.includes(kind)), kind).toBe(true);
    expect([...f.keys()].some((p) => /^xl\/media\/image\d+\.png$/.test(p))).toBe(true);
    expect(lintText(await lintPackage(res.bytes))).toBe('');
    // cells and values are all there; the formula that reads the table stays a formula (the table is copied under the same name)
    const out = await readBook(res.bytes, 'o.xlsx');
    const g = out.sheets.find((s) => s.name === 'オブジェクト_データ')!;
    expect(g.rows[1]![1]!.v).toBe(1);
    expect(g.rows[1]![2]!.f).toBe('SUM(tblData[値])');
    expect(g.rows[1]![2]!.v).toBe(3);
    // the sheet that only reads a copied sheet by name follows it
    expect(out.sheets.find((s) => s.name === 'オブジェクト_別')!.rows[0]![1]!.f).toBe('オブジェクト_データ!B2*2');
    expect(res.report.items.filter((i) => i.code === 'sheetCopyLimits')).toEqual([]);
    expect(res.layout.addedSheets!.map((s) => [s.name, s.complete])).toEqual([['オブジェクト_データ', true], ['オブジェクト_別', true]]);
  });

  it('what is still not copied (an OLE object, a header picture) is removed cleanly and named in ONE warning of the sheet', async () => {
    const def: PkgDef = {
      sheets: [{
        name: 'データ',
        data: row(1, str('A1', 'ID')) + row(2, str('A2', 'a')),
        tail: '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><legacyDrawingHF r:id="rId1"/><oleObjects><oleObject progId="Word.Document.12" shapeId="1025" r:id="rId2"/></oleObjects>',
        rels: [{ id: 'rId1', type: 'vmlDrawing', target: '../drawings/vmlDrawingHF1.vml' }, { id: 'rId2', type: 'oleObject', target: '../embeddings/oleObject1.bin' }],
      }],
      parts: [{ path: 'xl/drawings/vmlDrawingHF1.vml', xml: VML }, { path: 'xl/embeddings/oleObject1.bin', data: new Uint8Array([1, 2, 3]) }],
    };
    const res = await premergeCopies(await base(), [{ bytes: await pkgOf(def), fileName: 'x.xlsx' }], { dropOtherSheets: true });
    const f = await readFiles(res.bytes);
    for (const p of f.keys()) expect(p, p).not.toMatch(/vml|embeddings/i);
    expect(lintText(await lintPackage(res.bytes))).toBe('');
    const warns = res.report.items.filter((i) => i.code === 'sheetCopyLimits' && i.message.includes('x_データ'));
    expect(warns).toHaveLength(1);
    expect(warns[0].message).toMatch(/ヘッダー・フッターの画像/);
    expect(warns[0].message).toMatch(/埋め込みオブジェクト 1 個/);
  });

  it('a chart sheet is copied as a chart sheet and takes a name (round 9 phase 2); one without anything to show becomes an empty sheet with a warning', async () => {
    const def: PkgDef = {
      sheets: [
        { name: 'S1', data: row(1, num('A1', 1)) },
        { name: 'グラフ', kind: 'chartsheet', xml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetViews><sheetView workbookViewId="0"/></sheetViews><drawing r:id="rId1"/></chartsheet>`, rels: [{ id: 'rId1', type: 'drawing', target: '../drawings/drawing1.xml' }] },
        { name: 'S2', data: row(1, num('A1', 2)) },
      ],
      parts: [{ path: 'xl/drawings/drawing1.xml', xml: DRAWING.replace(/<xdr:twoCellAnchor>[\s\S]*<\/xdr:twoCellAnchor>/, ''), ct: CT.drawing }],
    };
    const res = await premergeCopies(await base(), [{ bytes: await pkgOf(def), fileName: 'g.xlsx', sheet: 'S1' }], { dropOtherSheets: true });
    const names = sheetsOf(await readFiles(res.bytes)).map((s) => s.name);
    expect(names).toEqual(['主', 'g_S1', 'g_グラフ', 'g_S2']);
    // its drawing holds nothing: there is no chart sheet to write, the planned name stays as an empty sheet
    expect(res.report.items.some((i) => i.code === 'sheetCopyLimits' && i.message.includes('「グラフ」') && /コピーできませんでした.*「g_グラフ」は空のシートになります/.test(i.message))).toBe(true);
    expect(lintText(await lintPackage(res.bytes))).toBe('');
  });
});

describe('insertWorksheetElement (the extension point for the object phase)', () => {
  const ws = (body: string) => `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/>${body}</worksheet>`;
  it('puts drawing / legacyDrawing / tableParts where the schema wants them (the nested extLst of a cfRule does not count)', () => {
    let x = ws('<mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells><conditionalFormatting sqref="A1"><cfRule type="dataBar" priority="1"><extLst><ext uri="{x}"/></extLst></cfRule></conditionalFormatting><pageMargins left="1" right="1" top="1" bottom="1" header="0" footer="0"/><extLst><ext uri="{y}"/></extLst>');
    x = insertWorksheetElement(x, 'tableParts', '<tableParts count="1"><tablePart r:id="rId9"/></tableParts>');
    x = insertWorksheetElement(x, 'drawing', '<drawing r:id="rId7"/>');
    x = insertWorksheetElement(x, 'legacyDrawing', '<legacyDrawing r:id="rId8"/>');
    const order = [...x.matchAll(/<(mergeCells|conditionalFormatting|pageMargins|drawing|legacyDrawing|tableParts|extLst)\b/g)].map((m) => m[1]);
    expect(order).toEqual(['mergeCells', 'conditionalFormatting', 'extLst', 'pageMargins', 'drawing', 'legacyDrawing', 'tableParts', 'extLst']);
    expect(x).toMatch(/xmlns:r="http:\/\/schemas.openxmlformats.org\/officeDocument\/2006\/relationships"/); // declared when a relationship id appears
    expect(x.indexOf('<drawing')).toBeLessThan(x.indexOf('<legacyDrawing'));
    expect(x.indexOf('<tableParts')).toBeLessThan(x.lastIndexOf('<extLst'));
  });
});

describe('a sheet that cannot be written back safely', () => {
  it('stays as an EMPTY sheet of its planned name (formulas of the other copies that read it do not turn into #REF!), with a warning', async () => {
    const def: PkgDef = {
      sheets: [
        { name: '壊れ', data: `<row r="1"><c r="A1"><v>1</v></c><c r="A1"><v>2</v></c></row>` }, // two cells at A1: cannot be written back
        { name: '正常', data: row(1, num('A1', 5) + fml('B1', '壊れ!A1+1', 2)) },
      ],
    };
    const res = await premergeCopies(await base(), [{ bytes: await pkgOf(def), fileName: 'x.xlsx', sheet: '正常' }], { dropOtherSheets: true });
    const names = sheetsOf(await readFiles(res.bytes)).map((s) => s.name);
    expect(names).toEqual(['主', 'x_壊れ', 'x_正常']);
    const out = await readBook(res.bytes, 'o.xlsx');
    expect(out.sheets.find((s) => s.name === 'x_壊れ')!.maxRow).toBe(0);
    expect(out.sheets.find((s) => s.name === 'x_正常')!.rows[0]![1]!.f).toBe('x_壊れ!A1+1');
    const w = res.report.items.find((i) => i.code === 'sheetCopyLimits' && i.message.includes('「壊れ」'))!;
    expect(w.message).toMatch(/コピーできませんでした。.*「x_壊れ」は空のシートになります/);
    expect(res.layout.addedSheets!.map((s) => [s.name, s.complete])).toEqual([['x_壊れ', false], ['x_正常', true]]);
    expect(lintText(await lintPackage(res.bytes))).toBe('');
  });
});
