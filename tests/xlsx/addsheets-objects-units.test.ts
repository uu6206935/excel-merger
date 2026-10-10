/**
 * The object phase of the sheet copier (task A2) on hand-made packages: every case the Excel-authored fixtures do not have — what cannot be
 * copied (slicers, SmartArt, ink, form controls / ActiveX placeholders, query / XML tables, pivot caches of external data), what falls
 * back (chart references that cannot follow, a chart that cannot be copied at all), name / id collisions with the base, theme conversion of
 * charts and shapes, new-style charts, Strict namespaces, VML renumbering, sheets whose XML the core copy reads incompletely.
 */
import { describe, expect, it } from 'vitest';
import { readBook } from '../../src/core/xlsx';
import { CT, fml, num, pkgOf, row, str, type PkgDef, type PartDef, type RelDef, type SheetDef } from './dropsheets-helpers';
import { lintObjects, lintObjectsText, objectsOf, openPkg } from './addsheets-objects-helpers';
import { premergeCopies, type FileIn } from './addsheets-helpers';

// ---------------------------------------------------------------------------------------------------------------------
// builders
// ---------------------------------------------------------------------------------------------------------------------

const BASE_REGION = { sheet: '主', headerRow: 1, firstDataRow: 2, lastDataRow: 3, firstCol: 1, lastCol: 2 };
const baseDef = (extra: Partial<PkgDef> = {}, sheet: Partial<SheetDef> = {}): PkgDef => ({
  sheets: [{ name: '主', data: row(1, str('A1', 'ID') + str('B1', '値')) + row(2, str('A2', 'x') + num('B2', 1)) + row(3, str('A3', 'y') + num('B3', 2)), ...sheet }],
  ...extra,
});
const base = async (extra: Partial<PkgDef> = {}, sheet: Partial<SheetDef> = {}): Promise<FileIn> => ({ bytes: await pkgOf(baseDef(extra, sheet)), fileName: '主.xlsx', region: BASE_REGION });

const XDR = 'xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"';
const wsDr = (...anchors: string[]) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr ${XDR}>${anchors.join('')}</xdr:wsDr>`;
const pos = (c: number, r: number) => `<xdr:from><xdr:col>${c}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${r}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>${c + 3}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${r + 4}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>`;
const chartAnchor = (id: number, rid: string, c = 4, r = 1) => `<xdr:twoCellAnchor>${pos(c, r)}<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="Chart ${id}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="${rid}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>`;
const picAnchor = (id: number, rid: string, c = 8, r = 1) => `<xdr:twoCellAnchor>${pos(c, r)}<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${id}" name="Picture ${id}"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:twoCellAnchor>`;
const spAnchor = (id: number, attrs = 'macro="" textlink=""', fill = '<a:solidFill><a:schemeClr val="accent1"/></a:solidFill>', extra = '') => `<xdr:twoCellAnchor>${pos(2, 8)}<xdr:sp ${attrs}><xdr:nvSpPr><xdr:cNvPr id="${id}" name="Shape ${id}">${extra}</xdr:cNvPr><xdr:cNvSpPr txBox="1"/></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${fill}</xdr:spPr><xdr:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="ja-JP"><a:latin typeface="+mn-lt"/></a:rPr><a:t>text ${id}</a:t></a:r></a:p></xdr:txBody></xdr:sp><xdr:clientData/></xdr:twoCellAnchor>`;
const slicerAnchor = (id: number) => `<xdr:twoCellAnchor>${pos(6, 12)}<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice xmlns:sle15="http://schemas.microsoft.com/office/drawing/2012/slicer" Requires="sle15"><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="Slicer ${id}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/drawing/2010/slicer"><sle:slicer xmlns:sle="http://schemas.microsoft.com/office/drawing/2010/slicer" name="S${id}"/></a:graphicData></a:graphic></xdr:graphicFrame></mc:Choice><mc:Fallback><xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="0" name=""/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1" cy="1"/></a:xfrm></xdr:spPr></xdr:sp></mc:Fallback></mc:AlternateContent><xdr:clientData/></xdr:twoCellAnchor>`;
const diagramAnchor = (id: number) => `<xdr:twoCellAnchor>${pos(0, 20)}<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="Diagram ${id}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" r:dm="rId7" r:lo="rId8" r:qs="rId9" r:cs="rId10"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>`;
const controlPlaceholder = (id: number) => `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" Requires="a14"><xdr:twoCellAnchor editAs="oneCell">${pos(0, 30)}<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="${id}" name="Check Box ${id}" hidden="1"><a:extLst><a:ext uri="{63B3BB69-23CF-44E3-9099-C40C66FF867C}"><a14:compatExt spid="_x0000_s${id}"/></a:ext></a:extLst></xdr:cNvPr><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm></xdr:spPr></xdr:sp><xdr:clientData/></xdr:twoCellAnchor></mc:Choice></mc:AlternateContent>`;

const CHART_NS = 'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const numRef = (f: string, pts: number[]) => `<c:numRef><c:f>${f}</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${pts.length}"/>${pts.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('')}</c:numCache></c:numRef>`;
const strRef = (f: string, pts: string[]) => `<c:strRef><c:f>${f}</c:f><c:strCache><c:ptCount val="${pts.length}"/>${pts.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('')}</c:strCache></c:strRef>`;
const ser = (i: number, tx: string, cat: string, val: string, sp = '') => `<c:ser><c:idx val="${i}"/><c:order val="${i}"/><c:tx>${tx}</c:tx>${sp}<c:cat>${cat}</c:cat><c:val>${val}</c:val></c:ser>`;
const chartXml = (sers: string, extra = '', title = '') => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<c:chartSpace ${CHART_NS}><c:date1904 val="0"/><c:chart>${title}<c:plotArea><c:layout/><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>${sers}<c:axId val="1"/><c:axId val="2"/></c:barChart><c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx></c:plotArea></c:chart>${extra}</c:chartSpace>`;
const plainChart = (sheet = 'データ') => chartXml(ser(0, strRef(`${sheet}!$B$1`, ['値']), strRef(`${sheet}!$A$2:$A$3`, ['a', 'b']), numRef(`${sheet}!$B$2:$B$3`, [1, 2])));

const dataSheet = (name: string, extra: Partial<SheetDef> = {}): SheetDef => ({ name, data: row(1, str('A1', 'ID') + str('B1', '値')) + row(2, str('A2', 'a') + num('B2', 1)) + row(3, str('A3', 'b') + num('B3', 2)), ...extra });
const drawingSheet = (name: string, drawingPath: string, tail = '', extraRels: RelDef[] = []): SheetDef => ({
  ...dataSheet(name),
  tail: `<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><drawing r:id="rId1"/>${tail}`,
  rels: [{ id: 'rId1', type: 'drawing', target: `../drawings/${drawingPath}` }, ...extraRels],
});
const PNG = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 0, 1, 0, 0, 0, 1, 8, 6, 0, 0, 0, 31, 21, 196, 137]);

async function copy(bytes: Uint8Array, name = 'obj.xlsx', b?: FileIn, extra: Record<string, boolean> = { dropOtherSheets: true }) {
  const res = await premergeCopies(b ?? (await base()), [{ bytes, fileName: name }], extra);
  const pkg = await openPkg(res.bytes);
  return { res, pkg, objs: await objectsOf(res.bytes) };
}
const limits = (res: Awaited<ReturnType<typeof copy>>['res']) => res.report.items.filter((i) => i.code === 'sheetCopyLimits').map((i) => i.message);

// ---------------------------------------------------------------------------------------------------------------------

describe('charts: references that cannot follow keep their values', () => {
  it('a workbook name, an external workbook and a missing sheet become literal data; the other series follow the renamed sheet', async () => {
    const chart = chartXml(
      ser(0, strRef('データ!$B$1', ['値']), strRef('データ!$A$2:$A$3', ['a', 'b']), numRef('データ!$B$2:$B$3', [1, 2])) +
      ser(1, strRef('[0]!NamedName', ['名前系列']), strRef('データ!$A$2:$A$3', ['a', 'b']), numRef('[0]!NamedValues', [5, 6])) +
      ser(2, strRef('データ!$B$1', ['値']), strRef('[1]Other!$A$2:$A$3', ['p', 'q']), numRef('[1]Other!$B$2:$B$3', [7, 8])) +
      ser(3, strRef('データ!$B$1', ['値']), strRef('消えた!$A$2:$A$3', ['x', 'y']), numRef('消えた!$B$2:$B$3', [9, 10])),
      '',
      `<c:title><c:tx>${strRef('[0]!TitleName', ['リンクした見出し'])}</c:tx><c:overlay val="0"/></c:title>`,
    );
    const bytes = await pkgOf({
      sheets: [drawingSheet('データ', 'drawing1.xml')],
      parts: [{ path: 'xl/drawings/drawing1.xml', xml: wsDr(chartAnchor(2, 'rId1')), ct: CT.drawing }, { path: 'xl/charts/chart1.xml', xml: chart, ct: CT.chart }],
      partRels: { 'xl/drawings/drawing1.xml': [{ id: 'rId1', type: 'chart', target: '../charts/chart1.xml' }] },
    });
    const { res, pkg } = await copy(bytes);
    expect(lintObjectsText(await lintObjects(res.bytes))).toBe('');
    const out = [...pkg.files.entries()].find(([n]) => /^xl\/charts\/chart\d+\.xml$/.test(n))![1];
    expect(out).toContain('<c:f>obj_データ!$B$2:$B$3</c:f>');
    // series 2: name + values literal, category follows
    expect(out).toContain('<c:tx><c:v>名前系列</c:v></c:tx>');
    expect(out).toContain('<c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="2"/><c:pt idx="0"><c:v>5</c:v></c:pt><c:pt idx="1"><c:v>6</c:v></c:pt></c:numLit>');
    expect(out).toContain('<c:strLit><c:ptCount val="2"/><c:pt idx="0"><c:v>p</c:v></c:pt><c:pt idx="1"><c:v>q</c:v></c:pt></c:strLit>');
    expect(out).toContain('<c:strLit><c:ptCount val="2"/><c:pt idx="0"><c:v>x</c:v></c:pt><c:pt idx="1"><c:v>y</c:v></c:pt></c:strLit>');
    // the title linked to a name becomes text
    expect(out).toContain('<c:rich><a:bodyPr/><a:p><a:r><a:t>リンクした見出し</a:t></a:r></a:p></c:rich>');
    expect(out).not.toMatch(/\[0\]!|\[1\]Other|消えた!/);
    const w = limits(res);
    expect(w).toHaveLength(1);
    expect(w[0]).toMatch(/グラフの参照先をコピーできなかった系列・見出し 7 個は、元の値のまま残しました/);
  });

  it('a series that reads a defined name of one plain range keeps reading the cells (the range, renamed); a name that computes stays literal data', async () => {
    const chart = chartXml(
      ser(0, strRef('[0]!Hdr', ['値']), strRef('データ!Cats', ['a', 'b']), numRef('[0]!Rng', [1, 2])) +
      ser(1, strRef('データ!$B$1', ['値']), strRef('データ!$A$2:$A$3', ['a', 'b']), numRef('[0]!Dyn', [5, 6])),
    );
    const bytes = await pkgOf({
      sheets: [drawingSheet('データ', 'drawing1.xml')],
      names: '<definedName name="Rng">データ!$B$2:$B$3</definedName><definedName name="Hdr">データ!$B$1</definedName><definedName name="Cats" localSheetId="0">データ!$A$2:$A$3</definedName><definedName name="Dyn">OFFSET(データ!$B$2,0,0,2,1)</definedName>',
      parts: [{ path: 'xl/drawings/drawing1.xml', xml: wsDr(chartAnchor(2, 'rId1')), ct: CT.drawing }, { path: 'xl/charts/chart1.xml', xml: chart, ct: CT.chart }],
      partRels: { 'xl/drawings/drawing1.xml': [{ id: 'rId1', type: 'chart', target: '../charts/chart1.xml' }] },
    });
    const { res, pkg } = await copy(bytes);
    expect(lintObjectsText(await lintObjects(res.bytes))).toBe('');
    const out = [...pkg.files.entries()].find(([n]) => /^xl\/charts\/chart\d+\.xml$/.test(n))![1];
    expect(out).toContain('<c:f>obj_データ!$B$2:$B$3</c:f>');
    expect(out).toContain('<c:f>obj_データ!$B$1</c:f>');
    // a name of the sheet itself is the copy's name of the same text (the core copy writes it): the reference only follows the sheet
    expect(out).toContain('<c:f>obj_データ!Cats</c:f>');
    expect(out).not.toMatch(/\[0\]!|Rng|Hdr/);
    // OFFSET(...) cannot be a plain reference: its cached values stay as literal data
    expect(out).toContain('<c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="2"/><c:pt idx="0"><c:v>5</c:v></c:pt><c:pt idx="1"><c:v>6</c:v></c:pt></c:numLit>');
    expect(limits(res)[0]).toMatch(/グラフの参照先をコピーできなかった系列・見出し 1 個は、元の値のまま残しました/);
  });

  it('a data-label range (c15 extension) that cannot follow is dropped with its extension, the chart stays', async () => {
    const c15 = (f: string) => `<c:extLst><c:ext uri="{02D57815-91ED-43cb-92C2-25804820EDAC}" xmlns:c15="http://schemas.microsoft.com/office/drawing/2012/chart"><c15:datalabelsRange><c15:f>${f}</c15:f></c15:datalabelsRange></c:ext></c:extLst>`;
    const chart = chartXml(
      ser(0, strRef('データ!$B$1', ['値']), strRef('データ!$A$2:$A$3', ['a', 'b']), numRef('データ!$B$2:$B$3', [1, 2])).replace('</c:ser>', `${c15('データ!$A$2:$A$3')}</c:ser>`) +
      ser(1, strRef('データ!$B$1', ['値']), strRef('データ!$A$2:$A$3', ['a', 'b']), numRef('データ!$B$2:$B$3', [1, 2])).replace('</c:ser>', `${c15('[1]X!$A$1:$A$2')}</c:ser>`),
    );
    const bytes = await pkgOf({
      sheets: [drawingSheet('データ', 'drawing1.xml')],
      parts: [{ path: 'xl/drawings/drawing1.xml', xml: wsDr(chartAnchor(2, 'rId1')), ct: CT.drawing }, { path: 'xl/charts/chart1.xml', xml: chart, ct: CT.chart }],
      partRels: { 'xl/drawings/drawing1.xml': [{ id: 'rId1', type: 'chart', target: '../charts/chart1.xml' }] },
    });
    const { res, pkg } = await copy(bytes);
    expect(lintObjectsText(await lintObjects(res.bytes))).toBe('');
    const out = [...pkg.files.entries()].find(([n]) => /^xl\/charts\/chart\d+\.xml$/.test(n))![1];
    expect(out).toContain('<c15:f>obj_データ!$A$2:$A$3</c15:f>');
    expect(out).not.toContain('[1]X');
    expect((out.match(/<c15:datalabelsRange>/g) ?? []).length).toBe(1);
  });

  it('a chart whose part is missing is left out and named; the other anchors of the drawing stay; nothing dangles', async () => {
    const bytes = await pkgOf({
      sheets: [drawingSheet('データ', 'drawing1.xml')],
      parts: [{ path: 'xl/drawings/drawing1.xml', xml: wsDr(chartAnchor(2, 'rId1'), spAnchor(3)), ct: CT.drawing }],
      partRels: { 'xl/drawings/drawing1.xml': [{ id: 'rId1', type: 'chart', target: '../charts/gone.xml' }] },
    });
    const { res, objs } = await copy(bytes);
    expect(lintObjectsText(await lintObjects(res.bytes))).toBe('');
    const o = objs.find((s) => s.name === 'obj_データ')!;
    expect(o.charts).toBe(0);
    expect(o.shapes).toBe(1);
    expect(limits(res)).toEqual([expect.stringMatching(/次のものはコピーできませんでした: グラフ 1 個/)]);
  });
});

describe('drawings: what is left out', () => {
  it('slicers, SmartArt, ink and the placeholders of form controls are left out; charts, pictures and shapes stay; the warning counts them', async () => {
    const ink = `<xdr:twoCellAnchor>${pos(0, 40)}<xdr:contentPart xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId11"/><xdr:clientData/></xdr:twoCellAnchor>`;
    const bytes = await pkgOf({
      sheets: [drawingSheet('データ', 'drawing1.xml')],
      parts: [
        { path: 'xl/drawings/drawing1.xml', xml: wsDr(chartAnchor(2, 'rId1'), picAnchor(3, 'rId2'), spAnchor(4), slicerAnchor(5), slicerAnchor(6), diagramAnchor(7), controlPlaceholder(1025), ink), ct: CT.drawing },
        { path: 'xl/charts/chart1.xml', xml: plainChart(), ct: CT.chart },
        { path: 'xl/media/image1.png', data: PNG },
      ],
      partRels: { 'xl/drawings/drawing1.xml': [{ id: 'rId1', type: 'chart', target: '../charts/chart1.xml' }, { id: 'rId2', type: 'image', target: '../media/image1.png' }] },
    });
    const { res, objs, pkg } = await copy(bytes);
    expect(lintObjectsText(await lintObjects(res.bytes))).toBe('');
    const o = objs.find((s) => s.name === 'obj_データ')!;
    expect([o.charts, o.pictures, o.shapes, o.slicers, o.controlPlaceholders]).toEqual([1, 1, 1, 0, 0]);
    const drawing = [...pkg.files.entries()].find(([n]) => /^xl\/drawings\/drawing\d+\.xml$/.test(n))![1];
    expect(drawing).not.toMatch(/slicer|diagram|compatExt|contentPart|dgm:/);
    const w = limits(res);
    expect(w).toHaveLength(1);
    expect(w[0]).toMatch(/次のものはコピーできませんでした: スライサー・タイムライン 2 個、SmartArt 1 個、その他の図 1 個/);
    expect(w[0]).toContain('シート「obj_データ」');
  });

  it('a drawing with nothing left over writes no drawing part and no <drawing> element', async () => {
    const bytes = await pkgOf({
      sheets: [drawingSheet('データ', 'drawing1.xml')],
      parts: [{ path: 'xl/drawings/drawing1.xml', xml: wsDr(slicerAnchor(5), diagramAnchor(7)), ct: CT.drawing }],
    });
    const { res, pkg } = await copy(bytes);
    expect(lintObjectsText(await lintObjects(res.bytes))).toBe('');
    expect([...pkg.files.keys()].filter((n) => /drawings\/drawing/.test(n))).toEqual([]);
    expect([...pkg.files.entries()].find(([n]) => /worksheets\/sheet2\.xml$/.test(n))![1]).not.toContain('<drawing');
  });

  it('pictures: one image part for the pictures of every sheet of the file that share it; macros are removed; textlink and sheet links follow the copies', async () => {
    const hl = (rid: string) => `<a:hlinkClick xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="${rid}"/>`;
    const sheetA = drawingSheet('データ', 'drawing1.xml');
    const sheetB = drawingSheet('別', 'drawing2.xml');
    const bytes = await pkgOf({
      sheets: [sheetA, sheetB],
      parts: [
        { path: 'xl/drawings/drawing1.xml', xml: wsDr(picAnchor(2, 'rId1'), spAnchor(3, 'macro="[0]!Button1_Click" textlink="データ!$A$2"', undefined, hl('rId2')), spAnchor(4, 'macro="" textlink="=\'別\'!B2"', undefined, hl('rId3'))), ct: CT.drawing },
        { path: 'xl/drawings/drawing2.xml', xml: wsDr(picAnchor(2, 'rId1'), spAnchor(3, 'macro="" textlink="消えた!A1"', undefined, hl('rId2'))), ct: CT.drawing },
        { path: 'xl/media/image1.png', data: PNG },
      ],
      partRels: {
        'xl/drawings/drawing1.xml': [{ id: 'rId1', type: 'image', target: '../media/image1.png' }, { id: 'rId2', type: 'hyperlink', target: "#'別'!A1" }, { id: 'rId3', type: 'hyperlink', target: 'https://example.com/?a=1&b=2', external: true }],
        'xl/drawings/drawing2.xml': [{ id: 'rId1', type: 'image', target: '../media/image1.png' }, { id: 'rId2', type: 'hyperlink', target: '#消えた!A1' }],
      },
    });
    const { res, pkg } = await copy(bytes);
    expect(lintObjectsText(await lintObjects(res.bytes))).toBe('');
    expect([...pkg.names].filter((n) => /^xl\/media\/.+/.test(n))).toHaveLength(1);
    const ds = [...pkg.files.entries()].filter(([n]) => /^xl\/drawings\/drawing\d+\.xml$/.test(n)).map(([, x]) => x);
    expect(ds[0]).toContain('macro=""');
    expect(ds[0]).not.toContain('Button1_Click');
    expect(ds[0]).toContain('textlink="obj_データ!$A$2"');
    expect(ds[0]).toMatch(/textlink="obj_別!B2"/);
    // a text link to a sheet that does not exist is cleared, the shape stays; its hyperlink to the missing sheet goes
    expect(ds[1]).toContain('textlink=""');
    expect(ds[1]).not.toContain('hlinkClick');
    const rels = [...pkg.files.entries()].filter(([n]) => /^xl\/drawings\/_rels\/drawing\d+\.xml\.rels$/.test(n)).map(([, x]) => x);
    expect(rels[0]).toContain('Target="#obj_別!A1"');
    expect(rels[0]).toContain('Target="https://example.com/?a=1&amp;b=2" TargetMode="External"');
    expect(rels[1]).not.toContain('消えた');
  });
});

describe('notes and form controls (one VML for both)', () => {
  const VML = `<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/></v:shapetype><v:shape id="_x0000_s1025" type="#_x0000_t202" style='position:absolute;width:96pt;height:60pt;z-index:1;visibility:hidden'><x:ClientData ObjectType="Note"><x:Anchor>2, 15, 0, 15, 3, 71, 3, 19</x:Anchor><x:Row>1</x:Row><x:Column>0</x:Column></x:ClientData></v:shape><v:shapetype id="_x0000_t201" coordsize="21600,21600" o:spt="201" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/></v:shapetype><v:shape id="_x0000_s1026" type="#_x0000_t201" style='position:absolute'><x:ClientData ObjectType="Checkbox"><x:FmlaLink>$E$1</x:FmlaLink></x:ClientData></v:shape><v:shape id="_x0000_s1027" type="#_x0000_t201" style='position:absolute'><x:ClientData ObjectType="Button"><x:Anchor>0, 26, 17, 1, 2, 2, 18, 14</x:Anchor></x:ClientData></v:shape></xml>`;
  const COMMENTS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors><author>me</author><author>you</author></authors><commentList><comment ref="A2" authorId="1" shapeId="0"><text><r><t>hello</t></r></text></comment></commentList></comments>`;
  const sheet = (): SheetDef => ({
    ...dataSheet('データ'),
    tail: '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><legacyDrawing r:id="rId1"/>',
    rels: [{ id: 'rId1', type: 'vmlDrawing', target: '../drawings/vmlDrawing1.vml' }, { id: 'rId2', type: 'comments', target: '../comments1.xml' }],
  });

  it('the VML keeps the note, drops the controls and their shapetype, is renumbered (idmap / ids); authors and anchors stay; the controls are named', async () => {
    const bytes = await pkgOf({ sheets: [sheet()], parts: [{ path: 'xl/drawings/vmlDrawing1.vml', xml: VML }, { path: 'xl/comments1.xml', xml: COMMENTS, ct: CT.comments }] });
    const { res, pkg } = await copy(bytes);
    expect(lintObjectsText(await lintObjects(res.bytes))).toBe('');
    const vml = [...pkg.files.entries()].find(([n]) => /vmlDrawing\d+\.vml$/.test(n))![1];
    expect(vml).toContain('ObjectType="Note"');
    expect(vml).not.toMatch(/Checkbox|Button|_x0000_t201/);
    expect(vml).toContain('<x:Anchor>2, 15, 0, 15, 3, 71, 3, 19</x:Anchor>');
    expect(vml).toMatch(/idmap v:ext="edit" data="1"/);
    expect(vml).toMatch(/id="_x0000_s1025"/);
    const c = [...pkg.files.entries()].find(([n]) => /^xl\/comments\d+\.xml$/.test(n))![1];
    expect(c).toContain('<author>you</author>');
    expect(c).toMatch(/<comment ref="A2" authorId="0"/);
    const w = limits(res);
    expect(w).toHaveLength(1);
    expect(w[0]).toMatch(/次のものはコピーできませんでした: フォームコントロール 2 個/);
    expect(w[0]).not.toMatch(/コメント/);
  });

  it('two copies with notes do not share VML ids (data="1" and data="2")', async () => {
    const mk = async () => pkgOf({ sheets: [sheet()], parts: [{ path: 'xl/drawings/vmlDrawing1.vml', xml: VML }, { path: 'xl/comments1.xml', xml: COMMENTS, ct: CT.comments }] });
    const res = await premergeCopies(await base(), [{ bytes: await mk(), fileName: 'a.xlsx' }, { bytes: await mk(), fileName: 'b.xlsx' }], { dropOtherSheets: true });
    expect(lintObjectsText(await lintObjects(res.bytes))).toBe('');
    const pkg = await openPkg(res.bytes);
    const vmls = [...pkg.files.entries()].filter(([n]) => /\.vml$/.test(n)).map(([, x]) => x);
    expect(vmls).toHaveLength(2);
    expect(vmls[0].match(/idmap[^>]*data="(\d+)"/)![1]).not.toBe(vmls[1].match(/idmap[^>]*data="(\d+)"/)![1]);
    expect(vmls[1]).toContain('_x0000_s2049');
  });

  it('comments without their VML are not copied and stay reported', async () => {
    const s = sheet();
    s.rels = [{ id: 'rId2', type: 'comments', target: '../comments1.xml' }];
    s.tail = '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>';
    const bytes = await pkgOf({ sheets: [s], parts: [{ path: 'xl/comments1.xml', xml: COMMENTS, ct: CT.comments }] });
    const { res, pkg } = await copy(bytes);
    expect(lintObjectsText(await lintObjects(res.bytes))).toBe('');
    expect([...pkg.files.keys()].filter((n) => /comments|vml/i.test(n))).toEqual([]);
    expect(limits(res)[0]).toMatch(/コメント 1 件/);
  });
});

describe('tables', () => {
  const T = (id: number, name: string, ref: string, extra = '') => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="${id}" name="${name}" displayName="${name}" ref="${ref}"${extra}><autoFilter ref="${ref}"/><tableColumns count="2"><tableColumn id="1" name="ID"/><tableColumn id="2" name="値"/></tableColumns><tableStyleInfo name="TableStyleMedium2" showRowStripes="1"/></table>`;
  const tableSheet = (name: string, tables: Array<{ id: number; name: string }>, formula?: string): SheetDef => ({
    ...dataSheet(name),
    data: row(1, str('A1', 'ID') + str('B1', '値') + str('D1', '計')) + row(2, str('A2', 'a') + num('B2', 1) + (formula ? fml('D2', formula, 3) : '')) + row(3, str('A3', 'b') + num('B3', 2)),
    tail: `<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><tableParts count="${tables.length}">${tables.map((_t, i) => `<tablePart r:id="rId${i + 1}"/>`).join('')}</tableParts>`,
    rels: tables.map((t, i) => ({ id: `rId${i + 1}`, type: 'table', target: `../tables/table${t.id}.xml` })),
  });

  it('a table keeps its name unless the base has it as a table or a defined name: 売上表 -> 売上表_2, Tn -> Tn_2; ids are workbook-unique; formulas follow', async () => {
    const baseFile = await base({
      names: '<definedName name="Tn">主!$A$1</definedName>',
      parts: [{ path: 'xl/tables/table1.xml', xml: T(5, '売上表', 'A1:B3'), ct: CT.table }],
    }, { tail: '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><tableParts count="1"><tablePart r:id="rId1"/></tableParts>', rels: [{ id: 'rId1', type: 'table', target: '../tables/table1.xml' }] });
    const bytes = await pkgOf({
      sheets: [tableSheet('データ', [{ id: 1, name: '売上表' }, { id: 2, name: 'Tn' }], 'SUM(売上表[値])+SUM(Tn[値])+ROWS(売上表)')],
      parts: [{ path: 'xl/tables/table1.xml', xml: T(1, '売上表', 'A1:B3'), ct: CT.table }, { path: 'xl/tables/table2.xml', xml: T(2, 'Tn', 'F1:G3'), ct: CT.table }],
    });
    const res = await premergeCopies(baseFile, [{ bytes, fileName: 'obj.xlsx' }], { dropOtherSheets: false });
    expect(lintObjectsText(await lintObjects(res.bytes))).toBe('');
    const objs = await objectsOf(res.bytes);
    expect(objs.flatMap((s) => s.tables.map((t) => `${s.name}:${t.name}`)).sort()).toEqual(['obj_データ:Tn_2', 'obj_データ:売上表_2', '主:売上表']);
    const pkg = await openPkg(res.bytes);
    const ids = [...pkg.files.entries()].filter(([n]) => /^xl\/tables\//.test(n)).map(([, x]) => /<table\b[^>]*\sid="(\d+)"/.exec(x)![1]);
    expect(new Set(ids).size).toBe(3);
    const out = await readBook(res.bytes, 'o.xlsx');
    expect(out.sheets.find((s) => s.name === 'obj_データ')!.rows[1]![3]!.f).toBe('SUM(売上表_2[値])+SUM(Tn_2[値])+ROWS(売上表_2)');
    expect(res.report.items.filter((i) => i.code === 'tableRenamed').map((i) => i.message).join(' ')).toMatch(/テーブル「売上表」は、名前が重なったため「売上表_2」にしました/);
    expect(limits(res)).toEqual([]);
  });

  it('query tables and XML-mapped tables stay as cells (their structured references became values) and are named', async () => {
    const q = T(1, 'Q1', 'A1:B3', ' tableType="queryTable" connectionId="1"');
    const x = T(2, 'X1', 'F1:G3', ' tableType="xml"');
    const bytes = await pkgOf({
      sheets: [tableSheet('データ', [{ id: 1, name: 'Q1' }, { id: 2, name: 'X1' }], 'SUM(Q1[値])')],
      parts: [{ path: 'xl/tables/table1.xml', xml: q, ct: CT.table }, { path: 'xl/tables/table2.xml', xml: x, ct: CT.table }],
    });
    const { res, pkg, objs } = await copy(bytes);
    expect(lintObjectsText(await lintObjects(res.bytes))).toBe('');
    expect(objs.find((s) => s.name === 'obj_データ')!.tables).toEqual([]);
    expect([...pkg.files.keys()].filter((n) => /tables\//.test(n))).toEqual([]);
    const m = limits(res).join('\n');
    expect(m).toMatch(/テーブル 2 個（外部データ接続・XML 対応のため、テーブルの機能は外し、セルの値と書式だけを残しました）/);
    expect(m).toMatch(/テーブル参照の数式 1 件は値にしました/);
  });

  it('calculated-column and totals formulas follow renamed sheets / tables; a formula that cannot follow is left out (the column is an ordinary column)', async () => {
    const t = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="T1" displayName="T1" ref="A1:B4" totalsRowCount="1"><autoFilter ref="A1:B3"/><tableColumns count="2"><tableColumn id="1" name="ID" totalsRowLabel="計"/><tableColumn id="2" name="値" totalsRowFunction="custom"><calculatedColumnFormula>データ!$D$1+T1[[#This Row],[ID]]</calculatedColumnFormula><totalsRowFormula>SUM(別!A1:A3)+[外]!X</totalsRowFormula></tableColumn></tableColumns><tableStyleInfo name="TableStyleMedium2"/></table>`;
    const bytes = await pkgOf({
      sheets: [tableSheet('データ', [{ id: 1, name: 'T1' }]), dataSheet('別')],
      parts: [{ path: 'xl/tables/table1.xml', xml: t, ct: CT.table }],
    });
    const { res, pkg } = await copy(bytes);
    expect(lintObjectsText(await lintObjects(res.bytes))).toBe('');
    const out = [...pkg.files.entries()].find(([n]) => /^xl\/tables\//.test(n))![1];
    expect(out).toContain('<calculatedColumnFormula>obj_データ!$D$1+T1[[#This Row],[ID]]</calculatedColumnFormula>');
    expect(out).not.toContain('totalsRowFormula');
    expect(out).toMatch(/totalsRowFunction="none"/);
  });
});

describe('formats of tables and pivot tables (dxf ids of the source styles.xml)', () => {
  const DXFS = '<dxfs count="3"><dxf><font><color rgb="FFFF0000"/></font></dxf><dxf><font><b/></font></dxf><dxf><fill><patternFill><bgColor rgb="FF00FF00"/></patternFill></fill></dxf></dxfs>';
  const styles = (): string => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="游ゴシック"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="標準" xfId="0" builtinId="0"/></cellStyles>${DXFS}<tableStyles count="0"/></styleSheet>`;

  it('a table keeps its header / data / totals formats: the dxfs are imported into the target and the ids follow; named cell styles of the table are left out', async () => {
    const t = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="表1" displayName="表1" ref="A1:B3" headerRowDxfId="2" dataDxfId="1" headerRowCellStyle="標準" dataCellStyle="標準"><autoFilter ref="A1:B3"/><tableColumns count="2"><tableColumn id="1" name="ID" dataDxfId="0"/><tableColumn id="2" name="値" dataDxfId="1" uniqueName="u"/></tableColumns><tableStyleInfo name="TableStyleMedium2"/></table>`;
    const s: SheetDef = { ...dataSheet('データ'), tail: '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><tableParts count="1"><tablePart r:id="rId1"/></tableParts>', rels: [{ id: 'rId1', type: 'table', target: '../tables/table1.xml' }] };
    // the base gets a dxf of its own first: the copy's ids start behind it
    const b = await base({ styles: styles().replace('FFFF0000', 'FF0000FF') });
    const src = await pkgOf({ sheets: [s], styles: styles(), parts: [{ path: 'xl/tables/table1.xml', xml: t, ct: CT.table }] });
    const { res, pkg } = await copy(src, 'obj.xlsx', b);
    expect(lintObjectsText(await lintObjects(res.bytes))).toBe('');
    const out = [...pkg.files.entries()].find(([n]) => /^xl\/tables\//.test(n))![1];
    const dxfs = [...pkg.files.get('xl/styles.xml')!.matchAll(/<dxf>[\s\S]*?<\/dxf>/g)].map((m) => m[0]);
    const at = (attr: string) => dxfs[Number(new RegExp(` ${attr}="(\\d+)"`).exec(out)![1])];
    expect(at('headerRowDxfId')).toContain('00FF00');
    expect(out).toMatch(/<table [^>]*dataDxfId="\d+"/);
    expect(at('dataDxfId')).toContain('<b/>');
    expect(dxfs[Number(/<tableColumn id="1" name="ID" dataDxfId="(\d+)"/.exec(out)![1])]).toContain('FFFF0000');
    expect(out).not.toMatch(/CellStyle|uniqueName/);
  });

  it('a pivot table keeps its formats (dxfs imported); a format whose dxf does not exist is dropped with its element, the count follows', async () => {
    const pivot = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="P" cacheId="1" dataCaption="値"><location ref="D1:E3" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/><pivotFields count="2"><pivotField axis="axisRow" showAll="0"><items count="3"><item x="0"/><item x="1"/><item t="default"/></items></pivotField><pivotField dataField="1" showAll="0"/></pivotFields><rowFields count="1"><field x="0"/></rowFields><rowItems count="3"><i><x/></i><i><x v="1"/></i><i t="grand"><x/></i></rowItems><colItems count="1"><i/></colItems><dataFields count="1"><dataField name="合計" fld="1" baseField="0" baseItem="0"/></dataFields><formats count="2"><format dxfId="2"><pivotArea outline="0" collapsedLevelsAreSubtotals="1" fieldPosition="0"/></format><format dxfId="9"><pivotArea outline="0" fieldPosition="0"/></format></formats><pivotTableStyleInfo name="PivotStyleLight16"/></pivotTableDefinition>`;
    const cache = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" recordCount="2"><cacheSource type="worksheet"><worksheetSource ref="A1:B3" sheet="データ"/></cacheSource><cacheFields count="2"><cacheField name="ID" numFmtId="0"><sharedItems count="2"><s v="a"/><s v="b"/></sharedItems></cacheField><cacheField name="値" numFmtId="0"><sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1" containsInteger="1" minValue="1" maxValue="2"/></cacheField></cacheFields></pivotCacheDefinition>`;
    const s: SheetDef = { ...dataSheet('データ'), tail: '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>', rels: [{ id: 'rId1', type: 'pivotTable', target: '../pivotTables/pivotTable1.xml' }] };
    const src = await pkgOf({
      sheets: [s], styles: styles(),
      parts: [{ path: 'xl/pivotTables/pivotTable1.xml', xml: pivot, ct: CT.pivotTable }, { path: 'xl/pivotCache/pivotCacheDefinition1.xml', xml: cache, ct: CT.pivotCacheDef }],
      partRels: { 'xl/pivotTables/pivotTable1.xml': [{ id: 'rId1', type: 'pivotCacheDefinition', target: '../pivotCache/pivotCacheDefinition1.xml' }] },
      wbRels: [{ id: 'rId50', type: 'pivotCacheDefinition', target: 'pivotCache/pivotCacheDefinition1.xml' }],
      wbExtra: '<pivotCaches><pivotCache cacheId="1" r:id="rId50"/></pivotCaches>',
    });
    const { res, pkg } = await copy(src);
    expect(lintObjectsText(await lintObjects(res.bytes))).toBe('');
    const px = [...pkg.files.entries()].find(([n]) => /^xl\/pivotTables\/pivotTable\d+\.xml$/.test(n))![1];
    expect(px).toMatch(/<formats count="1"><format dxfId="\d+">/);
    expect((px.match(/<format /g) ?? []).length).toBe(1);
    // the cache has no records part: the copy asks Excel to refresh it when it is opened
    expect([...pkg.files.entries()].find(([n]) => /pivotCacheDefinition\d+\.xml$/.test(n))![1]).toMatch(/refreshOnLoad="1"/);
  });
});

describe('a chart picture fill whose image is gone', () => {
  it('the fill loses its image reference, the chart is copied, nothing dangles', async () => {
    const chart = chartXml(ser(0, strRef('データ!$B$1', ['値']), strRef('データ!$A$2:$A$3', ['a', 'b']), numRef('データ!$B$2:$B$3', [1, 2]), '<c:spPr><a:blipFill><a:blip r:embed="rId5"/><a:stretch><a:fillRect/></a:stretch></a:blipFill></c:spPr>'));
    const bytes = await pkgOf({
      sheets: [drawingSheet('データ', 'drawing1.xml')],
      parts: [{ path: 'xl/drawings/drawing1.xml', xml: wsDr(chartAnchor(2, 'rId1')), ct: CT.drawing }, { path: 'xl/charts/chart1.xml', xml: chart, ct: CT.chart }],
      partRels: { 'xl/drawings/drawing1.xml': [{ id: 'rId1', type: 'chart', target: '../charts/chart1.xml' }], 'xl/charts/chart1.xml': [{ id: 'rId5', type: 'image', target: '../media/gone.png' }] },
    });
    const { res, objs } = await copy(bytes);
    expect(lintObjectsText(await lintObjects(res.bytes))).toBe('');
    expect(objs.find((s) => s.name === 'obj_データ')!.charts).toBe(1);
  });
});

describe('pivot tables', () => {
  const CACHE = (src: string, extra = '') => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1" refreshedBy="x" refreshedDate="45000" createdVersion="8" refreshedVersion="8" recordCount="2"${extra}>${src}<cacheFields count="2"><cacheField name="ID" numFmtId="0"><sharedItems count="2"><s v="a"/><s v="b"/></sharedItems></cacheField><cacheField name="値" numFmtId="179"><sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1" containsInteger="1" minValue="1" maxValue="2"/></cacheField></cacheFields></pivotCacheDefinition>`;
  const RECORDS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2"><r><x v="0"/><n v="1"/></r><r><x v="1"/><n v="2"/></r></pivotCacheRecords>`;
  const PIVOT = (name: string, cacheId: number, extra = '') => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="${name}" cacheId="${cacheId}" dataCaption="値"><location ref="D1:E3" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/><pivotFields count="2"><pivotField axis="axisRow" showAll="0"><items count="3"><item x="0"/><item x="1"/><item t="default"/></items></pivotField><pivotField dataField="1" numFmtId="179" showAll="0"/></pivotFields><rowFields count="1"><field x="0"/></rowFields><rowItems count="3"><i><x/></i><i><x v="1"/></i><i t="grand"><x/></i></rowItems><colItems count="1"><i/></colItems><dataFields count="1"><dataField name="合計" fld="1" baseField="0" baseItem="0" numFmtId="179"/></dataFields>${extra}<pivotTableStyleInfo name="PivotStyleLight16"/></pivotTableDefinition>`;
  const STYLES = (fmt: string) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="179" formatCode="${fmt}"/></numFmts><fonts count="1"><font><sz val="11"/><name val="游ゴシック"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="標準" xfId="0" builtinId="0"/></cellStyles><dxfs count="1"><dxf><font><b/></font></dxf></dxfs><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/></styleSheet>`;
  const pivotSheet = (name: string, pivots: string[]): SheetDef => ({
    ...dataSheet(name),
    tail: '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>',
    rels: pivots.map((p, i) => ({ id: `rId${i + 1}`, type: 'pivotTable', target: `../pivotTables/${p}` })),
  });
  const def = (source: string, pivots: Array<{ name: string; cacheId: number }>, extraCache = ''): PkgDef => ({
    sheets: [pivotSheet('データ', pivots.map((_p, i) => `pivotTable${i + 1}.xml`))],
    styles: STYLES('0&quot;歳&quot;'),
    parts: [
      ...pivots.map((p, i): PartDef => ({ path: `xl/pivotTables/pivotTable${i + 1}.xml`, xml: PIVOT(p.name, p.cacheId), ct: CT.pivotTable })),
      { path: 'xl/pivotCache/pivotCacheDefinition1.xml', xml: CACHE(source, extraCache), ct: CT.pivotCacheDef },
      { path: 'xl/pivotCache/pivotCacheRecords1.xml', xml: RECORDS, ct: CT.pivotCacheRec },
    ],
    partRels: {
      ...Object.fromEntries(pivots.map((_p, i) => [`xl/pivotTables/pivotTable${i + 1}.xml`, [{ id: 'rId1', type: 'pivotCacheDefinition', target: '../pivotCache/pivotCacheDefinition1.xml' }]])),
      'xl/pivotCache/pivotCacheDefinition1.xml': [{ id: 'rId1', type: 'pivotCacheRecords', target: 'pivotCacheRecords1.xml' }],
    },
    wbRels: [{ id: 'rId50', type: 'pivotCacheDefinition', target: 'pivotCache/pivotCacheDefinition1.xml' }],
    wbExtra: `<pivotCaches><pivotCache cacheId="${pivots[0].cacheId}" r:id="rId50"/></pivotCaches>`,
  });

  it('a range source: sheet name, number format ids, cache id and pivot name follow the target (the base has cacheId 7 and a pivot of the same name)', async () => {
    // a base that has a pivot table of its own (cacheId 7, name P) on its data sheet
    const b = await pkgOf({ ...def('<cacheSource type="worksheet"><worksheetSource ref="A1:B3" sheet="主"/></cacheSource>', [{ name: 'P', cacheId: 7 }]), sheets: [{ ...pivotSheet('主', ['pivotTable1.xml']), data: row(1, str('A1', 'ID') + str('B1', '値')) + row(2, str('A2', 'a') + num('B2', 1)) + row(3, str('A3', 'b') + num('B3', 2)) }] });
    const src = await pkgOf(def('<cacheSource type="worksheet"><worksheetSource ref="A1:B3" sheet="データ"/></cacheSource>', [{ name: 'P', cacheId: 3 }, { name: 'Q', cacheId: 3 }]));
    const res = await premergeCopies({ bytes: b, fileName: '主.xlsx', region: BASE_REGION }, [{ bytes: src, fileName: 'obj.xlsx' }], { dropOtherSheets: false });
    expect(lintObjectsText(await lintObjects(res.bytes))).toBe('');
    const objs = await objectsOf(res.bytes);
    const copy1 = objs.find((s) => s.name === 'obj_データ')!;
    expect(copy1.pivots.map((p) => p.name)).toEqual(['P_2', 'Q']);
    expect(copy1.pivots.map((p) => p.cacheId)).toEqual(['8', '8']);
    expect(copy1.pivots.map((p) => p.source)).toEqual(['obj_データ!A1:B3', 'obj_データ!A1:B3']);
    expect(limits(res)).toEqual([]);
    const pkg = await openPkg(res.bytes);
    // number formats: the custom format 179 of the source is the target's id for "0歳" (the base styles define 179 as the same text here: found, not duplicated)
    const px = [...pkg.files.entries()].filter(([n]) => /^xl\/pivotTables\/pivotTable\d+\.xml$/.test(n)).map(([, x]) => x);
    const styles = pkg.files.get('xl/styles.xml')!;
    for (const x of px) for (const m of x.matchAll(/numFmtId="(\d+)"/g)) expect(styles).toContain(`numFmtId="${m[1]}"`);
  });

  it.each([
    ['external data (connection)', '<cacheSource type="external" connectionId="1"/>'],
    ['several ranges (consolidation)', '<cacheSource type="consolidation"><consolidation><pages count="0"/></consolidation></cacheSource>'],
    ['a defined name', '<cacheSource type="worksheet"><worksheetSource name="MyData"/></cacheSource>'],
    ['a sheet that is not there', '<cacheSource type="worksheet"><worksheetSource ref="A1:B3" sheet="消えた"/></cacheSource>'],
    ['a sheet of another workbook', '<cacheSource type="worksheet"><worksheetSource ref="A1:B3" sheet="x" r:id="rId9"/></cacheSource>'],
  ])('a cache that reads %s: the pivot stays as cell values, named in the warning, nothing is written for it', async (_label, source) => {
    const src = await pkgOf(def(source, [{ name: 'P', cacheId: 3 }]));
    const { res, pkg, objs } = await copy(src);
    expect(lintObjectsText(await lintObjects(res.bytes))).toBe('');
    expect(objs.find((s) => s.name === 'obj_データ')!.pivots).toEqual([]);
    expect([...pkg.files.keys()].filter((n) => /pivot/i.test(n))).toEqual([]);
    expect(pkg.files.get('xl/workbook.xml')).not.toContain('pivotCache');
    expect(limits(res)).toEqual([expect.stringMatching(/ピボットテーブル 1 個（元のデータを参照できないため、セルの値だけを残しました）/)]);
    // the cells are there
    expect((await readBook(res.bytes, 'o.xlsx')).sheets.find((s) => s.name === 'obj_データ')!.rows[1]![1]!.v).toBe(1);
  });
});

describe('pivot charts', () => {
  const pivotChart = (pivot: string) => chartXml(
    ser(0, strRef('データ!$B$1', ['値']), strRef('データ!$A$2:$A$3', ['a', 'b']), numRef('データ!$B$2:$B$3', [1, 2])),
    '',
  ).replace('<c:chart>', `<c:pivotSource><c:name>[b.xlsx]${pivot}</c:name><c:fmtId val="0"/></c:pivotSource><c:chart><c:pivotFmts><c:pivotFmt><c:idx val="0"/></c:pivotFmt></c:pivotFmts>`);
  const PIVOT = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="P1" cacheId="1" dataCaption="値"><location ref="D1:E3" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/><pivotFields count="1"><pivotField dataField="1" showAll="0"/></pivotFields><dataFields count="1"><dataField name="合計" fld="0" baseField="0" baseItem="0"/></dataFields></pivotTableDefinition>`;
  const cache = (source: string) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" recordCount="0">${source}<cacheFields count="1"><cacheField name="値" numFmtId="0"><sharedItems/></cacheField></cacheFields></pivotCacheDefinition>`;
  const make = async (source: string, chartPivot: string) => pkgOf({
    sheets: [{
      ...drawingSheet('データ', 'drawing1.xml'),
      rels: [{ id: 'rId1', type: 'drawing', target: '../drawings/drawing1.xml' }, { id: 'rId2', type: 'pivotTable', target: '../pivotTables/pivotTable1.xml' }],
    }],
    parts: [
      { path: 'xl/drawings/drawing1.xml', xml: wsDr(chartAnchor(2, 'rId1')), ct: CT.drawing },
      { path: 'xl/charts/chart1.xml', xml: pivotChart(chartPivot), ct: CT.chart },
      { path: 'xl/pivotTables/pivotTable1.xml', xml: PIVOT, ct: CT.pivotTable },
      { path: 'xl/pivotCache/pivotCacheDefinition1.xml', xml: cache(source), ct: CT.pivotCacheDef },
    ],
    partRels: {
      'xl/drawings/drawing1.xml': [{ id: 'rId1', type: 'chart', target: '../charts/chart1.xml' }],
      'xl/pivotTables/pivotTable1.xml': [{ id: 'rId1', type: 'pivotCacheDefinition', target: '../pivotCache/pivotCacheDefinition1.xml' }],
    },
    wbRels: [{ id: 'rId50', type: 'pivotCacheDefinition', target: 'pivotCache/pivotCacheDefinition1.xml' }],
    wbExtra: '<pivotCaches><pivotCache cacheId="1" r:id="rId50"/></pivotCaches>',
  });

  it('a pivot chart follows its pivot table copy: sheet and pivot names in <c:pivotSource> are the copy\'s', async () => {
    const { res, pkg } = await copy(await make('<cacheSource type="worksheet"><worksheetSource ref="A1:B3" sheet="データ"/></cacheSource>', 'データ!P1'));
    expect(lintObjectsText(await lintObjects(res.bytes))).toBe('');
    const c = [...pkg.files.entries()].find(([n]) => /^xl\/charts\/chart\d+\.xml$/.test(n))![1];
    expect(c).toContain('<c:name>[b.xlsx]obj_データ!P1</c:name>');
    expect(c).toContain('<c:pivotFmts>');
  });

  it('a pivot chart whose pivot table stays as cell values is an ordinary chart: no pivot source, no pivot formats', async () => {
    const { res, pkg, objs } = await copy(await make('<cacheSource type="external" connectionId="1"/>', 'データ!P1'));
    expect(lintObjectsText(await lintObjects(res.bytes))).toBe('');
    const c = [...pkg.files.entries()].find(([n]) => /^xl\/charts\/chart\d+\.xml$/.test(n))![1];
    expect(c).not.toMatch(/pivotSource|pivotFmts/);
    expect(objs.find((s) => s.name === 'obj_データ')!.charts).toBe(1);
    expect(limits(res)[0]).toMatch(/ピボットテーブル 1 個/);
  });
});

describe('themes: charts and shapes keep the colours / fonts of their source', () => {
  it('a source with another theme: schemeClr becomes the source colour, theme fonts become the font names; a pie chart ref stays', async () => {
    const theme = (acc1: string, minor: string) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="t"><a:themeElements><a:clrScheme name="t"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="${acc1}"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme><a:fontScheme name="t"><a:majorFont><a:latin typeface="Georgia"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="${minor}"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="t"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>`;
    const sp = `<c:spPr><a:solidFill><a:schemeClr val="accent1"><a:lumMod val="60000"/><a:lumOff val="40000"/></a:schemeClr></a:solidFill><a:ln><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:ln></c:spPr>`;
    const chart = chartXml(ser(0, strRef('データ!$B$1', ['値']), strRef('データ!$A$2:$A$3', ['a', 'b']), numRef('データ!$B$2:$B$3', [1, 2]), sp), '<c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr><a:latin typeface="+mn-lt"/></a:defRPr></a:pPr><a:endParaRPr lang="ja-JP"/></a:p></c:txPr>');
    const bytes = await pkgOf({
      sheets: [drawingSheet('データ', 'drawing1.xml')],
      parts: [{ path: 'xl/drawings/drawing1.xml', xml: wsDr(chartAnchor(2, 'rId1'), spAnchor(3, 'macro="" textlink=""', '<a:solidFill><a:schemeClr val="accent1"/></a:solidFill><a:ln><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>')), ct: CT.drawing }, { path: 'xl/charts/chart1.xml', xml: chart, ct: CT.chart }],
      partRels: { 'xl/drawings/drawing1.xml': [{ id: 'rId1', type: 'chart', target: '../charts/chart1.xml' }] },
    });
    // the source's theme is replaced by editing the package afterwards (pkgOf has a fixed empty theme: the reader falls back to the default Office palette)
    const JSZip = (await import('jszip')).default;
    const z = await JSZip.loadAsync(bytes);
    z.file('xl/theme/theme1.xml', theme('C0504D', 'Meiryo UI'));
    const src = await z.generateAsync({ type: 'uint8array' });
    // a base with another theme (accent1 = 4472C4, minor font Calibri)
    const bz = await JSZip.loadAsync(await pkgOf(baseDef()));
    bz.file('xl/theme/theme1.xml', theme('4472C4', 'Calibri'));
    const b: FileIn = { bytes: await bz.generateAsync({ type: 'uint8array' }), fileName: '主.xlsx', region: BASE_REGION };
    const { res, pkg } = await copy(src, 'obj.xlsx', b);
    expect(lintObjectsText(await lintObjects(res.bytes))).toBe('');
    const c = [...pkg.files.entries()].find(([n]) => /^xl\/charts\/chart\d+\.xml$/.test(n))![1];
    expect(c).toContain('<a:srgbClr val="C0504D"><a:lumMod val="60000"/><a:lumOff val="40000"/></a:srgbClr>');
    expect(c).toContain('<a:srgbClr val="000000"/>');
    expect(c).toContain('<a:latin typeface="Meiryo UI"/>');
    expect(c).not.toContain('schemeClr');
    const d = [...pkg.files.entries()].find(([n]) => /^xl\/drawings\/drawing\d+\.xml$/.test(n))![1];
    expect(d).toContain('<a:srgbClr val="C0504D"/>');
    expect(d).toContain('<a:schemeClr val="phClr"/>'); // placeholder colours stay
    expect(d).toContain('<a:latin typeface="Meiryo UI"/>');
    // the same themes: nothing is rewritten
    const same = await copy(await pkgOf({
      sheets: [drawingSheet('データ', 'drawing1.xml')],
      parts: [{ path: 'xl/drawings/drawing1.xml', xml: wsDr(chartAnchor(2, 'rId1'), spAnchor(3)), ct: CT.drawing }, { path: 'xl/charts/chart1.xml', xml: chart, ct: CT.chart }],
      partRels: { 'xl/drawings/drawing1.xml': [{ id: 'rId1', type: 'chart', target: '../charts/chart1.xml' }] },
    }));
    const c2 = [...same.pkg.files.entries()].find(([n]) => /^xl\/charts\/chart\d+\.xml$/.test(n))![1];
    expect(c2).toContain('<a:schemeClr val="accent1">');
    expect(c2).toContain('<a:latin typeface="+mn-lt"/>');
  });
});

describe('new-style charts (chartEx), chart style parts, embedded workbooks, chart sheets', () => {
  it('a chartEx chart follows the renamed sheet (cx:f); one whose formula cannot follow is left out and named', async () => {
    const CX = 'xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"';
    const cx = (f: string) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<cx:chartSpace ${CX} xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><cx:chartData><cx:data id="0"><cx:strDim type="cat"><cx:f>${f}</cx:f></cx:strDim></cx:data></cx:chartData><cx:chart><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="funnel" uniqueId="{1}"><cx:dataId val="0"/></cx:series></cx:plotAreaRegion></cx:plotArea></cx:chart></cx:chartSpace>`;
    const anchor = (id: number, rid: string, r: number) => `<xdr:twoCellAnchor>${pos(4, r)}<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="cx1" xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex"><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="Chart ${id}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/drawing/2014/chartex"><cx:chart ${CX} r:id="${rid}"/></a:graphicData></a:graphic></xdr:graphicFrame></mc:Choice><mc:Fallback><xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="0" name=""/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1" cy="1"/></a:xfrm></xdr:spPr></xdr:sp></mc:Fallback></mc:AlternateContent><xdr:clientData/></xdr:twoCellAnchor>`;
    const bytes = await pkgOf({
      sheets: [drawingSheet('データ', 'drawing1.xml')],
      parts: [
        { path: 'xl/drawings/drawing1.xml', xml: wsDr(anchor(2, 'rId1', 1), anchor(3, 'rId2', 10)), ct: CT.drawing },
        { path: 'xl/charts/chartEx1.xml', xml: cx('データ!$A$2:$A$3'), ct: 'application/vnd.ms-office.chartex+xml' },
        { path: 'xl/charts/chartEx2.xml', xml: cx('[1]X!$A$2:$A$3'), ct: 'application/vnd.ms-office.chartex+xml' },
      ],
      partRels: { 'xl/drawings/drawing1.xml': [{ id: 'rId1', type: 'http://schemas.microsoft.com/office/2014/relationships/chartEx', target: '../charts/chartEx1.xml' }, { id: 'rId2', type: 'http://schemas.microsoft.com/office/2014/relationships/chartEx', target: '../charts/chartEx2.xml' }] },
    });
    const { res, pkg } = await copy(bytes);
    expect(lintObjectsText(await lintObjects(res.bytes))).toBe('');
    const ex = [...pkg.files.entries()].filter(([n]) => /^xl\/charts\/chartEx\d+\.xml$/.test(n));
    expect(ex).toHaveLength(1);
    expect(ex[0][1]).toContain('<cx:f>obj_データ!$A$2:$A$3</cx:f>');
    expect(limits(res)).toEqual([expect.stringMatching(/次のものはコピーできませんでした: グラフ 1 個/)]);
  });

  it('chart style / colour parts, an embedded workbook and user shapes of a chart are copied with it (once per chart, own names)', async () => {
    const chart = (rid: string) => chartXml(ser(0, strRef('データ!$B$1', ['値']), strRef('データ!$A$2:$A$3', ['a', 'b']), numRef('データ!$B$2:$B$3', [1, 2])), `<c:externalData r:id="${rid}"><c:autoUpdate val="0"/></c:externalData><c:userShapes r:id="rId9"/>`);
    const bytes = await pkgOf({
      sheets: [drawingSheet('データ', 'drawing1.xml')],
      parts: [
        { path: 'xl/drawings/drawing1.xml', xml: wsDr(chartAnchor(2, 'rId1'), chartAnchor(3, 'rId2', 4, 10)), ct: CT.drawing },
        { path: 'xl/charts/chart1.xml', xml: chart('rId1'), ct: CT.chart },
        { path: 'xl/charts/chart2.xml', xml: chart('rId1'), ct: CT.chart },
        { path: 'xl/charts/style1.xml', xml: '<cs:chartStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" id="201"/>', ct: 'application/vnd.ms-office.chartstyle+xml' },
        { path: 'xl/charts/colors1.xml', xml: '<cs:colorStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" meth="cycle" id="10"/>', ct: 'application/vnd.ms-office.chartcolorstyle+xml' },
        { path: 'xl/drawings/drawing9.xml', xml: '<c:userShapes xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>', ct: 'application/vnd.openxmlformats-officedocument.drawingml.chartshapes+xml' },
        { path: 'xl/embeddings/Microsoft_Excel_Sheet1.xlsx', data: new Uint8Array([80, 75, 3, 4]) },
      ],
      partRels: {
        'xl/drawings/drawing1.xml': [{ id: 'rId1', type: 'chart', target: '../charts/chart1.xml' }, { id: 'rId2', type: 'chart', target: '../charts/chart2.xml' }],
        'xl/charts/chart1.xml': [{ id: 'rId1', type: 'package', target: '../embeddings/Microsoft_Excel_Sheet1.xlsx' }, { id: 'rId9', type: 'chartUserShapes', target: '../drawings/drawing9.xml' }, { id: 'rId5', type: 'http://schemas.microsoft.com/office/2011/relationships/chartStyle', target: 'style1.xml' }, { id: 'rId6', type: 'http://schemas.microsoft.com/office/2011/relationships/chartColorStyle', target: 'colors1.xml' }],
        'xl/charts/chart2.xml': [{ id: 'rId1', type: 'package', target: '../embeddings/Microsoft_Excel_Sheet1.xlsx' }, { id: 'rId9', type: 'chartUserShapes', target: '../drawings/drawing9.xml' }],
      },
    });
    const { res, pkg } = await copy(bytes);
    expect(lintObjectsText(await lintObjects(res.bytes))).toBe('');
    expect(pkg.names.filter((n) => /^xl\/charts\/(style|colors)\d+\.xml$/.test(n)).sort()).toEqual(['xl/charts/colors1.xml', 'xl/charts/style1.xml']);
    expect(pkg.names.filter((n) => /^xl\/embeddings\/.+/.test(n))).toHaveLength(2);
    expect(pkg.names.filter((n) => /^xl\/drawings\/drawing\d+\.xml$/.test(n))).toHaveLength(3); // the sheet's + 2 user-shape parts
    expect(limits(res)).toEqual([]);
  });

  it('a chart sheet is copied as a chart sheet and takes a name like the worksheets (the naming rule sees it); one without a usable chart becomes an empty sheet with a warning', async () => {
    const CS = (drawingRid: boolean) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetPr codeName="Chart1"/><sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>${drawingRid ? '<drawing r:id="rId1"/>' : ''}</chartsheet>`;
    const abs = `<xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="9000000" cy="6000000"/><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId1"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:absoluteAnchor>`;
    const bytes = await pkgOf({
      sheets: [
        dataSheet('データ'),
        { name: 'グラフ', kind: 'chartsheet', xml: CS(true), rels: [{ id: 'rId1', type: 'drawing', target: '../drawings/drawing1.xml' }] },
        { name: '空のグラフ', kind: 'chartsheet', xml: CS(false) },
        { name: 'グラフ2', kind: 'chartsheet', xml: CS(true), rels: [{ id: 'rId1', type: 'drawing', target: '../drawings/drawing2.xml' }] },
      ],
      parts: [
        { path: 'xl/drawings/drawing1.xml', xml: wsDr(abs), ct: CT.drawing },
        { path: 'xl/drawings/drawing2.xml', xml: wsDr(slicerAnchor(4)), ct: CT.drawing },
        { path: 'xl/charts/chart1.xml', xml: plainChart(), ct: CT.chart },
      ],
      partRels: { 'xl/drawings/drawing1.xml': [{ id: 'rId1', type: 'chart', target: '../charts/chart1.xml' }] },
    });
    const { res, pkg, objs } = await copy(bytes);
    expect(lintObjectsText(await lintObjects(res.bytes))).toBe('');
    expect(res.layout.addedSheets!.map((s) => s.name)).toEqual(['obj_データ', 'obj_グラフ', 'obj_空のグラフ', 'obj_グラフ2']);
    expect(objs.map((s) => `${s.name}:${s.kind}:${s.charts}`)).toEqual(['主:worksheet:0', 'obj_データ:worksheet:0', 'obj_グラフ:chartsheet:1', 'obj_空のグラフ:worksheet:0', 'obj_グラフ2:worksheet:0']);
    const cs = [...pkg.files.entries()].find(([n]) => /^xl\/chartsheets\/sheet\d+\.xml$/.test(n))![1];
    expect(cs).not.toMatch(/tabSelected|codeName/);
    expect(cs).toMatch(/<drawing r:id="rId1"\/><\/chartsheet>/);
    const w = limits(res);
    expect(w).toHaveLength(2);
    expect(w.join('\n')).toMatch(/「空のグラフ」はコピーできませんでした。.*「obj_空のグラフ」は空のシートになります/);
    expect(w.join('\n')).toMatch(/「グラフ2」はコピーできませんでした。.*「obj_グラフ2」は空のシートになります/);
  });
});

describe('sheets whose XML the core copy reads incompletely', () => {
  it('a table on a sheet with form controls (nested mc:AlternateContent before <tableParts>) is still copied', async () => {
    const t = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="T1" displayName="T1" ref="A1:B3"><autoFilter ref="A1:B3"/><tableColumns count="2"><tableColumn id="1" name="ID"/><tableColumn id="2" name="値"/></tableColumns><tableStyleInfo name="TableStyleMedium2"/></table>`;
    const ctl = (n: number) => `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="x14"><control shapeId="${2049 + n}" r:id="rId${3 + n}" name="Check Box ${n}"><controlPr defaultSize="0"/></control></mc:Choice></mc:AlternateContent>`;
    const s: SheetDef = {
      ...dataSheet('データ'),
      tail: `<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="x14"><controls>${ctl(1)}${ctl(2)}</controls></mc:Choice></mc:AlternateContent><tableParts count="1"><tablePart r:id="rId1"/></tableParts>`,
      rels: [{ id: 'rId1', type: 'table', target: '../tables/table1.xml' }],
    };
    const { res, objs } = await copy(await pkgOf({ sheets: [s], parts: [{ path: 'xl/tables/table1.xml', xml: t, ct: CT.table }] }));
    expect(lintObjectsText(await lintObjects(res.bytes))).toBe('');
    expect(objs.find((o) => o.name === 'obj_データ')!.tables.map((x) => x.name)).toEqual(['T1']);
  });
});

describe('known defect of the core copy (A1, copysheet.ts splitTop): reported to the lead, not fixed here', () => {
  // splitTop finds the end of an element by the first closing tag of its own name: <controls> sits in <mc:AlternateContent> and holds one <mc:AlternateContent> per control,
  // so everything after the first control (<tableParts>, <extLst> with x14 conditional formats / validations / sparklines / slicer lists) is read as nothing.
  // The table is copied nevertheless (the object phase reads the sheet XML itself); the extension list is lost. `it.fails` flips to a failure when the defect is fixed:
  // then turn it into a plain `it`.
  it.fails('the extension list (x14 conditional format) after the form controls of a sheet is kept in the copy', async () => {
    const ctl = (n: number) => `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="x14"><control shapeId="${2049 + n}" r:id="rId${3 + n}" name="Check Box ${n}"><controlPr defaultSize="0"/></control></mc:Choice></mc:AlternateContent>`;
    const ext = '<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:conditionalFormattings><x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x14:cfRule type="dataBar" id="{00000000-0000-0000-0000-000000000001}"><x14:dataBar minLength="0" maxLength="100"><x14:cfvo type="autoMin"/><x14:cfvo type="autoMax"/></x14:dataBar></x14:cfRule><xm:sqref>B2:B3</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst>';
    const s: SheetDef = {
      ...dataSheet('データ'),
      tail: `<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="x14"><controls>${ctl(1)}${ctl(2)}</controls></mc:Choice></mc:AlternateContent>${ext}`,
    };
    const { pkg } = await copy(await pkgOf({ sheets: [s] }));
    expect([...pkg.files.entries()].find(([n]) => /worksheets\/sheet2\.xml$/.test(n))![1]).toContain('<x14:cfRule');
  });
});
