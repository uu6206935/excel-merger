/**
 * Sheet features the Excel goldens do not contain: page breaks, ignored errors, protected ranges, sparklines,
 * autofilter filter columns / sort state, hyperlinks of deleted cells, pivot table locations, dataTable refs,
 * form-control links, split panes.
 */
import { describe, expect, it } from 'vitest';
import type { LoggedOp, StructOp } from '../../src/core/types';
import { replayLog } from '../../src/core/xlsx/replay';
import { buildXlsx, part } from '../golden/build';

const st = (op: StructOp): LoggedOp => ({ kind: 'struct', sheet: 'S', op });
const data = Array.from({ length: 10 }, (_x, i) => `<row r="${i + 1}">${'ABCDEFGH'.split('').map((c, j) => `<c r="${c}${i + 1}"><v>${(i + 1) * 10 + j}</v></c>`).join('')}</row>`).join('');

async function book(tail: string, head = ''): Promise<Uint8Array> {
  return buildXlsx([{ name: 'S', data, tail, head }, { name: 'O', data: '<row r="1"><c r="A1"><v>1</v></c></row>' }]);
}
async function sheetXml(bytes: Uint8Array): Promise<string> {
  return (await part(bytes, 'xl/worksheets/sheet1.xml'))!;
}

describe('sheet features under structural ops', () => {
  it('page breaks follow their rows / columns and disappear with them', async () => {
    const tail = '<rowBreaks count="2" manualBreakCount="2"><brk id="3" max="16383" man="1"/><brk id="7" max="16383" man="1"/></rowBreaks>' +
      '<colBreaks count="1" manualBreakCount="1"><brk id="4" max="1048575" man="1"/></colBreaks>';
    const x1 = await sheetXml(await replayLog(await book(tail), [st({ type: 'insert', axis: 'row', at: 2, count: 2 }), st({ type: 'delete', axis: 'col', at: 5, count: 1 })]));
    expect(x1).toContain('<brk id="5" max="16383" man="1"/><brk id="9" max="16383" man="1"/>');
    // a break belongs to the row / column that starts the next page: deleting column E removes the break after D
    expect(x1).not.toContain('<colBreaks');
    const x2 = await sheetXml(await replayLog(await book(tail), [st({ type: 'delete', axis: 'row', at: 4, count: 1 })]));
    // deleting row 4 (which starts the page after the break at 3) removes that break; the other one shifts
    expect(x2).toContain('<rowBreaks count="1" manualBreakCount="1"><brk id="6" max="16383" man="1"/></rowBreaks>');
  });

  it('ignoredErrors / protectedRanges sqrefs are adjusted and dropped when deleted', async () => {
    const tail = '<protectedRanges><protectedRange name="p1" sqref="B2:C5"/><protectedRange name="p2" sqref="E2"/></protectedRanges>' +
      '<ignoredErrors><ignoredError sqref="C3:C9 E1" numberStoredAsText="1"/></ignoredErrors>';
    const x = await sheetXml(await replayLog(await book(tail), [st({ type: 'delete', axis: 'col', at: 5, count: 1 }), st({ type: 'insert', axis: 'col', at: 1, count: 1 })]));
    expect(x).toContain('<protectedRanges><protectedRange name="p1" sqref="C2:D5"/></protectedRanges>');
    expect(x).toContain('<ignoredError sqref="D3:D9" numberStoredAsText="1"/>');
  });

  it('autoFilter: ref, filterColumn ids (relative to the filter) and sortState follow the columns', async () => {
    const tail = '<autoFilter ref="B1:F10"><filterColumn colId="1"><filters><filter val="1"/></filters></filterColumn><filterColumn colId="3"><filters><filter val="2"/></filters></filterColumn>' +
      '<sortState ref="B2:F10"><sortCondition ref="E2:E10"/></sortState></autoFilter>';
    const x = await sheetXml(await replayLog(await book(tail), [st({ type: 'move', axis: 'col', from: 5, count: 1, to: 2 })]));
    // E moved before B: like any range (R2b R10) the filter becomes C1:F10 (the block lands left of it);
    // E's filter column leaves the range, old C (colId 1) is now D = index 1 of C..F
    expect(x).toMatch(/<autoFilter ref="C1:F10"><filterColumn colId="1"><filters><filter val="1"\/>/);
    expect(x).not.toContain('<filter val="2"/>');
    expect(x).toContain('<sortState ref="C2:F10"><sortCondition ref="B2:B10"/>');
    const z = await sheetXml(await replayLog(await book(tail), [st({ type: 'move', axis: 'col', from: 5, count: 1, to: 3 })]));
    // E moved before C, inside the filter: ids follow their columns (C -> D = 2, E -> C = 1)
    expect(z).toMatch(/<autoFilter ref="B1:F10"><filterColumn colId="2"><filters><filter val="1"\/><\/filters><\/filterColumn><filterColumn colId="1">/);
    const y = await sheetXml(await replayLog(await book(tail), [st({ type: 'delete', axis: 'col', at: 3, count: 1 })]));
    expect(y).toMatch(/<autoFilter ref="B1:E10"><filterColumn colId="2">/);
    expect(y).not.toContain('colId="1"');
  });

  it('hyperlinks follow their cells; a deleted cell loses its hyperlink; locations are formulas', async () => {
    const tail = '<hyperlinks><hyperlink ref="C3" location="S!F5" display="x"/><hyperlink ref="D4" location="O!A1"/></hyperlinks>';
    const x = await sheetXml(await replayLog(await book(tail), [st({ type: 'delete', axis: 'col', at: 4, count: 1 }), st({ type: 'insert', axis: 'col', at: 1, count: 1 })]));
    expect(x).toContain('<hyperlinks><hyperlink ref="D3" location="S!F5" display="x"/></hyperlinks>');
  });

  it('sparklines: data ranges are formulas, locations follow their cells', async () => {
    const tail = '<extLst><ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x14:sparklineGroup displayEmptyCellsAs="gap"><x14:colorSeries rgb="FF376092"/><x14:sparklines><x14:sparkline><xm:f>S!A2:D2</xm:f><xm:sqref>H2</xm:sqref></x14:sparkline><x14:sparkline><xm:f>S!A3:D3</xm:f><xm:sqref>H3</xm:sqref></x14:sparkline></x14:sparklines></x14:sparklineGroup></x14:sparklineGroups></ext></extLst>';
    const x = await sheetXml(await replayLog(await book(tail), [st({ type: 'insert', axis: 'col', at: 2, count: 1 }), st({ type: 'delete', axis: 'row', at: 3, count: 1 })]));
    expect(x).toContain('<x14:sparkline><xm:f>S!A2:E2</xm:f><xm:sqref>I2</xm:sqref></x14:sparkline>');
    expect(x).not.toContain('S!A3:E3');
    const y = await sheetXml(await replayLog(await book(tail), [st({ type: 'delete', axis: 'row', at: 2, count: 2 })]));
    expect(y).not.toContain('sparklineGroups'); // every sparkline gone: group and container removed
  });

  it('split (not frozen) panes keep their split; frozen panes follow the insert / delete steps', async () => {
    const head = '';
    const split = await buildXlsx([{ name: 'S', data, head }]);
    const withPane = (await part(split, 'xl/worksheets/sheet1.xml'))!;
    expect(withPane).toContain('<sheetView workbookViewId="0"/>');
    const frozen = await buildXlsx([{ name: 'S', data }]);
    // inject a frozen pane through the builder's sheetViews: use a tail-less custom XML instead
    const JSZip = (await import('jszip')).default;
    const z = await JSZip.loadAsync(frozen);
    const sx = (await z.file('xl/worksheets/sheet1.xml')!.async('string')).replace('<sheetView workbookViewId="0"/>',
      '<sheetView workbookViewId="0"><pane xSplit="3" ySplit="1" topLeftCell="D2" activePane="bottomRight" state="frozen"/><selection pane="topRight" activeCell="D1" sqref="D1"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/><selection pane="bottomRight" activeCell="E5" sqref="E5"/></sheetView>');
    z.file('xl/worksheets/sheet1.xml', sx);
    const fb = await z.generateAsync({ type: 'uint8array' });
    const a = await sheetXml(await replayLog(fb, [st({ type: 'insert', axis: 'col', at: 2, count: 2 })]));
    expect(a).toMatch(/<pane xSplit="5" ySplit="1" topLeftCell="F2" activePane="bottomRight" state="frozen"\/>/);
    const b = await sheetXml(await replayLog(fb, [st({ type: 'delete', axis: 'col', at: 3, count: 2 })]));
    expect(b).toMatch(/<pane xSplit="2" ySplit="1" topLeftCell="C2"/);
    const c = await sheetXml(await replayLog(fb, [st({ type: 'move', axis: 'col', from: 5, count: 1, to: 2 })]));
    expect(c).toMatch(/<pane xSplit="4" ySplit="1" topLeftCell="E2"/); // insert step inside the frozen block
    const d = await sheetXml(await replayLog(fb, [st({ type: 'move', axis: 'col', from: 6, count: 2, to: 8 })]));
    expect(d).toMatch(/<pane xSplit="3" ySplit="1" topLeftCell="D2"/); // unchanged
    expect(d).toContain('activeCell="E5" sqref="E5"');
  });

  it('pivot table locations on the op sheet, dataTable refs and form-control links', async () => {
    const pivot = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="P" cacheId="1"><location ref="J3:K8" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/></pivotTableDefinition>';
    const ctrl = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<formControlPr xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" objectType="CheckBox" fmlaLink="$C$2" lockText="1"/>';
    const bytes = await buildXlsx([{
      name: 'S',
      data: data + '<row r="12"><c r="B12"><f t="dataTable" ref="B12:C13" dt2D="0" dtr="0" r1="A1"/><v>1</v></c></row>',
      rels: [
        { id: 'rId7', type: 'pivotTable', path: 'xl/pivotTables/pivotTable1.xml', xml: pivot, contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml' },
        { id: 'rId8', type: 'ctrlProp', path: 'xl/ctrlProps/ctrlProp1.xml', xml: ctrl, contentType: 'application/vnd.ms-excel.controlproperties+xml' },
      ],
    }]);
    const out = await replayLog(bytes, [st({ type: 'insert', axis: 'col', at: 1, count: 1 })]);
    expect(await part(out, 'xl/pivotTables/pivotTable1.xml')).toContain('<location ref="K3:L8"');
    expect(await part(out, 'xl/ctrlProps/ctrlProp1.xml')).toContain('fmlaLink="$D$2"');
    expect(await sheetXml(out)).toContain('<f t="dataTable" ref="C12:D13" dt2D="0" dtr="0" r1="B1"/>');
  });
});
