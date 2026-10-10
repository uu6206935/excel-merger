/**
 * The "kitchen sink" source workbook of the sheet-copier tests: every worksheet feature the core copy has to carry over, written by hand.
 *
 *  Data      the busy sheet: frozen panes, tab colour, zoom, hidden / outlined columns, column + row styles, merges, autoFilter + sortState, all
 *            cell types (rich text, inline strings, booleans, errors, dates, cached formula results, shared / array / dynamic-array
 *            formulas), formulas into other sheets, CF (classic + x14 data bar), DV (lists into other sheets, names, dates), hyperlinks
 *            (external + internal), page setup with a printer-settings relationship, extension list (x14, sparklines, unknown)
 *  Calc      the sheet Data reads (names Rate / Chain / Choices) and a sheet-scoped name that hides a workbook name (Shadow)
 *  Bad       formulas that cannot be copied as formulas (external, 3D, call-out, uncopied table, unusable name) next to ones that can
 *  Q'uote    a sheet name with an apostrophe and a space
 *  Hidden    state hidden, VHidden: veryHidden
 */
import {
  OTHER_COLORS, buildPackage, boolc, cn, errc, fml, grid, inl, num, row, sst, themeXml, type PkgSpec, type SheetSpec,
  MAIN_NS,
} from './addsheets-build';

export const KITCHEN_NAME = 'キッチン.xlsx';

const NS_MC = 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac xr" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision"';
const X14 = 'xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"';

export const KITCHEN_STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="${MAIN_NS}"><numFmts count="2"><numFmt numFmtId="164" formatCode="0.0&quot;pt&quot;"/><numFmt numFmtId="165" formatCode="yyyy&quot;年&quot;m&quot;月&quot;"/></numFmts><fonts count="4"><font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font><font><b/><sz val="11"/><color theme="5" tint="0.39997558519241921"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font><font><i/><sz val="11"/><color rgb="FFFF0000"/><name val="Calibri"/><family val="2"/></font><font><sz val="14"/><color indexed="10"/><name val="Arial"/><family val="2"/></font></fonts><fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor theme="4" tint="0.59999389629810485"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="3"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color theme="1"/></left><right style="thin"><color theme="1"/></right><top style="thin"><color theme="1"/></top><bottom style="thin"><color theme="1"/></bottom><diagonal/></border><border><left/><right/><top/><bottom style="thick"><color rgb="FF0000FF"/></bottom><diagonal/></border></borders><cellStyleXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" applyFont="1"/></cellStyleXfs><cellXfs count="9"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1" applyProtection="1"><alignment horizontal="center" vertical="center" wrapText="1"/><protection locked="0"/></xf><xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="10" fontId="2" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="1" applyFont="1"/><xf numFmtId="4" fontId="3" fillId="3" borderId="2" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"/></cellXfs><cellStyles count="2"><cellStyle name="Normal" xfId="0" builtinId="0"/><cellStyle name="Title" xfId="1" builtinId="15"/></cellStyles><dxfs count="3"><dxf><font><color rgb="FFFF0000"/></font><fill><patternFill><bgColor rgb="FFFFFF00"/></patternFill></fill></dxf><dxf><font><b/><color theme="5"/></font><numFmt numFmtId="166" formatCode="0.00%"/></dxf><dxf><border><left style="thin"><color theme="1"/></left></border></dxf></dxfs><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/></styleSheet>`;

const RICH_TARO = '<r><rPr><b/><sz val="11"/><color theme="5"/><rFont val="Calibri"/><family val="2"/><scheme val="minor"/></rPr><t>太</t></r><r><rPr><sz val="11"/><color rgb="FF0000FF"/><rFont val="Calibri"/><family val="2"/></rPr><t xml:space="preserve"> 郎</t></r><rPh sb="0" eb="1"><t>タ</t></rPh><phoneticPr fontId="0" type="Hiragana"/>';

export const KITCHEN_SST: Array<string | { xml: string }> = [
  'ID', '氏名', { xml: RICH_TARO }, 'A001', 'A002', 'A003', '花子', 'Score', 'Date', 'Rate', 'Flag', 'Err', 'Sum', 'Cross', 'Shared', 'Arr', 'Dyn', 'Inline', '_x000D_ cr', ' lead',
];

const METADATA = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<metadata xmlns="${MAIN_NS}" xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray"><metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/></metadataTypes><futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata><cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>`;

function dataSheet(): SheetSpec {
  const hdr = ['ID', '氏名', 'Score', 'Date', 'Rate', 'Flag', 'Err', 'Sum', 'Cross', 'Shared', 'Arr', 'Dyn', 'Inline'];
  const rows: string[] = [];
  rows.push(row(1, hdr.map((_h, i) => sst(`${cn(i + 1)}1`, [0, 1, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17][i], 3)).join(''), 'ht="24" customHeight="1"'));
  for (let r = 2; r <= 11; r++) {
    const cells = [
      sst(`A${r}`, 3 + ((r - 2) % 3)),
      r === 2 ? sst(`B${r}`, 2) : r === 3 ? sst(`B${r}`, 6, 7) : inl(`B${r}`, `名${r}`),
      num(`C${r}`, r * 10, 1),
      num(`D${r}`, 44000 + r, r % 2 ? 5 : 2),
      num(`E${r}`, r / 100, 6),
      boolc(`F${r}`, r % 2 === 0),
      r === 2 ? fml(`G${r}`, 'NA()', '#N/A', 'e') : r === 3 ? errc(`G${r}`, '#DIV/0!') : '',
      r === 2 ? fml(`H${r}`, 'SUM($C$2:$C$11)', 650, 'n', 8) : '',
      fml(`I${r}`, `Calc!B${r}+Lookup!A2`, r * 10 + 7 + 2, 'n'),
      r === 2 ? fml(`J${r}`, 'C2*2', 40, 'n', 0, 't="shared" ref="J2:J11" si="0"') : `<c r="J${r}"><f t="shared" si="0"/><v>${r * 20}</v></c>`,
      r === 2 ? fml(`K${r}`, 'C2:C3*10', 200, 'n', 0, 't="array" ref="K2:K3"') : r === 3 ? '<c r="K3"><v>300</v></c>' : '',
      r === 2 ? `<c r="L2" cm="1"><f t="array" ref="L2:L4">_xlfn.SEQUENCE(3)</f><v>1</v></c>` : r === 3 ? '<c r="L3"><v>2</v></c>' : r === 4 ? '<c r="L4"><v>3</v></c>' : '',
      inl(`M${r}`, r === 2 ? '  spaced  ' : `m${r}`),
    ];
    rows.push(row(r, cells.join(''), r === 5 ? 'ht="30" customHeight="1" hidden="1"' : r === 6 ? 'outlineLevel="1"' : r === 7 ? 's="7" customFormat="1"' : ''));
  }
  rows.push(row(12, `<c r="A12" s="4" t="inlineStr"><is><t>merged title</t></is></c><c r="B12" s="4"/><c r="C12" s="4"/>`));
  const tail = [
    '<sheetProtection sheet="1" objects="1" scenarios="1"/>',
    '<autoFilter ref="A1:M11"><filterColumn colId="2"><customFilters><customFilter operator="greaterThan" val="25"/></customFilters></filterColumn><sortState ref="A2:M11"><sortCondition ref="C2:C11"/></sortState></autoFilter>',
    '<mergeCells count="1"><mergeCell ref="A12:C12"/></mergeCells>',
    '<phoneticPr fontId="0" type="Hiragana"/>',
    '<conditionalFormatting sqref="C2:C11"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>50</formula></cfRule><cfRule type="colorScale" priority="2"><colorScale><cfvo type="min"/><cfvo type="max"/><color theme="4"/><color rgb="FF63BE7B"/></colorScale></cfRule></conditionalFormatting>',
    '<conditionalFormatting sqref="E2:E11"><cfRule type="dataBar" priority="3"><dataBar><cfvo type="min"/><cfvo type="max"/><color theme="5"/></dataBar><extLst><ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:id>{11111111-2222-4333-8444-555555555555}</x14:id></ext></extLst></cfRule></conditionalFormatting>',
    '<conditionalFormatting sqref="D2:D11"><cfRule type="expression" dxfId="1" priority="4"><formula>Calc!$A$1&gt;C2</formula></cfRule><cfRule type="top10" dxfId="2" priority="5" rank="3"/></conditionalFormatting>',
    '<conditionalFormatting sqref="B2:B11"><cfRule type="iconSet" priority="6"><iconSet iconSet="3Arrows"><cfvo type="percent" val="0"/><cfvo type="percent" val="33"/><cfvo type="percent" val="67"/></iconSet></cfRule></conditionalFormatting>',
    '<dataValidations count="5">' +
      '<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A2:A11" xr:uid="{AAAAAAAA-0000-4000-8000-000000000001}"><formula1>Calc!$A$1:$A$5</formula1></dataValidation>' +
      '<dataValidation type="list" allowBlank="1" showErrorMessage="1" sqref="B2:B11"><formula1>"a,b,c"</formula1></dataValidation>' +
      '<dataValidation type="whole" operator="between" allowBlank="1" showErrorMessage="1" sqref="C2:C11"><formula1>1</formula1><formula2>100</formula2></dataValidation>' +
      '<dataValidation type="date" operator="between" allowBlank="1" showErrorMessage="1" sqref="D2:D11"><formula1>43831</formula1><formula2>44196</formula2></dataValidation>' +
      '<dataValidation type="list" allowBlank="1" showErrorMessage="1" sqref="M2:M11"><formula1>Choices</formula1></dataValidation>' +
    '</dataValidations>',
    '<hyperlinks><hyperlink ref="A2" r:id="rIdH1" tooltip="site"/><hyperlink ref="A3" location="Calc!B2" display="to calc"/><hyperlink ref="A4" location="Lookup!A1"/></hyperlinks>',
    '<printOptions horizontalCentered="1"/><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>',
    '<pageSetup paperSize="9" orientation="landscape" r:id="rIdP1"/>',
    '<headerFooter><oddHeader>&amp;CKitchen</oddHeader></headerFooter>',
    '<rowBreaks count="1" manualBreakCount="1"><brk id="8" max="16383" man="1"/></rowBreaks>',
    '<ignoredErrors><ignoredError sqref="I2:I11" formula="1"/></ignoredErrors>',
    `<extLst>` +
      `<ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" ${X14}><x14:conditionalFormattings><x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x14:cfRule type="dataBar" id="{11111111-2222-4333-8444-555555555555}"><x14:dataBar minLength="0" maxLength="100" border="1" negativeBarBorderColorSameAsPositive="0"><x14:cfvo type="autoMin"/><x14:cfvo type="autoMax"/><x14:borderColor theme="5"/><x14:negativeFillColor rgb="FFFF0000"/><x14:negativeBorderColor rgb="FFFF0000"/><x14:axisColor rgb="FF000000"/></x14:dataBar></x14:cfRule><xm:sqref>E2:E11</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext>` +
      `<ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}" ${X14}><x14:dataValidations count="1"><x14:dataValidation type="list" allowBlank="1" showErrorMessage="1"><x14:formula1><xm:f>Calc!$B$1:$B$3</xm:f></x14:formula1><xm:sqref>F2:F11</xm:sqref></x14:dataValidation></x14:dataValidations></ext>` +
      `<ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}" ${X14}><x14:sparklineGroups><x14:sparklineGroup displayEmptyCellsAs="gap"><x14:colorSeries theme="4" tint="-0.499984740745262"/><x14:colorNegative theme="5"/><x14:colorAxis rgb="FF000000"/><x14:colorMarkers theme="4" tint="-0.499984740745262"/><x14:colorFirst theme="4" tint="0.39997558519241921"/><x14:colorLast theme="4" tint="0.39997558519241921"/><x14:colorHigh theme="4"/><x14:colorLow theme="4"/><x14:sparklines><x14:sparkline><xm:f>Calc!A2:C2</xm:f><xm:sqref>N2</xm:sqref></x14:sparkline><x14:sparkline><xm:f>Data!C2:C11</xm:f><xm:sqref>N3</xm:sqref></x14:sparkline></x14:sparklines></x14:sparklineGroup></x14:sparklineGroups></ext>` +
      `<ext uri="{AAAAAAAA-BBBB-4CCC-8DDD-EEEEEEEEEEEE}" ${X14}><x14:unknownThing/></ext>` +
    `</extLst>`,
  ].join('');
  return {
    name: 'Data',
    rootAttrs: `${NS_MC} xr:uid="{00000000-0001-0000-0000-000000000000}"`,
    pr: '<sheetPr codeName="Sheet1" filterMode="0"><tabColor theme="5"/><outlinePr summaryBelow="0"/><pageSetUpPr fitToPage="1"/></sheetPr>',
    views: '<sheetView tabSelected="1" zoomScale="85" zoomScaleNormal="85" showGridLines="0" workbookViewId="0"><pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/><selection pane="topRight" activeCell="B1" sqref="B1"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/><selection pane="bottomRight" activeCell="B2" sqref="B2"/><pivotSelection pane="bottomRight" showHeader="1" activeRow="2" activeCol="1" click="1"><pivotArea type="data" outline="0" fieldPosition="0"/></pivotSelection></sheetView>',
    head: '<cols><col min="1" max="1" width="10.5" style="3" customWidth="1"/><col min="2" max="3" width="14" customWidth="1"/><col min="4" max="4" width="9" hidden="1" customWidth="1"/><col min="5" max="5" width="12" outlineLevel="1" customWidth="1"/></cols>',
    data: rows.join(''),
    tail,
    rels: [
      { id: 'rIdH1', type: 'hyperlink', target: 'https://example.com/a?b=1&c=2', external: true },
      { id: 'rIdP1', type: 'printerSettings', target: '../printerSettings/printerSettings1.bin' },
    ],
  };
}

export async function kitchenSource(over: Partial<PkgSpec> = {}): Promise<Uint8Array> {
  const calc: SheetSpec = {
    name: 'Calc',
    data: [
      row(1, num('A1', 10) + fml('B1', 'Rate*2', 20) + fml('C1', 'Chain+1', 41)),
      row(2, num('A2', 20) + fml('B2', 'Data!C2+1', 21) + fml('C2', 'Shadow', 100)),
      row(3, num('A3', 30) + num('B3', 3) + fml('C3', "'Q''uote Sheet'!A1+1", 6)),
      row(4, num('A4', 40) + num('B4', 4)),
      row(5, num('A5', 50) + num('B5', 5)),
      row(6, num('A6', 60) + num('B6', 6)),
    ].join(''),
  };
  const bad: SheetSpec = {
    name: 'Bad',
    data: [
      row(1, fml('A1', '[1]Ext!A1', 5) + fml('B1', 'SUM(Data:Calc!A1)', 111) + fml('C1', 'HYPERLINK("#Data!A1","go")', 'go', 's') + fml('D1', 'INDIRECT("Calc!A1")', 10)),
      row(2, fml('A2', 'WEBSERVICE("http://example.com")', 'web', 's') + fml('B2', 'Table1[Col]', 7) + fml('C2', 'Rate*3', 30) + fml('D2', 'ExtName+1', 6)),
      row(3, fml('A3', 'Gone!A1', 9) + fml('B3', "'Q''uote Sheet'!A1", 1) + fml('C3', 'SUM(Calc!A1:A3)', 60) + fml('D3', 'Local+1', 2)),
      row(4, fml('A4', 'MacroFn+1', 4) + fml('B4', 'Chain', 40) + fml('C4', 'Unused2', 1)),
    ].join(''),
  };
  const quote: SheetSpec = { name: "Q'uote Sheet", data: row(1, num('A1', 5) + fml('B1', "Calc!A1+A1", 15)) };
  const hidden: SheetSpec = { name: 'Hidden', state: 'hidden', data: grid([['h', 1], ['i', 2]]) };
  const vhidden: SheetSpec = { name: 'VHidden', state: 'veryHidden', data: grid([['v', 1]]) };
  const lookup: SheetSpec = { name: 'Lookup', data: row(1, inl('A1', 'lookup')) + row(2, num('A2', 7) + fml('B2', 'Data!C2*Rate', 200)) };
  const names = [
    '<definedName name="Rate">Calc!$A$1</definedName>',
    '<definedName name="Chain">Rate*2</definedName>',
    '<definedName name="Choices">Calc!$A$1:$A$5</definedName>',
    '<definedName name="Local" localSheetId="0">Data!$C$2:$C$11</definedName>',
    '<definedName name="_xlnm.Print_Area" localSheetId="0">Data!$A$1:$H$12</definedName>',
    '<definedName name="_xlnm.Print_Titles" localSheetId="0">Data!$1:$1</definedName>',
    '<definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">Data!$A$1:$M$11</definedName>',
    '<definedName name="ExtName">[1]Ext!$A$1</definedName>',
    '<definedName name="MacroFn" function="1" vbProcedure="0">Calc!$A$1</definedName>',
    '<definedName name="Unused">Calc!$C$3</definedName>',
    '<definedName name="Orphan">Calc!$C$4</definedName>',
    '<definedName name="Shadow">Calc!$A$1</definedName>',
    '<definedName name="Shadow" localSheetId="1">Calc!$A$2</definedName>',
    '<definedName name="Unused2">Unused</definedName>',
  ].join('');
  return buildPackage({
    sheets: [dataSheet(), calc, bad, lookup, quote, hidden, vhidden],
    sst: KITCHEN_SST,
    styles: KITCHEN_STYLES,
    theme: themeXml(OTHER_COLORS, 'Cambria', 'Calibri', 'Other'),
    names,
    metadata: METADATA,
    parts: [{ path: 'xl/printerSettings/printerSettings1.bin', xml: 'PRINTER', ct: 'application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings' }],
    app: true,
    ...over,
  });
}
