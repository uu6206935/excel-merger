// Generates tests/fixtures/xlsx/apply-probe.xlsx: a hand-written package whose cellXfs exercise the applyX
// attributes against a non-default parent cell style (xfId=1). Excel's rendering of these cells is recorded by
// com/record-probes.ps1 into excel-probes.json (section "apply"), which the reader must reproduce.
// Run: node tests/fixtures/xlsx/gen/make-apply-probe.mjs
import JSZip from 'jszip';
import { writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const here = dirname(fileURLToPath(import.meta.url));
const out = join(here, '..', 'apply-probe.xlsx');

const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const RNS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

const styles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="${NS}">
<numFmts count="1"><numFmt numFmtId="164" formatCode="0.000"/></numFmts>
<fonts count="2">
<font><sz val="11"/><color theme="1"/><name val="游ゴシック"/><family val="3"/><charset val="128"/></font>
<font><b/><sz val="14"/><color rgb="FFFF0000"/><name val="Arial"/><family val="2"/></font>
</fonts>
<fills count="4">
<fill><patternFill patternType="none"/></fill>
<fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF00FF00"/><bgColor indexed="64"/></patternFill></fill>
</fills>
<borders count="2">
<border><left/><right/><top/><bottom/><diagonal/></border>
<border><left style="thin"><color rgb="FF0000FF"/></left><right style="thin"><color rgb="FF0000FF"/></right><top style="thin"><color rgb="FF0000FF"/></top><bottom style="thin"><color rgb="FF0000FF"/></bottom><diagonal/></border>
</borders>
<cellStyleXfs count="2">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
<xf numFmtId="14" fontId="1" fillId="2" borderId="1"><alignment horizontal="center"/></xf>
</cellStyleXfs>
<cellXfs count="7">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="1"/>
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="1" applyNumberFormat="0" applyFont="0" applyFill="0" applyBorder="0" applyAlignment="0"/>
<xf numFmtId="164" fontId="1" fillId="3" borderId="1" xfId="0" applyNumberFormat="0" applyFont="0" applyFill="0" applyBorder="0" applyAlignment="0"><alignment horizontal="right"/></xf>
<xf numFmtId="164" fontId="1" fillId="3" borderId="1" xfId="0"><alignment horizontal="right"/></xf>
<xf numFmtId="4" fontId="1" fillId="3" borderId="0" xfId="1" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left"/></xf>
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="1" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"/>
</cellXfs>
<cellStyles count="2"><cellStyle name="Custom" xfId="1"/><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`;

let rows = '';
for (let i = 0; i < 7; i++) rows += `<row r="${i + 1}"><c r="A${i + 1}" s="${i}"><v>1234.5</v></c></row>`;
const sheet = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="${NS}" xmlns:r="${RNS}"><sheetData>${rows}</sheetData></worksheet>`;

const workbook = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="${NS}" xmlns:r="${RNS}"><sheets><sheet name="Probe" sheetId="1" r:id="rId1"/></sheets></workbook>`;

const wbRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="${RNS}/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="${RNS}/styles" Target="styles.xml"/>
</Relationships>`;

const rootRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="${RNS}/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;

const ct = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`;

const zip = new JSZip();
zip.file('[Content_Types].xml', ct);
zip.file('_rels/.rels', rootRels);
zip.file('xl/workbook.xml', workbook);
zip.file('xl/_rels/workbook.xml.rels', wbRels);
zip.file('xl/styles.xml', styles);
zip.file('xl/worksheets/sheet1.xml', sheet);
const buf = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
writeFileSync(out, buf);
console.log('wrote', out, buf.length);
