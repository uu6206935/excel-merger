/**
 * Round-2 patcher features (no Excel needed): dynamic-array writes (xl/metadata.xml), CellWrite.numFmt (xf reuse,
 * custom ids, column widening), the 'unmerge' op, 'toValues' over array / spill masters, table header renames,
 * KeyFillEntry sourceCols / joiner / splitPart, rich-text imports.
 * The Excel comparisons of the same features are golden cases (tests/golden/cases.ts: r-*, u-*, h-*, h2-*).
 */
import fs from 'node:fs';
import path from 'node:path';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import type { CellWrite, LoggedOp } from '../../src/core/types';
import type { PremergeInput } from '../../src/core/contracts';
import { buildPremerged } from '../../src/core/xlsx/premerge';
import { replayLog } from '../../src/core/xlsx/replay';
import { numberTextPx } from '../../src/core/xlsx/patch/autowidth';
import { NEW_METADATA_XML, ensureDynamicArrayMetadata } from '../../src/core/xlsx/patch/metadata';
import { Sst } from '../../src/core/xlsx/patch/sst';
import { Styles } from '../../src/core/xlsx/patch/styles';
import { STYLES, buildXlsx, listParts, part, partBytes } from '../golden/build';
import { GOLDEN_FIX, keyFillScenario, richScenario, source } from '../golden/premerge-scenario';
import { snapshot, styleSigs } from '../golden/snapshot';

const read = (f: string) => new Uint8Array(fs.readFileSync(path.join(GOLDEN_FIX, f)));
const JP = /[ぁ-んァ-ン一-龥]/;
const cells = (sheet: string, writes: CellWrite[], reason: 'edit' | 'toValues' = 'edit'): LoggedOp => ({ kind: 'cells', sheet, writes, reason });
const w = (row: number, col: number, cell: CellWrite['cell']): CellWrite => ({ row, col, cell });

async function small(head = ''): Promise<Uint8Array> {
  return buildXlsx([
    {
      name: 'S',
      head,
      data: '<row r="1"><c r="A1" s="1" t="s"><v>0</v></c><c r="B1"><v>1</v></c></row><row r="2"><c r="A2"><v>2</v></c></row>',
    },
    { name: 'T', data: '<row r="1"><c r="A1"><f>S!A2</f><v>2</v></c></row>' },
  ], { sst: ['見出し'] });
}

async function sheetXml(bytes: Uint8Array, n = 1): Promise<string> {
  return (await part(bytes, `xl/worksheets/sheet${n}.xml`))!;
}

// ---------------------------------------------------------------------------------------------------------------
describe('dynamic arrays (CellWrite.dyn)', () => {
  it('creates xl/metadata.xml (+ relationship + content type) and stores <c cm><f t="array" ref=master>', async () => {
    const out = await replayLog(await small(), [cells('S', [w(4, 3, { t: 'n', v: 1, f: '_xlfn.SEQUENCE(3)', dyn: true }), w(5, 1, { t: 'z', v: null, f: 'SUM(A1:A2)' })])]);
    expect(await part(out, 'xl/metadata.xml')).toBe(NEW_METADATA_XML);
    expect(await part(out, 'xl/_rels/workbook.xml.rels')).toMatch(/Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/sheetMetadata" Target="metadata\.xml"/);
    expect(await part(out, '[Content_Types].xml')).toContain('<Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/>');
    const xml = await sheetXml(out);
    expect(xml).toContain('<c r="C4" cm="1"><f t="array" ref="C4">_xlfn.SEQUENCE(3)</f><v>1</v></c>');
    expect(xml).toContain('<c r="A5"><f>SUM(A1:A2)</f></c>'); // not dynamic: plain formula
    const snap = await snapshot(out);
    expect(snap.sheets[0].formulas.C4).toBe('_xlfn.SEQUENCE(3) {dynamic}');
  });

  it('reuses the XLDAPR cell metadata of a book that already has it (metadata part untouched)', async () => {
    const input = read('r2-arrays.xlsx');
    const out = await replayLog(input, [cells('A', [w(9, 1, { t: 'z', v: null, f: '_xlfn.SEQUENCE(2)', dyn: true })])]);
    expect(await part(out, 'xl/metadata.xml')).toBe(await part(input, 'xl/metadata.xml'));
    expect(await sheetXml(out)).toContain('<c r="A9" cm="1"><f t="array" ref="A9">_xlfn.SEQUENCE(2)</f></c>');
    expect((await listParts(out)).filter((p) => /metadata/.test(p))).toEqual(['xl/metadata.xml']);
  });

  it('a dynamic formula written before an insert follows it (ref included)', async () => {
    const out = await replayLog(await small(), [
      cells('S', [w(4, 3, { t: 'z', v: null, f: '_xlfn.SEQUENCE(3)*A2', dyn: true })]),
      { kind: 'struct', sheet: 'S', op: { type: 'insert', axis: 'col', at: 1, count: 1 } },
    ]);
    expect(await sheetXml(out)).toContain('<c r="D4" cm="1"><f t="array" ref="D4">_xlfn.SEQUENCE(3)*B2</f></c>');
  });

  it('ensureDynamicArrayMetadata merges into existing metadata (other types and indexes kept)', () => {
    // rich-value metadata only (images in cells): XLDAPR becomes type 2, its own future block and a new cell block
    const rich = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">' +
      '<metadataTypes count="1"><metadataType name="XLRICHVALUE" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1"/></metadataTypes>' +
      '<futureMetadata name="XLRICHVALUE" count="1"><bk><extLst><ext uri="{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}"><xlrd:rvb i="0"/></ext></extLst></bk></futureMetadata>' +
      '<valueMetadata count="1"><bk><rc t="1" v="0"/></bk></valueMetadata></metadata>';
    const r = ensureDynamicArrayMetadata(rich);
    expect(r.cm).toBe(1);
    expect(r.xml).toContain('<metadataTypes count="2"><metadataType name="XLRICHVALUE"');
    expect(r.xml).toMatch(/<metadataType name="XLDAPR"[^>]*cellMeta="1"\/><\/metadataTypes>/);
    expect(r.xml).toContain('xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray"');
    expect(r.xml).toContain('<futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata>');
    // element order: metadataTypes, futureMetadata (x2), cellMetadata, valueMetadata
    expect(r.xml.indexOf('<cellMetadata count="1"><bk><rc t="2" v="0"/></bk></cellMetadata>')).toBeGreaterThan(r.xml.lastIndexOf('</futureMetadata>'));
    expect(r.xml.indexOf('<cellMetadata')).toBeLessThan(r.xml.indexOf('<valueMetadata'));
    expect(r.xml).toContain('<valueMetadata count="1"><bk><rc t="1" v="0"/></bk></valueMetadata>');
    // idempotent: the merged part is reused as is
    expect(ensureDynamicArrayMetadata(r.xml)).toEqual(r);
    expect(ensureDynamicArrayMetadata(NEW_METADATA_XML)).toEqual({ xml: NEW_METADATA_XML, cm: 1 });
    // an XLDAPR type whose cell block is a collapsed record: a second block is added
    const collapsed = NEW_METADATA_XML.replace('fCollapsed="0"', 'fCollapsed="1"');
    const c = ensureDynamicArrayMetadata(collapsed);
    expect(c.cm).toBe(2);
    expect(c.xml).toContain('<futureMetadata name="XLDAPR" count="2">');
    expect(c.xml).toContain('<cellMetadata count="2"><bk><rc t="1" v="0"/></bk><bk><rc t="1" v="1"/></bk></cellMetadata>');
  });
});

// ---------------------------------------------------------------------------------------------------------------
describe('CellWrite.numFmt', () => {
  it('reuses an identical xf, keeps built-in ids for ja-JP built-in codes, appends custom codes >= 164', async () => {
    const out = await replayLog(await small(), [cells('S', [
      w(3, 1, { t: 'd', v: 46295, numFmt: 'yyyy/m/d' }), // xf 0 + id 14 = existing xf 3
      w(3, 2, { t: 'n', v: 0.5, numFmt: '0%' }), // built-in 9
      w(3, 3, { t: 'n', v: 12.5, numFmt: '0.0"pt"' }), // existing custom 164
      w(3, 4, { t: 'n', v: 12.5, numFmt: '0.0"kg"' }), // new custom 165
      w(1, 1, { t: 's', v: '見出し', numFmt: '@' }), // styled cell (bold, fill, border): only the number format changes
      w(4, 1, { t: 'n', v: 1, numFmt: '0.0"kg"' }), // same code again: same xf
    ])]);
    const xml = await sheetXml(out);
    const styles = (await part(out, 'xl/styles.xml'))!;
    expect(xml).toMatch(/<c r="A3" s="3"><v>46295<\/v><\/c>/);
    expect(styles).toContain('<numFmt numFmtId="165" formatCode="0.0&quot;kg&quot;"/>');
    expect(styles).toMatch(/<numFmts count="2">/);
    const sigs = styleSigs(styles);
    const sOf = (ref: string) => sigs[Number(new RegExp(`<c r="${ref}" s="(\\d+)"`).exec(xml)?.[1] ?? 0)];
    expect(sOf('B3')).toMatch(/^nf=#9 /);
    expect(sOf('C3')).toMatch(/^nf=0\.0&quot;pt&quot; |^nf=0\.0"pt" /);
    expect(sOf('D3')).toContain('kg');
    expect(sOf('A4')).toBe(sOf('D3'));
    expect(sOf('A1')).toMatch(/^nf=#49 font=<font><b\/>/);
    expect(sOf('A1')).toContain('patternType="solid"');
    // exactly one new xf per distinct (xf, code) pair: 0+9, 0+165, 1+49 (0+14 and 0+164 existed)
    expect((styles.match(/<cellXfs count="(\d+)"/) ?? [])[1]).toBe(String(4 + 4));
  });

  it('widens a standard-width column like Excel when the formatted number no longer fits (not user widths)', async () => {
    const out = await replayLog(await small('<cols><col min="4" max="4" width="6" customWidth="1"/><col min="5" max="5" width="3" hidden="1" customWidth="1"/></cols>'), [cells('S', [
      w(3, 1, { t: 'd', v: 46295, numFmt: 'yyyy/m/d' }), // 2026/9/30 -> 9.63 chars = stored 10.25
      w(3, 2, { t: 'n', v: 0.5, numFmt: '0%' }), // fits
      w(3, 3, { t: 'n', v: 1234567, numFmt: '#,##0' }), // 1,234,567 -> 8.88 = 9.5
      w(4, 3, { t: 'n', v: 123456789, numFmt: '#,##0' }), // auto-fitted (bestFit) columns keep growing: 11 -> 11.625
      w(3, 4, { t: 'd', v: 46295, numFmt: 'yyyy/m/d' }), // user width: kept
      w(3, 5, { t: 'd', v: 46295, numFmt: 'yyyy/m/d' }), // hidden: kept
      w(3, 6, { t: 'z', v: null, f: 'A3', numFmt: 'yyyy/m/d' }), // formula: value unknown, no widening
    ])]);
    const head = /<cols>[\s\S]*?<\/cols>/.exec(await sheetXml(out))![0];
    expect(head).toBe('<cols><col min="1" max="1" width="10.25" bestFit="1" customWidth="1"/><col min="3" max="3" width="11.625" bestFit="1" customWidth="1"/>' +
      '<col min="4" max="4" width="6" customWidth="1"/><col min="5" max="5" width="3" hidden="1" customWidth="1"/></cols>');
  });

  it('text widths reproduce the widths Excel chose (tests/golden/com/probe-numfmt-width.ps1)', () => {
    // ColumnWidth (chars) Excel set -> content px = chars * 8
    const probe: Array<[string, number]> = [
      ['2026/1/1', 8.63], ['2026/9/30', 9.63], ['2026/10/10', 10.75], ['2026/9/30 9:30', 14.13], ['123,456,789', 11],
      ['12,345,678', 9.88], ['1,234,567', 8.88], ['12345.60%', 9.88], ['$1,234,567', 10], ['1,234,567.89', 11.63],
      ['12345678901', 12.13], ['2026-09-30', 10.75], ['-1,234,567', 9.75], ['2026年9月30日', 14],
    ];
    for (const [text, chars] of probe) expect(Math.abs(numberTextPx(text) - Math.round(chars * 8)), text).toBeLessThanOrEqual(1);
    // texts that fit the default 8.38 (67 px) were not widened
    for (const text of ['12:34:20', 'R8.9.30', '50%', '9/30']) expect(numberTextPx(text), text).toBeLessThanOrEqual(67);
    // a 16 pt cell (max digit width 12): 14.63 -> 117 px
    expect(Math.abs(numberTextPx('2026/9/30', 12) - 117)).toBeLessThanOrEqual(2);
  });

  it('Styles.withNumFmt on a styles part without numFmts creates the section in schema order', () => {
    const st = new Styles(STYLES.replace(/<numFmts[\s\S]*?<\/numFmts>/, ''));
    const i = st.withNumFmt(0, '0.000');
    expect(i).toBe(4);
    const x = st.serialize();
    expect(x).toMatch(/<styleSheet [^>]*><numFmts count="1"><numFmt numFmtId="164" formatCode="0.000"\/><\/numFmts><fonts/);
    expect(st.withNumFmt(0, 'General')).toBe(0);
    expect(st.withNumFmt(0, '0.000')).toBe(i);
  });
});

// ---------------------------------------------------------------------------------------------------------------
describe("LoggedOp 'unmerge'", () => {
  it('centerAcross: removes the merge and gives every cell of the area horizontal="centerContinuous" (other alignment kept)', async () => {
    const input = read('torture.xlsx');
    const before = (await snapshot(input)).sheets.find((s) => s.name === 'Data')!;
    expect(before.merges).toContain('A1:F1');
    const out = await replayLog(input, [{ kind: 'unmerge', sheet: 'Data', refs: ['$A$1:$F$1'], centerAcross: true }]);
    const s = (await snapshot(out)).sheets.find((x) => x.name === 'Data')!;
    expect(s.merges).not.toContain('A1:F1');
    expect(s.values.A1).toBe(before.values.A1);
    const fillOf = (sig: string) => /fill=(<fill>.*?<\/fill>)/.exec(sig)?.[1];
    for (const c of 'ABCDEF') {
      expect(s.styles[`${c}1`], c).toContain('horizontal="centerContinuous"');
      // the rest of the style (font, border, vertical alignment) is the original one; since round 3 the fill is the top-left
      // cell's (every other cell of the former area receives it, see patch-round3.test.ts)
      const strip = (sig: string) => sig.replace(/<alignment[^>]*>/, '').replace(/fill=<fill>.*?<\/fill>/, 'fill=-');
      const orig = strip(before.styles[`${c}1`] ?? before.styles.A1 ?? '');
      if (before.styles[`${c}1`]) expect(strip(s.styles[`${c}1`])).toBe(orig);
      expect(fillOf(s.styles[`${c}1`]), c).toBe(fillOf(before.styles.A1));
    }
    // the merge no longer blocks a column move through it
    await expect(replayLog(input, [{ kind: 'struct', sheet: 'Data', op: { type: 'move', axis: 'col', from: 3, count: 1, to: 8 } }])).rejects.toThrow(JP);
    await expect(replayLog(input, [{ kind: 'unmerge', sheet: 'Data', refs: ['A1:F1'], centerAcross: true }, { kind: 'struct', sheet: 'Data', op: { type: 'move', axis: 'col', from: 3, count: 1, to: 8 } }])).resolves.toBeInstanceOf(Uint8Array);
  });

  it('creates missing cells of the area with the aligned style; without centerAcross only the merge goes', async () => {
    const book = await buildXlsx([{
      name: 'M',
      data: '<row r="1"><c r="A1" s="1" t="s"><v>0</v></c></row><row r="3"><c r="B3"><v>1</v></c></row>',
      tail: '<mergeCells count="2"><mergeCell ref="A1:C2"/><mergeCell ref="B3:D3"/></mergeCells>',
    }], { sst: ['タイトル'] });
    const out = await replayLog(book, [{ kind: 'unmerge', sheet: 'M', refs: ['A1:C2'], centerAcross: true }]);
    const xml = await sheetXml(out);
    expect(xml).toContain('<mergeCells count="1"><mergeCell ref="B3:D3"/></mergeCells>');
    const styles = styleSigs((await part(out, 'xl/styles.xml'))!);
    for (const ref of ['A1', 'B1', 'C1', 'A2', 'B2', 'C2']) {
      const s = Number(new RegExp(`<c r="${ref}" s="(\\d+)"`).exec(xml)?.[1]);
      expect(styles[s], ref).toContain('horizontal="centerContinuous"');
    }
    // A1 kept its bold / fill / border (xf 1) besides the alignment; empty cells got xf 0 + alignment
    const a1 = styles[Number(/<c r="A1" s="(\d+)"/.exec(xml)![1])];
    expect(a1).toContain('<b/>');
    expect(a1).toContain('patternType="solid"');
    const plain = await replayLog(book, [{ kind: 'unmerge', sheet: 'M', refs: ['B3:D3', 'X9:Y9'], centerAcross: false }]);
    const px = await sheetXml(plain);
    expect(px).toContain('<mergeCells count="1"><mergeCell ref="A1:C2"/></mergeCells>');
    expect(await part(plain, 'xl/styles.xml')).toBe(await part(book, 'xl/styles.xml'));
    const both = await replayLog(book, [{ kind: 'unmerge', sheet: 'M', refs: ['A1:C2', 'B3:D3'], centerAcross: false }]);
    expect(await sheetXml(both)).not.toContain('mergeCell');
    await expect(replayLog(book, [{ kind: 'unmerge', sheet: 'M', refs: ['not a ref'], centerAcross: true }])).rejects.toThrow(JP);
  });
});

// ---------------------------------------------------------------------------------------------------------------
describe("'toValues' over array / spill masters", () => {
  it('a legacy (CSE) array master turns the whole array into values; edits inside an array are still refused', async () => {
    const input = read('r2-arrays.xlsx');
    await expect(replayLog(input, [cells('A', [w(3, 2, { t: 'n', v: 1 })])])).rejects.toThrow(/配列/);
    await expect(replayLog(input, [cells('A', [w(3, 2, { t: 'n', v: 1 })], 'toValues')])).rejects.toThrow(/配列/); // not the master
    const out = await replayLog(input, [cells('A', [w(2, 2, { t: 'n', v: 10 })], 'toValues'), cells('A', [w(3, 2, { t: 'n', v: 99 })])]);
    const s = (await snapshot(out)).sheets[0];
    expect(s.formulas.B2).toBeUndefined();
    expect(s.values.B2).toBe('n:10');
    expect(s.values.B3).toBe('n:99'); // the array is gone: the cell can be edited
    expect(s.values.B4).toBe('n:30');
    expect(s.formulas.K2).toBe('SUM(B2:B4)');
  });

  it('a spill master: values stay, t="array" / ref / cm go, "formula string" results become shared strings', async () => {
    const out = await replayLog(read('r2-arrays.xlsx'), [cells('A', [w(2, 7, { t: 's', v: '1x' }), w(2, 4, { t: 'n', v: 1 })], 'toValues')]);
    const xml = await sheetXml(out);
    expect(xml).not.toMatch(/<c r="G2"[^>]*cm=/);
    expect(xml).not.toMatch(/<c r="D2"[^>]*cm=/);
    expect(xml).toMatch(/<c r="G3"[^>]*t="s"><v>\d+<\/v><\/c>/);
    const s = (await snapshot(out)).sheets[0];
    expect(s.formulas.G2).toBeUndefined();
    expect(s.values.G2).toBe('s:1x');
    expect(s.values.G3).toBe('s:2x');
    expect(s.values.G4).toBe('s:3x');
    expect(s.values.E4).toBe('n:6');
    // a formula referring to the former spill keeps its text (Excel: #REF! on recalculation, as in golden r-tovalues)
    expect(s.formulas.K3).toBe('SUM(_xlfn.ANCHORARRAY(D2))');
  });

  it('an edit over a spill master clears the spilled values (the new value / formula replaces the spill)', async () => {
    const s = (await snapshot(await replayLog(read('r2-arrays.xlsx'), [cells('A', [w(2, 4, { t: 'n', v: 5 })])]))).sheets[0];
    expect(s.values.D2).toBe('n:5');
    for (const ref of ['D3', 'D4', 'E2', 'E3', 'E4']) expect(s.values[ref], ref).toBeUndefined();
  });
});

// ---------------------------------------------------------------------------------------------------------------
describe('table header rename (a write into a table header cell)', () => {
  it('renames the column in the table part, calculated / totals formulas, names and every sheet', async () => {
    const out = await replayLog(read('r2-table.xlsx'), [cells('T', [w(1, 2, { t: 's', v: '得点' })])]);
    const snap = await snapshot(out);
    const t = snap.sheets.find((x) => x.name === 'T')!;
    expect(t.values.B1).toBe('s:得点');
    expect(t.tables[0]).toBe('Tbl ref=A1:C5 hdr=1 tot=1 af=A1:C4 cols=[ID|得点|Twice] calc=[Tbl[[#This Row],[得点]]*2] totals=[SUM(Tbl[得点])*2|ID: label=集計|得点:custom|Twice:sum]');
    expect(t.formulas.B5).toBe('SUM(Tbl[得点])*2');
    expect(snap.names).toEqual(['ScoreCol = Tbl[得点]']);
    const u = snap.sheets.find((x) => x.name === 'U')!;
    expect(u.formulas).toMatchObject({ A1: 'SUM(Tbl[得点])', A2: 'Tbl[[#Headers],[得点]]', A3: 'SUM(Tbl[[得点]:[Twice]])', A5: 'SUM(Tbl[Twice])' });
  });

  it('unique names (case-insensitive: Score -> score2), 列N for a cleared header, numbers as text', async () => {
    const out = await replayLog(read('r2-table.xlsx'), [cells('T', [w(1, 3, { t: 's', v: 'score' }), w(1, 1, null)]), cells('T', [w(1, 2, { t: 'n', v: 2024 })])]);
    const t = (await snapshot(out)).sheets.find((x) => x.name === 'T')!;
    expect(t.tables[0]).toContain('cols=[列1|2024|score2]');
    expect(t.values).toMatchObject({ A1: 's:列1', B1: 's:2024', C1: 's:score2' });
    expect(t.formulas.C5).toBe('SUBTOTAL(109,Tbl[score2])');
  });

  it('a rename followed by a delete of that column makes its references #REF! (never the old name)', async () => {
    const out = await replayLog(read('r2-table.xlsx'), [cells('T', [w(1, 2, { t: 's', v: 'P' })]), { kind: 'struct', sheet: 'T', op: { type: 'delete', axis: 'col', at: 2, count: 1 } }]);
    const snap = await snapshot(out);
    const u = snap.sheets.find((x) => x.name === 'U')!;
    expect(u.formulas.A1).toBe('SUM(#REF!)');
    expect(u.formulas.A5).toBe('SUM(Tbl[Twice])');
  });
});

// ---------------------------------------------------------------------------------------------------------------
describe('table totals row (a write into a totals cell redefines the column totals entry)', () => {
  const tableXml = async (b: Uint8Array) => (await part(b, 'xl/tables/table1.xml'))!;
  const col = (xml: string, name: string) => new RegExp(`<tableColumn\\b[^>]*name="${name}"[^>]*?(?:/>|>[\\s\\S]*?</tableColumn>)`).exec(xml)![0];

  it('a constant becomes the column label (numbers / booleans as text, the cell holds that text), functions are dropped', async () => {
    const out = await replayLog(read('r2-table.xlsx'), [cells('T', [w(5, 1, { t: 's', v: '計' }), w(5, 2, { t: 'n', v: 5 }), w(5, 3, { t: 'b', v: true })])]);
    const x = await tableXml(out);
    expect(col(x, 'ID')).toMatch(/totalsRowLabel="計"\/>$/);
    expect(col(x, 'Score')).toMatch(/totalsRowLabel="5"\/>$/);
    expect(col(x, 'Score')).not.toMatch(/totalsRowFunction|totalsRowFormula/);
    expect(col(x, 'Twice')).toContain('totalsRowLabel="TRUE"');
    expect(col(x, 'Twice')).not.toContain('totalsRowFunction');
    expect(col(x, 'Twice')).toContain('<calculatedColumnFormula>Tbl[[#This Row],[Score]]*2</calculatedColumnFormula>');
    const t = (await snapshot(out)).sheets.find((s) => s.name === 'T')!;
    expect(t.values).toMatchObject({ A5: 's:計', B5: 's:5', C5: 's:TRUE' });
    expect(t.formulas.B5).toBeUndefined();
  });

  it('SUBTOTAL(n, Table[this column]) is a built-in function, anything else custom with its formula', async () => {
    const f = (c: number, text: string) => w(5, c, { t: 'z', v: null, f: text });
    const out = await replayLog(read('r2-table.xlsx'), [cells('T', [f(1, 'SUBTOTAL(109,Tbl[Score])'), f(2, 'SUBTOTAL(104,Tbl[Score])'), f(3, 'SUBTOTAL(103,Tbl[Twice])')])]);
    const x = await tableXml(out);
    expect(col(x, 'ID')).toContain('totalsRowFunction="custom"><totalsRowFormula>SUBTOTAL(109,Tbl[Score])</totalsRowFormula>');
    expect(col(x, 'Score')).toMatch(/totalsRowFunction="max"\/>$/);
    expect(col(x, 'Twice')).toMatch(/totalsRowFunction="count"[^>]*><calculatedColumnFormula>[^<]*<\/calculatedColumnFormula><\/tableColumn>$/);
    expect(x).not.toContain('totalsRowLabel');
    const fns: Record<string, string> = { 101: 'average', 102: 'countNums', 103: 'count', 104: 'max', 105: 'min', 107: 'stdDev', 109: 'sum', 110: 'var' };
    for (const [n, fn] of Object.entries(fns)) {
      const o = await replayLog(read('r2-table.xlsx'), [cells('T', [f(2, `SUBTOTAL(${n},Tbl[Score])`)])]);
      expect(col(await tableXml(o), 'Score'), n).toContain(`totalsRowFunction="${fn}"`);
    }
    // not a built-in: other function numbers, spaces, other columns, plain formulas
    for (const text of ['SUBTOTAL(108,Tbl[Score])', 'SUBTOTAL(109, Tbl[Score])', 'SUBTOTAL(109,Tbl[Twice])', 'SUM(B2:B4)', '1+1']) {
      const o = await replayLog(read('r2-table.xlsx'), [cells('T', [f(2, text)])]);
      expect(col(await tableXml(o), 'Score'), text).toContain(`totalsRowFunction="custom"><totalsRowFormula>${text}</totalsRowFormula>`);
    }
  });

  it('clearing removes label / function / formula; a formula dyn flag is ignored', async () => {
    const out = await replayLog(read('r2-table.xlsx'), [cells('T', [w(5, 1, null), w(5, 2, null), w(5, 3, { t: 's', v: '' })])]);
    const x = await tableXml(out);
    expect(x).not.toMatch(/totalsRow(Label|Function|Formula)/);
    const t = (await snapshot(out)).sheets.find((s) => s.name === 'T')!;
    expect(t.values.A5).toBeUndefined();
    expect(t.formulas.B5).toBeUndefined();
    const dyn = await replayLog(read('r2-table.xlsx'), [cells('T', [w(5, 2, { t: 'z', v: null, f: 'MAX(Tbl[Score])', dyn: true })])]);
    expect(await sheetXml(dyn)).toContain('<c r="B5"><f>MAX(Tbl[Score])</f></c>');
  });

  it('ops before the write are not applied to the written formula, ops after it are', async () => {
    const ins = (at: number) => ({ kind: 'struct', sheet: 'T', op: { type: 'insert', axis: 'col', at, count: 1 } }) as LoggedOp;
    const before = await replayLog(read('r2-table.xlsx'), [ins(1), cells('T', [w(5, 3, { t: 'z', v: null, f: 'SUM(C2:C4)' })])]);
    expect(col(await tableXml(before), 'Score')).toContain('<totalsRowFormula>SUM(C2:C4)</totalsRowFormula>');
    const after = await replayLog(read('r2-table.xlsx'), [cells('T', [w(5, 2, { t: 'z', v: null, f: 'SUM(B2:B4)' })]), ins(1)]);
    expect(col(await tableXml(after), 'Score')).toContain('<totalsRowFormula>SUM(C2:C4)</totalsRowFormula>');
    expect(await tableXml(after)).toContain('ref="B1:D5"');
  });

  it('a write below a table without a totals row, and into a totals row after a rename, behave normally', async () => {
    const out = await replayLog(read('r2-table.xlsx'), [cells('T', [w(1, 2, { t: 's', v: '得点' })]), cells('T', [w(5, 2, { t: 'z', v: null, f: 'SUBTOTAL(109,Tbl[得点])' })])]);
    expect(col(await tableXml(out), '得点')).toMatch(/totalsRowFunction="sum"\/>$/);
    const plain = await replayLog(read('g-table.xlsx'), [cells('名簿', [w(16, 1, { t: 's', v: 'x' })])]);
    expect(await tableXml(plain)).not.toMatch(/totalsRow(Label|Function|Formula|Count)/);
  });
});

// ---------------------------------------------------------------------------------------------------------------
describe('premerge: KeyFillEntry sourceCols / joiner / splitPart', () => {
  async function books(): Promise<{ base: Uint8Array; src: Uint8Array }> {
    const base = await buildXlsx([{
      name: 'B',
      data: '<row r="1">' + ['ID', '姓', '名', '氏名', '番号'].map((h, i) => `<c r="${'ABCDE'[i]}1" t="inlineStr"><is><t>${h}</t></is></c>`).join('') + '</row>' +
        '<row r="2"><c r="A2" t="inlineStr"><is><t>S1</t></is></c><c r="E2"><v>1</v></c></row><row r="3"><c r="A3" t="inlineStr"><is><t>S2</t></is></c><c r="E3"><v>2</v></c></row>',
    }]);
    const src = await buildXlsx([{
      name: 'S',
      data: '<row r="1">' + ['ID', '氏名', '姓', '名', '番号'].map((h, i) => `<c r="${'ABCDE'[i]}1" t="inlineStr"><is><t>${h}</t></is></c>`).join('') + '</row>' +
        '<row r="2"><c r="A2" t="inlineStr"><is><t>S9</t></is></c><c r="B2" t="inlineStr"><is><t>山田 太郎</t></is></c><c r="C2" t="inlineStr"><is><t>佐藤</t></is></c><c r="D2" t="inlineStr"><is><t>花子</t></is></c><c r="E2"><v>123</v></c></row>' +
        '<row r="3"><c r="A3" t="inlineStr"><is><t>S8</t></is></c><c r="B3" t="inlineStr"><is><t>鈴木</t></is></c><c r="D3" t="inlineStr"><is><t>次郎</t></is></c><c r="E3"><v>7</v></c></row>',
    }]);
    return { base, src };
  }
  function input(b: { base: Uint8Array; src: Uint8Array }, keyFill: PremergeInput['keyFill']['s']): PremergeInput {
    return {
      base: source('b', 'base.xlsx', { sheet: 'B', headerRow: 1, firstDataRow: 2, lastDataRow: 3, firstCol: 1, lastCol: 5 }, 'base', 'full', b.base),
      sources: [source('s', 'src.xlsx', { sheet: 'S', headerRow: 1, firstDataRow: 2, lastDataRow: 3, firstCol: 1, lastCol: 5 }, 'src', 'full', b.src)],
      rows: [
        { baseRow: 2, matches: { s: null } }, { baseRow: 3, matches: { s: null } },
        { baseRow: null, appendedFrom: { sourceId: 's', row: 2 }, matches: { s: { row: 2, via: 'ID', flags: [] } } },
        { baseRow: null, appendedFrom: { sourceId: 's', row: 3 }, matches: { s: { row: 3, via: 'ID', flags: [] } } },
      ],
      imports: [],
      keyFill: { s: keyFill },
      options: { formulaPolicy: 'keep-if-possible', extendTable: true, headerConflict: 'suffix' },
    };
  }

  it('joins several columns (default full-width space, custom joiner) and splits on white space', async () => {
    const r = await buildPremerged(input(await books(), [
      { baseCol: 1, sourceCol: 1 },
      { baseCol: 2, sourceCol: 2, splitPart: 0 },
      { baseCol: 3, sourceCol: 2, splitPart: 1 },
      { baseCol: 4, sourceCols: [3, 4] },
      { baseCol: 5, sourceCols: [1, 5], joiner: '-' },
    ]));
    const s = (await snapshot(r.bytes)).sheets[0];
    expect([s.values.A4, s.values.B4, s.values.C4, s.values.D4, s.values.E4]).toEqual(['s:S9', 's:山田', 's:太郎', 's:佐藤　花子', 's:S9-123']);
    // one part only: part 1 is empty, an empty column is skipped by the join
    expect([s.values.A5, s.values.B5, s.values.C5, s.values.D5, s.values.E5]).toEqual(['s:S8', 's:鈴木', undefined, 's:次郎', 's:S8-7']);
    expect(r.layout.appended).toEqual([{ row: 4, sourceId: 's', sourceRow: 2 }, { row: 5, sourceId: 's', sourceRow: 3 }]);
  });

  it('on the Excel-authored premerge scenario (appended S009 row of pm-src.xlsx)', async () => {
    const r = await buildPremerged(keyFillScenario());
    const s = (await snapshot(r.bytes)).sheets.find((x) => x.name === '学生')!;
    expect([s.values.A7, s.values.B7, s.values.C7]).toEqual(['s:S009-Z', 's:学生Z　Z', 's:学生Z']);
    expect(s.formulas.D9).toBe('SUM(D2:D7)');
  });

  it('a single part of a number stays a number; bad entries are ignored', async () => {
    const r = await buildPremerged(input(await books(), [
      { baseCol: 5, sourceCol: 5, splitPart: 0 },
      { baseCol: 2, sourceCols: [] as number[], sourceCol: 2, splitPart: 1 },
      { baseCol: 3, sourceCol: 99999 },
      { baseCol: 4, sourceCol: 2, splitPart: -1 },
    ]));
    const s = (await snapshot(r.bytes)).sheets[0];
    expect(s.values.E4).toBe('n:123');
    expect(s.values.B4).toBe('s:太郎');
    expect(s.values.C4).toBeUndefined();
    expect(s.values.D4).toBeUndefined();
  });
});

// ---------------------------------------------------------------------------------------------------------------
describe('premerge: rich text (formatting runs / phonetic readings) survives the import', () => {
  it('copies the <si> runs into the base shared strings and keeps the run fonts', async () => {
    const r = await buildPremerged(richScenario());
    const xml = await sheetXml(r.bytes);
    const sst = (await part(r.bytes, 'xl/sharedStrings.xml'))!;
    const items = [...sst.matchAll(/<si>([\s\S]*?)<\/si>/g)].map((m) => m[1]);
    const idx = (ref: string) => Number(new RegExp(`<c r="${ref}"[^>]*t="s"[^>]*><v>(\\d+)</v>`).exec(xml)?.[1]);
    // E2 = "Bold normal": run 1 takes the cell font (bold red, imported with the cell style), run 2 its own <rPr>
    const e2 = items[idx('E2')];
    expect(e2).toBe('<r><t>Bold</t></r><r><rPr><sz val="11"/><color theme="1"/><rFont val="游ゴシック"/><family val="2"/><charset val="128"/><scheme val="minor"/></rPr><t xml:space="preserve"> normal</t></r>');
    const styles = (await part(r.bytes, 'xl/styles.xml'))!;
    const sig = styleSigs(styles)[Number(/<c r="E2" s="(\d+)"/.exec(xml)![1])];
    expect(sig).toMatch(/font=<font><b\/>/);
    expect(sig).toContain('FFFF0000');
    // the appended row's key cell (keyFill) keeps the base format, so run 1 gets the source cell font as <rPr>
    expect(items[idx('B7')]).toBe('<r><rPr><b/><sz val="11"/><color rgb="FFFF0000"/><rFont val="游ゴシック"/><family val="3"/><charset val="128"/><scheme val="minor"/></rPr><t>Bold</t></r>' +
      '<r><rPr><sz val="11"/><color theme="1"/><rFont val="游ゴシック"/><family val="2"/><charset val="128"/><scheme val="minor"/></rPr><t xml:space="preserve"> normal</t></r>');
    expect(xml).toMatch(/<c r="B7" s="\d+" t="s">|<c r="B7" t="s">/);
    // E3 = 山田 with its reading: phoneticPr points at the imported 6 pt font of the source
    const e3 = items[idx('E3')];
    expect(e3).toMatch(/^<t>山田<\/t><rPh sb="0" eb="2"><t>ヤマダ<\/t><\/rPh><phoneticPr fontId="(\d+)"\/>$/);
    const fontId = Number(/fontId="(\d+)"/.exec(e3)![1]);
    const fonts = [...(/<fonts\b[^>]*>([\s\S]*?)<\/fonts>/.exec(styles)![1]).matchAll(/<font>[\s\S]*?<\/font>|<font\/>/g)].map((m) => m[0]);
    expect(fonts[fontId]).toContain('<sz val="6"/>');
    // plain strings stay plain
    expect(items[idx('F2')]).toBe('<t>plain</t>');
    const snap = await snapshot(r.bytes);
    expect(snap.sheets[0].values).toMatchObject({ E1: 's:NAME', E2: 's:Bold normal', E3: 's:山田', F2: 's:plain', B7: 's:Bold normal', A7: 's:S001' });
  });

  it('Sst.addRich deduplicates by XML and never collides with plain items', () => {
    const s = new Sst('<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2"><si><t>a</t></si><si><r><t>a</t></r></si></sst>');
    expect(s.add('a')).toBe(0);
    expect(s.addRich('<r><t>a</t></r>')).toBe(1);
    expect(s.addRich('<r><rPr><b/></rPr><t>a</t></r>')).toBe(2);
    expect(s.addRich('<r><rPr><b/></rPr><t>a</t></r>')).toBe(2);
    expect(s.text(2)).toBe('a');
    expect(s.serialize()).toContain('uniqueCount="3"');
  });
});

// ---------------------------------------------------------------------------------------------------------------
describe('package hygiene of the round-2 writes', () => {
  it('every new part is a valid zip entry and the untouched parts keep their bytes', async () => {
    const input = read('r2-arrays.xlsx');
    const out = await replayLog(input, [cells('A', [w(9, 1, { t: 'z', v: null, f: '_xlfn.SEQUENCE(2)', dyn: true }), w(9, 3, { t: 'd', v: 46295, numFmt: 'yyyy/m/d' })])]);
    const zin = await JSZip.loadAsync(input);
    for (const p of Object.keys(zin.files)) {
      if (zin.files[p].dir || /worksheets\/sheet1\.xml|workbook\.xml$|styles\.xml|calcChain|Content_Types|workbook\.xml\.rels/.test(p)) continue;
      expect(await partBytes(out, p), p).toEqual(await partBytes(input, p));
    }
  });
});
