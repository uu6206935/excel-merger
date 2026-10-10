// tools/lib/package-diff.mjs — the package ("normalized XML") view of tools/compare.mjs. Builds tiny .xlsx packages
// in memory and checks that known-benign differences are filtered (and counted) while real ones are reported.
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { compareInspections, formatComparison } from '../../tools/lib/compare-core.mjs';
import { comparePackages, FILTER_LABELS, isoDateSerial, parseXml } from '../../tools/lib/package-diff.mjs';
import { book, sheet } from './helpers';

const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const RNS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PR = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OD = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

// xf 0 default · xf 1 bold · xf 2 bold again (same RESOLVED style as xf 1) · xf 3 date (yyyy/m/d) · xf 4 percent
const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="${NS}">
<fonts count="2"><font><sz val="11"/><name val="游ゴシック"/><family val="2"/><charset val="128"/><scheme val="minor"/></font><font><b/><sz val="11"/><name val="游ゴシック"/><family val="3"/></font></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="5"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/><xf fontId="1" numFmtId="0" borderId="0" fillId="0" xfId="0"/><xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="9" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>
<cellStyles count="1"><cellStyle name="標準" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`;

interface Pkg {
  sheetData: string;
  /** everything after </sheetData> */
  tail?: string;
  dimension?: string;
  selection?: string;
  cols?: string;
  sst: string[];
  calcChain?: boolean;
}

async function xlsx(p: Pkg): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>`,
  );
  zip.file('_rels/.rels', `<?xml version="1.0"?><Relationships xmlns="${PR}"><Relationship Id="rId1" Type="${OD}/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
  zip.file(
    'xl/workbook.xml',
    `<?xml version="1.0"?><workbook xmlns="${NS}" xmlns:r="${RNS}"><sheets><sheet name="名簿" sheetId="1" r:id="rId1"/></sheets><definedNames><definedName name="志望度列">名簿!$B:$B</definedName></definedNames><calcPr calcId="191029"/></workbook>`,
  );
  zip.file(
    'xl/_rels/workbook.xml.rels',
    `<?xml version="1.0"?><Relationships xmlns="${PR}"><Relationship Id="rId1" Type="${OD}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${OD}/styles" Target="styles.xml"/><Relationship Id="rId3" Type="${OD}/sharedStrings" Target="sharedStrings.xml"/>${p.calcChain ? `<Relationship Id="rId4" Type="${OD}/calcChain" Target="calcChain.xml"/>` : ''}</Relationships>`,
  );
  zip.file('xl/styles.xml', STYLES);
  zip.file('xl/sharedStrings.xml', `<?xml version="1.0"?><sst xmlns="${NS}" count="${p.sst.length}" uniqueCount="${p.sst.length}">${p.sst.map((s) => `<si><t>${s}</t></si>`).join('')}</sst>`);
  if (p.calcChain) zip.file('xl/calcChain.xml', `<?xml version="1.0"?><calcChain xmlns="${NS}"><c r="B2" i="1"/></calcChain>`);
  zip.file(
    'xl/worksheets/sheet1.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="${NS}" xmlns:r="${RNS}"><dimension ref="${p.dimension ?? 'A1:D4'}"/><sheetViews><sheetView tabSelected="1" workbookViewId="0">${p.selection ?? '<selection activeCell="A1" sqref="A1"/>'}</sheetView></sheetViews><sheetFormatPr defaultRowHeight="18.75"/>${p.cols ?? '<cols><col min="1" max="1" width="12.5" customWidth="1"/></cols>'}<sheetData>${p.sheetData}</sheetData>${p.tail ?? '<mergeCells count="1"><mergeCell ref="C1:D1"/></mergeCells>'}</worksheet>`,
  );
  return zip.generateAsync({ type: 'uint8array' });
}

// A: shared strings [学生ID, 氏名], shared formula B3:B4, bold header via xf 1
const A: Pkg = {
  sst: ['学生ID', '氏名'],
  sheetData:
    '<row r="1"><c r="A1" s="1" t="s"><v>0</v></c><c r="B1" s="1" t="s"><v>1</v></c></row>' +
    '<row r="2"><c r="A2"><v>1</v></c><c r="B2"><f>A2*2</f><v>2</v></c></row>' +
    '<row r="3"><c r="A3"><v>2</v></c><c r="B3"><f t="shared" ref="B3:B4" si="0">A3*2</f><v>4</v></c></row>' +
    '<row r="4"><c r="A4"><v>3</v></c><c r="B4"><f t="shared" si="0"/><v>6</v></c></row>',
};

// B: the same workbook written differently — reversed sharedStrings, header style via the duplicate xf 2, plain
// formulas instead of a shared one, stale cached values, other attribute order, an empty styled-as-default <c/>,
// another dimension / selection and a calcChain part
const B: Pkg = {
  sst: ['氏名', '学生ID'],
  calcChain: true,
  dimension: 'A1:D5',
  selection: '<selection activeCell="C7" sqref="C7"/>',
  sheetData:
    '<row r="1" spans="1:4"><c t="s" s="2" r="A1"><v>1</v></c><c r="B1" t="s" s="2"><v>0</v></c></row>' +
    '<row r="2"><c r="A2"><v>1</v></c><c r="B2"><f>A2*2</f><v>0</v></c><c r="C2"/></row>' +
    '<row r="3"><c r="A3"><v>2</v></c><c r="B3"><f>A3*2</f><v>4</v></c></row>' +
    '<row r="4"><c r="A4"><v>3</v></c><c r="B4"><f>A4*2</f><v>6</v></c></row>',
};

const diffGroups = (rows: Array<{ ok: boolean; group: string; filtered?: boolean }>) => rows.filter((r) => !r.ok && !r.filtered).map((r) => r.group);

describe('package view (tools/lib/package-diff.mjs)', () => {
  it('parses machine-written XML (entities, CDATA, self-closing, prefixes)', () => {
    const x = parseXml('<?xml version="1.0"?><x:a xmlns:x="u" k="1&amp;2"><b/><c>t&lt;<![CDATA[<raw>]]></c></x:a>')!;
    expect(x.local).toBe('a');
    expect(x.attrs.k).toBe('1&2');
    expect(x.children.map((c) => c.local)).toEqual(['b', 'c']);
    expect(x.children[1].text).toBe('t<<raw>');
  });

  it('two writings of the same workbook are identical; every benign difference is counted as filtered', async () => {
    const r = await comparePackages(await xlsx(A), await xlsx(B));
    expect(diffGroups(r.rows), formatComparison(r.rows, { onlyDiffs: true })).toEqual([]);
    const keys = r.filtered.filter((f) => f.count != null).map((f) => f.key);
    for (const k of ['calcChain', 'xfIndex', 'sstIndex', 'sharedFormula', 'cachedValues', 'emptyCells', 'dimension', 'selection']) expect(keys, k).toContain(k);
    // always-ignored categories are listed (without a count) so the report says what is never compared
    expect(r.filtered.filter((f) => f.count == null).map((f) => f.key)).toEqual(['xmlNoise', 'styleNoise', 'otherParts']);
    // the filtered categories are rendered as a separate, labelled list
    const text = formatComparison(r.rows, { onlyDiffs: true });
    expect(text).toContain('Filtered (known-benign, not counted)');
    expect(text).toContain(FILTER_LABELS.calcChain.slice(0, 10));
  });

  it('a t="d" cell (an ISO date, Strict Open XML) is the same value as the date serial Excel writes for it', async () => {
    expect(isoDateSerial('2024-04-12')).toBe(45394);
    expect(isoDateSerial('2024-04-12T12:00:00')).toBe(45394.5);
    expect(isoDateSerial('2024-04-12', true)).toBe(45394 - 1462);
    expect(isoDateSerial('1900-01-15')).toBeNull(); // (before the 1900 leap-year bug: not converted)
    expect(isoDateSerial('abc')).toBeNull();
    const D: Pkg = { ...A, tail: '', sheetData: '<row r="1"><c r="A1" s="3" t="d"><v>2024-04-12</v></c></row>' };
    const N: Pkg = { ...A, tail: '', sheetData: '<row r="1"><c r="A1" s="3"><v>45394</v></c></row>' };
    const r = await comparePackages(await xlsx(D), await xlsx(N));
    expect(diffGroups(r.rows), formatComparison(r.rows, { onlyDiffs: true })).toEqual([]);
    const O: Pkg = { ...A, tail: '', sheetData: '<row r="1"><c r="A1" s="3"><v>45395</v></c></row>' };
    expect(diffGroups((await comparePackages(await xlsx(D), await xlsx(O))).rows)).toContain('values');
  });

  it('reports real differences: formula text, resolved style, merges, widths, names', async () => {
    const C: Pkg = {
      ...A,
      cols: '<cols><col min="1" max="1" width="20" customWidth="1"/></cols>',
      tail: '',
      sheetData:
        '<row r="1"><c r="A1" s="3" t="s"><v>0</v></c><c r="B1" s="1" t="s"><v>1</v></c></row>' +
        '<row r="2"><c r="A2"><v>1</v></c><c r="B2"><f>A2*3</f><v>3</v></c></row>' +
        '<row r="3"><c r="A3"><v>2</v></c><c r="B3"><f t="shared" ref="B3:B4" si="0">A3*2</f><v>4</v></c></row>' +
        '<row r="4"><c r="A4"><v>5</v></c><c r="B4"><f t="shared" si="0"/><v>10</v></c></row>',
    };
    const r = await comparePackages(await xlsx(A), await xlsx(C));
    const groups = diffGroups(r.rows);
    expect(groups).toEqual(expect.arrayContaining(['formulas', 'values', 'styles', 'merges', 'cols']));
    const detail = (g: string) => r.rows.find((x) => x.group === g && !x.ok)!.detail;
    expect(detail('formulas')).toContain('B2: =A2*2 -> =A2*3');
    expect(detail('values')).toContain('A4: 3 -> 5');
    expect(detail('styles')).toMatch(/A1: fmt: General -> yyyy\/m\/d, font: .*-> /);
    expect(detail('cols')).toContain('A 幅 12.5 -> 20');
    // the cached value of the changed formula is not reported separately (it is part of the formula difference)
    expect(detail('values')).not.toContain('B2');
  });

  it('dynamic-array formulas: the spill extent and spilled values are filtered, the formula kind is not', async () => {
    const spill = (ref: string, vals: number[], dyn = true) => ({
      sst: ['x'],
      sheetData:
        `<row r="1"><c r="A1"${dyn ? ' cm="1"' : ''}><f t="array" ref="${ref}">_xlfn.UNIQUE(B1:B9)</f><v>${vals[0]}</v></c></row>` +
        vals
          .slice(1)
          .map((v, i) => `<row r="${i + 2}"><c r="A${i + 2}"><v>${v}</v></c></row>`)
          .join(''),
    });
    const r = await comparePackages(await xlsx(spill('A1:A3', [1, 2, 3])), await xlsx(spill('A1:A2', [1, 5])));
    expect(diffGroups(r.rows)).toEqual([]);
    expect(r.filtered.map((f) => f.key)).toContain('spill');
    // a legacy CSE array formula vs a dynamic one IS a difference (Excel shows {=…} vs a spill)
    const r2 = await comparePackages(await xlsx(spill('A1:A3', [1, 2, 3])), await xlsx(spill('A1:A3', [1, 2, 3], false)));
    expect(diffGroups(r2.rows)).toEqual(['formulas']);
  });

  it('percent / date number formats are compared as format codes (built-in ids resolved)', async () => {
    const withFmt = (s: number) => ({ sst: ['x'], sheetData: `<row r="1"><c r="A1" s="${s}"><v>0.5</v></c></row>` });
    const same = await comparePackages(await xlsx(withFmt(4)), await xlsx(withFmt(4)));
    expect(diffGroups(same.rows)).toEqual([]);
    const other = await comparePackages(await xlsx(withFmt(4)), await xlsx(withFmt(3)));
    expect(other.rows.find((x) => x.group === 'styles')!.detail).toContain('fmt: 0% -> yyyy/m/d');
  });
});

describe('Excel view filters (compare-core compareInspections)', () => {
  it('values of volatile functions are filtered, other value differences are not', () => {
    const mk = (now: number, x: number) => book([sheet({ name: 'S', cells: { A1: now, B1: x, C1: 1 }, formulas: { A1: '=NOW()', B1: '=C1*2' } })]);
    const same = compareInspections(mk(45000.1, 2), mk(45000.2, 2));
    expect(same.filter((r) => !r.ok)).toEqual([]);
    expect(same.some((r) => r.filtered && r.group === 'values' && /揮発性/.test(r.detail))).toBe(true);
    const diff = compareInspections(mk(45000.1, 2), mk(45000.2, 3));
    expect(diff.filter((r) => !r.ok).map((r) => r.group)).toEqual(['values']);
  });
});
