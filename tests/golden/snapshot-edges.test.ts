/**
 * The golden snapshot merges shared border edges (a line between two neighbouring cells is stored on either of them:
 * Excel moves A1.right to B1.left when it rewrites a border) and compares only what is DRAWN. These tests pin both
 * sides of that rule: the same drawn line is equal wherever it is stored, every real difference still shows.
 */
import { describe, expect, it } from 'vitest';
import { buildXlsx } from './build';
import { diffSnap, snapshot } from './snapshot';

const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';
const line = (tag: string, style: string) => `<${tag} style="${style}"><color indexed="64"/></${tag}>`;
const border = (l = '', r = '', t = '', b = '') => `<border>${l || '<left/>'}${r || '<right/>'}${t || '<top/>'}${b || '<bottom/>'}<diagonal/></border>`;

/** xf index = position in this list + 1: [border xml, fill id] */
const BORDERS: Array<[string, number]> = [
  [border('', line('right', 'thin')), 0], // 1  right thin
  [border(line('left', 'thin')), 0], // 2  left thin
  [border('', '', '', line('bottom', 'thin')), 0], // 3  bottom thin
  [border('', '', line('top', 'thin')), 0], // 4  top thin
  [border(line('left', 'medium')), 0], // 5  left medium
  [border(line('left', 'thin'), line('right', 'thin')), 0], // 6  left + right thin
  [border(line('left', 'thin')), 1], // 7  left thin + yellow fill
];

const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet ${NS}><fonts count="1"><font><sz val="11"/><name val="游ゴシック"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="${BORDERS.length + 1}">${border()}${BORDERS.map((b) => b[0]).join('')}</borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="${BORDERS.length + 1}"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>${BORDERS.map((b, i) => `<xf numFmtId="0" fontId="0" fillId="${b[1]}" borderId="${i + 1}" xfId="0" applyBorder="1"/>`).join('')}</cellXfs><cellStyles count="1"><cellStyle name="標準" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/></styleSheet>`;

/** a one-sheet book; cells = [ref, xf] (styled, empty) */
async function book(cells: Array<[string, number]>): Promise<Uint8Array> {
  const rows = new Map<number, Array<[string, number]>>();
  for (const [ref, xf] of cells) {
    const r = Number(/\d+/.exec(ref)![0]);
    rows.set(r, [...(rows.get(r) ?? []), [ref, xf]]);
  }
  const data = [...rows.entries()].sort((a, b) => a[0] - b[0])
    .map(([r, cs]) => `<row r="${r}">${cs.sort((a, b) => (a[0] < b[0] ? -1 : 1)).map(([ref, xf]) => `<c r="${ref}" s="${xf}"/>`).join('')}</row>`).join('');
  return buildXlsx([{ name: 'S', data }], { styles: STYLES });
}

async function diff(a: Array<[string, number]>, b: Array<[string, number]>): Promise<string[]> {
  return diffSnap(await snapshot(await book(a)), await snapshot(await book(b))).map((d) => d.path);
}

describe('snapshot: shared border edges are merged', () => {
  it('a line stored on the left cell equals the same line stored on the right cell (A1.right = B1.left)', async () => {
    expect(await diff([['A1', 1]], [['B1', 2]])).toEqual([]);
  });

  it('the same for a line between two rows (H1.bottom = H2.top)', async () => {
    expect(await diff([['H1', 3]], [['H2', 4]])).toEqual([]);
  });

  it('a line stored on both cells equals the line stored on one of them', async () => {
    expect(await diff([['A1', 1], ['B1', 2]], [['B1', 2]])).toEqual([]);
    expect(await diff([['A1', 1], ['B1', 2]], [['A1', 1]])).toEqual([]);
  });

  it('a cell that only exists for its border (empty, xf with one side) equals the neighbour that draws it', async () => {
    // A1 draws the line on its right side; B1 does not exist in the first file, holds the left side in the second
    expect(await diff([['A1', 1], ['C1', 6]], [['B1', 2], ['C1', 6]])).toEqual([]);
  });

  it('a line that is not drawn at all still differs', async () => {
    expect((await diff([['A1', 1]], [['A1', 0]])).length).toBeGreaterThan(0);
    expect((await diff([['A1', 1]], [])).length).toBeGreaterThan(0);
  });

  it('a line in another weight still differs', async () => {
    // A1.left thin against A1.left medium
    expect((await diff([['A1', 2]], [['A1', 5]])).length).toBeGreaterThan(0);
    // the neighbour's facing side is merged, the weights still have to agree: B1.left thin vs A1.right medium is not representable
    // here, but B1.left medium against A1.right thin must differ
    expect((await diff([['A1', 1]], [['B1', 5]])).length).toBeGreaterThan(0);
  });

  it('a line one cell further away still differs', async () => {
    expect((await diff([['A1', 1]], [['C1', 2]])).length).toBeGreaterThan(0);
    expect((await diff([['H1', 3]], [['H3', 4]])).length).toBeGreaterThan(0);
  });

  it('the outside edge of a cell is not mirrored anywhere (A1.left equals A1.left only)', async () => {
    expect(await diff([['A1', 2]], [['A1', 2]])).toEqual([]);
    expect((await diff([['A1', 2]], [['B1', 2]])).length).toBeGreaterThan(0);
  });

  it('a different fill still differs when the edges agree', async () => {
    expect((await diff([['A1', 2]], [['A1', 7]])).length).toBeGreaterThan(0);
  });

  it('the reported path stays [sheet].styles.<cell> (the keys known deviations are matched on)', async () => {
    const paths = await diff([['A1', 1]], []);
    expect(paths.every((p) => p.startsWith('[S].styles.'))).toBe(true);
  });
});

// ---------------------------------------------------------------------------------------------------------------
// the quote prefix counts only where it shows (a text constant)
// ---------------------------------------------------------------------------------------------------------------
const QSTYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet ${NS}><fonts count="1"><font><sz val="11"/><name val="游ゴシック"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1">${border()}</borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" quotePrefix="1"/></cellXfs><cellStyles count="1"><cellStyle name="標準" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/></styleSheet>`;
const qbook = (cell: string) => buildXlsx([{ name: 'S', data: `<row r="1">${cell}</row>` }], { styles: QSTYLES, sst: ['abc'] });

describe('snapshot: a quote prefix counts only on a text constant', () => {
  const d = async (a: string, b: string) => diffSnap(await snapshot(await qbook(a)), await snapshot(await qbook(b))).map((x) => x.path);

  it('on text it is a real difference', async () => {
    expect((await d('<c r="A1" s="1" t="s"><v>0</v></c>', '<c r="A1" t="s"><v>0</v></c>')).length).toBeGreaterThan(0);
  });

  it('on a number, a boolean, a formula or an empty cell it is invisible (equal with and without)', async () => {
    expect(await d('<c r="A1" s="1"><v>5</v></c>', '<c r="A1"><v>5</v></c>')).toEqual([]);
    expect(await d('<c r="A1" s="1" t="b"><v>1</v></c>', '<c r="A1" t="b"><v>1</v></c>')).toEqual([]);
    expect(await d('<c r="A1" s="1"><f>1+1</f><v>2</v></c>', '<c r="A1"><f>1+1</f><v>2</v></c>')).toEqual([]);
    expect(await d('<c r="A1" s="1"/>', '<c r="A1"/>')).toEqual([]);
  });
});
