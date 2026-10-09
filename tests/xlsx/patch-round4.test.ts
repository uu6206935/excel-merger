/**
 * Round-4 patcher parity fixes (the Excel comparison of the wrap-text rule is golden case q-wrap, tests/golden/cases.ts):
 *  1. a text constant with a line feed (LF, CR LF; not a lone CR — measured) gets WrapText on its xf (Excel does this by itself
 *     when the text is written);
 *  2. CellWrite.quote semantics: quote === true -> quotePrefix, every other write adds none (a prefix the cell already
 *     had is cleared by a non-quote write: the documented difference to Excel COM, golden QUOTE_COM);
 *  3. text with CR / control characters is written the way Excel writes it (_xHHHH_), so it reads back unchanged.
 */
import { describe, expect, it } from 'vitest';
import type { CellWrite, LoggedOp } from '../../src/core/types';
import { Styles } from '../../src/core/xlsx/patch/styles';
import { replayLog } from '../../src/core/xlsx/replay';
import { readBook } from '../../src/core/xlsx/read';
import { buildXlsx, part } from '../golden/build';
import { snapshot } from '../golden/snapshot';
import { R3_STYLES } from '../golden/tools/make-r3-fixture';

const cells = (sheet: string, writes: CellWrite[], reason: 'edit' | 'toValues' = 'edit'): LoggedOp => ({ kind: 'cells', sheet, writes, reason });
const w = (row: number, col: number, cell: CellWrite['cell']): CellWrite => ({ row, col, cell });

/** xf element (open tag + alignment) that cell `ref` of sheet1 uses in a package */
async function xfOf(bytes: Uint8Array, ref: string): Promise<string> {
  const xml = (await part(bytes, 'xl/worksheets/sheet1.xml'))!;
  const s = Number(new RegExp(`<c r="${ref}"[^>]*?\\ss="(\\d+)"`).exec(xml)?.[1] ?? 0);
  const styles = (await part(bytes, 'xl/styles.xml'))!;
  const xfs = [...(/<cellXfs\b[\s\S]*?<\/cellXfs>/.exec(styles)![0]).matchAll(/<xf\b[^>]*?(?:\/>|>[\s\S]*?<\/xf>)/g)].map((m) => m[0]);
  return xfs[s];
}
const xfCount = async (bytes: Uint8Array) => Number(/<cellXfs count="(\d+)"/.exec((await part(bytes, 'xl/styles.xml'))!)![1]);

/** R3_STYLES xf 4: theme-filled title (centre, vertical centre); 0 = default. A1 of the book is styled, the rest is empty. */
async function book(): Promise<Uint8Array> {
  return buildXlsx([{ name: 'S', data: '<row r="1"><c r="A1" s="4" t="s"><v>0</v></c></row>' }], { sst: ['タイトル'], styles: R3_STYLES });
}

describe('wrap text for a text constant with a line break', () => {
  it('LF and CRLF derive an xf with wrapText="1" (other attributes of the xf kept); a lone CR does not, like Excel (golden q-wrap)', async () => {
    const out = await replayLog(await book(), [cells('S', [
      w(3, 1, { t: 's', v: 'a\nb' }), w(3, 2, { t: 's', v: 'a\r\nb' }), w(3, 3, { t: 's', v: 'a\rb' }), w(1, 1, { t: 's', v: 'x\ny' }),
    ])]);
    for (const ref of ['A3', 'B3']) {
      const xf = await xfOf(out, ref);
      expect(xf, ref).toContain('wrapText="1"');
      expect(xf, ref).toContain('applyAlignment="1"');
    }
    // Excel leaves WrapText alone for a text whose only line break is a CR (measured, golden q-wrap cell C7)
    expect((await xfOf(out, 'C3')) ?? '').not.toContain('wrapText');
    // A1 had its own style (fill / alignment centre + vertical centre): everything stays, wrap comes on top
    const a1 = await xfOf(out, 'A1');
    expect(a1).toContain('wrapText="1"');
    const before = await xfOf(await book(), 'A1');
    const attr = (x: string, n: string) => new RegExp(`\\s${n}="([^"]*)"`).exec(x)?.[1];
    for (const n of ['fillId', 'fontId', 'borderId', 'numFmtId']) expect(attr(a1, n), n).toBe(attr(before, n));
    expect(a1).toContain('horizontal="center"');
    expect(a1).toContain('vertical="center"');
    // the cells of one kind share ONE xf (derived once)
    const s = (await snapshot(out)).sheets[0];
    expect(s.styles.A3).toBe(s.styles.B3);
    expect(s.styles.A3).toContain('wrapText="1"');
  });

  it('a cell whose xf already wraps keeps its xf (no new xf), a text without line break never wraps', async () => {
    const first = await replayLog(await book(), [cells('S', [w(3, 1, { t: 's', v: 'a\nb' })])]);
    const n1 = await xfCount(first);
    const again = await replayLog(first, [cells('S', [w(3, 1, { t: 's', v: 'c\nd' }), w(4, 1, { t: 's', v: 'plain' })])]);
    expect(await xfCount(again)).toBe(n1);
    expect(await xfOf(again, 'A3')).toBe(await xfOf(first, 'A3'));
    expect(await xfOf(again, 'A4') ?? '').not.toContain('wrapText');
    const s = (await snapshot(again)).sheets[0];
    expect(s.values.A3).toBe('s:c\nd');
  });

  it('formulas, numbers, booleans and a cleared cell never wrap (only text constants)', async () => {
    const out = await replayLog(await book(), [cells('S', [
      w(3, 1, { t: 's', v: 'a\nb', f: '"a"&CHAR(10)&"b"' }), // a formula with a text result
      w(3, 2, { t: 'n', v: 12 }), w(3, 3, { t: 'b', v: true }), w(3, 4, null),
    ])]);
    for (const ref of ['A3', 'B3', 'C3']) expect(await xfOf(out, ref) ?? '', ref).not.toContain('wrapText');
  });

  it('with a number format (text format @) and with quote the wrap comes on top of both', async () => {
    const out = await replayLog(await book(), [cells('S', [
      w(3, 1, { t: 's', v: '1\n2', numFmt: '@' }), w(3, 2, { t: 's', v: '00\n1', quote: true }), w(3, 3, { t: 's', v: '0\n1', numFmt: '@', quote: true }),
    ])]);
    const a = await xfOf(out, 'A3'), b = await xfOf(out, 'B3'), c = await xfOf(out, 'C3');
    expect(a).toContain('wrapText="1"');
    expect(a).toContain('numFmtId="49"');
    expect(a).not.toContain('quotePrefix');
    expect(b).toContain('wrapText="1"');
    expect(b).toContain('quotePrefix="1"');
    expect(c).toContain('wrapText="1"');
    expect(c).toContain('numFmtId="49"');
    expect(c).toContain('quotePrefix="1"');
  });

  it('the text is stored the way Excel stores it: LF literal, CR as _x000D_, and it reads back unchanged', async () => {
    const out = await replayLog(await book(), [cells('S', [w(3, 1, { t: 's', v: 'a\r\nb' }), w(3, 2, { t: 's', v: 'l1\nl2' }), w(3, 3, { t: 's', v: 'tab\there' })])]);
    const sst = (await part(out, 'xl/sharedStrings.xml'))!;
    expect(sst).toContain('a_x000D_\nb');
    expect(sst).toContain('l1\nl2');
    expect(sst).not.toMatch(/\r/);
    const book2 = await readBook(out, 'x.xlsx');
    const rows = book2.sheets[0].rows;
    expect(rows[2]![0]!.v).toBe('a\r\nb');
    expect(rows[2]![1]!.v).toBe('l1\nl2');
    expect(rows[2]![2]!.v).toBe('tab\there');
    // the snapshot agrees
    expect((await snapshot(out)).sheets[0].values.A3).toBe('s:a\r\nb');
  });

  it('a column default style (no cell of its own yet) is wrapped too: the new cell takes the column xf first', async () => {
    const b = await buildXlsx([{ name: 'S', head: '<cols><col min="1" max="1" width="20" style="1" customWidth="1"/></cols>', data: '<row r="1"><c r="B1" t="s"><v>0</v></c></row>' }], { sst: ['x'] });
    const out = await replayLog(b, [cells('S', [w(2, 1, { t: 's', v: 'a\nb' })])]);
    const xf = await xfOf(out, 'A2');
    expect(xf).toContain('wrapText="1"');
    expect(xf).toContain('fillId="2"'); // column style 1 (bold, yellow fill, border) kept
  });

  it('Styles.withWrapText reuses an identical existing xf and is idempotent', () => {
    const st = new Styles(R3_STYLES);
    const w1 = st.withWrapText(0);
    expect(w1).toBeGreaterThan(0);
    expect(st.hasWrapText(0)).toBe(false);
    expect(st.hasWrapText(w1)).toBe(true);
    expect(st.withWrapText(w1)).toBe(w1);
    expect(st.withWrapText(0)).toBe(w1);
  });
});

describe('CellWrite.quote semantics (types.ts)', () => {
  const q = cells;

  it('quote === true adds quotePrefix to a text constant, every other write adds none', async () => {
    const out = await replayLog(await book(), [q('S', [
      w(3, 1, { t: 's', v: '00123', quote: true }), w(3, 2, { t: 's', v: '00123' }), w(3, 3, { t: 's', v: '00123', quote: false }),
      // ignored for numbers, booleans and formulas
      w(3, 4, { t: 'n', v: 5, quote: true }), w(3, 5, { t: 'b', v: true, quote: true }), w(3, 6, { t: 'z', v: null, f: '1+1', quote: true }),
    ])]);
    const has = async (ref: string) => /quotePrefix="1"/.test((await xfOf(out, ref)) ?? '');
    expect(await has('A3')).toBe(true);
    for (const ref of ['B3', 'C3', 'D3', 'E3', 'F3']) expect(await has(ref), ref).toBe(false);
  });

  it('a prefix the cell already had is CLEARED by a non-quote write (Excel COM keeps it: documented difference)', async () => {
    const first = await replayLog(await book(), [q('S', [w(3, 1, { t: 's', v: '00123', quote: true })])]);
    expect(await xfOf(first, 'A3')).toContain('quotePrefix="1"');
    const over = await replayLog(first, [q('S', [w(3, 1, { t: 's', v: 'abc' })])]);
    expect(await xfOf(over, 'A3') ?? '').not.toContain('quotePrefix');
    const num = await replayLog(first, [q('S', [w(3, 1, { t: 'n', v: 7 })])]);
    expect(await xfOf(num, 'A3') ?? '').not.toContain('quotePrefix');
    // a cleared cell keeps its style, prefix included
    const cleared = await replayLog(first, [q('S', [w(3, 1, null)])]);
    expect(await xfOf(cleared, 'A3')).toContain('quotePrefix="1"');
  });

  it('the code says so (the difference to Excel COM is documented where the decision is made)', async () => {
    const fs = await import('node:fs');
    const src = fs.readFileSync(new URL('../../src/core/xlsx/patch/engine.ts', import.meta.url), 'utf8');
    expect(src).toMatch(/DELIBERATE DIFFERENCE TO THE EXCEL ENGINE/);
    expect(src).toMatch(/Excel through COM keeps the flag/);
  });
});
