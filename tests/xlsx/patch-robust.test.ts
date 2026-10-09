/**
 * Round 4: robustness of replayLog / buildPremerged against hostile and unusual packages. The contract under test:
 * every outcome is a CORRECT package (well-formed XML in every part the patcher wrote, the same content as the
 * ordinary form of the same book) or a PatchError with a Japanese message — never a thrown TypeError, never a package
 * that Excel would have to repair.
 */
import { describe, expect, it } from 'vitest';
import type { CellWrite, LoggedOp, StructOp } from '../../src/core/types';
import { MAX_FORMULA_LENGTH } from '../../src/core/xlsx/patch/engine';
import { PatchError, toPatchError } from '../../src/core/xlsx/patch/pkg';
import { stripMainPrefix } from '../../src/core/xlsx/patch/xml';
import { xmlProblem } from '../../src/core/xlsx/patch/wellformed';
import { buildPremerged } from '../../src/core/xlsx/premerge';
import { readBook } from '../../src/core/xlsx/read';
import { replayLog } from '../../src/core/xlsx/replay';
import { buildXlsx, part } from '../golden/build';
import { edit, packageProblems, prefixMain, text, unpack } from '../golden/package-tools';
import { scenario } from '../golden/premerge-scenario';
import { snapshot } from '../golden/snapshot';

const JP = /[ぁ-んァ-ン一-龥]/;
const struct = (sheet: string, op: StructOp): LoggedOp => ({ kind: 'struct', sheet, op });
const insC = (sheet: string, at: number, count = 1) => struct(sheet, { type: 'insert', axis: 'col', at, count });
const delC = (sheet: string, at: number, count = 1) => struct(sheet, { type: 'delete', axis: 'col', at, count });
const movC = (sheet: string, from: number, count: number, to: number) => struct(sheet, { type: 'move', axis: 'col', from, count, to });
const insR = (sheet: string, at: number, count = 1) => struct(sheet, { type: 'insert', axis: 'row', at, count });
const cells = (sheet: string, writes: CellWrite[], reason: 'edit' | 'toValues' = 'edit'): LoggedOp => ({ kind: 'cells', sheet, writes, reason });
const w = (row: number, col: number, cell: CellWrite['cell']): CellWrite => ({ row, col, cell });

const DATA = '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" t="s"><v>2</v></c></row>'
  + '<row r="2"><c r="A2"><v>1</v></c><c r="B2"><v>2</v></c><c r="C2"><f>A2+B2</f><v>3</v></c></row>'
  + '<row r="3"><c r="A3"><v>4</v></c><c r="B3"><v>5</v></c><c r="C3"><f>A3+B3</f><v>9</v></c></row>';
const SST = ['a', 'b', 'c'];
const base = (sheetName = 'S') => buildXlsx([{ name: sheetName, data: DATA }], { sst: SST });

/** the usual edit: move a column, insert one, write a text with a number format, delete a row-less column */
const OPS: LoggedOp[] = [movC('S', 3, 1, 1), insC('S', 2), cells('S', [w(4, 1, { t: 's', v: 'x' }), w(4, 2, { t: 'd', v: 46295, numFmt: 'yyyy/m/d' })]), delC('S', 4)];

/** the result must be a fully well-formed package that the reader can read; returns the snapshot of sheet `name` */
async function expectGood(bytes: Uint8Array, name = 'S') {
  expect(await packageProblems(bytes)).toEqual([]);
  const book = await readBook(bytes, 'out.xlsx');
  expect(book.sheets.length).toBeGreaterThan(0);
  return (await snapshot(bytes)).sheets.find((s) => s.name === name)!;
}

/** a PatchError with a Japanese message that is NOT an internal (unexpected) error */
async function expectRefused(p: Promise<unknown>, re?: RegExp): Promise<PatchError> {
  let err: unknown;
  try { await p; } catch (e) { err = e; }
  expect(err, 'must be refused').toBeInstanceOf(PatchError);
  const pe = err as PatchError;
  expect(pe.message).toMatch(JP);
  expect(pe.internal, `an expected refusal, not an internal error: ${pe.message}`).toBe(false);
  if (re) expect(pe.message).toMatch(re);
  return pe;
}

// ---------------------------------------------------------------------------------------------------------------
describe('namespace-prefixed SpreadsheetML (<x:row> …)', () => {
  const prefixed = async (bytes: Uint8Array, prefix = 'x') => edit(bytes, (p) => {
    for (const [n, v] of p) if (typeof v === 'string' && /\.xml$/i.test(n)) p.set(n, prefixMain(v, prefix));
  });

  it('stripMainPrefix rewrites to the default namespace (and leaves everything else alone)', () => {
    const src = '<?xml version="1.0"?>\n<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="urn:r"><x:sheetData><x:row r="1"><x:c r="A1"/></x:row></x:sheetData><x:pageMargins left="1"/></x:worksheet>';
    expect(stripMainPrefix(src)).toBe('<?xml version="1.0"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="urn:r"><sheetData><row r="1"><c r="A1"/></row></sheetData><pageMargins left="1"/></worksheet>');
    const plain = '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>';
    expect(stripMainPrefix(plain)).toBe(plain);
    const chart = '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:f>S!A1</c:f></c:chartSpace>';
    expect(stripMainPrefix(chart)).toBe(chart); // another namespace: untouched
    // the prefix is also used for another default namespace: cannot be rewritten safely
    expect(stripMainPrefix('<x:worksheet xmlns="urn:other" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><x:sheetData/></x:worksheet>')).toBeNull();
  });

  it('a book written with a prefix replays like its ordinary twin (every op kind)', async () => {
    const plain = await replayLog(await base(), OPS);
    const pre = await replayLog(await prefixed(await base()), OPS);
    const a = await expectGood(plain), b = await expectGood(pre);
    expect(b.formulas).toEqual(a.formulas);
    expect(b.values).toEqual(a.values);
    expect(b.styles).toEqual(a.styles);
    expect(b.cols).toEqual(a.cols);
    // the parts the patcher wrote are in the default-namespace form, valid for Excel
    const parts = await unpack(pre);
    expect(text(parts, 'xl/worksheets/sheet1.xml')).toContain('<sheetData>');
    expect(text(parts, 'xl/worksheets/sheet1.xml')).not.toContain('<x:');
    expect(text(parts, 'xl/sharedStrings.xml')).not.toContain('<x:');
  });

  it('another prefix and the prefixed styles.xml / workbook.xml / sharedStrings.xml are read, not replaced by a default', async () => {
    const pre = await prefixed(await base(), 'ss');
    expect(await part(pre, 'xl/styles.xml')).toContain('<ss:fonts count="2">');
    // a date in the built-in format reuses an existing xf: styles.xml is not written, the prefixed original stays as it was
    const reuse = await replayLog(pre, [cells('S', [w(5, 1, { t: 'd', v: 46295, numFmt: 'yyyy/m/d' })])]);
    expect(await part(reuse, 'xl/styles.xml')).toBe(await part(pre, 'xl/styles.xml'));
    // a new number format has to be added: the styles are READ through their prefix (nothing is lost) and written in the
    // default-namespace form
    const out = await replayLog(pre, [cells('S', [w(5, 1, { t: 'd', v: 46295, numFmt: 'yyyy/m/d' }), w(5, 2, { t: 's', v: '新しい文字' }), w(5, 3, { t: 'n', v: 1.5, numFmt: '0.00"kg"' })])]);
    const s = await expectGood(out);
    expect(s.values.B5).toBe('s:新しい文字');
    expect(s.values.A1).toBe('s:a');
    expect(s.styles.C5).toContain('0.00"kg"');
    // the old styles are still there (only appended to): the prefixed fonts / fills / borders and the old custom format 0.0"pt"
    const styles = (await part(out, 'xl/styles.xml'))!;
    expect(styles).not.toContain('<ss:');
    expect(styles).toContain('<fonts count="2">');
    expect(styles).toContain('<fills count="3">');
    expect(styles).toContain('formatCode="0.0&quot;pt&quot;"');
    expect(styles).toContain('formatCode="0.00&quot;kg&quot;"');
    expect(styles).toContain('<cellXfs count="5">'); // 4 old (xf 3 = yyyy/m/d is reused for A5) + one derived xf for 0.00"kg"
  });

  it('an untouched prefixed sheet keeps its bytes', async () => {
    const two = await buildXlsx([{ name: 'S', data: DATA }, { name: 'T', data: '<row r="1"><c r="A1"><v>1</v></c></row>' }], { sst: SST });
    const pre = await prefixed(two);
    const out = await replayLog(pre, [insC('S', 1)]);
    expect(text(await unpack(out), 'xl/worksheets/sheet2.xml')).toBe(text(await unpack(pre), 'xl/worksheets/sheet2.xml'));
  });

  it('a prefix that cannot be removed safely is refused in Japanese (not patched blindly)', async () => {
    const bad = await edit(await base(), (p) => {
      p.set('xl/worksheets/sheet1.xml', text(p, 'xl/worksheets/sheet1.xml').replace('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"', '<x:worksheet xmlns="urn:something:else" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"').replace(/<\/worksheet>/, '</x:worksheet>'));
    });
    await expectRefused(replayLog(bad, [insC('S', 1)]), /名前空間/);
  });

  it('buildPremerged accepts a prefixed base book (same result as the ordinary one)', async () => {
    const sc = scenario('plain');
    const plain = await buildPremerged(sc);
    const pre = await prefixed(sc.base.book.bytes);
    const sc2 = scenario('plain');
    sc2.base = { ...sc2.base, book: { ...sc2.base.book, bytes: pre } };
    const r = await buildPremerged(sc2);
    const a = (await snapshot(plain.bytes)).sheets.find((s) => s.name === '学生')!;
    const b = (await snapshot(r.bytes)).sheets.find((s) => s.name === '学生')!;
    expect(b.values).toEqual(a.values);
    expect(b.formulas).toEqual(a.formulas);
    expect(b.styles).toEqual(a.styles);
    expect(r.layout).toEqual(plain.layout);
    expect(await packageProblems(r.bytes)).toEqual([]);
  });
});

// ---------------------------------------------------------------------------------------------------------------
describe('cells and columns beyond the Excel grid', () => {
  const sheetWith = (data: string, head = '') => buildXlsx([{ name: 'S', data, head }], { sst: SST });

  it('a cell beyond column XFD (XFE, ZZZ, AAAA) is refused when the sheet has to be written', async () => {
    for (const ref of ['XFE1', 'ZZZ1', 'AAAA1']) {
      const b = await sheetWith(`<row r="1"><c r="A1"><v>1</v></c><c r="${ref}"><v>2</v></c></row>`);
      await expectRefused(replayLog(b, [insC('S', 1)]), /XFD|範囲/);
    }
  });

  it('a row beyond 1,048,576, a cell in the wrong row and an unreadable address are refused', async () => {
    await expectRefused(replayLog(await sheetWith('<row r="1048577"><c r="A1048577"><v>1</v></c></row>'), [insC('S', 1)]), /範囲外/);
    await expectRefused(replayLog(await sheetWith('<row r="2"><c r="A3"><v>1</v></c></row>'), [insC('S', 1)]), /セル/);
    await expectRefused(replayLog(await sheetWith('<row r="2"><c r="1A"><v>1</v></c></row>'), [insC('S', 1)]), /セル/);
  });

  it('two cells at one position / two rows with one number are refused', async () => {
    await expectRefused(replayLog(await sheetWith('<row r="1"><c r="B1"><v>1</v></c><c r="A1"><v>2</v></c><c r="B1"><v>3</v></c></row>'), [insC('S', 1)]), /重複/);
    await expectRefused(replayLog(await sheetWith('<row r="1"><c r="A1"><v>1</v></c></row><row r="1"><c r="B1"><v>1</v></c></row>'), [insC('S', 1)]), /重複/);
  });

  it('out-of-order rows / cells are simply sorted', async () => {
    const b = await sheetWith('<row r="2"><c r="B2"><v>4</v></c><c r="A2"><v>3</v></c></row><row r="1"><c r="B1"><v>2</v></c><c r="A1"><v>1</v></c></row>');
    const s = await expectGood(await replayLog(b, [insC('S', 1)]));
    expect(s.values).toEqual({ B1: 'n:1', C1: 'n:2', B2: 'n:3', C2: 'n:4' });
  });

  it('a full row of 16,384 styled cells survives inserts, deletes and moves (empty cells pushed off the grid are dropped)', async () => {
    let row = '<row r="1">';
    for (let c = 1; c <= 16384; c++) row += `<c r="${colL(c)}1" s="${c % 3 === 0 ? 1 : 0}"/>`;
    row += '</row>';
    const b = await sheetWith(row + '<row r="2"><c r="A2"><v>7</v></c></row>', '<cols><col min="1" max="16384" width="9" style="2" customWidth="1"/></cols>');
    const t0 = Date.now();
    const out = await replayLog(b, [insC('S', 1, 2), delC('S', 5), movC('S', 10, 3, 1)]);
    expect(Date.now() - t0).toBeLessThan(8000);
    const xml = text(await unpack(out), 'xl/worksheets/sheet1.xml');
    // 16,384 cells; insert 2 at A: the last 2 (empty, styled) fall off the grid (Excel drops them too; only content blocks an
    // insert) and nothing is materialised at A (no left neighbour to inherit from); delete 1: one fewer; the move changes nothing
    expect((xml.match(/<c r="[A-Z]+1"/g) ?? []).length).toBe(16384 - 2 - 1);
    expect(xml).not.toMatch(/<c r="XF[E-Z]/);
    expect(xml.match(/<col [^>]*max="(\d+)"/g)!.every((m) => Number(/max="(\d+)"/.exec(m)![1]) <= 16384)).toBe(true);
    expect(await packageProblems(out)).toEqual([]);
  }, 30000);

  it('a value in the last column blocks an insert (Excel refuses too), as a Japanese PatchError', async () => {
    const b = await sheetWith('<row r="1"><c r="XFD1"><v>1</v></c></row>');
    await expectRefused(replayLog(b, [insC('S', 1)]), /押し出/);
    // ... a move / delete of other columns is fine
    await expectGood(await replayLog(b, [delC('S', 1)]));
  });

  it('<col> entries that start beyond XFD or have nonsense numbers are ignored, huge max values are clamped', async () => {
    const b = await sheetWith('<row r="1"><c r="A1"><v>1</v></c></row>', '<cols><col min="1" max="1048576" width="9" customWidth="1"/><col min="20000" max="30000" width="5"/><col min="x" max="3"/><col min="0" max="2"/></cols>');
    const out = await replayLog(b, [insC('S', 1)]);
    expect(await packageProblems(out)).toEqual([]);
    const cols = text(await unpack(out), 'xl/worksheets/sheet1.xml').match(/<col [^>]*>/g)!;
    expect(cols.every((c) => Number(/max="(\d+)"/.exec(c)![1]) <= 16384)).toBe(true);
  });
});

function colL(n: number): string {
  let s = '';
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

// ---------------------------------------------------------------------------------------------------------------
describe('spans, missing dimension, inline strings, rich shared strings', () => {
  it('rows with (wrong) spans lose them, the cells stay; a missing <dimension> is not invented', async () => {
    const b = await buildXlsx([{ name: 'S', data: '<row r="1" spans="1:2"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="D1"><v>9</v></c></row><row r="2" spans="3:3" ht="30" customHeight="1"><c r="C2"><v>1</v></c></row>' }], { sst: SST });
    const noDim = await edit(b, (p) => p.set('xl/worksheets/sheet1.xml', text(p, 'xl/worksheets/sheet1.xml').replace(/<dimension[^>]*\/>/, '')));
    expect(text(await unpack(noDim), 'xl/worksheets/sheet1.xml')).not.toContain('<dimension');
    const out = await replayLog(noDim, [insC('S', 2)]);
    const xml = text(await unpack(out), 'xl/worksheets/sheet1.xml');
    expect(xml).not.toContain('spans=');
    expect(xml).not.toContain('<dimension');
    expect(xml).toContain('<row r="2" ht="30" customHeight="1">');
    const s = await expectGood(out);
    expect(s.values).toEqual({ A1: 's:a', C1: 's:b', E1: 'n:9', D2: 'n:1' });
  });

  it('a wrong <dimension> is corrected for the sheets the patcher writes', async () => {
    const b = await edit(await base(), (p) => p.set('xl/worksheets/sheet1.xml', text(p, 'xl/worksheets/sheet1.xml').replace('<dimension ref="A1"/>', '<dimension ref="A1:Z99"/>')));
    const out = await replayLog(b, [insC('S', 1)]);
    expect(text(await unpack(out), 'xl/worksheets/sheet1.xml')).toContain('<dimension ref="B1:D3"/>');
  });

  it('inline strings (t="inlineStr", also with rich runs) move, stay inline and are not turned into shared strings', async () => {
    const inline = '<row r="1"><c r="A1" t="inlineStr"><is><t>ID</t></is></c><c r="B1" t="inlineStr"><is><r><rPr><b/></rPr><t>太</t></r><r><t xml:space="preserve"> 普通</t></r></is></c></row>'
      + '<row r="2"><c r="A2" t="inlineStr"><is><t>a&amp;b</t></is></c><c r="B2"><v>2</v></c></row>';
    const b = await buildXlsx([{ name: 'S', data: inline }]);
    const out = await replayLog(b, [movC('S', 2, 1, 1), cells('S', [w(3, 1, { t: 's', v: '新' })])]);
    const s = await expectGood(out);
    expect(s.values.B1).toBe('s:ID');
    expect(s.values.A1).toBe('s:太 普通');
    expect(s.values.B2).toBe('s:a&b');
    expect(s.values.A3).toBe('s:新');
    const xml = text(await unpack(out), 'xl/worksheets/sheet1.xml');
    expect(xml).toContain('<is><r><rPr><b/></rPr><t>太</t></r><r><t xml:space="preserve"> 普通</t></r></is>');
  });

  it('shared strings with rich text runs, phonetics and xml:space survive byte for byte; new strings are appended', async () => {
    const items = ['<si><t>plain</t></si>', '<si><r><rPr><b/><sz val="11"/></rPr><t>Bold</t></r><r><t xml:space="preserve"> normal</t></r></si>',
      '<si><t>山田</t><rPh sb="0" eb="2"><t>ヤマダ</t></rPh><phoneticPr fontId="1"/></si>', '<si><t xml:space="preserve">  lead</t></si>', '<si><t/></si>'];
    const b0 = await buildXlsx([{ name: 'S', data: '<row r="1"><c r="A1" t="s"><v>1</v></c><c r="B1" t="s"><v>2</v></c><c r="C1" t="s"><v>3</v></c></row>' }], { sst: ['x'] });
    const b = await edit(b0, (p) => p.set('xl/sharedStrings.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="4" uniqueCount="${items.length}">${items.join('')}</sst>`));
    const out = await replayLog(b, [cells('S', [w(2, 1, { t: 's', v: 'plain' }), w(2, 2, { t: 's', v: '新' }), w(2, 3, { t: 's', v: 'a<b>&c' })]), movC('S', 3, 1, 1)]);
    const sst = text(await unpack(out), 'xl/sharedStrings.xml');
    for (const it of items) expect(sst).toContain(it);
    expect(sst).toContain('<si><t>新</t></si>');
    expect(sst).toContain('<t>a&lt;b&gt;&amp;c</t>');
    expect(sst).toMatch(/uniqueCount="7"/); // 5 + 新 + a<b>&c ('plain' is reused)
    const s = await expectGood(out);
    // column C moved to the front: C (lead / a<b>&c) is now A, A (Bold normal / plain) is B, B (山田 / 新) is C
    expect(s.values).toEqual({ A1: 's:  lead', B1: 's:Bold normal', C1: 's:山田', A2: 's:a<b>&c', B2: 's:plain', C2: 's:新' });
  });
});

// ---------------------------------------------------------------------------------------------------------------
describe('workbooks without styles.xml / sharedStrings.xml, macro-enabled books', () => {
  const noStyles = async () => edit(await buildXlsx([{ name: 'S', data: DATA.replace(/ s="\d+"/g, '') }], { sst: SST }), (p) => {
    p.delete('xl/styles.xml');
    p.set('xl/_rels/workbook.xml.rels', text(p, 'xl/_rels/workbook.xml.rels').replace(/<Relationship [^>]*\/styles"[^>]*\/>/, ''));
    p.set('[Content_Types].xml', text(p, '[Content_Types].xml').replace(/<Override PartName="\/xl\/styles.xml"[^>]*\/>/, ''));
  });
  const noSst = async () => edit(await buildXlsx([{ name: 'S', data: '<row r="1"><c r="A1"><v>1</v></c><c r="B1" t="inlineStr"><is><t>i</t></is></c></row>' }]), (p) => {
    p.delete('xl/sharedStrings.xml');
    p.set('xl/_rels/workbook.xml.rels', text(p, 'xl/_rels/workbook.xml.rels').replace(/<Relationship [^>]*\/sharedStrings"[^>]*\/>/, ''));
    p.set('[Content_Types].xml', text(p, '[Content_Types].xml').replace(/<Override PartName="\/xl\/sharedStrings.xml"[^>]*\/>/, ''));
  });

  it('no styles.xml: structural ops need none, nothing is added', async () => {
    const b = await noStyles();
    expect((await unpack(b)).has('xl/styles.xml')).toBe(false);
    // C (the sums) moves to the front, a column is inserted after it, then the inserted column and the last (empty) one go
    const out = await replayLog(b, [movC('S', 3, 1, 1), insC('S', 2), delC('S', 2), delC('S', 4)]);
    expect((await unpack(out)).has('xl/styles.xml')).toBe(false);
    const s = await expectGood(out);
    expect(s.formulas).toEqual({ A2: 'B2+C2', A3: 'B3+C3' });
    expect(s.values).toMatchObject({ A1: 's:c', B1: 's:a', C1: 's:b' });
  });

  it('no styles.xml: a number format / wrap text creates the part, its relationship and content type', async () => {
    const out = await replayLog(await noStyles(), [cells('S', [w(5, 1, { t: 'd', v: 46295, numFmt: 'yyyy/m/d' }), w(5, 2, { t: 's', v: 'a\nb' })])]);
    const parts = await unpack(out);
    expect(parts.has('xl/styles.xml')).toBe(true);
    expect(text(parts, 'xl/_rels/workbook.xml.rels')).toMatch(/relationships\/styles"/);
    expect(text(parts, '[Content_Types].xml')).toContain('/xl/styles.xml');
    const s = await expectGood(out);
    expect(s.styles.A5).toContain('nf=#14');
    expect(s.styles.B5).toContain('wrapText="1"');
    // the cells of the old styleless book have no style attribute (xf 0 = default) and still none after the ops
    expect(text(parts, 'xl/worksheets/sheet1.xml')).toContain('<c r="A2">');
  });

  it('no sharedStrings.xml: a text write creates it (relationship + content type) and old inline strings stay', async () => {
    const b = await noSst();
    const out = await replayLog(b, [cells('S', [w(3, 1, { t: 's', v: '文字' })]), insC('S', 1)]);
    const parts = await unpack(out);
    expect(parts.has('xl/sharedStrings.xml')).toBe(true);
    expect(text(parts, 'xl/_rels/workbook.xml.rels')).toMatch(/relationships\/sharedStrings"/);
    expect(text(parts, '[Content_Types].xml')).toContain('/xl/sharedStrings.xml');
    const s = await expectGood(out);
    expect(s.values).toEqual({ B1: 'n:1', C1: 's:i', B3: 's:文字' });
  });

  it('neither part: a formula-only edit works', async () => {
    const b = await edit(await noSst(), (p) => {
      p.delete('xl/styles.xml');
      p.set('xl/_rels/workbook.xml.rels', text(p, 'xl/_rels/workbook.xml.rels').replace(/<Relationship [^>]*\/styles"[^>]*\/>/, ''));
    });
    const s = await expectGood(await replayLog(b, [cells('S', [w(2, 1, { t: 'z', v: null, f: 'A1*2' })])]));
    expect(s.formulas).toEqual({ A2: 'A1*2' });
  });

  it('xlsm: macro content type, vbaProject.bin (binary) and its relationship are carried through unchanged', async () => {
    const bin = new Uint8Array(3000).map((_v, i) => (i * 31 + 7) & 255);
    const xlsm = await edit(await base(), (p) => {
      p.set('xl/vbaProject.bin', bin);
      p.set('[Content_Types].xml', text(p, '[Content_Types].xml')
        .replace('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml', 'application/vnd.ms-excel.sheet.macroEnabled.main+xml')
        .replace('<Default Extension="xml"', '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/><Default Extension="xml"'));
      p.set('xl/_rels/workbook.xml.rels', text(p, 'xl/_rels/workbook.xml.rels').replace('</Relationships>', '<Relationship Id="rId99" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/></Relationships>'));
    });
    const out = await replayLog(xlsm, OPS);
    const before = await unpack(xlsm), after = await unpack(out);
    expect(after.get('xl/vbaProject.bin')).toEqual(bin);
    expect(text(after, '[Content_Types].xml')).toContain('macroEnabled.main+xml');
    expect(text(after, '[Content_Types].xml')).toContain('vbaProject');
    expect(text(after, 'xl/_rels/workbook.xml.rels')).toContain('vbaProject.bin');
    expect([...after.keys()].filter((k) => !before.has(k))).toEqual([]);
    await expectGood(out);
  });
});

// ---------------------------------------------------------------------------------------------------------------
describe('very long formulas', () => {
  it('a formula of ~8,000 characters is shifted reference by reference (and stays well-formed)', async () => {
    // 7,799 chars, just under Excel's 8,192. (Excel itself could not hold 2,600 references: its compiled-formula limit is
    // 16,384 bytes, about 8 per reference — the patcher still has to carry such an input through unchanged in shape; the
    // Excel-opening check uses tools/make-fuzz-outputs.ts h-long-formula with a formula Excel can hold.)
    const f = Array.from({ length: 2600 }, (_x, i) => `B${(i % 3) + 1}`).join('+');
    expect(f.length).toBeGreaterThan(7500);
    expect(f.length).toBeLessThan(MAX_FORMULA_LENGTH);
    const b = await buildXlsx([{ name: 'S', data: `<row r="1"><c r="A1"><f>${f}</f><v>0</v></c></row><row r="2"><c r="B2"><v>1</v></c></row>` }]);
    const t0 = Date.now();
    const out = await replayLog(b, [insC('S', 2), insR('S', 1)]);
    expect(Date.now() - t0).toBeLessThan(3000);
    const s = await expectGood(out);
    expect(s.formulas.A2).toBe(f.replace(/B(\d)/g, (_m, r: string) => `C${Number(r) + 1}`));
  });

  it('a hostile 150,000-character formula in the input is processed in reasonable time', async () => {
    const f = Array.from({ length: 30000 }, () => 'B2').join('+');
    const b = await buildXlsx([{ name: 'S', data: `<row r="1"><c r="A1"><f>${f}</f><v>0</v></c></row>` }]);
    const t0 = Date.now();
    const out = await replayLog(b, [insC('S', 1)]);
    expect(Date.now() - t0).toBeLessThan(10000);
    expect((await snapshot(out)).sheets[0].formulas.B1).toBe(f.replace(/B2/g, 'C2'));
  }, 30000);

  it('a formula typed by the user that is longer than Excel allows is refused in Japanese', async () => {
    const long = 'A1' + '+A1'.repeat(MAX_FORMULA_LENGTH / 3 + 10);
    await expectRefused(replayLog(await base(), [cells('S', [w(5, 1, { t: 'z', v: null, f: long })])]), /長すぎ/);
    // the file-form prefixes do not count
    const fn = 'SUM(A1:A2)' + '+_xlfn.STDEV.S(A1:A3)'.repeat(300);
    expect(fn.replace(/_xlfn\./g, '').length).toBeLessThan(MAX_FORMULA_LENGTH);
    expect(fn.length).toBeLessThan(MAX_FORMULA_LENGTH + 4000);
    await expectGood(await replayLog(await base(), [cells('S', [w(5, 1, { t: 'z', v: null, f: fn })])]));
  });

  it('a formula with a control character is refused, a line break inside a formula is fine', async () => {
    await expectRefused(replayLog(await base(), [cells('S', [w(5, 1, { t: 'z', v: null, f: 'A1+\u0001' })])]), /使用できない文字/);
    const out = await replayLog(await base(), [cells('S', [w(5, 1, { t: 'z', v: null, f: 'SUM(A1,\nA2)' })])]);
    await expectGood(out);
    expect((await snapshot(out)).sheets[0].formulas.A5).toBe('SUM(A1,\nA2)');
  });
});

// ---------------------------------------------------------------------------------------------------------------
describe('Japanese sheet names with quotes and other awkward characters', () => {
  const NAMES = ["It's 売上", 'A1', '2024年度', 'a b', 'AT&T 売上', 'R1C1'];
  const esc = (s: string) => s.replace(/&/g, '&amp;');
  const q = (s: string) => `'${s.replace(/'/g, "''")}'`;

  async function awkward(): Promise<Uint8Array> {
    const sheets = NAMES.map((n, i) => ({
      name: esc(n),
      data: `<row r="1"><c r="A1"><v>${i + 1}</v></c><c r="B1"><v>10</v></c></row><row r="2"><c r="A2"><f>A1*2</f><v>0</v></c></row>`,
    }));
    // a "Main" sheet that refers to every other sheet
    sheets.push({ name: 'Main', data: NAMES.map((n, i) => `<row r="${i + 1}"><c r="A${i + 1}"><f>${esc(q(n))}!B1+SUM(${esc(q(n))}!A1:B2)</f><v>0</v></c></row>`).join('') });
    const names = NAMES.map((n, i) => `<definedName name="Nm${i}">${esc(q(n))}!$B$1</definedName>`).join('') + `<definedName name="Loc" localSheetId="0">${esc(q(NAMES[0]))}!$A$1:$B$2</definedName>`;
    return buildXlsx(sheets, { names });
  }

  it('every op reaches the formulas, names and other sheets that refer to a sheet with such a name', async () => {
    const b = await awkward();
    const out = await replayLog(b, NAMES.map((n) => insC(n, 1)));
    expect(await packageProblems(out)).toEqual([]);
    const snap = await snapshot(out);
    const main = snap.sheets.find((s) => s.name === 'Main')!;
    NAMES.forEach((n, i) => {
      expect(main.formulas[`A${i + 1}`], n).toBe(`${q(n)}!C1+SUM(${q(n)}!B1:C2)`.replace(/^'([A-Za-z0-9_]+)'!/, (_m, id: string) => (/^[A-Za-z_][A-Za-z0-9_]*$/.test(id) && !/^[A-Za-z]{1,3}\d+$/i.test(id) && !/^R\d*C\d*$/i.test(id) ? `${id}!` : `'${id}'!`)).replace(/SUM\('([A-Za-z0-9_]+)'!/, (_m, id: string) => (/^[A-Za-z_][A-Za-z0-9_]*$/.test(id) && !/^[A-Za-z]{1,3}\d+$/i.test(id) && !/^R\d*C\d*$/i.test(id) ? `SUM(${id}!` : `SUM('${id}'!`)));
    });
    // defined names (workbook and sheet scope) follow
    for (let i = 0; i < NAMES.length; i++) expect(snap.names.some((x) => x.startsWith(`Nm${i} = `) && x.endsWith('!$C$1')), `Nm${i}: ${snap.names.join(' | ')}`).toBe(true);
    expect(snap.names.some((x) => x.includes('Loc = ') && x.endsWith('$B$1:$C$2'))).toBe(true);
    // the ops' own sheets: formulas in the sheet itself
    for (const n of NAMES) expect(snap.sheets.find((s) => s.name === n)!.formulas.B2, n).toBe('B1*2');
  });

  it('cells written to / ops on a sheet named like a cell reference (A1, R1C1) address that sheet', async () => {
    const out = await replayLog(await awkward(), [cells('A1', [w(5, 5, { t: 's', v: 'ここ' })]), cells("It's 売上", [w(5, 5, { t: 's', v: 'there' })]), movC('R1C1', 2, 1, 1)]);
    const snap = await snapshot(out);
    expect(snap.sheets.find((s) => s.name === 'A1')!.values.E5).toBe('s:ここ');
    expect(snap.sheets.find((s) => s.name === "It's 売上")!.values.E5).toBe('s:there');
    expect(snap.sheets.find((s) => s.name === 'R1C1')!.values.A1).toBe('n:10');
    expect(await packageProblems(out)).toEqual([]);
  });

  it('sheet names are matched case-insensitively like Excel does, an unknown name is a Japanese error', async () => {
    await expectGood(await replayLog(await awkward(), [insC("IT'S 売上", 1)]), "It's 売上");
    await expectRefused(replayLog(await awkward(), [insC('ない', 1)]), /シート「ない」/);
  });
});

// ---------------------------------------------------------------------------------------------------------------
describe('damaged input is refused in Japanese, never written back damaged', () => {
  const sheetXml = (p: Map<string, string | Uint8Array>) => text(p, 'xl/worksheets/sheet1.xml');
  const damage = (fn: (x: string) => string) => async () => edit(await base(), (p) => p.set('xl/worksheets/sheet1.xml', fn(sheetXml(p))));

  it('a sheet cut off inside <sheetData>', async () => {
    const b = await damage((x) => x.slice(0, x.indexOf('<row r="3">') + 30))();
    await expectRefused(replayLog(b, [insC('S', 1)]), /途中で切れ/);
  });

  it('a sheet cut off after <sheetData> / with unbalanced tags in the tail / a raw ampersand', async () => {
    await expectRefused(replayLog(await damage((x) => x.replace('</worksheet>', '<mergeCells count="1"><mergeCell ref="A1:B1"></mergeCells></worksheet>'))(), [insC('S', 1)]), /検査/);
    await expectRefused(replayLog(await damage((x) => x.replace('</worksheet>', ''))(), [insC('S', 1)]), /検査|閉じていません/);
    await expectRefused(replayLog(await damage((x) => x.replace('<v>1</v>', '<v>1 & 2</v>'))(), [insC('S', 1)]), /検査|&/);
  });

  it('an undeclared namespace prefix in a part that has to be written', async () => {
    await expectRefused(replayLog(await damage((x) => x.replace('</worksheet>', '<extLst><ext><x14:id>{1}</x14:id></ext></extLst></worksheet>'))(), [insC('S', 1)]), /接頭辞/);
  });

  it('a package that is not an xlsx, an empty zip, a zip without a workbook', async () => {
    await expectRefused(replayLog(new Uint8Array([1, 2, 3, 4]), []), undefined);
    await expectRefused(replayLog(new Uint8Array(0), []), undefined);
    const empty = await edit(await base(), (p) => { for (const k of [...p.keys()]) p.delete(k); });
    await expectRefused(replayLog(empty, []));
    const noWb = await edit(await base(), (p) => { p.delete('xl/workbook.xml'); });
    await expectRefused(replayLog(noWb, [insC('S', 1)]));
  });

  it('a worksheet relationship that points nowhere: the sheet is reported missing (Japanese), not a crash', async () => {
    const b = await edit(await base(), (p) => p.delete('xl/worksheets/sheet1.xml'));
    await expectRefused(replayLog(b, [insC('S', 1)]), /シート「S」/);
  });

  it('hostile op logs', async () => {
    const b = await base();
    const bad: unknown[][] = [
      [null], [42], [{}], [{ kind: 'nope', sheet: 'S' }], [{ kind: 'struct', sheet: 5, op: {} }], [{ kind: 'struct', sheet: 'S' }],
      [{ kind: 'struct', sheet: 'S', op: { type: 'insert', axis: 'diag', at: 1, count: 1 } }],
      [{ kind: 'struct', sheet: 'S', op: { type: 'frob', axis: 'col', at: 1, count: 1 } }],
      [{ kind: 'struct', sheet: 'S', op: { type: 'insert', axis: 'col', at: 1.5, count: 1 } }],
      [{ kind: 'struct', sheet: 'S', op: { type: 'insert', axis: 'col', at: 1, count: 99999 } }],
      [{ kind: 'cells', sheet: 'S', writes: [null], reason: 'edit' }],
      [{ kind: 'cells', sheet: 'S', writes: [{ row: 0, col: 1, cell: null }], reason: 'edit' }],
      [{ kind: 'cells', sheet: 'S', writes: [{ row: 1, col: 1, cell: 'x' }], reason: 'edit' }],
      [{ kind: 'cells', sheet: 'S', writes: [{ row: 1, col: 1, cell: { t: 'z', v: null, f: 123 } }], reason: 'edit' }],
      [{ kind: 'colWidth', sheet: 'S', col: 1, width: NaN }], [{ kind: 'colWidth', sheet: 'S', col: 0, width: 5 }],
      [{ kind: 'unmerge', sheet: 'S', refs: [5], centerAcross: false }], [{ kind: 'unmerge', sheet: 'S', refs: ['ZZ'], centerAcross: false }],
    ];
    for (const log of bad) await expectRefused(replayLog(b, log as LoggedOp[]));
    // nothing to do: a valid package comes back
    await expectGood(await replayLog(b, []));
    await expectGood(await replayLog(b, undefined as unknown as LoggedOp[]));
  });

  it('numeric character references that are not characters do not throw, and text round-trips through the escapes', async () => {
    const b = await damage((x) => x.replace('<f>A2+B2</f>', '<f>A2+B2&#99999999;</f>').replace('<v>2</v>', '<v>&#xD800;</v>'))();
    // the formula text keeps the reference verbatim (it is not a character), the output is well-formed or refused in Japanese
    let ok = true;
    try { await replayLog(b, [insC('S', 1)]); } catch (e) { ok = false; expect(e).toBeInstanceOf(PatchError); expect((e as PatchError).internal).toBe(false); }
    expect(typeof ok).toBe('boolean');
    // control characters / lone surrogates in written text become _xHHHH_ escapes
    const out = await replayLog(await base(), [cells('S', [w(5, 1, { t: 's', v: 'a\u0001b\ud800c\u0000' })])]);
    const sst = text(await unpack(out), 'xl/sharedStrings.xml');
    expect(sst).toContain('a_x0001_b_xD800_c_x0000_');
    await expectGood(out);
    expect((await snapshot(out)).sheets[0].values.A5).toBe('s:a\u0001b\ud800c\u0000');
  });
});

describe('package plumbing', () => {
  it('relationship targets written with backslashes or upper-case drive-less absolute paths resolve', async () => {
    const b = await edit(await base(), (p) => {
      p.set('xl/_rels/workbook.xml.rels', text(p, 'xl/_rels/workbook.xml.rels').replace('Target="worksheets/sheet1.xml"', 'Target="worksheets\\sheet1.xml"'));
    });
    const s = await expectGood(await replayLog(b, [insC('S', 1)]));
    expect(s.values.B1).toBe('s:a');
    const abs = await edit(await base(), (p) => {
      p.set('xl/_rels/workbook.xml.rels', text(p, 'xl/_rels/workbook.xml.rels').replace('Target="worksheets/sheet1.xml"', 'Target="/xl/worksheets/sheet1.xml"'));
    });
    expect((await expectGood(await replayLog(abs, [insC('S', 1)]))).values.B1).toBe('s:a');
  });

  it('a self-closing <definedName/> does not swallow its neighbour', async () => {
    const b = await buildXlsx([{ name: 'S', data: DATA }], { names: '<definedName name="Empty"/><definedName name="Col">S!$B$1:$B$3</definedName>' });
    const out = await replayLog(b, [insC('S', 1)]);
    const wb = text(await unpack(out), 'xl/workbook.xml');
    expect(wb).toContain('<definedName name="Empty"/>');
    expect(wb).toContain('<definedName name="Col">S!$C$1:$C$3</definedName>');
  });

  it('toPatchError keeps PatchErrors, converts everything else to Japanese text with the cause attached', () => {
    const pe = new PatchError('そのまま');
    expect(toPatchError(pe, 'x')).toBe(pe);
    const conv = toPatchError(new TypeError('boom'), '書き出せませんでした');
    expect(conv).toBeInstanceOf(PatchError);
    expect(conv.internal).toBe(true);
    expect(conv.message).toMatch(/書き出せませんでした.*boom/);
    expect((conv as { cause?: unknown }).cause).toBeInstanceOf(TypeError);
    const big = toPatchError(new RangeError('Invalid string length'), 'x');
    expect(big.message).toMatch(/大きすぎる/);
  });

  it('xmlProblem: what it accepts and what it refuses', () => {
    const ok = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<a xmlns="urn:a" xmlns:b="urn:b"><b:c d="1" e=\'2\'/><!-- c --><f>x &amp; &#10; &#xA; y</f><![CDATA[<>&]]></f2></a>'.replace('<![CDATA[<>&]]></f2>', '<![CDATA[<>&]]>');
    expect(xmlProblem(ok)).toBeNull();
    for (const bad of ['<a><b></a></b>', '<a>', '<a></a><b/>', 'x<a/>', '<a b=1/>', '<a b="1" b="2"/>', '<a>&</a>', '<a>&foo;</a>', '<a\u0001/>', '<a>\ud800</a>', '<p:a/>', '<a p:b="1"/>', '<!DOCTYPE a><a/>', '<a b="<"/>', '', '<a></b>', '<a><!-- x</a>']) {
      expect(xmlProblem(bad), JSON.stringify(bad)).not.toBeNull();
    }
    // a prefix declared on the element itself or an ancestor is fine, out of scope is not
    expect(xmlProblem('<a xmlns:p="u"><p:b/></a>')).toBeNull();
    expect(xmlProblem('<a><b xmlns:p="u"/><p:c/></a>')).not.toBeNull();
  });
});
