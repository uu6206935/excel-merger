// Round-4 doc parity: ColProps.customWidth = "set by the user WITHOUT bestFit" (the reader drops the flag of an
// auto-fitted column), so the doc widens auto AND auto-fitted (bestFit) columns exactly like the JS patcher, and a width
// the user sets (colWidth op) ends auto-fit in both.
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { createEditor, deleteColumns, insertColumns, moveColumns, setCellInput, setColumnWidth } from '../../src/core/doc';
import { readBook, replayLog } from '../../src/core/xlsx';
import { compare, golden } from './compare';
import { recol } from './fuzzcore';
import { ok } from './helpers';

// g-table 名簿: A..F carry customWidth="1" (user widths); J and K share <col min="10" max="11" width="9" outlineLevel="1"/>
// (an auto width: no customWidth) and hold nothing under row 13, so dates typed into J15 / J16 are the only widening
// candidates. 'auto' strips the flags of every <col>, 'bestFit' makes them bestFit="1" customWidth="1" (auto-fitted by
// Excel), 'user' gives J/K customWidth="1" (a width the user set).
const REGION = { headerRow: 3, firstDataRow: 4, lastDataRow: 13 };
const LAYOUT = { sheet: '名簿', insertAt: 0, imported: [], appendAt: null, appended: [] };
const J = 9;

async function editorOf(mode: 'auto' | 'bestFit' | 'user') {
  const raw = golden('g-table.xlsx');
  let bytes: Uint8Array;
  if (mode === 'user') {
    const zip = await JSZip.loadAsync(raw);
    const name = 'xl/worksheets/sheet1.xml';
    const xml = await zip.file(name)!.async('string');
    const out = xml.replace('width="9" outlineLevel="1"/>', 'width="9" outlineLevel="1" customWidth="1"/>');
    expect(out).not.toBe(xml);
    zip.file(name, out);
    bytes = await zip.generateAsync({ type: 'uint8array' });
  } else {
    bytes = await recol(raw, mode);
  }
  const book = await readBook(bytes, 'g-table.xlsx');
  return { bytes, book, s: createEditor(book, LAYOUT, {}, REGION) };
}

describe('a text constant with a line feed turns wrap text on (Excel does, the patcher mirrors it; golden q-wrap)', () => {
  const PLAIN = { sheet: '学生', insertAt: 0, imported: [], appendAt: null, appended: [] };
  const PLAIN_REGION = { headerRow: 1, firstDataRow: 2, lastDataRow: 6 };
  const plain = async () => {
    const bytes = golden('pm-base-plain.xlsx');
    const book = await readBook(bytes, 'pm-base-plain.xlsx');
    return { bytes, s0: createEditor(book, PLAIN, {}, PLAIN_REGION) };
  };
  const wrapOf = (s: ReturnType<typeof createEditor>, col: number, row: number) => s.doc.styles[s.doc.columns[col].cells[row - 1]?.s ?? 0]?.wrap === true;

  it('LF and CR LF switch it on, a lone CR and formulas / numbers do not; the style keeps everything else (doc == replay)', async () => {
    const { bytes, s0 } = await plain();
    let s = ok(setCellInput(s0, 'b2', 2, 'a\nb'));
    s = ok(setCellInput(s, 'b3', 3, 'a\r\nb'));
    s = ok(setCellInput(s, 'b3', 4, 'a\rb'));
    s = ok(setCellInput(s, 'b3', 5, '=B2&"x"'));
    s = ok(setCellInput(s, 'b4', 6, '12'));
    expect([wrapOf(s, 1, 2), wrapOf(s, 2, 3), wrapOf(s, 2, 4), wrapOf(s, 2, 5), wrapOf(s, 3, 6)]).toEqual([true, true, false, false, false]);
    // the derived style is the cell's own style plus wrap (fill / border / font stay)
    const before = s0.doc.styles[s0.doc.columns[1].cells[1]!.s ?? 0];
    expect(s.doc.styles[s.doc.columns[1].cells[1]!.s ?? 0]).toEqual({ ...before, wrap: true });
    // typing the same text again into a cell that already wraps changes nothing (no new style)
    const again = ok(setCellInput(s, 'b2', 2, 'c\nd'));
    expect(again.doc.styles).toBe(s.doc.styles);
    const after = await readBook(await replayLog(bytes, s.log), 'after.xlsx');
    expect(compare(s.doc, after)).toEqual([]);
  });

  it('a table header with a line feed and a totals label (doc == replay)', async () => {
    const bytes = golden('torture.xlsx');
    const book = await readBook(bytes, 'torture.xlsx');
    const s0 = createEditor(book, { sheet: 'Tbl', insertAt: 0, imported: [], appendAt: null, appended: [] }, {}, { headerRow: 1, firstDataRow: 2, lastDataRow: 11 });
    let s = ok(setCellInput(s0, s0.doc.columns[0].id, 1, 'x\ny')); // header cell
    s = ok(setCellInput(s, s.doc.columns[0].id, 12, 'sum\nof')); // totals row label
    expect(wrapOf(s, 0, 1)).toBe(true);
    expect(wrapOf(s, 0, 12)).toBe(true);
    const after = await readBook(await replayLog(bytes, s.log), 'after.xlsx');
    expect(compare(s.doc, after)).toEqual([]);
  });

  it('a formula result with a line feed that is converted to a value wraps like any constant written (doc == replay)', async () => {
    const z = await JSZip.loadAsync(golden('pm-base-plain.xlsx'));
    const name = 'xl/worksheets/sheet1.xml';
    const xml = await z.file(name)!.async('string');
    const cell = '<c r="E2" t="str"><f>B2&amp;CHAR(10)&amp;C2</f><v>x&#10;y</v></c>';
    const out = xml.replace('<c r="D2" s="5"><v>60</v></c></row>', `<c r="D2" s="5"><v>60</v></c>${cell}</row>`);
    expect(out).not.toBe(xml);
    z.file(name, out);
    const bytes = await z.generateAsync({ type: 'uint8array' });
    const book = await readBook(bytes, 'lf.xlsx');
    const s0 = createEditor(book, PLAIN, {}, PLAIN_REGION);
    expect(s0.doc.columns[4].cells[1]).toMatchObject({ t: 's', v: 'x\ny', f: 'B2&CHAR(10)&C2' });
    const s = ok(deleteColumns(s0, [s0.doc.columns[1].id], { convertDependents: true })); // 氏名 (B): E2 depends on it
    const c = s.doc.columns[3].cells[1]!; // E -> D
    expect(c).toMatchObject({ t: 's', v: 'x\ny' });
    expect(c.f).toBeUndefined();
    expect(s.doc.styles[c.s ?? 0].wrap).toBe(true);
    const after = await readBook(await replayLog(bytes, s.log), 'after.xlsx');
    expect(compare(s.doc, after)).toEqual([]);
  });
});

describe('CF block identity follows the ops like the patcher regroups blocks (fuzz seed 40057)', () => {
  // g-table 名簿 has three one-rule CF blocks. After the delete the cellIs rule (90) and a fragment of the third block
  // (G4>F4) cover the same cells F4:F13, but they are different blocks: the fragment's reference to the moved column must
  // not stop the 90 rule from spreading onto the block that lands next to it. The doc used to group rules by area only.
  it('rules of different blocks that coincide in area do not veto each other', async () => {
    const bytes = golden('g-table.xlsx');
    const book = await readBook(bytes, 'g-table.xlsx');
    let s = createEditor(book, LAYOUT, {}, REGION);
    const ids = (idx: number[]) => idx.map((i) => s.doc.columns[i].id);
    s = ok(moveColumns(s, ids([7, 4, 1]), 8, { unmergeTitles: true }));
    s = ok(deleteColumns(s, ids([2]), { convertDependents: false }));
    s = ok(moveColumns(s, ids([0, 6, 3]), 6, { unmergeTitles: true }));
    const after = await readBook(await replayLog(bytes, s.log), 'after.xlsx');
    expect(compare(s.doc, after)).toEqual([]);
    const gt90 = (rules: typeof s.doc.features.rules) => rules.filter((r) => r.kind === 'cf' && r.formulas.join() === '90').map((r) => r.sqref);
    expect(gt90(s.doc.features.rules)).toEqual(['D4:F13']);
    expect(gt90(after.sheets[0].features.rules)).toEqual(['D4:F13']);
  });
});

describe('customWidth flag as read from <col>', () => {
  it('customWidth without bestFit = user-set; no flag / bestFit customWidth = auto', async () => {
    expect((await editorOf('user')).s.doc.columns[J].props).toEqual({ width: 9, customWidth: true, outlineLevel: 1 });
    for (const mode of ['auto', 'bestFit'] as const) {
      const { s } = await editorOf(mode);
      expect(s.doc.columns[J].props).toEqual({ width: 9, outlineLevel: 1 });
      expect(s.doc.columns[1].props).toEqual({ width: 14.625 });
    }
  });
});

describe('auto and auto-fitted columns are widened by a typed date, user widths are not (doc == replay)', () => {
  for (const mode of ['auto', 'bestFit'] as const) {
    it(`${mode}: the doc widens like the patcher and keeps the column auto`, async () => {
      const { bytes, s: s0 } = await editorOf(mode);
      let s = ok(setCellInput(s0, s0.doc.columns[J].id, 15, '2026年10月20日'));
      const w1 = s.doc.columns[J].props.width!;
      expect(w1).toBeGreaterThan(9);
      expect(s.doc.columns[J].props.customWidth).toBeUndefined();
      // a wider value widens it again (Excel: an auto-fitted column keeps growing)
      s = ok(setCellInput(s, s.doc.columns[J].id, 16, '令和8年10月20日'));
      expect(s.doc.columns[J].props.width!).toBeGreaterThan(w1);
      // the neighbour that shared the <col> span with it is untouched
      expect(s.doc.columns[J + 1].props.width).toBe(9);
      // replay: same widths, same flag (auto: the patcher's bestFit customWidth reads back as no flag)
      const after = await readBook(await replayLog(bytes, s.log), 'after.xlsx');
      expect(compare(s.doc, after)).toEqual([]);
      expect(after.sheets[0].colProps[J]?.customWidth).toBeUndefined();
      expect(after.sheets[0].colProps[J]?.width).toBeCloseTo(s.doc.columns[J].props.width!, 6);
      expect(after.sheets[0].colProps[J + 1]?.width).toBe(9);
    });
  }

  it('user: the width is kept by the doc and by the replay', async () => {
    const { bytes, s: s0 } = await editorOf('user');
    const s = ok(setCellInput(s0, s0.doc.columns[J].id, 15, '令和8年10月20日'));
    expect(s.doc.columns[J].props).toEqual({ width: 9, customWidth: true, outlineLevel: 1 });
    const after = await readBook(await replayLog(bytes, s.log), 'after.xlsx');
    expect(compare(s.doc, after)).toEqual([]);
    expect(after.sheets[0].colProps[J]).toMatchObject({ width: 9, customWidth: true });
  });

  for (const mode of ['auto', 'bestFit'] as const) {
    it(`${mode}: a width the user sets afterwards ends auto-fit (no widening any more, flag = user-set)`, async () => {
      const { bytes, s: s0 } = await editorOf(mode);
      let s = ok(setCellInput(s0, s0.doc.columns[J].id, 15, '2026年10月20日')); // widens J
      s = ok(setColumnWidth(s, s.doc.columns[J].id, 9));
      s = ok(setCellInput(s, s.doc.columns[J].id, 16, '令和8年10月20日'));
      expect(s.doc.columns[J].props).toMatchObject({ width: 9, customWidth: true });
      const after = await readBook(await replayLog(bytes, s.log), 'after.xlsx');
      expect(compare(s.doc, after)).toEqual([]);
      expect(after.sheets[0].colProps[J]).toMatchObject({ width: 9, customWidth: true });
    });
  }

  for (const mode of ['auto', 'bestFit'] as const) {
    it(`${mode}: the text is measured in the font of the cell it is typed into, also when the format creates a new xf (fuzz seed 20048)`, async () => {
      // row 1 of g-table holds the big title font: the date typed into B1 (after D moved to the front, B = the old A) needs
      // a wider column than the default font would (doc used to measure with the default font: 11.375 instead of 14)
      const { bytes, s: s0 } = await editorOf(mode);
      let s = ok(moveColumns(s0, [s0.doc.columns[3].id], 0));
      s = ok(setCellInput(s, s.doc.columns[1].id, 1, '2026/10/20'));
      expect(s.doc.columns[1].props.width).toBe(14);
      const after = await readBook(await replayLog(bytes, s.log), 'after.xlsx');
      expect(compare(s.doc, after)).toEqual([]);
    });
  }

  it('an inserted column takes the width AND the auto flag of its left neighbour (doc == replay)', async () => {
    for (const mode of ['auto', 'bestFit', 'user'] as const) {
      const { bytes, s: s0 } = await editorOf(mode);
      let s = ok(setCellInput(s0, s0.doc.columns[J].id, 15, '2026年10月20日')); // J widened (auto modes) or kept (user)
      s = ok(insertColumns(s, J + 1, 1)); // right of J
      // a date typed into the new column: widened if and only if J is auto
      s = ok(setCellInput(s, s.doc.columns[J + 1].id, 15, '令和8年10月20日'));
      const after = await readBook(await replayLog(bytes, s.log), 'after.xlsx');
      expect(compare(s.doc, after), mode).toEqual([]);
    }
  });
});
