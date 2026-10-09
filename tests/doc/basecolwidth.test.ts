// Default column width from <sheetFormatPr baseColWidth> (Mac Excel writes baseColWidth="10" and no defaultColWidth).
//
// Reproduction (doc engineer): g-table 名簿 with baseColWidth="10", a date typed into G15 (column G has no <col>). The doc
// assumed the standard 8 characters (72 px for 游ゴシック 11) and widened G to 9.25, while the JS patcher derives the default
// column from baseColWidth (88 px) and keeps it: doc != readBook(replayLog). The reader now reports the sheet's default
// column width (SheetGrid.defaultColWidth, see sheetDefaultColWidth in xlsx/read.ts) so the doc follows the patcher — and
// both follow Excel (measured in a hidden Excel instance: baseColWidth 10 + 游ゴシック 11 = 88 px, tests/xlsx/read-defaultcolwidth).
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { createEditor, setCellInput } from '../../src/core/doc';
import { readBook, replayLog } from '../../src/core/xlsx';
import { compare, golden } from './compare';
import { ok } from './helpers';

const REGION = { headerRow: 3, firstDataRow: 4, lastDataRow: 13 };
const LAYOUT = { sheet: '名簿', insertAt: 0, imported: [], appendAt: null, appended: [] };
const G = 6; // 0-based doc column of G: tblMain A3:G13, no <col> element for it
const FMT = '<sheetFormatPr defaultRowHeight="18.75" outlineLevelCol="1"/>'; // g-table's own

/** g-table.xlsx with the sheet's <sheetFormatPr> replaced (null = left as the file has it: no baseColWidth). */
async function gTable(format: string | null): Promise<Uint8Array> {
  const raw = golden('g-table.xlsx');
  if (format === null) return raw;
  const zip = await JSZip.loadAsync(raw);
  const name = 'xl/worksheets/sheet1.xml';
  const xml = await zip.file(name)!.async('string');
  expect(xml).toContain(FMT);
  zip.file(name, xml.replace(FMT, format));
  return zip.generateAsync({ type: 'uint8array' });
}

async function editorOf(format: string | null) {
  const bytes = await gTable(format);
  const book = await readBook(bytes, 'g-table.xlsx');
  return { bytes, book, s0: createEditor(book, LAYOUT, {}, REGION) };
}

/** the doc after typing `text` into G15 and what the patcher's replay of the same log reads back as */
async function typeIntoG15(format: string | null, text: string) {
  const { bytes, s0 } = await editorOf(format);
  const s = ok(setCellInput(s0, s0.doc.columns[G].id, 15, text));
  const after = await readBook(await replayLog(bytes, s.log), 'after.xlsx');
  return { s, after, diffs: compare(s.doc, after) };
}

describe('baseColWidth: the doc widens a typed date like the patcher (and Excel) does', () => {
  it('Mac Excel: baseColWidth="10" — 2026/1/1 fits the 88 px default column: no widening in the doc, none in the replay', async () => {
    const { s, after, diffs } = await typeIntoG15('<sheetFormatPr baseColWidth="10" defaultRowHeight="18.75" outlineLevelCol="1"/>', '2026/1/1');
    expect(diffs).toEqual([]);
    expect(s.doc.columns[G].props.width).toBeUndefined(); // used to be 9.25 in the doc while the file kept the column without a width
    expect(after.sheets[0].colProps[G]?.width).toBeUndefined();
    expect(s.log).toHaveLength(1); // the typed value only: no width op
  });

  it('the same sheet with the reader\'s default column width removed reproduces the old disagreement (negative control)', async () => {
    const bytes = await gTable('<sheetFormatPr baseColWidth="10" defaultRowHeight="18.75" outlineLevelCol="1"/>');
    const book = await readBook(bytes, 'g-table.xlsx');
    expect(book.sheets[0].defaultColWidth).toBeGreaterThan(10); // what the fix reports
    book.sheets[0].defaultColWidth = undefined; // what the reader reported before: the doc assumes 8 characters
    const s0 = createEditor(book, LAYOUT, {}, REGION);
    const s = ok(setCellInput(s0, s0.doc.columns[G].id, 15, '2026/1/1'));
    expect(s.doc.columns[G].props.width).toBe(9.25);
    const after = await readBook(await replayLog(bytes, s.log), 'after.xlsx');
    expect(compare(s.doc, after)).toEqual(['G width doc 9.25 file undefined']);
  });

  it('baseColWidth absent or 8 (the schema default): 2026/1/1 does NOT fit 72 px — both widen G to 9.25', async () => {
    for (const fmt of [null, '<sheetFormatPr baseColWidth="8" defaultRowHeight="18.75" outlineLevelCol="1"/>']) {
      const { s, after, diffs } = await typeIntoG15(fmt, '2026/1/1');
      expect(diffs, String(fmt)).toEqual([]);
      expect(s.doc.columns[G].props.width, String(fmt)).toBe(9.25);
      expect(after.sheets[0].colProps[G]?.width, String(fmt)).toBeCloseTo(9.25, 6);
      expect(after.sheets[0].colProps[G]?.customWidth, String(fmt)).toBeUndefined(); // auto-fitted: bestFit customWidth reads back as auto
    }
  });

  it('the default column grows with baseColWidth: 2026/12/31 (86 px) widens the 88 px column of base 10, fits base 11 (96 px) and up', async () => {
    for (const base of [null, 8, 9, 10, 11, 12, 15]) {
      const fmt = base === null ? null : `<sheetFormatPr baseColWidth="${base}" defaultRowHeight="18.75" outlineLevelCol="1"/>`;
      const { s, after, diffs } = await typeIntoG15(fmt, '2026/12/31');
      expect(diffs, `base ${base}`).toEqual([]);
      const w = s.doc.columns[G].props.width;
      if (base === null || base <= 10) {
        expect(w, `base ${base}`).toBe(11.375); // (86 + 5 px) / 8
        expect(after.sheets[0].colProps[G]?.width, `base ${base}`).toBeCloseTo(11.375, 6);
      } else {
        expect(w, `base ${base}`).toBeUndefined();
        expect(after.sheets[0].colProps[G]?.width, `base ${base}`).toBeUndefined();
      }
    }
  });

  it('a date wider than any default column widens it in both, to the same width', async () => {
    for (const base of [10, 12, 15, 20]) {
      const fmt = `<sheetFormatPr baseColWidth="${base}" defaultRowHeight="18.75" outlineLevelCol="1"/>`;
      const { s, after, diffs } = await typeIntoG15(fmt, '令和8年10月20日');
      expect(diffs, `base ${base}`).toEqual([]);
      const w = s.doc.columns[G].props.width;
      if (base === 20) expect(w).toBeUndefined(); // 168 px holds it
      else {
        expect(w, `base ${base}`).toBe(16.5);
        expect(after.sheets[0].colProps[G]?.width, `base ${base}`).toBeCloseTo(16.5, 6);
      }
    }
  });

  it('an explicit defaultColWidth beats baseColWidth in the doc and the patcher', async () => {
    // 9.5 chars = 76 px (stored): 2026/1/1 (69 px + 5) fits, so nothing is widened although baseColWidth 10 would allow even more
    const fit = await typeIntoG15('<sheetFormatPr baseColWidth="10" defaultColWidth="9.5" defaultRowHeight="18.75" outlineLevelCol="1"/>', '2026/1/1');
    expect(fit.diffs).toEqual([]);
    expect(fit.s.doc.columns[G].props.width).toBeUndefined();
    // 7 chars = 56 px: even a short date does not fit and both widen although baseColWidth 10 would hold it
    const narrow = await typeIntoG15('<sheetFormatPr baseColWidth="10" defaultColWidth="7" defaultRowHeight="18.75" outlineLevelCol="1"/>', '2026/1/1');
    expect(narrow.diffs).toEqual([]);
    expect(narrow.s.doc.columns[G].props.width).toBe(9.25);
  });

  it('a non-positive defaultColWidth is ignored by both (baseColWidth decides)', async () => {
    for (const dcw of ['0', '-3']) {
      const { s, diffs } = await typeIntoG15(`<sheetFormatPr baseColWidth="10" defaultColWidth="${dcw}" defaultRowHeight="18.75" outlineLevelCol="1"/>`, '2026/1/1');
      expect(diffs, dcw).toEqual([]);
      expect(s.doc.columns[G].props.width, dcw).toBeUndefined();
    }
  });

  it('a user width, an auto width with a <col>, and the other columns are unaffected by the default column width', async () => {
    const { bytes, s0 } = await editorOf('<sheetFormatPr baseColWidth="10" defaultRowHeight="18.75" outlineLevelCol="1"/>');
    // J / K share <col width="9" outlineLevel="1"/> (an auto width of 9 chars): 2026/1/1 needs more than 9, so J widens as before
    let s = ok(setCellInput(s0, s0.doc.columns[9].id, 15, '2026/1/1'));
    expect(s.doc.columns[9].props.width).toBeGreaterThan(9);
    // C is a user width (customWidth): never touched
    s = ok(setCellInput(s, s.doc.columns[2].id, 15, '令和8年10月20日'));
    expect(s.doc.columns[2].props.width).toBe(9.625);
    // G: still fits the default column
    s = ok(setCellInput(s, s.doc.columns[G].id, 16, '2026/1/1'));
    expect(s.doc.columns[G].props.width).toBeUndefined();
    const after = await readBook(await replayLog(bytes, s.log), 'after.xlsx');
    expect(compare(s.doc, after)).toEqual([]);
  });
});
