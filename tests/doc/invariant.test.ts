// Invariant (docs/ARCHITECTURE.md §6): for every op sequence, readBook(replayLog(premerged, log)) equals the doc
// after the same ops — cell values / formulas / styles of the merged sheet, every formula on the other sheets,
// names and sheet features. Runs real workbooks through the real reader and the JS package patcher.
import { describe, expect, it } from 'vitest';
import { colToNum, numToCol } from '../../src/core/formula/excelRef';
import { createEditor, deleteColumns, insertColumns, moveColumns, setCellInput, setColumnWidth } from '../../src/core/doc';
import { tableBox } from '../../src/core/doc/tables';
import { readBook, replayLog } from '../../src/core/xlsx';
import type { EditorSnapshot } from '../../src/core/types';
import type { OpResult } from '../../src/core/contracts';
import { compare, golden, idAt } from './compare';

interface Case {
  file: string;
  sheet: string;
  region: { headerRow: number; firstDataRow: number; lastDataRow: number };
  /** first column the scripts touch (torture.xlsx has a title merged over A1:F1) */
  offset: number;
  /** scripts known to differ because of a gap in the package patcher (it.fails: turns red once fixed) */
  knownGaps?: string[];
  /** the 'unmerge' script really unmerges a title (a title merge blocks its moves) */
  titleBlocks?: boolean;
  /** run only these scripts (default: all) */
  only?: string[];
}
const CASES: Case[] = [
  { file: 'torture.xlsx', sheet: 'Data', region: { headerRow: 3, firstDataRow: 4, lastDataRow: 23 }, offset: 6 },
  { file: 'torture.xlsx', sheet: 'Data', region: { headerRow: 3, firstDataRow: 4, lastDataRow: 23 }, offset: 0, titleBlocks: true },
  // header of a table column: Excel renames the table column + structured refs (replayLog does it since round 2)
  // 'dynamic' inserts a table column right of 「点数1」 while 「点数2」 exists: Excel (measured, doc/tables.ts) names it
  // 「点数12」; replayLog names new table columns the same way since round 3 (it was the it.fails gap of round 2)
  { file: 'g-table.xlsx', sheet: '名簿', region: { headerRow: 3, firstDataRow: 4, lastDataRow: 13 }, offset: 0 },
  { file: 'g-table.xlsx', sheet: '名簿', region: { headerRow: 3, firstDataRow: 4, lastDataRow: 13 }, offset: 7, titleBlocks: true },
  { file: 'pm-base-plain.xlsx', sheet: '学生', region: { headerRow: 1, firstDataRow: 2, lastDataRow: 6 }, offset: 0 },
  // a legacy array (B2:B4), two dynamic arrays (D2:E4, G2:G4) and readers of a spill on another sheet (B!A2 = A!G2#)
  { file: 'r2-arrays.xlsx', sheet: 'A', region: { headerRow: 1, firstDataRow: 2, lastDataRow: 4 }, offset: 0 },
  // round 3: five title merges above the header row (A1:F1 band with a box on A1 only, a 3 x 3 area A3:C5 whose cells carry
  // their own date / number formats and borders, a theme-coloured band, plain / own-border bands): unmerging spreads the
  // top-left cell's fill and the border sides of each edge (multi-row areas included), like Excel's UnMerge + the copy
  { file: 'r3-misc.xlsx', sheet: 'Band', region: { headerRow: 12, firstDataRow: 13, lastDataRow: 12 }, offset: 0, titleBlocks: true, only: ['unmerge', 'moves'] },
  // round 3: a table WITH a totals row (A1:D12, row 12): constants typed into totals cells become the column labels (text)
  { file: 'torture.xlsx', sheet: 'Tbl', region: { headerRow: 1, firstDataRow: 2, lastDataRow: 11 }, offset: 0, only: ['totals', 'quoted', 'edits', 'numFmt', 'header'] },
];

type Step = (s: EditorSnapshot, o: number) => OpResult;

/** Op scripts (column k = offset + k); a step the doc refuses (like Excel would) is skipped. */
const SCRIPTS: Record<string, Step[]> = {
  moves: [
    (s, o) => moveColumns(s, [idAt(s, o + 1)], s.doc.columns.length),
    (s, o) => moveColumns(s, [idAt(s, o + 3), idAt(s, o + 4)], o),
    (s, o) => moveColumns(s, [idAt(s, o), idAt(s, o + 2), idAt(s, o + 5)], o + 4),
    (s, o) => moveColumns(s, [idAt(s, o + 2)], o),
  ],
  insertDelete: [
    (s, o) => insertColumns(s, o + 2, 2),
    (s, o) => deleteColumns(s, [idAt(s, o + 4)], { convertDependents: true }),
    (s, o) => deleteColumns(s, [idAt(s, o), idAt(s, o + 6)], { convertDependents: true }),
    (s, o) => insertColumns(s, o + 1, 1),
    (s, o) => deleteColumns(s, [idAt(s, o + 3)], { convertDependents: false }),
  ],
  edits: [
    (s, o) => setCellInput(s, idAt(s, o + 1), s.doc.firstDataRow, '１，２３４'),
    (s, o) => setCellInput(s, idAt(s, o + 2), s.doc.firstDataRow + 1, '=SUM(A1:B2)*2'),
    (s, o) => setCellInput(s, idAt(s, o), s.doc.firstDataRow + 2, ''),
    (s, o) => setColumnWidth(s, idAt(s, o + 3), 21.5),
    (s, o) => moveColumns(s, [idAt(s, o + 2)], o),
    (s, o) => setCellInput(s, idAt(s, o), s.doc.firstDataRow + 3, 'テキスト'),
  ],
  header: [
    (s, o) => setCellInput(s, idAt(s, o + 1), s.doc.headerRow, '新しい見出し'),
    (s, o) => moveColumns(s, [idAt(s, o + 1)], o),
  ],
  // round 2: typed dates / % / currency set the cell's number format (CellWrite.numFmt)
  numFmt: [
    (s, o) => setCellInput(s, idAt(s, o + 1), s.doc.firstDataRow, '2026/10/20'),
    (s, o) => setCellInput(s, idAt(s, o + 2), s.doc.firstDataRow, '50%'),
    (s, o) => setCellInput(s, idAt(s, o + 3), s.doc.firstDataRow + 1, '¥1,000'),
    (s, o) => setCellInput(s, idAt(s, o + 1), s.doc.firstDataRow + 2, '10/20'),
    (s, o) => setCellInput(s, idAt(s, o + 4), s.doc.firstDataRow + 3, 'R8.10.20'),
    (s, o) => setCellInput(s, idAt(s, o + 4), s.doc.firstDataRow + 4, '1,000'),
    (s, o) => moveColumns(s, [idAt(s, o + 1)], o + 4),
  ],
  // round 2: typed formulas stored as dynamic arrays (CellWrite.dyn), moved, then converted to values by a delete
  dynamic: [
    (s, o) => setCellInput(s, idAt(s, o + 2), s.doc.firstDataRow, `=${numToCol(o + 2)}${s.doc.firstDataRow}:${numToCol(o + 2)}${s.doc.firstDataRow + 2}`),
    (s, o) => setCellInput(s, idAt(s, o + 3), s.doc.firstDataRow + 1, '=SEQUENCE(2)'),
    (s, o) => moveColumns(s, [idAt(s, o + 3)], o),
    (s, o) => insertColumns(s, o + 1, 1),
    (s) => deleteColumns(s, [dynSourceId(s)], { convertDependents: true }),
  ],
  // round 2: deleting the column that feeds arrays converts them (toValues over the MASTER replaces the whole array,
  // readers of a converted spill on other sheets follow); the remaining spill block moves as a unit
  arrays: [
    (s, o) => deleteColumns(s, [idAt(s, o)], { convertDependents: true }),
    (s, o) => moveColumns(s, [idAt(s, o + 2), idAt(s, o + 3)], o + 8),
    (s, o) => insertColumns(s, o + 1, 1),
  ],
  // round 3: a block inside a table moved to the gap right after the table's last column (c2 + 1) is an in-table
  // reorder (table columns reordered, range unchanged). g-table: table A3:G13 (7 columns)
  tableReorder: [
    (s, o) => moveColumns(s, [idAt(s, o + 1)], o + 7),
    (s, o) => moveColumns(s, [idAt(s, o + 2), idAt(s, o + 3)], o + 7),
    (s, o) => moveColumns(s, [idAt(s, o)], o + 7),
    (s, o) => moveColumns(s, [idAt(s, o + 4), idAt(s, o + 6)], o + 7),
  ],
  // round 3: quote-prefixed text (leading apostrophe) is stored as the text without the apostrophe; ordinary text as is
  quoted: [
    (s, o) => setCellInput(s, idAt(s, o + 1), s.doc.firstDataRow, "'00123"),
    (s, o) => setCellInput(s, idAt(s, o + 2), s.doc.firstDataRow + 1, "'2026/10/20"),
    (s, o) => setCellInput(s, idAt(s, o), s.doc.firstDataRow + 2, "'abc"),
    (s, o) => setCellInput(s, idAt(s, o + 3), s.doc.firstDataRow + 3, '03-1234-5678'),
    (s, o) => setCellInput(s, idAt(s, o + 1), s.doc.firstDataRow + 4, "'TRUE"),
  ],
  // round 3: a table's totals row: constants typed there are text (the column labels), formulas stay formulas
  totals: [
    (s, o) => setCellInput(s, idAt(s, o + 1), totalsRowOf(s), '123'),
    (s, o) => setCellInput(s, idAt(s, o + 2), totalsRowOf(s), '2026/10/20'),
    (s, o) => setCellInput(s, idAt(s, o), totalsRowOf(s), 'TRUE'),
    (s, o) => setCellInput(s, idAt(s, o + 3), totalsRowOf(s), '=SUBTOTAL(104,[金額])'),
    (s, o) => setCellInput(s, idAt(s, o + 2), totalsRowOf(s), '=SUBTOTAL(101,[単価])'),
    (s, o) => setCellInput(s, idAt(s, o + 1), totalsRowOf(s), '50%'),
    (s, o) => setCellInput(s, idAt(s, o + 3), totalsRowOf(s), ''),
  ],
  // round 2: title merges above the header row are unmerged (centerContinuous) when they block a move
  unmerge: [
    (s, o) => moveColumns(s, [idAt(s, o + 1)], s.doc.columns.length, { unmergeTitles: true }),
    (s, o) => moveColumns(s, [idAt(s, o + 3), idAt(s, o + 4)], o, { unmergeTitles: true }),
    (s, o) => setCellInput(s, idAt(s, o), s.doc.firstDataRow, '2026/10/20'),
  ],
};

/** row of the totals row of the first table that has one (else the row below the data) */
function totalsRowOf(s: EditorSnapshot): number {
  for (const t of s.doc.features.tables) {
    const b = t.totalsRow ? tableBox(t) : null;
    if (b) return b.r2;
  }
  return s.doc.lastDataRow + 1;
}

/** id of the column the first dynamic-array formula of the doc reads (the 'dynamic' script deletes it). */
function dynSourceId(s: EditorSnapshot): string {
  for (const col of s.doc.columns) {
    for (const cell of col.cells) {
      const m = cell?.dyn && cell.f ? /^([A-Z]+)\d+:/.exec(cell.f) : null;
      if (m) return idAt(s, colToNum(m[1]) - 1);
    }
  }
  return idAt(s, 0);
}

function run(s0: EditorSnapshot, steps: Step[], o: number): { s: EditorSnapshot; skipped: string[] } {
  let s = s0;
  const skipped: string[] = [];
  steps.forEach((step, i) => {
    const r = step(s, o);
    if (r.ok) s = r.snapshot;
    else skipped.push(`step ${i}: ${r.reason}`);
  });
  return { s, skipped };
}

describe('doc == readBook(replayLog(premerged, log))', () => {
  for (const c of CASES) {
    for (const [name, steps] of Object.entries(SCRIPTS)) {
      if (c.only && !c.only.includes(name)) continue;
      const gap = c.knownGaps?.includes(name);
      (gap && !process.env.DOC_SHOW_GAPS ? it.fails : it)(`${c.file} ${c.sheet} @${c.offset}: ${name}${gap ? ' (KNOWN GAP in replayLog)' : ''}`, async () => {
        const bytes = golden(c.file);
        const book = await readBook(bytes, c.file);
        const s0 = createEditor(book, { sheet: c.sheet, insertAt: 0, imported: [], appendAt: null, appended: [] }, {}, c.region);
        // the untouched doc equals the file
        expect(compare(s0.doc, book)).toEqual([]);
        const { s, skipped } = run(s0, steps, c.offset);
        if (name === 'unmerge' && c.titleBlocks) expect(s.log.some((e) => e.kind === 'unmerge')).toBe(true);
        if (skipped.length) console.log(`${c.file} ${name} skipped:`, skipped.join(" | "));
        if (!s.log.length) {
          // every step refused (like Excel): nothing to replay
          expect(skipped.length).toBe(steps.length);
          return;
        }
        const out = await replayLog(bytes, s.log);
        const after = await readBook(out, c.file);
        expect(compare(s.doc, after)).toEqual([]);
      }, 60_000);
    }
  }
});
