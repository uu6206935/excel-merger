// Internal consistency of the MergedDoc after random op sequences (no package involved, so it also catches mistakes the doc
// and the patcher would make the same way): every column has rowCount cells, merges / tables / arrays / filter / freeze /
// hyperlinks stay inside the sheet, a table's header cells carry the column names, array masters and their ranges agree, every
// formula keeps tokenizing losslessly, style indices exist, widths are sane, versions move forward. Only violations that an op
// INTRODUCED count (a fixture may start with an odd state). DOC_FUZZ_COUNT=<n> sequences per workbook (default 60).
import { describe, expect, it } from 'vitest';
import { createEditor } from '../../src/core/doc';
import { tableBox } from '../../src/core/doc/tables';
import { parseBox } from '../../src/core/doc/util';
import { MAX_COL, tokenize } from '../../src/core/formula/excelRef';
import { readBook } from '../../src/core/xlsx';
import type { EditorSnapshot, MergedDoc } from '../../src/core/types';
import { insideTableScenario, richScenario, scenario } from '../golden/premerge-scenario';
import { golden } from './compare';
import { recol, rng, step } from './fuzzcore';
import { buildSamplePremerged, premergedOf, samplesAvailable, type SamplePremerged } from './sampleMerge';

const lossless = (f: string) => tokenize(f).map((t) => t.text).join('') === f;

/** every rule of the doc that is broken, as a stable text */
function problems(doc: MergedDoc): Set<string> {
  const out = new Set<string>();
  const n = doc.columns.length;
  const bad = (msg: string) => out.add(msg);
  const ids = new Set<string>();
  doc.columns.forEach((c, i) => {
    if (ids.has(c.id)) bad(`duplicate column id ${c.id}`);
    ids.add(c.id);
    if (c.cells.length !== doc.rowCount) bad(`column ${i + 1} has ${c.cells.length} cells, rowCount ${doc.rowCount}`);
    const w = c.props.width;
    if (w !== undefined && !(Number.isFinite(w) && w >= 0 && w <= 255)) bad(`column ${i + 1} width ${w}`);
    if (c.props.s !== undefined && !(c.props.s >= 0 && c.props.s < doc.styles.length)) bad(`column ${i + 1} style ${c.props.s} missing`);
    c.cells.forEach((cell, r) => {
      if (!cell) return;
      if (cell.s !== undefined && !(cell.s >= 0 && cell.s < doc.styles.length)) bad(`cell col ${i + 1} row ${r + 1} style ${cell.s} missing (${doc.styles.length} styles)`);
      if (cell.f !== undefined && !lossless(cell.f)) bad(`formula col ${i + 1} row ${r + 1} does not tokenize losslessly: ${cell.f}`);
    });
  });
  // merges: inside the sheet, more than one cell, no overlap
  const merges = doc.features.merges;
  merges.forEach((m, k) => {
    if (!(m.r1 >= 1 && m.c1 >= 1 && m.r1 <= m.r2 && m.c1 <= m.c2 && m.c2 <= n && m.r2 <= doc.rowCount)) bad(`merge ${JSON.stringify(m)} outside the sheet (${n} columns, ${doc.rowCount} rows)`);
    if (m.r1 === m.r2 && m.c1 === m.c2) bad(`merge of one cell ${JSON.stringify(m)}`);
    for (const o of merges.slice(k + 1)) if (m.r1 <= o.r2 && o.r1 <= m.r2 && m.c1 <= o.c2 && o.c1 <= m.c2) bad(`merges overlap ${JSON.stringify(m)} ${JSON.stringify(o)}`);
  });
  // tables: geometry, column list, header cells = column names
  for (const t of doc.features.tables) {
    const b = tableBox(t);
    if (!b) {
      bad(`table ${t.name}: ref ${t.ref} not understood`);
      continue;
    }
    if (!(b.c1 >= 1 && b.c2 <= n && b.r1 >= 1 && b.r2 <= doc.rowCount)) bad(`table ${t.name} ${t.ref} outside the sheet`);
    if (t.columns.length !== b.c2 - b.c1 + 1) bad(`table ${t.name}: ${t.columns.length} column names for ${b.c2 - b.c1 + 1} columns`);
    if (new Set(t.columns.map((c) => c.toLowerCase())).size !== t.columns.length) bad(`table ${t.name}: duplicate column names ${t.columns.join('|')}`);
    if (t.headerRow) {
      t.columns.forEach((name, i) => {
        const cell = doc.columns[b.c1 - 1 + i]?.cells[b.r1 - 1];
        const text = cell && (cell.t === 's' || cell.t === 'n' || cell.t === 'd') ? String(cell.v) : '';
        if (text !== name) bad(`table ${t.name}: header cell of column ${i + 1} is "${text}", the column is named "${name}"`);
      });
    }
  }
  // arrays: master carries `arr`, the rest of the range holds no formula
  for (const a of doc.features.arrays) {
    if (!(a.c1 >= 1 && a.c2 <= n && a.r1 >= 1 && a.r2 <= doc.rowCount)) bad(`array ${JSON.stringify(a)} outside the sheet`);
    const master = doc.columns[a.c1 - 1]?.cells[a.r1 - 1];
    if (!master || master.f === undefined || master.arr === undefined) bad(`array ${JSON.stringify(a)} has no master formula`);
    for (let c = a.c1; c <= a.c2; c++) {
      for (let r = a.r1; r <= a.r2; r++) {
        if (c === a.c1 && r === a.r1) continue;
        if (doc.columns[c - 1]?.cells[r - 1]?.f !== undefined) bad(`array ${JSON.stringify(a)}: a second formula at col ${c} row ${r}`);
      }
    }
  }
  doc.columns.forEach((col, i) =>
    col.cells.forEach((cell, r) => {
      if (cell?.arr === undefined) return;
      const box = parseBox(cell.arr.replace(/\$/g, ''));
      if (!box || box.c1 !== i + 1 || box.r1 !== r + 1) bad(`cell col ${i + 1} row ${r + 1}: arr ${cell.arr} does not start at the cell`);
    }),
  );
  // (a split may lie beyond the used columns: deleting every frozen column keeps the split, measured with Excel)
  const fr = doc.features.freeze;
  if (fr && !(fr.cols >= 0 && fr.cols < MAX_COL)) bad(`freeze ${fr.cols} columns`);
  if (doc.features.autoFilter) {
    const b = parseBox(doc.features.autoFilter.replace(/\$/g, ''));
    if (!b || b.c1 < 1 || b.c2 > n) bad(`autoFilter ${doc.features.autoFilter} outside the ${n} columns`);
  }
  for (const h of doc.features.hyperlinks) {
    const b = parseBox(h.ref.replace(/\$/g, ''));
    if (!b || b.c1 < 1 || b.c2 > n) bad(`hyperlink ${h.ref} outside the ${n} columns`);
  }
  for (const r of doc.features.rules) {
    for (const part of r.sqref.split(/\s+/).filter(Boolean)) {
      const b = parseBox(part.replace(/\$/g, ''));
      if (!b || b.c1 < 1 || b.c2 > MAX_COL) bad(`rule area ${part} invalid`);
    }
    for (const f of r.formulas) if (!lossless(f)) bad(`rule formula does not tokenize losslessly: ${f}`);
  }
  for (const nm of doc.names) if (!lossless(nm.refersTo)) bad(`name ${nm.name} does not tokenize losslessly: ${nm.refersTo}`);
  const ext = new Set<string>();
  for (const e of doc.externalFormulas) {
    const k = `${e.sheet}!${e.col},${e.row}`;
    if (ext.has(k)) bad(`external formula twice at ${k}`);
    ext.add(k);
    if (!lossless(e.f)) bad(`external formula ${k} does not tokenize losslessly: ${e.f}`);
  }
  return out;
}

const COUNT = Number(process.env.DOC_FUZZ_COUNT ?? 60);
const STEPS = Number(process.env.DOC_FUZZ_STEPS ?? 25);

function walk(label: string, make: () => EditorSnapshot, seed0: number, totalsRow?: number): string[] {
  const failures: string[] = [];
  for (let k = 0; k < COUNT; k++) {
    const r = rng(seed0 + k);
    let s = make();
    const baseline = problems(s.doc);
    const trail: string[] = [];
    for (let i = 0; i < STEPS; i++) {
      const { res, what } = step(s, r, false, totalsRow);
      trail.push(res.ok ? what : `${what} [refused]`);
      if (!res.ok) continue;
      const next = res.snapshot;
      const fresh = [...problems(next.doc)].filter((p) => !baseline.has(p));
      if (next.doc !== s.doc && next.doc.version <= s.doc.version) fresh.push(`version ${s.doc.version} -> ${next.doc.version} although the doc changed`);
      if (next !== s && next.log.length < s.log.length) fresh.push('the log shrank');
      if (fresh.length) {
        failures.push(`${label} seed ${seed0 + k}: ${fresh.slice(0, 3).join(' | ')}\n    ops: ${trail.join('; ')}`);
        break;
      }
      s = next;
    }
  }
  return failures;
}

const GOLD: Array<[string, string, { headerRow: number; firstDataRow: number; lastDataRow: number }, { totalsRow?: number; cols?: 'auto' | 'bestFit' }?]> = [
  ['torture.xlsx', 'Data', { headerRow: 3, firstDataRow: 4, lastDataRow: 23 }],
  ['torture.xlsx', 'Tbl', { headerRow: 1, firstDataRow: 2, lastDataRow: 11 }, { totalsRow: 12 }],
  ['g-table.xlsx', '名簿', { headerRow: 3, firstDataRow: 4, lastDataRow: 13 }, { cols: 'auto' }],
  ['pm-base-plain.xlsx', '学生', { headerRow: 1, firstDataRow: 2, lastDataRow: 6 }],
  ['pm-base-table.xlsx', '学生', { headerRow: 1, firstDataRow: 2, lastDataRow: 6 }],
  ['r2-table.xlsx', 'T', { headerRow: 1, firstDataRow: 2, lastDataRow: 4 }],
  ['r2-arrays.xlsx', 'A', { headerRow: 1, firstDataRow: 2, lastDataRow: 4 }],
  ['r3-misc.xlsx', 'Tb', { headerRow: 1, firstDataRow: 2, lastDataRow: 4 }],
  ['r3-misc.xlsx', 'Band', { headerRow: 1, firstDataRow: 3, lastDataRow: 11 }],
];

describe('the consistency checks themselves are not vacuous', () => {
  it('they flag a wrong header cell, a short column, a duplicate id, a missing style and a broken array', async () => {
    const book = await readBook(golden('torture.xlsx'), 'torture.xlsx');
    const s = createEditor(book, { sheet: 'Tbl', insertAt: 0, imported: [], appendAt: null, appended: [] }, {}, { headerRow: 1, firstDataRow: 2, lastDataRow: 11 });
    expect([...problems(s.doc)]).toEqual([]);
    const cols = s.doc.columns.slice();
    cols[0] = { ...cols[0], cells: cols[0].cells.map((c, r) => (r === 0 ? { t: 's' as const, v: 'zzz' } : c)) };
    expect([...problems({ ...s.doc, columns: cols })].join('|')).toMatch(/header cell of column 1 is "zzz"/);
    const short = s.doc.columns.slice();
    short[1] = { ...short[1], cells: short[1].cells.slice(1) };
    expect([...problems({ ...s.doc, columns: short })].join('|')).toMatch(/column 2 has \d+ cells/);
    const dup = s.doc.columns.slice();
    dup[1] = { ...dup[1], id: dup[0].id };
    expect([...problems({ ...s.doc, columns: dup })].join('|')).toMatch(/duplicate column id/);
    const styled = s.doc.columns.slice();
    styled[2] = { ...styled[2], cells: styled[2].cells.map((c, r) => (r === 3 ? { t: 'n' as const, v: 1, s: 9999 } : c)) };
    expect([...problems({ ...s.doc, columns: styled })].join('|')).toMatch(/style 9999 missing/);
    const arr = { ...s.doc, features: { ...s.doc.features, arrays: [{ r1: 2, c1: 1, r2: 3, c2: 1 }] } };
    expect([...problems(arr)].join('|')).toMatch(/has no master formula/);
  });
});

describe('the doc stays internally consistent under random ops', () => {
  for (const [file, sheet, region, opt] of GOLD) {
    it(`${file} ${sheet}`, async () => {
      const bytes = opt?.cols ? await recol(golden(file), opt.cols) : golden(file);
      const book = await readBook(bytes, file);
      const make = () => createEditor(book, { sheet, insertAt: 0, imported: [], appendAt: null, appended: [] }, {}, region);
      expect(walk(`${file}/${sheet}`, make, 31000, opt?.totalsRow)).toEqual([]);
    });
  }
  const merged: Array<[string, () => Promise<SamplePremerged>, boolean]> = [
    ['sample merge (01 + 02 03 05)', () => buildSamplePremerged({}), true],
    ['sample merge, full join', () => buildSamplePremerged({ joinType: 'full' }), true],
    ['sample merge, xls base', () => buildSamplePremerged({ base: '04_適性検査.xls', others: ['01_学生名簿.xlsx'] }), true],
    ['scenario table', () => premergedOf(scenario('table')), false],
    ['scenario plain', () => premergedOf(scenario('plain')), false],
    ['scenario imports inside a table', () => premergedOf(insideTableScenario()), false],
    ['scenario rich text', () => premergedOf(richScenario()), false],
  ];
  for (const [name, build, needsSamples] of merged) {
    it.skipIf(needsSamples && !samplesAvailable())(name, async () => {
      const p = await build();
      expect(walk(name, () => createEditor(p.book, p.layout, p.sourceInfo, p.region), 33000, p.region.lastDataRow + 2)).toEqual([]);
    });
  }
});
