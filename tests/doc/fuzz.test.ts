// Randomized op sequences (seeded, reproducible): every sequence of moves / inserts / deletes / typed cells / widths the
// doc accepts must read back from replayLog exactly as the doc shows it (same invariant as invariant.test.ts, but over
// combinations nobody scripted: spill cuts, title unmerges, typed dates, conversions of array masters ...).
// DOC_FUZZ_STEPS=<n> changes the length of a sequence. A failure names the fixture and seed; re-run with DOC_FUZZ_SEED=<seed> to replay one sequence.
import { describe, expect, it } from 'vitest';
import { commit, createEditor, initHistory, redo, undo } from '../../src/core/doc';
import { readBook, replayLog } from '../../src/core/xlsx';
import { compare, golden } from './compare';
import { cfBlocks, cfRules, recol, rng, step, withRulesOn } from './fuzzcore';

interface Fixture {
  file: string;
  sheet: string;
  region: { headerRow: number; firstDataRow: number; lastDataRow: number };
  /** no column inserts (none needed any more: the patcher names new table columns like Excel since round 3) */
  noInsert?: boolean;
  /** row of a table's totals row: typed texts also go there (constants become the column labels) */
  totalsRow?: number;
  /**
   * rewrite every <col> of the package: 'auto' = no customWidth (an untouched column: typed dates widen it), 'bestFit' =
   * bestFit="1" customWidth="1" (a column Excel auto-fitted: read back as auto, widened again by a wider value)
   */
  cols?: 'auto' | 'bestFit';
  /** CF / DV XML added to another sheet's part (rules of other sheets that read the merged sheet: doc.externalRules) */
  ext?: { path: string; xml: string };
  /** CF / DV XML that REPLACES the merged sheet's own classic rules (xl/worksheets/sheet1.xml): the CF stress variants */
  rules?: string;
}

// CF stress: several blocks with several rules each (every rule type family), multi-area blocks, relative / absolute
// formulas, a colour scale with a formula cfvo, data validations; moves / inserts / deletes fragment and regroup them
const CF_PLAIN =
  cfBlocks([
    ['A2:D6', [cfRules.expr('$D2>70'), cfRules.cellIs('between', '60', '80')]],
    ['B2:C6', [cfRules.expr('B2=$B$3'), cfRules.scale()]],
    ['D2:D6', [cfRules.bar(), cfRules.top(2)]],
    ['A2:B6 D2:D6', [cfRules.expr('$A2<>""')]],
    ['C2:D6', [cfRules.expr('C2>D2'), cfRules.expr('$B2="x"'), cfRules.dup()]],
    ['A2:A6', [cfRules.scale('$D$2'), cfRules.expr('A2=B2')]],
  ]) +
  '<dataValidations count="2"><dataValidation type="list" allowBlank="1" sqref="A2:A6"><formula1>$C$2:$C$6</formula1></dataValidation>' +
  '<dataValidation type="whole" operator="between" allowBlank="1" sqref="D2:D6"><formula1>$D$2</formula1><formula2>100</formula2></dataValidation></dataValidations>';
const CF_GTABLE =
  cfBlocks([
    ['A4:K13', [cfRules.expr('$D4>=80'), cfRules.cellIs('lessThan', '60')]],
    ['B4:C13 E4:F13', [cfRules.expr('B4<>C4'), cfRules.scale()]],
    ['D4:G13', [cfRules.expr('D4>E4'), cfRules.expr('$B4=""'), cfRules.bar()]],
    ['H4:I13', [cfRules.dup(), cfRules.top(3)]],
  ]) +
  '<dataValidations count="1"><dataValidation type="list" allowBlank="1" sqref="C4:C13"><formula1>$A$4:$A$13</formula1></dataValidation></dataValidations>';

// rules of the OTHER sheets that read the merged sheet (no golden fixture has any): expression / cellIs CF and list / whole DV
const cf = (sqref: string, ...formulas: string[]) =>
  `<conditionalFormatting sqref="${sqref}">${formulas.map((f, i) => `<cfRule type="${i % 2 ? 'cellIs' : 'expression'}"${i % 2 ? ' operator="greaterThan"' : ''} priority="${i + 1}"><formula>${f.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')}</formula></cfRule>`).join('')}</conditionalFormatting>`;
const dv = (...units: Array<[string, string, string?]>) =>
  `<dataValidations count="${units.length}">${units.map(([sqref, f1, f2]) => `<dataValidation type="${f2 ? 'whole' : 'list'}"${f2 ? ' operator="between"' : ''} allowBlank="1" sqref="${sqref}"><formula1>${f1}</formula1>${f2 ? `<formula2>${f2}</formula2>` : ''}</dataValidation>`).join('')}</dataValidations>`;
const EXT_GTABLE = { path: 'xl/worksheets/sheet2.xml', xml: cf('A1:B5', 'COUNTIF(名簿!$B$4:$B$13,A1)>0', '名簿!$D$4') + dv(['C1:C5', '名簿!$C$4:$C$13'], ['D1:D5', '名簿!$E$4', 'MAX(名簿!$F$4:$F$13)']) };
const EXT_TORTURE = { path: 'xl/worksheets/sheet2.xml', xml: cf('A1:C6', 'SUM(Data!$E$4:$E$23)>A1', 'Data!$B$4') + dv(['D1:D4', 'Data!$C$4:$C$23'], ['E1:E4', 'Data!$E$4', 'MAX(Data!$F$4:$F$23)']) };

const FIXTURES: Fixture[] = [
  { file: 'torture.xlsx', sheet: 'Data', region: { headerRow: 3, firstDataRow: 4, lastDataRow: 23 } },
  { file: 'r2-arrays.xlsx', sheet: 'A', region: { headerRow: 1, firstDataRow: 2, lastDataRow: 4 } },
  { file: 'pm-base-plain.xlsx', sheet: '学生', region: { headerRow: 1, firstDataRow: 2, lastDataRow: 6 } },
  { file: 'g-table.xlsx', sheet: '名簿', region: { headerRow: 3, firstDataRow: 4, lastDataRow: 13 } },
  // a table with a totals row (A1:D12)
  { file: 'torture.xlsx', sheet: 'Tbl', region: { headerRow: 1, firstDataRow: 2, lastDataRow: 11 }, totalsRow: 12 },
  // the rest of tests/fixtures/golden (round 4: every fixture that has a data sheet)
  { file: 'pm-base-table.xlsx', sheet: '学生', region: { headerRow: 1, firstDataRow: 2, lastDataRow: 6 } },
  { file: 'r2-table.xlsx', sheet: 'T', region: { headerRow: 1, firstDataRow: 2, lastDataRow: 4 } },
  { file: 'r2-rich.xlsx', sheet: 'S', region: { headerRow: 1, firstDataRow: 2, lastDataRow: 3 } },
  { file: 'pm-src.xlsx', sheet: '回答', region: { headerRow: 1, firstDataRow: 2, lastDataRow: 6 } },
  { file: 'r3-misc.xlsx', sheet: 'Tb', region: { headerRow: 1, firstDataRow: 2, lastDataRow: 4 } },
  { file: 'r3-misc.xlsx', sheet: 'Outline', region: { headerRow: 1, firstDataRow: 3, lastDataRow: 3 } },
  // title merges above the header row (unmerge + spread), as in invariant.test.ts
  { file: 'r3-misc.xlsx', sheet: 'Band', region: { headerRow: 12, firstDataRow: 13, lastDataRow: 12 } },
  { file: 'r3-misc.xlsx', sheet: 'Band', region: { headerRow: 1, firstDataRow: 3, lastDataRow: 11 } },
  // round 4: ColProps.customWidth = user-set WITHOUT bestFit. Auto / auto-fitted columns are widened by typed dates in the
  // doc exactly as the patcher does; a width set with setColumnWidth ends auto-fit (compare.ts also compares the flag)
  { file: 'torture.xlsx', sheet: 'Data', region: { headerRow: 3, firstDataRow: 4, lastDataRow: 23 }, cols: 'auto' },
  { file: 'torture.xlsx', sheet: 'Data', region: { headerRow: 3, firstDataRow: 4, lastDataRow: 23 }, cols: 'bestFit' },
  { file: 'g-table.xlsx', sheet: '名簿', region: { headerRow: 3, firstDataRow: 4, lastDataRow: 13 }, cols: 'auto' },
  { file: 'g-table.xlsx', sheet: '名簿', region: { headerRow: 3, firstDataRow: 4, lastDataRow: 13 }, cols: 'bestFit' },
  { file: 'pm-base-plain.xlsx', sheet: '学生', region: { headerRow: 1, firstDataRow: 2, lastDataRow: 6 }, cols: 'auto' },
  { file: 'pm-base-plain.xlsx', sheet: '学生', region: { headerRow: 1, firstDataRow: 2, lastDataRow: 6 }, cols: 'bestFit' },
  // round 4: CF / DV on the other sheets that read the merged sheet (they follow the ops like formulas)
  { file: 'g-table.xlsx', sheet: '名簿', region: { headerRow: 3, firstDataRow: 4, lastDataRow: 13 }, ext: EXT_GTABLE },
  { file: 'torture.xlsx', sheet: 'Data', region: { headerRow: 3, firstDataRow: 4, lastDataRow: 23 }, ext: EXT_TORTURE },
  // CF / DV stress on the merged sheet itself
  { file: 'pm-base-plain.xlsx', sheet: '学生', region: { headerRow: 1, firstDataRow: 2, lastDataRow: 6 }, rules: CF_PLAIN },
  { file: 'g-table.xlsx', sheet: '名簿', region: { headerRow: 3, firstDataRow: 4, lastDataRow: 13 }, rules: CF_GTABLE },
];

/** ops per sequence (DOC_FUZZ_STEPS for longer ones) */
const STEPS = Number(process.env.DOC_FUZZ_STEPS ?? 9);
const only = process.env.DOC_FUZZ_SEED ? Number(process.env.DOC_FUZZ_SEED) : null;
// DOC_FUZZ_COUNT=<n> sequences per fixture (big run: 300); DOC_FUZZ_SEED_BASE=<n> moves the seed range to explore new ones
const SEED_BASE = Number(process.env.DOC_FUZZ_SEED_BASE ?? 1000);
const SEEDS = only !== null ? [only] : Array.from({ length: Number(process.env.DOC_FUZZ_COUNT ?? 30) }, (_, i) => SEED_BASE + i);

describe('random op sequences: doc == readBook(replayLog)', () => {
  for (const fx of FIXTURES) {
    it(`${fx.file} ${fx.sheet} (header row ${fx.region.headerRow}${fx.cols ? `, cols ${fx.cols}` : ''}${fx.ext ? ', rules on another sheet' : ''}${fx.rules ? ', CF stress' : ''}): ${SEEDS.length} sequences`, async () => {
      let bytes = fx.cols ? await recol(golden(fx.file), fx.cols) : golden(fx.file);
      if (fx.ext) bytes = await withRulesOn(bytes, fx.ext.path, fx.ext.xml);
      if (fx.rules) bytes = await withRulesOn(bytes, 'xl/worksheets/sheet1.xml', fx.rules, true);
      const book = await readBook(bytes, fx.file);
      const failures: string[] = [];
      if (fx.ext) {
        // the variant really carries rules on the other sheet (2 CF + 2 DV)
        const s0 = createEditor(book, { sheet: fx.sheet, insertAt: 0, imported: [], appendAt: null, appended: [] }, {}, fx.region);
        expect(s0.doc.externalRules.map((e) => `${e.rule.kind}/${e.rule.type}`).sort()).toEqual(['cf/cellIs', 'cf/expression', 'dv/list', 'dv/whole']);
      }
      if (fx.rules) {
        // the stress variant really carries its rules (classic CF rules + DVs, plus whatever x14 the fixture has)
        const s0 = createEditor(book, { sheet: fx.sheet, insertAt: 0, imported: [], appendAt: null, appended: [] }, {}, fx.region);
        const classic = s0.doc.features.rules.filter((x) => !x.ext);
        expect(classic.filter((x) => x.kind === 'cf').length).toBe(fx.rules === CF_PLAIN ? 12 : 9);
        expect(classic.filter((x) => x.kind === 'dv').length).toBe(fx.rules === CF_PLAIN ? 2 : 1);
      }
      for (const seed of SEEDS) {
        const r = rng(seed * 7919 + fx.file.length);
        let s = createEditor(book, { sheet: fx.sheet, insertAt: 0, imported: [], appendAt: null, appended: [] }, {}, fx.region);
        const trail: string[] = [];
        for (let i = 0; i < STEPS; i++) {
          const { res, what } = step(s, r, !!fx.noInsert, fx.totalsRow);
          trail.push(res.ok ? what : `${what} [refused]`);
          if (res.ok) s = res.snapshot;
        }
        if (!s.log.length) continue;
        let diffs: string[];
        try {
          const after = await readBook(await replayLog(bytes, s.log), fx.file);
          diffs = compare(s.doc, after);
        } catch (e) {
          diffs = [`replay threw: ${(e as Error).message}`];
        }
        if (diffs.length) failures.push(`seed ${seed}: ${diffs.slice(0, 4).join(' | ')}\n    ops: ${trail.join('; ')}`);
      }
      expect(failures).toEqual([]);
    }, 180_000);
  }
});

// Undo / redo interleaved with the ops: whatever snapshot the history presents, its doc must still equal the replay of ITS
// log (snapshots are immutable and share structure; a cache or a WeakMap keyed by a shared object must not leak between them).
describe('random op sequences with undo / redo in between: present.doc == readBook(replayLog(present.log))', () => {
  const pick = FIXTURES.filter((f) => (f.file === 'torture.xlsx' && (f.sheet === 'Data' || f.sheet === 'Tbl') && !f.cols && !f.ext) || (f.file === 'g-table.xlsx' && (f.cols === 'auto' || f.ext)) || (f.file === 'r3-misc.xlsx' && f.region.headerRow === 12));
  for (const fx of pick) {
    it(`${fx.file} ${fx.sheet}${fx.cols ? ` (cols ${fx.cols})` : ''}${fx.ext ? ' (rules on another sheet)' : ''}: ${SEEDS.length} sequences`, async () => {
      let bytes = fx.cols ? await recol(golden(fx.file), fx.cols) : golden(fx.file);
      if (fx.ext) bytes = await withRulesOn(bytes, fx.ext.path, fx.ext.xml);
      const book = await readBook(bytes, fx.file);
      const failures: string[] = [];
      for (const seed of SEEDS) {
        const r = rng(seed * 6151 + fx.file.length);
        let h = initHistory(createEditor(book, { sheet: fx.sheet, insertAt: 0, imported: [], appendAt: null, appended: [] }, {}, fx.region));
        const trail: string[] = [];
        for (let i = 0; i < Math.max(STEPS, 12); i++) {
          const x = r();
          if (x < 0.15) {
            h = undo(h);
            trail.push('undo');
          } else if (x < 0.25) {
            h = redo(h);
            trail.push('redo');
          } else {
            const { res, what } = step(h.present, r, !!fx.noInsert, fx.totalsRow);
            trail.push(res.ok ? what : `${what} [refused]`);
            if (res.ok && res.snapshot !== h.present) h = commit(h, res.snapshot);
          }
        }
        if (!h.present.log.length) continue;
        let diffs: string[];
        try {
          diffs = compare(h.present.doc, await readBook(await replayLog(bytes, h.present.log), fx.file));
        } catch (e) {
          diffs = [`replay threw: ${(e as Error).message}`];
        }
        if (diffs.length) failures.push(`seed ${seed}: ${diffs.slice(0, 4).join(' | ')}\n    ops: ${trail.join('; ')}`);
      }
      expect(failures).toEqual([]);
    }, 180_000);
  }
});
