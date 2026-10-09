// "Never produce #REF!" (docs/ARCHITECTURE.md key decision 4) and the delete dry run, as a property over random op
// sequences (the fuzzcore generator): moves and inserts never add a #REF! to any formula of the workbook, a delete that
// converts its dependents to values leaves none in a formula either (names and CF / DV rules cannot be converted: the
// dry run lists them), and analyzeDelete's dry run (findBreaks) predicts exactly the names that turn into #REF!.
// DOC_FUZZ_COUNT=<n> sequences per workbook (default 40).
import { describe, expect, it } from 'vitest';
import { analyzeDelete, createEditor } from '../../src/core/doc';
import { deleteOpsFor } from '../../src/core/doc/delete';
import { findBreaks } from '../../src/core/doc/deps';
import { readBook } from '../../src/core/xlsx';
import type { EditorSnapshot, MergedDoc } from '../../src/core/types';
import { insideTableScenario, scenario } from '../golden/premerge-scenario';
import { golden } from './compare';
import { rng, step } from './fuzzcore';
import { buildSamplePremerged, premergedOf, samplesAvailable, type SamplePremerged } from './sampleMerge';

const hasRef = (f: string) => f.includes('#REF!');

/** formulas of the merged sheet and of the other sheets that contain #REF! */
function refFormulas(d: MergedDoc): number {
  let n = 0;
  for (const c of d.columns) for (const x of c.cells) if (x?.f !== undefined && hasRef(x.f)) n++;
  for (const e of d.externalFormulas) if (hasRef(e.f)) n++;
  return n;
}
/** defined names (without the hidden filter database, which follows the autofilter) that contain #REF! */
const refNames = (d: MergedDoc): number => d.names.filter((x) => hasRef(x.refersTo) && !/_FilterDatabase$/i.test(x.name)).length;

const GOLD = [
  ['torture.xlsx', 'Data', { headerRow: 3, firstDataRow: 4, lastDataRow: 23 }],
  ['g-table.xlsx', '名簿', { headerRow: 3, firstDataRow: 4, lastDataRow: 13 }],
  ['pm-base-plain.xlsx', '学生', { headerRow: 1, firstDataRow: 2, lastDataRow: 6 }],
  ['pm-base-table.xlsx', '学生', { headerRow: 1, firstDataRow: 2, lastDataRow: 6 }],
  ['torture.xlsx', 'Tbl', { headerRow: 1, firstDataRow: 2, lastDataRow: 11 }],
  ['r2-table.xlsx', 'T', { headerRow: 1, firstDataRow: 2, lastDataRow: 4 }],
  ['r2-arrays.xlsx', 'A', { headerRow: 1, firstDataRow: 2, lastDataRow: 4 }],
] as const;

const COUNT = Number(process.env.DOC_FUZZ_COUNT ?? 40);
const STEPS = 20;

function walk(label: string, make: () => EditorSnapshot, seed0: number): string[] {
  const bad: string[] = [];
  for (let k = 0; k < COUNT; k++) {
    const r = rng(seed0 + k);
    let s = make();
    const trail: string[] = [];
    for (let i = 0; i < STEPS; i++) {
      const { res, what } = step(s, r, false, undefined);
      trail.push(res.ok ? what : `${what} [refused]`);
      if (!res.ok) continue;
      const fail = (msg: string) => bad.push(`${label} seed ${seed0 + k}: ${msg}\n    ops: ${trail.join('; ')}`);
      const isDelete = what.startsWith('delete');
      if (isDelete) {
        // the dry run lists exactly the names the delete turns into #REF! (they cannot be converted to values)
        const idx = what.replace(/^delete /, '').replace(/ \(convert\)$/, '').split(',').map(Number);
        const predicted = findBreaks(s.doc, deleteOpsFor(idx), new Set(idx)).names.filter((x) => !/_FilterDatabase$/i.test(x.name)).length;
        if (refNames(res.snapshot.doc) - refNames(s.doc) !== predicted) fail(`${what}: names predicted ${predicted}, actual ${refNames(res.snapshot.doc) - refNames(s.doc)}`);
        const impact = analyzeDelete(s, idx.map((i) => s.doc.columns[i].id));
        const actualBreaks = refFormulas(res.snapshot.doc) - refFormulas(s.doc);
        if (!what.includes('(convert)') && actualBreaks > 0 && impact.breakCount === 0) fail(`${what}: ${actualBreaks} formulas became #REF! but analyzeDelete found none`);
      }
      if (!what.startsWith('type') && (!isDelete || what.includes('(convert)')) && refFormulas(res.snapshot.doc) > refFormulas(s.doc)) {
        fail(`${what} added #REF! formulas (${refFormulas(s.doc)} -> ${refFormulas(res.snapshot.doc)})`);
      }
      s = res.snapshot;
    }
  }
  return bad;
}

describe('no #REF! from moves, inserts and converting deletes; the delete dry run predicts the names', () => {
  for (const [file, sheet, region] of GOLD) {
    it(`${file} ${sheet}`, async () => {
      const book = await readBook(golden(file), file);
      expect(walk(`${file}/${sheet}`, () => createEditor(book, { sheet, insertAt: 0, imported: [], appendAt: null, appended: [] }, {}, region), 9000)).toEqual([]);
    });
  }
  const scenarios: Array<[string, () => Promise<SamplePremerged>, boolean]> = [
    ['sample merge (01 + 02 03 05)', () => buildSamplePremerged({}), true],
    ['sample merge, full join', () => buildSamplePremerged({ joinType: 'full' }), true],
    ['scenario table', () => premergedOf(scenario('table')), false],
    ['scenario plain', () => premergedOf(scenario('plain')), false],
    ['scenario imports inside a table', () => premergedOf(insideTableScenario()), false],
  ];
  for (const [name, build, needsSamples] of scenarios) {
    it.skipIf(needsSamples && !samplesAvailable())(name, async () => {
      const p = await build();
      expect(walk(name, () => createEditor(p.book, p.layout, p.sourceInfo, p.region), 7000)).toEqual([]);
    });
  }
});
