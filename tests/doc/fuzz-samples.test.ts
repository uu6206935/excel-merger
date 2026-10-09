// Randomized op sequences (same generator as fuzz.test.ts) over REAL merge results: the app's sample folder (samples/*)
// merged through readBook -> keys -> buildPremerged (tests/doc/sampleMerge.ts), i.e. a base sheet with a merged title, a
// frozen pane, an autofilter, CF / DV (incl. x14 lists), a hidden helper column, aggregate rows below the table, other
// sheets that reference the list (charts, pivot, names) plus imported columns / appended rows with source styles.
// A second group runs the same over the patcher's Excel-authored premerge scenarios (tests/golden/premerge-scenario.ts:
// Excel Table / AutoFilter bases, imports inside a table, appended rows, rich text, key fill), which need no samples/.
// doc vs readBook(replayLog(premerged, log)) must agree for every sequence the doc accepts.
// A failure names the variant and seed; re-run with DOC_FUZZ_SEED=<seed> (and -t "<variant>") to replay one sequence.
// DOC_FUZZ_COUNT=300 is the big run (default 20 keeps the suite quick). Skipped when samples/ was not generated.
import { describe, expect, it } from 'vitest';
import { createEditor } from '../../src/core/doc';
import { readBook, replayLog } from '../../src/core/xlsx';
import { compare } from './compare';
import { rng, step, TEXTS } from './fuzzcore';
import { insideTableScenario, keyFillScenario, richScenario, scenario } from '../golden/premerge-scenario';
import { buildSamplePremerged, premergedOf, samplesAvailable, type SamplePremerged } from './sampleMerge';

interface Variant {
  name: string;
  build: () => Promise<SamplePremerged>;
}
const sample = (name: string, opts: Parameters<typeof buildSamplePremerged>[0]): Variant => ({ name, build: () => buildSamplePremerged(opts) });
const VARIANTS: Variant[] = [
  // the demo of samples/README.md: base 01 + 02 (ID) + 03 (氏名) + 05 (氏名＋大学)
  sample('base 01 + 02 03 05', {}),
  // join 'full': rows that exist only in the other files are appended below the table (before the aggregate rows)
  sample('base 01 + 02 03 05, full join', { joinType: 'full' }),
  // the legacy .xls (フリガナ key) as well: xls-origin style records next to xlsx ones
  sample('base 01 + 02 03 04 05', { others: ['02_面接評価.xlsx', '03_連絡先.csv', '04_適性検査.xls', '05_研修出欠.xlsx'] }),
  // other bases: no title merge / no other sheets / a CSV or xls base
  sample('base 02 + 01', { base: '02_面接評価.xlsx', others: ['01_学生名簿.xlsx'] }),
  sample('base 02 + 01, full join', { base: '02_面接評価.xlsx', others: ['01_学生名簿.xlsx'], joinType: 'full' }),
  sample('base 05 + 01', { base: '05_研修出欠.xlsx', others: ['01_学生名簿.xlsx'] }),
  sample('base 03 (csv) + 01 05', { base: '03_連絡先.csv', others: ['01_学生名簿.xlsx', '05_研修出欠.xlsx'] }),
  sample('base 04 (xls) + 01', { base: '04_適性検査.xls', others: ['01_学生名簿.xlsx'] }),
];

// the patcher's Excel-authored premerge scenarios (tests/golden/premerge-scenario.ts): a base that is an Excel Table or has
// an AutoFilter + CF + names + a chart sheet, a source with styled / formula columns, an appended row (join 'full'),
// columns imported INSIDE a table, rich text. They do not need samples/.
const SCENARIOS: Variant[] = [
  { name: 'scenario plain (autofilter base, appended row)', build: () => premergedOf(scenario('plain')) },
  { name: 'scenario table (ListObject extended over the imports)', build: () => premergedOf(scenario('table')) },
  { name: 'scenario imports inside a table (g-table)', build: () => premergedOf(insideTableScenario()) },
  { name: 'scenario rich text source', build: () => premergedOf(richScenario()) },
  { name: 'scenario key fill (combined / split key cells)', build: () => premergedOf(keyFillScenario()) },
];

/** what a user might type into the merged sample: typed ids, formulas over the list, aggregates, lookups */
const SAMPLE_TEXTS = [
  ...TEXTS,
  '000100', "'000100", '=B4&C4', '=B4*2', '=SUM(J4:K4)', '=COUNTIF(B:B,B4)', '=SUM(K4:K63)', '=AVERAGE(K:K)', '=VLOOKUP(A4,A:B,2,FALSE)',
  '=INDEX(B4:B63,2)', '=IF(K4>=4,"◎","")', '=$A$4:$A$20', '=UNIQUE(D4:D63)', '=TEXT(I4,"yyyy/m/d")', '2027/4/1', 'メモ', '  ', '=SUM(A:A)', '=名簿!B4',
  // the premerge scenarios (data rows 2..6, sheet 学生 / 名簿): same-row formulas, column aggregates, structured refs, other sheets
  '=B2&C2', '=D2*2', '=SUM(D2:D6)', '=COUNTA(B2:B6)', '=AVERAGE(D:D)', '=学生!B2', '=Report!A1', '=SUM(tblStudents[点数])', '=tblStudents[[#This Row],[点数]]*2', '=[@点数]+1',
];

/** ops per sequence (DOC_FUZZ_STEPS for longer ones) */
const STEPS = Number(process.env.DOC_FUZZ_STEPS ?? 10);
const only = process.env.DOC_FUZZ_SEED ? Number(process.env.DOC_FUZZ_SEED) : null;
const COUNT = Number(process.env.DOC_FUZZ_COUNT ?? 20);
const SEED_BASE = Number(process.env.DOC_FUZZ_SEED_BASE ?? 5000); // DOC_FUZZ_SEED_BASE=<n> explores a new seed range
const SEEDS = only !== null ? [only] : Array.from({ length: COUNT }, (_, i) => SEED_BASE + i);

function runVariants(variants: Variant[]) {
  for (const v of variants) {
    it(`${v.name}: ${SEEDS.length} sequences`, async () => {
      const p: SamplePremerged = await v.build();
      const failures: string[] = [];
      for (const seed of SEEDS) {
        const r = rng(seed * 104729 + v.name.length);
        let s = createEditor(p.book, p.layout, p.sourceInfo, p.region);
        const trail: string[] = [];
        for (let i = 0; i < STEPS; i++) {
          // typed texts also go to the aggregate rows below the table (first one = 2 rows under the last data row)
          const { res, what } = step(s, r, false, s.doc.lastDataRow + 2, SAMPLE_TEXTS);
          trail.push(res.ok ? what : `${what} [refused]`);
          if (res.ok) s = res.snapshot;
        }
        if (!s.log.length) continue;
        let diffs: string[];
        try {
          const after = await readBook(await replayLog(p.bytes, s.log), 'after.xlsx');
          diffs = compare(s.doc, after);
        } catch (e) {
          diffs = [`replay threw: ${(e as Error).message}`];
        }
        if (diffs.length) failures.push(`seed ${seed}: ${diffs.slice(0, 4).join(' | ')}\n    ops: ${trail.join('; ')}`);
      }
      expect(failures).toEqual([]);
    }, 600_000);
  }
}

describe.skipIf(!samplesAvailable())('random op sequences on merged samples: doc == readBook(replayLog)', () => runVariants(VARIANTS));
describe('random op sequences on the premerge scenarios: doc == readBook(replayLog)', () => runVariants(SCENARIOS));
