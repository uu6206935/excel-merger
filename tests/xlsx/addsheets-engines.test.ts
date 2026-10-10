/**
 * The editor (src/core/doc) and the built-in export engine (replayLog) on premerged packages that CONTAIN copied sheets:
 *  - createEditor accepts the package, the doc lists the copies among the sheet names;
 *  - the random op sequences of the doc fuzz (tests/doc/fuzzcore.ts: moves incl. non-contiguous, inserts, deletes, typed cells, widths) read
 *    back from replayLog exactly as the doc shows them, copies included (tests/doc/compare.ts);
 *  - the copies are not changed by any op on the merged sheet: after the replay every copied sheet equals what it was in the premerged
 *    package (values, formulas, styles, widths ...);
 *  - formulas typed into the merged sheet may read a copy.
 */
import fs from 'node:fs';
import path from 'node:path';
import { describe, expect, it } from 'vitest';
import { createEditor, moveColumns, setCellInput } from '../../src/core/doc';
import { analyzeSources, defaultPlan, detectRegion, joinSources, pickDataSheet } from '../../src/core/keys';
import type { KeySource } from '../../src/core/contracts';
import type { JoinType } from '../../src/core/types';
import { readBook, replayLog } from '../../src/core/xlsx';
import { compare } from '../doc/compare';
import { rng, step, TEXTS } from '../doc/fuzzcore';
import { SAMPLES_DIR, premergedOf, samplesAvailable, type SamplePremerged } from '../doc/sampleMerge';
import { compareSheets, lintPackage, lintText } from './addsheets-helpers';

async function load(name: string, id: string): Promise<KeySource> {
  const book = await readBook(new Uint8Array(fs.readFileSync(path.join(SAMPLES_DIR, name))), name, { id });
  const sheet = pickDataSheet(book);
  return { id, book, region: detectRegion(book, sheet), label: name.replace(/\.[^.]+$/, '') };
}

async function withCopies(baseName: string, otherNames: string[], opts: { joinType?: JoinType; dropOtherSheets?: boolean } = {}): Promise<SamplePremerged> {
  const base = await load(baseName, 's1');
  const others = await Promise.all(otherNames.map((n, i) => load(n, `s${i + 2}`)));
  const analysis = analyzeSources(base, others);
  const join = joinSources(analysis, { ...defaultPlan(analysis), joinType: opts.joinType ?? 'left' });
  return premergedOf({
    base, sources: others, rows: join.rows, imports: join.defaultImports, keyFill: join.keyFill, keyPairs: join.keyPairs,
    options: { formulaPolicy: 'keep-if-possible', extendTable: true, headerConflict: 'suffix', includeSourceSheets: true, dropOtherSheets: opts.dropOtherSheets ?? false },
  });
}

const ALL = ['02_面接評価.xlsx', '03_連絡先.csv', '04_適性検査.xls', '05_研修出欠.xlsx'];
const VARIANTS: Array<{ name: string; build: () => Promise<SamplePremerged> }> = [
  { name: 'base 01 + 02 03 04 05 (all sheets kept) + copies', build: () => withCopies('01_学生名簿.xlsx', ALL) },
  { name: 'base 01 + 02 03 05, join full, single-sheet base + copies', build: () => withCopies('01_学生名簿.xlsx', ['02_面接評価.xlsx', '03_連絡先.csv', '05_研修出欠.xlsx'], { joinType: 'full', dropOtherSheets: true }) },
  { name: 'base 02 + 01 05 + copies', build: () => withCopies('02_面接評価.xlsx', ['01_学生名簿.xlsx', '05_研修出欠.xlsx']) },
  { name: 'base 05 + 01 full + copies', build: () => withCopies('05_研修出欠.xlsx', ['01_学生名簿.xlsx'], { joinType: 'full' }) },
];

const SEEDS = Array.from({ length: Number(process.env.DOC_FUZZ_COUNT ?? 12) }, (_x, i) => 9000 + i);
const STEPS = Number(process.env.DOC_FUZZ_STEPS ?? 10);
const TEXT = [...TEXTS, '=B4&C4', '=SUM(K4:K63)', '=COUNTIF(B:B,B4)', '=VLOOKUP(A4,A:B,2,FALSE)', '=名簿!B4', 'メモ', '000100'];

describe.skipIf(!samplesAvailable())('the editor and replayLog on packages with copied sheets', () => {
  for (const v of VARIANTS) {
    it(`${v.name}: linter, createEditor, ${SEEDS.length} random op sequences (doc == replayLog), copies untouched`, async () => {
      const p = await v.build();
      expect(lintText(await lintPackage(p.bytes))).toBe('');
      expect(p.layout.addedSheets!.length).toBeGreaterThanOrEqual(3);
      const s0 = createEditor(p.book, p.layout, p.sourceInfo, p.region);
      for (const a of p.layout.addedSheets!) expect(s0.doc.sheetNames).toContain(a.name);
      const before = p.book;
      const failures: string[] = [];
      let ran = 0;
      for (const seed of SEEDS) {
        const r = rng(seed * 7919 + v.name.length);
        let s = s0;
        const trail: string[] = [];
        for (let i = 0; i < STEPS; i++) {
          const { res, what } = step(s, r, false, s.doc.lastDataRow + 2, TEXT);
          trail.push(res.ok ? what : `${what} [refused]`);
          if (res.ok) s = res.snapshot;
        }
        if (!s.log.length) continue;
        ran++;
        try {
          const bytes = await replayLog(p.bytes, s.log);
          const after = await readBook(bytes, 'after.xlsx');
          const diffs = compare(s.doc, after);
          // every copy is exactly what it was
          for (const a of p.layout.addedSheets!) {
            for (const d of compareSheets(before, a.name, after, a.name, { rename: new Map(), styles: true })) diffs.push(`copy ${a.name} ${d.where}: ${d.detail}`);
          }
          const lint = await lintPackage(bytes);
          if (lint.length) diffs.push(`linter: ${lintText(lint).split('\n')[0]}`);
          if (diffs.length) failures.push(`seed ${seed}: ${diffs.slice(0, 4).join(' | ')}\n    ops: ${trail.join('; ')}`);
        } catch (e) {
          failures.push(`seed ${seed}: replay threw: ${(e as Error).message}\n    ops: ${trail.join('; ')}`);
        }
      }
      expect(ran).toBeGreaterThan(0);
      expect(failures).toEqual([]);
    }, 600_000);
  }

  it('a formula typed into the merged sheet may read a copy; moving columns keeps it and the copy', async () => {
    const p = await withCopies('01_学生名簿.xlsx', ['02_面接評価.xlsx'], { dropOtherSheets: true });
    const copy = p.layout.addedSheets![0].name;
    const s0 = createEditor(p.book, p.layout, p.sourceInfo, p.region);
    const typed = setCellInput(s0, s0.doc.columns[0].id, s0.doc.firstDataRow, `=COUNTA('${copy}'!A:A)`);
    expect(typed.ok).toBe(true);
    if (!typed.ok) return;
    const moved = moveColumns(typed.snapshot, [typed.snapshot.doc.columns[0].id], 3, { unmergeTitles: true });
    if (!moved.ok) return;
    const after = await readBook(await replayLog(p.bytes, moved.snapshot.log), 'after.xlsx');
    expect(compare(moved.snapshot.doc, after)).toEqual([]);
    expect(compareSheets(p.book, copy, after, copy, { rename: new Map(), styles: true })).toEqual([]);
  });
});
