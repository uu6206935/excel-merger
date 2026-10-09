// A real premerged workbook built from the app's sample folder (samples/*): 01_学生名簿.xlsx (base) + the other samples,
// through the real pipeline (readBook -> keys -> buildPremerged), exactly like the controller does. Shared by the doc
// fuzz / invariant tests that want a merge result instead of a hand-made fixture.
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { analyzeSources, defaultPlan, detectRegion, joinSources, pickDataSheet } from '../../src/core/keys';
import { buildPremerged, readBook } from '../../src/core/xlsx';
import type { KeySource, PremergeInput, PremergeLayout } from '../../src/core/contracts';
import type { JoinType, MergedDoc, SourceBook } from '../../src/core/types';

const here = path.dirname(fileURLToPath(import.meta.url));
export const SAMPLES_DIR = path.resolve(here, '../../samples');

export const SAMPLE_FILES = ['01_学生名簿.xlsx', '02_面接評価.xlsx', '03_連絡先.csv', '04_適性検査.xls', '05_研修出欠.xlsx'] as const;

export const samplesAvailable = (): boolean => SAMPLE_FILES.every((f) => fs.existsSync(path.join(SAMPLES_DIR, f)));

export interface SamplePremerged {
  /** the premerged package */
  bytes: Uint8Array;
  layout: PremergeLayout;
  /** the premerged package read back (what createEditor is given) */
  book: SourceBook;
  sourceInfo: MergedDoc['sourceInfo'];
  region: { headerRow: number; firstDataRow: number; lastDataRow: number };
  /** name of the merged sheet */
  sheet: string;
  /** labels of the files that were merged in (others than the base) */
  merged: string[];
}

/**
 * Merge `others` into the base sample (default: every other sample that has usable join keys). The base is always the
 * first file; `joinType` 'full' also appends rows that exist only in the other files.
 */
export async function buildSamplePremerged(
  opts: { base?: string; others?: string[]; joinType?: JoinType } = {},
): Promise<SamplePremerged> {
  const baseName = opts.base ?? SAMPLE_FILES[0];
  const otherNames = opts.others ?? SAMPLE_FILES.filter((f) => f !== baseName && !f.startsWith('04_'));
  const load = async (name: string, id: string): Promise<KeySource> => {
    const book = await readBook(new Uint8Array(fs.readFileSync(path.join(SAMPLES_DIR, name))), name, { id });
    const sheet = pickDataSheet(book);
    return { id, book, region: detectRegion(book, sheet), label: name.replace(/\.[^.]+$/, '') };
  };
  const base = await load(baseName, 's1');
  const others = await Promise.all(otherNames.map((n, i) => load(n, `s${i + 2}`)));
  const analysis = analyzeSources(base, others);
  const plan = { ...defaultPlan(analysis), joinType: opts.joinType ?? 'left' };
  const join = joinSources(analysis, plan);
  return premergedOf({
    base,
    sources: others,
    rows: join.rows,
    imports: join.defaultImports,
    keyFill: join.keyFill,
    options: { formulaPolicy: 'keep-if-possible', extendTable: true, headerConflict: 'suffix' },
  });
}

/** A premerge input (a hand-made scenario or the keys module's plan) through buildPremerged -> readBook, ready for createEditor. */
export async function premergedOf(input: PremergeInput): Promise<SamplePremerged> {
  const { base, sources: others } = input;
  const res = await buildPremerged(input);
  const book = await readBook(res.bytes, base.book.fileName, { id: 'merged' });
  const colors = ['#2563eb', '#15803d', '#c2410c', '#7c3aed', '#0e7490'];
  const sourceInfo: MergedDoc['sourceInfo'] = {};
  [base, ...others].forEach((k, i) => (sourceInfo[k.id] = { label: k.label, color: colors[i % colors.length] }));
  let last = base.region.lastDataRow;
  for (const a of res.layout.appended) if (a.row > last) last = a.row;
  return {
    bytes: res.bytes,
    layout: res.layout,
    book,
    sourceInfo,
    region: { headerRow: base.region.headerRow, firstDataRow: base.region.firstDataRow, lastDataRow: last },
    sheet: res.layout.sheet,
    merged: others.map((o) => o.label),
  };
}
