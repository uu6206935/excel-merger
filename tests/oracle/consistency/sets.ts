/**
 * Scenario sets of the consistency oracle (task V1): which base + other files are merged, and with which merge parameters.
 *  - smp-*  : the app's own samples (samples/*): 01_学生名簿 (title merge, hidden helper column, CF/DV, 集計 sheet with whole-column and
 *             cross-sheet references, names, chart, pivot) and 02_面接評価 (平均/判定 formulas, a table sheet) as bases;
 *  - rich-* : the generated "rich base" (gen.ts: relative / absolute / mixed / whole-column / cross-sheet / name / structured / lookup
 *             formulas, total row, report sheet, CF + DV) with the formula sources of tests/fixtures/formulaimport (F1's corpus: shuffled
 *             order, formulas of every kind of the round-8 list, with and without stored results);
 *  - row-*  : the rich base with generated second files whose rows are reversed / shuffled / have extra and missing people / duplicates.
 */
import fs from 'node:fs';
import path from 'node:path';
import type { KeySource } from '../../../src/core/contracts';
import { buildRichBase, buildRowSource, ROW_ORDERS, type BaseVariant, type RowOrder } from './gen';
import { FIXTURES, loadKeySourceBytes, loadKeySourceFile, SAMPLES, type MergeParams } from './lib';

export interface Built {
  base: KeySource;
  others: KeySource[];
}

export interface Variant {
  name: string;
  params: Partial<MergeParams>;
}

export interface Ctx {
  /** directory with the authored (Excel-saved) twins of the generated workbooks: gen/cached/<name>.xlsx, raw ones in gen/raw */
  genDir: string;
}

export interface SetDef {
  id: string;
  title: string;
  variants: Variant[];
  /** relative weight when scenarios are distributed over the sets */
  weight?: number;
  /** the key texts are written the same way in both files: a placed row must agree on a key pair (srcCheck row-person) */
  strictKeys?: boolean;
  build(ctx: Ctx): Promise<Built>;
}

/** a generated workbook as the oracle uses it: the Excel-saved twin (results stored) when it exists, else the raw one */
export function genBytes(ctx: Ctx, name: string, cached: boolean): Uint8Array {
  const c = path.join(ctx.genDir, 'cached', `${name}.xlsx`);
  const r = path.join(ctx.genDir, 'raw', `${name}.xlsx`);
  if (cached && fs.existsSync(c)) return new Uint8Array(fs.readFileSync(c));
  return new Uint8Array(fs.readFileSync(r));
}

/** Every generated workbook (name -> bytes) the sets need: written to gen/raw, authored by Excel into gen/cached. */
export async function generatedWorkbooks(): Promise<Map<string, Uint8Array>> {
  const out = new Map<string, Uint8Array>();
  for (const v of ['plain', 'table', 'title', 'quoted'] as BaseVariant[]) out.set(`base-${v}`, (await buildRichBase(v)).bytes);
  for (const o of ROW_ORDERS) {
    out.set(`rows-${o}`, (await buildRowSource(o, 11)).bytes);
    out.set(`rowstbl-${o}`, (await buildRowSource(o, 12, { table: true })).bytes);
  }
  return out;
}

const LEFT_DROP: Variant = { name: 'left+drop', params: { joinType: 'left', dropOtherSheets: true } };
const FULL_DROP: Variant = { name: 'full+drop', params: { joinType: 'full', dropOtherSheets: true } };
const LEFT_KEEP: Variant = { name: 'left+keep', params: { joinType: 'left', dropOtherSheets: false } };
const FULL_KEEP: Variant = { name: 'full+keep', params: { joinType: 'full', dropOtherSheets: false } };
const ALL4 = [LEFT_DROP, FULL_DROP, LEFT_KEEP, FULL_KEEP];
/** every column of the source is imported, the key columns too (their formulas are the ones that read key columns; headers get the file-name suffix) */
const ALLCOLS_LEFT: Variant = { name: 'allcols+left+keep', params: { joinType: 'left', dropOtherSheets: false, imports: 'all' } };
const ALLCOLS_FULL: Variant = { name: 'allcols+full+drop', params: { joinType: 'full', dropOtherSheets: true, imports: 'all' } };

const sample = (f: string, id: string) => loadKeySourceFile(path.join(SAMPLES, f), id);

const FIXTURE_ITEMS = [
  'f01_samerow__shuf', 'f02_shared__shuf', 'f03_keycol__shuf', 'f04_lookups__shuf', 'f05_ranges__shuf', 'f06_running__shuf', 'f07_othersheet__shuf',
  'f08_names__shuf', 'f09_tables__shuf', 'f10_arrays__shuf', 'f11_values__shuf', 'f12_footer__shuf', 'f13_hyperlink__shuf', 'f14_twoheader__shuf', 'f16_unsafe__shuf',
];

export function buildSets(): SetDef[] {
  const sets: SetDef[] = [];
  // ---- the app's samples
  sets.push({
    id: 'smp-demo',
    title: 'base 01_学生名簿 + 02 (ID) + 03 (氏名) + 05 (氏名＋大学): the demo of samples/README.md',
    variants: ALL4,
    weight: 3,
    build: async () => ({ base: await sample('01_学生名簿.xlsx', 's1'), others: [await sample('02_面接評価.xlsx', 's2'), await sample('03_連絡先.csv', 's3'), await sample('05_研修出欠.xlsx', 's4')] }),
  });
  sets.push({
    id: 'smp-all5',
    title: 'base 01 + 02 03 04 (.xls, フリガナ) 05',
    variants: [LEFT_KEEP, FULL_DROP],
    weight: 2,
    build: async () => ({ base: await sample('01_学生名簿.xlsx', 's1'), others: [await sample('02_面接評価.xlsx', 's2'), await sample('03_連絡先.csv', 's3'), await sample('04_適性検査.xls', 's4'), await sample('05_研修出欠.xlsx', 's5')] }),
  });
  sets.push({
    id: 'smp-02base',
    title: 'base 02_面接評価 (平均 / 判定 formulas, a table sheet) + 01',
    variants: [LEFT_DROP, FULL_DROP, LEFT_KEEP, FULL_KEEP],
    weight: 3,
    build: async () => ({ base: await sample('02_面接評価.xlsx', 's1'), others: [await sample('01_学生名簿.xlsx', 's2')] }),
  });
  sets.push({
    id: 'smp-05base',
    title: 'base 05_研修出欠 (出席回数) + 01 + 02',
    variants: [LEFT_DROP, FULL_KEEP],
    weight: 1,
    build: async () => ({ base: await sample('05_研修出欠.xlsx', 's1'), others: [await sample('01_学生名簿.xlsx', 's2'), await sample('02_面接評価.xlsx', 's3')] }),
  });
  // ---- the rich base with F1's formula sources
  for (const item of FIXTURE_ITEMS) {
    for (const cached of [true, false]) {
      sets.push({
        id: `rich-${item.replace(/__shuf$/, '')}-${cached ? 'c' : 'n'}`,
        strictKeys: true,
        title: `rich base + ${item} (${cached ? 'results stored, saved by Excel' : 'no stored results (openpyxl / XlsxWriter)'})`,
        variants: /f03_keycol|f09_tables|f04_lookups|f08_names/.test(item) ? [...ALL4, ALLCOLS_LEFT, ALLCOLS_FULL] : ALL4,
        build: async (ctx) => {
          const base = await loadKeySourceBytes(genBytes(ctx, 'base-plain', true), '名簿_rich.xlsx', 'b');
          const src = await loadKeySourceFile(path.join(FIXTURES, cached ? 'cached' : 'nocache', `${item}.xlsx`), 's2');
          return { base, others: [src] };
        },
      });
    }
  }
  sets.push({
    id: 'rich-g01-c',
    strictKeys: true,
    title: 'the app sample list (01_学生名簿 copy) + F1 g01 source (results stored)',
    variants: [LEFT_DROP, FULL_KEEP],
    build: async () => ({ base: await loadKeySourceFile(path.join(FIXTURES, 'base_samples.xlsx'), 'b'), others: [await loadKeySourceFile(path.join(FIXTURES, 'cached', 'g01_samples__shuf.xlsx'), 's2')] }),
  });
  sets.push({
    id: 'rich-g01-n',
    strictKeys: true,
    title: 'the app sample list (01_学生名簿 copy) + F1 g01 source (no stored results)',
    variants: [LEFT_DROP, FULL_KEEP],
    build: async () => ({ base: await loadKeySourceFile(path.join(FIXTURES, 'base_samples.xlsx'), 'b'), others: [await loadKeySourceFile(path.join(FIXTURES, 'nocache', 'g01_samples__shuf.xlsx'), 's2')] }),
  });
  sets.push({
    id: 'rich-multi',
    strictKeys: true,
    title: 'rich base + f01 + f04 + f09 (three formula sources at once)',
    variants: ALL4,
    weight: 2,
    build: async (ctx) => ({
      base: await loadKeySourceBytes(genBytes(ctx, 'base-plain', true), '名簿_rich.xlsx', 'b'),
      others: [
        await loadKeySourceFile(path.join(FIXTURES, 'cached', 'f01_samerow__shuf.xlsx'), 's2'),
        await loadKeySourceFile(path.join(FIXTURES, 'nocache', 'f04_lookups__shuf.xlsx'), 's3'),
        await loadKeySourceFile(path.join(FIXTURES, 'cached', 'f09_tables__shuf.xlsx'), 's4'),
      ],
    }),
  });
  // ---- table / title bases
  for (const bv of ['table', 'title', 'quoted'] as BaseVariant[]) {
    sets.push({
      id: `rich-${bv}`,
      strictKeys: true,
      title: `rich ${bv} base + rows-shuf source`,
      variants: ALL4,
      weight: 2,
      build: async (ctx) => ({
        base: await loadKeySourceBytes(genBytes(ctx, `base-${bv}`, true), `名簿_${bv}.xlsx`, 'b'),
        others: [await loadKeySourceBytes(genBytes(ctx, bv === 'table' ? 'rowstbl-shuf' : 'rows-shuf', true), '評価.xlsx', 's2')],
      }),
    });
  }
  // ---- the row-rearrangement family
  for (const o of ROW_ORDERS) {
    for (const cached of [true, false]) {
      sets.push({
        id: `row-${o}-${cached ? 'c' : 'n'}`,
        strictKeys: true,
        title: `rich base + a source whose rows are ${o} (${cached ? 'results stored' : 'no stored results'})`,
        variants: [...ALL4, ALLCOLS_LEFT, ALLCOLS_FULL],
        weight: o === 'same' ? 0.5 : 1,
        build: async (ctx) => ({
          base: await loadKeySourceBytes(genBytes(ctx, 'base-plain', true), '名簿_rich.xlsx', 'b'),
          others: [await loadKeySourceBytes(genBytes(ctx, `rows-${o}`, cached), '評価.xlsx', 's2')],
        }),
      });
    }
    sets.push({
      id: `rowtbl-${o}`,
      strictKeys: true,
      title: `rich table base + a source Table whose rows are ${o}`,
      variants: [LEFT_DROP, FULL_KEEP],
      weight: 0.5,
      build: async (ctx) => ({
        base: await loadKeySourceBytes(genBytes(ctx, 'base-table', true), '名簿_table.xlsx', 'b'),
        others: [await loadKeySourceBytes(genBytes(ctx, `rowstbl-${o}`, true), '評価.xlsx', 's2')],
      }),
    });
  }
  return sets;
}

export type { RowOrder };
