/**
 * The app's own sample files (samples/, read with the app's reader) against samples/truth.json: regions (the 人数 /
 * 平均 footer below samples/01's table is not data), matches per file, check items and their other-file rows.
 */
import fs from 'node:fs';
import path from 'node:path';
import { beforeAll, describe, expect, it } from 'vitest';
import type { JoinResult, KeyAnalysis, KeySource } from '../../src/core/contracts';
import type { SourceBook } from '../../src/core/types';
import { readBook } from '../../src/core/xlsx';
import { analyzeSources, defaultPlan, detectRegion, joinSources, pickDataSheet } from '../../src/core/keys';
import { buildTable } from '../../src/core/keys/region';

const dir = path.resolve(__dirname, '../../samples');
const FILES = { B: '01_学生名簿.xlsx', IV: '02_面接評価.xlsx', CT: '03_連絡先.csv', AP: '04_適性検査.xls', AT: '05_研修出欠.xlsx' } as const;
type Id = keyof typeof FILES;
const TRUTH_KEY: Record<Exclude<Id, 'B'>, string> = { IV: 'interview', CT: 'contacts', AP: 'aptitude', AT: 'attendance' };

interface TruthFile {
  headerRow: number;
  rows: Array<{ row: number; baseRow: number | null; resubmitted?: boolean; extra?: boolean }>;
  expected: { matchedBaseRows: number; unmatchedBaseRows: number[]; extraRows: number[] };
}
const truth = JSON.parse(fs.readFileSync(path.join(dir, 'truth.json'), 'utf8')) as {
  base: { sheet: string; headerRow: number; firstDataRow: number; lastDataRow: number; footerRows: number[]; students: unknown[]; sameNamePair: number[] };
} & Record<string, TruthFile>;

const books = {} as Record<Id, SourceBook>;
const srcs = {} as Record<Id, KeySource>;
let an: KeyAnalysis;
let res: JoinResult;

beforeAll(async () => {
  for (const [id, f] of Object.entries(FILES) as Array<[Id, string]>) {
    const b = await readBook(new Uint8Array(fs.readFileSync(path.join(dir, f))), f, { id });
    books[id] = b;
    const sheet = pickDataSheet(b);
    srcs[id] = { id, book: b, region: detectRegion(b, sheet), label: f.replace(/\.[^.]+$/, '') };
  }
  an = analyzeSources(srcs.B, [srcs.IV, srcs.CT, srcs.AP, srcs.AT]);
  res = joinSources(an, defaultPlan(an));
});

describe('samples: regions', () => {
  it('01_学生名簿: header row 3, 60 students in rows 4–63; the 人数 / 平均 footer (65–66) below a blank row is not data', () => {
    const tb = truth.base;
    expect(srcs.B.region).toEqual({ sheet: tb.sheet, headerRow: tb.headerRow, firstDataRow: tb.firstDataRow, lastDataRow: tb.lastDataRow, firstCol: 1, lastCol: 12 });
    expect(tb.footerRows).toEqual([65, 66]);
    const t = buildTable(srcs.B);
    expect(t.rows.length).toBe(60);
    expect(t.rows.length).toBe(tb.students.length);
    expect(t.excelRow).toEqual(Array.from({ length: 60 }, (_, i) => i + 4));
    expect(t.skippedRows).toEqual([]);
    // the user extending the region over the footer still never matches the footer rows
    const ext = buildTable({ ...srcs.B, region: detectRegion(books.B, '名簿', { lastDataRow: 66 }) });
    expect(ext.excelRow).toEqual(t.excelRow);
    expect(ext.skippedRows).toEqual([64, 65, 66]);
  });

  it('the other samples: sheet, header row and last data row', () => {
    expect(srcs.IV.region).toMatchObject({ sheet: '面接評価', headerRow: 1, firstDataRow: 2, lastDataRow: 64, firstCol: 1, lastCol: 9 });
    expect(srcs.CT.region).toMatchObject({ headerRow: 1, firstDataRow: 2, lastDataRow: 48, firstCol: 1, lastCol: 3 });
    expect(srcs.AP.region).toMatchObject({ sheet: '適性検査結果', headerRow: 1, firstDataRow: 2, lastDataRow: 43, firstCol: 1, lastCol: 8 });
    expect(srcs.AT.region).toMatchObject({ sheet: '出欠', headerRow: 2, firstDataRow: 3, lastDataRow: 34, firstCol: 1, lastCol: 7 });
    for (const id of ['IV', 'CT', 'AP', 'AT'] as const) {
      const tf = truth[TRUTH_KEY[id]];
      expect(srcs[id].region.headerRow).toBe(tf.headerRow);
      expect(srcs[id].region.lastDataRow).toBe(Math.max(...tf.rows.map((r) => r.row)));
    }
  });

  it('opts.headerRow reproduces the detected region; opts.lastDataRow is taken as given', () => {
    expect(detectRegion(books.B, '名簿', { headerRow: 3 })).toEqual(srcs.B.region);
    expect(detectRegion(books.AT, '出欠', { headerRow: 2 })).toEqual(srcs.AT.region);
    expect(detectRegion(books.B, '名簿', { lastDataRow: 40 })).toMatchObject({ headerRow: 3, firstDataRow: 4, lastDataRow: 40 });
  });
});

describe('samples: join against truth.json', () => {
  it('summary: 基準 60 名, every student matched by at least one file', () => {
    expect(res.summary).toEqual({ baseRows: 60, matchedAny: 60, matchedAll: 19, appended: 0 });
    expect(res.rows.map((r) => r.baseRow)).toEqual(Array.from({ length: 60 }, (_, i) => i + 4));
    // no check item points at the footer rows
    for (const it of res.items) for (const r of it.rows ?? []) expect(r).toBeLessThanOrEqual(63);
  });

  it('selected keys per file', () => {
    const sel = (id: string) => an.pairs[id].candidates.filter((c) => c.selected).sort((x, y) => (x.tier ?? 99) - (y.tier ?? 99)).map((c) => c.label);
    expect(sel('IV')).toEqual(['ID＋氏名', 'ID', '氏名']);
    expect(sel('CT')[0]).toBe('氏名＋電話番号');
    expect(sel('AP')).toEqual(['氏名カナ']);
    expect(sel('AT')[0]).toBe('氏名＋大学');
  });

  for (const id of ['IV', 'CT', 'AP', 'AT'] as const) {
    it(`${FILES[id]}: matches, unmatched base rows and other-only rows as in truth.json`, () => {
      const tf = truth[TRUTH_KEY[id]];
      const byBase = new Map(res.rows.map((r) => [r.baseRow, r.matches[id]?.row ?? null]));
      // every match the tool makes is the true one
      for (const [baseRow, row] of byBase) {
        if (row === null) continue;
        const t = tf.rows.find((x) => x.row === row);
        expect(t?.baseRow, `${FILES[id]} row ${row} → base ${baseRow}`).toBe(baseRow);
      }
      const unmatched = res.rows.filter((r) => !r.matches[id]).map((r) => r.baseRow);
      const ambiguousRows = res.items.find((x) => x.code === 'ambiguous' && x.sourceId === id)?.rows ?? [];
      // every truly unmatched row is unmatched; the only extra ones are 同姓同名 rows with identical keys (04: kana only),
      // reported as ambiguous instead of guessed
      for (const r of tf.expected.unmatchedBaseRows) expect(unmatched).toContain(r);
      const extra = unmatched.filter((r) => !tf.expected.unmatchedBaseRows.includes(r as number));
      for (const r of extra) expect(ambiguousRows).toContain(r);
      expect(extra.length).toBe(id === 'AP' ? 1 : 0);
      expect(res.reports[id].matched).toBe(tf.expected.matchedBaseRows - extra.length);
      expect(res.items.find((x) => x.code === 'unmatchedOther' && x.sourceId === id)?.otherRows).toEqual(tf.expected.extraRows);
      // the report lists every row of the other file that joined nothing: other-only rows + candidates held back
      // for an ambiguous base row
      const held = res.reports[id].ambiguous.flatMap((a) => a.candidates.map((c) => c.excelRow));
      expect(res.reports[id].unmatchedOther.map((r) => r.excelRow)).toEqual([...new Set([...tf.expected.extraRows, ...held])].sort((a, b) => a - b));
    });
  }

  it('02: resubmitted rows are dropped by the policy (duplicate item names them); 04: the 同姓同名 pair is ambiguous', () => {
    const dup = res.items.find((x) => x.code === 'duplicate' && x.sourceId === 'IV');
    const resub = truth.interview.rows.filter((r) => r.resubmitted);
    expect(dup?.otherRows).toEqual(resub.map((r) => r.row));
    expect(dup?.rows).toEqual(resub.map((r) => r.baseRow));
    const amb = res.items.find((x) => x.code === 'ambiguous' && x.sourceId === 'AP');
    expect(amb?.rows).toEqual(truth.base.sameNamePair);
    const apRow = truth.aptitude.rows.find((r) => r.baseRow !== null && truth.base.sameNamePair.includes(r.baseRow))?.row;
    expect(amb?.otherRows).toEqual([apRow]);
  });

  it('keyFill and full join: 8 other-only rows appended (02: 3, 03: 2, 04: 2, 05: 1)', () => {
    expect(res.keyFill.IV).toEqual([{ baseCol: 1, sourceCol: 1 }, { baseCol: 2, sourceCol: 2 }]);
    expect(res.keyFill.AP).toEqual([{ baseCol: 3, sourceCol: 2 }]);
    expect(res.keyFill.AT).toEqual([{ baseCol: 2, sourceCol: 1 }, { baseCol: 4, sourceCol: 2 }]);
    const full = joinSources(an, { ...defaultPlan(an), joinType: 'full' });
    expect(full.summary).toEqual({ baseRows: 60, matchedAny: 60, matchedAll: 19, appended: 8 });
    const appended = full.rows.filter((r) => r.baseRow === null).map((r) => `${r.appendedFrom?.sourceId}:${r.appendedFrom?.row}`);
    expect(appended).toEqual(['IV:11', 'IV:30', 'IV:49', 'CT:8', 'CT:28', 'AP:4', 'AP:37', 'AT:20']);
  });
});
