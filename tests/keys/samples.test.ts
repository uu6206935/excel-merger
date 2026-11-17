/**
 * interview-scheduler sample files (copied to tests/fixtures/keys/interview): data-sheet choice, regions and the
 * cross-file results of research R6 §9.2, now through the facade.
 */
import path from 'node:path';
import { describe, expect, it } from 'vitest';
import type { KeySource } from '../../src/core/contracts';
import { analyzeSources, defaultPlan, detectRegion, joinSources, pickDataSheet } from '../../src/core/keys';
import { FIXTURES, bookFromFile, sourceOf } from './helpers';

const dir = path.join(FIXTURES, 'interview');
const book = (f: string, id: string) => bookFromFile(path.join(dir, f), id);
const load = (f: string, id: string, label: string, sheet?: string): KeySource => sourceOf(book(f, id), id, label, sheet);

describe('data sheet and region detection', () => {
  it('picks the list-style 備考 sheet over the 横持ち answer sheet', () => {
    for (const f of ['学生回答.xlsx', '職場回答.xlsx', '人事回答.xlsx']) expect(pickDataSheet(book(f, f))).toBe('備考');
  });
  it('picks 名簿 in the scheduler workbook and the only sheet elsewhere', () => {
    expect(pickDataSheet(book('面接日程_調整用_20260928.xlsx', 'x'))).toBe('名簿');
    expect(pickDataSheet(book('学生-職場マッチング.xlsx', 'x'))).toBe('学生-職場マッチング');
  });
  it('regions', () => {
    expect(detectRegion(book('学生回答.xlsx', 'x'), '備考')).toEqual({ sheet: '備考', headerRow: 1, firstDataRow: 2, lastDataRow: 31, firstCol: 1, lastCol: 3 });
    expect(detectRegion(book('学生回答.xlsx', 'x'), 'Sheet1').headerRow).toBe(2);
    expect(detectRegion(book('学生-職場マッチング.xlsx', 'x'), '学生-職場マッチング')).toEqual({ sheet: '学生-職場マッチング', headerRow: 1, firstDataRow: 2, lastDataRow: 31, firstCol: 1, lastCol: 5 });
    // the old-format CSV has 12 long question headers and many short answer cells: header is row 1, not row 24
    expect(detectRegion(book('学生回答_旧形式_サンプル.csv', 'x'), 'Sheet1')).toMatchObject({ headerRow: 1, firstDataRow: 2, lastDataRow: 31, firstCol: 1 });
  });
  it('the 横持ち sheet carries a layout hint', () => {
    const an = analyzeSources(load('学生-職場マッチング.xlsx', 'M', 'マッチング'), [load('学生回答.xlsx', 'S', '学生回答', 'Sheet1')]);
    expect(an.pairs.S.other.layoutHint).toMatch(/横持ち/);
  });
});

describe('cross-file joins (R6 §9.2)', () => {
  const M = load('学生-職場マッチング.xlsx', 'M', 'マッチング');
  const ST = load('学生回答.xlsx', 'ST', '学生回答');
  const WP = load('職場回答.xlsx', 'WP', '職場回答');
  const HR = load('人事回答.xlsx', 'HR', '人事回答');
  const CSV = load('学生回答_旧形式_サンプル.csv', 'CSV', '旧形式');
  const MEIBO = load('面接日程_調整用_20260928.xlsx', 'MEIBO', '名簿', '名簿');
  const MIHAI = load('面接日程_調整用_20260928.xlsx', 'MIHAI', '未配置', '未配置');

  it('マッチング ⟷ 学生回答 / 職場回答 / 人事回答', () => {
    const an = analyzeSources(M, [ST, WP, HR]);
    // base profile: No = 行番号, 学生ID = 学籍番号 family, 職場ID = 組織コード (not a person key)
    const types = Object.fromEntries(an.base.columns.map((c) => [c.header, c.type]));
    expect(types).toMatchObject({ No: 'rowNumber', 学生ID: 'studentId', 学生名: 'name', 職場ID: 'orgCode' });
    expect(an.pairs.ST.mode).toBe('oneToOne');
    expect(an.pairs.WP.mode).toBe('lookup');
    const selST = an.pairs.ST.candidates.filter((c) => c.selected).map((c) => [c.label, c.family, c.baseColumns[0].header, c.otherColumns[0].header]);
    expect(selST).toEqual([['ID＋氏名', 'id', '学生ID', '応募者コード'], ['ID', 'id', '学生ID', '応募者コード'], ['氏名', 'name', '学生名', '漢字氏名']]);
    const selWP = an.pairs.WP.candidates.filter((c) => c.selected);
    expect(selWP.map((c) => [c.kind, c.family, c.baseColumns[0].header, c.otherColumns[0].header])).toEqual([['discovered', 'id', '職場ID', '応募者コード']]);
    expect(an.pairs.HR.candidates.filter((c) => c.selected)).toEqual([]);
    // chips: ID / 氏名 always present, per source
    expect(an.chips.map((c) => c.family)).toEqual(['id', 'name']);
    expect(an.chips[0].perSource.ST?.label).toBe('ID＋氏名');
    expect(an.chips[0].perSource.HR).toBeNull();

    const res = joinSources(an, defaultPlan(an));
    expect(res.summary).toEqual({ baseRows: 30, matchedAny: 30, matchedAll: 0, appended: 0 });
    expect(res.reports.ST.matched).toBe(30);
    expect(res.reports.ST.perTier['ID＋氏名']).toBe(30);
    expect(res.reports.WP.matched).toBe(30); // every student gets its workplace row (N:1 lookup)
    expect(Object.keys(res.reports.WP.flags)).toEqual([]); // 職場名 ⟷ 漢字氏名 consistency check: 0 disagreements
    expect(res.reports.HR.matched).toBe(0);
    expect(res.items.find((i) => i.code === 'noKey')?.sourceId).toBe('HR');
    expect(res.items.some((i) => i.code === 'duplicateBase')).toBe(false);
    expect(res.items.find((i) => i.code === 'lookup')?.severity).toBe('info');
    // keyFill / imports: key columns duplicating a base key column are not imported
    expect(res.keyFill.ST).toEqual([{ baseCol: 2, sourceCol: 1 }, { baseCol: 3, sourceCol: 2 }]);
    expect(res.keyFill.WP).toEqual([{ baseCol: 4, sourceCol: 1 }]);
    expect(res.defaultImports).toEqual([
      { sourceId: 'ST', col: 3 },
      { sourceId: 'WP', col: 2 },
      { sourceId: 'WP', col: 3 },
      { sourceId: 'HR', col: 1 },
      { sourceId: 'HR', col: 2 },
      { sourceId: 'HR', col: 3 },
    ]);
    // rows: base order, Excel rows
    expect(res.rows.map((r) => r.baseRow)).toEqual(Array.from({ length: 30 }, (_, i) => i + 2));
    expect(res.rows[0].matches.ST).toEqual({ row: 2, via: 'ID＋氏名', flags: [] });
    expect(res.rows[0].matches.WP?.flags).toEqual(['duplicateInBase']);
    expect(res.rows[0].matches.HR).toBeNull();
  });

  it('マッチング ⟷ 旧形式 CSV / 名簿', () => {
    const an = analyzeSources(M, [CSV, MEIBO]);
    const res = joinSources(an, defaultPlan(an));
    expect(res.reports.CSV.matched).toBe(30);
    expect(res.reports.MEIBO.matched).toBe(30);
    expect(res.reports.MEIBO.unmatchedOther.length).toBe(14); // staff / HR rows of 名簿
    expect(res.summary.matchedAll).toBe(30);
  });

  it('人事回答 / 職場回答 ⟷ 名簿 and 学生回答 ⟷ 未配置', () => {
    for (const [base, n] of [[HR, 8], [WP, 6]] as const) {
      const an = analyzeSources(base, [MEIBO]);
      const res = joinSources(an, defaultPlan(an));
      expect(res.reports.MEIBO.matched).toBe(n);
      expect(res.reports.MEIBO.plan.map((p) => p.name)).toEqual(['ID＋氏名', 'ID', '氏名']);
    }
    const an = analyzeSources(ST, [MIHAI]);
    expect(joinSources(an, defaultPlan(an)).reports.MIHAI.matched).toBe(30);
  });

  it('名簿 as base, three others', () => {
    const an = analyzeSources(MEIBO, [M, HR, WP]);
    const res = joinSources(an, { ...defaultPlan(an), joinType: 'full' });
    expect(res.summary).toEqual({ baseRows: 44, matchedAny: 44, matchedAll: 0, appended: 0 });
  });
});
