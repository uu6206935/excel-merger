/**
 * Multi-file join semantics (facade analyzeSources / defaultPlan / joinSources) on hand-built SourceBooks.
 */
import { describe, expect, it } from 'vitest';
import type { JoinPlan, KeyAnalysis, KeySource } from '../../src/core/contracts';
import type { MergedRowPlan } from '../../src/core/types';
import { analyzeSources, defaultPlan, joinSources } from '../../src/core/keys';
import { type CellSpec, sourceFromRows } from './helpers';

const HEAD_BASE = ['応募者ID', '氏名', 'フリガナ', '大学'];
const BASE_ROWS: CellSpec[][] = [
  HEAD_BASE,
  ['1001', '山田 太郎', 'ヤマダ タロウ', '東京大学'],
  ['1002', '鈴木 花子', 'スズキ ハナコ', '慶應義塾大学'],
  ['1003', '佐藤 翔太', 'サトウ ショウタ', '早稲田大学'],
  ['1004', '佐藤 翔太', 'サトウ ショウタ', '慶應義塾大学'],
  ['1005', '髙橋 一郎', 'タカハシ イチロウ', '京都大学'],
  ['1006', '渡邊 恵', 'ワタナベ メグミ', '大阪大学'],
  ['1007', '伊藤 陽菜', 'イトウ ヒナ', '一橋大学'],
  ['1008', '中村 優子', 'ナカムラ ユウコ', '明治大学'],
];
const base = () => sourceFromRows('base', BASE_ROWS, '名簿');

function run(b: KeySource, others: KeySource[], patch: Partial<JoinPlan> = {}, analysis?: KeyAnalysis) {
  const an = analysis ?? analyzeSources(b, others);
  const plan = { ...defaultPlan(an), ...patch };
  return { an, plan, res: joinSources(an, plan) };
}
const byBase = (rows: MergedRowPlan[]) => new Map(rows.filter((r) => r.baseRow !== null).map((r) => [r.baseRow as number, r]));

describe('full join over three files', () => {
  // 面接 (S1): 1001 matched, 2001/2002 new; 内定 (S2): 1002 matched, 2001 (also in S1 → ONE appended row), 3001 new
  const s1 = () => sourceFromRows('S1', [
    ['受験番号', '名前', '面接評価'],
    ['1001', '山田太郎', 'A'],
    ['2001', '田中 一郎', 'B'],
    ['2002', '小林 大輝', 'C'],
    ['1003', '佐藤 翔太', 'A'],
    ['1005', '高橋 一郎', 'B'],
  ]);
  const s2 = () => sourceFromRows('S2', [
    ['ID', '氏名', '結果'],
    ['1002', '鈴木 花子', '内定'],
    ['2001', '田中 一郎', '内定'],
    ['3001', '加藤 結衣', '保留'],
    ['1006', '渡辺 恵', '内定'],
  ]);

  it('appends rows only in other files once, in source order, and matches later files against them', () => {
    const { res } = run(base(), [s1(), s2()], { joinType: 'full' });
    expect(res.summary).toEqual({ baseRows: 8, matchedAny: 5, matchedAll: 0, appended: 3 });
    const appended = res.rows.filter((r) => r.baseRow === null);
    expect(appended.map((r) => r.appendedFrom)).toEqual([
      { sourceId: 'S1', row: 3 },
      { sourceId: 'S1', row: 4 },
      { sourceId: 'S2', row: 4 },
    ]);
    // 田中 一郎 (2001): appended from S1, matched by S2 through the base ID column filled from S1
    expect(appended[0].matches).toEqual({ S1: { row: 3, via: '追加行', flags: [] }, S2: { row: 3, via: 'ID＋氏名', flags: [] } });
    expect(appended[1].matches).toEqual({ S1: { row: 4, via: '追加行', flags: [] }, S2: null });
    expect(appended[2].matches).toEqual({ S1: null, S2: { row: 4, via: '追加行', flags: [] } });
    // key columns of each pair are filled into the base key columns for appended rows
    expect(res.keyFill.S1).toEqual([{ baseCol: 1, sourceCol: 1 }, { baseCol: 2, sourceCol: 2 }]);
    expect(res.keyFill.S2).toEqual([{ baseCol: 1, sourceCol: 1 }, { baseCol: 2, sourceCol: 2 }]);
    expect(res.defaultImports).toEqual([{ sourceId: 'S1', col: 3 }, { sourceId: 'S2', col: 3 }]);
    // base rows in base order first
    expect(res.rows.slice(0, 8).map((r) => r.baseRow)).toEqual([2, 3, 4, 5, 6, 7, 8, 9]);
    const b = byBase(res.rows);
    expect(b.get(2)?.matches.S1).toEqual({ row: 2, via: 'ID＋氏名', flags: [] });
    expect(b.get(6)?.matches.S1?.flags).toEqual(['variant']); // 髙橋 ⟷ 高橋
    expect(b.get(7)?.matches.S2?.flags).toEqual(['variant']); // 渡邊 ⟷ 渡辺
    expect(res.items.filter((i) => i.code === 'unmatchedOther').map((i) => i.severity)).toEqual(['info', 'info']);
  });

  it('left join: nothing appended, other-only rows reported', () => {
    const { res } = run(base(), [s1(), s2()]);
    expect(res.summary.appended).toBe(0);
    expect(res.rows.length).toBe(8);
    expect(res.reports.S1.unmatchedOther.map((r) => r.excelRow)).toEqual([3, 4]);
    const it = res.items.find((i) => i.code === 'unmatchedOther' && i.sourceId === 'S1');
    expect(it?.message).toContain('にだけある行: 2件');
    expect(it?.severity).toBe('warning');
    const un = res.items.find((i) => i.code === 'unmatched' && i.sourceId === 'S1');
    expect(un?.message).toBe('「S1」に見つからない基準の行: 5件');
    expect(un?.rows).toEqual([3, 5, 7, 8, 9]); // 1002, 1004, 1006, 1007, 1008 (base sheet rows)
  });

  it('a person appended from S1 is not appended again when S2 lacks an ID but has the name', () => {
    const s2b = sourceFromRows('S2', [
      ['ID', '氏名', '結果'],
      ['1002', '鈴木 花子', '内定'],
      [null, '田中 一郎', '内定'],
      ['1006', '渡辺 恵', '内定'],
      ['1001', '山田 太郎', '辞退'],
    ]);
    const { res } = run(base(), [s1(), s2b], { joinType: 'full' });
    const appended = res.rows.filter((r) => r.baseRow === null);
    expect(appended.length).toBe(2);
    expect(appended[0].matches.S2).toEqual({ row: 3, via: '氏名', flags: ['fallback'] });
  });
});

describe('manual links', () => {
  const other = () => sourceFromRows('O', [
    ['受験番号', '名前', '評価'],
    ['1001', '山田 太郎', 'A'],
    ['1002', '鈴木 花子', 'B'],
    [null, '伊藤 陽奈', 'C'], // typo of 伊藤 陽菜, no ID → suggestion only
    ['1008', '中村 優子', 'A'],
  ]);

  it('a suggestion accepted as a manual link joins the rows', () => {
    const r0 = run(base(), [other()]);
    const sug = r0.res.reports.O.suggestions.find((s) => s.base.excelRow === 8);
    expect(sug?.other.excelRow).toBe(4);
    expect(byBase(r0.res.rows).get(8)?.matches.O).toBeNull();
    const { res } = run(base(), [other()], { manualLinks: { O: [{ baseRow: 8, otherRow: 4 }] } }, r0.an);
    // still cross-checked: 陽菜 / 陽奈 is a one-character difference (要確認)
    expect(byBase(res.rows).get(8)?.matches.O).toEqual({ row: 4, via: '手動', flags: ['manual', 'conflict'] });
    expect(res.reports.O.flags['氏名1文字違い']?.count).toBe(1);
    expect(res.reports.O.perTier['手動']).toBe(1);
    expect(res.reports.O.suggestions.find((s) => s.base.excelRow === 8)).toBeUndefined();
  });

  it('manual links override tier matches', () => {
    const { res } = run(base(), [other()], { manualLinks: { O: [{ baseRow: 2, otherRow: 3 }] } });
    const b = byBase(res.rows);
    expect(b.get(2)?.matches.O?.row).toBe(3);
    expect(b.get(2)?.matches.O?.flags).toContain('manual');
    expect(b.get(2)?.matches.O?.flags).toContain('conflict'); // 山田 太郎 ⟷ 鈴木 花子 is flagged even when forced
    expect(b.get(3)?.matches.O).toBeNull(); // 鈴木's row was taken by the manual link
  });

  it('invalid manual links are reported, not applied', () => {
    const { res } = run(base(), [other()], { manualLinks: { O: [{ baseRow: 99, otherRow: 2 }, { baseRow: 2, otherRow: 3 }, { baseRow: 2, otherRow: 5 }] } });
    const bad = res.items.filter((i) => i.code === 'manualInvalid');
    expect(bad.length).toBe(2);
    expect(bad[0].message).toContain('基準の行がデータ範囲外');
    expect(bad[1].message).toContain('同じ基準の行に複数の手動結合');
    expect(byBase(res.rows).get(2)?.matches.O?.row).toBe(3);
  });
});

describe('duplicate policy', () => {
  const dup = () => sourceFromRows('D', [
    ['受験番号', '名前', '評価'],
    ['1001', '山田 太郎', 'A'],
    ['1002', '鈴木 花子', 'B'],
    ['1003', '佐藤 翔太', 'C'],
    ['1002', '鈴木 花子', 'A（再提出）'],
    ['1005', '髙橋 一郎', 'B'],
  ]);

  it('first / last pick one of identical submissions', () => {
    const first = run(base(), [dup()]).res;
    expect(byBase(first.rows).get(3)?.matches.D).toEqual({ row: 3, via: 'ID＋氏名', flags: ['duplicateInSource'] });
    const last = run(base(), [dup()], { duplicatePolicy: 'last' }).res;
    expect(byBase(last.rows).get(3)?.matches.D).toEqual({ row: 5, via: 'ID＋氏名', flags: ['duplicateInSource'] });
    const item = last.items.find((i) => i.code === 'duplicate');
    expect(item?.message).toBe('「D」でキーが重複している行: 1件（最後の行を採用）');
    expect(item?.rows).toEqual([3]);
    // the dropped duplicate is neither unmatched nor appended
    expect(last.reports.D.unmatchedOther).toEqual([]);
    const full = run(base(), [dup()], { joinType: 'full' }).res;
    expect(full.summary.appended).toBe(0);
  });

  it('skip leaves the row unjoined and reports it as ambiguous', () => {
    const { res } = run(base(), [dup()], { duplicatePolicy: 'skip' });
    expect(byBase(res.rows).get(3)?.matches.D).toBeNull();
    expect(res.reports.D.ambiguous.map((a) => [a.base.excelRow, a.candidates.map((c) => c.excelRow)])).toEqual([[3, [3, 5]]]);
    expect(res.items.find((i) => i.code === 'ambiguous')?.rows).toEqual([3]);
    // not appended in a full join either (they belong to a base row)
    expect(run(base(), [dup()], { duplicatePolicy: 'skip', joinType: 'full' }).res.summary.appended).toBe(0);
  });

  it('rows sharing an ID but disagreeing on the name are not merged by the policy', () => {
    const bad = sourceFromRows('D', [
      ['受験番号', '名前', '評価'],
      ['1001', '山田 太郎', 'A'],
      ['1002', '中村 優子', 'B'], // wrong ID typed
      ['1002', '鈴木 花子', 'C'],
      ['1003', '佐藤 翔太', 'A'],
    ]);
    const { res } = run(base(), [bad]);
    // ID＋氏名 is unique per row: it picks the right row at once (the wrongly typed 中村 優子 / 1002 row has another first-tier key)
    expect(byBase(res.rows).get(3)?.matches.D).toEqual({ row: 4, via: 'ID＋氏名', flags: [] });
    // 中村 優子 (base 1008) must not receive the row with ID 1002: the ID tier vetoes it
    expect(byBase(res.rows).get(9)?.matches.D).toBeNull();
  });

  it('duplicates among new rows are appended once (first) or all (skip)', () => {
    const n = () => sourceFromRows('N', [
      ['受験番号', '名前', '評価'],
      ['1001', '山田 太郎', 'A'],
      ['2001', '田中 一郎', 'B'],
      ['2001', '田中 一郎', 'B（再提出）'],
      ['1002', '鈴木 花子', 'A'],
    ]);
    const first = run(base(), [n()], { joinType: 'full' }).res;
    expect(first.rows.filter((r) => r.baseRow === null).map((r) => [r.appendedFrom?.row, r.matches.N?.flags])).toEqual([[3, ['duplicateInSource']]]);
    expect(first.items.find((i) => i.code === 'duplicate')?.message).toContain('追加しなかった重複行: 4行目');
    const last = run(base(), [n()], { joinType: 'full', duplicatePolicy: 'last' }).res;
    expect(last.rows.filter((r) => r.baseRow === null).map((r) => r.appendedFrom?.row)).toEqual([4]);
    const skip = run(base(), [n()], { joinType: 'full', duplicatePolicy: 'skip' }).res;
    expect(skip.summary.appended).toBe(2);
  });

  it('base rows sharing an ID all receive the row (flagged)', () => {
    const b2 = sourceFromRows('base', [...BASE_ROWS, ['1002', '鈴木 花子', 'スズキ ハナコ', '慶應義塾大学']], '名簿');
    const o = sourceFromRows('O', [['受験番号', '名前', '評価'], ['1001', '山田 太郎', 'A'], ['1002', '鈴木 花子', 'B'], ['1003', '佐藤 翔太', 'C']]);
    const { res } = run(b2, [o]);
    const b = byBase(res.rows);
    expect(b.get(3)?.matches.O).toEqual({ row: 3, via: 'ID＋氏名', flags: ['duplicateInBase'] });
    expect(b.get(10)?.matches.O).toEqual({ row: 3, via: 'ID＋氏名', flags: ['duplicateInBase'] });
    expect(res.items.find((i) => i.code === 'duplicateBase')?.rows).toEqual([3, 10]);
    // 結合しない: neither base row is joined; both are listed as ambiguous (and as base duplicates)
    const skip = run(b2, [o], { duplicatePolicy: 'skip' }).res;
    const bs = byBase(skip.rows);
    expect([bs.get(3)?.matches.O, bs.get(10)?.matches.O]).toEqual([null, null]);
    expect(skip.reports.O.ambiguous.map((a) => a.base.excelRow)).toEqual([3, 10]);
    expect(skip.items.find((i) => i.code === 'duplicateBase')?.rows).toEqual([3, 10]);
  });
});

describe('key normalisation in joins', () => {
  it('leading zeros: numbers formatted 000000, zero-padded text and plain numbers join', () => {
    const b = sourceFromRows('base', [
      ['社員番号', '氏名', '部署'],
      [{ v: 115, fmt: '000000' }, '山田 太郎', '営業'],
      ['000116', '鈴木 花子', '人事'],
      [117, '佐藤 翔太', '総務'],
      ['000118', '高橋 一郎', '経理'],
    ]);
    const o = sourceFromRows('O', [
      ['社員番号', '氏名', '評価'],
      ['000115', '山田 太郎', 'A'],
      [116, '鈴木 花子', 'B'],
      ['117', '佐藤 翔太', 'C'],
      ['１１８', '高橋 一郎', 'D'], // full-width digits
    ]);
    const { an, res } = run(b, [o]);
    expect(res.reports.O.perTier['ID＋氏名']).toBe(4);
    for (const r of res.rows) expect(r.matches.O?.via).toBe('ID＋氏名');
    expect(res.rows[0].matches.O?.flags).toEqual([]); // 000115 ⟷ 000115 (display text of the number cell)
    expect(res.rows[1].matches.O?.flags).toEqual(['variant']); // 000116 ⟷ 116
    // 先頭ゼロを無視 off: 000116 ≠ 116, and different IDs veto the 氏名 tier (never merged)
    const an2 = analyzeSources(b, [o], { leadingZeros: false });
    expect(an2.options.leadingZeros).toBe(false);
    const res2 = joinSources(an2, defaultPlan(an2));
    expect(res2.rows[0].matches.O?.via).toBe('ID＋氏名'); // the number cell SHOWS 000115
    expect(res2.rows[1].matches.O).toBeNull();
    expect(res2.rows[2].matches.O?.via).toBe('ID＋氏名'); // 117 ⟷ '117'
    expect(res2.items.find((i) => i.code === 'vetoed')?.rows).toEqual([3, 5]);
    expect(an.options.leadingZeros).toBe(true);
  });

  it('full-width / hyphen variants of alphanumeric IDs', () => {
    const b = sourceFromRows('base', [['会員番号', '氏名'], ['AB-0012', '山田 太郎'], ['AB-0013', '鈴木 花子'], ['AB-0014', '佐藤 翔太']]);
    const o = sourceFromRows('O', [['会員番号', '購入額'], ['ＡＢ－００１２', 100], ['ab0013', 200], ['AB‐0014', 300]]);
    const { res } = run(b, [o]);
    expect(res.rows.map((r) => r.matches.O?.row)).toEqual([2, 3, 4]);
  });

  it('異体字 in names (name-only files) and the itaiji option', () => {
    const b = sourceFromRows('base', [['氏名', '部署'], ['髙橋 一郎', '営業'], ['渡邊 恵', '人事'], ['山﨑 翔', '総務'], ['齋藤 美咲', '経理']]);
    const o = sourceFromRows('O', [['お名前', '回答'], ['高橋一郎', 'はい'], ['渡辺 恵', 'いいえ'], ['山崎　翔', 'はい'], ['斎藤 美咲', 'はい']]);
    const { res } = run(b, [o]);
    expect(res.rows.map((r) => [r.matches.O?.row, r.matches.O?.flags])).toEqual([[2, ['variant']], [3, ['variant']], [4, ['variant']], [5, ['variant']]]);
    expect(res.items.find((i) => i.code === 'variant')?.message).toContain('4件');
    const an2 = analyzeSources(b, [o], { itaiji: false });
    const res2 = joinSources(an2, defaultPlan(an2));
    // 﨑 (U+FA11) has no NFKC decomposition: without the 異体字 table nothing joins, all are suggested instead
    expect(res2.rows.map((r) => r.matches.O?.row ?? null)).toEqual([null, null, null, null]);
    expect(res2.reports.O.suggestions.map((s) => [s.base.excelRow, s.other.excelRow])).toEqual(expect.arrayContaining([[2, 2], [3, 3], [4, 4], [5, 5]]));
  });

  it('斉藤 ⟷ 斎藤 are different surnames: suggested, never auto-merged', () => {
    const b = sourceFromRows('base', [['氏名', 'フリガナ'], ['斉藤 健太', 'サイトウ ケンタ'], ['山田 太郎', 'ヤマダ タロウ'], ['鈴木 花子', 'スズキ ハナコ']]);
    const o = sourceFromRows('O', [['氏名', 'ふりがな', '評価'], ['斎藤 健太', 'さいとう けんた', 'A'], ['山田 太郎', 'やまだ たろう', 'B'], ['鈴木 花子', 'すずき はなこ', 'C']]);
    const { res } = run(b, [o]);
    expect(res.rows[0].matches.O).toBeNull();
    expect(res.reports.O.suggestions[0]).toMatchObject({ base: { excelRow: 2 }, other: { excelRow: 2 } });
    expect(res.reports.O.conflicts.length).toBe(1); // カナ tier vetoed by the kanji name tier
    const sug = res.items.find((i) => i.code === 'suggestion');
    expect(sug?.rows).toEqual([2]);
    expect(sug?.otherRows).toEqual([2]);
  });
});

describe('fallbacks and composites', () => {
  it('missing IDs in one file fall back to 氏名＋大学 (同姓同名 in the base)', () => {
    const o = sourceFromRows('O', [
      ['受験番号', '名前', '学校名', '評価'],
      ['1001', '山田 太郎', '東京大学', 'A'],
      [null, '佐藤 翔太', '慶応義塾大学', 'B'], // 1004 (慶應), no ID
      [null, '佐藤 翔太', '早稲田大学', 'C'], // 1003 (早稲田), no ID
      ['1002', '鈴木 花子', '慶應', 'A'],
      [null, '伊藤 陽菜', '一橋大学', 'B'],
    ]);
    const { an, res } = run(base(), [o]);
    expect(an.pairs.O.candidates.filter((c) => c.selected).map((c) => c.label)).toEqual(['ID＋氏名', 'ID', '氏名＋大学', '氏名']);
    const b = byBase(res.rows);
    expect(b.get(4)?.matches.O).toEqual({ row: 4, via: '氏名＋大学', flags: ['fallback'] });
    expect(b.get(5)?.matches.O).toEqual({ row: 3, via: '氏名＋大学', flags: ['fallback'] });
    expect(b.get(8)?.matches.O?.via).toBe('氏名＋大学');
    expect(b.get(3)?.matches.O?.via).toBe('ID＋氏名');
    const fb = res.items.find((i) => i.code === 'fallback');
    expect(fb?.severity).toBe('warning');
    expect(fb?.rows).toEqual([4, 5, 8]);
    // keyFill: ID and both composite parts
    expect(res.keyFill.O).toEqual([{ baseCol: 1, sourceCol: 1 }, { baseCol: 2, sourceCol: 2 }, { baseCol: 4, sourceCol: 3 }]);
    expect(res.defaultImports).toEqual([{ sourceId: 'O', col: 4 }]);
  });

  it('同姓同名 without a distinguishing column is never auto-merged', () => {
    const o = sourceFromRows('O', [
      ['名前', '評価'],
      ['佐藤 翔太', 'A'],
      ['山田 太郎', 'B'],
      ['鈴木 花子', 'C'],
      ['中村 優子', 'D'],
    ]);
    const { res } = run(base(), [o]);
    const b = byBase(res.rows);
    expect(b.get(4)?.matches.O).toBeNull();
    expect(b.get(5)?.matches.O).toBeNull();
    expect(b.get(2)?.matches.O?.row).toBe(3);
    expect(res.reports.O.ambiguous.map((a) => a.base.excelRow)).toEqual([4, 5]);
    expect(res.items.find((i) => i.code === 'duplicateBase')?.rows).toEqual([4, 5]);
    // new row? no: the 佐藤 row belongs to a base row (ambiguous), it is not appended
    expect(run(base(), [o], { joinType: 'full' }).res.summary.appended).toBe(0);
  });

  it('split 姓 / 名 columns join a 氏名 column and fill it (姓＋全角スペース＋名); they are not imported again', () => {
    const o = sourceFromRows('O', [
      ['姓', '名', '大学', '学年'],
      ['山田', '太郎', '東京大学', 4],
      ['佐藤', '翔太', '早稲田大学', 3],
      ['佐藤', '翔太', '慶應義塾大学', 4],
      ['鈴木', '花子', '慶應義塾大学', 2],
    ]);
    const { res } = run(base(), [o]);
    const b = byBase(res.rows);
    expect(b.get(4)?.matches.O?.row).toBe(3);
    expect(b.get(5)?.matches.O?.row).toBe(4);
    // tier 氏名＋大学: base 氏名 ⟷ 姓＋名 (virtual) and 大学 ⟷ 大学
    expect(res.keyFill.O).toEqual([{ baseCol: 2, sourceCols: [1, 2], joiner: '　' }, { baseCol: 4, sourceCol: 3 }]);
    expect(res.defaultImports.map((c) => c.col)).toEqual([4]);
  });
});

describe('keyFill between 氏名 and 姓 / 名', () => {
  const NAME_BASE: CellSpec[][] = [
    ['氏名', 'メール', '部署'],
    ['山田 太郎', 'yamada@example.jp', '営業'],
    ['鈴木 花子', 'suzuki@example.jp', '人事'],
    ['佐藤 翔太', 'sato@example.jp', '開発'],
    ['伊藤 陽菜', 'ito@example.jp', '総務'],
  ];
  const SPLIT_BASE: CellSpec[][] = [
    ['姓', '名', '部署'],
    ['山田', '太郎', '営業'],
    ['鈴木', '花子', '人事'],
    ['佐藤', '翔太', '開発'],
    ['伊藤', '陽菜', '総務'],
  ];

  it('base 氏名 ⟷ other 姓 + 名: one entry combining both columns; appended rows are matched by later files through it', () => {
    const s1 = sourceFromRows('S1', [
      ['姓', '名', '評価'],
      ['山田', '太郎', 'A'],
      ['田中', '一郎', 'B'], // new person
      ['鈴木', '花子', 'C'],
    ]);
    const s2 = sourceFromRows('S2', [
      ['氏名', '結果'],
      ['田中　一郎', '内定'], // the person appended from S1
      ['佐藤 翔太', '保留'],
    ]);
    const { res } = run(sourceFromRows('base', NAME_BASE), [s1, s2], { joinType: 'full' });
    expect(res.keyFill.S1).toEqual([{ baseCol: 1, sourceCols: [1, 2], joiner: '　' }]);
    expect(res.keyFill.S2).toEqual([{ baseCol: 1, sourceCol: 1 }]);
    expect(res.defaultImports).toEqual([{ sourceId: 'S1', col: 3 }, { sourceId: 'S2', col: 2 }]);
    const appended = res.rows.filter((r) => r.baseRow === null);
    expect(appended.map((r) => r.appendedFrom)).toEqual([{ sourceId: 'S1', row: 3 }]);
    // 田中 一郎 of S2 finds the row appended from S1 (base 氏名 filled from 姓＋名), no second appended row
    expect(appended[0].matches.S2).toEqual({ row: 2, via: '氏名', flags: [] });
    expect(res.summary.appended).toBe(1);
  });

  it('base 姓 / 名 ⟷ other 氏名: two entries splitting the full name (part 0 / 1)', () => {
    const s1 = sourceFromRows('S1', [
      ['氏名', '評価'],
      ['山田　太郎', 'A'],
      ['田中 一郎', 'B'], // new person
      ['伊藤陽菜', 'C'],
    ]);
    const s2 = sourceFromRows('S2', [
      ['姓', '名', '結果'],
      ['田中', '一郎', '内定'], // the person appended from S1
      ['鈴木', '花子', '保留'],
    ]);
    const { res } = run(sourceFromRows('base', SPLIT_BASE), [s1, s2], { joinType: 'full' });
    expect(res.keyFill.S1).toEqual([{ baseCol: 1, sourceCol: 1, splitPart: 0 }, { baseCol: 2, sourceCol: 1, splitPart: 1 }]);
    expect(res.keyFill.S2).toEqual([{ baseCol: 1, sourceCol: 1 }, { baseCol: 2, sourceCol: 2 }]);
    expect(res.defaultImports).toEqual([{ sourceId: 'S1', col: 2 }, { sourceId: 'S2', col: 3 }]);
    const b = byBase(res.rows);
    expect(b.get(2)?.matches.S1?.row).toBe(2);
    expect(b.get(5)?.matches.S1?.row).toBe(4); // 伊藤陽菜 without a space still matches 伊藤 + 陽菜
    const appended = res.rows.filter((r) => r.baseRow === null);
    expect(appended.map((r) => r.appendedFrom)).toEqual([{ sourceId: 'S1', row: 3 }]);
    // S2's 田中 / 一郎 finds the appended row through base 姓 / 名 filled by splitting 「田中 一郎」
    expect(appended[0].matches.S2?.row).toBe(2);
    expect(res.summary.appended).toBe(1);
  });

  it('both sides split: column by column', () => {
    const o = sourceFromRows('O', [['姓', '名', '評価'], ['山田', '太郎', 'A'], ['鈴木', '花子', 'B']]);
    const { res } = run(sourceFromRows('base', SPLIT_BASE), [o]);
    expect(res.keyFill.O).toEqual([{ baseCol: 1, sourceCol: 1 }, { baseCol: 2, sourceCol: 2 }]);
  });
});

describe('CheckItem.otherRows (rows of the other file)', () => {
  it('unmatchedOther lists the other-only rows', () => {
    const o = sourceFromRows('O', [['受験番号', '名前', '評価'], ['1001', '山田 太郎', 'A'], ['2001', '田中 一郎', 'B'], ['2002', '小林 大輝', 'C']]);
    for (const joinType of ['left', 'full'] as const) {
      const { res } = run(base(), [o], { joinType });
      const it = res.items.find((i) => i.code === 'unmatchedOther');
      expect(it?.otherRows).toEqual([3, 4]);
      expect(it?.rows).toBeUndefined();
    }
  });

  it('duplicate lists the rows dropped by the policy; ambiguous lists the candidates', () => {
    const dup = sourceFromRows('D', [
      ['受験番号', '名前', '評価'],
      ['1001', '山田 太郎', 'A'],
      ['1002', '鈴木 花子', 'B'],
      ['1002', '鈴木 花子', 'A（再提出）'],
      ['2001', '田中 一郎', 'B'],
      ['2001', '田中 一郎', 'B（再提出）'],
    ]);
    const first = run(base(), [dup], { joinType: 'full' }).res.items.find((i) => i.code === 'duplicate');
    expect(first?.rows).toEqual([3]);
    expect(first?.otherRows).toEqual([4, 6]); // 1002 row 4 dropped (row 3 kept), new 2001 row 6 not appended
    const last = run(base(), [dup], { duplicatePolicy: 'last' }).res.items.find((i) => i.code === 'duplicate');
    expect(last?.otherRows).toEqual([3]);
    const skip = run(base(), [dup], { duplicatePolicy: 'skip' }).res.items.find((i) => i.code === 'ambiguous');
    expect(skip?.rows).toEqual([3]);
    expect(skip?.otherRows).toEqual([3, 4]);
  });

  it('conflicts name the other file rows too', () => {
    const o = sourceFromRows('O', [['受験番号', '名前', '評価'], ['1001', '中村 恵子', 'A'], ['1002', '鈴木 花子', 'B']]);
    // ID agrees, name does not: the default key ID＋氏名 differs, so the row is not joined (vetoed) instead of being joined with an error flag
    const it = run(base(), [o]).res.items.find((i) => i.code === 'vetoed');
    expect(it?.rows).toEqual([2]);
    expect(it?.otherRows).toEqual([2]);
  });
});

describe('robustness', () => {
  it('no other files', () => {
    const an = analyzeSources(base(), []);
    expect(an.chips.map((c) => c.perSource)).toEqual([{}, {}]);
    const res = joinSources(an, defaultPlan(an));
    expect(res.summary).toEqual({ baseRows: 8, matchedAny: 0, matchedAll: 0, appended: 0 });
    expect(res.rows.length).toBe(8);
  });

  it('empty / missing sheets and inverted regions degrade to empty tables', () => {
    const empty = sourceFromRows('E', [['受験番号', '名前']]);
    const missing: KeySource = { ...sourceFromRows('M', [['a', 'b'], [1, 2]]), region: { sheet: 'nope', headerRow: 1, firstDataRow: 2, lastDataRow: 9, firstCol: 1, lastCol: 2 } };
    const inverted: KeySource = { ...sourceFromRows('I', [['ID', '氏名'], ['1001', '山田 太郎']]), region: { sheet: 'Sheet1', headerRow: 5, firstDataRow: 3, lastDataRow: 1, firstCol: 4, lastCol: 2 } };
    const { res } = run(base(), [empty, missing, inverted], { joinType: 'full' });
    expect(res.summary.appended).toBe(0);
    expect(res.items.filter((i) => i.code === 'emptySource').map((i) => i.sourceId)).toEqual(['E', 'M', 'I']);
    const emptyBase = sourceFromRows('B', []);
    const r2 = run(emptyBase, [base()]);
    expect(r2.res.rows).toEqual([]);
    expect(r2.res.items[0]).toMatchObject({ severity: 'error', code: 'emptySource' });
  });

  it('unknown candidate ids and missing tiers are tolerated', () => {
    const o = sourceFromRows('O', [['受験番号', '評価'], ['1001', 'A'], ['1002', 'B'], ['1003', 'C']]);
    const r1 = run(base(), [o], { tiers: { O: ['single:9|9', 'bogus'] } });
    expect(r1.res.items.map((i) => i.code)).toEqual(expect.arrayContaining(['planInvalid', 'noKey']));
    expect(r1.res.rows.every((r) => r.matches.O === null)).toBe(true);
    const r2 = run(base(), [o], { tiers: {} });
    expect(r2.res.items.some((i) => i.code === 'noKey')).toBe(true);
    // the same other file twice / the base passed as an "other" are ignored
    const an = analyzeSources(base(), [o, o, base()]);
    expect(Object.keys(an.pairs)).toEqual(['O']);
  });

  it('joinSources without the internal state throws a Japanese error', () => {
    const { an, plan } = run(base(), []);
    expect(() => joinSources({ ...an, internal: undefined }, plan)).toThrow(/キーの分析結果がありません/);
  });

  it('formula key cells without a cached value are blank and reported', () => {
    const b = sourceFromRows('base', [
      ['ID', '氏名'],
      ['1001', '山田 太郎'],
      [{ v: null, f: 'A2+1' }, '鈴木 花子'],
      ['1003', '佐藤 翔太'],
    ]);
    const o = sourceFromRows('O', [['ID', '氏名', '評価'], ['1001', '山田 太郎', 'A'], ['1002', '鈴木 花子', 'B'], ['1003', '佐藤 翔太', 'C']]);
    const { res } = run(b, [o]);
    expect(res.rows[1].matches.O?.via).toBe('氏名');
    expect(res.items.find((i) => i.code === 'noCachedValue')?.message).toContain('計算結果が保存されていない数式が 1 個');
  });

  it('RowRef previews are display texts', () => {
    const b = sourceFromRows('base', [['社員番号', '氏名', '入社日'], [{ v: 115, fmt: '000000' }, '山田 太郎', new Date(Date.UTC(2020, 3, 1))], [116, '鈴木 花子', null], [117, '佐藤 翔太', null]]);
    const o = sourceFromRows('O', [['社員番号', '評価'], ['999', 'A'], ['116', 'B'], ['117', 'C']]);
    const { res } = run(b, [o]);
    expect(res.reports.O.unmatchedBase[0]).toEqual({ excelRow: 2, preview: ['000115', '山田 太郎', '2020/4/1'] });
    expect(res.reports.O.unmatchedOther[0]).toEqual({ excelRow: 2, preview: ['999', 'A'] });
  });
});
