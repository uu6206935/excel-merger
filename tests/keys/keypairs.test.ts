/**
 * JoinResult.keyPairs: per source the aligned (base column, source column) pairs of every enabled key tier that matched a
 * row — produced for every join (also 'left'), independent of keyFill. Columns are 1-based sheet columns.
 */
import { describe, expect, it } from 'vitest';
import type { JoinPlan, KeyAnalysis, KeySource } from '../../src/core/contracts';
import { addManualCandidate, analyzeSources, defaultPlan, joinSources } from '../../src/core/keys';
import { type CellSpec, sourceFromRows } from './helpers';

const BASE_ROWS: CellSpec[][] = [
  ['応募者ID', '氏名', 'フリガナ', '大学', '部署'],
  ['1001', '山田 太郎', 'ヤマダ タロウ', '東京大学', '営業'],
  ['1002', '鈴木 花子', 'スズキ ハナコ', '慶應義塾大学', '開発'],
  ['1003', '佐藤 翔太', 'サトウ ショウタ', '早稲田大学', '営業'],
  ['1004', '佐藤 翔太', 'サトウ ショウタ', '慶應義塾大学', '人事'],
  ['1005', '髙橋 一郎', 'タカハシ イチロウ', '京都大学', '開発'],
  ['1006', '渡邊 恵', 'ワタナベ メグミ', '大阪大学', '人事'],
  ['1007', '伊藤 陽菜', 'イトウ ヒナ', '一橋大学', '営業'],
  ['1008', '中村 優子', 'ナカムラ ユウコ', '明治大学', '開発'],
];
const base = () => sourceFromRows('base', BASE_ROWS, '名簿');

function run(b: KeySource, others: KeySource[], patch: Partial<JoinPlan> = {}, analysis?: KeyAnalysis) {
  const an = analysis ?? analyzeSources(b, others);
  const plan: JoinPlan = { ...defaultPlan(an), ...patch };
  return { an, plan, res: joinSources(an, plan) };
}

/** candidate id of the pair with this label */
function candId(an: KeyAnalysis, sourceId: string, label: string): string {
  const c = an.pairs[sourceId].candidates.find((x) => x.label === label);
  if (!c) throw new Error(`no candidate "${label}": ${an.pairs[sourceId].candidates.map((x) => x.label).join(' / ')}`);
  return c.id;
}

describe('keyPairs: single columns', () => {
  const s1 = () => sourceFromRows('S1', [
    ['受験番号', '名前', '面接評価'],
    ['1001', '山田太郎', 'A'],
    ['1002', '鈴木 花子', 'B'],
    ['1003', '佐藤 翔太', 'C'],
  ]);

  it('ID + 氏名: both pairs, although the ID tier took every row and the 氏名 tier matched none itself', () => {
    const { plan, res } = run(base(), [s1()]);
    // default key = ID＋氏名 together; the ID and 氏名 tiers behind it only catch rows where one of the two is blank
    expect(plan.tiers.S1.length).toBe(3);
    expect(res.reports.S1.perTier).toEqual({ 'ID＋氏名': 3, ID: 0, 氏名: 0 });
    expect(res.keyPairs).toEqual({ S1: [{ baseCol: 1, sourceCol: 1 }, { baseCol: 2, sourceCol: 2 }] });
  });

  it("joinType 'left' and 'full' give the same pairs; keyPairs does not depend on keyFill", () => {
    const left = run(base(), [s1()]).res;
    const full = run(base(), [s1()], { joinType: 'full' }).res;
    expect(full.keyPairs).toEqual(left.keyPairs);
    // ... and it is produced when keyFill is not what produced it: only the 氏名 tier enabled
    const only = run(base(), [s1()], { tiers: { S1: [candId(analyzeSources(base(), [s1()]), 'S1', '氏名')] } });
    expect(only.res.keyPairs).toEqual({ S1: [{ baseCol: 2, sourceCol: 2 }] });
  });

  it('only the enabled tiers count: a disabled 氏名 tier gives no pair', () => {
    const an = analyzeSources(base(), [s1()]);
    const { res } = run(base(), [s1()], { tiers: { S1: [candId(an, 'S1', 'ID')] } }, an);
    expect(res.keyPairs).toEqual({ S1: [{ baseCol: 1, sourceCol: 1 }] });
  });

  it('the pairs are listed in tier priority order', () => {
    const an = analyzeSources(base(), [s1()]);
    const ids = [candId(an, 'S1', '氏名'), candId(an, 'S1', 'ID')];
    const { res } = run(base(), [s1()], { tiers: { S1: ids } }, an);
    expect(res.keyPairs?.S1).toEqual([{ baseCol: 2, sourceCol: 2 }, { baseCol: 1, sourceCol: 1 }]);
  });

  it('is plain JSON data and does not change the other results', () => {
    const { res } = run(base(), [s1()]);
    expect(JSON.parse(JSON.stringify(res.keyPairs))).toEqual(res.keyPairs);
    expect(res.keyFill.S1).toEqual([{ baseCol: 1, sourceCol: 1 }, { baseCol: 2, sourceCol: 2 }]);
    expect(res.defaultImports).toEqual([{ sourceId: 'S1', col: 3 }]);
  });
});

describe('keyPairs: composite tiers', () => {
  // 佐藤 翔太 exists twice in the base (different 大学): the name alone is not enough, 氏名＋大学 is
  const s1 = () => sourceFromRows('S1', [
    ['氏名', '大学', '評価'],
    ['山田 太郎', '東京大学', 'A'],
    ['佐藤 翔太', '早稲田大学', 'B'],
    ['佐藤 翔太', '慶應義塾大学', 'C'],
    ['鈴木 花子', '慶應義塾大学', 'D'],
  ]);

  it('氏名＋大学 gives one pair per part; 氏名 alone adds nothing new', () => {
    const { plan, res } = run(base(), [s1()]);
    expect(plan.tiers.S1[0]).toMatch(/^composite:/);
    expect(res.keyPairs).toEqual({ S1: [{ baseCol: 2, sourceCol: 1 }, { baseCol: 4, sourceCol: 2 }] });
  });

  it('a composite tier alone', () => {
    const an = analyzeSources(base(), [s1()]);
    const { res } = run(base(), [s1()], { tiers: { S1: [candId(an, 'S1', '氏名＋大学')] } }, an);
    expect(res.keyPairs?.S1).toEqual([{ baseCol: 2, sourceCol: 1 }, { baseCol: 4, sourceCol: 2 }]);
  });

  it('ID + composite: the ID pair first, then the parts', () => {
    const s = sourceFromRows('S1', [
      ['受験番号', '氏名', '大学', '評価'],
      ['1001', '山田 太郎', '東京大学', 'A'],
      ['1003', '佐藤 翔太', '早稲田大学', 'B'],
      ['1004', '佐藤 翔太', '慶應義塾大学', 'C'],
    ]);
    const an = analyzeSources(base(), [s]);
    const { res } = run(base(), [s], { tiers: { S1: [candId(an, 'S1', 'ID'), candId(an, 'S1', '氏名＋大学')] } }, an);
    expect(res.keyPairs?.S1).toEqual([{ baseCol: 1, sourceCol: 1 }, { baseCol: 2, sourceCol: 2 }, { baseCol: 4, sourceCol: 3 }]);
  });
});

describe('keyPairs: virtual 姓＋名 columns give no pair', () => {
  it('base 氏名 ⟷ other 姓＋名: no entry at all', () => {
    const s = sourceFromRows('S1', [
      ['姓', '名', '評価'],
      ['山田', '太郎', 'A'],
      ['鈴木', '花子', 'B'],
      ['伊藤', '陽菜', 'C'],
    ]);
    const { res } = run(base(), [s]);
    expect(res.summary.matchedAny).toBe(3);
    expect(res.keyFill.S1[0].sourceCols).toEqual([1, 2]);
    expect(res.keyPairs).toEqual({});
  });

  it('ID + 姓／名: only the ID pair', () => {
    const s = sourceFromRows('S1', [
      ['受験番号', '姓', '名', '評価'],
      ['1001', '山田', '太郎', 'A'],
      ['1002', '鈴木', '花子', 'B'],
      ['1007', '伊藤', '陽菜', 'C'],
    ]);
    const { res } = run(base(), [s]);
    expect(res.keyPairs).toEqual({ S1: [{ baseCol: 1, sourceCol: 1 }] });
  });

  it('base 姓／名 ⟷ other 氏名: the base side is virtual, no pair but the ID', () => {
    const b = sourceFromRows('base', [
      ['応募者ID', '姓', '名', '大学'],
      ['1001', '山田', '太郎', '東京大学'],
      ['1002', '鈴木', '花子', '慶應義塾大学'],
      ['1003', '佐藤', '翔太', '早稲田大学'],
    ]);
    const s = sourceFromRows('S1', [
      ['受験番号', '氏名', '評価'],
      ['1001', '山田 太郎', 'A'],
      ['1002', '鈴木 花子', 'B'],
    ]);
    const { res } = run(b, [s]);
    expect(res.summary.matchedAny).toBe(2);
    expect(res.keyPairs).toEqual({ S1: [{ baseCol: 1, sourceCol: 1 }] });
    const noId = run(b, [sourceFromRows('S1', [['氏名', '評価'], ['山田 太郎', 'A'], ['鈴木 花子', 'B']])]).res;
    expect(noId.summary.matchedAny).toBe(2);
    expect(noId.keyPairs).toEqual({});
  });

  it('姓／名 on both sides: the two column pairs are direct pairs', () => {
    const b = sourceFromRows('base', [
      ['応募者ID', '姓', '名', '大学'],
      ['1001', '山田', '太郎', '東京大学'],
      ['1002', '鈴木', '花子', '慶應義塾大学'],
      ['1003', '佐藤', '翔太', '早稲田大学'],
    ]);
    const s = sourceFromRows('S1', [
      ['姓', '名', '評価'],
      ['山田', '太郎', 'A'],
      ['鈴木', '花子', 'B'],
    ]);
    const { res } = run(b, [s]);
    expect(res.summary.matchedAny).toBe(2);
    expect(res.keyPairs).toEqual({ S1: [{ baseCol: 2, sourceCol: 1 }, { baseCol: 3, sourceCol: 2 }] });
  });
});

describe('keyPairs: ID variants pair the same columns', () => {
  it('ID（英字の接頭辞を除く）', () => {
    const s = sourceFromRows('S1', [
      ['受験番号', '評価'],
      ['S1001', 'A'],
      ['S1002', 'B'],
      ['S1003', 'C'],
      ['S1005', 'D'],
    ]);
    const an = analyzeSources(base(), [s]);
    expect(an.pairs.S1.candidates.some((c) => c.label === 'ID（英字の接頭辞を除く）')).toBe(true);
    const { plan, res } = run(base(), [s]);
    expect(plan.tiers.S1).toEqual([candId(an, 'S1', 'ID（英字の接頭辞を除く）')]);
    expect(res.summary.matchedAny).toBe(4);
    expect(res.keyPairs).toEqual({ S1: [{ baseCol: 1, sourceCol: 1 }] });
  });

  it('ID（桁数違い・末尾一致）', () => {
    const b = sourceFromRows('base', [
      ['応募者ID', '氏名'],
      ['20230001', '山田 太郎'],
      ['20230002', '鈴木 花子'],
      ['20230003', '佐藤 翔太'],
      ['20230004', '高橋 一郎'],
    ]);
    const s = sourceFromRows('S1', [
      ['受験番号', '評価'],
      ['230001', 'A'],
      ['230002', 'B'],
      ['230003', 'C'],
      ['230009', 'D'],
    ]);
    const an = analyzeSources(b, [s]);
    const c = an.pairs.S1.candidates.find((x) => x.label === 'ID（桁数違い・末尾一致）');
    expect(c).toBeDefined();
    const { res } = run(b, [s], { tiers: { S1: [c!.id] } }, an);
    expect(res.summary.matchedAny).toBe(3);
    expect(res.keyPairs).toEqual({ S1: [{ baseCol: 1, sourceCol: 1 }] });
  });

  it('plain ID and the variant over the same columns are one pair', () => {
    const s = sourceFromRows('S1', [
      ['受験番号', '評価'],
      ['1001', 'A'],
      ['S1002', 'B'],
      ['S1003', 'C'],
    ]);
    const an = analyzeSources(base(), [s]);
    const ids = an.pairs.S1.candidates.filter((c) => c.family === 'id').map((c) => c.id);
    expect(ids.length).toBeGreaterThanOrEqual(1);
    const { res } = run(base(), [s], { tiers: { S1: ids } }, an);
    expect(res.keyPairs?.S1).toEqual([{ baseCol: 1, sourceCol: 1 }]);
  });
});

describe('keyPairs: manual candidates', () => {
  const s1 = () => sourceFromRows('S1', [
    ['受験番号', '所属', '評価'],
    ['1001', '営業', 'A'],
    ['1002', '開発', 'B'],
    ['1003', '営業', 'C'],
    ['1004', '人事', 'D'],
  ]);

  it('a manual column pair enabled next to the ID tier', () => {
    const an0 = analyzeSources(base(), [s1()]);
    const an = addManualCandidate(an0, 'S1', [5], [2]);
    const manual = an.pairs.S1.candidates[an.pairs.S1.candidates.length - 1];
    expect(manual.id).toBe('manual:5|2');
    const { res } = run(base(), [s1()], { tiers: { S1: [candId(an, 'S1', 'ID'), manual.id] } }, an);
    expect(res.keyPairs?.S1).toEqual([{ baseCol: 1, sourceCol: 1 }, { baseCol: 5, sourceCol: 2 }]);
  });

  it('a manual composite gives one pair per part', () => {
    const an = addManualCandidate(analyzeSources(base(), [s1()]), 'S1', [1, 5], [1, 2]);
    const manual = an.pairs.S1.candidates[an.pairs.S1.candidates.length - 1];
    expect(manual.kind).toBe('composite');
    const { res } = run(base(), [s1()], { tiers: { S1: [manual.id] } }, an);
    expect(res.keyPairs?.S1).toEqual([{ baseCol: 1, sourceCol: 1 }, { baseCol: 5, sourceCol: 2 }]);
  });

  it('a manual pair that never agrees on a matched row gives no pair', () => {
    // 部署 ⟷ 評価 (A / B / C / D): not one equal value
    const an = addManualCandidate(analyzeSources(base(), [s1()]), 'S1', [5], [3]);
    const manual = an.pairs.S1.candidates[an.pairs.S1.candidates.length - 1];
    const { res } = run(base(), [s1()], { tiers: { S1: [candId(an, 'S1', 'ID'), manual.id] } }, an);
    expect(res.summary.matchedAny).toBe(4);
    expect(res.keyPairs?.S1).toEqual([{ baseCol: 1, sourceCol: 1 }]);
  });

  it('manual links alone (no tier) give no pair', () => {
    const { res } = run(base(), [s1()], { tiers: { S1: [] }, manualLinks: { S1: [{ baseRow: 2, otherRow: 2 }] } });
    expect(res.summary.matchedAny).toBe(1);
    expect(res.keyPairs).toEqual({});
  });
});

describe('keyPairs: sources without a usable key', () => {
  const good = () => sourceFromRows('S1', [
    ['受験番号', '名前', '面接評価'],
    ['1001', '山田太郎', 'A'],
    ['1002', '鈴木 花子', 'B'],
  ]);
  const keyless = () => sourceFromRows('S2', [
    ['商品', '数量'],
    ['りんご', 3],
    ['みかん', 5],
  ]);

  it('an excluded file (empty tiers) has no entry; the others keep theirs', () => {
    const { plan, res } = run(base(), [good(), keyless()]);
    expect(plan.tiers.S2).toEqual([]);
    expect(res.keyPairs).toEqual({ S1: [{ baseCol: 1, sourceCol: 1 }, { baseCol: 2, sourceCol: 2 }] });
    expect('S2' in (res.keyPairs ?? {})).toBe(false);
  });

  it('no tier at all: an empty object, never undefined', () => {
    const { res } = run(base(), [keyless()]);
    expect(res.keyPairs).toEqual({});
  });

  it('plan.tiers with unknown candidate ids are ignored', () => {
    const { res } = run(base(), [good()], { tiers: { S1: ['nope', 'single:1|1'] } });
    expect(res.keyPairs).toEqual({ S1: [{ baseCol: 1, sourceCol: 1 }] });
  });

  it('a tier that matched no row gives no pair', () => {
    const s = sourceFromRows('S1', [
      ['受験番号', '名前', '面接評価'],
      ['9001', '田中 一郎', 'A'],
      ['9002', '小林 大輝', 'B'],
    ]);
    // the auto analysis offers no candidate for these files; the user forces ID ⟷ 受験番号
    const an = addManualCandidate(analyzeSources(base(), [s]), 'S1', [1], [1]);
    const manual = an.pairs.S1.candidates[an.pairs.S1.candidates.length - 1];
    const { res } = run(base(), [s], { tiers: { S1: [manual.id] } }, an);
    expect(res.summary.matchedAny).toBe(0);
    expect(res.keyPairs).toEqual({});
  });
});

describe('keyPairs: sheet columns, several files, appended rows', () => {
  it('columns are sheet columns (1-based) even when the table does not start in column A', () => {
    const b = sourceFromRows('base', BASE_ROWS.map((r) => [null, null, ...r]), '名簿');
    const s = sourceFromRows('S1', [
      [null, '受験番号', '評価', '名前'],
      [null, '1001', 'A', '山田 太郎'],
      [null, '1002', 'B', '鈴木 花子'],
      [null, '1003', 'C', '佐藤 翔太'],
    ]);
    const { res } = run(b, [s]);
    expect(b.region.firstCol).toBe(3);
    expect(s.region.firstCol).toBe(2);
    expect(res.keyPairs).toEqual({ S1: [{ baseCol: 3, sourceCol: 2 }, { baseCol: 4, sourceCol: 4 }] });
  });

  it('three files, full join: every source has its pairs, also when it matches a row appended from an earlier file', () => {
    const s1 = sourceFromRows('S1', [
      ['受験番号', '名前', '面接評価'],
      ['1001', '山田太郎', 'A'],
      ['2001', '田中 一郎', 'B'],
      ['2002', '小林 大輝', 'C'],
      ['1003', '佐藤 翔太', 'A'],
    ]);
    const s2 = sourceFromRows('S2', [
      ['ID', '氏名', '結果'],
      ['1002', '鈴木 花子', '内定'],
      ['2001', '田中 一郎', '内定'],
      ['3001', '加藤 結衣', '保留'],
    ]);
    const full = run(base(), [s1, s2], { joinType: 'full' }).res;
    expect(full.summary.appended).toBe(3);
    expect(full.keyPairs?.S1).toEqual([{ baseCol: 1, sourceCol: 1 }, { baseCol: 2, sourceCol: 2 }]);
    expect(full.keyPairs?.S2).toEqual([{ baseCol: 1, sourceCol: 1 }, { baseCol: 2, sourceCol: 2 }]);
    const left = run(base(), [s1, s2]).res;
    expect(left.keyPairs).toEqual(full.keyPairs);
  });

  it('a source that matches only rows appended from an earlier file still gets its pair', () => {
    const s1 = sourceFromRows('S1', [
      ['受験番号', '名前', '面接評価'],
      ['1001', '山田太郎', 'A'],
      ['2001', '田中 一郎', 'B'],
    ]);
    const s2 = sourceFromRows('S2', [
      ['志望ID', '結果'],
      ['2001', '内定'],
      ['3001', '保留'],
    ]);
    // no auto candidate (S2 shares nothing with the base itself): the user forces ID ⟷ 志望ID
    const an = addManualCandidate(analyzeSources(base(), [s1, s2]), 'S2', [1], [1]);
    const manual = an.pairs.S2.candidates[an.pairs.S2.candidates.length - 1];
    const full = run(base(), [s1, s2], { joinType: 'full', tiers: { ...defaultPlan(an).tiers, S2: [manual.id] } }, an).res;
    const appended = full.rows.filter((r) => r.baseRow === null);
    expect(appended.map((r) => r.matches.S2?.via)).toEqual([manual.label, '追加行']);
    expect(full.keyPairs?.S2).toEqual([{ baseCol: 1, sourceCol: 1 }]);
    // a left join cannot reach the appended row: S2 matches nothing → no entry
    const left = run(base(), [s1, s2], { tiers: { ...defaultPlan(an).tiers, S2: [manual.id] } }, an).res;
    expect(left.keyPairs?.S2).toBeUndefined();
    expect(left.keyPairs?.S1).toEqual([{ baseCol: 1, sourceCol: 1 }, { baseCol: 2, sourceCol: 2 }]);
  });

  it('different key columns per file', () => {
    const s1 = sourceFromRows('S1', [
      ['メモ', '受験番号', '評価'],
      ['x', '1001', 'A'],
      ['y', '1002', 'B'],
    ]);
    const s2 = sourceFromRows('S2', [
      ['氏名', '大学', '結果'],
      ['山田 太郎', '東京大学', '内定'],
      ['佐藤 翔太', '早稲田大学', '保留'],
      ['佐藤 翔太', '慶應義塾大学', '辞退'],
    ]);
    const { res } = run(base(), [s1, s2]);
    expect(res.keyPairs).toEqual({
      S1: [{ baseCol: 1, sourceCol: 2 }],
      S2: [{ baseCol: 2, sourceCol: 1 }, { baseCol: 4, sourceCol: 2 }],
    });
  });
});
