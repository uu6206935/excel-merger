/**
 * addManualCandidate (the user picks the key columns themselves) and files without any usable key.
 */
import path from 'node:path';
import { describe, expect, it } from 'vitest';
import type { JoinPlan, KeyAnalysis, KeySource } from '../../src/core/contracts';
import { addManualCandidate, analyzeSources, defaultPlan, joinSources } from '../../src/core/keys';
import { type CellSpec, FIXTURES, bookFromFile, sourceFromRows, sourceOf } from './helpers';

const BASE_ROWS: CellSpec[][] = [
  ['応募者ID', '氏名', '大学', '電話', '部署'],
  ['1001', '山田 太郎', '東京大学', '090-1111-2222', '営業'],
  ['1002', '鈴木 花子', '慶應義塾大学', '090-3333-4444', '人事'],
  ['1003', '佐藤 翔太', '早稲田大学', '080-5555-6666', '営業'],
  ['1004', '中村 優子', '京都大学', '080-7777-8888', '開発'],
  ['1005', '髙橋 一郎', '京都大学', '070-1212-3434', '人事'],
  ['1006', '渡邊 恵', '大阪大学', '070-5656-7878', '開発'],
];
const base = () => sourceFromRows('base', BASE_ROWS, '名簿');

/** a file that only has a contact number (typed as an ID by its header/values) and answers: no auto key at all */
const phoneOnly = (id = 'O') => sourceFromRows(id, [
  ['連絡先', 'メモ', '評価'],
  ['09011112222', 'x1', 'A'],
  ['09033334444', 'x2', 'B'],
  ['08055556666', 'x3', 'C'],
  ['08077778888', 'x4', 'D'],
  ['07012123434', 'x5', 'A'],
]);

const planWith = (an: KeyAnalysis, tiers: Record<string, string[]>, patch: Partial<JoinPlan> = {}): JoinPlan => ({ ...defaultPlan(an), tiers, ...patch });
const last = <T>(xs: T[]): T => xs[xs.length - 1];
const matchRows = (res: ReturnType<typeof joinSources>, id: string) => res.rows.map((r) => r.matches[id]?.row ?? null);

describe('addManualCandidate: single column', () => {
  it('adds a candidate with real statistics, selected:false, registered for joinSources', () => {
    const an = analyzeSources(base(), [phoneOnly()]);
    expect(an.pairs.O.candidates).toEqual([]); // nothing usable was found automatically
    const an2 = addManualCandidate(an, 'O', [4], [1]); // 電話 ⟷ 連絡先
    expect(an2).not.toBe(an);
    expect(an2.pairs).not.toBe(an.pairs);
    expect(an2.pairs.O.candidates).not.toBe(an.pairs.O.candidates);
    expect(an2.pairs.O.candidates.length).toBe(1);
    const c = an2.pairs.O.candidates[0];
    expect(c).toMatchObject({
      id: 'manual:4|1', kind: 'single', label: '電話 ⟷ 連絡先', family: 'other',
      baseColumns: [{ col: 4, header: '電話' }], otherColumns: [{ col: 1, header: '連絡先' }],
      matchRate: 1, coverage: 0.833, matchedOther: 5, nonBlankOther: 5, matchedBase: 5, nonBlankBase: 6,
      uniqBase: 1, uniqOther: 1, dupRowsBase: 0, dupRowsOther: 0, blankBase: 0, blankOther: 0, ambiguous: 0,
      cardinality: '1:1', verdict: '良好', selected: false,
    });
    expect(c.tier).toBeUndefined();
    expect(c.score).toBeGreaterThan(0.8);
    // the untouched parts of the analysis are shared, the pair's other fields kept
    expect(an2.baseId).toBe('base');
    expect(an2.base).toBe(an.base);
    expect(an2.options).toBe(an.options);
    expect(an2.pairs.O.other).toBe(an.pairs.O.other);
    expect(an2.pairs.O.mode).toBe('oneToOne');
    // not part of the default plan (selected:false) — the UI adds the id
    expect(defaultPlan(an2).tiers.O).toEqual([]);
    const res = joinSources(an2, planWith(an2, { O: [c.id] }));
    expect(matchRows(res, 'O')).toEqual([2, 3, 4, 5, 6, null]);
    expect(res.rows[0].matches.O).toEqual({ row: 2, via: '電話 ⟷ 連絡先', flags: [] });
    expect(res.reports.O.plan).toEqual([{ order: 1, name: '電話 ⟷ 連絡先' }]);
    expect(res.reports.O.perTier).toEqual({ '電話 ⟷ 連絡先': 5 });
    expect(res.reports.O.unmatchedBase.map((r) => r.excelRow)).toEqual([7]);
    expect(res.items.some((i) => i.code === 'noKey')).toBe(false);
    // the key columns are not imported again; the base key cell of appended rows would be filled from the source
    expect(res.keyFill.O).toEqual([{ baseCol: 4, sourceCol: 1 }]);
    expect(res.defaultImports).toEqual([{ sourceId: 'O', col: 2 }, { sourceId: 'O', col: 3 }]);
  });

  it('copy-on-write: the analysis it was given does not know the new candidate', () => {
    const an = analyzeSources(base(), [phoneOnly()]);
    const an2 = addManualCandidate(an, 'O', [4], [1]);
    const res = joinSources(an, planWith(an, { O: ['manual:4|1'] }));
    expect(res.items.map((i) => i.code)).toEqual(expect.arrayContaining(['planInvalid', 'noKey']));
    expect(matchRows(res, 'O').every((r) => r === null)).toBe(true);
    expect(matchRows(joinSources(an2, planWith(an2, { O: ['manual:4|1'] })), 'O')).toEqual([2, 3, 4, 5, 6, null]);
  });

  it('normalises both columns the same way: the type the detector got right wins over a misread column', () => {
    // 連絡先 was typed as an ID, 電話 as a phone number: the phone normaliser makes 090-1111-2222 ⟷ 09011112222 equal
    const an = addManualCandidate(analyzeSources(base(), [phoneOnly()]), 'O', [4], [1]);
    expect(an.pairs.O.candidates[0].matchRate).toBe(1);
    // with the profile column order reversed (base 連絡先 ⟷ other 電話) the same holds
    const b2 = sourceFromRows('base', [
      ['連絡先', '氏名'],
      ['09011112222', '山田 太郎'],
      ['09033334444', '鈴木 花子'],
      ['08055556666', '佐藤 翔太'],
    ], '名簿');
    const o2 = sourceFromRows('O', [['電話', '評価'], ['090-1111-2222', 'A'], ['090-3333-4444', 'B'], ['080-5555-6666', 'C']]);
    const c2 = last(addManualCandidate(analyzeSources(b2, [o2]), 'O', [1], [1]).pairs.O.candidates);
    expect(c2).toMatchObject({ id: 'manual:1|1', matchRate: 1, coverage: 1, verdict: '良好' });
  });

  // non-unique columns: the value-overlap discovery never offers them, so a manual key over them is new
  const FRUITS = ['赤い りんご', '青い ぶどう', '赤い りんご', '青い ぶどう', '黄色い バナナ', '赤い りんご'];
  const NAMES6 = ['山田 太郎', '鈴木 花子', '佐藤 翔太', '中村 優子', '田中 一郎', '髙橋 次郎'];

  it('columns nothing recognised: generic text keys (全角空白・半角空白・大文字小文字)', () => {
    const b = sourceFromRows('base', [['氏名', 'X列'], ...NAMES6.map((n, i) => [n, FRUITS[i]])]);
    const o = sourceFromRows('O', [['名前', 'Y列', '評価'], ...NAMES6.map((n, i) => [n, FRUITS[i].replace(' ', '　'), 'A'])]);
    const an = analyzeSources(b, [o]);
    expect(an.base.columns[1].type).toBe('unknown');
    expect(an.pairs.O.candidates.some((c) => c.baseColumns[0].col === 2)).toBe(false);
    const c = last(addManualCandidate(an, 'O', [2], [2]).pairs.O.candidates);
    expect(c).toMatchObject({
      id: 'manual:2|2', kind: 'single', label: 'X列 ⟷ Y列', selected: false, matchedOther: 6, nonBlankOther: 6, matchedBase: 6, cardinality: 'N:M', verdict: '良好',
    });
    expect(c.matchRate).toBe(1);
    expect(c.uniqBase).toBeLessThan(0.6);
    expect(c.ambiguous).toBe(5); // every other row but the バナナ one hits several base rows
  });

  it('uses the analysis options (generic keys: 先頭ゼロを無視 off → 007 ≠ 7)', () => {
    const num = ['007', '008', '007', '008', '009', '007'];
    const b = sourceFromRows('base', [['氏名', 'X列'], ...NAMES6.map((n, i) => [n, num[i]])]);
    const o = sourceFromRows('O', [['名前', 'Y列', '評価'], ...NAMES6.map((n, i) => [n, String(Number(num[i])), 'A'])]);
    const col = (opts?: Parameters<typeof analyzeSources>[2]) => last(addManualCandidate(analyzeSources(b, [o], opts), 'O', [2], [2]).pairs.O.candidates);
    expect(col().matchedOther).toBe(6);
    expect(col({ leadingZeros: false }).matchedOther).toBe(0);
  });

  it('uses the analysis options (name keys: 敬称を除去 off → 山田太郎さん ≠ 山田 太郎)', () => {
    const b = sourceFromRows('base', [['氏名', '部署'], ['山田 太郎', '営業'], ['鈴木 花子', '人事'], ['佐藤 翔太', '開発']]);
    const o = sourceFromRows('O', [['通称', '評価'], ['山田太郎さん', 'A'], ['鈴木花子さん', 'B'], ['佐藤翔太さん', 'C']]);
    const col = (opts?: Parameters<typeof analyzeSources>[2]) => {
      const an = analyzeSources(b, [o], opts);
      expect(an.pairs.O.candidates).toEqual([]);
      return last(addManualCandidate(an, 'O', [1], [1]).pairs.O.candidates);
    };
    expect(col()).toMatchObject({ id: 'manual:1|1', label: '氏名 ⟷ 通称', family: 'name', matchRate: 1, verdict: '良好' });
    expect(col({ honorifics: false })).toMatchObject({ matchRate: 0, verdict: '使用不可' });
  });

  it('family follows the detected types: an ID column → id, a name column → name, nothing recognised → other', () => {
    const pts = [10, 20, 10, 20, 30, 10];
    const b = sourceFromRows('base', [['応募者ID', '氏名', '点数'], ...NAMES6.map((n, i) => [String(1001 + i), n, pts[i]])]);
    const o = sourceFromRows('O', [['受験番号', 'お名前', '評価点'], ...NAMES6.map((n, i) => [String(1001 + i), n, pts[i]])]);
    const an = analyzeSources(b, [o]);
    const add = (bc: number[], oc: number[]) => last(addManualCandidate(an, 'O', bc, oc).pairs.O.candidates);
    expect(add([3], [3])).toMatchObject({ id: 'manual:3|3', label: '点数 ⟷ 評価点', family: 'other', matchRate: 1 });
    expect(add([2], [3])).toMatchObject({ id: 'manual:2|3', family: 'name', matchRate: 0 }); // 氏名 ⟷ 評価点
    expect(add([1], [3])).toMatchObject({ id: 'manual:1|3', family: 'id', matchRate: 0 }); // 応募者ID ⟷ 評価点
    // a composite is an id when an ID is part of it, else a name when a name is part of it
    expect(add([1, 3], [1, 3]).family).toBe('id');
    expect(add([3, 2], [3, 2]).family).toBe('name');
  });

  it('chips are rebuilt: a manual 氏名 candidate appears in the chip of the file that had none, other files keep theirs', () => {
    const b = sourceFromRows('base', [['氏名', '部署'], ['山田 太郎', '営業'], ['鈴木 花子', '人事'], ['佐藤 翔太', '開発']]);
    const o = sourceFromRows('O', [['通称', '評価'], ['山田太郎さん', 'A'], ['鈴木花子さん', 'B'], ['佐藤翔太さん', 'C']]);
    const p = sourceFromRows('P', [['名前', '結果'], ['山田 太郎', 'A'], ['鈴木 花子', 'B'], ['佐藤 翔太', 'C']]);
    const an = analyzeSources(b, [o, p]);
    const nameChip = (a: KeyAnalysis) => a.chips.find((c) => c.family === 'name')!;
    expect(nameChip(an).perSource.O).toBeNull();
    expect(nameChip(an).perSource.P).toMatchObject({ label: '氏名' });
    const an2 = addManualCandidate(an, 'O', [1], [1]);
    expect(nameChip(an2).perSource.O).toEqual({ candidateId: 'manual:1|1', label: '氏名 ⟷ 通称', matchRate: 1 });
    expect(nameChip(an2).perSource.P).toEqual(nameChip(an).perSource.P);
    expect(an2.chips.find((c) => c.family === 'id')).toEqual(an.chips.find((c) => c.family === 'id'));
    expect(nameChip(an).perSource.O).toBeNull(); // the old analysis is unchanged
  });

  it('a label used by another candidate of the pair is qualified with the column letters', () => {
    const v1 = [1, 2, 1, 2, 3, 1];
    const v2 = [5, 5, 6, 6, 7, 7];
    const b = sourceFromRows('base', [['氏名', 'メモ', 'メモ'], ...NAMES6.map((n, i) => [n, v1[i], v2[i]])]);
    const o = sourceFromRows('O', [['お名前', 'メモ', '評価'], ...NAMES6.map((n, i) => [n, v1[i], 'A'])]);
    const an1 = addManualCandidate(analyzeSources(b, [o]), 'O', [2], [2]);
    const an2 = addManualCandidate(an1, 'O', [3], [2]);
    const labels = an2.pairs.O.candidates.map((c) => c.label);
    expect(new Set(labels).size).toBe(labels.length);
    expect(labels.slice(-2)).toEqual(['メモ ⟷ メモ', 'メモ ⟷ メモ（C ⟷ B）']);
    expect(an2.pairs.O.candidates.slice(-2).map((c) => c.id)).toEqual(['manual:2|2', 'manual:3|2']);
  });
});

describe('addManualCandidate: composite', () => {
  // 氏名 is unique in both files: 氏名＋大学 is not offered automatically, the user adds it
  const other = () => sourceFromRows('R', [
    ['名前', '学校名', '評価'],
    ['山田 太郎', '東京大学', 'A'],
    ['鈴木 花子', '慶応義塾大学', 'B'],
    ['佐藤 翔太', '早稲田', 'C'],
    ['中村 優子', '京都大学', 'D'],
    ['田中 一郎', '東京大学', 'E'],
  ]);

  it('氏名＋大学: kind composite, label of both header lists, family name, univ abbreviations normalised', () => {
    const an = analyzeSources(base(), [other()]);
    expect(an.pairs.R.candidates.some((c) => c.kind === 'composite')).toBe(false);
    const an2 = addManualCandidate(an, 'R', [2, 3], [1, 2]);
    const c = last(an2.pairs.R.candidates);
    expect(an2.pairs.R.candidates.length).toBe(an.pairs.R.candidates.length + 1);
    expect(c).toMatchObject({
      id: 'manual:2+3|1+2', kind: 'composite', label: '氏名＋大学 ⟷ 名前＋学校名', family: 'name', selected: false,
      baseColumns: [{ col: 2, header: '氏名' }, { col: 3, header: '大学' }],
      otherColumns: [{ col: 1, header: '名前' }, { col: 2, header: '学校名' }],
      matchedOther: 4, nonBlankOther: 5, matchedBase: 4, nonBlankBase: 6, matchRate: 0.8, verdict: '良好', cardinality: '1:1',
    });
    const res = joinSources(an2, planWith(an2, { R: [c.id] }));
    expect(matchRows(res, 'R')).toEqual([2, 3, 4, 5, null, null]);
    expect(res.rows[1].matches.R?.via).toBe('氏名＋大学 ⟷ 名前＋学校名');
    expect(res.keyFill.R).toEqual([{ baseCol: 2, sourceCol: 1 }, { baseCol: 3, sourceCol: 2 }]);
    expect(res.defaultImports).toEqual([{ sourceId: 'R', col: 3 }]);
    // the other candidates of the pair are untouched
    expect(an2.pairs.R.candidates.slice(0, -1)).toEqual(an.pairs.R.candidates);
  });

  it('大学名の表記ゆれを統一 off changes the statistics (real options, not defaults)', () => {
    const an = analyzeSources(base(), [other()], { universityAbbrev: false });
    const c = last(addManualCandidate(an, 'R', [2, 3], [1, 2]).pairs.R.candidates);
    expect(c.matchedOther).toBeLessThan(4);
  });

  it('氏名＋部署 tells two 同姓同名 rows apart where 氏名 alone is ambiguous', () => {
    const b = sourceFromRows('base', [
      ['氏名', '部署', '入社年'],
      ['佐藤 翔太', '営業', 2019],
      ['佐藤 翔太', '開発', 2021],
      ['鈴木 花子', '人事', 2020],
    ]);
    const o = sourceFromRows('O', [
      ['氏名', '所属', '評価'],
      ['佐藤 翔太', '開発', 'A'],
      ['佐藤 翔太', '営業', 'B'],
      ['鈴木 花子', '人事', 'C'],
    ]);
    const an = analyzeSources(b, [o]);
    const plain = joinSources(an, defaultPlan(an));
    expect(matchRows(plain, 'O')).toEqual([null, null, 4]); // both 佐藤 rows are ambiguous
    const an2 = addManualCandidate(an, 'O', [1, 2], [1, 2]);
    const c = last(an2.pairs.O.candidates);
    expect(c).toMatchObject({ kind: 'composite', label: '氏名＋部署 ⟷ 氏名＋所属', family: 'name', matchRate: 1, verdict: '良好', uniqBase: 1 });
    const res = joinSources(an2, planWith(an2, { O: [c.id, ...defaultPlan(an2).tiers.O] }));
    expect(matchRows(res, 'O')).toEqual([3, 2, 4]);
    expect(res.rows[0].matches.O?.via).toBe('氏名＋部署 ⟷ 氏名＋所属');
  });

  it('three columns', () => {
    const b = sourceFromRows('base', [['氏名', '大学', '部署', '電話'], ...BASE_ROWS.slice(1).map((r) => [r[1], r[2], r[4], r[3]])]);
    const o = sourceFromRows('O', [['名前', '学校名', '所属', 'TEL', '評価'], ...BASE_ROWS.slice(1, 5).map((r) => [r[1], r[2], r[4], r[3], 'A'])]);
    const an = addManualCandidate(analyzeSources(b, [o]), 'O', [1, 2, 3], [1, 2, 3]);
    const c = last(an.pairs.O.candidates);
    expect(c).toMatchObject({ kind: 'composite', label: '氏名＋大学＋部署 ⟷ 名前＋学校名＋所属', matchedOther: 4 });
    expect(joinSources(an, planWith(an, { O: [c.id] })).reports.O.matched).toBe(4);
  });
});

describe('addManualCandidate: invalid input', () => {
  const an = analyzeSources(base(), [phoneOnly('O'), phoneOnly('P')]);

  it.each<[string, number[], number[], RegExp]>([
    ['no base column', [], [1], /両方で1列以上/],
    ['no other column', [4], [], /両方で1列以上/],
    ['both empty', [], [], /両方で1列以上/],
    ['more than 3 columns', [1, 2, 3, 4], [1, 2, 3, 1], /最大3列まで/],
    ['different counts', [2, 3], [1], /列数（2）と「O」の列数（1）が違います/],
    ['column 0', [0], [1], /列の指定が正しくありません/],
    ['negative column', [4], [-1], /列の指定が正しくありません/],
    ['fractional column', [1.5], [1], /列の指定が正しくありません/],
    ['NaN', [Number.NaN], [1], /列の指定が正しくありません/],
    ['base column right of the region', [9], [1], /基準ファイルの列Iは、読み込んだ表の範囲（A〜E列）の外/],
    ['other column right of the region', [4], [4], /「O」の列Dは、読み込んだ表の範囲（A〜C列）の外/],
    ['repeated base column', [2, 2], [1, 2], /基準ファイルで同じ列（列B）が重複/],
    ['repeated other column', [2, 3], [1, 1], /「O」で同じ列（列A）が重複/],
  ])('%s', (_name, baseCols, otherCols, re) => {
    expect(() => addManualCandidate(an, 'O', baseCols, otherCols)).toThrow(re);
  });

  it('unknown file, the base itself, missing analysis, non-array input', () => {
    expect(() => addManualCandidate(an, 'nope', [1], [1])).toThrow(/結合キーの分析結果が見つかりません/);
    expect(() => addManualCandidate(an, 'base', [1], [1])).toThrow(/基準ファイル自身/);
    expect(() => addManualCandidate({ ...an, internal: undefined }, 'O', [1], [1])).toThrow(/キーの分析結果がありません/);
    expect(() => addManualCandidate(an, 'O', undefined as unknown as number[], [1])).toThrow(/列を選んでください/);
  });

  it('errors are plain Errors and leave the analysis usable', () => {
    expect(() => addManualCandidate(an, 'O', [], [])).toThrow(Error);
    const ok = addManualCandidate(an, 'O', [4], [1]);
    expect(last(ok.pairs.O.candidates).id).toBe('manual:4|1');
    expect(an.pairs.O.candidates).toEqual([]);
  });

  it('a column range narrower than the sheet: columns outside the region are refused', () => {
    const src = sourceFromRows('O', [['連絡先', 'メモ', '評価'], ['09011112222', 'x1', 'A'], ['09033334444', 'x2', 'B']]);
    const narrow: KeySource = { ...src, region: { ...src.region, firstCol: 2, lastCol: 3 } };
    const a2 = analyzeSources(base(), [narrow]);
    expect(() => addManualCandidate(a2, 'O', [4], [1])).toThrow(/「O」の列Aは、読み込んだ表の範囲（B〜C列）の外/);
    expect(last(addManualCandidate(a2, 'O', [4], [2]).pairs.O.candidates).otherColumns).toEqual([{ col: 2, header: 'メモ' }]);
  });
});

describe('addManualCandidate: idempotence', () => {
  it('the same columns again return the analysis itself', () => {
    const an = analyzeSources(base(), [phoneOnly()]);
    const a1 = addManualCandidate(an, 'O', [4], [1]);
    const a2 = addManualCandidate(a1, 'O', [4], [1]);
    expect(a2).toBe(a1);
    expect(a1.pairs.O.candidates.length).toBe(1);
  });

  it('columns of an automatic candidate return the analysis unchanged (ID ⟷ ID, 氏名 ⟷ 氏名)', () => {
    const o = sourceFromRows('O', [['受験番号', '名前', '評価'], ['1001', '山田 太郎', 'A'], ['1002', '鈴木 花子', 'B'], ['1003', '佐藤 翔太', 'C']]);
    const an = analyzeSources(base(), [o]);
    const ids = an.pairs.O.candidates.filter((c) => c.selected).map((c) => c.id);
    expect(ids.length).toBeGreaterThanOrEqual(2);
    expect(addManualCandidate(an, 'O', [1], [1])).toBe(an);
    expect(addManualCandidate(an, 'O', [2], [2])).toBe(an);
  });

  it('電話 ⟷ 電話: listed automatically when the numbers overlap (unchanged), added by hand when they do not', () => {
    const phones = BASE_ROWS.slice(1, 5).map((r) => r[3]);
    const same = sourceFromRows('O', [['電話', '評価'], ...phones.map((p) => [p, 'A'])]);
    const an = analyzeSources(base(), [same]);
    // (Round 5: a phone number that is unique on both sides is an automatic tier of its own, like the e-mail address)
    expect(an.pairs.O.candidates.map((c) => [c.kind, c.family, c.baseColumns[0].col, c.selected])).toEqual([['single', 'other', 4, true]]);
    expect(addManualCandidate(an, 'O', [4], [1])).toBe(an);
    // none of the numbers is in the base: the automatic list drops the hopeless candidate, the user can still add it
    const other = sourceFromRows('P', [['電話', '評価'], ['03-1111-0001', 'A'], ['03-1111-0002', 'B'], ['03-1111-0003', 'C']]);
    const an2 = analyzeSources(base(), [other]);
    expect(an2.pairs.P.candidates).toEqual([]);
    const an3 = addManualCandidate(an2, 'P', [4], [1]);
    expect(last(an3.pairs.P.candidates)).toMatchObject({ id: 'manual:4|1', label: '電話 ⟷ 電話', kind: 'single', family: 'other', matchRate: 0, verdict: '使用不可', selected: false });
    expect(addManualCandidate(an3, 'P', [4], [1])).toBe(an3);
  });

  it('a composite in another column order is the same key; a crossed pairing is a different one', () => {
    const o = sourceFromRows('R', [['名前', '学校名', '評価'], ['山田 太郎', '東京大学', 'A'], ['鈴木 花子', '慶應義塾大学', 'B'], ['佐藤 翔太', '早稲田大学', 'C']]);
    const an = analyzeSources(base(), [o]);
    const a1 = addManualCandidate(an, 'R', [2, 3], [1, 2]);
    expect(addManualCandidate(a1, 'R', [3, 2], [2, 1])).toBe(a1);
    const crossed = addManualCandidate(a1, 'R', [2, 3], [2, 1]);
    expect(crossed).not.toBe(a1);
    expect(last(crossed.pairs.R.candidates)).toMatchObject({ id: 'manual:2+3|2+1', verdict: '使用不可' });
  });

  it('several manual candidates accumulate; each pair is independent', () => {
    const an = analyzeSources(base(), [phoneOnly('O'), phoneOnly('P')]);
    const a1 = addManualCandidate(an, 'O', [4], [1]);
    const a2 = addManualCandidate(a1, 'P', [4], [1]);
    const a3 = addManualCandidate(a2, 'O', [5], [2]);
    expect(a3.pairs.O.candidates.map((c) => c.id)).toEqual(['manual:4|1', 'manual:5|2']);
    expect(a3.pairs.P.candidates.map((c) => c.id)).toEqual(['manual:4|1']);
    // earlier analyses keep their own lists
    expect(a1.pairs.O.candidates.length).toBe(1);
    expect(a2.pairs.O.candidates.length).toBe(1);
    const res = joinSources(a3, planWith(a3, { O: ['manual:4|1'], P: ['manual:4|1'] }));
    expect(matchRows(res, 'O')).toEqual([2, 3, 4, 5, 6, null]);
    expect(matchRows(res, 'P')).toEqual([2, 3, 4, 5, 6, null]);
  });
});

describe('manual candidates in multi-file full joins (appended rows)', () => {
  it('a base column the detector typed differently: later files find the appended rows through the filled key column', () => {
    // base: 連絡先 is typed as an ID (digits); the files have 電話 (phone). The manual key normalises both as phone numbers
    const b = sourceFromRows('base', [
      ['連絡先', '氏名'],
      ['09011112222', '山田 太郎'],
      ['09033334444', '鈴木 花子'],
      ['08055556666', '佐藤 翔太'],
    ], '名簿');
    const s1 = sourceFromRows('S1', [['電話', '面接'], ['090-1111-2222', 'A'], ['090-9999-0000', 'B'], ['080-5555-6666', 'C']]);
    const s2 = sourceFromRows('S2', [['電話番号', '結果'], ['09099990000', '内定'], ['090-3333-4444', '保留']]);
    let an = analyzeSources(b, [s1, s2]);
    expect(an.pairs.S1.candidates).toEqual([]);
    an = addManualCandidate(addManualCandidate(an, 'S1', [1], [1]), 'S2', [1], [1]);
    const res = joinSources(an, planWith(an, { S1: ['manual:1|1'], S2: ['manual:1|1'] }, { joinType: 'full' }));
    expect(res.summary).toEqual({ baseRows: 3, matchedAny: 3, matchedAll: 0, appended: 1 });
    const appended = res.rows.filter((r) => r.baseRow === null);
    expect(appended.map((r) => r.appendedFrom)).toEqual([{ sourceId: 'S1', row: 3 }]);
    // S2's 09099990000 finds the row appended from S1 (base 連絡先 filled with S1's 090-9999-0000)
    expect(appended[0].matches.S2).toEqual({ row: 2, via: '連絡先 ⟷ 電話番号', flags: [] });
    expect(res.rows[1].matches.S2?.row).toBe(3);
    expect(res.keyFill.S1).toEqual([{ baseCol: 1, sourceCol: 1 }]);
  });

  it('composite: parts keep their normaliser on the extended base table', () => {
    const b = sourceFromRows('base', [
      ['氏名', '部署', '入社年'],
      ['山田 太郎', '営業', 2019],
      ['鈴木 花子', '人事', 2020],
    ]);
    const s1 = sourceFromRows('S1', [['氏名', '所属', '評価'], ['山田 太郎', '営業', 'A'], ['田中 一郎', 'ＩＴ部', 'B']]);
    const s2 = sourceFromRows('S2', [['名前', '所属部署', '結果'], ['田中 一郎', 'IT部', '内定'], ['鈴木 花子', '人事', '保留']]);
    let an = analyzeSources(b, [s1, s2]);
    an = addManualCandidate(addManualCandidate(an, 'S1', [1, 2], [1, 2]), 'S2', [1, 2], [1, 2]);
    const id1 = last(an.pairs.S1.candidates).id;
    const id2 = last(an.pairs.S2.candidates).id;
    const res = joinSources(an, planWith(an, { S1: [id1], S2: [id2] }, { joinType: 'full' }));
    const appended = res.rows.filter((r) => r.baseRow === null);
    expect(appended.map((r) => r.appendedFrom)).toEqual([{ sourceId: 'S1', row: 3 }]);
    expect(appended[0].matches.S2?.row).toBe(2); // 田中 一郎 / IT部 (full-width ＩＴ部 in S1)
    expect(res.summary.appended).toBe(1);
    expect(res.keyFill.S1).toEqual([{ baseCol: 1, sourceCol: 1 }, { baseCol: 2, sourceCol: 2 }]);
  });
});

describe('files without any usable key', () => {
  const okFile = () => sourceFromRows('K', [['受験番号', '名前', '評価'], ['1001', '山田 太郎', 'A'], ['1002', '鈴木 花子', 'B'], ['2001', '田中 一郎', 'C']]);
  const strange = () => sourceFromRows('N', [['項目1', '項目2'], ['a1', 'b1'], ['a2', 'b2'], ['a3', 'b3'], ['a4', 'b4']]);

  it('analyzeSources: no candidate at all (nothing selected); defaultPlan gives an empty tier list', () => {
    const an = analyzeSources(base(), [okFile(), strange()]);
    expect(an.pairs.N.candidates).toEqual([]);
    expect(an.pairs.N.candidates.some((c) => c.selected)).toBe(false);
    expect(an.chips.map((c) => c.perSource.N)).toEqual([null, null]);
    expect(an.pairs.K.candidates.some((c) => c.selected)).toBe(true);
    const plan = defaultPlan(an);
    expect(plan.tiers.N).toEqual([]);
    expect(plan.tiers.K.length).toBeGreaterThan(0);
  });

  it('completely different ID systems and names: candidates are empty or all 使用不可 — never selected', () => {
    const o = sourceFromRows('X', [['受験番号', '名前', '評価'], ['9001', '小林 大輝', 'A'], ['9002', '加藤 結衣', 'B'], ['9003', '吉田 蓮', 'C']]);
    const an = analyzeSources(base(), [o]);
    expect(an.pairs.X.candidates.every((c) => !c.selected && c.verdict === '使用不可')).toBe(true);
    expect(defaultPlan(an).tiers.X).toEqual([]);
    const res = joinSources(an, defaultPlan(an));
    expect(matchRows(res, 'X').every((r) => r === null)).toBe(true);
    expect(res.items.some((i) => i.code === 'noKey')).toBe(true);
  });

  it('a barely overlapping key is listed but not selected: detection rule = no candidate with selected === true', () => {
    const rows: CellSpec[][] = [['応募者ID', '氏名']];
    for (let i = 0; i < 20; i++) rows.push([String(1000 + i), `氏${i}`]);
    const other: CellSpec[][] = [['応募者ID', '評価'], ['1001', 'A']];
    for (let i = 0; i < 9; i++) other.push([String(5000 + i), 'B']);
    const an = analyzeSources(sourceFromRows('base', rows), [sourceFromRows('O', other)]);
    expect(an.pairs.O.candidates.length).toBeGreaterThan(0);
    expect(an.pairs.O.candidates.every((c) => !c.selected)).toBe(true);
    expect(defaultPlan(an).tiers.O).toEqual([]);
  });

  it('joinSources with an empty tier list: no matches, nothing appended, noKey only', () => {
    const an = analyzeSources(base(), [strange()]);
    const tierCases: Array<Record<string, string[]>> = [{ N: [] }, {}]; // an empty list / the source missing from plan.tiers
    for (const tiers of tierCases) {
      for (const joinType of ['left', 'full'] as const) {
        const res = joinSources(an, planWith(an, tiers, { joinType }));
        expect(res.summary).toEqual({ baseRows: 6, matchedAny: 0, matchedAll: 0, appended: 0 });
        expect(res.rows.length).toBe(6);
        expect(matchRows(res, 'N').every((r) => r === null)).toBe(true);
        expect(res.reports.N.matched).toBe(0);
        expect(res.reports.N.plan).toEqual([]);
        expect(res.keyFill.N).toEqual([]);
        expect(res.items.filter((i) => i.sourceId === 'N').map((i) => [i.code, i.severity])).toEqual([['noKey', 'error']]);
      }
    }
  });

  it('the other files of the same join are not affected (full join, rows appended once)', () => {
    const an = analyzeSources(base(), [okFile(), strange()]);
    const res = joinSources(an, planWith(an, { ...defaultPlan(an).tiers, N: [] }, { joinType: 'full' }));
    expect(res.summary.appended).toBe(1); // 田中 一郎 of K only
    expect(res.rows.filter((r) => r.baseRow === null).map((r) => r.appendedFrom)).toEqual([{ sourceId: 'K', row: 4 }]);
    expect(matchRows(res, 'K').slice(0, 2)).toEqual([2, 3]);
    expect(res.rows.every((r) => r.matches.N === null)).toBe(true);
    expect(res.items.filter((i) => i.code === 'noKey').map((i) => i.sourceId)).toEqual(['N']);
    // the same result as when the keyless file is not part of the analysis, for the files that can be joined
    const an2 = analyzeSources(base(), [okFile()]);
    const res2 = joinSources(an2, { ...defaultPlan(an2), joinType: 'full' });
    expect(res.rows.map((r) => r.matches.K)).toEqual(res2.rows.map((r) => r.matches.K));
    expect(res.keyFill.K).toEqual(res2.keyFill.K);
  });

  it('a manual link alone is a key: the file is not keyless any more', () => {
    const an = analyzeSources(base(), [strange()]);
    const res = joinSources(an, planWith(an, { N: [] }, { manualLinks: { N: [{ baseRow: 3, otherRow: 4 }] } }));
    expect(matchRows(res, 'N')).toEqual([null, 4, null, null, null, null]); // base row 3 (鈴木) ⟷ row 4 of the file
    expect(res.items.some((i) => i.code === 'noKey')).toBe(false);
    expect(res.reports.N.perTier['手動']).toBe(1);
  });

  it('once the user adds a manual key, the same file joins', () => {
    const o = sourceFromRows('N', [['項目1', '項目2'], ['a1', '09011112222'], ['a2', '09033334444'], ['a3', '08055556666']]);
    const an = analyzeSources(base(), [o]);
    expect(defaultPlan(an).tiers.N).toEqual([]);
    const an2 = addManualCandidate(an, 'N', [4], [2]);
    const res = joinSources(an2, planWith(an2, { N: [last(an2.pairs.N.candidates).id] }));
    expect(matchRows(res, 'N')).toEqual([2, 3, 4, null, null, null]);
    expect(res.items.some((i) => i.code === 'noKey')).toBe(false);
  });
});

describe('interview sample files: 人事回答 has no key in common with the マッチング sheet', () => {
  const dir = path.join(FIXTURES, 'interview');
  const M = sourceOf(bookFromFile(path.join(dir, '学生-職場マッチング.xlsx'), 'M'), 'M', 'マッチング');
  const HR = sourceOf(bookFromFile(path.join(dir, '人事回答.xlsx'), 'HR'), 'HR', '人事回答');

  it('no candidate, empty default tiers; left and full joins contribute nothing; a manual (hopeless) key is listed as 使用不可', () => {
    const an = analyzeSources(M, [HR]);
    expect(an.pairs.HR.candidates).toEqual([]);
    expect(defaultPlan(an).tiers.HR).toEqual([]);
    for (const joinType of ['left', 'full'] as const) {
      const res = joinSources(an, { ...defaultPlan(an), joinType });
      expect(res.summary).toEqual({ baseRows: 30, matchedAny: 0, matchedAll: 0, appended: 0 });
      expect(res.items.filter((i) => i.severity === 'error').map((i) => i.code)).toEqual(['noKey']);
    }
    // 職場ID (col 4) ⟷ 応募者コード (col 1): different numbering systems
    const an2 = addManualCandidate(an, 'HR', [4], [1]);
    expect(last(an2.pairs.HR.candidates)).toMatchObject({ id: 'manual:4|1', kind: 'single', family: 'id', matchRate: 0, verdict: '使用不可', selected: false });
    expect(defaultPlan(an2).tiers.HR).toEqual([]);
  });
});
