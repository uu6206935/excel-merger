import { beforeAll, describe, expect, it } from 'vitest';
import type { JoinResult, KeyAnalysis } from '../../src/core/contracts';
import {
  addedCandidate,
  candidateByColumns,
  carryOverPlan,
  checkGroups,
  checkItemWhere,
  chipViews,
  columnOptions,
  effectiveImports,
  excludedSourceIds,
  extraChips,
  fileKeyColumns,
  fileKeyLine,
  importChoices,
  importKey,
  isManualCandidate,
  joinSummary,
  keySources,
  manualKeyHint,
  matchRate,
  matchRateText,
  matchRateTone,
  moveCandidate,
  orderedCandidates,
  otherCandidateCount,
  plannedColumns,
  previewRows,
  readyFiles,
  rowsList,
  setManualLink,
  sourcesWithoutKeys,
  summaryText,
  summaryUnit,
  tierLabels,
  toggleCandidate,
  toggleFamily,
  unmatchedEverywhere,
  verdictTone,
} from '../../src/state/keysPlan';
import type { FileEntry } from '../../src/state/types';
import { joined, svc, type Joined } from './fixtures';

let j: Joined;
const labels = { s1: '学生回答', s2: '職場回答', s3: '人事回答' };

beforeAll(async () => {
  j = await joined();
});

describe('sources', () => {
  it('only usable files take part; the chosen base comes first', () => {
    const f = (id: string, extra: Partial<FileEntry>): FileEntry => ({ id, fileName: `${id}.xlsx`, size: 1, color: '#000', status: 'ready', ...extra });
    const files: FileEntry[] = [
      f('s1', { book: j.base.book, sheet: j.base.region.sheet, region: j.base.region }),
      f('s2', { book: j.others[0].book, sheet: j.others[0].region.sheet, region: j.others[0].region }),
      f('s3', { status: 'loading' }),
      f('s4', { book: j.others[1].book, sheet: 'x', region: { ...j.others[1].region, lastDataRow: 0, firstDataRow: 2 } }), // no data rows
    ];
    expect(readyFiles(files).map((x) => x.id)).toEqual(['s1', 's2']);
    const ks = keySources(files, 's2')!;
    expect(ks.base.id).toBe('s2');
    expect(ks.others.map((o) => o.id)).toEqual(['s1']);
    expect(ks.base.label).toBe('s2');
    expect(keySources(files.slice(0, 1), 's1')).toBeNull();
  });
});

describe('chips and plan edits', () => {
  it('ID / 氏名 chips are both on with the default plan', () => {
    const chips = chipViews(j.analysis, j.plan, labels);
    expect(chips.map((c) => [c.family, c.on, c.partial, c.available])).toEqual([
      ['id', true, false, true],
      ['name', true, false, true],
    ]);
    expect(chips[0].details.join('\n')).toContain('職場回答');
  });

  it('toggling a family removes it everywhere and puts it back at its natural priority', () => {
    const off = toggleFamily(j.analysis, j.plan, 'id', false);
    for (const s of ['s2', 's3']) expect(off.tiers[s]).toEqual(['name']);
    expect(chipViews(j.analysis, off, labels)[0].on).toBe(false);
    const on = toggleFamily(j.analysis, off, 'id', true);
    for (const s of ['s2', 's3']) expect(on.tiers[s]).toEqual(['id', 'name']);
  });

  it('partial chip when only some pairs use a family', () => {
    const p = toggleCandidate(j.plan, 's2', 'name', false);
    const name = chipViews(j.analysis, p, labels)[1];
    expect(name.on).toBe(false);
    expect(name.partial).toBe(true);
  });

  it('candidate checkbox, priority and ordering', () => {
    const p1 = toggleCandidate(j.plan, 's3', 'name+univ', true);
    expect(p1.tiers.s3).toEqual(['id', 'name', 'name+univ']);
    const p2 = moveCandidate(p1, 's3', 'name+univ', -1);
    expect(p2.tiers.s3).toEqual(['id', 'name+univ', 'name']);
    expect(moveCandidate(p2, 's3', 'id', -1)).toBe(p2); // already first
    const ord = orderedCandidates(j.analysis.pairs.s3, p2).map((x) => [x.c.id, x.priority]);
    expect(ord).toEqual([
      ['id', 1],
      ['name+univ', 2],
      ['name', 3],
    ]);
    expect(tierLabels(j.analysis, p2, 's3')).toBe('ID → 氏名＋大学 → 氏名');
    expect(tierLabels(j.analysis, { ...p2, tiers: { ...p2.tiers, s3: [] } }, 's3')).toBe('なし');
  });

  it('extra chips / other candidates / sources without keys', () => {
    expect(extraChips(j.analysis, j.plan)).toEqual([]);
    expect(otherCandidateCount(j.analysis, j.plan)).toBe(1); // 氏名＋大学 of 人事回答
    const none = { ...j.plan, tiers: { ...j.plan.tiers, s2: [] } };
    expect(sourcesWithoutKeys(j.analysis, none)).toEqual(['s2']);
    const linked = setManualLink(none, 's2', 5, 6, true);
    expect(sourcesWithoutKeys(j.analysis, linked)).toEqual([]);
  });

  it('manual links: one link per base row / other row, and removable', () => {
    let p = setManualLink(j.plan, 's2', 5, 6, true);
    p = setManualLink(p, 's2', 5, 7, true); // replaces the link of base row 5
    expect(p.manualLinks.s2).toEqual([{ baseRow: 5, otherRow: 7 }]);
    p = setManualLink(p, 's2', 5, 7, false);
    expect(p.manualLinks.s2).toEqual([]);
  });

  it('carry-over keeps join type / policy always, tiers and links only for the same files', () => {
    const prev = { ...setManualLink(toggleFamily(j.analysis, j.plan, 'name', false), 's2', 3, 4, true), joinType: 'full' as const, duplicatePolicy: 'last' as const };
    const fresh = svc.defaultPlan(j.analysis);
    const same = carryOverPlan(fresh, prev, { sameFiles: true, analysis: j.analysis });
    expect(same.tiers.s2).toEqual(['id']);
    expect(same.manualLinks.s2).toEqual([{ baseRow: 3, otherRow: 4 }]);
    expect(same.joinType).toBe('full');
    expect(same.duplicatePolicy).toBe('last');
    const other = carryOverPlan(fresh, prev, { sameFiles: false, analysis: j.analysis });
    expect(other.tiers.s2).toEqual(['id', 'name']);
    expect(other.manualLinks.s2).toBeUndefined();
    expect(other.joinType).toBe('full');
    // a candidate that no longer exists -> the fresh tiers
    const gone = carryOverPlan(fresh, { ...prev, tiers: { ...prev.tiers, s2: ['nope'] } }, { sameFiles: true, analysis: j.analysis });
    expect(gone.tiers.s2).toEqual(['id', 'name']);
    expect(carryOverPlan(fresh, null, { sameFiles: true, analysis: j.analysis })).toBe(fresh);
  });

  it('verdict tones', () => {
    expect(['良好', '一部一致', '不一致多数', '使用不可', '?'].map(verdictTone)).toEqual(['good', 'partial', 'poor', 'bad', 'bad']);
  });
});

describe('import columns', () => {
  it('defaults = every non-key column; overrides win; key columns can be imported explicitly', () => {
    const ch = importChoices(j.analysis, j.plan, j.join, {}, 's2');
    expect(ch.map((c) => [c.header, c.on, c.isKey])).toEqual([
      ['学生ID', false, true],
      ['氏名', false, true],
      ['配属先', true, false],
      ['評価', true, false],
      ['面接日', true, false],
      ['コメント', true, false],
    ]);
    const sheet = j.analysis.pairs.s2.other.sheet;
    const ov = { [importKey('s2', sheet, 4)]: false, [importKey('s2', sheet, 1)]: true };
    const eff = effectiveImports(j.analysis, j.plan, j.join, ov).filter((i) => i.sourceId === 's2');
    expect(eff.map((i) => i.col)).toEqual([1, 3, 5, 6]);
    expect(effectiveImports(null, j.plan, j.join, {})).toEqual([]);
  });

  it('planned columns = base columns + imports, conflicting headers get the file name', () => {
    const sheet = j.analysis.pairs.s2.other.sheet;
    const cols = plannedColumns(j.analysis, j.plan, j.join, { [importKey('s2', sheet, 2)]: true });
    expect(cols.slice(0, 3).map((c) => c.header)).toEqual(['No', '学生ID', '氏名']);
    expect(cols.filter((c) => !c.imported)).toHaveLength(11);
    expect(cols.find((c) => c.key === 'i:s2:2')?.header).toBe('氏名（職場回答）');
    expect(cols.filter((c) => c.imported)).toHaveLength(8);
  });
});

describe('summary and checks', () => {
  it('summary counts rows matched by every file', () => {
    const s = joinSummary(j.join);
    expect(s).toMatchObject({ baseRows: 30, matched: 27, partial: 2, unmatched: 1, appended: 0 });
    expect(s.duplicates).toBe(1);
    expect(summaryText(s, '名')).toBe(`基準 30 名 / 一致 27 / 一部一致 2 / 未一致 1 / 重複 1${s.review ? ` / 要確認 ${s.review}` : ''}`);
    expect(summaryText({ ...s, baseRows: 10000, matched: 9999, partial: 0, unmatched: 1, review: 0 }, '行')).toBe('基準 10,000 行 / 一致 9,999 / 未一致 1 / 重複 1');
    expect(summaryUnit(j.analysis, j.plan)).toBe('名');
    expect(summaryUnit(null, null)).toBe('行');
  });

  it('check groups: rows in no file (未一致) and rows missing in some files, with the files that miss them', () => {
    const groups = checkGroups(j.join, labels, { count: 2, rows: [] });
    const un = groups.find((g) => g.id === 'unmatched')!;
    expect(un.count).toBe(1);
    expect(un.rows.map((r) => r.row)).toEqual([22]);
    expect(un.rows[0].sub).toBe('どのファイルにも該当なし');
    const part = groups.find((g) => g.id === 'partial')!;
    expect(part.rows.map((r) => [r.row, r.sub])).toEqual([
      [9, '職場回答 に該当なし'],
      [23, '人事回答 に該当なし'],
    ]);
    expect([...unmatchedEverywhere(j.join)]).toEqual([22]);
    expect(unmatchedEverywhere(null).size).toBe(0);
    expect(groups.find((g) => g.id === 'duplicate')!.count).toBe(1);
    expect(groups.find((g) => g.id === 'review')!.rows.some((r) => r.sub?.includes('氏名が異なる'))).toBe(true);
    expect(groups.find((g) => g.id === 'toValues')!.count).toBe(2);
    expect(groups.find((g) => g.id === 'otherOnly')!.count).toBe(2);
  });

  it('the 要確認 badge counts like the summary line (a flag by its count, conflicts / ambiguous rows once each)', () => {
    const base = Object.values(j.join.reports)[0];
    const join = {
      ...j.join,
      reports: {
        s2: {
          ...base,
          conflicts: [{ base: { excelRow: 5, preview: ['A'] }, other: { excelRow: 6, preview: ['B'] }, message: '氏名が異なる' }],
          ambiguous: [{ base: { excelRow: 7, preview: ['C'] }, candidates: [{ excelRow: 8, preview: [] }, { excelRow: 9, preview: [] }], tier: 1 }],
          flags: { '重複（ルールで1行を採用）': { level: 'warn', count: 3, examples: ['9行目 ⟷ 23行目'] }, 異体字の違い: { level: 'info', count: 5, examples: [] } },
        },
      },
    };
    const review = checkGroups(join, labels).find((g) => g.id === 'review')!;
    expect(review.count).toBe(joinSummary(join).review);
    expect(review.count).toBe(1 + 1 + 3);
    expect(review.rows).toHaveLength(3);
  });

  it('rows of the OTHER file (CheckItem.otherRows) read 「職場回答 の 12行目」 and sit on top of their group, uncounted', () => {
    const groups = checkGroups(j.join, labels);
    const other = groups.find((g) => g.id === 'otherOnly')!;
    expect(other.rows[0].text).toMatch(/^(職場回答|人事回答) の \d+行目$/);
    const dup = groups.find((g) => g.id === 'duplicate')!;
    expect(dup.count).toBe(1); // summaries do not change the counts
    expect(dup.summaries?.[0].text).toContain('「職場回答」でキーが重複している行');
    expect(dup.summaries?.[0].sub).toMatch(/^職場回答 の \d+行目・\d+行目 \/ 基準の \d+行目$/);
    expect(other.summaries?.length).toBeGreaterThan(0);

    // items without otherRows are not repeated; unknown codes land in 要確認
    const extra: JoinResult = {
      ...j.join,
      items: [
        { severity: 'warning', code: 'unmatched', message: 'x', sourceId: 's2', rows: [5] },
        { severity: 'warning', code: 'somethingNew', message: '新しい確認事項', sourceId: 's3', otherRows: [12] },
      ],
    };
    const g2 = checkGroups(extra, labels);
    expect(g2.find((g) => g.id === 'unmatched')!.summaries).toBeUndefined();
    expect(g2.find((g) => g.id === 'review')!.summaries).toEqual([{ key: 'it1', row: null, text: '新しい確認事項', sub: '人事回答 の 12行目', tone: 'warning' }]);
  });

  it('where-text of check items', () => {
    expect(rowsList([12, 13, 12])).toBe('12行目・13行目');
    expect(rowsList([1, 2, 3, 4, 5, 6, 7])).toBe('1行目・2行目・3行目・4行目・5行目 ほか 2 件');
    expect(checkItemWhere({ severity: 'info', code: 'a', message: 'm', sourceId: 's2', otherRows: [12] }, labels)).toBe('職場回答 の 12行目');
    expect(checkItemWhere({ severity: 'info', code: 'a', message: 'm', sourceId: 'zz', otherRows: [3], rows: [4] }, labels)).toBe('このファイル の 3行目 / 基準の 4行目');
    expect(checkItemWhere({ severity: 'info', code: 'a', message: 'm' }, labels)).toBeUndefined();
  });

  it('preview rows per file', () => {
    const rows = previewRows(j.analysis, j.plan, j.join);
    expect(rows.map((r) => [r.sourceId, r.rows, r.keys, r.matched])).toEqual([
      ['s1', 30, '基準', null],
      ['s2', 30, 'ID → 氏名', 28],
      ['s3', 28, 'ID → 氏名', 28],
    ]);
    expect(rows[0].unmatched).toBe(1);
  });

  it('key columns of each card', () => {
    expect(fileKeyLine(fileKeyColumns(j.analysis, j.plan, 's1'))).toBe('ID: 学生ID / 氏名: 氏名');
    expect(fileKeyLine(fileKeyColumns(j.analysis, j.plan, 's3'))).toBe('ID: 学籍番号 / 氏名: 名前');
    expect(fileKeyColumns(null, null, 's1')).toEqual([]);
    const withUniv = toggleCandidate(j.plan, 's3', 'name+univ', true);
    expect(fileKeyLine(fileKeyColumns(j.analysis, withUniv, 's3'))).toBe('ID: 学籍番号 / 氏名: 名前 / 氏名＋大学: 大学名');
  });

  it('a join result with appended rows shows them in the summary', () => {
    const full = svc.joinSources(j.analysis, { ...j.plan, joinType: 'full' }) as JoinResult;
    expect(full.summary.appended).toBe(2);
    expect(summaryText(joinSummary(full), '名')).toContain('追加 2');
    expect(full.rows.filter((r) => r.baseRow === null)).toHaveLength(2);
    const a: KeyAnalysis = j.analysis;
    expect(Object.keys(a.pairs)).toEqual(['s2', 's3']);
  });
});

describe('excluded files in the pure helpers', () => {
  const none = (): typeof j.plan => ({ ...j.plan, tiers: { ...j.plan.tiers, s3: [] } });

  it('excludedSourceIds = the files without a key (a manual link counts as a way to join)', () => {
    expect(excludedSourceIds(j.analysis, j.plan)).toEqual([]);
    expect(excludedSourceIds(j.analysis, none())).toEqual(['s3']);
    expect(excludedSourceIds(j.analysis, setManualLink(none(), 's3', 4, 5, true))).toEqual([]);
    expect(excludedSourceIds(null, null)).toEqual([]);
  });

  it('their columns are neither imported nor planned', () => {
    const plan = none();
    const eff = effectiveImports(j.analysis, plan, j.join, {});
    expect(eff.some((i) => i.sourceId === 's3')).toBe(false);
    expect(eff.some((i) => i.sourceId === 's2')).toBe(true);
    expect(plannedColumns(j.analysis, plan, j.join, {}).some((c) => c.sourceId === 's3')).toBe(false);
    // the choices of the detail modal are still listed (the user may prepare them)
    expect(importChoices(j.analysis, plan, j.join, {}, 's3').length).toBeGreaterThan(0);
  });

  it('the preview row says 結合対象外 and has no numbers', () => {
    const rows = previewRows(j.analysis, none(), j.join);
    expect(rows[2]).toMatchObject({ sourceId: 's3', excluded: true, keys: '結合対象外', matched: null, unmatched: null, duplicates: null, blanks: null });
    expect(rows[1].excluded).toBe(false);
  });

  it('summaryText appends the number of excluded files only when there are some', () => {
    const sum = joinSummary(j.join);
    expect(summaryText(sum, '名', 2)).toBe(`${summaryText(sum, '名')} / 結合対象外 2 ファイル`);
  });
});

describe('match rate sentence (Power Query style)', () => {
  it('counts the rows of the other file / of the base that found a partner', () => {
    const m = matchRate(j.analysis, j.join, 's2')!;
    expect(m.baseRows).toBe(30);
    expect(m.otherRows).toBe(30);
    expect(m.matchedBase).toBe(30 - j.join.reports.s2.unmatchedBase.length);
    expect(m.matchedOther).toBe(30 - j.join.reports.s2.unmatchedOther.length);
    expect(matchRateText({ otherRows: 63, baseRows: 60, matchedOther: 58, matchedBase: 58 })).toBe('このキーで 58 / 63 行が一致します（基準 60 行中 58 行が見つかりました）');
    expect(matchRateText({ otherRows: 12345, baseRows: 10000, matchedOther: 12000, matchedBase: 9999 })).toBe('このキーで 12,000 / 12,345 行が一致します（基準 10,000 行中 9,999 行が見つかりました）');
    expect(matchRate(j.analysis, null, 's2')).toBeNull();
    expect(matchRate(j.analysis, j.join, 'zz')).toBeNull();
  });

  it('follows the selected keys live', () => {
    const idOnly = toggleCandidate(toggleCandidate(j.plan, 's2', 'name', false), 's3', 'name', false);
    const a = matchRate(j.analysis, j.join, 's2')!;
    const b = matchRate(j.analysis, svc.joinSources(j.analysis, idOnly), 's2')!;
    expect(b.matchedOther).toBeLessThanOrEqual(a.matchedOther);
    const nothing = svc.joinSources(j.analysis, { ...j.plan, tiers: { s2: [], s3: [] } });
    expect(matchRate(j.analysis, nothing, 's2')!.matchedOther).toBe(0);
  });

  it('tone by the share of the other file that matched', () => {
    expect(matchRateTone({ otherRows: 10, matchedOther: 9, baseRows: 10, matchedBase: 9 })).toBe('good');
    expect(matchRateTone({ otherRows: 10, matchedOther: 6, baseRows: 10, matchedBase: 6 })).toBe('partial');
    expect(matchRateTone({ otherRows: 10, matchedOther: 1, baseRows: 10, matchedBase: 1 })).toBe('poor');
    expect(matchRateTone({ otherRows: 0, matchedOther: 0, baseRows: 0, matchedBase: 0 })).toBe('poor');
  });
});

describe('manual key picker helpers', () => {
  it('every column of the table by header text + column letter', () => {
    const cols = j.analysis.base.columns;
    const opts = columnOptions(j.analysis.base);
    expect(opts).toHaveLength(cols.length);
    expect(opts.slice(0, 3).map((o) => o.label)).toEqual(['No（A列）', '学生ID（B列）', '氏名（C列）']);
    expect(columnOptions({ ...j.analysis.base, columns: [{ ...cols[0], header: '  ' }] })[0].label).toBe('（見出しなし）（A列）');
  });

  it('the live hint: detected types, why 「追加」 is not possible yet', () => {
    const cols = j.analysis.base.columns;
    const id = cols.find((c) => c.header === '学生ID')!;
    const name = cols.find((c) => c.header === '氏名')!;
    const otherId = j.analysis.pairs.s2.other.columns.find((c) => c.header === '学生ID')!;
    expect(manualKeyHint([undefined], [undefined])).toMatchObject({ tone: 'none', ready: false });
    expect(manualKeyHint([id], [undefined]).text).toContain('選んでください');
    expect(manualKeyHint([id], [])).toMatchObject({ ready: false });
    expect(manualKeyHint([id, name], [otherId]).text).toContain('同じ数の列');
    expect(manualKeyHint([id, id], [otherId, otherId]).text).toContain('重複');
    const ok = manualKeyHint([id], [otherId]);
    expect(ok).toMatchObject({ tone: 'ok', ready: true });
    expect(ok.text).toContain(`基準 学生ID（${id.typeLabel}）`);
    const diff = manualKeyHint([name], [otherId]);
    expect(diff).toMatchObject({ tone: 'warn', ready: true });
    expect(diff.text).toContain('種類が違います');
  });

  it('manual candidates are recognised by id or by the stored list; columns find their candidate', () => {
    expect(isManualCandidate({ id: 'manual:3|2' })).toBe(true);
    expect(isManualCandidate({ id: 'single:3|2' })).toBe(false);
    expect(isManualCandidate({ id: 'x' }, [{ id: 'x' }])).toBe(true);
    const pair = j.analysis.pairs.s2;
    const idc = pair.candidates.find((c) => c.id === 'id')!;
    expect(candidateByColumns(pair, idc.baseColumns.map((c) => c.col), idc.otherColumns.map((c) => c.col))?.id).toBe('id');
    expect(candidateByColumns(pair, [99], [99])).toBeUndefined();
  });

  it('addedCandidate finds the new candidate, else the one with the same columns', () => {
    const before = j.analysis.pairs.s2;
    const after = svc.addManualCandidate(j.analysis, 's2', [1], [1]);
    expect(addedCandidate(before, after.pairs.s2, [1], [1])?.id).toBe('manual:1|1');
    // nothing new (the keys module returned the same analysis): the candidate over the same columns
    const idc = before.candidates.find((c) => c.id === 'id')!;
    expect(addedCandidate(before, before, idc.baseColumns.map((c) => c.col), idc.otherColumns.map((c) => c.col))?.id).toBe('id');
    expect(addedCandidate(before, undefined, [1], [1])).toBeUndefined();
  });
});
