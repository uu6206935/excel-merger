import { describe, expect, it } from 'vitest';
import {
  colToNum,
  compareInspections,
  diffSummary,
  extractRefs,
  formatComparison,
  headerIndex,
  headerRowGuess,
  headerText,
  numToCol,
  parseRange,
  parseRanges,
  summarize,
} from '../../tools/lib/compare-core.mjs';
import { book, sheet } from './helpers';

describe('A1 helpers', () => {
  it('converts columns', () => {
    expect(colToNum('A')).toBe(1);
    expect(colToNum('Z')).toBe(26);
    expect(colToNum('AA')).toBe(27);
    expect(colToNum('XFD')).toBe(16384);
    expect(numToCol(16384)).toBe('XFD');
    expect(numToCol(28)).toBe('AB');
  });
  it('parses ranges', () => {
    expect(parseRange('$B$2:$D$9')).toEqual({ r1: 2, c1: 2, r2: 9, c2: 4 });
    expect(parseRange('K:K')).toEqual({ r1: 1, c1: 11, r2: 1048576, c2: 11 });
    expect(parseRange('3:5')).toEqual({ r1: 3, c1: 1, r2: 5, c2: 16384 });
    expect(parseRange('C7')).toEqual({ r1: 7, c1: 3, r2: 7, c2: 3 });
    expect(parseRange('名簿')).toBeNull();
    expect(parseRanges('A1:B2,D4 名簿!E5')).toHaveLength(3);
  });
});

describe('extractRefs', () => {
  const texts = (f: string, s: string | null = 'S') => extractRefs(f, s).map((r) => `${r.sheet}!${r.text.replace(/^.*!/, '')}`);

  it('finds qualified and unqualified refs', () => {
    expect(texts('=COUNTIF(名簿!$D:$D,$A5)')).toEqual(['名簿!$D:$D', 'S!$A5']);
    expect(texts("=SUM('集計 シート'!B2:B9)*A1")).toEqual(['集計 シート!B2:B9', 'S!A1']);
    expect(texts("='It''s'!C3")).toEqual(["It's!C3"]);
    expect(texts('=SUM(3:3)')).toEqual(['S!3:3']);
  });
  it('ignores strings, functions, names, structured refs and booleans', () => {
    expect(texts('=IF(A1="B2","C3",LOG10(D4))')).toEqual(['S!A1', 'S!D4']);
    expect(texts('=SUM(志望度列)+基準日')).toEqual([]);
    expect(texts('=SUM(Table1[金額])+Table1[[#This Row],[x]]')).toEqual([]);
    expect(texts('=TRUE+FALSE')).toEqual([]);
    expect(texts('=ATAN2(1,2)')).toEqual([]);
  });
  it('handles #REF! and 3-D references without inventing refs', () => {
    expect(texts('=COUNTIF(名簿!#REF!,$A5)')).toEqual(['S!$A5']);
    expect(texts('=SUM(Sheet1:Sheet3!A1)')).toEqual([]);
  });
  it('keeps every reference of a long formula', () => {
    const f = '=IFERROR(SUMPRODUCT((名簿!$D$4:$D$63=$A5)*名簿!$K$4:$K$63)/B5,"-")';
    expect(texts(f)).toEqual(['名簿!$D$4:$D$63', 'S!$A5', '名簿!$K$4:$K$63', 'S!B5']);
    const r = extractRefs(f, 'S')[0];
    expect(r.box).toEqual({ r1: 4, c1: 4, r2: 63, c2: 4 });
    expect(r.qualified).toBe(true);
  });
  it('reads a chart SERIES formula', () => {
    expect(texts('=SERIES(集計!$B$4,集計!$A$5:$A$14,集計!$B$5:$B$14,1)')).toEqual(['集計!$B$4', '集計!$A$5:$A$14', '集計!$B$5:$B$14']);
  });
});

describe('headers + comparison', () => {
  const s = sheet({
    name: '名簿',
    cells: { A1: 'タイトル', A3: '学生ID', B3: '氏名', C3: '大学', A4: '000100', B4: '佐藤', C4: '東大' },
  });
  it('guesses the header row and reads header texts', () => {
    expect(headerRowGuess(s)).toBe(3);
    expect(headerText(s, 3, 2)).toBe('氏名');
    expect(headerIndex(s, 3).get('大学')).toEqual([3]);
  });
  it('reports identical dumps as identical and lists differences', () => {
    const a = book([s]);
    expect(compareInspections(a, a).every((r) => r.ok)).toBe(true);
    const b = book([sheet({ name: '名簿', cells: { A1: 'タイトル', A3: '学生ID', B3: '氏名', C3: '大学', A4: '000100', B4: '佐藤', C4: '京大' } })]);
    const rows = compareInspections(a, b);
    const bad = rows.filter((r) => !r.ok);
    expect(bad.map((r) => r.group)).toEqual(['values']);
    expect(bad[0].detail).toContain('C4');
    expect(formatComparison(rows, { onlyDiffs: true })).toContain('| values | 名簿 | DIFF |');
    expect(diffSummary(rows)).toBe('values 1');
    // identical: no empty table, only the filtered list (if any)
    expect(formatComparison(compareInspections(a, a), { onlyDiffs: true })).toBe('');
    const withFiltered = [...compareInspections(a, a), { group: 'values', sheet: '名簿', ok: true, detail: '揮発性関数 1 セル', filtered: true }];
    expect(formatComparison(withFiltered, { onlyDiffs: true })).toBe('Filtered (known-benign, not counted):\n- [values / 名簿] 揮発性関数 1 セル');
    expect(diffSummary(withFiltered)).toBe('');
    expect(summarize(a)).toContain('[名簿]');
  });
  it('ignores the re-laid-out widths of pivot sheets unless asked', () => {
    const pivots = [{ name: 'p', range: 'A3:F15', source: '名簿!R3C1:R63C12', rowFields: ['大学'], columnFields: [], dataFields: ['人数'] }];
    const cols = (w: number) => [{ col: 1, letter: 'A', width: w, hidden: false }];
    const a = book([sheet({ name: 'ピボット', extra: { pivots, cols: cols(12.38) } })]);
    const b = book([sheet({ name: 'ピボット', extra: { pivots, cols: cols(16) } })]);
    expect(compareInspections(a, b).every((r) => r.ok)).toBe(true);
    expect(compareInspections(a, b, { pivotLayout: true }).filter((r) => !r.ok).map((r) => r.group)).toEqual(['cols']);
  });
});
