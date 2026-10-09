import { describe, expect, it } from 'vitest';
import { SHEET_NAME_MAX, cutUnits, fileBaseName, joinSheetName, planSheetNames, sanitizeSheetName, sheetNameKey, withSuffix } from '../../src/core/xlsx/sheetnames';

const names = (p: ReturnType<typeof planSheetNames>) => p.map((x) => x.name);
const lone = /[\ud800-\udbff](?![\udc00-\udfff])|(?<![\ud800-\udbff])[\udc00-\udfff]/;

describe('fileBaseName', () => {
  it('drops the folder and a known extension only', () => {
    expect(fileBaseName('D:\\x\\02_面接評価.xlsx')).toBe('02_面接評価');
    expect(fileBaseName('/a/b/データ.v2.XLSM')).toBe('データ.v2');
    expect(fileBaseName('名簿.csv')).toBe('名簿');
    expect(fileBaseName('名簿')).toBe('名簿');
    expect(fileBaseName('report.final')).toBe('report.final');
  });
});

describe('planSheetNames: <file>_<sheet>', () => {
  it('uses the file name and the sheet name', () => {
    const p = planSheetNames(['名簿'], [{ file: '02_面接評価.xlsx', sheets: ['面接評価', 'Sheet2'] }]);
    expect(names(p)).toEqual(['02_面接評価_面接評価', '02_面接評価_Sheet2']);
    expect(p.every((x) => !x.suffixed)).toBe(true);
    expect(p[0]).toMatchObject({ file: '02_面接評価.xlsx', original: '面接評価' });
  });

  it('keeps the order of files and sheets', () => {
    const p = planSheetNames([], [
      { file: 'b.xlsx', sheets: ['2', '1'] },
      { file: 'a.xlsx', sheets: ['x'] },
    ]);
    expect(names(p)).toEqual(['b_2', 'b_1', 'a_x']);
  });

  it('a later sheet gets (1), (2) ... when the name is taken; the earlier one keeps its name', () => {
    const p = planSheetNames(['名簿', '名簿_名簿'], [
      { file: '名簿.xlsx', sheets: ['名簿'] },
      { file: '名簿.xlsm', sheets: ['名簿'] },
      { file: '名簿.xlsx', sheets: ['名簿'] },
    ]);
    expect(names(p)).toEqual(['名簿_名簿(1)', '名簿_名簿(2)', '名簿_名簿(3)']);
    expect(p.every((x) => x.suffixed)).toBe(true);
  });

  it('never renames the existing sheets and compares case-insensitively / ignoring full width', () => {
    expect(names(planSheetNames(['A_B'], [{ file: 'a.xlsx', sheets: ['b'] }]))).toEqual(['a_b(1)']);
    expect(names(planSheetNames(['ＡＢ_ｃ'], [{ file: 'ab.xlsx', sheets: ['C'] }]))).toEqual(['ab_C(1)']);
  });

  it('the same sheet name twice inside one file (case only) is made unique too', () => {
    expect(names(planSheetNames([], [{ file: 'a.xlsx', sheets: ['b_c', 'B_C'] }]))).toEqual(['a_b_c', 'a_B_C(1)']);
  });

  it('is deterministic and does not mutate its input', () => {
    const ex = ['x'];
    const rq = [{ file: 'a.xlsx', sheets: ['s'] }];
    expect(planSheetNames(ex, rq)).toEqual(planSheetNames(ex, rq));
    expect(ex).toEqual(['x']);
    expect(rq).toEqual([{ file: 'a.xlsx', sheets: ['s'] }]);
  });
});

describe('Excel limits', () => {
  it('never exceeds 31 UTF-16 code units, also with the (n) suffix', () => {
    const long = 'とても長いファイル名のエクセルブック_2026年度版';
    const p = planSheetNames([], [{ file: long + '.xlsx', sheets: ['成績一覧表_最終版', '成績一覧表_最終版', 'x'.repeat(40)] }]);
    for (const x of p) expect(x.name.length).toBeLessThanOrEqual(SHEET_NAME_MAX);
    expect(new Set(p.map((x) => sheetNameKey(x.name))).size).toBe(p.length);
    expect(p[1].suffixed).toBe(true);
    expect(p[1].name.endsWith('(1)')).toBe(true);
  });

  it('shortens the file part first and keeps the sheet name', () => {
    expect(joinSheetName('とても長いファイル名のエクセルブック_2026年度版', '成績一覧表_最終版')).toBe('とても長いファイル名のエクセルブック_20_成績一覧表_最終版');
    expect(joinSheetName('A', 'x'.repeat(40))).toBe('A_' + 'x'.repeat(29));
  });

  it('does not split a surrogate pair', () => {
    const n = joinSheetName('𠮷野家注文一覧データベース', '𠮷'.repeat(20));
    expect(n.length).toBeLessThanOrEqual(31);
    expect(lone.test(n)).toBe(false);
    expect(lone.test(cutUnits('a𠮷b', 2))).toBe(false);
    expect(withSuffix('𠮷'.repeat(20), 1).length).toBeLessThanOrEqual(31);
  });

  it('replaces characters Excel forbids and apostrophes at both ends', () => {
    expect(sanitizeSheetName("[a]:b*c?d/e\\f")).toBe('_a__b_c_d_e_f');
    expect(sanitizeSheetName("'abc'")).toBe('abc');
    expect(sanitizeSheetName('')).toBe('Sheet');
    expect(sanitizeSheetName("'''")).toBe('Sheet');
    expect(joinSheetName('bad:name?', '[s]')).toBe('bad_name__' + '_s_');
  });

  it('property: random names always give valid, unique names', () => {
    const alphabet = ['a', 'B', '名', '簿', '𠮷', ' ', '_', "'", '[', ']', ':', '?', '/', '\\', '1', '(', ')'];
    let seed = 12345;
    const rnd = (n: number) => ((seed = (seed * 1103515245 + 12345) & 0x7fffffff), seed % n);
    const word = (max: number) => Array.from({ length: rnd(max) + 1 }, () => alphabet[rnd(alphabet.length)]).join('');
    for (let round = 0; round < 300; round++) {
      const existing = Array.from({ length: rnd(3) }, () => word(8));
      const reqs = Array.from({ length: rnd(3) + 1 }, () => ({ file: word(40) + '.xlsx', sheets: Array.from({ length: rnd(4) + 1 }, () => word(40)) }));
      const p = planSheetNames(existing, reqs);
      const keys = new Set(existing.map(sheetNameKey));
      for (const x of p) {
        expect(x.name.length).toBeGreaterThan(0);
        expect(x.name.length).toBeLessThanOrEqual(SHEET_NAME_MAX);
        expect(/[\\/?*[\]:]/.test(x.name)).toBe(false);
        expect(/^'|'$/.test(x.name)).toBe(false);
        expect(lone.test(x.name)).toBe(false);
        expect(keys.has(sheetNameKey(x.name))).toBe(false);
        keys.add(sheetNameKey(x.name));
      }
      expect(p.length).toBe(reqs.reduce((a, r) => a + r.sheets.length, 0));
    }
  });
});
