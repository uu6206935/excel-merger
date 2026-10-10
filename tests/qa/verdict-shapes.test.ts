// tools/lib/verdict.mjs + copies.mjs: the output shapes of round 9 (--single-sheet, --with-copies --source ...) on synthetic inspect dumps (no Excel).
import { describe, expect, it } from 'vitest';
import { checkOutput, formatVerdict } from '../../tools/lib/verdict.mjs';
import { expectedCopies, sheetNamesOf, textSheetName } from '../../tools/lib/copies.mjs';
import type { InspectResult, InspectSheet } from '../../tools/lib/inspect.mjs';
import { book, sheet } from './helpers';

// base: 名簿 (a list with a validation list that reads マスタ) + 集計 (reads 名簿) + マスタ; names: 志望度列 -> 名簿, 大学一覧 -> マスタ
function baseBook(): InspectResult {
  return book(
    [
      sheet({
        name: '名簿',
        cells: { A1: '学生ID', B1: '氏名', C1: '大学', D1: '志望度', A2: '000100', B2: '佐藤', C2: '東大', D2: 3, A3: '000107', B3: '鈴木', C3: '京大', D3: 5, A4: '000114', B4: '高橋', C4: '東大', D4: 4, C6: '合計', D6: 12, E6: 2 },
        formulas: { D6: '=SUM(D2:D4)', E6: '=COUNTIF(マスタ!$A$2:$A$3,C2)' },
        extra: { dv: [{ range: 'C2:C4', type: 3, operator: 1, formula1: '=大学一覧', formula2: '', alertStyle: 1 }] },
      }),
      sheet({
        name: '集計',
        cells: { A1: '大学', B1: '人数', A2: '東大', B2: 2, B3: 12 },
        formulas: { B2: '=COUNTIF(名簿!$C:$C,A2)', B3: '=SUM(志望度列)' },
        extra: { charts: [{ name: 'g', type: 51, title: 't', topLeft: 'E2', series: [{ name: '人数', formula: '=SERIES(集計!$B$1,集計!$A$2:$A$2,集計!$B$2:$B$2,1)' }] }] },
      }),
      sheet({ name: 'マスタ', cells: { A1: '大学', A2: '東大', A3: '京大' } }),
    ],
    {
      names: [
        { name: '志望度列', scope: null, refersTo: '=名簿!$D:$D', visible: true },
        { name: '大学一覧', scope: null, refersTo: '=マスタ!$A$2:$A$3', visible: true },
      ],
    },
  );
}

/** the app's default output: ONLY the merged sheet (a column 面接 imported; the formulas that read other sheets became values) */
function singleOutput(): InspectResult {
  return book(
    [
      sheet({
        name: '名簿',
        cells: { A1: '学生ID', B1: '氏名', C1: '大学', D1: '志望度', E1: '面接', A2: '000100', B2: '佐藤', C2: '東大', D2: 3, E2: 'A', A3: '000107', B3: '鈴木', C3: '京大', D3: 5, E3: 'B', A4: '000114', B4: '高橋', C4: '東大', D4: 4, E4: 'A', C6: '合計', D6: 12, E6: 2 },
        formulas: { D6: '=SUM(D2:D4)' }, // E6 read マスタ: now a value
        extra: { dv: [] }, // the list is a literal list now (or dropped): the base's rule no longer shows a reference
      }),
    ],
    { names: [{ name: '志望度列', scope: null, refersTo: '=名簿!$D:$D', visible: true }] },
  );
}

const codes = (v: ReturnType<typeof checkOutput>, level: string) => v.items.filter((i) => i.level === level).map((i) => i.code);

describe('checkOutput: the default is unchanged for outputs that keep every sheet', () => {
  it('a single-sheet output FAILS without --single-sheet, with a hint', () => {
    const v = checkOutput(baseBook(), singleOutput());
    expect(codes(v, 'fail')).toContain('sheets-lost');
    expect(v.items.find((i) => i.code === 'sheets-lost')!.message).toContain('--single-sheet');
    expect(v.ok).toBe(false);
  });
});

describe('checkOutput --single-sheet', () => {
  it('passes the single-sheet output: removed sheets / names / validation lists are INFO, formulas that read them are values', () => {
    const v = checkOutput(baseBook(), singleOutput(), { singleSheet: true });
    expect(codes(v, 'fail'), formatVerdict(v)).toEqual([]);
    expect(v.ok).toBe(true);
    const info = codes(v, 'info');
    expect(info).toEqual(expect.arrayContaining(['sheets-dropped', 'names-dropped', 'dv-literalized']));
    expect(codes(v, 'pass')).toEqual(expect.arrayContaining(['sheets', 'no-dangling-ref']));
    // the name that points at the merged sheet is still checked: losing it fails
    const lost = singleOutput();
    lost.names = [];
    expect(codes(checkOutput(baseBook(), lost, { singleSheet: true }), 'fail')).toContain('names-lost');
  });

  it('fails when the output still holds the other sheets, or when a formula reads a sheet that is not there', () => {
    const more = singleOutput();
    more.sheets!.push(sheet({ name: '集計', cells: { A1: '大学' } }));
    expect(codes(checkOutput(baseBook(), more, { singleSheet: true }), 'fail')).toContain('not-single-sheet');
    const dangling = singleOutput();
    dangling.sheets![0].formulas.E6 = '=COUNTIF(マスタ!$A$2:$A$3,C2)';
    const v = checkOutput(baseBook(), dangling, { singleSheet: true });
    expect(codes(v, 'fail')).toContain('dangling-sheet-ref');
    expect(v.items.find((i) => i.code === 'dangling-sheet-ref')!.examples![0]).toContain('名簿!E6');
  });

  it('a validation list that reads the merged sheet itself is not forgiven', () => {
    const b = baseBook();
    b.sheets![0].dv = [{ range: 'C2:C4', type: 3, operator: 1, formula1: '=$C$2:$C$3', formula2: '', alertStyle: 1 } as never];
    const o = singleOutput();
    o.sheets![0].dv = [];
    expect(codes(checkOutput(b, o, { singleSheet: true }), 'fail')).toContain('dv');
  });
});

// ---------------------------------------------------------------------------------------------------------------------------------------
// --with-copies
// ---------------------------------------------------------------------------------------------------------------------------------------

function sourceBook(): InspectResult {
  return book(
    [
      sheet({ name: '面接', cells: { A1: '学生ID', B1: '評価', A2: '000100', B2: 4, A3: '000107', B3: 5, B5: 9 }, formulas: { B5: '=SUM(B2:B3)' }, extra: { index: 1, merges: ['A1:B1'], cols: [{ col: 1, letter: 'A', width: 10, hidden: false }] }}),
      sheet({ name: 'メモ', cells: { A1: '備考', A2: 'ok', B2: 3 }, formulas: { B2: '=面接!B2-1' }, extra: { index: 2 } }),
    ],
    { file: 'ob.xlsx', path: 'C:\\ob.xlsx' },
  );
}

/** the copies as the app writes them: <file>_<sheet>, references renamed */
function copySheets(file = 'ob'): InspectSheet[] {
  return [
    sheet({ name: `${file}_面接`, cells: { A1: '学生ID', B1: '評価', A2: '000100', B2: 4, A3: '000107', B3: 5, B5: 9 }, formulas: { B5: '=SUM(B2:B3)' }, extra: { index: 3, merges: ['A1:B1'], cols: [{ col: 1, letter: 'A', width: 10, hidden: false }] } }),
    sheet({ name: `${file}_メモ`, cells: { A1: '備考', A2: 'ok', B2: 3 }, formulas: { B2: `=${file}_面接!B2-1` }, extra: { index: 4 } }),
  ];
}

function withCopiesOutput(): InspectResult {
  const o = baseBook();
  const merged = singleOutput().sheets![0];
  merged.dv = baseBook().sheets![0].dv; // the base's other sheets are kept: the validation list still reads マスタ
  merged.formulas.E6 = '=COUNTIF(マスタ!$A$2:$A$3,C2)';
  o.sheets = [merged, ...baseBook().sheets!.slice(1).map((s, i) => ({ ...s, index: i + 2 })), ...copySheets()];
  return o;
}

const SRC = [{ name: 'ob.xlsx', dump: sourceBook() }];

describe('checkOutput --with-copies', () => {
  it('names the copies by the rule and passes when every copy equals its source', () => {
    const exp = expectedCopies('名簿', ['集計', 'マスタ'], SRC);
    expect(exp.map((e) => e.name)).toEqual(['ob_面接', 'ob_メモ']);
    const v = checkOutput(baseBook(), withCopiesOutput(), { withCopies: true, sources: SRC });
    expect(codes(v, 'fail'), formatVerdict(v)).toEqual([]);
    expect(codes(v, 'pass')).toEqual(expect.arrayContaining(['copy-names', 'copy-equal']));
    expect(v.items.find((i) => i.code === 'copied-sheets')!.message).toContain('ob_面接');
    expect(v.stats.copies).toEqual({ compared: 2 });
  });

  it('fails on a missing / wrongly named / wrongly ordered copy', () => {
    const wrongName = withCopiesOutput();
    wrongName.sheets![4].name = 'ob_メモ書き';
    const v = checkOutput(baseBook(), wrongName, { withCopies: true, sources: SRC });
    expect(codes(v, 'fail')).toEqual(expect.arrayContaining(['copy-missing', 'copy-unexpected']));
    const swapped = withCopiesOutput();
    const [a, b] = [swapped.sheets![3], swapped.sheets![4]];
    swapped.sheets![3] = { ...b, index: 4 };
    swapped.sheets![4] = { ...a, index: 5 };
    expect(codes(checkOutput(baseBook(), swapped, { withCopies: true, sources: SRC }), 'fail')).toContain('copy-order');
  });

  it('fails when a copy is not its source: a value, a formula, a merge, a hidden column, hyperlinks, new errors', () => {
    const o1 = withCopiesOutput();
    o1.sheets![3].values.B2 = 40;
    expect(codes(checkOutput(baseBook(), o1, { withCopies: true, sources: SRC }), 'fail')).toContain('copy-differs');
    const o2 = withCopiesOutput();
    o2.sheets![4].formulas.B2 = '=ob_面接!B3-1';
    expect(codes(checkOutput(baseBook(), o2, { withCopies: true, sources: SRC }), 'fail')).toContain('copy-differs');
    const o3 = withCopiesOutput();
    o3.sheets![3].merges = [];
    expect(codes(checkOutput(baseBook(), o3, { withCopies: true, sources: SRC }), 'fail')).toContain('copy-differs');
    const o4 = withCopiesOutput();
    o4.sheets![3].cols = [{ col: 1, letter: 'A', width: 10, hidden: true }];
    expect(codes(checkOutput(baseBook(), o4, { withCopies: true, sources: SRC }), 'fail')).toContain('copy-differs');
    const o5 = withCopiesOutput();
    o5.sheets![3].hyperlinks = 2;
    expect(codes(checkOutput(baseBook(), o5, { withCopies: true, sources: SRC }), 'fail')).toContain('copy-differs');
    const o6 = withCopiesOutput();
    o6.sheets![4].errors = { B2: '#REF!' };
    delete o6.sheets![4].values.B2;
    const v6 = checkOutput(baseBook(), o6, { withCopies: true, sources: SRC });
    expect(codes(v6, 'fail')).toEqual(expect.arrayContaining(['copy-new-error', 'copy-differs']));
  });

  it('the errors of a copy are not counted against the base file; a source error is not new', () => {
    const src = sourceBook();
    src.sheets![1].values = { A1: '備考', A2: 'ok' };
    src.sheets![1].errors = { B2: '#REF!' };
    const o = withCopiesOutput();
    o.sheets![4].values = { A1: '備考', A2: 'ok' };
    o.sheets![4].errors = { B2: '#REF!' };
    const v = checkOutput(baseBook(), o, { withCopies: true, sources: [{ name: 'ob.xlsx', dump: src }] });
    expect(codes(v, 'fail'), formatVerdict(v)).toEqual([]);
  });

  it('a name that is taken gets (1): two files with the same name, and a name of the base', () => {
    const exp = expectedCopies('名簿', ['集計'], [SRC[0], SRC[0]]);
    expect(exp.map((e) => e.name)).toEqual(['ob_面接', 'ob_メモ', 'ob_面接(1)', 'ob_メモ(1)']);
    const o = withCopiesOutput();
    o.sheets = [...o.sheets!.filter((s) => s.name !== 'マスタ'), ...copySheets('ob').map((s, i) => ({ ...s, name: s.name + '(1)', index: 6 + i, formulas: Object.fromEntries(Object.entries(s.formulas).map(([k, f]) => [k, f.replace('ob_面接', 'ob_面接(1)')])) }))];
    const v = checkOutput(baseBook(), o, { withCopies: true, sources: [SRC[0], SRC[0]] });
    expect(codes(v, 'fail').filter((c) => c.startsWith('copy-m') || c.startsWith('copy-un'))).toEqual([]);
  });

  it('the 1904 date system moves dates by 1462 days: not a difference', () => {
    const src = sourceBook();
    src.date1904 = true;
    const o = withCopiesOutput();
    o.sheets![3].values.B2 = 4 - 1462; // (a date: the same day in the 1900 system is the serial minus 1462 ... the sign is the source's)
    const v = checkOutput(baseBook(), o, { withCopies: true, sources: [{ name: 'ob.xlsx', dump: { ...src, sheets: src.sheets!.map((s) => ({ ...s })) } }] });
    // the copy holds 4 - 1462: the source (1904) 4 + 1462... -> neither equals: this is a difference
    expect(codes(v, 'fail')).toContain('copy-differs');
    const o2 = withCopiesOutput();
    o2.sheets![3].values.B2 = 4 + 1462; // a 1904 source: the 1900 copy is serial + 1462
    const v2 = checkOutput(baseBook(), o2, { withCopies: true, sources: [{ name: 'ob.xlsx', dump: src }] });
    expect(codes(v2, 'fail'), formatVerdict(v2)).toEqual([]);
    expect(codes(v2, 'info')).toContain('copy-date-system');
  });

  it('combines with --single-sheet: the merged sheet and the copies only', () => {
    const o = singleOutput();
    o.sheets = [...o.sheets!, ...copySheets().map((s, i) => ({ ...s, index: 2 + i }))];
    const v = checkOutput(baseBook(), o, { singleSheet: true, withCopies: true, sources: SRC });
    expect(codes(v, 'fail'), formatVerdict(v)).toEqual([]);
    // ... but without --with-copies the extra sheets are a failure of the single-sheet expectation
    expect(codes(checkOutput(baseBook(), o, { singleSheet: true }), 'fail')).toContain('not-single-sheet');
  });

  it('a source Excel could not read is a warning, the names are still checked', () => {
    const v = checkOutput(baseBook(), withCopiesOutput(), { withCopies: true, sources: [{ name: 'ob.xlsx', dump: { ...sourceBook(), openMode: 'failed', sheets: undefined, error: 'x' } as InspectResult }] });
    expect(codes(v, 'warn')).toContain('copy-source');
  });

  it('a .csv source is not compared by value (Excel parses it by its own rules), an .xls source only by value (the app reads values only)', () => {
    // csv: "07062200284" is text in the app, the number 7062200284 in Excel
    const csv = sourceBook();
    csv.sheets = [csv.sheets![0]];
    // (Excel names the sheet after the file it opened - here a temporary name -, the app after the file name: c.csv -> c)
    csv.sheets[0].name = '8600e48f87e9';
    csv.sheets[0].values.A2 = 7062200284;
    const o = withCopiesOutput();
    o.sheets = [o.sheets![0], o.sheets![1], o.sheets![2], { ...o.sheets![3], name: 'c_c', values: { ...o.sheets![3].values, A2: '07062200284' }, formulas: {} }];
    const vc = checkOutput(baseBook(), o, { withCopies: true, sources: [{ name: 'c.csv', dump: csv }] });
    expect(codes(vc, 'fail'), formatVerdict(vc)).toEqual([]);
    expect(codes(vc, 'info')).toContain('copy-text-source');
    // xls: formulas and styles are not copied (values only), a changed value still fails
    const xls = sourceBook();
    const ox = withCopiesOutput();
    ox.sheets![3].formulas = {};
    ox.sheets![3].merges = [];
    ox.sheets![4].formulas = {};
    const v1 = checkOutput(baseBook(), ox, { withCopies: true, sources: [{ name: 'ob.xls', dump: xls }] });
    expect(codes(v1, 'fail').filter((c) => c.startsWith('copy-m') || c.startsWith('copy-u') || c === 'copy-order'), formatVerdict(v1)).toEqual([]);
    ox.sheets![3].values.B2 = 99;
    // (the copy names follow the file name: ob.xls -> ob_面接)
    expect(codes(checkOutput(baseBook(), ox, { withCopies: true, sources: [{ name: 'ob.xls', dump: xls }] }), 'fail')).toContain('copy-differs');
  });

  it('the only sheet of a text file is named after the file, like the app does (31 characters, forbidden characters become _)', () => {
    expect(textSheetName('03_連絡先.csv')).toBe('03_連絡先');
    expect(textSheetName('C:\\data\\a:b*c.tsv')).toBe('a_b_c');
    expect(textSheetName('x'.repeat(40) + '.txt')).toBe('x'.repeat(31));
    expect(textSheetName('.csv')).toBe('.csv'); // (no stem: the whole name, as the reader does)
    const csv = sourceBook();
    csv.sheets = [{ ...csv.sheets![0], name: 'abcdef123456' }];
    expect(expectedCopies('名簿', [], [{ name: '03_連絡先.csv', dump: csv }]).map((e) => e.name)).toEqual(['03_連絡先_03_連絡先']);
  });

  it('a pivot table renamed "_2" by Excel (a name is unique in the workbook) is the same pivot table', () => {
    const src = sourceBook();
    const piv = { name: 'P', range: 'A3:F15', source: 'R3C1:R63C12', rowFields: ['大学'], columnFields: ['学年'], dataFields: ['人数'] };
    src.sheets![0].pivots = [piv as never];
    const o = withCopiesOutput();
    o.sheets![3].pivots = [{ ...piv, name: 'P_2' } as never];
    const v = checkOutput(baseBook(), o, { withCopies: true, sources: [{ name: 'ob.xlsx', dump: src }] });
    expect(codes(v, 'fail'), formatVerdict(v)).toEqual([]);
    o.sheets![3].pivots = [{ ...piv, range: 'A3:F20' } as never];
    expect(codes(checkOutput(baseBook(), o, { withCopies: true, sources: [{ name: 'ob.xlsx', dump: src }] }), 'fail')).toContain('copy-differs');
  });

  it('charts of a Strict source in a normal output are a known limit (a warning, like the app says it), not a failure', () => {
    const src = sourceBook();
    src.sheets![0].charts = [{ name: 'g', type: 51, title: 't', topLeft: 'E2', series: [{ name: 's', formula: '=SERIES(面接!$B$1,面接!$A$2:$A$3,面接!$B$2:$B$3,1)' }] } as never];
    const o = withCopiesOutput(); // the copy has no chart
    const plain = checkOutput(baseBook(), o, { withCopies: true, sources: [{ name: 'ob.xlsx', dump: src }] });
    expect(codes(plain, 'fail')).toContain('copy-differs');
    const strict = checkOutput(baseBook(), o, { withCopies: true, sources: [{ name: 'ob.xlsx', dump: src, strict: true }] });
    expect(codes(strict, 'fail'), formatVerdict(strict)).toEqual([]);
    expect(codes(strict, 'warn')).toContain('copy-leftout');
    // both Strict (the copy is a Strict copy of a Strict sheet): the charts must be there
    expect(codes(checkOutput(baseBook(), o, { withCopies: true, sources: [{ name: 'ob.xlsx', dump: src, strict: true }], outStrict: true }), 'fail')).toContain('copy-differs');
    // ... and other differences of a Strict copy still fail
    const o2 = withCopiesOutput();
    o2.sheets![3].values.B2 = 40;
    expect(codes(checkOutput(baseBook(), o2, { withCopies: true, sources: [{ name: 'ob.xlsx', dump: src, strict: true }] }), 'fail')).toContain('copy-differs');
  });

  it('chart sheets are part of the names (their position is not known)', () => {
    const src = sourceBook();
    src.chartSheets = ['グラフ'];
    expect(sheetNamesOf(src)).toEqual(['面接', 'メモ', 'グラフ']);
    expect(expectedCopies('名簿', [], [{ name: 'ob.xlsx', dump: src }]).map((e) => e.name)).toEqual(['ob_面接', 'ob_メモ', 'ob_グラフ']);
  });
});
