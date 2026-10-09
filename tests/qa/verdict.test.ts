import { describe, expect, it } from 'vitest';
import { checkOutput, formatVerdict } from '../../tools/lib/verdict.mjs';
import type { InspectResult } from '../../tools/lib/inspect.mjs';
import { book, sheet } from './helpers';

// base: 名簿 (header row 1: 学生ID 氏名 大学 志望度, rows 2..4, footer row 6) + 集計 + name 志望度列
function baseBook(): InspectResult {
  return book(
    [
      sheet({
        name: '名簿',
        cells: {
          A1: '学生ID', B1: '氏名', C1: '大学', D1: '志望度',
          A2: '000100', B2: '佐藤', C2: '東大', D2: 3,
          A3: '000107', B3: '鈴木', C3: '京大', D3: 5,
          A4: '000114', B4: '高橋', C4: '東大', D4: 4,
          C6: '平均', D6: 4,
        },
        formulas: { D6: '=AVERAGE(D2:D4)' },
        extra: {
          cf: [{ type: 1, appliesTo: 'D2:D4', operator: 7, formula1: '=4', formula2: '', priority: 1 }],
          freeze: { frozen: true, splitRow: 1, splitCol: 1 },
        },
      }),
      sheet({
        name: '集計',
        cells: { A1: '大学', B1: '人数', C1: '平均', A2: '東大', B2: 2, C2: 4, B3: 12 },
        formulas: { B2: '=COUNTIF(名簿!$C:$C,A2)', C2: '=AVERAGE(名簿!D2:D4)', B3: '=SUM(志望度列)' },
        extra: { charts: [{ name: 'g', type: 51, title: 't', topLeft: 'E2', series: [{ name: '人数', formula: '=SERIES(集計!$B$1,集計!$A$2:$A$2,集計!$B$2:$B$2,1)' }] }] },
      }),
    ],
    { names: [{ name: '志望度列', scope: null, refersTo: '=名簿!$D:$D', visible: true }] },
  );
}

// a correct export: 志望度 moved to B, one imported column E (面接), everything adjusted
function goodOutput(): InspectResult {
  return book(
    [
      sheet({
        name: '名簿',
        cells: {
          A1: '学生ID', B1: '志望度', C1: '氏名', D1: '大学', E1: '面接',
          A2: '000100', B2: 3, C2: '佐藤', D2: '東大', E2: 'A',
          A3: '000107', B3: 5, C3: '鈴木', D3: '京大', E3: 'B',
          A4: '000114', B4: 4, C4: '高橋', D4: '東大', E4: 'A',
          D6: '平均', B6: 4,
        },
        formulas: { B6: '=AVERAGE(B2:B4)' },
        extra: {
          cf: [{ type: 1, appliesTo: 'B2:B4', operator: 7, formula1: '=4', formula2: '', priority: 1 }],
          freeze: { frozen: true, splitRow: 1, splitCol: 1 },
        },
      }),
      sheet({
        name: '集計',
        cells: { A1: '大学', B1: '人数', C1: '平均', A2: '東大', B2: 2, C2: 4, B3: 12 },
        formulas: { B2: '=COUNTIF(名簿!$D:$D,A2)', C2: '=AVERAGE(名簿!B2:B4)', B3: '=SUM(志望度列)' },
        extra: { charts: [{ name: 'g', type: 51, title: 't', topLeft: 'E2', series: [{ name: '人数', formula: '=SERIES(集計!$B$1,集計!$A$2:$A$2,集計!$B$2:$B$2,1)' }] }] },
      }),
    ],
    { names: [{ name: '志望度列', scope: null, refersTo: '=名簿!$B:$B', visible: true }] },
  );
}

const codes = (v: ReturnType<typeof checkOutput>, level: string) => v.items.filter((i) => i.level === level).map((i) => i.code);

describe('checkOutput', () => {
  it('passes a correct export', () => {
    const v = checkOutput(baseBook(), goodOutput());
    expect(codes(v, 'fail')).toEqual([]);
    expect(v.ok).toBe(true);
    expect(v.stats.mergedSheet).toBe('名簿');
    expect(v.stats.addedHeaders).toEqual(['面接']);
    expect(v.stats.formulasChecked).toBe(4);
    expect(formatVerdict(v)).toContain('合格');
  });

  it('fails when the output needs repair or cannot be opened', () => {
    const v = checkOutput(baseBook(), { ...goodOutput(), openMode: 'repair', repairLog: ['removedRecord: x'] });
    expect(codes(v, 'fail')).toContain('repair');
    const f = checkOutput(baseBook(), { ...goodOutput(), openMode: 'failed', sheets: undefined });
    expect(codes(f, 'fail')).toEqual(['open-failed']);
  });

  it('fails when an other-sheet formula now points to another header', () => {
    const out = goodOutput();
    out.sheets![1].formulas.B2 = '=COUNTIF(名簿!$C:$C,A2)'; // not adjusted: now 氏名
    const v = checkOutput(baseBook(), out);
    expect(codes(v, 'fail')).toContain('formula-columns');
    expect(v.items.find((i) => i.code === 'formula-columns')!.examples![0]).toContain('名簿!氏名');
  });

  it('fails when a merged-sheet formula (footer) points to another header', () => {
    const out = goodOutput();
    out.sheets![0].formulas.B6 = '=AVERAGE(C2:C4)';
    expect(codes(checkOutput(baseBook(), out), 'fail')).toContain('formula-columns');
  });

  it('fails when a defined name was not adjusted', () => {
    const out = goodOutput();
    out.names![0].refersTo = '=名簿!$D:$D';
    expect(codes(checkOutput(baseBook(), out), 'fail')).toContain('formula-columns');
  });

  it('fails on new #REF! / #NAME? and lost names or charts', () => {
    const out = goodOutput();
    out.sheets![1].formulas.C2 = '=AVERAGE(名簿!#REF!)';
    out.sheets![1].errors = { C2: '#REF!', B3: '#NAME?' };
    out.sheets![1].errorCounts = { '#REF!': 1, '#NAME?': 1 };
    out.sheets![1].charts = [];
    out.names = [];
    const v = checkOutput(baseBook(), out);
    const f = codes(v, 'fail');
    expect(f).toEqual(expect.arrayContaining(['new-ref-text', 'new-error-#REF!', 'new-error-#NAME?', 'names-lost', 'charts-lost']));
    expect(v.ok).toBe(false);
  });

  it('treats formulas converted to values as info, and a vanished formula as a failure', () => {
    const out = goodOutput();
    delete out.sheets![1].formulas.B2;
    const v = checkOutput(baseBook(), out);
    expect(codes(v, 'info')).toContain('formulas-to-values');
    expect(v.ok).toBe(true);
    const out2 = goodOutput();
    delete out2.sheets![1].formulas.B2;
    delete out2.sheets![1].values.B2;
    expect(codes(checkOutput(baseBook(), out2), 'fail')).toContain('formulas-lost');
  });

  it('follows the footer when rows were appended (join full)', () => {
    const out = goodOutput();
    const s = out.sheets![0];
    s.values.A5 = '000601';
    s.values.C5 = '新井';
    delete s.values.D6;
    delete s.values.B6;
    delete s.formulas.B6;
    s.values.D7 = '平均';
    s.values.B7 = 4;
    s.formulas.B7 = '=AVERAGE(B2:B5)';
    s.maxRow = 7;
    const v = checkOutput(baseBook(), out);
    expect(codes(v, 'fail')).toEqual([]);
  });

  it('fails when CF no longer sits on the same header column', () => {
    const out = goodOutput();
    out.sheets![0].cf[0].appliesTo = 'D2:D4'; // now on 大学
    expect(codes(checkOutput(baseBook(), out), 'fail')).toContain('cf');
  });

  it('a deleted column leaving the pivot source / CF / DV is info; losing a column that still exists fails', () => {
    const pivot = (source: string) => [{ name: 'p', range: 'A3:C9', source, rowFields: ['氏名'], columnFields: [], dataFields: ['合計'] }];
    const mk = (cols: 2 | 3, source: string, rules: boolean) =>
      book([
        sheet({
          name: '名簿',
          cells: {
            A1: '学生ID', B1: '氏名', A2: '000100', B2: '佐藤', A3: '000107', B3: '鈴木', A4: '000114', B4: '高橋',
            ...(cols === 3 ? { C1: '志望度', C2: 3, C3: 5, C4: 4 } : {}),
          },
          extra: rules
            ? {
                cf: [{ type: 1, appliesTo: 'C2:C4', operator: 7, formula1: '=4', formula2: '', priority: 1 }],
                dv: [{ range: 'C2:C4', type: 3, operator: null, formula1: '1,2,3,4,5', formula2: null, alertStyle: 1 }],
              }
            : {},
        }),
        sheet({ name: 'ピボット', cells: { A1: 'x' }, extra: { pivots: pivot(source) } }),
      ]);
    const base = mk(3, '名簿!R1C1:R4C3', true);
    const v = checkOutput(base, mk(2, '名簿!R1C1:R4C2', false));
    expect(codes(v, 'fail')).toEqual([]);
    expect(codes(v, 'info')).toEqual(expect.arrayContaining(['pivot-source-deleted', 'cf-deleted', 'dv-deleted', 'removed-columns']));
    // 氏名 is still in the sheet but no longer in the pivot source -> a real failure
    const bad = mk(3, '名簿!R1C1:R4C1', true);
    expect(codes(checkOutput(base, bad), 'fail')).toContain('pivot-source');
  });

  it('join full: reports which bounded ranges over the table grew over the appended rows', () => {
    // same columns as the base, one row appended (row 5), footer pushed down to row 7
    const out = baseBook();
    const s = out.sheets![0];
    Object.assign(s.values, { A5: '000601', B5: '新井', C5: '京大', D5: 2 });
    delete s.values.C6;
    delete s.values.D6;
    delete s.formulas.D6;
    s.values.C7 = '平均';
    s.values.D7 = 3.5;
    s.formulas.D7 = '=AVERAGE(D2:D5)'; // grew with the table
    s.cf = [{ type: 1, appliesTo: 'D2:D5', operator: 7, formula1: '=4', formula2: '', priority: 1 }];
    s.maxRow = 7;
    // 集計!C2 =AVERAGE(名簿!D2:D4) was NOT extended -> the appended row is missing from that average
    const v = checkOutput(baseBook(), out);
    expect(codes(v, 'fail')).toEqual([]);
    expect(v.stats.appended).toEqual({ rows: 1, firstRow: 5, lastRow: 5, ranges: 3, extended: 2, notExtended: 1 });
    expect(codes(v, 'info')).toEqual(expect.arrayContaining(['appended-rows', 'appended-ranges']));
    const excluded = v.items.find((i) => i.code === 'appended-ranges-excluded')!;
    expect(excluded.level).toBe('warn');
    expect(excluded.examples).toEqual(['集計!C2: 名簿!D2:D4 → 名簿!D2:D4']);
    const extended = v.items.find((i) => i.code === 'appended-ranges')!.examples!;
    expect(extended).toEqual(expect.arrayContaining(['名簿!D7: D2:D4 → D2:D5', '条件付き書式（志望度 列）: D2:D4']));
    // whole-column references (COUNTIF(名簿!$C:$C), 志望度列 = 名簿!$D:$D) always include new rows: not listed
    expect(JSON.stringify(v.items)).not.toContain('$C:$C →');
  });
});

// CF / DV ranges that GREW over a column Excel itself inserted next to a column carrying the rule (insert right of
// 志望度: verified Excel behaviour, both engines do it) are fine; only a rule that no longer covers the column with
// the original header is a failure.
describe('checkOutput: CF / DV ranges that grew over an inserted column', () => {
  const withRules = (cols: number, cf: string, dv: string | null, cells: Record<string, string | number> = {}) => {
    const b = baseBook();
    const s = b.sheets![0];
    s.maxCol = cols;
    Object.assign(s.values, cells);
    s.cf = [{ type: 1, appliesTo: cf, operator: 7, formula1: '=4', formula2: '', priority: 1 }];
    s.dv = dv ? [{ range: dv, type: 3, operator: null, formula1: '1,2,3,4,5', formula2: null, alertStyle: 1 }] : [];
    return b;
  };
  const base = () => withRules(4, 'D2:D4', 'D2:D4');

  it('passes a CF / DV that grew over the new empty column right of the rule column and reports it as info', () => {
    const v = checkOutput(base(), withRules(5, 'D2:E4', 'D2:E4'));
    expect(codes(v, 'fail')).toEqual([]);
    expect(v.ok).toBe(true);
    expect(codes(v, 'pass')).toEqual(expect.arrayContaining(['cf']));
    expect(codes(v, 'info')).toEqual(expect.arrayContaining(['cf-grown', 'dv-grown']));
    const cf = v.items.find((i) => i.code === 'cf-grown')!;
    expect(cf.level).toBe('info');
    expect(cf.examples).toEqual(['D2:D4 → D2:E4（広がった先: 見出しのない列）']);
    expect(v.items.find((i) => i.code === 'dv-grown')!.examples).toEqual(['D2:D4 → D2:E4（広がった先: 見出しのない列）']);
    expect(v.stats.ruleGrowth).toEqual({ cf: 1, dv: 1 });
  });

  it('also passes when the new column already got a header (typed by the user) and names it in the info', () => {
    const v = checkOutput(base(), withRules(5, 'D2:E4', 'D2:E4', { E1: 'メモ' }));
    expect(codes(v, 'fail')).toEqual([]);
    expect(v.items.find((i) => i.code === 'cf-grown')!.examples![0]).toContain('メモ');
  });

  it('an unchanged rule is not reported as grown', () => {
    const v = checkOutput(base(), withRules(5, 'D2:D4', 'D2:D4'));
    expect(codes(v, 'fail')).toEqual([]);
    expect(codes(v, 'info')).not.toEqual(expect.arrayContaining(['cf-grown']));
    expect(v.stats.ruleGrowth).toEqual({ cf: 0, dv: 0 });
  });

  it('fails when the grown rule no longer covers the column with the original header', () => {
    // the rule was moved onto the new column E only: 志望度 (D) lost it
    const v = checkOutput(base(), withRules(5, 'E2:E4', 'D2:E4'));
    const f = v.items.filter((i) => i.level === 'fail').map((i) => i.code);
    expect(f).toContain('cf');
    expect(f).not.toContain('dv'); // the DV still covers D (and grew over E)
    expect(v.ok).toBe(false);
    // DV the other way round
    const w = checkOutput(base(), withRules(5, 'D2:E4', 'E2:E4'));
    expect(w.items.filter((i) => i.level === 'fail').map((i) => i.code)).toEqual(['dv']);
  });

  it('a rule of another kind / another formula is not accepted just because it covers the header', () => {
    const out = withRules(5, 'D2:E4', 'D2:E4');
    out.sheets![0].cf[0].type = 2; // not the same kind of rule any more
    expect(codes(checkOutput(base(), out), 'fail')).toContain('cf');
    const out2 = withRules(5, 'D2:E4', 'D2:E4');
    out2.sheets![0].dv[0].formula1 = '1,2,3'; // another list
    expect(codes(checkOutput(base(), out2), 'fail')).toContain('dv');
  });

  it('a rule that lost only a column the user deleted is info (shrunk), not a failure', () => {
    // base rule on C:D (大学+志望度); the user deleted 大学 (C) -> rule on the remaining 志望度
    const b = withRules(4, 'C2:D4', null);
    const out = book([sheet({ name: '名簿', cells: { A1: '学生ID', B1: '氏名', C1: '志望度', A2: '000100', B2: '佐藤', C2: 3, A3: '000107', B3: '鈴木', C3: 5, A4: '000114', B4: '高橋', C4: 4 }, extra: { cf: [{ type: 1, appliesTo: 'C2:C4', operator: 7, formula1: '=4', formula2: '', priority: 1 }] } }), b.sheets![1]], { names: b.names });
    const v = checkOutput(b, out);
    expect(codes(v, 'fail')).not.toContain('cf');
    expect(codes(v, 'info')).toContain('cf-shrunk');
  });
});
