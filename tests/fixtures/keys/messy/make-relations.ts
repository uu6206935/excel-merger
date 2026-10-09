/**
 * Messy RELATIONS between the files: typos in every name, 同姓同名, one-to-many (both directions), a file with more rows
 * than the base joined with "full", two files without an ID that only share 氏名＋大学, every file keyed differently.
 */
import { type MessyCase, PEOPLE, type Person, Sheet, build, fullKana, fullName, range, shuffle } from './lib';

const order = (xs: Person[], seed: number) => shuffle(xs, seed);
export const cases: MessyCase[] = [];

// 19 ------------------------------------------------------------------------
cases.push((() => {
  const TYPO: Record<string, string> = {
    p0: '山田　太朗', p1: '鈴木　健大', p2: '佐藤　結', p7: '小林　華子', p8: '加藤　美早', p9: '吉田　大暉', p14: '林　一朗', p17: '山崎　択也',
  };
  const base = build('名簿', ['氏名', 'フリガナ', '大学'], (s) => {
    for (const p of range(0, 25)) s.add(p, [fullName(p), fullKana(p), p.univ]);
  });
  const s1 = build('回答', ['氏名', '評価'], (s) => {
    for (const p of order(range(0, 25), 191)) s.add(p, [TYPO[p.id] ?? fullName(p), 'ABC'[p.sid % 3]]);
  });
  // s2: correct kana, names with a typo in 3 rows (joined through フリガナ and flagged) and kana typos in 3 other rows
  const KANA_TYPO: Record<string, string> = { p10: 'ヤマグチ　ユウカ', p11: 'マツモト　ユウマ'.replace('ユウマ', 'ユウメ'), p12: 'イノウエ　アヤヤ' };
  const NAME_TYPO: Record<string, string> = { p3: '高橋　連', p4: '渡辺　さくらん', p5: '伊藤　直樹樹' };
  const s2 = build('回答', ['氏名', 'フリガナ', '点数'], (s) => {
    for (const p of order(range(0, 20), 192)) s.add(p, [NAME_TYPO[p.id] ?? fullName(p), KANA_TYPO[p.id] ?? fullKana(p), 50 + (p.sid % 50)]);
  });
  return {
    name: 'typos-never-silent',
    note: 'every name is unique but 8 rows have a one-character typo: they must NOT be joined silently (suggestions); s2: typos in names but correct kana (joined via フリガナ, flagged), typos in kana only',
    files: [base.file('base', '名簿.xlsx'), s1.file('s1', '回答_誤字あり.xlsx'), s2.file('s2', '適性_誤字とカナ.xlsx')],
    expect: { matched: { s1: 17, s2: 17 }, surfaced: { s1: 8, s2: 3 } },
  };
})());

// 20 ------------------------------------------------------------------------
cases.push((() => {
  const t1 = 'n1';
  const t2 = 'n2';
  const base = build('名簿', ['氏名', '大学', '学部'], (s) => {
    for (const p of range(0, 20)) s.add(p, [fullName(p), p.univ, p.fac]);
    s.add(t1, ['田中　太郎', '東京大学', '法学部']);
    s.add(t2, ['田中　太郎', '早稲田大学', '商学部']);
  });
  const s1 = build('回答', ['氏名', '大学', '評価'], (s) => {
    s.add(t2, ['田中　太郎', '早稲田大学', 'B']);
    for (const p of order(range(0, 16), 201)) s.add(p, [fullName(p), p.univ, 'ABC'[p.sid % 3]]);
    s.add(t1, ['田中　太郎', '東京大学', 'A']);
  });
  const s2 = build('回答', ['氏名', '評価'], (s) => {
    for (const p of order(range(0, 16), 202)) s.add(p, [fullName(p), 'ABC'[p.sid % 3]]);
    s.add(t2, ['田中　太郎', 'B']);
    s.add(t1, ['田中　太郎', 'A']);
  });
  // an extra namesake from a third university that is not in the base; only ONE base namesake answered
  const s3 = build('回答', ['氏名', '大学', '評価'], (s) => {
    for (const p of order(range(0, 10), 203)) s.add(p, [fullName(p), p.univ, 'ABC'[p.sid % 3]]);
    s.add(t2, ['田中　太郎', '早稲田大学', 'B']);
    s.add('n3', ['田中　太郎', '京都大学', 'C']);
  });
  return {
    name: 'namesakes-university',
    note: 'two different 田中 太郎 in the base (東京大学 / 早稲田大学). s1 has both with 大学 (order reversed), s2 has no 大学 (must stay ambiguous), s3 has one of them plus a third namesake',
    files: [base.file('base', '名簿_同姓同名.xlsx'), s1.file('s1', '回答_大学付き.xlsx'), s2.file('s2', '回答_大学なし.xlsx'), s3.file('s3', '回答_別の田中さん.xlsx')],
    expect: { matched: { s1: 18, s2: 16, s3: 11 }, surfaced: { s2: 2 }, plan: { s1: ['氏名＋大学', '氏名'] } },
  };
})());

cases.push((() => {
  const t1 = 'n1';
  const t2 = 'n2';
  const d1 = new Date(Date.UTC(2001, 4, 5));
  const d2 = new Date(Date.UTC(2001, 4, 6));
  const base = build('名簿', ['氏名', '大学', '学部', '生年月日'], (s) => {
    for (const p of range(0, 16)) s.add(p, [fullName(p), p.univ, p.fac, p.birth]);
    s.add(t1, ['伊藤　翔', '東京大学', '法学部', d1]);
    s.add(t2, ['伊藤　翔', '東京大学', '法学部', d2]);
  });
  const s1 = build('回答', ['氏名', '大学', '学部', '生年月日', '評価'], (s) => {
    for (const p of order(range(0, 14), 204)) s.add(p, [fullName(p), p.univ, p.fac, p.birth, 'ABC'[p.sid % 3]]);
    s.add(t2, ['伊藤　翔', '東京大学', '法学部', d2, 'B']);
    s.add(t1, ['伊藤　翔', '東京大学', '法学部', d1, 'A']);
  });
  const s2 = build('回答', ['氏名', '大学', '学部', '評価'], (s) => {
    for (const p of order(range(0, 14), 205)) s.add(p, [fullName(p), p.univ, p.fac, 'ABC'[p.sid % 3]]);
    s.add(t2, ['伊藤　翔', '東京大学', '法学部', 'B']);
    s.add(t1, ['伊藤　翔', '東京大学', '法学部', 'A']);
  });
  return {
    name: 'namesakes-same-university-faculty',
    note: 'two different 伊藤 翔: same 大学 and 学部, only the 生年月日 differs. s1 carries it, s2 does not (→ ambiguous, never guessed)',
    files: [base.file('base', '名簿_同姓同名同学部.xlsx'), s1.file('s1', '回答_生年月日付き.xlsx'), s2.file('s2', '回答_生年月日なし.xlsx')],
    expect: { matched: { s1: 16, s2: 14 }, surfaced: { s2: 2 } },
  };
})());

// 21 ------------------------------------------------------------------------
function interviewHistory(seed: number): { base: Sheet; hist: Sheet; rowsOf: Record<string, number[]> } {
  const base = build('名簿', ['学籍番号', '氏名', '大学'], (s) => {
    for (const p of range(0, 20)) s.add(p, [String(p.sid), fullName(p), p.univ]);
  });
  const rowsOf: Record<string, number[]> = {};
  const hist = build('面接履歴', ['学籍番号', '氏名', '面接日', '評価'], (s) => {
    let n = 0;
    for (const p of order(range(0, 15), seed)) {
      const times = 1 + (Number(p.id.slice(1)) % 3); // 1, 2 or 3 interviews
      for (let k = 0; k < times; k++) {
        s.add(p, [String(p.sid), fullName(p), new Date(Date.UTC(2026, 5, 1 + k * 7 + (n % 5))), 'ABC'[(p.sid + k) % 3]]);
        (rowsOf[p.id] ??= []).push(s.rows.length);
        n++;
      }
    }
  });
  return { base, hist, rowsOf };
}
export const INTERVIEW_ROWS: Record<string, number[]> = interviewHistory(211).rowsOf;
for (const policy of ['first', 'last', 'skip'] as const) {
  cases.push((() => {
    const { base, hist } = interviewHistory(211);
    return {
      name: `one-to-many-${policy}`,
      note: `the base has one row per person, the other file 1–3 interview rows per person (policy ${policy})`,
      files: [base.file('base', '名簿.xlsx'), hist.file('s1', '面接履歴.xlsx')],
      plan: { duplicatePolicy: policy },
      expect: policy === 'skip' ? { matched: { s1: 5 }, surfaced: { s1: 10 } } : { matched: { s1: 15 }, items: ['s1:duplicate'] },
    };
  })());
}

// 22 ------------------------------------------------------------------------
cases.push((() => {
  const base = build('応募履歴', ['学籍番号', '氏名', '企業', '結果'], (s) => {
    for (const p of range(0, 12)) {
      const k = 1 + (Number(p.id.slice(1)) % 4);
      for (let c = 0; c < k; c++) s.add(p, [String(p.sid), fullName(p), `企業${String.fromCharCode(65 + c)}`, 'ABC'[(p.sid + c) % 3]]);
    }
  });
  const s1 = build('名簿', ['学籍番号', '氏名', '大学', '電話番号'], (s) => {
    for (const p of order(range(0, 20), 221)) s.add(p, [String(p.sid), fullName(p), p.univ, p.tel]);
  });
  return {
    name: 'one-to-many-lookup-by-id',
    note: 'the BASE has several rows per person (応募履歴), the other file is a person master with one row each → VLOOKUP-style',
    files: [base.file('base', '応募履歴.xlsx'), s1.file('s1', '学生名簿.xlsx')],
    expect: { matched: { s1: 30 }, plan: { s1: ['ID', '氏名'] }, noItems: ['s1:duplicateBase'] },
  };
})());

cases.push((() => {
  // no IDs: a person who applied twice is indistinguishable from two namesakes by the name alone
  const base = build('応募履歴', ['氏名', '企業'], (s) => {
    for (const p of range(0, 10)) s.add(p, [fullName(p), '企業A']);
    for (const p of range(0, 3)) s.add(p, [fullName(p), '企業B']);
  });
  const s1 = build('名簿', ['氏名', '大学'], (s) => {
    for (const p of order(range(0, 14), 222)) s.add(p, [fullName(p), p.univ]);
  });
  return {
    name: 'one-to-many-name-only',
    note: 'base rows repeat the same name (3 of 10 people applied twice = 46% of the rows repeat) and there is no ID: a VLOOKUP-style join, reported as such',
    files: [base.file('base', '応募履歴_ID無し.xlsx'), s1.file('s1', '名簿_大学.xlsx')],
    expect: { matched: { s1: 13 }, items: ['s1:lookup'] },
  };
})());

cases.push((() => {
  // no IDs, ONE person applied twice (2 of 11 rows repeat): the name alone cannot tell a repeat from 同姓同名
  const base = build('応募履歴', ['氏名', '企業'], (s) => {
    for (const p of range(0, 10)) s.add(p, [fullName(p), '企業A']);
    s.add(PEOPLE[0], [fullName(PEOPLE[0]), '企業B']);
  });
  const s1 = build('名簿', ['氏名', '大学'], (s) => {
    for (const p of order(range(0, 14), 223)) s.add(p, [fullName(p), p.univ]);
  });
  return {
    name: 'one-to-many-name-only-few',
    note: 'a person listed twice in the base without any ID: the two rows are ambiguous (flagged), everybody else joins',
    files: [base.file('base', '応募履歴_重複少.xlsx'), s1.file('s1', '名簿_大学.xlsx')],
    expect: { matched: { s1: 9 }, surfaced: { s1: 2 }, items: ['s1:duplicateBase'] },
  };
})());

// 23 ------------------------------------------------------------------------
cases.push((() => {
  const base = build('名簿', ['姓', '名', '大学'], (s) => {
    for (const p of range(0, 10)) s.add(p, [p.sei, p.mei, p.univ]);
  });
  const s1 = build('回答', ['氏名', '大学', '評価'], (s) => {
    const rows: Array<[Person, string]> = [];
    for (const p of order(range(0, 10), 231)) rows.push([p, fullName(p)]);
    for (const p of range(20, 32)) rows.push([p, p.id === 'p25' ? `${p.sei}${p.mei}` : fullName(p)]); // p25 without a space
    for (const p of range(20, 23)) rows.push([p, fullName(p)]); // submitted twice
    shuffle(rows, 232).forEach(([p, nm]) => s.add(p, [nm, p.univ, 'ABC'[p.sid % 3]]));
  });
  return {
    name: 'more-rows-than-base-full',
    note: 'the other file has 25 rows for a 10-person base: 10 known people, 12 new ones (3 of them twice), one new name without a space; joined with "full"',
    files: [base.file('base', '名簿_姓名別.xlsx'), s1.file('s1', '回答_全員.xlsx')],
    plan: { joinType: 'full' },
    expect: { matched: { s1: 10 }, appended: 12 },
  };
})());

// 24 ------------------------------------------------------------------------
cases.push((() => {
  const t1 = 'n1';
  const t2 = 'n2';
  const base = build('名簿', ['学籍番号', '氏名', '大学'], (s) => {
    for (const p of range(0, 16)) s.add(p, [String(p.sid), fullName(p), p.univ]);
    s.add(t1, ['20239901', '田中　太郎', '東京大学']);
    s.add(t2, ['20239902', '田中　太郎', '早稲田大学']);
  });
  const extras = range(20, 24);
  const s1 = build('面接', ['氏名', '大学', '面接評価'], (s) => {
    for (const p of order(range(0, 14), 241)) s.add(p, [fullName(p), p.univ, 'ABC'[p.sid % 3]]);
    s.add(t2, ['田中　太郎', '早稲田大学', 'B']);
    s.add(t1, ['田中　太郎', '東京大学', 'A']);
    for (const p of extras) s.add(p, [fullName(p), p.univ, 'C']);
  });
  const s2 = build('適性', ['氏名', '大学', '点数'], (s) => {
    for (const p of order(range(2, 16), 242)) s.add(p, [fullName(p), p.univ, 50 + (p.sid % 50)]);
    s.add(t1, ['田中　太郎', '東京大学', 61]);
    s.add(t2, ['田中　太郎', '早稲田大学', 72]);
    for (const p of extras.slice(0, 2)) s.add(p, [fullName(p), p.univ, 80]);
    s.add('x99', ['新井　誠', '明治大学', 55]);
  });
  return {
    name: 'two-files-without-id-full',
    note: 'two files without an ID share only 氏名＋大学 (a 同姓同名 pair exists); joined with "full": s2 must find the rows s1 appended',
    files: [base.file('base', '名簿_ID付き.xlsx'), s1.file('s1', '面接_ID無し.xlsx'), s2.file('s2', '適性_ID無し.xlsx')],
    plan: { joinType: 'full' },
    expect: { matched: { s1: 16, s2: 16 }, appended: 5 },
  };
})());

// 26 ------------------------------------------------------------------------
cases.push((() => {
  const base = build('名簿', ['学籍番号', '氏名', 'フリガナ', 'メールアドレス', '電話番号', '大学'], (s) => {
    for (const p of range(0, 20)) s.add(p, [String(p.sid), fullName(p), fullKana(p), p.email, p.tel, p.univ]);
  });
  const s1 = build('回答', ['学籍番号', '評価'], (s) => {
    for (const p of order(range(0, 14), 261)) s.add(p, [`S${String(p.sid).slice(0, 4)}-${String(p.sid).slice(4)}`, 'ABC'[p.sid % 3]]);
  });
  const s2 = build('回答', ['お名前', '点数'], (s) => {
    for (const p of order(range(2, 16), 262)) s.add(p, [`${p.sei}${p.mei}様`, 50 + (p.sid % 50)]);
  });
  const s3 = build('回答', ['Eメール', 'コメント'], (s) => {
    for (const p of order(range(4, 18), 263)) s.add(p, [` ${p.email.toUpperCase()}`, null]);
  });
  const s4 = build('回答', ['携帯', '備考'], (s) => {
    for (const p of order(range(6, 20), 264)) s.add(p, [p.tel.replace(/-/g, ''), null]);
  });
  const s5 = build('回答', ['フリガナ', '区分'], (s) => {
    for (const p of order(range(0, 12), 265)) s.add(p, [`${p.seiKana}${p.meiKana}`, '学部']);
  });
  return {
    name: 'every-file-keyed-differently',
    note: 'one base (ID/氏名/メール/電話), five files each keyed by something else: prefixed ID, 氏名+様, 大文字メール, 携帯, フリガナ',
    files: [base.file('base', '名簿_全項目.xlsx'), s1.file('s1', '回答_ID.xlsx'), s2.file('s2', '回答_お名前.xlsx'), s3.file('s3', '回答_メール.xlsx'), s4.file('s4', '回答_携帯.xlsx'), s5.file('s5', '回答_フリガナ.xlsx')],
    expect: { matched: { s1: 14, s2: 14, s3: 14, s4: 14, s5: 12 } },
  };
})());
