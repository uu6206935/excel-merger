/**
 * Messy LAYOUT: dates of birth in every notation, paged exports (blank rows, repeated header rows, page markers,
 * subtotals), hidden columns / sheets, two-row merged headers, referral columns, bare "No." columns, files without a key.
 */
import { type CellSpec, type MessyCase, PEOPLE, type Person, Sheet, build, fullName, range, serial, shuffle, warekiOf } from './lib';

const order = (xs: Person[], seed: number) => shuffle(xs, seed);
export const cases: MessyCase[] = [];

// 14 ------------------------------------------------------------------------
cases.push((() => {
  const tanaka1 = 'n1';
  const tanaka2 = 'n2';
  const d1 = new Date(Date.UTC(2001, 4, 5));
  const d2 = new Date(Date.UTC(2002, 6, 7));
  const base = build('名簿', ['氏名', '生年月日', '大学'], (s) => {
    for (const p of range(0, 20)) s.add(p, [fullName(p), p.birth, p.univ]);
    s.add(tanaka1, ['田中　太郎', d1, '東京大学']);
    s.add(tanaka2, ['田中　太郎', d2, '早稲田大学']);
  });
  const styles = ['H13.4.1', '平成13年4月1日', 'h13/04/01'] as const;
  const s1 = build('回答', ['氏名', '生年月日', '評価'], (s) => {
    order(range(0, 16), 141).forEach((p, i) => {
      const y = p.birth.getUTCFullYear();
      const v: CellSpec = i % 4 === 0 ? `${y}/${p.birth.getUTCMonth() + 1}/${p.birth.getUTCDate()}` : warekiOf(p.birth, styles[i % 3]);
      s.add(p, [fullName(p), v, 'ABC'[p.sid % 3]]);
    });
    s.add(tanaka2, ['田中　太郎', warekiOf(d2, '平成13年4月1日'), 'B']);
    s.add(tanaka1, ['田中　太郎', warekiOf(d1, 'H13.4.1'), 'A']);
  });
  const s2 = build('回答', ['氏名', '生年月日', '点数'], (s) => {
    order(range(2, 18), 142).forEach((p, i) => s.add(p, [fullName(p), i % 2 ? serial(p.birth) : Number(`${p.birth.getUTCFullYear()}${String(p.birth.getUTCMonth() + 1).padStart(2, '0')}${String(p.birth.getUTCDate()).padStart(2, '0')}`), 50 + (p.sid % 50)]));
    s.add(tanaka1, ['田中　太郎', serial(d1), 70]);
    s.add(tanaka2, ['田中　太郎', serial(d2), 80]);
  });
  const s3 = build('回答', ['氏名', '生年月日', 'メモ'], (s) => {
    for (const p of order(range(0, 10), 143)) s.add(p, [fullName(p), `${p.birth.getUTCFullYear()}年${p.birth.getUTCMonth() + 1}月${p.birth.getUTCDate()}日生`, null]);
    s.add(tanaka2, ['田中　太郎', `${d2.getUTCFullYear()}年${d2.getUTCMonth() + 1}月${d2.getUTCDate()}日生`, null]);
  });
  return {
    name: 'birthdate-notations',
    note: '生年月日 as 和暦 (H13.4.1 / 平成13年4月1日 / h13/04/01), 2001/4/1 text, serial numbers, 20010401, "2001年4月1日生"; two different 田中 太郎',
    files: [base.file('base', '名簿_生年月日.xlsx'), s1.file('s1', '回答_和暦.xlsx'), s2.file('s2', '回答_シリアル値.xlsx'), s3.file('s3', '回答_生まれ.xlsx')],
    expect: { matched: { s1: 18, s2: 18, s3: 11 } },
  };
})());

// 15 ------------------------------------------------------------------------
function paged(people: Person[], pageSize: number, title: boolean): Sheet {
  const s = new Sheet('一覧');
  const header = ['氏名', '大学', '評価'];
  if (title) s.raw(['応募者一覧（2026年度）']).raw([]);
  s.raw(header);
  const forms = ['小計（ページ計）', '東京大学　小計', 'Subtotal', '◆ページ計'];
  let n = 0;
  people.forEach((p, i) => {
    s.add(p, [fullName(p), p.univ, 50 + (p.sid % 50)]);
    n++;
    if ((i + 1) % pageSize === 0 && i + 1 < people.length) {
      s.raw([forms[Math.floor(i / pageSize) % forms.length], null, n]);
      s.raw([]).raw([]);
      s.raw(['　　― ' + (Math.floor(i / pageSize) + 2) + ' ―']);
      s.raw(header);
      n = 0;
    }
  });
  s.raw(['合計', null, people.length]);
  return s;
}

cases.push((() => {
  const base = build('名簿', ['氏名', '大学', '学部'], (s) => {
    for (const p of range(0, 24)) s.add(p, [fullName(p), p.univ, p.fac]);
  });
  const s1 = paged([...order(range(2, 20), 151), ...range(24, 27)], 5, true);
  return {
    name: 'paged-export-as-other',
    note: 'paged export: title, blank rows, the header repeated on every page, page markers "― 2 ―", subtotal rows in four notations, 合計; joined with "full"',
    files: [base.file('base', '名簿.xlsx'), s1.file('s1', '応募者一覧_ページ出力.xlsx')],
    plan: { joinType: 'full' },
    expect: { matched: { s1: 18 }, appended: 3 },
  };
})());

cases.push((() => {
  const base = paged(range(0, 18), 6, true);
  const s1 = build('回答', ['氏名', '面接評価'], (s) => {
    for (const p of order(range(0, 18), 152)) s.add(p, [fullName(p), 'ABC'[p.sid % 3]]);
  });
  return {
    name: 'paged-export-as-base',
    note: 'the BASE is the paged export: the repeated header / marker / subtotal rows must not become base rows',
    files: [base.file('base', '応募者一覧_ページ出力.xlsx'), s1.file('s1', '面接結果.xlsx')],
    expect: { matched: { s1: 18 } },
  };
})());

// 16 ------------------------------------------------------------------------
cases.push((() => {
  const base = build('名簿', ['氏名', '内部ID', '大学'], (s) => {
    for (const p of range(0, 20)) s.add(p, [fullName(p), `X-${String(p.sid).slice(4)}`, p.univ]);
    s.hide(2);
  });
  const old = build('旧データ', ['氏名', '内部ID'], (s) => {
    for (const p of range(20, 30)) s.add(p, [fullName(p), `X-${String(p.sid).slice(4)}`]);
  }).spec();
  old.state = 'hidden';
  const s1 = build('回答', ['内部ID', '評価'], (s) => {
    for (const p of [...order(range(0, 14), 161), PEOPLE[25]]) s.add(p, [`X-${String(p.sid).slice(4)}`, 'ABC'[p.sid % 3]]);
  });
  const f = base.file('base', '名簿_非表示列.xlsx', [old]);
  // the hidden sheet comes FIRST in the workbook (the default active sheet is hidden)
  f.sheets = [old, f.sheets[0]];
  f.dataSheet = '名簿';
  return {
    name: 'hidden-column-and-sheet',
    note: 'a hidden 内部ID column carries the key; a hidden sheet 旧データ is the first sheet of the workbook',
    files: [f, s1.file('s1', '回答_内部ID.xlsx')],
    expect: { matched: { s1: 14 }, plan: { s1: ['ID'] } },
  };
})());

// 17 ------------------------------------------------------------------------
cases.push((() => {
  const base = build('名簿', ['学籍番号', '氏名', '大学'], (s) => {
    for (const p of range(0, 20)) s.add(p, [String(p.sid), fullName(p), p.univ]);
  });
  // A: keys merged over both header rows
  const a = new Sheet('面接評価');
  a.raw(['学籍番号', '氏名', '面接評価', null, null, '備考']);
  a.raw([null, null, '評価1', '評価2', '評価3', null]);
  a.merge(1, 1, 2, 1).merge(1, 2, 2, 2).merge(1, 3, 1, 5).merge(1, 6, 2, 6);
  for (const p of order(range(2, 16), 171)) a.add(p, [String(p.sid), fullName(p), 3 + (p.sid % 3), 3 + ((p.sid + 1) % 3), 3 + ((p.sid + 2) % 3), null]);
  // B: only the group title is merged; the key headers sit in the first header row, the second row has blanks under them
  const b = new Sheet('面接評価');
  b.raw(['学籍番号', '氏名', '面接評価', null, null]);
  b.raw([null, null, '評価1', '評価2', '評価3']);
  b.merge(1, 3, 1, 5);
  for (const p of order(range(0, 12), 172)) b.add(p, [String(p.sid), fullName(p), 3 + (p.sid % 3), 3 + ((p.sid + 1) % 3), 3 + ((p.sid + 2) % 3)]);
  return {
    name: 'two-row-merged-header',
    note: '「面接評価」 merged over 評価1/評価2/評価3; in A the key headers are merged over both rows, in B they are not',
    files: [base.file('base', '名簿.xlsx'), a.file('s1', '面接評価_結合ヘッダーA.xlsx'), b.file('s2', '面接評価_結合ヘッダーB.xlsx')],
    expect: { matched: { s1: 14, s2: 12 }, plan: { s1: ['ID＋氏名', 'ID', '氏名'], s2: ['ID＋氏名', 'ID', '氏名'] } },
  };
})());

// 18 ------------------------------------------------------------------------
cases.push((() => {
  const base = build('名簿', ['氏名', '大学'], (s) => {
    for (const p of range(0, 20)) s.add(p, [fullName(p), p.univ]);
  });
  // a referral list: the only names are the REFERRERS (people of the base) and the people they referred (not in the base)
  const s1 = build('紹介', ['紹介者', '紹介された人', 'ポイント'], (s) => {
    for (let i = 0; i < 12; i++) {
      const r = PEOPLE[(i * 5) % 20];
      const x = PEOPLE[20 + i];
      s.add(`ref${i}`, [fullName(r), fullName(x), 10 + i]);
    }
  });
  return {
    name: 'referral-columns',
    note: 'the other file has only 紹介者 (names of base people) and 紹介された人: a referrer is not the person of the row',
    files: [base.file('base', '名簿.xlsx'), s1.file('s1', '紹介制度.xlsx')],
    expect: { matched: { s1: 0 }, items: ['s1:noKey'] },
  };
})());

// 31 ------------------------------------------------------------------------
cases.push((() => {
  const base = build('名簿', ['No', '氏名', '大学'], (s) => {
    range(0, 20).forEach((p, i) => s.add(p, [i + 1, fullName(p), p.univ]));
  });
  const s1 = build('回答', ['No.', '氏名', '評価'], (s) => {
    order(range(2, 16), 311).forEach((p, i) => s.add(p, [i + 1, fullName(p), 'ABC'[p.sid % 3]]));
  });
  // a score ranking: "No" is the rank in a different order, no name at all
  const s2 = build('回答', ['No', '評価'], (s) => {
    order(range(0, 14), 312).forEach((p, i) => s.add(p, [((i * 5) % 14) + 1, 'ABC'[p.sid % 3]]));
  });
  return {
    name: 'bare-no-columns',
    note: "'No' columns are row counters of each file (1, 2, 3 …), never the same person across files; s2 has nothing else",
    files: [base.file('base', '名簿_No付き.xlsx'), s1.file('s1', '回答_No付き.xlsx'), s2.file('s2', '順位表.xlsx')],
    expect: { matched: { s1: 14, s2: 0 }, plan: { s1: ['氏名'] }, items: ['s2:noKey'], unsurfaced: { s2: 14 } },
  };
})());

// 32 ------------------------------------------------------------------------
cases.push((() => {
  const base = build('名簿', ['氏名', '大学', '学部'], (s) => {
    for (const p of range(0, 20)) s.add(p, [fullName(p), p.univ, p.fac]);
  });
  const s1 = build('集計', ['大学', '学部', '評価'], (s) => {
    range(0, 10).forEach((p, i) => s.add(`u${i}`, [p.univ, p.fac, 'ABC'[i % 3]]));
  });
  const s2 = build('集計', ['学部', '性別', '評価'], (s) => {
    for (let i = 0; i < 8; i++) s.add(`g${i}`, [range(0, 8)[i].fac, i % 2 ? '男' : '女', 'ABC'[i % 3]]);
  });
  return {
    name: 'no-person-key-at-all',
    note: 'other files have only categories (大学 / 学部 / 性別): joining people on those would fan one row out over many people',
    files: [base.file('base', '名簿.xlsx'), s1.file('s1', '大学別集計.xlsx'), s2.file('s2', '学部別集計.xlsx')],
    expect: { matched: { s1: 0, s2: 0 }, items: ['s1:noKey', 's2:noKey'] },
  };
})());

// 19 ------------------------------------------------------------------------
cases.push((() => {
  // the students are grouped by university and the 大学 cell is MERGED over the rows of its group: only the first row has the text
  const groups: Array<[string, Person[]]> = [['東京大学', range(0, 4)], ['早稲田大学', range(4, 8)], ['京都大学', range(8, 12)]];
  const base = new Sheet('名簿');
  base.raw(['氏名', '大学', '学部']);
  let row = 1;
  for (const [u, ps] of groups) {
    const r1 = row + 1;
    ps.forEach((p, i) => {
      base.add(p, [fullName(p), i === 0 ? u : null, p.fac]);
      row++;
    });
    base.merge(r1, 2, row, 2);
  }
  // two namesakes in different universities (merged groups of one row each): only 氏名＋大学 tells them apart
  base.add('d1', ['田中　太郎', '東京大学', '法学部']).add('d2', ['田中　太郎', '早稲田大学', '商学部']);
  const flat = build('回答', ['氏名', '大学', '評価'], (s) => {
    for (const [u, ps] of groups) for (const p of ps) s.add(p, [fullName(p), u, 'ABC'[p.sid % 3]]);
    s.add('d1', ['田中　太郎', '東京大学', 'A']);
    s.add('d2', ['田中　太郎', '早稲田大学', 'B']);
  });
  return {
    name: 'vertical-merged-university',
    note: '大学 merged over the rows of each group (only the first row carries the text); two 田中 太郎 need 氏名＋大学',
    files: [base.file('base', '名簿_大学結合セル.xlsx'), flat.file('s1', '回答_大学付き.xlsx')],
    expect: { matched: { s1: 14 } },
  };
})());

// 20 ------------------------------------------------------------------------
cases.push((() => {
  const s = new Sheet('一覧');
  s.raw(['氏名', '大学', '評価']);
  const people = order(range(0, 18), 201);
  people.slice(0, 6).forEach((p) => s.add(p, [fullName(p), p.univ, 'A']));
  s.raw(['※ 以下は2次募集分']);
  people.slice(6, 12).forEach((p) => s.add(p, [fullName(p), p.univ, 'B']));
  s.raw(['---- 辞退者 ----']);
  people.slice(12).forEach((p) => s.add(p, [fullName(p), p.univ, 'C']));
  const base = build('名簿', ['氏名', '大学'], (b) => {
    for (const p of range(0, 24)) b.add(p, [fullName(p), p.univ]);
  });
  return {
    name: 'remark-lines-inside-list',
    note: 'remark lines typed between the rows: "※ 以下は2次募集分", "---- 辞退者 ----"; joined with "full" (they must not become people)',
    files: [base.file('base', '名簿.xlsx'), s.file('s1', '応募状況_区切り行あり.xlsx')],
    plan: { joinType: 'full' },
    expect: { matched: { s1: 18 }, appended: 0 },
  };
})());

// 21 ------------------------------------------------------------------------
cases.push((() => {
  // a list pasted in two pieces with 5 blank rows between them: the table ends at the gap, the rest is NOT silently dropped
  const s = new Sheet('名簿');
  s.raw(['氏名', '大学', '評価']);
  const people = order(range(0, 18), 211);
  people.slice(0, 10).forEach((p) => s.add(p, [fullName(p), p.univ, 'A']));
  for (let i = 0; i < 5; i++) s.raw([]);
  people.slice(10).forEach((p) => s.add(p, [fullName(p), p.univ, 'B']));
  const base = build('名簿', ['氏名', '大学'], (b) => {
    for (const p of range(0, 18)) b.add(p, [fullName(p), p.univ]);
  });
  return {
    name: 'list-in-two-pieces',
    note: 'the other file is one list in two pieces (5 blank rows between them): only the first piece is the table, the second is reported',
    files: [base.file('base', '名簿.xlsx'), s.file('s1', '応募状況_二分割.xlsx')],
    expect: { matched: { s1: 10 }, items: ['s1:dataBelow'], unsurfaced: { s1: 8 } },
  };
})());

// 22 ------------------------------------------------------------------------
cases.push((() => {
  // an application form template: the row under the header shows a sample entry ("例）山田 太郎") — a sample is nobody
  const s = new Sheet('応募フォーム');
  s.raw(['No', '氏名', '大学', '志望']);
  s.raw(['例）', fullName(PEOPLE[0]), '東京大学', '営業']);
  order(range(1, 12), 221).forEach((p, i) => s.add(p, [i + 1, fullName(p), p.univ, '営業']));
  const base = build('名簿', ['氏名', '大学'], (b) => {
    for (const p of range(0, 14)) b.add(p, [fullName(p), p.univ]);
  });
  return {
    name: 'form-template-example-row',
    note: 'the row under the header is a sample entry (例）山田 太郎): the real 山田 太郎 is NOT in this file and must not be joined to it',
    files: [base.file('base', '名簿.xlsx'), s.file('s1', '応募フォーム_記入例付き.xlsx')],
    plan: { joinType: 'full' },
    expect: { matched: { s1: 11 }, appended: 0 },
  };
})());
