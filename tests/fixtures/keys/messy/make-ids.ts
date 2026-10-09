/**
 * Messy IDENTIFIER data: e-mail (case, spaces, 全角＠, mailto:, "名前 <addr>"), phone (hyphens, parentheses, +81, lost
 * leading zero), student numbers with a letter prefix, leading zeros lost / kept, 全角 digits, scientific notation,
 * numbers of different lengths.
 */
import { type CellSpec, type MessyCase, type Person, PEOPLE, build, fullName, range, shuffle } from './lib';

const order = (xs: Person[], seed: number) => shuffle(xs, seed);
export const cases: MessyCase[] = [];

// 08 ------------------------------------------------------------------------
cases.push((() => {
  const base = build('名簿', ['氏名', 'メールアドレス', '大学'], (s) => {
    for (const p of range(0, 20)) s.add(p, [fullName(p), p.email, p.univ]);
  });
  const variant = (p: Person, i: number): string => {
    switch (i % 7) {
      case 0: return `${p.email.toUpperCase()} `; // capitals + trailing space
      case 1: return `mailto:${p.email}`;
      case 2: return p.email.replace('@', '＠'); // full-width at
      case 3: return `${p.sei} ${p.mei} <${p.email}>`; // mail-client "Name <addr>"
      case 4: return ` ${p.email}。`; // leading space + 。
      case 5: return `\t${p.email}\t`;
      default: return p.email;
    }
  };
  const s1 = build('回答', ['メールアドレス', '評価'], (s) => {
    order(range(0, 16), 81).forEach((p, i) => s.add(p, [variant(p, i), 'ABC'[p.sid % 3]]));
    s.add(PEOPLE[27], ['someone.else@example.com', 'C']);
  });
  return {
    name: 'email-variants',
    note: 'the other file has ONLY e-mail addresses: capitals / trailing spaces / 全角＠ / mailto: / "名前 <addr>" / 。',
    files: [base.file('base', '名簿_メール.xlsx'), s1.file('s1', 'アンケート_メールのみ.xlsx')],
    expect: { matched: { s1: 16 }, plan: { s1: ['メール'] } },
  };
})());

// 09 ------------------------------------------------------------------------
cases.push((() => {
  const base = build('名簿', ['氏名', '電話番号', '大学'], (s) => {
    for (const p of range(0, 20)) s.add(p, [fullName(p), p.tel, p.univ]);
  });
  const digits = (p: Person) => p.tel.replace(/-/g, '');
  const variant = (p: Person, i: number): CellSpec => {
    const d = digits(p); // 09012345678
    const a = d.slice(0, 3), b = d.slice(3, 7), c = d.slice(7);
    switch (i % 10) {
      case 0: return `(${a}) ${b}-${c}`;
      case 1: return `+81 ${a.slice(1)}-${b}-${c}`;
      case 2: return `+81-(0)${a.slice(1)}-${b}-${c}`;
      case 3: return d;
      case 4: return Number(d); // the leading zero was lost (Excel number)
      case 5: return d.replace(/\d/g, (x) => String.fromCharCode(0xff10 + Number(x))).replace(/^(.{3})(.{4})/, '$1－$2－'); // ０９０－１２３４－…
      case 6: return `TEL:${p.tel}`;
      case 7: return `${a} ${b} ${c}`;
      case 8: return `${a}.${b}.${c}`;
      default: return `${p.tel}（本人）`;
    }
  };
  const s1 = build('回答', ['電話番号', '評価'], (s) => {
    order(range(0, 18), 91).forEach((p, i) => s.add(p, [variant(p, i), 'ABC'[p.sid % 3]]));
    s.add(PEOPLE[28], ['080-0000-0000', 'B']);
  });
  return {
    name: 'phone-variants',
    note: 'the other file has ONLY phone numbers: hyphens / parentheses / +81 / 全角 / number without the leading 0 / TEL: / spaces / dots',
    files: [base.file('base', '名簿_電話.xlsx'), s1.file('s1', '連絡先_電話のみ.xlsx')],
    expect: { matched: { s1: 18 }, plan: { s1: ['電話番号'] } },
  };
})());

// 10 ------------------------------------------------------------------------
cases.push((() => {
  const prefixed = (p: Person) => `S${String(p.sid).slice(0, 4)}-${String(p.sid).slice(4)}`; // S2023-0012
  const base = build('名簿', ['学籍番号', '氏名', '大学'], (s) => {
    for (const p of range(0, 20)) s.add(p, [prefixed(p), fullName(p), p.univ]);
  });
  const s1 = build('回答', ['学籍番号', '評価'], (s) => {
    for (const p of [...order(range(1, 16), 101), PEOPLE[30]]) s.add(p, [String(p.sid), 'ABC'[p.sid % 3]]); // '20230012'
  });
  const s2 = build('回答', ['学生番号', '点数'], (s) => {
    for (const p of [...order(range(3, 19), 102), PEOPLE[31]]) s.add(p, [p.sid, 50 + (p.sid % 50)]); // number cells
  });
  const s3 = build('回答', ['学籍番号', '備考'], (s) => {
    order(range(0, 12), 103).forEach((p, i) => s.add(p, [i % 2 ? prefixed(p).toLowerCase() : prefixed(p).replace(/[A-Z0-9-]/g, (c) => (c === '-' ? '－' : String.fromCharCode(c.charCodeAt(0) + 0xfee0))), null]));
  });
  return {
    name: 'id-prefix',
    note: "base 'S2023-0012'; the others '20230012' (text / number), or 's2023-0012' / 全角 'Ｓ２０２３－００１２'",
    files: [base.file('base', '名簿_S付き.xlsx'), s1.file('s1', '回答_数字のみ.xlsx'), s2.file('s2', '適性_数値.xlsx'), s3.file('s3', '回答_表記ゆれ.xlsx')],
    expect: { matched: { s1: 15, s2: 16, s3: 12 }, plan: { s1: ['ID（英字の接頭辞を除く）'], s2: ['ID（英字の接頭辞を除く）'], s3: ['ID'] } },
  };
})());

// 10b: the prefix the other way round, and a prefix that carries identity (must NOT be stripped) ----------
cases.push((() => {
  const base = build('名簿', ['学籍番号', '氏名'], (s) => {
    for (const p of range(0, 16)) s.add(p, [String(p.sid), fullName(p)]);
  });
  const s1 = build('回答', ['学籍番号', '評価'], (s) => {
    for (const p of order(range(2, 14), 104)) s.add(p, [`S${p.sid}`, 'ABC'[p.sid % 3]]);
  });
  // B/M/D = 学部/修士/博士: the same digits in different letters are DIFFERENT people
  const s2 = build('回答', ['学籍番号', '備考'], (s) => {
    s.add('x1', ['B20230001', null]).add('x2', ['M20230001', null]).add('x3', ['D20230001', null]).add('x4', ['B20230002', null]);
    s.add('x5', ['M20230002', null]).add('x6', ['D20230002', null]).add('x7', ['B20230003', null]).add('x8', ['M20230003', null]);
  });
  return {
    name: 'id-prefix-reverse-and-identity',
    note: "base '20230012', other 'S20230012' (prefix on the other side); s2: B/M/D prefixes that make different people out of the same digits",
    files: [base.file('base', '名簿_数字.xlsx'), s1.file('s1', '回答_S付き.xlsx'), s2.file('s2', '課程別.xlsx')],
    expect: { matched: { s1: 12, s2: 0 } },
  };
})());

// 11 ------------------------------------------------------------------------
cases.push((() => {
  const code = (i: number) => 100 + i * 7; // 100, 107, …
  const base = build('名簿', ['学籍番号', '氏名'], (s) => {
    for (const p of range(0, 20)) s.add(p, [{ v: code(Number(p.id.slice(1))), fmt: '000000' }, fullName(p)]); // number shown as 000107
  });
  const s1 = build('回答', ['学籍番号', '評価'], (s) => {
    order(range(0, 20), 111).forEach((p, i) => {
      const c = code(Number(p.id.slice(1)));
      const z = String(c).padStart(6, '0');
      const v: CellSpec = [z, c, z.replace(/\d/g, (x) => String.fromCharCode(0xff10 + Number(x))) + ' ', `${c} `, ` ${z}`, `'${z}`][i % 6];
      s.add(p, [v, 'ABC'[p.sid % 3]]);
    });
  });
  return {
    name: 'id-leading-zeros',
    note: 'base: number 107 shown as 000107; other: "000107" / 107 / ０００１０７ + space / "107 " / " 000107" / \'000107',
    files: [base.file('base', '名簿_ゼロ埋め表示.xlsx'), s1.file('s1', '回答_ゼロ混在.xlsx')],
    expect: { matched: { s1: 20 }, plan: { s1: ['ID'] } },
  };
})());

// 12 ------------------------------------------------------------------------
cases.push((() => {
  const big = (p: Person) => 202300000000 + Number(p.id.slice(1)) * 98765432; // 12 digits
  const sci = (n: number) => n.toExponential(5).replace('e+', 'E+'); // what Excel writes to a CSV: 2.02345E+11
  const base = build('名簿', ['受験番号', '氏名'], (s) => {
    for (const p of range(0, 20)) s.add(p, [big(p), fullName(p)]);
    s.add('p19b', [202301000000, '森　特別']); // a real person whose number is exactly what the lossy text below reads back as
  });
  // half of the rows kept their real number, the other half went through text (2.02345E+11): the ID tier is usable,
  // but the lossy cells must never read back as somebody else's number
  const s1 = build('CSV', ['受験番号', '氏名', '評価'], (s) => {
    [...order(range(2, 16), 121), PEOPLE[33]].forEach((p, i) => s.add(p, [i % 2 ? sci(big(p)) : big(p), fullName(p), 'ABC'[p.sid % 3]]));
    s.add('p33b', [sci(202301234567), '別人　太郎', 'C']); // 202301234567 → '2.02301E+11': reads back as base's 202301000000
  });
  return {
    name: 'id-scientific-notation',
    note: 'the other file went through CSV: every 12-digit number became text like 2.02345E+11 (digits lost). Must never join on those',
    files: [base.file('base', '名簿_受験番号.xlsx'), s1.file('s1', '回答_指数表記.xlsx')],
    expect: { matched: { s1: 14 }, items: ['s1:lossyId'], plan: { s1: ['ID', '氏名'] } },
  };
})());

// 13 ------------------------------------------------------------------------
cases.push((() => {
  const base = build('名簿', ['学籍番号', '氏名'], (s) => {
    for (const p of range(0, 20)) s.add(p, [String(p.sid), fullName(p)]);
  });
  // 6-digit: the last two digits of the year + the serial (230001)
  const s1 = build('回答', ['学籍番号', '氏名', '評価'], (s) => {
    for (const p of [...order(range(2, 16), 131), PEOPLE[34]]) s.add(p, [String(p.sid).slice(2), fullName(p), 'ABC'[p.sid % 3]]);
  });
  // only the 6-digit number, no name at all
  const s2 = build('回答', ['学籍番号', '点数'], (s) => {
    for (const p of [...order(range(0, 14), 132), PEOPLE[35]]) s.add(p, [String(p.sid).slice(2), 50 + (p.sid % 50)]);
  });
  return {
    name: 'id-different-lengths',
    note: '8-digit 学籍番号 in the base, 6-digit (year without the century) in the others; s2 has no name to fall back on',
    files: [base.file('base', '名簿_8桁.xlsx'), s1.file('s1', '回答_6桁と氏名.xlsx'), s2.file('s2', '点数_6桁のみ.xlsx')],
    expect: { matched: { s1: 14, s2: 14 }, plan: { s1: ['ID（桁数違い・末尾一致）', '氏名'], s2: ['ID（桁数違い・末尾一致）'] }, items: ['s2:checkWarn'] },
  };
})());

// 14 ------------------------------------------------------------------------
cases.push((() => {
  const base = build('名簿', ['氏名', '電話番号', '大学'], (s) => {
    for (const p of range(0, 20)) s.add(p, [fullName(p), p.tel, p.univ]);
  });
  const s1 = build('連絡先', ['電話番号', 'メモ'], (s) => {
    order(range(0, 16), 141).forEach((p, i) => s.add(p, [i % 2 ? `${p.tel} / 03-1111-${String(2000 + i)}` : `${p.tel}、${p.tel.replace(/^090/, '080')}`, null]));
  });
  return {
    name: 'phone-two-numbers',
    note: 'two numbers in one cell: "090-… / 03-…" and "090-…、080-…": the first valid number is the key',
    files: [base.file('base', '名簿_電話.xlsx'), s1.file('s1', '連絡先_複数番号.xlsx')],
    expect: { matched: { s1: 16 }, plan: { s1: ['電話番号'] } },
  };
})());

// 15 ------------------------------------------------------------------------
cases.push((() => {
  // the two numberings overlap by accident: 1001…1020 are DIFFERENT people in the two systems
  const base = build('名簿', ['学籍番号', '氏名', '大学'], (s) => {
    range(0, 20).forEach((p, i) => s.add(p, [1001 + i, fullName(p), p.univ]));
  });
  const s1 = build('回答', ['学籍番号', '氏名', '評価'], (s) => {
    order(range(2, 16), 151).forEach((p, i) => s.add(p, [1001 + i, fullName(p), 'ABC'[p.sid % 3]]));
  });
  return {
    name: 'id-systems-overlap-by-accident',
    note: 'both files number their people 1001, 1002 … (different systems): the ID pairs different people, the names say so',
    files: [base.file('base', '名簿_学籍番号.xlsx'), s1.file('s1', '回答_別番号体系.xlsx')],
    expect: { matched: { s1: 14 }, plan: { s1: ['氏名'] } },
  };
})());

// 16 ------------------------------------------------------------------------
cases.push((() => {
  const FAMILY = 'family.suzuki@example.ac.jp';
  const base = build('名簿', ['氏名', 'メールアドレス'], (s) => {
    for (const p of range(0, 20)) s.add(p, [fullName(p), p.id === 'p1' || p.id === 'p11' ? FAMILY : p.email]);
  });
  const s1 = build('回答', ['メールアドレス', '評価'], (s) => {
    for (const p of order(range(0, 16), 161)) s.add(p, [p.id === 'p1' || p.id === 'p11' ? FAMILY : p.email, 'ABC'[p.sid % 3]]);
  });
  return {
    name: 'email-shared-family-address',
    note: 'two people share one family address (p1, p11): that address decides nothing, everybody else joins through the e-mail',
    files: [base.file('base', '名簿_メール.xlsx'), s1.file('s1', '回答_メールのみ.xlsx')],
    expect: { matched: { s1: 14 }, surfaced: { s1: 2 } },
  };
})());

// 17 ------------------------------------------------------------------------
cases.push((() => {
  const base = build('名簿', ['学籍番号', '氏名'], (s) => {
    for (const p of range(0, 16)) s.add(p, [p.sid, fullName(p)]);
  });
  const s1 = build('回答', ['学籍番号', '評価'], (s) => {
    for (const p of order(range(0, 12), 171)) s.add(p, [String(p.sid).replace(/\B(?=(\d{3})+(?!\d))/g, ','), 'ABC'[p.sid % 3]]);
  });
  return {
    name: 'id-thousands-separator',
    note: "IDs pasted with thousands separators as text: '20,230,012'",
    files: [base.file('base', '名簿_数値ID.xlsx'), s1.file('s1', '回答_カンマ付きID.xlsx')],
    expect: { matched: { s1: 12 }, plan: { s1: ['ID'] } },
  };
})());
