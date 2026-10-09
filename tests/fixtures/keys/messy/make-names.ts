/**
 * Messy NAME data: 姓/名 split vs 氏名 combined, spaces, フリガナ scripts, English name orders, 異体字 / 旧姓 / 読み仮名 in
 * parentheses, 敬称.  Every builder is deterministic (people come from lib.ts).
 */
import { type MessyCase, PEOPLE, type Person, build, fullKana, fullName, range, shuffle } from './lib';

// ---- half-width katakana / hiragana converters (fixtures only) ----
const HALF: Record<string, string> = {
  ア: 'ｱ', イ: 'ｲ', ウ: 'ｳ', エ: 'ｴ', オ: 'ｵ', カ: 'ｶ', キ: 'ｷ', ク: 'ｸ', ケ: 'ｹ', コ: 'ｺ', サ: 'ｻ', シ: 'ｼ', ス: 'ｽ', セ: 'ｾ', ソ: 'ｿ',
  タ: 'ﾀ', チ: 'ﾁ', ツ: 'ﾂ', テ: 'ﾃ', ト: 'ﾄ', ナ: 'ﾅ', ニ: 'ﾆ', ヌ: 'ﾇ', ネ: 'ﾈ', ノ: 'ﾉ', ハ: 'ﾊ', ヒ: 'ﾋ', フ: 'ﾌ', ヘ: 'ﾍ', ホ: 'ﾎ',
  マ: 'ﾏ', ミ: 'ﾐ', ム: 'ﾑ', メ: 'ﾒ', モ: 'ﾓ', ヤ: 'ﾔ', ユ: 'ﾕ', ヨ: 'ﾖ', ラ: 'ﾗ', リ: 'ﾘ', ル: 'ﾙ', レ: 'ﾚ', ロ: 'ﾛ', ワ: 'ﾜ', ヲ: 'ｦ', ン: 'ﾝ',
  ッ: 'ｯ', ャ: 'ｬ', ュ: 'ｭ', ョ: 'ｮ', ー: 'ｰ', '　': ' ',
};
const VOICED: Record<string, string> = {
  ガ: 'ｶ', ギ: 'ｷ', グ: 'ｸ', ゲ: 'ｹ', ゴ: 'ｺ', ザ: 'ｻ', ジ: 'ｼ', ズ: 'ｽ', ゼ: 'ｾ', ゾ: 'ｿ', ダ: 'ﾀ', ヂ: 'ﾁ', ヅ: 'ﾂ', デ: 'ﾃ', ド: 'ﾄ',
  バ: 'ﾊ', ビ: 'ﾋ', ブ: 'ﾌ', ベ: 'ﾍ', ボ: 'ﾎ',
};
export const toHalfKana = (s: string): string => [...s].map((c) => (VOICED[c] ? `${VOICED[c]}ﾞ` : HALF[c] ?? c)).join('');
export const toHira = (s: string): string => s.replace(/[ァ-ヶ]/g, (c) => String.fromCharCode(c.charCodeAt(0) - 0x60));

export const toKataOf = (s: string): string => s.replace(/[ぁ-ゖ]/g, (c) => String.fromCharCode(c.charCodeAt(0) + 0x60));
const order = (xs: Person[], seed: number) => shuffle(xs, seed);

export const cases: MessyCase[] = [];

// 01 ------------------------------------------------------------------------
cases.push((() => {
  const base = build('名簿', ['姓', '名', '大学', '学部'], (s) => {
    for (const p of range(0, 20)) s.add(p, [p.sei, p.mei, p.univ, p.fac]);
  });
  const s1 = build('回答', ['氏名', '面接評価'], (s) => {
    for (const p of [...order(range(2, 16), 11), PEOPLE[20], PEOPLE[21]]) s.add(p, [fullName(p, '　'), 'ABC'[p.sid % 3]]);
  });
  const s2 = build('回答', ['氏名', '点数'], (s) => {
    for (const p of [...order(range(5, 19), 12), PEOPLE[22]]) s.add(p, [fullName(p, ' '), 50 + (p.sid % 50)]);
  });
  const s3 = build('回答', ['氏名', 'コメント'], (s) => {
    for (const p of [...order(range(0, 10), 13), PEOPLE[23]]) s.add(p, [fullName(p, ''), '特になし']);
  });
  return {
    name: 'names-split-vs-combined',
    note: 'base has 姓/名 in two columns; the others have 氏名 with a full-width space, a half-width space, or no space',
    files: [base.file('base', '名簿_姓名別.xlsx'), s1.file('s1', '面接_全角スペース.xlsx'), s2.file('s2', '適性_半角スペース.xlsx'), s3.file('s3', '回答_スペース無し.xlsx')],
    expect: { matched: { s1: 14, s2: 14, s3: 10 }, plan: { s1: ['氏名'], s2: ['氏名'], s3: ['氏名'] } },
  };
})());

// 02 ------------------------------------------------------------------------
cases.push((() => {
  const base = build('名簿', ['氏名', '所属'], (s) => {
    for (const p of range(0, 20)) s.add(p, [fullName(p, ' '), p.univ]);
  });
  const s1 = build('回答', ['姓', '名', '評価'], (s) => {
    for (const p of [...order(range(3, 17), 21), PEOPLE[24]]) s.add(p, [p.sei, p.mei, 'ABC'[p.sid % 3]]);
  });
  const s2 = build('回答', ['氏名', '点数'], (s) => {
    for (const p of order(range(0, 12), 22)) s.add(p, [fullName(p, ''), 50 + (p.sid % 50)]);
  });
  const s3 = build('回答', ['姓', '名', 'セイ', 'メイ', '備考'], (s) => {
    for (const p of order(range(10, 18), 23)) s.add(p, [p.sei, p.mei, p.seiKana, p.meiKana, null]);
  });
  return {
    name: 'names-combined-vs-split',
    note: 'base has 氏名 (half-width space); others have 姓/名 columns (with kana columns), or 氏名 without a space',
    files: [base.file('base', '名簿_氏名.xlsx'), s1.file('s1', '回答_姓名別.xlsx'), s2.file('s2', '点数_スペース無し.xlsx'), s3.file('s3', '回答_姓名カナ付き.xlsx')],
    expect: { matched: { s1: 14, s2: 12, s3: 8 }, plan: { s1: ['氏名'], s2: ['氏名'] } },
  };
})());

// 03 ------------------------------------------------------------------------
cases.push((() => {
  const base = build('名簿', ['氏名', 'フリガナ'], (s) => {
    for (const p of range(0, 20)) s.add(p, [fullName(p), fullKana(p)]);
  });
  const s1 = build('回答', ['ふりがな', '評価'], (s) => {
    for (const p of [...order(range(2, 16), 31), PEOPLE[20]]) s.add(p, [toHira(fullKana(p)), 'ABC'[p.sid % 3]]);
  });
  const s2 = build('回答', ['カナ', '点数'], (s) => {
    for (const p of [...order(range(4, 18), 32), PEOPLE[21]]) s.add(p, [toHalfKana(fullKana(p, ' ')), 50 + (p.sid % 50)]);
  });
  const s3 = build('回答', ['読み', 'コメント'], (s) => {
    for (const p of order(range(0, 9), 33)) s.add(p, [fullKana(p, ''), '']);
  });
  return {
    name: 'kana-scripts',
    note: 'フリガナ in katakana / hiragana / half-width katakana / without a space; the others carry NO kanji name',
    files: [base.file('base', '名簿_カナ.xlsx'), s1.file('s1', '回答_ひらがな.xlsx'), s2.file('s2', '適性_半角カナ.xlsx'), s3.file('s3', '回答_読み.xlsx')],
    expect: { matched: { s1: 14, s2: 14, s3: 9 }, plan: { s1: ['氏名カナ'], s2: ['氏名カナ'], s3: ['氏名カナ'] } },
  };
})());

// 04 ------------------------------------------------------------------------
cases.push((() => {
  const base = build('名簿', ['氏名', 'セイ', 'メイ', '大学'], (s) => {
    for (const p of range(0, 20)) s.add(p, [fullName(p), p.seiKana, p.meiKana, p.univ]);
  });
  const s1 = build('回答', ['氏名カナ', '評価'], (s) => {
    for (const p of [...order(range(2, 16), 41), PEOPLE[20]]) s.add(p, [fullKana(p, ' '), 'ABC'[p.sid % 3]]);
  });
  const s2 = build('回答', ['姓カナ', '名カナ', '点数'], (s) => {
    for (const p of [...order(range(4, 18), 42), PEOPLE[21]]) s.add(p, [toHira(p.seiKana), toHira(p.meiKana), 50 + (p.sid % 50)]);
  });
  const s3 = build('回答', ['フリガナ', '備考'], (s) => {
    for (const p of order(range(0, 9), 43)) s.add(p, [toHira(fullKana(p, '')), null]);
  });
  return {
    name: 'kana-split-vs-combined',
    note: 'base has セイ/メイ columns (+ 氏名); others have a combined フリガナ (hiragana / katakana) or 姓カナ/名カナ',
    files: [base.file('base', '名簿_セイメイ.xlsx'), s1.file('s1', '回答_氏名カナ.xlsx'), s2.file('s2', '回答_姓名カナ別.xlsx'), s3.file('s3', '回答_ふりがな.xlsx')],
    expect: { matched: { s1: 14, s2: 14, s3: 9 } },
  };
})());

// 05 ------------------------------------------------------------------------
cases.push((() => {
  const base = build('名簿', ['Name', '大学'], (s) => {
    for (const p of range(0, 20)) s.add(p, [`${p.meiRoman} ${p.seiRoman}`, p.univ]);
  });
  const s1 = build('Answers', ['Name', 'Score'], (s) => {
    for (const p of [...order(range(2, 16), 51), PEOPLE[20]]) s.add(p, [`${p.seiRoman.toUpperCase()} ${p.meiRoman}`, 50 + (p.sid % 50)]);
  });
  const s2 = build('回答', ['氏名', '評価'], (s) => {
    for (const p of [...order(range(4, 16), 52), PEOPLE[21]]) s.add(p, [`${p.seiRoman}, ${p.meiRoman}`, 'ABC'[p.sid % 3]]);
  });
  const s3 = build('Answers', ['Full Name', 'Memo'], (s) => {
    for (const p of order(range(0, 10), 53)) s.add(p, [`${p.meiRoman.toUpperCase()}  ${p.seiRoman.toUpperCase()}`, null]);
  });
  return {
    name: 'english-name-orders',
    note: "'Taro Yamada' vs 'YAMADA Taro' vs 'Yamada, Taro' vs 'TARO  YAMADA' (double space, capitals)",
    files: [base.file('base', 'roster_en.xlsx'), s1.file('s1', 'answers_family_first.xlsx'), s2.file('s2', '回答_カンマ区切り.xlsx'), s3.file('s3', 'answers_caps.xlsx')],
    expect: { matched: { s1: 14, s2: 12, s3: 10 } },
  };
})());

// 06 ------------------------------------------------------------------------
cases.push((() => {
  // base spelling vs the other file's spelling of the same person
  const baseName = (p: Person): string => (p.id === 'p3' ? `髙橋　${p.mei}` : p.id === 'p4' ? `渡邊　${p.mei}` : p.id === 'p17' ? `山﨑　${p.mei}` : fullName(p));
  const otherName = (p: Person): string => {
    switch (p.id) {
      case 'p3': return `高橋　${p.mei}`;
      case 'p4': return `渡辺　${p.mei}`;
      case 'p15': return `斉藤　${p.mei}`; // 斎 → 斉: a DIFFERENT character people confuse → never auto-joined, suggested
      case 'p17': return `山崎　${p.mei}`;
      case 'p2': return `${p.sei}（旧姓：鈴木）　${p.mei}`;
      case 'p10': return `${p.sei}　${p.mei}（${p.seiKana}　${p.meiKana}）`;
      default: return fullName(p);
    }
  };
  const base = build('名簿', ['氏名', '大学'], (s) => {
    for (const p of range(0, 20)) s.add(p, [baseName(p), p.univ]);
  });
  const s1 = build('回答', ['氏名', '評価'], (s) => {
    for (const p of [...order(range(0, 18), 61), PEOPLE[25]]) s.add(p, [otherName(p), 'ABC'[p.sid % 3]]);
  });
  return {
    name: 'name-variants-maiden',
    note: '髙橋/高橋, 渡邊/渡辺, 山﨑/山崎 (異体字 → joined, flagged), 斎藤/斉藤 (confusable → suggestion only), 旧姓（）, 読み仮名（） in parentheses',
    files: [base.file('base', '名簿_旧字体.xlsx'), s1.file('s1', '回答_新字体.xlsx')],
    expect: { matched: { s1: 17 }, surfaced: { s1: 1 } },
  };
})());

// 07 ------------------------------------------------------------------------
cases.push((() => {
  const base = build('名簿', ['氏名', '大学'], (s) => {
    for (const p of range(0, 20)) s.add(p, [fullName(p), p.univ]);
  });
  const dec = (p: Person): string => {
    switch (p.id) {
      case 'p0': return `${fullName(p, ' ')} 様`;
      case 'p1': return `${p.sei}${p.mei}さん`;
      case 'p2': return `${p.sei} ${p.mei}殿`;
      case 'p3': return `${p.sei} ${p.mei}君`;
      case 'p4': return `${fullName(p, ' ')}　さま`;
      case 'p5': return `${p.sei}${p.mei}くん`;
      default: return fullName(p);
    }
  };
  const s1 = build('回答', ['お名前', '評価'], (s) => {
    for (const p of order(range(0, 10), 71)) s.add(p, [dec(p), 'ABC'[p.sid % 3]]);
  });
  return {
    name: 'honorifics',
    note: '様 / さん / 殿 / 君 / さま / くん after the name',
    files: [base.file('base', '名簿.xlsx'), s1.file('s1', 'お礼状リスト.xlsx')],
    expect: { matched: { s1: 10 } },
  };
})());

// 08 ------------------------------------------------------------------------
cases.push((() => {
  const base = build('名簿', ['氏名', 'フリガナ', '大学'], (s) => {
    for (const p of range(0, 20)) s.add(p, [fullName(p), fullKana(p), p.univ]);
  });
  // a web-form export: the reading is printed under the name INSIDE the same cell
  const s1 = build('回答', ['氏名', '評価'], (s) => {
    for (const p of [...order(range(2, 16), 81), PEOPLE[20]]) s.add(p, [`${fullName(p, ' ')}\n${fullKana(p, ' ')}`, 'ABC'[p.sid % 3]]);
  });
  // the same, written by Windows Excel (CR LF) and with the reading in parentheses on the same line
  const s2 = build('回答', ['氏名', '点数'], (s) => {
    order(range(0, 10), 82).forEach((p, i) => s.add(p, [i % 2 ? `${fullName(p, ' ')}\r\n${fullKana(p, ' ')}` : `${fullName(p, ' ')}（${fullKana(p, ' ')}）`, 50 + (p.sid % 50)]));
  });
  return {
    name: 'name-with-reading-line',
    note: '氏名 cells hold two lines: the kanji name and its フリガナ (LF / CR LF), or the reading in parentheses',
    files: [base.file('base', '名簿.xlsx'), s1.file('s1', '回答_読み仮名付き.xlsx'), s2.file('s2', '点数_読み仮名付き.xlsx')],
    expect: { matched: { s1: 14, s2: 10 } },
  };
})());

// 09 ------------------------------------------------------------------------
cases.push((() => {
  const base = build('名簿', ['氏名', 'フリガナ', '大学'], (s) => {
    for (const p of range(0, 20)) s.add(p, [fullName(p), fullKana(p), p.univ]);
  });
  const s1 = build('回答', ['お名前（姓）', 'お名前（名）', 'セイ（カナ）', 'メイ（カナ）', '評価'], (s) => {
    for (const p of [...order(range(2, 16), 91), PEOPLE[20]]) s.add(p, [p.sei, p.mei, p.seiKana, p.meiKana, 'ABC'[p.sid % 3]]);
  });
  const s2 = build('回答', ['お名前（姓）', 'お名前（名）', '評価'], (s) => {
    for (const p of order(range(0, 9), 92)) s.add(p, [p.sei, p.mei, 'ABC'[p.sid % 3]]);
  });
  return {
    name: 'split-name-header-variants',
    note: "web-form headers 'お名前（姓）/お名前（名）' and 'セイ（カナ）/メイ（カナ）' against a base with 氏名 / フリガナ",
    files: [base.file('base', '名簿.xlsx'), s1.file('s1', 'フォーム回答_姓名カナ.xlsx'), s2.file('s2', 'フォーム回答_姓名のみ.xlsx')],
    expect: { matched: { s1: 14, s2: 9 } },
  };
})());

// 10 ------------------------------------------------------------------------
cases.push((() => {
  const base = build('Roster', ['Name', '大学'], (s) => {
    for (const p of range(0, 20)) s.add(p, [`${p.meiRoman} ${p.seiRoman}`, p.univ]);
  });
  const s1 = build('Scores', ['Last Name', 'First Name', 'Score'], (s) => {
    for (const p of [...order(range(2, 16), 101), PEOPLE[20]]) s.add(p, [p.seiRoman, p.meiRoman, 50 + (p.sid % 50)]);
  });
  const s2 = build('Scores', ['Surname', 'Given name', 'Score'], (s) => {
    for (const p of order(range(0, 8), 102)) s.add(p, [p.seiRoman.toUpperCase(), p.meiRoman.toLowerCase(), 50 + (p.sid % 50)]);
  });
  return {
    name: 'english-split-columns',
    note: "a one-column 'Taro Yamada' roster against 'Last Name / First Name' columns (capitals, any case)",
    files: [base.file('base', 'Roster.xlsx'), s1.file('s1', 'scores_split.xlsx'), s2.file('s2', 'scores_caps.xlsx')],
    expect: { matched: { s1: 14, s2: 8 } },
  };
})());

// 11 ------------------------------------------------------------------------
cases.push((() => {
  // 山田さくら vs 山田サクラ: the same reading in two scripts — suggested (似た字), never joined on the name alone (R6 decision)
  const asKata = (p: Person): string => `${p.sei}　${toKataOf(p.mei)}`;
  const base = build('名簿', ['氏名', '大学'], (s) => {
    for (const p of range(0, 20)) s.add(p, [fullName(p), p.univ]);
  });
  const s1 = build('回答', ['氏名', '評価'], (s) => {
    for (const p of order(range(0, 16), 111)) s.add(p, [p.id === 'p4' ? asKata(p) : fullName(p), 'ABC'[p.sid % 3]]);
  });
  return {
    name: 'name-hiragana-vs-katakana',
    note: '渡辺 さくら (hiragana given name) vs 渡辺 サクラ: the same person, but only a suggestion — the name key does not fold the scripts',
    files: [base.file('base', '名簿_ひらがな.xlsx'), s1.file('s1', '回答_カタカナ.xlsx')],
    expect: { matched: { s1: 15 }, surfaced: { s1: 1 } },
  };
})());

// 12 ------------------------------------------------------------------------
cases.push((() => {
  // 伊藤 花子 / 井藤 花子: two different people, one reading
  const base = build('名簿', ['氏名', 'フリガナ'], (s) => {
    for (const p of range(0, 14)) s.add(p, [fullName(p), fullKana(p)]);
    s.add('h1', ['伊藤　花子', 'イトウ　ハナコ']).add('h2', ['井藤　花子', 'イトウ　ハナコ']);
  });
  const s1 = build('回答', ['フリガナ', '評価'], (s) => {
    for (const p of order(range(0, 12), 121)) s.add(p, [fullKana(p), 'ABC'[p.sid % 3]]);
    s.add('h2', ['イトウ　ハナコ', 'B']);
  });
  return {
    name: 'homophone-surnames',
    note: '伊藤 花子 and 井藤 花子 share one フリガナ: the reading alone cannot say which of them answered',
    files: [base.file('base', '名簿_同読み.xlsx'), s1.file('s1', '回答_フリガナのみ.xlsx')],
    expect: { matched: { s1: 12 }, surfaced: { s1: 1 }, plan: { s1: ['氏名カナ'] } }, // (both base rows are ambiguous; only h2 is really in the other file)
  };
})());

// 12 ------------------------------------------------------------------------
cases.push((() => {
  // the base lists MARRIED names with the maiden name in parentheses; the other file still uses the maiden names
  const MAIDEN: Record<string, string> = { p7: '加藤', p10: '井上' };
  const base = build('名簿', ['氏名', '大学'], (s) => {
    for (const p of range(0, 14)) s.add(p, [MAIDEN[p.id] ? `${p.sei}　${p.mei}（旧姓：${MAIDEN[p.id]}）` : fullName(p), p.univ]);
  });
  const s1 = build('回答', ['氏名', '評価'], (s) => {
    for (const p of order(range(0, 12), 121)) s.add(p, [MAIDEN[p.id] ? `${MAIDEN[p.id]}　${p.mei}` : fullName(p), 'ABC'[p.sid % 3]]);
  });
  return {
    name: 'maiden-name-in-one-file',
    note: 'the base says 小林 花子（旧姓：加藤）, the other file 加藤 花子: the same person under the other surname — suggested, not joined',
    files: [base.file('base', '名簿_旧姓付き.xlsx'), s1.file('s1', '回答_旧姓のみ.xlsx')],
    expect: { matched: { s1: 10 }, surfaced: { s1: 2 } },
  };
})());
