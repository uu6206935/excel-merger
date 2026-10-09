// Deterministic data for the sample / fixture workbooks in samples/ (owner: T).
// Pure JS (no I/O): tools/make-samples.mjs turns this into files (Excel COM for .xlsx/.xls, Node for the CSV),
// tests/qa/sample-data.test.ts checks the invariants the samples are meant to demonstrate.
//
// Cell input convention used in `rows` arrays (what the PowerShell builder writes through Range.Formula):
//   number            -> number (dates are Excel serials; the builder applies the date format)
//   '=...'            -> formula (English function names, comma separators = Range.Formula syntax)
//   "'..."            -> text forced with an apostrophe (Excel keeps the text, e.g. "'000123")
//   other strings     -> typed text
//   null              -> empty cell

/** 1900-system Excel serial of a calendar date. */
export function serial(y, m, d) {
  return Math.round((Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000);
}

/** Calendar date of an Excel serial (1900 system, serial > 60). */
export function fromSerial(n) {
  const t = new Date(Date.UTC(1899, 11, 30) + n * 86400000);
  return { y: t.getUTCFullYear(), m: t.getUTCMonth() + 1, d: t.getUTCDate() };
}

/** Small deterministic PRNG (mulberry32). */
export function rng(seed) {
  let a = seed >>> 0;
  const next = () => {
    a = (a + 0x6d2b79f5) >>> 0;
    let t = a;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
  return {
    next,
    int: (lo, hi) => lo + Math.floor(next() * (hi - lo + 1)),
    pick: (arr) => arr[Math.floor(next() * arr.length)],
    weighted: (pairs) => {
      const total = pairs.reduce((s, p) => s + p[1], 0);
      let x = next() * total;
      for (const [v, w] of pairs) {
        x -= w;
        if (x < 0) return v;
      }
      return pairs[pairs.length - 1][0];
    },
    shuffle: (arr) => {
      const a2 = arr.slice();
      for (let i = a2.length - 1; i > 0; i--) {
        const j = Math.floor(next() * (i + 1));
        [a2[i], a2[j]] = [a2[j], a2[i]];
      }
      return a2;
    },
  };
}

// ---------------------------------------------------------------------------------------------------------------
// Vocabulary
// ---------------------------------------------------------------------------------------------------------------

/** [kanji, katakana, romaji] */
export const FAMILIES = [
  ['佐藤', 'サトウ', 'sato'], ['鈴木', 'スズキ', 'suzuki'], ['田中', 'タナカ', 'tanaka'], ['伊藤', 'イトウ', 'ito'],
  ['渡辺', 'ワタナベ', 'watanabe'], ['山本', 'ヤマモト', 'yamamoto'], ['中村', 'ナカムラ', 'nakamura'], ['小林', 'コバヤシ', 'kobayashi'],
  ['加藤', 'カトウ', 'kato'], ['吉田', 'ヨシダ', 'yoshida'], ['山田', 'ヤマダ', 'yamada'], ['佐々木', 'ササキ', 'sasaki'],
  ['山口', 'ヤマグチ', 'yamaguchi'], ['松本', 'マツモト', 'matsumoto'], ['井上', 'イノウエ', 'inoue'], ['木村', 'キムラ', 'kimura'],
  ['林', 'ハヤシ', 'hayashi'], ['清水', 'シミズ', 'shimizu'], ['森', 'モリ', 'mori'], ['池田', 'イケダ', 'ikeda'],
  ['橋本', 'ハシモト', 'hashimoto'], ['阿部', 'アベ', 'abe'], ['石川', 'イシカワ', 'ishikawa'], ['前田', 'マエダ', 'maeda'],
  ['藤田', 'フジタ', 'fujita'], ['小川', 'オガワ', 'ogawa'], ['岡田', 'オカダ', 'okada'], ['後藤', 'ゴトウ', 'goto'],
  ['長谷川', 'ハセガワ', 'hasegawa'], ['村上', 'ムラカミ', 'murakami'], ['近藤', 'コンドウ', 'kondo'], ['石井', 'イシイ', 'ishii'],
  ['坂本', 'サカモト', 'sakamoto'], ['遠藤', 'エンドウ', 'endo'], ['青木', 'アオキ', 'aoki'], ['藤井', 'フジイ', 'fujii'],
  ['西村', 'ニシムラ', 'nishimura'], ['福田', 'フクダ', 'fukuda'], ['太田', 'オオタ', 'ota'], ['三浦', 'ミウラ', 'miura'],
  ['藤原', 'フジワラ', 'fujiwara'], ['岡本', 'オカモト', 'okamoto'], ['松田', 'マツダ', 'matsuda'], ['中川', 'ナカガワ', 'nakagawa'],
  ['原田', 'ハラダ', 'harada'], ['小野', 'オノ', 'ono'], ['竹内', 'タケウチ', 'takeuchi'], ['金子', 'カネコ', 'kaneko'],
];

/** [kanji, katakana, romaji] */
export const GIVENS = [
  ['翔太', 'ショウタ', 'shota'], ['美咲', 'ミサキ', 'misaki'], ['大輝', 'ダイキ', 'daiki'], ['陽菜', 'ヒナ', 'hina'],
  ['蓮', 'レン', 'ren'], ['結衣', 'ユイ', 'yui'], ['悠真', 'ユウマ', 'yuma'], ['さくら', 'サクラ', 'sakura'],
  ['颯太', 'ソウタ', 'sota'], ['葵', 'アオイ', 'aoi'], ['湊', 'ミナト', 'minato'], ['凛', 'リン', 'rin'],
  ['大和', 'ヤマト', 'yamato'], ['芽依', 'メイ', 'mei'], ['陽翔', 'ハルト', 'haruto'], ['美月', 'ミヅキ', 'mizuki'],
  ['優斗', 'ユウト', 'yuto'], ['彩花', 'アヤカ', 'ayaka'], ['健太', 'ケンタ', 'kenta'], ['真央', 'マオ', 'mao'],
  ['拓海', 'タクミ', 'takumi'], ['七海', 'ナナミ', 'nanami'], ['心春', 'コハル', 'koharu'], ['樹', 'イツキ', 'itsuki'],
  ['杏奈', 'アンナ', 'anna'], ['隼人', 'ハヤト', 'hayato'], ['詩織', 'シオリ', 'shiori'], ['遼', 'リョウ', 'ryo'],
  ['千尋', 'チヒロ', 'chihiro'], ['航', 'ワタル', 'wataru'], ['花音', 'カノン', 'kanon'], ['直樹', 'ナオキ', 'naoki'],
  ['莉子', 'リコ', 'riko'], ['海斗', 'カイト', 'kaito'], ['楓', 'カエデ', 'kaede'], ['亮介', 'リョウスケ', 'ryosuke'],
  ['奈々', 'ナナ', 'nana'], ['慎也', 'シンヤ', 'shinya'], ['萌', 'モエ', 'moe'], ['晴', 'ハル', 'haru'],
];

export const UNIVERSITIES = [
  { name: '慶應義塾大学', short: '慶應大学', domain: 'keio.jp', faculties: ['経済学部', '法学部', '商学部', '理工学部', '総合政策学部', '環境情報学部'] },
  { name: '早稲田大学', short: '早大', domain: 'waseda.jp', faculties: ['政治経済学部', '法学部', '商学部', '基幹理工学部', '国際教養学部', '文化構想学部'] },
  { name: '東京大学', short: '東大', domain: 'g.ecc.u-tokyo.ac.jp', faculties: ['経済学部', '法学部', '工学部', '教養学部', '理学部'] },
  { name: '京都大学', short: '京大', domain: 'st.kyoto-u.ac.jp', faculties: ['経済学部', '法学部', '工学部', '総合人間学部', '農学部'] },
  { name: '一橋大学', short: '一橋', domain: 'g.hit-u.ac.jp', faculties: ['経済学部', '商学部', '法学部', '社会学部'] },
  { name: '大阪大学', short: '阪大', domain: 'ecs.osaka-u.ac.jp', faculties: ['経済学部', '法学部', '工学部', '基礎工学部', '人間科学部'] },
  { name: '明治大学', short: '明大', domain: 'meiji.ac.jp', faculties: ['政治経済学部', '商学部', '経営学部', '理工学部', '情報コミュニケーション学部'] },
  { name: '上智大学', short: '上智', domain: 'eagle.sophia.ac.jp', faculties: ['経済学部', '法学部', '外国語学部', '理工学部', '総合グローバル学部'] },
  { name: '東北大学', short: '東北大', domain: 'dc.tohoku.ac.jp', faculties: ['経済学部', '法学部', '工学部', '理学部'] },
  { name: '名古屋大学', short: '名大', domain: 's.thers.ac.jp', faculties: ['経済学部', '法学部', '工学部', '情報学部'] },
];

/** [label, age on the reference date] */
export const GRADES = [
  ['学部3年', 20],
  ['学部4年', 21],
  ['修士1年', 22],
  ['修士2年', 23],
];

/** 基準日 (マスタ!H2) used by the 年齢 = DATEDIF(生年月日, 基準日, "Y") formulas. */
export const REF_DATE = { y: 2026, m: 4, d: 1 };

export const INTERVIEWERS = ['田村', '松井', '大野', '杉山'];
export const COMMENTS = [
  '志望動機が明確',
  '論理的に話せる',
  'コミュニケーション良好',
  'やや緊張気味',
  '専門知識が豊富',
  '回答が簡潔で分かりやすい',
  'リーダー経験あり',
  null,
  '二次面接で再確認',
  '入社意欲が高い',
];

/** old-form (異体字) -> common form, only the ones used in the samples */
export const ITAIJI = { 髙: '高', 齋: '斎', 邉: '辺', 﨑: '崎' };

// ---------------------------------------------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------------------------------------------

const H_KANA = {
  ア: 'ｱ', イ: 'ｲ', ウ: 'ｳ', エ: 'ｴ', オ: 'ｵ', カ: 'ｶ', キ: 'ｷ', ク: 'ｸ', ケ: 'ｹ', コ: 'ｺ', サ: 'ｻ', シ: 'ｼ', ス: 'ｽ', セ: 'ｾ', ソ: 'ｿ',
  タ: 'ﾀ', チ: 'ﾁ', ツ: 'ﾂ', テ: 'ﾃ', ト: 'ﾄ', ナ: 'ﾅ', ニ: 'ﾆ', ヌ: 'ﾇ', ネ: 'ﾈ', ノ: 'ﾉ', ハ: 'ﾊ', ヒ: 'ﾋ', フ: 'ﾌ', ヘ: 'ﾍ', ホ: 'ﾎ',
  マ: 'ﾏ', ミ: 'ﾐ', ム: 'ﾑ', メ: 'ﾒ', モ: 'ﾓ', ヤ: 'ﾔ', ユ: 'ﾕ', ヨ: 'ﾖ', ラ: 'ﾗ', リ: 'ﾘ', ル: 'ﾙ', レ: 'ﾚ', ロ: 'ﾛ', ワ: 'ﾜ', ヲ: 'ｦ',
  ン: 'ﾝ', ァ: 'ｧ', ィ: 'ｨ', ゥ: 'ｩ', ェ: 'ｪ', ォ: 'ｫ', ャ: 'ｬ', ュ: 'ｭ', ョ: 'ｮ', ッ: 'ｯ', ー: 'ｰ', '゙': 'ﾞ', '゚': 'ﾟ',
  '　': ' ',
};

/** full-width katakana -> half-width katakana (ｻﾄｳ ｼｮｳﾀ), as legacy systems export it */
export function toHalfKana(s) {
  return [...s.normalize('NFD')].map((ch) => H_KANA[ch] ?? ch).join('');
}

/** half-width digits -> full-width digits */
export function toFullDigits(s) {
  return s.replace(/[0-9]/g, (d) => String.fromCharCode(d.charCodeAt(0) + 0xfee0));
}

/** comparison key used only to keep the generated names unique (異体字 + spaces folded) */
export function foldName(s) {
  return [...s.replace(/[\s　]/g, '')].map((ch) => ITAIJI[ch] ?? ch).join('');
}

const pad = (n, w) => String(n).padStart(w, '0');

// ---------------------------------------------------------------------------------------------------------------
// Students (the base roster)
// ---------------------------------------------------------------------------------------------------------------

/** Fixed students at fixed positions (0-based index in the roster) — the cases the samples demonstrate. */
export const SPECIAL = {
  3: { family: ['髙橋', 'タカハシ', 'takahashi'], given: ['優斗', 'ユウト', 'yuto'], uni: '慶應義塾大学', note: '異体字 髙 (02では「高橋」)' },
  8: { family: ['佐藤', 'サトウ', 'sato'], given: ['翔太', 'ショウタ', 'shota'], uni: '慶應義塾大学', faculty: '総合政策学部', note: '同姓同名 (慶應)' },
  11: { family: ['高橋', 'タカハシ', 'takahashi'], given: ['健太', 'ケンタ', 'kenta'], uni: '東京大学', note: '高橋 (02では「髙橋」)' },
  17: { family: ['齋藤', 'サイトウ', 'saito'], given: ['美咲', 'ミサキ', 'misaki'], uni: '早稲田大学', note: '異体字 齋 (02では「斎藤」)' },
  23: { family: ['渡邉', 'ワタナベ', 'watanabe'], given: ['大輝', 'ダイキ', 'daiki'], uni: '京都大学', note: '異体字 邉 (02では「渡辺」)' },
  30: { family: ['山﨑', 'ヤマザキ', 'yamazaki'], given: ['彩花', 'アヤカ', 'ayaka'], uni: '一橋大学', note: '異体字 﨑 (02/05では「山崎」)' },
  37: { family: ['佐藤', 'サトウ', 'sato'], given: ['翔太', 'ショウタ', 'shota'], uni: '早稲田大学', faculty: '商学部', note: '同姓同名 (早稲田)' },
  44: { family: ['小野', 'オノ', 'ono'], given: ['さくら', 'サクラ', 'sakura'], uni: '上智大学', note: 'ひらがなの名' },
};

export const N_STUDENTS = 60;
/** base roster indices that have no e-mail address (COUNTBLANK demo) */
export const NO_EMAIL = [14, 52];

export function makeStudents() {
  const r = rng(20260401);
  const used = new Set();
  // names reserved for the special students (so random picks never collide with them after 異体字 folding)
  for (const sp of Object.values(SPECIAL)) used.add(foldName(sp.family[0] + sp.given[0]));
  const emails = new Set();
  const phones = new Set();
  const students = [];
  for (let i = 0; i < N_STUDENTS; i++) {
    const sp = SPECIAL[i];
    let family;
    let given;
    if (sp) {
      family = sp.family;
      given = sp.given;
    } else {
      for (;;) {
        family = r.pick(FAMILIES);
        given = r.pick(GIVENS);
        const k = foldName(family[0] + given[0]);
        if (!used.has(k)) {
          used.add(k);
          break;
        }
      }
    }
    const uni = sp?.uni ? UNIVERSITIES.find((u) => u.name === sp.uni) : r.weighted(UNIVERSITIES.map((u, k) => [u, k < 3 ? 3 : k < 6 ? 2 : 1]));
    const faculty = sp?.faculty ?? r.pick(uni.faculties);
    const [grade, baseAge] = r.weighted([
      [GRADES[0], 40],
      [GRADES[1], 35],
      [GRADES[2], 15],
      [GRADES[3], 10],
    ]);
    const age = baseAge + (r.next() < 0.1 ? 1 : 0); // 浪人・留年
    // born between (REF - (age+1) years + 1 day) and (REF - age years): DATEDIF(..., REF, "Y") = age
    const birth = serial(REF_DATE.y - age - 1, REF_DATE.m, REF_DATE.d + 1) + r.int(0, 364);
    const id = pad(100 + i * 7 + (i % 3), 6);
    const aspiration = r.weighted([
      [1, 5],
      [2, 15],
      [3, 30],
      [4, 30],
      [5, 20],
    ]);
    let email = null;
    if (!NO_EMAIL.includes(i)) {
      const kind = i % 3;
      if (kind === 0) email = `${given[2]}.${family[2]}@${uni.domain}`;
      else if (kind === 1) email = `${family[2]}.${given[2]}${pad((i * 37) % 100, 2)}@gmail.com`;
      else email = `${given[2][0]}${family[2]}${pad(fromSerial(birth).y % 100, 2)}@${uni.domain}`;
      let e = email;
      for (let k = 2; emails.has(e); k++) e = email.replace('@', `${k}@`);
      email = e;
      emails.add(email);
    }
    let phone;
    do {
      phone = `0${[90, 80, 70][r.int(0, 2)]}-${pad(r.int(1000, 9999), 4)}-${pad(r.int(0, 9999), 4)}`;
    } while (phones.has(phone));
    phones.add(phone);
    students.push({
      index: i,
      id,
      family: family[0],
      given: given[0],
      /** 氏名 as written in the base roster: 姓 + 全角スペース + 名 */
      name: `${family[0]}　${given[0]}`,
      kana: `${family[1]}　${given[1]}`,
      romaji: `${family[2]} ${given[2]}`,
      uni: uni.name,
      uniShort: uni.short,
      faculty,
      grade,
      birth,
      age,
      email,
      phone,
      aspiration,
      note: sp?.note ?? null,
    });
  }
  return students;
}

// ---------------------------------------------------------------------------------------------------------------
// 01_学生名簿.xlsx (BASE)
// ---------------------------------------------------------------------------------------------------------------

export const ROSTER = {
  sheet: '名簿',
  titleRow: 1,
  headerRow: 3,
  firstDataRow: 4,
  headers: ['学生ID', '氏名', 'フリガナ', '大学', '学部', '学年', 'メール', '電話', '生年月日', '年齢', '志望度', '重複チェック'],
  /** hidden helper column (L) */
  helperCol: 12,
};

export function rosterRows(students) {
  const first = ROSTER.firstDataRow;
  const last = first + students.length - 1;
  return students.map((s, k) => {
    const row = first + k;
    return [
      s.id, // column A is formatted as text ("@") before writing, so the leading zeros stay
      s.name,
      s.kana,
      s.uni,
      s.faculty,
      s.grade,
      s.email,
      s.phone,
      s.birth,
      `=DATEDIF(I${row},基準日,"Y")`,
      s.aspiration,
      `=COUNTIF($B$${first}:$B$${last},B${row})`,
    ];
  });
}

export function masterBlock(students) {
  const faculties = [];
  for (const u of UNIVERSITIES) for (const f of u.faculties) if (!faculties.includes(f)) faculties.push(f);
  void students;
  return {
    universities: UNIVERSITIES.map((u) => u.name),
    faculties,
    grades: GRADES.map((g) => g[0]),
    aspiration: [
      [1, '低い'],
      [2, 'やや低い'],
      [3, '普通'],
      [4, '高い'],
      [5, '第一志望'],
    ],
    refDate: serial(REF_DATE.y, REF_DATE.m, REF_DATE.d),
  };
}

/** 集計 sheet: formulas over 名簿 (whole-column and bounded ranges, names, SUMPRODUCT, VLOOKUP, INDEX/MATCH). */
export function summaryBlock(students) {
  const f = ROSTER.firstDataRow;
  const l = f + students.length - 1;
  const R = (col) => `名簿!$${col}$${f}:$${col}$${l}`;
  const uniRows = UNIVERSITIES.map((u, k) => {
    const row = 5 + k;
    return [
      u.name,
      `=COUNTIF(名簿!$D:$D,$A${row})`,
      `=IFERROR(AVERAGEIF(${R('D')},$A${row},${R('J')}),"-")`,
      `=COUNTIFS(${R('D')},$A${row},${R('K')},">=4")`,
      `=IFERROR(SUMPRODUCT((${R('D')}=$A${row})*${R('K')})/B${row},"-")`,
      `=IFERROR(B${row}/$B$${5 + UNIVERSITIES.length},0)`,
    ];
  });
  const totalRow = 5 + UNIVERSITIES.length; // 15
  const total = [
    '合計',
    `=SUM(B5:B${totalRow - 1})`,
    `=AVERAGE(名簿!J${f}:J${l})`,
    `=SUM(D5:D${totalRow - 1})`,
    '=AVERAGE(志望度列)',
    `=SUM(F5:F${totalRow - 1})`,
  ];
  const gradeHeaderRow = totalRow + 2; // 17
  const gradeRows = GRADES.map((g, k) => [g[0], `=COUNTIF(名簿!F:F,A${gradeHeaderRow + 1 + k})`]);
  const otherHeaderRow = gradeHeaderRow + GRADES.length + 2; // 23
  const lookupId = students[1].id;
  const others = [
    ['メール未登録', `=COUNTBLANK(名簿!G${f}:G${l})`],
    ['最年長', '=MAX(名簿!J:J)'],
    ['志望度5の人数', '=COUNTIF(志望度列,5)'],
    ['先頭の学生', `=名簿!B${f}`],
    [`ID ${lookupId} の氏名`, `=VLOOKUP("${lookupId}",名簿!$A$${f}:$K$${l},2,FALSE)`],
    ['志望度が最も高い学生', `=INDEX(${R('B')},MATCH(MAX(${R('K')}),${R('K')},0))`],
    ['同姓同名の件数', `=COUNTIF(名簿!L${f}:L${l},">1")`],
  ];
  return {
    sheet: '集計',
    title: '大学別 応募者集計',
    note: '※ 名簿シートから自動集計しています（手入力不要）',
    headerRow: 4,
    headers: ['大学', '人数', '平均年齢', '志望度4以上', '志望度平均', '構成比'],
    uniRows,
    totalRow,
    total,
    gradeHeaderRow,
    gradeRows,
    otherHeaderRow,
    others,
  };
}

// ---------------------------------------------------------------------------------------------------------------
// 02_面接評価.xlsx
// ---------------------------------------------------------------------------------------------------------------

export const INTERVIEW = {
  sheet: '面接評価',
  tableSheet: '評価テーブル',
  tableName: '面接評価表',
  headers: ['受験番号', '氏名', '面接官', '評価1', '評価2', '評価3', '平均', '判定', 'コメント'],
  /** base students that did not take the interview */
  missing: [27, 52],
  /** base students with a second (resubmitted) row */
  resubmitted: [5, 40],
  /** people that are not in the base roster */
  extras: [
    { id: '000601', name: '野口 大地' },
    { id: '000622', name: '菅原美優' },
    { id: '000648', name: '桑原　陸' },
  ],
};

/** 02 氏名 spelling for a base student (異体字 swapped, spaces varied) */
function interviewName(s, k) {
  let fam = s.family;
  if (s.index === 11) fam = '髙橋'; // the other direction: base 高橋 -> 髙橋
  else fam = [...fam].map((ch) => ITAIJI[ch] ?? ch).join('');
  let name;
  if (k % 4 === 0) name = `${fam} ${s.given}`;
  else if (k % 6 === 1) name = `${fam}　${s.given}`;
  else name = `${fam}${s.given}`;
  if (s.index === 20) name += ' '; // trailing space typed by accident
  return name;
}

/** 受験番号 as typed in 02: text with zeros / number without zeros / full-width text */
function interviewId(s) {
  if (s.index % 13 === 7) return { kind: 'fullwidth', input: "'" + toFullDigits(s.id) };
  if (s.index % 5 === 1) return { kind: 'number', input: Number(s.id) };
  return { kind: 'text', input: "'" + s.id };
}

export function interviewRows(students) {
  const r = rng(20260915);
  const takers = students.filter((s) => !INTERVIEW.missing.includes(s.index));
  const order = r.shuffle(takers);
  const entries = order.map((s) => ({ s, resub: false }));
  // resubmissions go further down the sheet (typed later)
  for (const idx of INTERVIEW.resubmitted) {
    const s = students[idx];
    const pos = entries.findIndex((e) => e.s === s);
    const at = Math.min(entries.length, pos + 12 + (idx % 7));
    entries.splice(at, 0, { s, resub: true });
  }
  // extras interleaved
  INTERVIEW.extras.forEach((x, k) => entries.splice(9 + k * 19, 0, { extra: x }));
  const rows = [];
  const meta = [];
  entries.forEach((e, k) => {
    const row = 2 + k;
    const score = () => {
      const base = e.s ? e.s.aspiration : 3;
      return Math.max(1, Math.min(5, Math.round(base * 0.5 + r.int(0, 3) - 0.5)));
    };
    const idCell = e.extra ? { kind: 'text', input: "'" + e.extra.id } : interviewId(e.s);
    const name = e.extra ? e.extra.name : interviewName(e.s, k);
    const comment = e.resub ? '再提出（差し替え）' : r.pick(COMMENTS);
    rows.push([
      idCell.input,
      name,
      INTERVIEWERS[(k * 3 + (e.s ? e.s.index : 1)) % INTERVIEWERS.length],
      score(),
      score(),
      score(),
      `=AVERAGE(D${row}:F${row})`,
      `=IF(G${row}>=3.5,"合格","保留")`,
      comment,
    ]);
    meta.push({ row, baseIndex: e.s ? e.s.index : null, idKind: idCell.kind, extra: !!e.extra, resubmitted: e.resub, name });
  });
  return { rows, meta };
}

// ---------------------------------------------------------------------------------------------------------------
// 03_連絡先.csv (Shift_JIS, no ID)
// ---------------------------------------------------------------------------------------------------------------

export const CONTACTS = {
  headers: ['氏名', 'メール', '電話'],
  extras: [
    { name: '西田 光', kana: 'ニシダ ヒカル', email: 'hikaru.nishida@gmail.com', phone: '080-3141-5926' },
    { name: '今井 さやか', kana: 'イマイ サヤカ', email: 'sayaka_imai@yahoo.co.jp', phone: '070-2718-2818' },
  ],
};

export function contactRows(students) {
  const picked = students.filter((s) => s.index % 4 !== 3);
  const list = [
    ...picked.map((s) => ({ s, kana: s.kana.replace('　', ' ') })),
    ...CONTACTS.extras.map((x) => ({ x, kana: x.kana })),
  ];
  list.sort((a, b) => a.kana.localeCompare(b.kana, 'ja') || (a.s?.index ?? 99) - (b.s?.index ?? 99));
  const meta = list.map((e, k) => ({ row: 2 + k, baseIndex: e.s ? e.s.index : null }));
  const rows = list.map((e, k) => {
    if (e.x) return [e.x.name, e.x.email, e.x.phone];
    const s = e.s;
    const fam = [...s.family].map((ch) => (ch === '﨑' || ch === '髙' ? ITAIJI[ch] : ch)).join(''); // CSV exports use JIS-level-1/2 forms
    let name = k % 5 === 2 ? `${fam}${s.given}` : `${fam} ${s.given}`;
    if (k === 6 || k === 19) name += ' 様';
    let email = s.email ?? `${s.romaji.split(' ')[1]}.${s.romaji.split(' ')[0]}@outlook.jp`;
    if (k % 7 === 3) email = email[0].toUpperCase() + email.slice(1);
    let phone = s.phone;
    if (k % 3 === 0) phone = phone.replace(/-/g, '');
    else if (k % 9 === 2) phone = phone.replace(/^(\d+)-(\d+)-(\d+)$/, '$1($2)$3');
    return [name, email, phone];
  });
  return { rows, meta };
}

/** CSV text (CRLF, quoted only when needed) — encode it as Shift_JIS (cp932) when writing */
export function toCsv(headers, rows) {
  const q = (v) => {
    const s = v == null ? '' : String(v);
    return /[",\r\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
  };
  return [headers, ...rows].map((r) => r.map(q).join(',')).join('\r\n') + '\r\n';
}

// ---------------------------------------------------------------------------------------------------------------
// 04_適性検査.xls (legacy .xls; 学籍番号 unrelated to 学生ID; kana only)
// ---------------------------------------------------------------------------------------------------------------

export const APTITUDE = {
  sheet: '適性検査結果',
  headers: ['学籍番号', '氏名カナ', '受検日', '言語', '非言語', '合計', '性格タイプ', '総合判定'],
  extras: [
    { kana: 'ミヤモト ユウキ' },
    { kana: 'オオツカ アヤネ' },
  ],
};

export function aptitudeRows(students) {
  const r = rng(20260803);
  const picked = students.filter((s) => s.index % 3 !== 1);
  const entries = [...picked.map((s) => ({ s })), ...APTITUDE.extras.map((x) => ({ x }))];
  // 学籍番号: 8-digit university student numbers, unrelated to 学生ID
  const nums = new Set();
  for (const e of entries) {
    let n;
    do n = 81900000 + r.int(0, 99999) * 10 + r.int(0, 9);
    while (nums.has(n));
    nums.add(n);
    e.num = n;
  }
  entries.sort((a, b) => a.num - b.num);
  const meta = [];
  const rows = entries.map((e, k) => {
    const row = 2 + k;
    const full = e.s ? e.s.kana : e.x.kana.replace(' ', '　');
    let kana;
    let kind;
    if (k % 5 === 0) {
      kana = toHalfKana(full);
      kind = 'halfwidth';
    } else if (k % 7 === 3) {
      kana = full.replace('　', '');
      kind = 'nospace';
    } else {
      kana = full.replace('　', ' ');
      kind = 'fullwidth';
    }
    meta.push({ row, baseIndex: e.s ? e.s.index : null, kanaKind: kind, kana });
    return [
      e.num,
      kana,
      serial(2026, 8, 3 + (k % 20)),
      r.int(35, 100),
      r.int(30, 100),
      `=D${row}+E${row}`,
      r.pick(['A', 'B', 'C', 'D', 'E']),
      `=IF(F${row}>=150,"A",IF(F${row}>=120,"B","C"))`,
    ];
  });
  return { rows, meta };
}

// ---------------------------------------------------------------------------------------------------------------
// 05_研修出欠.xlsx (氏名 + 大学 only: 同姓同名 needs the composite key 氏名＋大学)
// ---------------------------------------------------------------------------------------------------------------

export const ATTENDANCE = {
  sheet: '出欠',
  title: 'インターン研修 出欠表（2026年9月）',
  headerRow: 2,
  headers: ['氏名', '大学', '第1回(9/10)', '第2回(9/17)', '第3回(9/24)', '出席回数', '備考'],
  extras: [{ name: '新井 拓也', uni: '東京大学' }],
};

export function attendanceRows(students) {
  const r = rng(20260910);
  const picked = students.filter((s) => s.index % 2 === 0 || s.index === 37);
  const entries = [...picked.map((s) => ({ s })), ...ATTENDANCE.extras.map((x) => ({ x }))];
  const uniOrder = (name) => UNIVERSITIES.findIndex((u) => u.name === name);
  entries.sort((a, b) => uniOrder(a.s?.uni ?? a.x.uni) - uniOrder(b.s?.uni ?? b.x.uni) || (a.s?.index ?? 99) - (b.s?.index ?? 99));
  const first = ATTENDANCE.headerRow + 1;
  const meta = [];
  const rows = entries.map((e, k) => {
    const row = first + k;
    const mark = () => r.weighted([
      ['○', 80],
      ['×', 10],
      ['遅刻', 10],
    ]);
    let name;
    let uni;
    if (e.x) {
      name = e.x.name;
      uni = e.x.uni;
    } else {
      const s = e.s;
      const fam = [...s.family].map((ch) => ITAIJI[ch] ?? ch).join('');
      name = `${fam} ${s.given}`;
      const pair = s.index === 8 || s.index === 37;
      uni = !pair && k % 4 === 1 ? s.uniShort : s.uni; // 表記ゆれ (慶應大学 / 早大 / 東大 …)
    }
    meta.push({ row, baseIndex: e.s ? e.s.index : null, name, uni });
    return [
      name,
      uni,
      mark(),
      mark(),
      mark(),
      `=COUNTIF(C${row}:E${row},"○")+COUNTIF(C${row}:E${row},"遅刻")`,
      r.next() < 0.12 ? r.pick(['オンライン参加', '途中退出', '体調不良のため欠席連絡あり']) : null,
    ];
  });
  return { rows, meta };
}

// ---------------------------------------------------------------------------------------------------------------
// Everything, as one JSON-able spec for the PowerShell builder
// ---------------------------------------------------------------------------------------------------------------

export const FILES = {
  base: '01_学生名簿.xlsx',
  interview: '02_面接評価.xlsx',
  contacts: '03_連絡先.csv',
  aptitude: '04_適性検査.xls',
  attendance: '05_研修出欠.xlsx',
};

function summaryRefs(s) {
  const out = [];
  const hr = s.headerRow;
  s.uniRows.forEach((_, k) => {
    const row = hr + 1 + k;
    out.push({ cell: `B${row}`, headers: ['大学'] });
    out.push({ cell: `C${row}`, headers: ['大学', '年齢'] });
    out.push({ cell: `D${row}`, headers: ['大学', '志望度'] });
    out.push({ cell: `E${row}`, headers: ['大学', '志望度'] });
  });
  out.push({ cell: `C${s.totalRow}`, headers: ['年齢'] });
  s.gradeRows.forEach((_, k) => out.push({ cell: `B${s.gradeHeaderRow + 1 + k}`, headers: ['学年'] }));
  const o = s.otherHeaderRow + 1;
  out.push({ cell: `B${o}`, headers: ['メール'] });
  out.push({ cell: `B${o + 1}`, headers: ['年齢'] });
  out.push({ cell: `B${o + 3}`, headers: ['氏名'] });
  out.push({ cell: `B${o + 5}`, headers: ['氏名', '志望度'] });
  out.push({ cell: `B${o + 6}`, headers: ['重複チェック'] });
  return out;
}

/**
 * Ground truth for the samples (samples/truth.json): which row of every other file belongs to which base row.
 * Rows are 1-based Excel rows (for the CSV: line number = Excel row when opened in Excel).
 */
export function buildTruth(spec) {
  const R = spec.base.roster;
  const baseRow = (idx) => (idx == null ? null : R.firstDataRow + idx);
  const n = spec.students.length;
  const iv = spec.interview.meta;
  const ivBase = new Set(iv.filter((m) => m.baseIndex != null).map((m) => m.baseIndex));
  const dupIds = INTERVIEW.resubmitted.map((i) => spec.students[i].id);
  const map = (meta, extra = () => ({})) => meta.map((m) => ({ row: m.row, baseRow: baseRow(m.baseIndex), ...extra(m) }));
  const unmatchedBase = (meta) => {
    const s = new Set(meta.filter((m) => m.baseIndex != null).map((m) => m.baseIndex));
    return spec.students.filter((st) => !s.has(st.index)).map((st) => baseRow(st.index));
  };
  return {
    note: 'Generated by tools/make-samples.mjs from tools/lib/sample-data.mjs. Rows are 1-based Excel rows.',
    files: FILES,
    base: {
      file: FILES.base,
      sheet: R.sheet,
      sheets: ['名簿', '集計', 'ピボット', 'マスタ'],
      headerRow: R.headerRow,
      firstDataRow: R.firstDataRow,
      lastDataRow: R.firstDataRow + n - 1,
      footerRows: [R.firstDataRow + n + 1, R.firstDataRow + n + 2],
      headers: R.headers,
      hiddenColumns: ['L'],
      freeze: { cell: 'C4', rows: 3, cols: 2 },
      names: { 大学一覧: 'マスタ!$A$2:$A$20', 志望度列: '名簿!$K:$K', 基準日: 'マスタ!$H$2' },
      summarySheet: '集計',
      /** 集計 cells whose formulas reference single 名簿 columns: header texts of those columns (deduplicated) */
      summaryRefs: summaryRefs(spec.base.summary),
      students: spec.students.map((s) => ({
        row: baseRow(s.index),
        id: s.id,
        name: s.name,
        kana: s.kana,
        uni: s.uni,
        email: s.email,
        note: s.note,
      })),
      sameNamePair: [baseRow(8), baseRow(37)],
    },
    interview: {
      file: FILES.interview,
      sheet: INTERVIEW.sheet,
      headerRow: 1,
      key: 'ID (受験番号 ⟷ 学生ID)',
      rows: map(iv, (m) => ({ idKind: m.idKind, resubmitted: m.resubmitted, extra: m.extra })),
      expected: {
        baseRows: n,
        matchedBaseRows: ivBase.size,
        unmatchedBaseRows: unmatchedBase(iv),
        extraRows: iv.filter((m) => m.extra).map((m) => m.row),
        duplicateIds: dupIds,
      },
    },
    contacts: {
      file: FILES.contacts,
      encoding: 'Shift_JIS',
      headerRow: 1,
      key: '氏名 (同姓同名はメールで区別)',
      rows: map(spec.contacts.meta),
      expected: {
        matchedBaseRows: spec.contacts.meta.filter((m) => m.baseIndex != null).length,
        unmatchedBaseRows: unmatchedBase(spec.contacts.meta),
        extraRows: spec.contacts.meta.filter((m) => m.baseIndex == null).map((m) => m.row),
      },
    },
    aptitude: {
      file: FILES.aptitude,
      sheet: APTITUDE.sheet,
      headerRow: 1,
      key: 'フリガナ ⟷ 氏名カナ (学籍番号は学生IDと無関係)',
      rows: map(spec.aptitude.meta, (m) => ({ kanaKind: m.kanaKind })),
      expected: {
        matchedBaseRows: spec.aptitude.meta.filter((m) => m.baseIndex != null).length,
        unmatchedBaseRows: unmatchedBase(spec.aptitude.meta),
        extraRows: spec.aptitude.meta.filter((m) => m.baseIndex == null).map((m) => m.row),
      },
    },
    attendance: {
      file: FILES.attendance,
      sheet: ATTENDANCE.sheet,
      headerRow: ATTENDANCE.headerRow,
      key: '氏名＋大学 (同姓同名のため複合キーが必要)',
      rows: map(spec.attendance.meta, (m) => ({ uni: m.uni })),
      expected: {
        matchedBaseRows: spec.attendance.meta.filter((m) => m.baseIndex != null).length,
        unmatchedBaseRows: unmatchedBase(spec.attendance.meta),
        extraRows: spec.attendance.meta.filter((m) => m.baseIndex == null).map((m) => m.row),
      },
    },
  };
}

export function buildSpec() {
  const students = makeStudents();
  const interview = interviewRows(students);
  const aptitude = aptitudeRows(students);
  const attendance = attendanceRows(students);
  const contacts = contactRows(students);
  return {
    students,
    base: {
      title: '2027年卒 夏季インターン 応募者名簿',
      note: '※ 年齢はマスタシートの基準日時点。学生IDは6桁の文字列です。',
      roster: ROSTER,
      rows: rosterRows(students),
      master: masterBlock(students),
      summary: summaryBlock(students),
    },
    interview: { ...INTERVIEW, rows: interview.rows, meta: interview.meta },
    contacts: { headers: CONTACTS.headers, rows: contacts.rows, meta: contacts.meta },
    aptitude: { ...APTITUDE, rows: aptitude.rows, meta: aptitude.meta },
    attendance: { ...ATTENDANCE, rows: attendance.rows, meta: attendance.meta },
  };
}
