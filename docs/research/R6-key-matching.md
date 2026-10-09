# R6 — Automatic join-key detection and robust Japanese matching

Research and prototype for Excel統合ツール (Excel Merger). Everything below was run on this machine (Windows 11, Node 24.16, npm 11.13). The prototype lives in:

`C:\Users\81802\AppData\Local\Temp\claude\D--20260705-----------\2d64adba-cbb0-49f4-bf69-6262ff0972bb\scratchpad\research\r6`

In this report that folder is called `r6/`.

---

## 0. Summary

**Recommendation.** Port `r6/src/{normalize,detect,match,api}.ts` and `r6/dict/{key-synonyms,itaiji-name}.json` into the app as `src/core/keys/`. It is written in plain TypeScript with no DOM. Its only runtime dependency is **fastest-levenshtein**. Minified it is about 73 KB, or 27 KB gzipped, including both dictionaries. Run it in a Web Worker. It reads the cell values that the existing ExcelJS/SheetJS reader already produces, and it never writes to a workbook. That makes it independent of the "preserve formulas" export engines.

**How it works.**

1. **Header-row detection.** It scans the first 30 rows and picks the one with the most dictionary synonyms. Title rows and 横持ち (wide) sheets are handled.
2. **Column typing.** Each column gets a score from its header, a score from its values, and a uniqueness figure. These are combined with fixed formulas (§5).
3. **Candidate search between two sheets:**
   - same-type column pairs, e.g. 応募者ID ⟷ 受験番号 ⟷ 学生ID, or 氏名 ⟷ 姓＋名 built as a virtual column;
   - header-independent **value-overlap discovery** that is aware of primary-key/foreign-key relations (it found 職場ID ⟷ 応募者コード in the samples);
   - **composite keys** such as 氏名＋大学, 氏名＋生年月日, 氏名カナ＋大学 and 姓＋名, with dominated combinations pruned.
4. **Tiered join (fallback chain).** The order is ID → 氏名-based composite → カナ-based composite → メール → 氏名 → カナ.
   - If a higher tier's keys are present on both sides and differ, that pair is **vetoed** in the lower tiers. For example, a matching name with a different ID is never auto-merged.
   - A key that hits more than one row is reported as **ambiguous**. It is never picked at random.
5. **Cross-checks on matched pairs.** Examples: ID一致・氏名不一致 (error), 異体字の違い (info), 似た字 斎/斉 (warn), フリガナ不一致, 生年月日不一致, and discovered column pairs that disagree.
6. **Fuzzy suggestions for unmatched rows.** These are never applied automatically. They use kana, e-mail, phone, 異体字 and confusable folding, Levenshtein distance 1, and ID typos (Damerau distance 1), with Fellegi–Sunter-style additions for agreeing fields and penalties for disagreeing ones.

**Results.**

- On the interview-scheduler samples, every list-style sheet pair joins at 100% of the expected rows.
- On a synthetic "tricky" set the prototype made **0 wrong automatic matches**, and **every true pair it did not auto-match was surfaced as a top suggestion or an ambiguous choice**:

  | Base ⟷ other | Auto-matched correctly | Surfaced as suggestion or ambiguous |
  |---|---|---|
  | A ⟷ B | 24 of 29 | 5 |
  | A ⟷ C | 10 of 10 | 0 |
  | A ⟷ D | 20 of 20 | 0 |
  | B ⟷ C | 9 of 10 | 1 |

- 53 of 53 normaliser and header tests pass.
- A 20,000-row sheet takes about 1 s to read and about 1 s to profile, then about 1 s for candidates and the join.

---

## 1. Prior art and libraries (installed and tested in `r6/`)

| Library / work | Version | Verdict | Evidence |
|---|---|---|---|
| **`String.prototype.normalize('NFKC')`** (built in) | – | **use** (base of every normaliser) | Converts full-width ASCII to half-width, and half-width kana to full-width while composing dakuten (`ｶﾞｯｺｳ` → `ガッコウ`), plus ㈱, ①, № and CJK compatibility ideographs such as 塚 U+FA10. It does **not** fold 異体字: 髙/高, 﨑/崎 and 邊/辺 all stay distinct (verified). |
| **itaiji** (mziyut/itaiji-node, MIT; table from camelmasa/itaiji) | 1.2.0 | **use (vendored and patched)** | 1,404 old/variant → new pairs, 36 KB. Covers 髙﨑邊邉齋齊濵濱澤廣國德嶋嶌眞櫻瀧惠冨龍萬圓實壽與淺關會賴靜黑兒將條峯嵜曾桒栁槇籠蘆鷗瀨彌禮淸芲恆舘鐵驒萠凜堯遙螢聰穗亞. It misses 𠮷 and 﨔. It has a few **dangerous pairs for names**: 壬→閏 (wrong), 庄→荘, 苑→園, 椙→杉, 涌→湧, and 碕→埼 (should be 崎). `r6/scripts/build-itaiji.mjs` removes or overrides these and writes `dict/itaiji-name.json`: 1,415 pairs, 17 KB. |
| **itaiji-normalize** (tomatomerde; IPA MJ Shrink Map, code MIT, data CC BY-SA 2.1 JP) | 0.2.0 | **reference only** (mined at build time) | 30k source characters, but **2.5 MB minified / 582 KB gzipped** when bundled. MJ reduction does not fold 槇/槙, 籠/篭, 蘆/芦, 遙/遥, 凜/凛, 堯/尭 or 驒/騨 (both sides are JIS). Its 15k extra pairs include noise for names (嗎→罵, 蘤→花). I used it only to pick the surrogate-pair name kanji: 𠮷→吉, 𡈽→土, 𣘺→橋, 𥔎→崎, 𦚰→脇, 㟢→崎, 㠀→島 and so on. |
| @geolonia/normalize-japanese-addresses | 3.1.3 | reference only | Its `jisDai2.ts` is a JIS 第2→第1水準 and 旧→新字体 list; itaiji is a superset for our purpose. It also has address-specific variants ([之ノの], [ヶケが], 條/条, 渕/淵), which I reused as the **confusable** table used for suggestions only. |
| **wanakana** | 5.3.1 | **optional** (romaji name columns only) | `toKatakana('やまだ')` works. It **does not convert half-width kana** (`toKatakana('ﾔﾏﾀﾞ')` stays `ﾔﾏﾀﾞ`), so use NFKC for that. `toRomaji('サトウ ショウタ')` gives `satou shouta`. With the fold in §4 this matched *SATOH Shohta*, *Sato Shota*, *Ono Yuko* (オオノ), *Homma* (ホンマ) and *Jun-ichi* in 6 of 6 cases (`results/roman-sjis.txt`). 6.9 KB gzipped. |
| jaconv (kazuhikoarase) | 1.1.2 | not needed | Has `toZenKana`, `toHiragana`, `toKatakana` and `toHebon`. `toHebon` expects hiragana (katakana input comes back unchanged). NFKC plus a one-line hiragana→katakana shift covers what we need. |
| moji / japanese-string-utils / jaco | 0.5.1 / 0.5.0 / 5.0.0 | not needed | All give the same half-width/full-width and hiragana/katakana results as NFKC plus the code-point shift. |
| **encoding-japanese** | 2.4.0 | **use for CSV** | `Encoding.detect(buf)` returned `SJIS` for a cp932 CSV containing 髙 and ﾀｶﾊｼ, and `UTF8` for the BOM sample. Then `Encoding.convert(buf,{to:'UNICODE',from,type:'string'})` and `XLSX.read(text,{type:'string',raw:true})`. `raw:true` keeps `00123` as text. |
| **fastest-levenshtein** | 1.0.16 | **use** | 4 million comparisons of short names take 333 ms. `distance('田中憲一','田中健一')` = 1. |
| jaro-winkler | 0.2.8 | reject | Scores 佐藤/加藤 at 0.67 and 田中憲一/田中健一 at 0.87. It is too lenient on 2–4 character CJK names; "Levenshtein = 1 with length ≥ 3" is sharper. |
| fuse.js | 7.5.0 | optional, for the grid search box only | Bitap scoring on short CJK strings is flat: 田中健一 scores 0.5, 0.5 and 0.51 against three different names. Not suitable for matching. |
| fuzzball (TheFuzz port) | 2.2.6 | reject | `token_sort_ratio` is only useful for Latin names; `normName().latinSorted` does the same thing in one line. |
| kuroshiro / kuromoji | – | reject | Reading kanji needs a 20 MB dictionary and is still a guess for names. |
| **Auto-Join** (Zhu, He, Chaudhuri, VLDB 2017) and **Auto-FuzzyJoin** (Li et al., SIGMOD 2021) | – | ideas | Join by learned transformations, and fuzzy-join configuration without labels. Taken: pick the *most precise* normalisation first, then loosen only for suggestions. |
| **Auto-Suggest** (Yan & He, SIGMOD 2020) and **Auto-BI** (2023) | – | ideas | Join-column prediction features: distinct-value ratio (key-ness), value-overlap / containment, left-ness, value type, header similarity, PK–FK direction. These map to `uniqueness`, `containment`, header score, value score and the `lookup` mode. |
| **Power Query fuzzy merge** | – | ideas | Token Jaccard, default threshold 0.80, ignore case, "combine text parts", transformation table. The transformation table is the same idea as our itaiji and alias tables. Its threshold approach is not used for automatic merges. |
| **OpenRefine clustering** | – | ideas | Key collision: fingerprint (trim, lowercase, strip punctuation, sort unique tokens) and n-gram fingerprint; nearest neighbour with Levenshtein radius 1 and blocking of 6 characters. `latinSorted` is the fingerprint method. Suggestions use radius 1. |
| csvkit csvjoin | – | idea | Exact-match only, and the user names the columns. That is our "詳細画面で手動追加" fallback. |
| splink / dedupe | – | ideas | Fellegi–Sunter m/u weights, blocking, and per-field agreement/disagreement. The suggestion score is a hand-tuned version of this. There is no EM training, because data sizes and labels are too small. |

The reference app already contains the same primitives in `interview-scheduler/src/core/parse.ts`: `matchingIdText`, `digitsKey` (00123 ≡ 123), `nameKey` (spaces removed) and `isLossyExponentId`. The prototype generalises them. The app can keep one shared implementation.

---

## 2. Pipeline

```
workbook (ExcelJS/SheetJS values; cached formula RESULTS, never formulas)
  └─ per sheet: detectHeaderRow → Table{headers, rows, display(.w), numFmt, dateFmt, excelRow}
       └─ profileTable: per column → headerScore[type], valueScore[type], typeScore, type, uniqueness, keys[] (type-normalised)
            └─ findCandidates(base, other):
                 singles   (same type family: idFamily / name / kana / email / phone / …)
                 discovered(value overlap, PK–FK aware, any header)
                 composites(person part + {大学, 学部, 生年月日, メール, 電話, カナ}, 2–3 parts, dominated pruned)
            └─ planTiers → [ID, 氏名＋大学, 氏名カナ＋大学, メール, 氏名, 氏名カナ]   (UI: default selection)
            └─ chooseMode → oneToOne | lookup
            └─ executeJoin → pairs(+flags), unmatched, duplicates, ambiguous, conflicts, suggestions
```

Files:

| File | Contents |
|---|---|
| `r6/src/normalize.ts` | per-type normalisers and value-shape predicates |
| `r6/src/detect.ts` | header dictionary compilation, column profiling and typing, context rules |
| `r6/src/match.ts` | pair statistics, discovery, composites, tier planning, join, cross-check, suggestions |
| `r6/src/api.ts` | JSON DTOs for the UI (`analyzeKeys`, `runJoin`) |
| `r6/src/table.ts` | SheetJS/CSV reader and header-row detection (the app replaces it with its own reader) |

---

## 3. Header synonym dictionary — `r6/dict/key-synonyms.json`

Headers are compared after `headerNorm()`. The steps are:

1. NFKC, then lower-case.
2. Strip whitespace and the symbols `※*＊★☆◆◇■□●○・:：_-‐()（）［］[]【】「」『』<>〈〉《》.,、。?？!！'"/／`.
3. Strip note words: 必須, 任意, 半角(英数), 全角(カナ), ハイフンなし, スペースなし, 記入例, を入力してください, ください, …
4. Strip a question prefix such as `Q3.`, `問3` or `1.`.

Examples: `氏名（カナ）` → `氏名カナ`, `E-mail` → `email`, `携帯番号（ハイフンなし）` → `携帯番号`, `Q1. お名前（フルネーム）を入力してください` → `お名前フルネーム`.

Structure, per type:

- `rules`: a list of `{re, score, subtype?}`. The highest matching score wins.
- `negative`: a list of regexes that veto the type.
- `nonKeyHeaders`: never a key. These veto value-only typing: 備考, 区分, 理由, 状態, 結果, 評価, 日程, 開始, 住所, … (without this, a 区分 column full of kanji words looked like 氏名).
- `entityWords`:
  - `subject`: 学生, 応募者, 受験者, 候補者, 参加者, …
  - `staff`: 社員, 担当者, 面接官, 職場, …
  - `other`: 保護者, 紹介者, 会社, …
- `family`: which types can be joined with each other.

Types and main rules (scores):

| Type | Exact / strong (1.0) | Weaker | Negative (veto) |
|---|---|---|---|
| `id` | `id uid userid code コード key`; `(管理\|受験\|応募者?\|受付\|登録\|会員\|顧客\|個人\|候補者\|エントリー\|参加者\|回答者\|本人\|利用者\|受講者\|求職者\|申込者?\|志願者\|内定者\|整理\|識別\|問合せ)(番号\|no\|id\|コード\|code\|キー\|key\|number)`; English `(applicant\|candidate\|entry\|member\|customer\|user\|participant\|registration\|ref)(id\|no\|number\|code)` | `社員番号`-like → `employeeId` 0.95; `学生ID` 0.85; `*id / *コード` 0.7; `*番号 / *no` 0.45; contains id/番号/コード 0.35 | 電話, tel, fax, 携帯, 郵便, 口座, 部屋, 内線, カード, 免許, マイナンバー, パスポート, 学籍, メール, パスワード, 日程/日付/時間, 件数, 金額, 点数, スコア, 順位, 学年, 年齢, 職場id, 部署/大学/会社/商品…コード |
| `studentId` | `(学籍\|学生)(番号\|no\|id\|コード\|code\|nr)`, studentid, studentno | 学籍, 学生証 0.7 | 電話, メール |
| `orgCode` (not a person key) | `(職場\|部署\|部門\|会社\|企業\|大学\|学校\|学部\|店舗\|取引先\|拠点\|事業所\|支店)(id\|コード\|code\|番号\|no)` 0.9 | | |
| `rowNumber` | `no № # 連番 行番号 通し番号 番号 項番 seq` 0.9, **only if the values are 1..N** (sequence ≥ 0.8 starting at 0 or 1) | | |
| `name` | `氏名 お名前 名前 姓名 フルネーム 本名 漢字氏名 氏名漢字 name fullname`; `(学生\|応募者\|受験者\|…)(の)?(氏名\|名前\|姓名\|名\|フルネーム)` 0.95 | contains 氏名/名前/姓名/フルネーム 0.75; `name` 0.6 | カナ/かな/フリガナ/ヨミ/読み/kana/ローマ字/英字/romaji; 会社/企業/法人/団体/大学/学校/学部/学科/部署/職場/店舗; 担当/面接官/上司/保護者/親/緊急連絡/紹介者/推薦者/作成者/更新者/記入者; ファイル/シート/商品/品名/項目/イベント/セミナー/案件/職種/役職/資格; username/ユーザー名/ニックネーム/旧姓/通称 |
| `lastName` / `firstName` | `姓 氏 苗字 名字 lastname familyname surname` / `名 めい 下の名前 firstname givenname` | context rule: `名前` directly right of 姓/苗字 with values ≤ 3 characters → `firstName` | kana words |
| `nameKana` | `フリガナ ふりがな カナ かな ヨミ よみ 読み よみがな 読み仮名 カナ氏名 氏名カナ 氏名フリガナ 名前カナ お名前ふりがな セイメイ kana furigana yomi` | contains any of those 0.8 | 会社, 大学, 学部, 職場, 担当, 住所, 姓/セイ/名$/メイ$ (these are split-kana columns) |
| `lastNameKana` / `firstNameKana` | `セイ 姓カナ 姓フリガナ カナ姓 苗字カナ` / `メイ 名カナ 名フリガナ カナ名` | | |
| `nameRoman` | ローマ字, 英字, 英語表記, アルファベット, romaji, englishname | | |
| `email` | `メール メアド mail email emailaddress mailaddress メールアドレス eメール 電子メール pcメール 携帯メール` | contains メール/mail 0.85 | 送信, 配信, 件名, 本文, 日時, フラグ, 可否, 希望 |
| `phone` | `電話(番号) tel telephone phone 携帯(番号\|電話) mobile` | contains 電話/tel/phone/携帯 0.85; `連絡先` 0.5; `fax` 0.3 | 会社, 大学, 緊急, 保護者, 職場, 担当, 受付時間 |
| `university` (composite part) | `大学 大学名 学校 学校名 大学院 出身校 在籍校 所属大学 所属 university school college` | 0.85 / 0.7 | 学部, 学科, 研究科, 専攻, 学年, 卒業, 入学, コード, 住所, 区分, 偏差値 |
| `faculty` | `学部 学科 学部学科 研究科 専攻 faculty department major` | 0.8 | コード, 系統 |
| `birthDate` | `生年月日 誕生日 生年 生まれ birthday birthdate dob` | 0.85 | |
| `gender` (cross-check) | `性別 gender sex` | | |
| `graduation` (composite part) | 卒業年, 卒業予定, 卒年, graduation | 0.9 | |

Headers longer than 14 characters after normalisation are capped at 0.7. These are form questions such as お名前をフルネームで教えてください.

---

## 4. Normalisation per type (`r6/src/normalize.ts`)

**Common base.**

```ts
nfkc(s) = s.replace(/[\uFE00-\uFE0F]|\uDB40[\uDD00-\uDDEF]/g,'')   // SVS/IVS (葛󠄀 → 葛)
           .replace(/[\u200B-\u200F\u202A-\u202E\u2060-\u2064\uFEFF\u00AD]/g,'') // zero-width, BOM, soft hyphen
           .normalize('NFKC').trim()
```

**Blank-like values.** These count as blank:

- `null` or empty;
- whitespace only (including U+3000);
- Excel error literals such as `#N/A` and `#REF!`;
- placeholders `- ― ー なし 無し 不明 未定 未記入 未入力 該当なし n/a null none ? .`

| Type | Key (automatic matching) | Extra forms (flags and suggestions only) |
|---|---|---|
| **ID** (`normId`) | Steps: (1) if the cell is a **number** whose display text `.w` is `0\d+`, use the display text — number 115 with format `000000` shows 000115; (2) otherwise take `cellText`, where an integer becomes a `BigInt` string with no exponent or `.0`; (3) NFKC, remove all whitespace, map every hyphen variant `‐‑‒–—―−⁃﹘﹣－⸺⸻ー─━ｰ` to `-`, upper-case, strip a leading `'`, turn `123.0` into `123`, and expand `8.00E+09` when it is lossless. **Key** = all-digit IDs without leading zeros (`00123`≡`123`); otherwise the exact form with `-` and `_` removed (`TEST-0001`≡`TEST0001`). | `exact` form, used for the "ID表記の違いを補正" flag. `lossyExponent` = the ID's real digits are lost, so show a warning. |
| **氏名** (`normName`) | Steps: (1) NFKC; (2) cut parenthesised notes into `note` (`山田 花子（旧姓：佐藤）`); (3) remove all whitespace; (4) remove a trailing honorific 様/さま/殿/さん/くん/君/ちゃん/氏 when the name is longer than 2 characters; (5) remove `・･·•=＝.,、'` (ジョン・スミス≡ジョンスミス); (6) expand 々 and 〻 (佐々木≡佐佐木); (7) lower-case Latin; (8) unify long-vowel look-alikes inside katakana; (9) **異体字 fold** (itaiji-name, 1,415 pairs). Key = `folded`. | `strict` (not folded), used to show an "異体字の違い" info flag. `latinSorted` = sorted lower-case tokens (Smith John ≡ JOHN SMITH). `confusableFold` (斉齊齋→斎, 邉邊→辺, 澤→沢, 嶋嶌→島, 之の乃→ノ, ヶ→ケ, ヵ→カ, 籐→藤 …) is **for suggestions only**. |
| **カナ** (`normKana`) | NFKC (ﾔﾏﾀﾞ→ヤマダ), hiragana→katakana (`+0x60` over U+3041–3096 and ゝゞ), remove whitespace and `・`, turn `‐―−－─━ｰ~〜～-` into `ー`, remove a trailing サマ/サン/様. Key = `strict`. | `loose`: small→large kana (ァィゥェォッャュョヮヵヶ), ヂ→ジ, ヅ→ズ, ヴ→ブ, ヲ→オ, ヰ→イ, ヱ→エ, `ー` removed, so キョウコ≡キヨウコ. Suggestions only. |
| **メール** | NFKC, remove whitespace and `mailto:`, lower-case, ＠/(at)/[at] → `@`, `。`→`.`, validate `^[A-Za-z0-9._%+\-']+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$`. | `emailLoose`: drop `+tag`; for gmail also drop dots. Suggestions only. |
| **電話** | NFKC, strip a `tel:` prefix and an extension (内線/ext), keep digits only. `+81` or `0081` becomes `0`, and `+81 (0)90` also works. A **number cell** with 9–10 digits gets its **leading 0 restored** (flagged). Valid only when `^0\d{9,10}$`. | |
| **生年月日** | Rules in order: `Date` → ISO date (UTC; SheetJS and ExcelJS give UTC midnight). A number is an **Excel serial** only when the column has a date format (`XLSX.SSF.is_date(numFmt)`) or the header says 生年月日; 1900 system, serial 60 is the fake 1900-02-29. A `yyyymmdd` integer is also accepted. Strings: `2001/4/1`, `2001-04-01`, `2001.4.1`, `2001年4月1日(生)`, `20010401`, `m/d/yyyy`, **和暦** `令和/平成/昭和/大正/明治` or `R/H/S/T/M` + year or `元` + month + day (`H13.4.1`, `平成13年4月1日`, `令和元年5月1日`, `R1-5-1`). The result is validated, so `2001/2/30` gives an empty key. | |
| **大学** (`normOrg`) | NFKC, drop parentheses and whitespace, lower-case, **itaiji fold** (慶應≡慶応), drop the prefix 国立大学法人/公立大学法人/学校法人…/私立/国立…, `…大学院` → `…大学`, drop `キャンパス`, then a small alias table (東大, 京大, 阪大, 早大/早稲田, 慶大/慶応/SFC, 一橋, 東工大, 科学大, 上智, 明大, 立教, 中大, 法大, 同志社, 立命館, 関学, 関大, 筑波, 神大, 横国, 理科大, ICU). | |
| **generic** (discovery) | NFKC, remove whitespace, hyphens → `-`, upper-case, `123.0`→`123`, all-digit values lose leading zeros, itaiji fold. | |
| romaji ⇄ kana (optional) | `wanakana.toRomaji(kana)` then fold: `/[^a-z\s]/→''`, `(o)(u\|o\|h)(?![aiueo])→o`, `uu→u aa→a ii→i ee→e`, `m(?=[bmp])→n`, sort tokens. | Tested 6 of 6 (`r6/scripts/test-roman.mjs`). |

Kana is **not** folded inside the name key: 山田さくら and 山田サクラ are different strings. Kana matching is done on the kana column.

---

## 5. Column typing (`profileColumn`), with exact formulas

All ratios are computed over the non-blank cells, using an even sample of up to 2,000 cells. Keys and uniqueness use every row.

```
ratios: number, date (Date | number with date numFmt | parsable date text), email, phone, kana (strict kana 2–20),
        kanjiName  /^[Han々〆ヶ]{1,6} ?[Han Hiragana Katakana ー]{1,6}$/u with ≥1 Han,
        latinName  /^[A-Za-z][A-Za-z'.\-]*( +[A-Za-z][A-Za-z'.\-]*){1,3}$/,
        textish (2–30 chars, not numeric/date/email/url), univ /(大学|大学院|短期大学|短大|高等専門学校|高専|専門学校|university|college|institute)/,
        faculty /(学部|学科|研究科|専攻|学群|学類|学環|学域|学府|課程|コース|faculty|department)/,
        shortKanji (1–4 Han), idShape (int ≥0 | [A-Za-z0-9][A-Za-z0-9-_/.#]{1,23} with a digit),
        strongId (int ≥ 1000 | letters+digits | 0-padded | ≥ 4 digits), gender, sequence (+1 steps, sequence starts at 0/1),
        birthYear (share of years 1940–2015), idConsistency (share of the most common signature, e.g. AAAA99999)
gUniq  = distinct/nonBlank on the generic key
uniqF  = clamp((gUniq − 0.5)/0.45, 0, 1)          nameUniqF = clamp((gUniq − 0.3)/0.6, 0, 1)

valueScore:
  id = studentId = idShape × (0.5 + 0.5·idConsistency) × uniqF × (1 − date)
  name        = (shape + 0.4·max(0, textish − shape)) × (1 − 0.8·kana) × (1 − univ) × (1 − faculty) × (1 − email),  shape = max(kanjiName, latinName)
  lastName    = firstName = shortKanji × (1 − kana)
  nameKana    = kana;  nameRoman = latinName;  email = email;  phone = phone × (1 − email)
  university  = univ;  faculty = faculty;  birthDate = date × (0.4 + 0.6·birthYear);  gender = gender

typeScore(T):
  if headerScore h ≥ 0.5:   s = 0.6·h + 0.4·v;   if v < VETO[T] and nonBlank ≥ 3 → s ×= 0.6
        VETO = {email .3, phone .3, birthDate .3, nameKana .3, name .2, id .2, studentId .2, university .2, faculty .2}
  else (value-only):        vv = v × nameUniqF   (name, kana, roman, 姓/名, email, phone)
                            vv = v × strongId    (id family)
                            s  = 0.75·vv·(vv ≥ 0.8 ? 1 : 0.6) + 0.25·h
                            if the header is negated for T, or is a nonKeyHeader → s = min(s, 0.3)
                            value-only lastName/firstName/split-kana/studentId/employeeId/orgCode/rowNumber/faculty/graduation → min(s, 0.3)
  rowNumber: header No/番号 AND sequence ≥ 0.8 → 0.95
type = argmax typeScore; 'unknown' if max < 0.35.   confidence: high ≥ 0.75, medium ≥ 0.55, low otherwise.
keyScore (single-column joinability, used for sorting) =
  key types:            typeScore × (0.4 + 0.6·uniqueness) × (0.5 + 0.5·fill)
  composite-part types: typeScore × 0.5 × (0.5 + 0.5·fill)
  rowNumber / orgCode / unknown: 0
```

**Context rules.**

1. `名前` directly right of `姓` or `苗字`, with 80% or more of its values being 3 characters or fewer and containing no space, becomes `firstName`.
2. Staff-side columns (社員, 担当, 職場, 面接官) are flagged "本人のキーではない可能性" only when the same sheet also has a subject-side or neutral key column.

**Header row detection** (`r6/src/table.ts`):

- **Pass 1.** Scan the first 30 rows and take the row with the most cells that match a synonym with score ≥ 0.7, or a nonKeyHeader. Ties go to the earliest row. The row needs at least 2 non-empty cells and at least one non-empty row below it.
- **Pass 2**, used only when no row has any synonym. Score each row as +1 per short text cell, +0.5 per long form-question cell, −1 per number, −0.5 per ID-shaped value (`TEST00001`), and −0.5 per answer-like cell (`10:00`, ○×, NG). Take the highest.
- Merged header cells take the top-left value.

Before this rule, the old-format CSV picked row 24 as the header because its many short answer cells outscored the real header. The synonym-first rule fixed it.

**Wide (横持ち) layout hint.** If the header row, or the row above it, has 5 or more cells that are ID-shaped or kanji-name-shaped, make up at least 40% of the columns, and are *not* dictionary header words, the sheet gets the message 「横持ちレイアウト。行方向の結合キーは検出できません」. This fires on 人事回答, 学生回答 and 職場回答 `Sheet1`. It does not fire on title-row sheets.

---

## 6. Cross-file candidates and statistics (`match.ts`)

**Match-rate definitions**, for base A and other file B, per key pair:

```
matchRate (UI "一致率")  = |{b ∈ B non-blank : key(b) ∈ keys(A)}| / |B non-blank|      (spec: rows found in the base file)
coverageA                = |{a ∈ A non-blank : key(a) ∈ keys(B)}| / |A non-blank|
uniqA/uniqB              = distinct/non-blank ;  dupRows = rows whose key occurs > 1 ;  blank = rows with an empty key
ambiguousB               = B rows whose key hits > 1 A row
cardinality              = 1:1 | 1:N | N:1 | N:M   (from uniq < 1)
score (ranking)          = max(matchRate, coverageA) × min(uniqA, uniqB)² × √(min fill)
verdict                  = 良好 ≥ 0.8 · 一部一致 ≥ 0.5 · 不一致多数 ≥ 0.1 · 使用不可 < 0.1   (on max(matchRate, coverageA))
```

Both directions are shown because "base" is arbitrary. For example 人事回答.備考 (8 rows) against 名簿 (44 rows) gives matchRate 18.2% but coverage 100%, which is correct.

**Single candidates.** All pairs whose types are in the same family:

- `idFamily` = id, studentId, employeeId;
- `name` = name, or 姓＋名 **virtual** — the concatenated folded keys, so 姓+名 compares equal to a 氏名 column;
- `kana` = nameKana, or セイ＋メイ virtual;
- email, phone, university, faculty, birthDate, …

`orgCode` is excluded because it is not a person key.

**Discovery by value overlap** (header-independent, PK–FK aware):

- Eligible columns: not rowNumber, date or gender; distinct ≥ 3; fill ≥ 0.5; not a nonKeyHeader.
- The pair needs max(uniqA, uniqB) ≥ 0.9 (one side is a key) and must not already be a typed candidate.
- Skip the pair when both columns are 1..N sequences.
- `inter` = |distinct(A) ∩ distinct(B)| on the generic key. `containment = max(inter/|A|, inter/|B|)`.
- Accept when `inter ≥ min(3, smaller distinct count)` and `containment ≥ 0.3`. Rank by `containment × max(uniqA, uniqB)²`.
- Found in the samples: 職場ID (uniq 20%) ⟷ 応募者コード (100%) = N:1 lookup; 職場名 ⟷ 漢字氏名; マッチング職場 ⟷ 漢字氏名.
- Before the fill ≥ 0.5 rule it wrongly offered 備考 ⟷ 備考 (3 of 30 filled).

**Composite search.**

- Parts: the best column per signature among name, kana, roman, university, faculty, birthDate, email, phone, gender and graduation that exists in **both** files.
- Combinations are one person part (name, kana or roman) plus one or two other parts.
- Key = the parts joined with U+0001. It is empty if any part is blank.
- `score = pairScore − 0.03·(parts − 2) + bonus`, where bonus is name+university 0.05, name+birthDate 0.05, kana+university 0.04, name+kana 0.03, kana+birthDate 0.03, name+email 0.02.
- **Pruned:** combinations whose person part alone is already unique on both sides (the composite adds nothing), and supersets that are not more unique or better matching than a subset. This removed the useless 氏名カナ＋大学＋メール rows.

**Default plan (`planTiers`)**, which is the "基本の結合キー" pre-selected in the UI:

1. **ID**: the best idFamily pair with max(matchRate, coverage) ≥ 0.2, max uniq ≥ 0.9 and min uniq ≥ 0.5. Different ID systems do not overlap and are skipped automatically: 受験番号 ⟷ 学籍番号 = 0%.
2. **Composites**: the best one for each person part, in the order **氏名 → ローマ字 → カナ**. Each needs usable ≥ 0.1 and uniq ≥ 0.9 on both sides; exact per-row uniqueness is enforced at join time.
3. **メール**: only if uniq ≥ 0.98 on both sides, because family addresses are shared.
4. **氏名**.
5. **氏名カナ**.
6. If nothing typed is usable: the best discovered pair (≥ 0.3).

**Join mode (`chooseMode`).** If the first tier has uniqA < 0.9 and uniqB ≥ 0.99, the mode is `lookup` (VLOOKUP semantics: many base rows receive the same other row, e.g. 職場ID). Otherwise it is `oneToOne`.

---

## 7. Join semantics, cross-checks and suggestions

**Tier loop.** For each tier, and each unmatched A row with a non-blank key:

1. Look up the B rows that are still unused (in lookup mode, all B rows).
2. **Veto:** drop any candidate B row for which some *higher* tier has non-blank keys on both sides that differ. This is recorded as a conflict: "氏名＋大学は一致するがIDが異なる". Effects:
   - Same name with a different ID becomes a suggestion, not a merge.
   - A kanji name that differs blocks a kana-tier merge, so 斉藤/斎藤 and 田中憲一/健一 are suggested and not auto-merged.
   - A blank key does not veto, so a missing ID falls back to the name tiers.
3. Exactly one candidate, and the A key unique among unmatched A rows (1:1 mode) → match.
4. More than one candidate → **ambiguous**. Keep it for lower tiers; if still unresolved, list it with its candidates.
   - Example: 小林大輝 submitted twice in B.
   - The UI should offer 「先頭を採用 / 最終行を採用 / 手動で選ぶ」.

**Cross-checks on every matched pair.** Only columns belonging to the **same entity** as the key are compared, so 学生名 is not checked against a 職場ID match.

| Check | Level | Message |
|---|---|---|
| Name: strict equal | – | none |
| Name: equal only after 異体字 folding | info | 異体字の違い |
| Name: same Latin tokens | info | 姓名の順序違い |
| Name: equal after confusable fold | warn | 似た字の違い |
| Name: Levenshtein 1 with length ≥ 3 | warn | 氏名1文字違い |
| Name: anything else | **error** | ID一致・氏名不一致 |
| Kana differs | info if the loose forms are equal, otherwise warn | フリガナ不一致 |
| Birth date differs | warn | 生年月日不一致 |
| E-mail differs | info | メールが異なる |
| A discovered column pair with ≥ 0.8 match rate disagrees | warn | 「X」と「Y」が異なる |
| ID exact forms differ but digit keys match | info | ID表記の違いを補正 |

The UI should **aggregate flags by kind**. `api.ts runJoin().flags` does this (e.g. `ID表記の違いを補正: 20件`, `異体字の違い: 4件`).

**Suggestions** for unmatched rows. They are never applied automatically; the user accepts them.

- Positive signal (take the maximum):

  | Score | Signal |
  |---|---|
  | 0.90 | same kana (strict) |
  | 0.88 | same name (異体字 folded) |
  | 0.85 | same e-mail (loose) |
  | 0.85 | same Latin tokens |
  | 0.80 | confusable glyphs (斎/斉, ノ/之) |
  | 0.80 | same kana (loose) |
  | 0.75 | same phone |
  | 0.70 | name Levenshtein = 1 (length ≥ 3) |
  | 0.55 | ID one typo (Damerau distance 1, same length, ≥ 3 digits) |

- Adjustments:

  | Adjustment | Condition |
  |---|---|
  | +0.05 | same university; same birth date; same e-mail or phone (when not already the base signal); ID one typo on top of a stronger signal |
  | −0.35 | both IDs present and different, but only when an ID tier exists (受験番号 vs 学籍番号 are different numbering systems and are not compared) |
  | −0.15 | birth date differs |
  | −0.10 | university differs |

- Cap at 0.99. Show when ≥ 0.60. Show the top 3 per base row.

The weights started from the Fellegi–Sunter intuition and were tuned on the synthetic set. Before the disagreement penalties, 同姓同名 at another university was suggested at 0.90. After them it drops below the true pair, and there are 0 wrong suggestions.

---

## 8. UI DTO (`r6/src/api.ts`)

`analyzeKeys(base, other)` returns an `AnalysisDTO`:

```ts
interface ColumnDTO { col; header; type; typeLabel; confidence; typeScore; keyScore; uniqueness; fill; blank; dupRows; notes[] }
interface SheetDTO  { file; sheet; headerRow; rows; layoutHint?; columns: ColumnDTO[] }
interface CandidateDTO {
  id: 'single:1|0' | 'composite:2+4|1+3' | 'discovered:3|0'; kind: 'single'|'composite'|'discovered'; label;
  baseColumns: {col, header}[]; otherColumns: {col, header}[];
  matchRate; coverage; matchedOther; nonBlankOther; matchedBase; nonBlankBase;
  uniqBase; uniqOther; dupRowsBase; dupRowsOther; blankBase; blankOther; ambiguous; cardinality; score;
  verdict: '良好'|'一部一致'|'不一致多数'|'使用不可'; selected: boolean; tier?: number }
interface AnalysisDTO { base: SheetDTO; other: SheetDTO; mode: 'oneToOne'|'lookup'; candidates: CandidateDTO[] }
```

`runJoin(analysis)` returns:

```
{ mode, plan[{order,name}], matched, perTier{},
  unmatchedBase[{excelRow, preview[]}], unmatchedOther[…],
  duplicatesBase[{key,count}], duplicatesOther[…],
  ambiguous[{base, candidates[], tier}], conflicts[{base, other, message}],
  flags{kind: {level, count, examples[]}},
  suggestions[{base, other, score, reasons[]}] }
```

Examples: `r6/results/synthetic-A-vs-B.json`, `synthetic-A-vs-C.json`, `sample-matching-vs-student.json`, `sample-matching-vs-workplace.json`.

**Mapping to the requirements.**

| Requirement | UI behaviour |
|---|---|
| 3 | Main screen: show the `tier` 1 candidate of the ID family and of the name family as chips 「ID: 応募者ID ⟷ 受験番号 92%」「氏名: 氏名 ⟷ 名前 80%」. |
| 4, 5 | Detail screen: the `candidates` table (一致率, 重複, 空欄, 一意性, verdict colour) with checkboxes (`selected`) and drag to reorder the tiers. |
| 6 | Composites appear as their own rows. |
| 12 | Before joining, use `SheetDTO.columns` (blank and duplicate counts). After joining, use `runJoin` (matched per tier, unmatched lists with Excel row numbers, duplicates, ambiguous, conflicts, flags, suggestions with one-click accept). |

---

## 9. Test results

Re-run all of them with `npm test && npm run samples && npm run synthetic && npm run json` in `r6/`.

### 9.1 Normalisers and header dictionary — `r6/results/normalize-tests.txt`

53 passed, 0 failed. Cases covered:

- **IDs:** full-width IDs; numeric vs zero-padded (`123`≡`00123`); number with format `000000`; `8000000001` with no E-notation; lossy `8.00E+09`; `123.0`; `-`, `#N/A`, `なし` as blank.
- **Names:** spaces including U+3000; 髙/高, 渡邊/渡邉/渡辺, 山﨑/山崎/山嵜; 齋→斎 and 齊→斉 (斎≠斉 kept); 濵/濱; 澤廣國德嶋眞櫻瀧惠冨龍; 𠮷田; IVS 葛󠄀; 佐々木≡佐佐木; 様; 旧姓 note; Latin order; ジョン・スミス.
- **Kana:** half-width, hiragana and U+3000; long vowels `ユ-コ`/`ユ－コ`; loose forms.
- **E-mail and phone:** e-mail with full-width ＠; phone with `+81`, a numeric cell with the 0 lost, `03(1234)5678 内線12`, and full-width digits.
- **Dates:** serial 36982; serials 59 and 61 around the 1900 bug; H13.4.1, 平成13年4月1日, h13/04/01; 令和元年 and R1.5.1; S63.12.31; 4 Western formats; invalid 2/30; Date objects; serials accepted only with a date format.
- **Organisations and generic:** 慶應≡慶応≡慶應 (alias); 国立大学法人 prefix and 大学院 suffix; `00123`≡`123` in the generic key.
- **Headers:** 14 header groups, including 担当者名, 会社名 and 面接官氏名 → none; 番号 → rowNumber; 電話番号 → phone; 郵便番号 → none; 学生ID → studentId; form question → name 0.75; 職場ID and 大学コード → orgCode.

### 9.2 Interview-scheduler samples — `r6/results/samples.txt`

Twelve sheets from the 4 input xlsx files, the output xlsx and the old-format CSV. Profiling and all pairs take about 0.2–0.3 s.

**Detection.**

| Sheet | Header row | Detected keys |
|---|---|---|
| 人事回答 / 学生回答 / 職場回答 `Sheet1` | 2 | none, with **横持ちレイアウト** hint (row 1 = 8, 30 or 6 IDs across). Correct: persons are columns in this layout. |
| `備考` sheets (all three) | 1 | 応募者コード → ID 1.00 (uniq 100%); 漢字氏名 → 氏名 0.76 |
| 学生-職場マッチング | 1 | No → 行番号 (excluded); 学生ID → 学籍番号/idFamily 1.00; 学生名 → 氏名 0.73; 職場ID → 組織コード (uniq 20%, staff); 職場名 → not a person name |
| 学生回答_旧形式 CSV (UTF-8 BOM) | **1** (it was 24 before the fix) | 応募者コード → ID; 漢字氏名 → 氏名; 12 long question headers → unknown |
| 面接日程 名簿 | 1 | id → ID; 氏名 → 氏名; 区分 → unknown (was 氏名 before the nonKeyHeaders and uniqueness fix) |
| 面接日程 未配置 | 1 | 応募者コード → ID; 氏名 → 氏名 |
| 面接日程 調整用 / 枠組み | 1 | no keys (only 日程/開始/終了) |

**Cross-file results** (base A ⟷ other B):

| A ⟷ B | Plan | Match rate (B→A) / coverage of A | Result |
|---|---|---|---|
| マッチング ⟷ 学生回答.備考 | ID(学生ID⟷応募者コード) → 氏名 | 100% / 100% | 30 matched by ID, 0 flags |
| マッチング ⟷ 旧形式 CSV | ID → 氏名 | 100% / 100% | 30 |
| マッチング ⟷ **職場回答.備考** | **値一致（職場ID ⟷ 応募者コード）**, mode **lookup** | 100% / 100%, N:1 | 30 rows receive their workplace row. The typed pairs 学生ID⟷応募者コード and 学生名⟷漢字氏名 are 0% (使用不可). The second discovered pair 職場名 ⟷ 漢字氏名 is used as a consistency check: 0 disagreements. |
| マッチング ⟷ 面接日程.名簿 | ID → 氏名 | 68.2% / 100% | 30 matched; 14 staff and HR rows of 名簿 unmatched (expected) |
| 人事回答.備考 ⟷ 名簿 | ID → 氏名 | 18.2% / 100% | 8 matched |
| 職場回答.備考 ⟷ 名簿 | ID → 氏名 | 13.6% / 100% | 6 matched |
| 学生回答.備考 ⟷ 未配置, 名簿 ⟷ 未配置, … | ID → 氏名 | 100% | 30 |

### 9.3 Synthetic tricky set — `r6/results/synthetic.txt`

Generated by `r6/scripts/make-synthetic.ts`; truth is in `synthetic/truth.json`.

- **A 応募者名簿**: a title row, then the header on row 3; `応募者ID` as `000100` strings plus one number cell with format `000000`; ﾊﾝｶｸ kana; a phone number cell that lost its 0; Date cells.
- **B 一次面接結果**:
  - IDs: `受験番号` as numbers, full-width `０００１１４`, `'115'`, 4 missing, a duplicate submission, a transposed `000181`, and a new applicant 900.
  - Names: 高橋一郎 for A's 髙橋 一郎; 渡辺 惠; 山崎　翔; 斎藤 美咲 (齋); 斎藤 健太 for A's **斉藤**; typo 田中 健一 for 憲一 with no ID; 一之瀬 for 一ノ瀬 with no ID; 佐佐木; 鈴木 一郎 様; NGUYEN VAN AN; John Smith for Smith John; a completely different name (木村 拓也) on 中村 優子's ID.
  - 同姓同名: A has two 佐藤 翔太 at 早稲田 and 慶應; B has one of them with no ID.
  - Other fields: 慶応/慶應; +81 phone; 平成13年…, H13.4.1 and Date birth dates; upper-case e-mail.
- **C 学籍名簿**: 学籍番号 (a different ID system), 姓 / 名 / セイ / メイ split, 同姓同名 at different universities, and 5 people not in A.
- **D 謎の表**: header `参照` (unknown) holding A's IDs, plus random 0–100 scores under `データ`, which must not be typed as ID.

| Pair | Plan chosen | Auto-matched (correct / wrong) | Missed → surfaced | Wrong suggestions |
|---|---|---|---|---|
| A ⟷ B | ID → 氏名＋大学 → 氏名カナ＋大学 → 氏名 → 氏名カナ | **24 / 0** of 29 true pairs | 5 → all surfaced: 田中 typo 0.99, 小林 duplicate (ambiguous plus 0.99), 一之瀬 0.99, 井上 transposed ID 0.99 | 0 |
| A ⟷ C | 氏名＋大学(姓＋名＋大学) → … | **10 / 0** of 10 (同姓同名 resolved by 大学) | 0 | 0 |
| A ⟷ D | ID (`参照` typed as ID from values: strongId=1, score 0.75) | **20 / 0** of 20 | 0 | 0 |
| B ⟷ C | 氏名＋大学 → 氏名カナ＋大学 → … | **9 / 0** of 10 | 1 (斎藤/斉藤) → suggested 0.95. It was 0.60 while the 受験番号-vs-学籍番号 "IDが異なる" penalty was wrongly applied; see pitfall 11. | 0 |

Flags on A ⟷ B:

| Flag | Level | Count |
|---|---|---|
| ID表記の違いを補正 | info | 20 |
| 異体字の違い (髙橋/高橋, 渡邊恵/渡辺惠, 山﨑/山崎, 齋藤/斎藤) | info | 4 |
| 似た字の違い (斉藤/斎藤) | warn | 1 |
| **ID一致・氏名不一致** (中村 優子 / 木村 拓也) | **error** | 1 |
| フリガナ不一致 (same pair) | warn | 1 |
| 姓名の順序違い (Smith John / John Smith) | info | 1 |

Other results on A ⟷ B:

- Conflicts, meaning vetoed auto-merges that went to suggestions instead: 井上 (ID 118 ≠ 181), 田中 and 一ノ瀬 (name differs, kana equal).
- Composite pruning left 6 composite rows. 3-part combinations and combinations whose person part is already unique were removed.

### 9.4 Performance — `E_社員20000.xlsx` (20,000 rows × 5 columns)

| Step | Time |
|---|---|
| SheetJS read | 0.9–1.1 s |
| Profile | ~1.0 s, after sampling ratios to 2,000 cells (it was 3.0 s before) |
| Candidates + tiered self-join | 0.9–1.3 s, 20,000 matched |

Suggestions are O(unmatched A × unmatched B) with Levenshtein; 4 million pairs take ~0.33 s. Cap it: if the product is over 5 million, block on the first character of the kana or name, or on the university.

**Recommendation:** run `analyzeKeys` and `runJoin` in a Web Worker so the UI stays responsive.

### 9.5 Bundle size

| Bundle | Minified | Gzipped |
|---|---|---|
| Prototype matching core (esbuild ESM browser; normalize + detect + match + api + both dictionaries + fastest-levenshtein) | 73 KB | 27 KB |
| wanakana (optional) | 18 KB | 6.9 KB |
| itaiji-normalize | 2.5 MB | 582 KB |

itaiji-normalize would add 582 KB gzipped, so do not bundle it.

---

## 10. Pitfalls (found while testing, and what to do)

1. **Excel number formats that mean date.** SheetJS `cell.z` is `"General"` for plain numbers. My first date regex matched the `ge` in "General" (the 和暦 era token) and turned every numeric ID into a "date"; 受験番号 scored 0.39. **Use `XLSX.SSF.is_date(fmt)`** (or ExcelJS `numFmt` with the same check) and treat `General` as not a date.
2. **Zero-padded IDs.** Excel stores `00123` either as text or as the number 123 with format `00000`. Read the display text (`cell.w` / ExcelJS `cell.text`) for number cells whose display starts with 0, and key all-digit IDs without leading zeros. Keep the exact form so you can flag it.
3. **Large IDs.** Read integers with `BigInt(v).toString()`. `String(8000000001)` is fine, but `1e21` and above switch to exponent form. Detect **lossy E-notation IDs** (`8.00E+09` saved from CSV) and warn; they cannot be repaired.
4. **Phone numbers stored as numbers lose their leading 0** (9012345678). Restore it only for 9–10 digit values in a phone-typed column, and flag it.
5. **Header row is not row 1** (title rows, form exports, CSV with 12 long question headers). Choose by dictionary hits first and density second; the density-only rule chose a data row.
6. **横持ち (wide) sheets** (the interview-scheduler answer sheets): row-wise key detection is meaningless. Say so and point to the list-style sheet (`備考`). Offering 「行列を入れ替えて読み込む」 is a possible extra.
7. **Row-number columns** (`No` 1..N) overlap trivially across files. Exclude them, but only sequences that *start at 0 or 1*: ID ranges like `000100…000119` are also consecutive (a bug I hit that disabled discovery).
8. **Value-only typing false positives.**
   - A 区分 column of kanji words (受験者/職場/人事) looked like names. Fix: the nonKeyHeaders veto plus a name-uniqueness factor.
   - Random 0–100 scores looked like IDs. Fix: `strongId` (≥ 4 digits, zero-padded, or letters plus digits).
9. **Remark columns** (備考 filled on 3 of 30 rows) looked like a 100%-unique key. Discovery requires fill ≥ 0.5 and not a nonKeyHeader.
10. **Two entities in one sheet** (学生-職場マッチング: 学生ID/学生名 and 職場ID/職場名). Record the header entity and cross-check only same-entity columns. Without this, every workplace lookup produced a false 「ID一致・氏名不一致」 (学生名 vs 職場担当).
11. **Different ID systems** (受験番号 vs 学籍番号) never overlap. Do not use them for the ID tier or for the "IDが異なる" penalty in suggestions; the penalty applies only when an ID tier exists.
12. **Kana and kanji tiers.** Put **kanji-name tiers before kana tiers**, so a kanji mismatch vetoes a kana-only auto-merge. Kana collide (伊藤/井藤, 斎藤/斉藤); these must be suggestions.
13. **異体字 folding is two-edged.** 斎/斉 are *different* legal surnames and are kept apart (suggestion "似た字"). itaiji maps 庄→荘, 椙→杉, 涌→湧, 苑→園 and 壬→閏, which merge different people, so they are blocked in `build-itaiji.mjs`. Expose 「異体字を同一視する」 (default ON) in the detail screen. Always show an info flag when strict names differ.
14. **NFKC does not fold kanji variants**, and does not remove IVS selectors (葛 + U+E0100 survives NFKC). Strip VS explicitly before NFKC.
15. **wanakana does not handle half-width kana.** Use NFKC. Romaji ⇄ kana needs the long-vowel fold (satou ≡ sato ≡ satoh).
16. **Duplicates in the other file** (resubmitted forms) become *ambiguous*, not first-wins. The UI must offer a rule (先頭 / 最終 / 手動). Duplicates in the base with lookup semantics are expected: that is `mode: lookup`, chosen automatically when the base key repeats and the other key is unique.
17. **Formula key cells.** Keys must come from the **cached result** (SheetJS `v`, ExcelJS `cell.result`). Files written by libraries without calculation may lack cached values; the key is then blank. Warn with 「数式の計算結果が保存されていません。Excelで開いて保存するか Excel 本体モードを使用」.
18. **Summary rows** (合計, 小計, Total) and filtered or hidden rows in key columns. Exclude rows whose key cell is 合計, 計, 小計 or total from matching (keep them in place for export), and keep hidden rows (they are data).
19. **CSV encoding.** Shift_JIS (cp932, including IBM-extension 髙 and 﨑) must be detected with `encoding-japanese`. Read with `raw:true`, or SheetJS converts `00123` to 123 and `1-2` to a date.
20. **Time.** The first profiler took 3 s for 100,000 cells. Sample 2,000 cells for ratios; keys still use all rows. Run it in a Web Worker.
21. **Header dictionary regexes in JSON** must be compiled with flag `u`, because `\p{Script=Han}` is used elsewhere. Entries such as `名$` in negatives are regexes, not substrings. When editing TypeScript through scripts, check that `\d` and `\s` survived: two of my regexes were silently changed to `d{…}` and `s`.

---

## 11. Integration notes for the implementers

- Put the modules in `src/core/keys/{normalize,detect,match,api}.ts` and `src/core/keys/dict/*.json`. `table.ts` is only the prototype reader: build the same `Table` shape from the app's ExcelJS/SheetJS reader and keep the `display` (`.w`) and `dateFmt` arrays.
- **Web Worker API:** `analyze(sheets[], baseSheetId)` returns `AnalysisDTO` for every other sheet against the base. `join(plan)` returns the join report. Selecting keys in the detail screen means editing `plan` (tier list), then re-running `join`, which takes milliseconds.
- **Undo/redo:** key selection is plain state (a list of tiers plus options). Put it in the same undo stack as the grid operations (the reference app has `undoRedoAction` in `src/util/keys.ts`).
- **Options for 「詳細」** (defaults in brackets):

  | Option | Default |
  |---|---|
  | 異体字を同一視 | ON |
  | 先頭ゼロ無視 | ON |
  | 敬称を除去 | ON |
  | 旧姓の括弧を除去 | ON |
  | 大学名の略称を統一 | ON |
  | 重複行の扱い (先頭 / 最終 / 手動) | 手動 |
  | 結合方法 (基準に追加=左結合 / 全件=外部結合) | 左結合 |
  | 提案の自動採用 | never |

- **Merged output:** matched other rows add columns to the base row, using the cell values and styles from the other file (the export engines, R-other). Unmatched other rows are appended at the end in outer mode. Add a status column (「一致方法: ID / 氏名＋大学 / 提案採用」) only if the user asks for it, because it is optional and changes the sheet.
- **Rebuilding the itaiji table:** `npm i itaiji itaiji-normalize && node scripts/build-itaiji.mjs`. Edit `BLOCK`, `OVERRIDE` and `ADD` in that script.

## 12. Artifacts

All under `r6/`:

- `dict/key-synonyms.json` — header dictionary (types, rules, negatives, noteWords, nonKeyHeaders, entityWords, family)
- `dict/itaiji-name.json` — 1,415 variant→canonical pairs for names (generated)
- `src/normalize.ts`, `src/detect.ts`, `src/match.ts`, `src/api.ts`, `src/table.ts`, `src/report.ts`
- `scripts/build-itaiji.mjs`, `scripts/make-synthetic.ts`, `scripts/run-samples.ts`, `scripts/run-synthetic.ts`, `scripts/test-normalize.ts`, `scripts/dump-json.ts`, `scripts/test-libs.mjs`, `scripts/test-roman.mjs`, `scripts/explore/*` (library coverage and MJ mining)
- `samples/` (copies of the interview-scheduler inputs), `synthetic/*.xlsx`, `synthetic/truth.json`, `synthetic/sjis.csv`
- `results/samples.txt`, `results/synthetic.txt`, `results/normalize-tests.txt`, `results/libs.txt`, `results/roman-sjis.txt`, `results/*.json` (UI DTO examples), `results/bundle-test.js`
- `package.json` scripts: `npm test`, `npm run samples`, `npm run synthetic`, `npm run json`, `npm run itaiji`, `npm run typecheck` (tsc 7.0.2, clean)

Sources: [Auto-Suggest (SIGMOD 2020)](https://dl.acm.org/doi/10.1145/3318464.3389738) · [Auto-FuzzyJoin (arXiv 2103.04489)](https://arxiv.org/pdf/2103.04489) · [Auto-Join (VLDB 2017)](http://www.vldb.org/pvldb/vol10/p1034-he.pdf) · [Auto-BI (arXiv 2306.12515)](https://arxiv.org/pdf/2306.12515) · [Power Query fuzzy merge](https://learn.microsoft.com/en-us/power-query/merge-queries-fuzzy-match) · [How fuzzy matching works in Power Query](https://learn.microsoft.com/en-us/power-query/fuzzy-matching) · [OpenRefine clustering in depth](https://openrefine.org/docs/technical-reference/clustering-in-depth) · GitHub: mziyut/itaiji-node, camelmasa/itaiji, tomatomerde/itaiji-normalize, geolonia/normalize-japanese-addresses, WaniKani/WanaKana, kazuhikoarase/jaconv, polygonplanet/encoding.js, ka-weihe/fastest-levenshtein, krisk/Fuse.
