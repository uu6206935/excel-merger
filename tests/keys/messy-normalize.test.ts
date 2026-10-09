/**
 * Round 5 — the normalisers on messy cell texts (the fixtures in tests/fixtures/keys/messy test them end to end).
 */
import { describe, expect, it } from 'vitest';
import { suffixVariant, prefixVariant, stripPrefixKeys } from '../../src/core/keys/idvariants';
import {
  EMAIL_RE, MAX_KEY_TEXT, isKanjiNameShape, isLatinNameShape, isNoteLineText, isPageMarkerText, kanaStrict, nfkc, normEmail, normId, normName, normPhone,
  summaryLabelKind, idKeyOf,
} from '../../src/core/keys/normalize';

describe('nfkc fast path', () => {
  it('is exactly NFKC + trim on every character of the fast-path ranges (printable ASCII, hiragana, katakana, kanji, 々)', () => {
    const ranges: Array<[number, number]> = [[0x20, 0x7e], [0x3041, 0x3096], [0x30a1, 0x30fa], [0x30fc, 0x30fc], [0x3005, 0x3005], [0x4e00, 0x9fff]];
    let n = 0;
    for (const [a, b] of ranges) {
      for (let cp = a; cp <= b; cp++) {
        const c = String.fromCharCode(cp);
        expect(nfkc(`${c}${c}`)).toBe(`${c}${c}`.normalize('NFKC').trim());
        n++;
      }
    }
    expect(n).toBeGreaterThan(21000);
    expect(nfkc('  山田　太郎  ')).toBe('山田 太郎');
    expect(nfkc('ｶﾞｷﾞｸﾞ')).toBe('ガギグ');
    expect(nfkc('ヷヸヹヺ')).toBe('ヷヸヹヺ');
  });
});

describe('e-mail addresses in free text', () => {
  const addr = 'taro.yamada@example.ac.jp';
  it.each([
    [`山田 太郎 <${addr}>`], [`"山田" <${addr}>`], [`(${addr})`], [`${addr}、`], [` ${addr}。`], [`mailto:${addr}`], [addr.replace('@', '＠')], [addr.toUpperCase()],
    [`${addr}\n`], [`${addr} / other@example.com`], [`連絡先: ${addr}`], [`taro.yamada(at)example.ac.jp`],
  ])('%s', (v) => {
    expect(normEmail(v)).toBe(addr);
  });
  it('not an address', () => {
    for (const v of ['山田 太郎', 'taro@', '@example.com', 'a@b', '', '   ', 'n/a', 'taro at example dot com']) expect(normEmail(v)).toBe('');
    expect(EMAIL_RE.test(addr)).toBe(true);
  });
});

describe('phone numbers', () => {
  it('two numbers in one cell: the first valid one', () => {
    expect(normPhone('090-1234-5678 / 03-1111-2222').digits).toBe('09012345678');
    expect(normPhone('090-1234-5678、080-0000-1111').digits).toBe('09012345678');
    expect(normPhone('03-1111-2222；090-1234-5678').digits).toBe('0311112222');
    expect(normPhone('内線123 / 090-1234-5678').digits).toBe('09012345678');
  });
  it('still the old forms', () => {
    expect(normPhone('+81 90-1234-5678').digits).toBe('09012345678');
    expect(normPhone(9012345678)).toEqual({ digits: '09012345678', restoredZero: true });
    expect(normPhone('(090) 1234-5678').digits).toBe('09012345678');
    expect(normPhone('090,1234,5678').digits).toBe('09012345678');
    expect(normPhone('12345').digits).toBe('');
  });
});

describe('names', () => {
  it('the reading printed under the name in the same cell is not part of the name', () => {
    const plain = normName('山田 太郎');
    expect(normName('山田 太郎\nヤマダ タロウ')).toEqual(plain);
    expect(normName('山田 太郎\r\nやまだ たろう')).toEqual(plain);
    expect(normName('ヤマダ タロウ\n山田 太郎')).toEqual(plain);
    // a name that is ALL kana, or two lines that are both names, stays as it is
    expect(normName('ヤマダ\nタロウ').strict).toBe('ヤマダタロウ');
    expect(normName('山田\n太郎').strict).toBe('山田太郎');
    expect(isKanjiNameShape('山田 太郎\nヤマダ タロウ')).toBe(true);
    expect(isLatinNameShape('Taro Yamada\nタロウ')).toBe(true);
  });
  it('Latin names: any order, comma, capitals, double space', () => {
    const k = (s: string) => normName(s).latinSorted;
    expect(new Set(['Taro Yamada', 'YAMADA Taro', 'Yamada, Taro', 'TARO  YAMADA', 'yamada,taro'].map(k)).size).toBe(1);
    expect(isLatinNameShape('Yamada, Taro')).toBe(true);
    expect(isLatinNameShape('Yamada,')).toBe(false);
  });
  it('an essay is never a name: long texts have no key (and no quadratic regex)', () => {
    const long = '（'.repeat(MAX_KEY_TEXT + 1);
    expect(normName(long).strict).toBe('');
    expect(kanaStrict('ア'.repeat(MAX_KEY_TEXT + 1))).toBe('');
    expect(normEmail(`${'a'.repeat(MAX_KEY_TEXT)}@example.com`)).toBe('');
    expect(normName('山'.repeat(MAX_KEY_TEXT)).strict.length).toBe(MAX_KEY_TEXT);
  });
});

describe('IDs', () => {
  it('a number that lost its digits in text (2.02301E+11) has no key; an exact one keeps it', () => {
    expect(idKeyOf('2.02301E+11')).toBe('');
    expect(idKeyOf('2.02E+07')).toBe('');
    expect(idKeyOf('2.0230012E+07')).toBe('20230012');
    expect(idKeyOf(202301000000)).toBe('202301000000');
    expect(normId('2.02301E+11').lossyExponent).toBe(true); // still reported (lossyId check item)
  });
  it('letter prefix variants', () => {
    expect(stripPrefixKeys(['S20230012', '20230013', 'NO123', '']).keys).toEqual(['20230012', '20230013', '123', '']);
    const v = prefixVariant(['S20230001', 'S20230002', 'S20230003'], ['20230001', '20230002', '20230003']);
    expect(v?.a.keys).toEqual(['20230001', '20230002', '20230003']);
    // the prefix carries identity: stripping would collide
    expect(prefixVariant(['B20230001', 'M20230001', 'D20230001', 'B20230002'], ['20230001', '20230002', '20230003', '20230004'])).toBeNull();
    // nothing prefixed at all
    expect(prefixVariant(['1', '2', '3'], ['1', '2', '3'])).toBeNull();
    // both sides lettered (S… ⟷ T…): two numberings that share digits, not one numbering written twice
    expect(prefixVariant(['S0001', 'S0002', 'S0003'], ['T0001', 'T0002', 'T0003'])).toBeNull();
    // either side may carry the prefix
    expect(prefixVariant(['0001', '0002', '0003'], ['S0001', 'S0002', 'S0003'])?.b.keys).toEqual(['1', '2', '3']);
  });
  it('tail variant: the shorter number is the tail of the longer one', () => {
    const v = suffixVariant(['20230001', '20230002', '20230003', '20230004'], ['230001', '230002', '230003', '230009']);
    expect(v?.variant).toBe('suffix');
    expect(v?.a.keys).toEqual(['230001', '230002', '230003', '20230004']); // the long side's unmatched key stays (never matches)
    expect(v?.b.keys).toEqual(['230001', '230002', '230003', '230009']);
    // same length / too short / unrelated numbers: no variant
    expect(suffixVariant(['20230001', '20230002', '20230003'], ['30230001', '30230002', '30230003'])).toBeNull();
    expect(suffixVariant(['123', '456', '789'], ['3', '6', '9'])).toBeNull();
    expect(suffixVariant(['20230001', '20230002', '20230003'], ['111111', '222222', '333333'])).toBeNull();
  });
});

describe('list furniture', () => {
  it('page-number lines', () => {
    for (const v of ['- 2 -', '― 2 ―', 'Page 2 of 5', 'page 2', '2/5', '(2/5)', '第2ページ', '2ページ', 'p.2', '-2-', '＜ 3 ＞']) expect(isPageMarkerText(v), v).toBe(true);
    for (const v of ['2', '山田 太郎', 'S2023-0012', '2026/5/1', '', 'Total 5']) expect(isPageMarkerText(v), v).toBe(false);
  });
  it('remark lines', () => {
    for (const v of ['※ 以下は2次募集分', '＊辞退者', '---- 辞退者 ----', '注）集計は4/1時点', '以下余白', '◆ 2次募集', '=====']) expect(isNoteLineText(v), v).toBe(true);
    for (const v of ['山田 太郎', 'ABC-001', '以上太郎', '-5', '']) expect(isNoteLineText(v), v).toBe(v === '以上太郎');
  });
  it('qualified subtotal labels need a row without other text (checked by the region scan)', () => {
    expect(summaryLabelKind('東京大学　小計')).toBe('suffix');
    expect(summaryLabelKind('◆ページ計')).toBe('suffix');
    expect(summaryLabelKind('部門別計')).toBe('suffix');
    expect(summaryLabelKind('Group Total')).toBe('suffix');
    expect(summaryLabelKind('小計')).toBe('strong');
    expect(summaryLabelKind('山田 太郎')).toBeNull();
    expect(summaryLabelKind('設計')).toBeNull();
  });
});
