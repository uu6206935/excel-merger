/** Per-type normalisers and the header dictionary (ported from research R6 scripts/test-normalize.ts) + options. */
import { describe, expect, it } from 'vitest';
import {
  normId, normName, normKana, normEmail, normPhone, normDate, normOrg, normGeneric, excelSerialToIso, isBlankLike, idKeyOf,
  withNormalizeOptions, resolveNormalizeOptions, DEFAULT_NORMALIZE_OPTIONS, isSummaryText, summaryLabelKind, cellText,
} from '../../src/core/keys/normalize';
import { headerNorm, bestHeaderType } from '../../src/core/keys/detect';

const idk = (v: unknown, d?: string) => idKeyOf(v as never, d);
const IVS = String.fromCodePoint(0xe0100);
const SURR_YOSHI = String.fromCodePoint(0x20bb7); // 𠮷
const H = (h: string) => {
  const b = bestHeaderType(headerNorm(h));
  return b.type ? `${b.type}:${b.score}` : 'none';
};

describe('IDs', () => {
  it('numeric vs zero-padded', () => expect([idk(123), idk('00123')]).toEqual(['123', '123']));
  it('full-width', () => expect(idk('ＴＥＳＴ００００１')).toBe('TEST00001'));
  it('hyphen variants', () => expect([idk('test-0001'), idk('TEST－0001'), idk('TEST0001')]).toEqual(['TEST0001', 'TEST0001', 'TEST0001']));
  it('number shown as 000115 (numFmt 000000)', () => expect(normId(115, '000115').exact).toBe('000115'));
  it('large integer without E-notation', () => expect(idk(8000000001)).toBe('8000000001'));
  it('lossy 8.00E+09', () => expect(normId('8.00E+09').lossyExponent).toBe(true));
  it('123.0', () => expect(idk('123.0')).toBe('123'));
  it('blank-like values', () => expect([idk('-'), idk('#N/A'), idk('なし')]).toEqual(['', '', '']));
  it('spaces', () => expect(idk(' A 001 ')).toBe('A001'));
  it('literal apostrophe', () => expect(idk("'115")).toBe('115'));
  it('booleans / dates are never IDs', () => expect([idk(true), idk(new Date(0))]).toEqual(['', '']));
  it('0 is a valid ID value, not blank', () => {
    expect(isBlankLike('0')).toBe(false);
    expect(idk('0')).toBe('0');
  });
});

describe('names', () => {
  it('spaces incl. U+3000', () => expect([normName('山田　太郎').folded, normName('山田 太郎').folded, normName('山田太郎').folded]).toEqual(['山田太郎', '山田太郎', '山田太郎']));
  it('髙 = 高', () => expect(normName('髙橋 一郎').folded).toBe(normName('高橋一郎').folded));
  it('渡邊/渡邉/渡辺', () => expect([normName('渡邊').folded, normName('渡邉').folded, normName('渡辺').folded]).toEqual(['渡辺', '渡辺', '渡辺']));
  it('山﨑/山崎/山嵜', () => expect([normName('山﨑').folded, normName('山崎').folded, normName('山嵜').folded]).toEqual(['山崎', '山崎', '山崎']));
  it('齋→斎, 齊→斉 (斎 ≠ 斉 kept)', () => expect([normName('齋藤').folded, normName('齊藤').folded]).toEqual(['斎藤', '斉藤']));
  it('濵/濱', () => expect([normName('濵田').folded, normName('濱田').folded]).toEqual(['浜田', '浜田']));
  it('many old forms', () => expect(normName('澤廣國德嶋眞櫻瀧惠冨龍').folded).toBe('沢広国徳島真桜滝恵富竜'));
  it('surrogate 𠮷田', () => expect(normName(SURR_YOSHI + '田').folded).toBe('吉田'));
  it('IVS 葛+VS17', () => expect(normName('葛' + IVS + '城').folded).toBe('葛城'));
  it('佐々木 = 佐佐木', () => expect(normName('佐々木').folded).toBe(normName('佐佐木').folded));
  it('honorific', () => expect(normName('鈴木 一郎 様').folded).toBe('鈴木一郎'));
  it('maiden-name note', () => expect([normName('山田 花子（旧姓：佐藤）').folded, normName('山田 花子（旧姓：佐藤）').note]).toEqual(['山田花子', '(旧姓:佐藤)']));
  it('Latin order-insensitive', () => expect(normName('Smith John').latinSorted).toBe(normName('JOHN  SMITH').latinSorted));
  it('ジョン・スミス', () => expect(normName('ジョン・スミス').folded).toBe('ジョンスミス'));
  it('numbers are not names', () => expect(normName(123 as never).folded).toBe(''));
});

describe('kana', () => {
  it('half-width / hiragana / U+3000', () => expect([normKana('ﾔﾏﾀﾞ ﾀﾛｳ').strict, normKana('やまだ たろう').strict, normKana('ヤマダ　タロウ').strict]).toEqual(['ヤマダタロウ', 'ヤマダタロウ', 'ヤマダタロウ']));
  it('long vowel variants', () => expect([normKana('ユ-コ').strict, normKana('ユ－コ').strict, normKana('ユーコ').strict]).toEqual(['ユーコ', 'ユーコ', 'ユーコ']));
  it('loose forms', () => expect([normKana('キョウコ').loose === normKana('キヨウコ').loose, normKana('ハナヂ').loose === normKana('ハナジ').loose]).toEqual([true, true]));
});

describe('email / phone', () => {
  it('email full-width & upper', () => expect(normEmail('ＴＡＲＯ＠Example.CO.JP ')).toBe('taro@example.co.jp'));
  it('email invalid', () => expect(normEmail('taro at example')).toBe(''));
  it('phone +81', () => expect(normPhone('+81 90-1234-5678').digits).toBe('09012345678'));
  it('phone number cell lost 0', () => expect(normPhone(9012345678)).toEqual({ digits: '09012345678', restoredZero: true }));
  it('phone with extension', () => expect(normPhone('03(1234)5678 内線12').digits).toBe('0312345678'));
  it('phone full-width', () => expect(normPhone('０９０－１２３４－５６７８').digits).toBe('09012345678'));
});

describe('dates', () => {
  it('serial 36982', () => expect(excelSerialToIso(36982)).toBe('2001-04-01'));
  it('serials around the fake 1900-02-29', () => expect([excelSerialToIso(59), excelSerialToIso(61)]).toEqual(['1900-02-28', '1900-03-01']));
  it('和暦', () => expect([normDate('H13.4.1').iso, normDate('平成13年4月1日').iso, normDate('h13/04/01').iso]).toEqual(['2001-04-01', '2001-04-01', '2001-04-01']));
  it('令和元年', () => expect([normDate('令和元年5月1日').iso, normDate('R1.5.1').iso]).toEqual(['2019-05-01', '2019-05-01']));
  it('S63', () => expect(normDate('S63.12.31').iso).toBe('1988-12-31'));
  it('western forms', () => expect([normDate('2001/4/1').iso, normDate('2001-04-01').iso, normDate('2001年4月1日').iso, normDate('20010401').iso]).toEqual(['2001-04-01', '2001-04-01', '2001-04-01', '2001-04-01']));
  it('invalid 2/30', () => expect(normDate('2001/2/30').iso).toBe(''));
  it('Date object', () => expect(normDate(new Date(Date.UTC(2001, 3, 1))).iso).toBe('2001-04-01'));
  it('serial only with serialOk', () => expect([normDate(36982).iso, normDate(36982, true).iso]).toEqual(['', '2001-04-01']));
});

describe('organisations / generic', () => {
  it('慶應 variants', () => expect([normOrg('慶應義塾大学'), normOrg('慶応義塾大学'), normOrg('慶應')]).toEqual(['慶応義塾大学', '慶応義塾大学', '慶応義塾大学']));
  it('東大 variants', () => expect([normOrg('国立大学法人 東京大学'), normOrg('東京大学大学院'), normOrg('東大')]).toEqual(['東京大学', '東京大学', '東京大学']));
  it('generic 00123 = 123', () => expect([normGeneric('00123'), normGeneric(123)]).toEqual(['123', '123']));
});

describe('header dictionary', () => {
  it('name / kana', () => expect([H('氏名'), H('氏名（カナ）'), H('氏名カナ'), H('フリガナ')]).toEqual(['name:1', 'nameKana:1', 'nameKana:1', 'nameKana:1']));
  it('non-person names', () => expect([H('担当者名'), H('会社名'), H('大学名'), H('面接官氏名')]).toEqual(['none', 'none', 'university:1', 'none']));
  it('numbers', () => expect([H('番号'), H('No.'), H('電話番号'), H('郵便番号')]).toEqual(['rowNumber:0.9', 'rowNumber:0.9', 'phone:1', 'none']));
  it('ID family', () => expect([H('学籍番号'), H('学生番号'), H('学生ID'), H('受験番号'), H('応募者 ID'), H('社員番号')]).toEqual(['studentId:1', 'studentId:1', 'studentId:1', 'id:1', 'id:1', 'id:0.95']));
  it('email / phone', () => expect([H('E-mail'), H('メールアドレス（PC）'), H('携帯番号（ハイフンなし）')]).toEqual(['email:1', 'email:0.85', 'phone:1']));
  it('split names', () => expect([H('姓'), H('名'), H('セイ'), H('メイ'), H('苗字')]).toEqual(['lastName:1', 'firstName:1', 'lastNameKana:1', 'firstNameKana:1', 'lastName:1']));
  it('composite parts', () => expect([H('生年月日'), H('誕生日'), H('大学'), H('学部・学科'), H('性別')]).toEqual(['birthDate:1', 'birthDate:1', 'university:1', 'faculty:1', 'gender:1']));
  it('form question', () => expect(H('Q1. お名前（フルネーム）を入力してください')).toBe('name:0.75'));
  it('org codes', () => expect([H('職場ID'), H('大学コード')]).toEqual(['orgCode:0.9', 'orgCode:0.9']));
});

describe('misc', () => {
  it('summary labels', () => {
    expect(['合計', '小計', '総計', '計', 'Total', '合 計', '合計：'].map(isSummaryText)).toEqual([true, true, true, true, true, true, true]);
    expect(['計画', '会計', '山田 計', 123].map((v) => isSummaryText(v as never))).toEqual([false, false, false, false]);
  });
  it('summary label kinds: 合計 family strong, statistic labels weak, names / other words none', () => {
    const k = (v: unknown) => summaryLabelKind(v as never);
    expect(['合計', '【小計】', 'Grand Total', 'SUM', '合計（税込）'].map(k)).toEqual(['strong', 'strong', 'strong', 'strong', 'strong']);
    expect(['人数', '平均', '平均点', '件数', '最大', '最小値', '中央値', 'Average', 'COUNT', '■平均', '人数：'].map(k)).toEqual(Array(11).fill('weak'));
    expect(['平均太郎', '人数調整', '最大手', 'Amount', 'counter', '山田 花子', '志望度4以上', '', null, 60].map(k)).toEqual(Array(10).fill(null));
  });
  it('cellText of big integers / booleans', () => {
    expect(cellText(8000000001)).toBe('8000000001');
    expect(cellText(true)).toBe('TRUE');
    expect(cellText(null)).toBe('');
  });
});

describe('NormalizeOptions', () => {
  it('defaults are all on', () => {
    expect(resolveNormalizeOptions()).toEqual({ itaiji: true, leadingZeros: true, honorifics: true, maidenName: true, universityAbbrev: true });
    expect(resolveNormalizeOptions({ itaiji: false })).toEqual({ ...DEFAULT_NORMALIZE_OPTIONS, itaiji: false });
  });
  it('itaiji off keeps 髙 ≠ 高', () => {
    const o = resolveNormalizeOptions({ itaiji: false });
    withNormalizeOptions(o, () => {
      expect(normName('髙橋').folded).not.toBe(normName('高橋').folded);
      expect(normOrg('慶應義塾大学')).not.toBe(normOrg('慶応義塾大学'));
    });
    expect(normName('髙橋').folded).toBe(normName('高橋').folded); // restored afterwards
  });
  it('leadingZeros off keeps 00123 ≠ 123', () => {
    withNormalizeOptions(resolveNormalizeOptions({ leadingZeros: false }), () => {
      expect(idk('00123')).toBe('00123');
      expect(idk(123)).toBe('123');
      expect(normGeneric('00123')).toBe('00123');
    });
  });
  it('honorifics off keeps 様', () => {
    withNormalizeOptions(resolveNormalizeOptions({ honorifics: false }), () => expect(normName('鈴木 一郎 様').folded).toBe('鈴木一郎様'));
  });
  it('maidenName off keeps the note', () => {
    withNormalizeOptions(resolveNormalizeOptions({ maidenName: false }), () => expect(normName('山田 花子（旧姓：佐藤）').folded).not.toBe('山田花子'));
  });
  it('universityAbbrev off keeps 東大 ≠ 東京大学', () => {
    withNormalizeOptions(resolveNormalizeOptions({ universityAbbrev: false }), () => expect(normOrg('東大')).not.toBe(normOrg('東京大学')));
  });
  it('options restored after a throw', () => {
    expect(() => withNormalizeOptions(resolveNormalizeOptions({ leadingZeros: false }), () => {
      throw new Error('x');
    })).toThrow('x');
    expect(idk('00123')).toBe('123');
  });
});
