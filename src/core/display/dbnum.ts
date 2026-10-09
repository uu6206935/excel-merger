/**
 * [DBNum1] / [DBNum2] / [DBNum3] numerals of a number format (数字の表記), as the ja-JP Excel displays them (measured with
 * Range.Text, see tests/grid/numfmt-excel.test.ts):
 *   [DBNum1]  漢数字   一二三四五 (digits) / 百二十三 (General) — 〇 for zero
 *   [DBNum2]  大字     壱弐参四伍 (digits) / 壱百弐拾参 (General): 壱 is always written, units 拾 百 阡 萬
 *   [DBNum3]  全角     １２３４５ (digits) / １万２千３百４十５ (General): Arabic full-width digits with the kanji units
 * Number formats with digit placeholders (0, #,##0, 0.0 …) substitute digit by digit; "General" writes the value with the units.
 * Pure: no DOM, never throws.
 */

export type DbStyle = 1 | 2 | 3;

const DIGITS: Record<DbStyle, string> = { 1: '〇一二三四五六七八九', 2: '〇壱弐参四伍六七八九', 3: '０１２３４５６７８９' };
const SMALL_UNITS: Record<DbStyle, [string, string, string]> = { 1: ['十', '百', '千'], 2: ['拾', '百', '阡'], 3: ['十', '百', '千'] };
const BIG_UNITS: Record<DbStyle, string[]> = { 1: ['', '万', '億', '兆'], 2: ['', '萬', '億', '兆'], 3: ['', '万', '億', '兆'] };

/** Every ASCII digit of the text replaced (digit by digit): 12,345 -> 一二,三四五. */
export function digitWise(text: string, style: DbStyle): string {
  const d = DIGITS[style];
  return text.replace(/[0-9]/g, (c) => d[Number(c)]);
}

/** 0 .. 9999 inside one group of four digits. */
function group(g: number, style: DbStyle): string {
  const d = DIGITS[style];
  const u = SMALL_UNITS[style];
  const th = Math.floor(g / 1000) % 10;
  const h = Math.floor(g / 100) % 10;
  const t = Math.floor(g / 10) % 10;
  const o = g % 10;
  let s = '';
  const add = (val: number, unit: string) => {
    if (val === 0) return;
    // DBNum1 / DBNum3 leave the 1 out before 十 百 千 (百二十三); DBNum2 always writes 壱 (壱百弐拾参)
    s += (val === 1 && style !== 2 ? '' : d[val]) + unit;
  };
  add(th, u[2]);
  add(h, u[1]);
  add(t, u[0]);
  if (o > 0) s += d[o];
  return s;
}

/** A whole number with the positional units: 12345 -> 一万二千三百四十五, 100 -> 百, 10000 -> 一万. */
export function positional(n: bigint | number, style: DbStyle): string {
  let v = typeof n === 'bigint' ? n : BigInt(Math.trunc(n));
  if (v < 0n) v = -v;
  if (v === 0n) return DIGITS[style][0];
  const big = BIG_UNITS[style];
  let out = '';
  for (let i = 0; v > 0n && i < big.length; i++) {
    const g = Number(v % 10000n);
    v /= 10000n;
    if (g > 0) out = group(g, style) + big[i] + out;
  }
  return out;
}

/**
 * The text of a General number ('12.5', '-25', '1234567') written with the numerals; scientific notation and anything that is
 * not a plain decimal is returned as it is.
 */
export function generalText(text: string, style: DbStyle): string {
  const m = /^(-?)(\d+)(?:\.(\d+))?$/.exec(text.trim());
  if (!m) return text;
  const frac = m[3] ? '.' + digitWise(m[3], style) : '';
  return m[1] + positional(BigInt(m[2]), style) + frac;
}

/**
 * A two-letter date token (mm dd hh ss yy ee): below 10 the 〇 is written (〇五), from 10 on the units are used (十, 二十五);
 * ja-JP Excel: [DBNum1]mm → 〇一 (Jan) / 十 (Oct) / 十二 (Dec), dd 25 → 二十五, [DBNum2]mm/dd → 壱拾/〇壱.
 */
export function pad2(n: number, style: DbStyle): string {
  const v = Math.trunc(n);
  return v < 10 ? DIGITS[style][0] + DIGITS[style][Math.max(0, v)] : positional(v, style);
}
