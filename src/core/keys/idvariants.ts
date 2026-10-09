/**
 * ID key variants between two files whose ID columns do not match literally but are obviously the same numbering:
 *   'prefix'  S2023-0012 ⟷ 20230012 (a letter prefix on one or both sides; the prefix must not carry identity: B/M/D20230001
 *             are three people, so stripping it is refused when it makes keys collide);
 *   'suffix'  20230012 ⟷ 230012 (the shorter number is the tail of the longer one for most rows).
 * Both return per-row keys usable by pairStats / executeJoin (equal strings = same person) and a recipe that rebuilds the
 * keys on another profile of the same column (the base extended by appended rows).
 */
import type { ColumnProfile } from './detect';
import { getNormalizeOptions } from './normalize';

const PREFIX_RE = /^([A-Z]{1,4})[.#:]?(\d+)$/;
const DIGITS_RE = /^\d+$/;

const zeros = (d: string): string => (getNormalizeOptions().leadingZeros ? d.replace(/^0+(?=\d)/, '') : d);

export interface Stripped {
  keys: string[];
  nonBlank: number;
  prefixed: number;
}

/** keys of an ID column with a leading letter prefix removed ('S20230012' → '20230012'); other keys unchanged */
export function stripPrefixKeys(keys: readonly string[]): Stripped {
  const out = new Array<string>(keys.length);
  let nonBlank = 0;
  let prefixed = 0;
  for (let i = 0; i < keys.length; i++) {
    const k = keys[i];
    if (!k) {
      out[i] = '';
      continue;
    }
    nonBlank++;
    const m = PREFIX_RE.exec(k);
    if (m) {
      prefixed++;
      out[i] = zeros(m[2]);
    } else out[i] = k;
  }
  return { keys: out, nonBlank, prefixed };
}

function distinct(keys: readonly string[]): number {
  const s = new Set<string>();
  for (const k of keys) if (k) s.add(k);
  return s.size;
}

export interface KeyVariant {
  variant: 'prefix' | 'suffix';
  label: string;
  a: { keys: string[]; keysFor?: (p: ColumnProfile) => string[] };
  b: { keys: string[]; keysFor?: (p: ColumnProfile) => string[] };
}

/** 'prefix' variant of a column pair, or null (no prefix anywhere / the prefix carries identity) */
export function prefixVariant(ka: readonly string[], kb: readonly string[]): KeyVariant | null {
  const sa = stripPrefixKeys(ka);
  const sb = stripPrefixKeys(kb);
  const rate = (s: Stripped) => (s.nonBlank ? s.prefixed / s.nonBlank : 0);
  // the letters are a formatting difference only when ONE side carries them: S2023-0012 ⟷ T2023-0012 (two lettered systems
  // that happen to share digits) are two numberings, not one
  const oneSided = (rate(sa) >= 0.5 && rate(sb) < 0.1) || (rate(sb) >= 0.5 && rate(sa) < 0.1);
  if (!oneSided) return null;
  // the prefix may only be dropped when that creates no new collisions inside a column
  for (const [orig, s] of [[ka, sa], [kb, sb]] as const) {
    if (s.prefixed && distinct(s.keys) < distinct(orig) * 0.99) return null;
  }
  const f = (p: ColumnProfile) => stripPrefixKeys(p.keys).keys;
  return { variant: 'prefix', label: 'ID（英字の接頭辞を除く）', a: { keys: sa.keys, keysFor: f }, b: { keys: sb.keys, keysFor: f } };
}

const MIN_TAIL = 4;
const MAX_LENGTHS = 4;
const MAX_EXTRA = 8;

function digitShare(keys: readonly string[]): { n: number; share: number; lengths: Map<number, number> } {
  let n = 0;
  let d = 0;
  const lengths = new Map<number, number>();
  for (const k of keys) {
    if (!k) continue;
    n++;
    if (DIGITS_RE.test(k)) {
      d++;
      lengths.set(k.length, (lengths.get(k.length) ?? 0) + 1);
    }
  }
  return { n, share: n ? d / n : 0, lengths };
}

function medianLength(lengths: Map<number, number>): number {
  const all = [...lengths.entries()].sort((x, y) => x[0] - y[0]);
  const total = all.reduce((s, e) => s + e[1], 0);
  let acc = 0;
  for (const [len, c] of all) {
    acc += c;
    if (acc >= total / 2) return len;
  }
  return 0;
}

/**
 * 'suffix' variant: all-digit IDs of different lengths where the shorter ID is the tail of the longer one (year without the
 * century, serial without the faculty code …). The longer side's key is replaced by its tail when that tail is one of the
 * shorter side's IDs; every other key stays as it is (and so never matches).
 */
export function suffixVariant(ka: readonly string[], kb: readonly string[]): KeyVariant | null {
  const da = digitShare(ka);
  const db = digitShare(kb);
  if (da.n < 3 || db.n < 3 || da.share < 0.9 || db.share < 0.9) return null;
  const ma = medianLength(da.lengths);
  const mb = medianLength(db.lengths);
  if (ma === mb || !ma || !mb) return null;
  const aIsShort = ma < mb;
  const shortKeys = aIsShort ? ka : kb;
  const longKeys = aIsShort ? kb : ka;
  const sets = new Map<number, Set<string>>();
  for (const k of shortKeys) {
    if (k && DIGITS_RE.test(k) && k.length >= MIN_TAIL) {
      let s = sets.get(k.length);
      if (!s) sets.set(k.length, (s = new Set()));
      s.add(k);
    }
  }
  if (!sets.size || sets.size > MAX_LENGTHS) return null;
  const lens = [...sets.keys()].sort((x, y) => y - x);
  const remap = (keys: readonly string[]): { out: string[]; hits: number; tails: number } => {
    const out = new Array<string>(keys.length);
    const tails = new Set<string>();
    let hits = 0;
    for (let i = 0; i < keys.length; i++) {
      const k = keys[i];
      out[i] = k;
      if (!k || !DIGITS_RE.test(k)) continue;
      for (const l of lens) {
        if (l >= k.length || k.length - l > MAX_EXTRA) continue;
        const t = k.slice(k.length - l);
        if (sets.get(l)!.has(t)) {
          out[i] = t;
          hits++;
          tails.add(t);
          break;
        }
      }
    }
    return { out, hits, tails: tails.size };
  };
  const r = remap(longKeys);
  const longNonBlank = aIsShort ? db.n : da.n;
  // most of the longer column must end in a known short ID, and (nearly) every tail must come from ONE long ID
  if (r.hits < longNonBlank * 0.5 || r.tails < r.hits * 0.95) return null;
  const fLong = (p: ColumnProfile) => remap(p.keys).out;
  const long = { keys: r.out, keysFor: fLong };
  const short = { keys: [...shortKeys] };
  return { variant: 'suffix', label: 'ID（桁数違い・末尾一致）', a: aIsShort ? short : long, b: aIsShort ? long : short };
}
