/**
 * Generator of the randomised messy lists of messy-fuzz.test.ts (Round 5) — randomised lists against the ground truth: whatever columns two files share and however their values are
 * written (prefixes, zeros, 全角, 様, 異体字, typos, blanks, resubmissions, 同姓同名 …) the default plan must never join two
 * different people. (What is left unjoined is fine — it must show up as unmatched / suggestion / ambiguous.)
 *
 * Unavoidable "wrong" matches are kept out of the generator: a typo never produces another person's ID / name / address, and a
 * pair of namesakes is always in BOTH files or in neither (a namesake that exists in one file only cannot be told apart by anyone).
 */
import type { KeySource } from '../../src/core/contracts';
import { type Person, PEOPLE, fullKana, fullName, serial, warekiOf } from '../fixtures/keys/messy/lib';
import { toHalfKana, toHira } from '../fixtures/keys/messy/make-names';
import { type CellSpec, sourceFromRows } from './helpers';

class Rng {
  private s: number;
  constructor(seed: number) {
    this.s = (seed * 2654435761) >>> 0 || 1;
  }
  next(): number {
    this.s = (Math.imul(this.s, 1664525) + 1013904223) >>> 0;
    return this.s / 4294967296;
  }
  int(n: number): number {
    return Math.floor(this.next() * n);
  }
  chance(p: number): boolean {
    return this.next() < p;
  }
  pick<T>(xs: readonly T[]): T {
    return xs[this.int(xs.length)];
  }
  shuffle<T>(xs: readonly T[]): T[] {
    const a = [...xs];
    for (let i = a.length - 1; i > 0; i--) {
      const j = this.int(i + 1);
      [a[i], a[j]] = [a[j], a[i]];
    }
    return a;
  }
}

/** a person of the fuzz universe: the fixture cast + namesake pairs (same 氏名, different everything else) */
interface Fp extends Person {
  nameKey: string;
}

function universe(rng: Rng): { people: Fp[]; pairs: Array<[Fp, Fp]> } {
  const people: Fp[] = PEOPLE.slice(0, 34).map((p) => ({ ...p, nameKey: fullName(p) }));
  const pairs: Array<[Fp, Fp]> = [];
  for (let k = 0; k < 2; k++) {
    const a = PEOPLE[34 + k * 2];
    const b = PEOPLE[35 + k * 2];
    const mk = (p: Person, q: Person, i: number): Fp => ({ ...p, sei: q.sei, mei: q.mei, seiKana: q.seiKana, meiKana: q.meiKana, seiRoman: q.seiRoman, meiRoman: q.meiRoman, id: `n${k}${i}`, nameKey: fullName(q),
      email: `ns${k}${i}@example.ac.jp`, sid: 20239000 + k * 10 + i, univ: i ? '早稲田大学' : '東京大学', tel: `080-${k}${i}00-1234`, birth: new Date(Date.UTC(2000 + k, 3 + i, 10 + i)) });
    pairs.push([mk(a, a, 0), mk(b, a, 1)]);
  }
  void rng;
  return { people, pairs };
}

/** RowMatch flags that put a row into the 要確認 lists (2nd-tier key, name / birth date conflict, duplicate rows) */
export const FLAGGED: ReadonlySet<string> = new Set(['fallback', 'conflict', 'duplicateInSource', 'duplicateInBase', 'manual']);

type Col = 'id' | 'name' | 'split' | 'kana' | 'email' | 'phone' | 'univ' | 'birth';
const ALL: Col[] = ['id', 'name', 'split', 'kana', 'email', 'phone', 'univ', 'birth'];

export interface Fuzz {
  base: KeySource;
  other: KeySource;
  baseTruth: Map<number, string>;
  otherTruth: Map<number, string>;
  joinType: 'left' | 'full';
  policy: 'first' | 'last' | 'skip';
  desc: string;
}

function idText(p: Person, rng: Rng, style: number): CellSpec {
  const s = String(p.sid);
  switch (style) {
    case 1: return `S${s.slice(0, 4)}-${s.slice(4)}`;
    case 2: return p.sid;
    case 3: return s.replace(/\d/g, (x) => String.fromCharCode(0xff10 + Number(x)));
    case 4: return ` ${s} `;
    case 5: return `S${s}`;
    default: return rng.chance(0.5) ? s : p.sid;
  }
}

/** a typo'd ID is nobody's number — and no two typos are the same number */
let typoCounter = 0;

function makeFile(people: Fp[], cols: Col[], rng: Rng, prefix: string, dupRate: number, style: { id: number; name: number; kana: number }): { rows: CellSpec[][]; truth: Map<number, string> } {
  const header: CellSpec[] = [];
  for (const c of cols) {
    if (c === 'id') header.push(prefix === 'b' ? '学籍番号' : '学生番号');
    else if (c === 'name') header.push('氏名');
    else if (c === 'split') header.push('姓', '名');
    else if (c === 'kana') header.push('フリガナ');
    else if (c === 'email') header.push('メールアドレス');
    else if (c === 'phone') header.push('電話番号');
    else if (c === 'univ') header.push('大学');
    else header.push('生年月日');
  }
  header.push('備考');
  const rows: CellSpec[][] = [header];
  const truth = new Map<number, string>();
  const addRow = (p: Fp, dup: boolean) => {
    const r: CellSpec[] = [];
    const namesake = p.id.startsWith('n');
    for (const c of cols) {
      switch (c) {
        case 'id': {
          if (rng.chance(0.1)) r.push(null);
          else if (!dup && rng.chance(0.06)) r.push(String(20249000 + typoCounter++));
          else r.push(idText(p, rng, style.id));
          break;
        }
        case 'name': {
          let n = fullName(p, style.name === 1 ? ' ' : style.name === 2 ? '' : '　');
          if (rng.chance(0.15)) n = `${n}${rng.pick(['様', 'さん', ' 様'])}`;
          if (rng.chance(0.06)) n = p.nameKey.replace(/[髙高]/, '髙');
          // (a typo replaces the LAST character of the given name: with a one-character given name two different people could
          // end up with the very same typo'd name, which nobody can tell apart)
          if (!dup && !namesake && p.mei.length >= 2 && rng.chance(0.05)) n = `${p.sei}${p.mei.slice(0, -1)}${rng.pick(['朗', '輝', '樹'])}`;
          // (a namesake's name is never blank: a blank there makes the two indistinguishable for EVERYONE)
          r.push(rng.chance(0.03) && !namesake ? null : n);
          break;
        }
        case 'split': r.push(p.sei, p.mei); break;
        case 'kana': r.push(rng.chance(0.03) && !namesake ? null : style.kana === 1 ? toHira(fullKana(p)) : style.kana === 2 ? toHalfKana(fullKana(p, ' ')) : fullKana(p)); break;
        case 'email': r.push(rng.chance(0.1) ? null : rng.chance(0.3) ? ` ${p.email.toUpperCase()}` : p.email); break;
        case 'phone': r.push(rng.chance(0.1) ? null : rng.pick([p.tel, p.tel.replace(/-/g, ''), `+81 ${p.tel.slice(1)}`, Number(p.tel.replace(/-/g, ''))])); break;
        case 'univ': r.push(rng.chance(0.1) ? null : p.univ); break;
        default: r.push(rng.chance(0.1) ? null : rng.pick<CellSpec>([p.birth, warekiOf(p.birth, 'H13.4.1'), serial(p.birth), `${p.birth.getUTCFullYear()}/${p.birth.getUTCMonth() + 1}/${p.birth.getUTCDate()}`]));
      }
    }
    r.push(rng.chance(0.2) ? `メモ${rng.int(50)}` : null);
    rows.push(r);
    truth.set(rows.length, p.id);
  };
  for (const p of people) {
    addRow(p, false);
    if (rng.chance(dupRate)) addRow(p, true);
  }
  return { rows, truth };
}

export function fuzz(seed: number): Fuzz {
  const rng = new Rng(seed);
  const { people, pairs } = universe(rng);
  const pickPeople = (n: number, withPairs: boolean): Fp[] => {
    const xs = rng.shuffle(people).slice(0, n);
    return withPairs ? rng.shuffle([...xs, ...pairs.flat()]) : xs;
  };
  // a pair of namesakes is in both files or in neither
  const withPairs = rng.chance(0.5);
  const baseP = pickPeople(14 + rng.int(14), withPairs);
  const otherP = pickPeople(10 + rng.int(18), withPairs);
  const shared = (): Col[] => ALL.filter(() => rng.chance(0.5));
  let cb = shared();
  let co = shared();
  if (!cb.length) cb = ['name'];
  if (!co.length) co = ['kana'];
  const sb = { id: rng.int(3), name: rng.int(3), kana: rng.int(3) };
  const so = { id: rng.int(6), name: rng.int(3), kana: rng.int(3) };
  const b = makeFile(baseP, cb, rng, 'b', 0.05, sb);
  const o = makeFile(otherP, co, rng, 'o', 0.25, so);
  const base = sourceFromRows('base', b.rows);
  const other = sourceFromRows('o', o.rows);
  return {
    base, other, baseTruth: b.truth, otherTruth: o.truth,
    joinType: rng.chance(0.3) ? 'full' : 'left', policy: rng.pick(['first', 'last', 'skip'] as const),
    desc: `seed ${seed}: base [${cb.join(',')}] other [${co.join(',')}] pairs=${withPairs}`,
  };
}

export interface FuzzMulti {
  base: KeySource;
  others: KeySource[];
  baseTruth: Map<number, string>;
  /** per other source id */
  otherTruth: Record<string, Map<number, string>>;
  policy: 'first' | 'last' | 'skip';
  desc: string;
}

/** a base and 2–3 other files, each with its own columns / notations (sequential "full" joins match appended rows of earlier files) */
export function fuzzMulti(seed: number): FuzzMulti {
  const rng = new Rng(seed * 31 + 7);
  const { people, pairs } = universe(rng);
  const withPairs = rng.chance(0.4);
  const pick = (n: number): Fp[] => {
    const xs = rng.shuffle(people).slice(0, n);
    return withPairs ? rng.shuffle([...xs, ...pairs.flat()]) : xs;
  };
  const shared = (): Col[] => {
    const c = ALL.filter(() => rng.chance(0.5));
    return c.length ? c : ['name'];
  };
  const cb = shared();
  const b = makeFile(pick(12 + rng.int(10)), cb, rng, 'b', 0.05, { id: rng.int(3), name: rng.int(3), kana: rng.int(3) });
  const n = 2 + rng.int(2);
  const others: KeySource[] = [];
  const otherTruth: Record<string, Map<number, string>> = {};
  const descs: string[] = [];
  for (let k = 0; k < n; k++) {
    const co = shared();
    const o = makeFile(pick(10 + rng.int(14)), co, rng, 'o', 0.2, { id: rng.int(6), name: rng.int(3), kana: rng.int(3) });
    others.push(sourceFromRows(`o${k}`, o.rows));
    otherTruth[`o${k}`] = o.truth;
    descs.push(`o${k}[${co.join(',')}]`);
  }
  return { base: sourceFromRows('base', b.rows), others, baseTruth: b.truth, otherTruth, policy: rng.pick(['first', 'last', 'skip'] as const), desc: `seed ${seed}: base [${cb.join(',')}] ${descs.join(' ')} pairs=${withPairs}` };
}
