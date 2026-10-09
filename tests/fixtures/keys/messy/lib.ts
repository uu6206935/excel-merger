/**
 * Shared building blocks of the "messy real data" fixtures (Round 5): a deterministic cast of 40 people, a tiny
 * sheet builder that records WHICH PERSON every row is (the ground truth), and the xlsx writer used by
 * `npx tsx tests/fixtures/keys/messy/write-all.ts` (outputs are committed under ./out/).
 *
 * Nothing here depends on src/ (the fixtures must stay valid when the code under test changes).
 */
import * as XLSX from 'xlsx';
import type { CellSpec, SheetSpec } from '../../../keys/helpers';

export type { CellSpec, SheetSpec };

// ---------------------------------------------------------------------------
// people
// ---------------------------------------------------------------------------

export interface Person {
  /** ground-truth identity ('p0' … 'p39') */
  id: string;
  sei: string;
  mei: string;
  seiKana: string;
  meiKana: string;
  seiRoman: string;
  meiRoman: string;
  univ: string;
  fac: string;
  email: string;
  tel: string;
  /** 8-digit student number 2023xxxx */
  sid: number;
  birth: Date;
}

const SEI: Array<[string, string, string]> = [
  ['山田', 'ヤマダ', 'Yamada'], ['鈴木', 'スズキ', 'Suzuki'], ['佐藤', 'サトウ', 'Sato'], ['高橋', 'タカハシ', 'Takahashi'], ['渡辺', 'ワタナベ', 'Watanabe'],
  ['伊藤', 'イトウ', 'Ito'], ['中村', 'ナカムラ', 'Nakamura'], ['小林', 'コバヤシ', 'Kobayashi'], ['加藤', 'カトウ', 'Kato'], ['吉田', 'ヨシダ', 'Yoshida'],
  ['山口', 'ヤマグチ', 'Yamaguchi'], ['松本', 'マツモト', 'Matsumoto'], ['井上', 'イノウエ', 'Inoue'], ['木村', 'キムラ', 'Kimura'], ['林', 'ハヤシ', 'Hayashi'],
  ['斎藤', 'サイトウ', 'Saito'], ['清水', 'シミズ', 'Shimizu'], ['山崎', 'ヤマザキ', 'Yamazaki'], ['森', 'モリ', 'Mori'], ['池田', 'イケダ', 'Ikeda'],
];
const MEI: Array<[string, string, string]> = [
  ['太郎', 'タロウ', 'Taro'], ['花子', 'ハナコ', 'Hanako'], ['一郎', 'イチロウ', 'Ichiro'], ['健太', 'ケンタ', 'Kenta'], ['美咲', 'ミサキ', 'Misaki'],
  ['翔', 'ショウ', 'Sho'], ['結衣', 'ユイ', 'Yui'], ['大輝', 'ダイキ', 'Daiki'], ['陽菜', 'ヒナ', 'Hina'], ['蓮', 'レン', 'Ren'],
  ['優子', 'ユウコ', 'Yuko'], ['拓也', 'タクヤ', 'Takuya'], ['さくら', 'サクラ', 'Sakura'], ['悠真', 'ユウマ', 'Yuma'], ['葵', 'アオイ', 'Aoi'],
  ['直樹', 'ナオキ', 'Naoki'], ['彩', 'アヤ', 'Aya'], ['誠', 'マコト', 'Makoto'], ['瞳', 'ヒトミ', 'Hitomi'], ['隆', 'タカシ', 'Takashi'],
];
const UNIV = ['東京大学', '慶應義塾大学', '早稲田大学', '京都大学', '大阪大学', '東北大学', '名古屋大学', '九州大学', '北海道大学', '一橋大学', '上智大学', '明治大学', '立教大学', '青山学院大学'];
const FAC = ['経済学部', '法学部', '工学部', '文学部', '理学部', '総合政策学部', '商学部', '農学部'];

export function person(i: number): Person {
  const s = SEI[i % 20];
  const m = MEI[(i * 3 + Math.floor(i / 20)) % 20];
  return {
    id: `p${i}`,
    sei: s[0], mei: m[0], seiKana: s[1], meiKana: m[1], seiRoman: s[2], meiRoman: m[2],
    univ: UNIV[(i * 5 + 1) % UNIV.length],
    fac: FAC[(i * 3) % FAC.length],
    email: `${m[2].toLowerCase()}.${s[2].toLowerCase()}@example.ac.jp`,
    tel: `090-${String(1000 + ((i * 37) % 9000))}-${String((i * 7919) % 10000).padStart(4, '0')}`,
    sid: 20230000 + i + 1,
    birth: new Date(Date.UTC(2001 + (i % 4), i % 12, 1 + ((i * 3) % 28))),
  };
}

export const PEOPLE: Person[] = Array.from({ length: 40 }, (_, i) => person(i));

/** deterministic permutation (LCG), so every regenerated fixture is byte-stable */
export function shuffle<T>(xs: readonly T[], seed: number): T[] {
  const a = [...xs];
  let s = seed >>> 0 || 1;
  for (let i = a.length - 1; i > 0; i--) {
    s = (Math.imul(s, 1664525) + 1013904223) >>> 0;
    const j = s % (i + 1);
    [a[i], a[j]] = [a[j], a[i]];
  }
  return a;
}

/** people i0 … i1-1 */
export const range = (i0: number, i1: number): Person[] => PEOPLE.slice(i0, i1);

export const fullName = (p: Person, sep = '　') => `${p.sei}${sep}${p.mei}`;
export const fullKana = (p: Person, sep = '　') => `${p.seiKana}${sep}${p.meiKana}`;
export const roman = (p: Person) => `${p.meiRoman} ${p.seiRoman}`;

/** Excel 1900 serial of a UTC date */
export const serial = (d: Date) => Math.round(d.getTime() / 86400000 + 25569);

const pad = (n: number, w = 2) => String(n).padStart(w, '0');
export function warekiOf(d: Date, style: 'H13.4.1' | '平成13年4月1日' | 'h13/04/01'): string {
  const y = d.getUTCFullYear(), m = d.getUTCMonth() + 1, day = d.getUTCDate();
  const ymd = y * 10000 + m * 100 + day;
  let era: [string, string, number];
  if (ymd >= 20190501) era = ['令和', 'R', 2018];
  else if (ymd >= 19890108) era = ['平成', 'H', 1988];
  else era = ['昭和', 'S', 1925];
  const ey = y - era[2];
  switch (style) {
    case 'H13.4.1': return `${era[1]}${ey}.${m}.${day}`;
    case '平成13年4月1日': return `${era[0]}${ey === 1 ? '元' : ey}年${m}月${day}日`;
    default: return `${era[1].toLowerCase()}${pad(ey)}/${pad(m)}/${pad(day)}`;
  }
}

// ---------------------------------------------------------------------------
// files with ground truth
// ---------------------------------------------------------------------------

export interface MessyFile {
  /** source id in the analysis ('base', 's1', 's2' …) */
  id: string;
  fileName: string;
  sheets: SheetSpec[];
  /** name of the data sheet (default: the first one) */
  dataSheet?: string;
  /** Excel row (1-based) of the data sheet → person identity. Rows that are not people are absent. */
  truth: Record<number, string>;
  /** hidden columns (1-based) of the data sheet — written to the xlsx, not used by the tests */
  hiddenCols?: number[];
}

export interface CaseExpect {
  /** correct automatic matches per other source id (exact) */
  matched?: Record<string, number>;
  /** base rows whose person is in the other file, were NOT joined automatically (on purpose) and must be surfaced (suggestion / ambiguous) */
  surfaced?: Record<string, number>;
  /** base rows whose person is in the other file but is neither joined nor surfaced (default 0: a file without a key …) */
  unsurfaced?: Record<string, number>;
  /** candidate labels the auto plan selects, in tier order, per other source */
  plan?: Record<string, string[]>;
  /** check-item codes that must exist / must not exist for a source (`id:code`) */
  items?: string[];
  noItems?: string[];
  /** appended rows of the 'full' join */
  appended?: number;
  /** free text: the mess this case contains */
  note?: string;
}

export interface MessyCase {
  name: string;
  note: string;
  files: MessyFile[];
  plan?: { joinType?: 'left' | 'full'; duplicatePolicy?: 'first' | 'last' | 'skip' };
  expect: CaseExpect;
}

/** Sheet builder: rows are pushed in file order; add() records the ground truth of a person row. */
export class Sheet {
  rows: CellSpec[][] = [];
  truth: Record<number, string> = {};
  merges: Array<{ r1: number; c1: number; r2: number; c2: number }> = [];
  hiddenCols: number[] = [];
  constructor(public name = 'Sheet1') {}
  /** a row that is not a person (title, header, blank, subtotal …) */
  raw(cells: CellSpec[] = []): this {
    this.rows.push(cells);
    return this;
  }
  /** a person row; `who` = ground-truth identity (Person or explicit id) */
  add(who: Person | string, cells: CellSpec[]): this {
    this.rows.push(cells);
    this.truth[this.rows.length] = typeof who === 'string' ? who : who.id;
    return this;
  }
  merge(r1: number, c1: number, r2: number, c2: number): this {
    this.merges.push({ r1, c1, r2, c2 });
    return this;
  }
  hide(...cols: number[]): this {
    this.hiddenCols.push(...cols);
    return this;
  }
  spec(): SheetSpec {
    return { name: this.name, rows: this.rows, merges: this.merges.length ? this.merges : undefined };
  }
  file(id: string, fileName: string, extraSheets: SheetSpec[] = []): MessyFile {
    const f: MessyFile = { id, fileName, sheets: [this.spec(), ...extraSheets], truth: { ...this.truth } };
    if (this.hiddenCols.length) f.hiddenCols = [...this.hiddenCols];
    return f;
  }
}

/** Builds a sheet: header row(s) first, then persons through the callback. */
export function build(name: string, header: CellSpec[], fill: (s: Sheet) => void, pre: CellSpec[][] = []): Sheet {
  const s = new Sheet(name);
  for (const r of pre) s.raw(r);
  s.raw(header);
  fill(s);
  return s;
}

// ---------------------------------------------------------------------------
// xlsx writer (SheetJS; deterministic content)
// ---------------------------------------------------------------------------

function toCell(spec: CellSpec): XLSX.CellObject | null {
  if (spec === null || spec === undefined) return null;
  const o = typeof spec === 'object' && !(spec instanceof Date) ? spec : { v: spec };
  const v = o.v;
  let cell: XLSX.CellObject;
  if (v instanceof Date) cell = { t: 'n', v: serial(v), z: o.fmt ?? 'yyyy/m/d' };
  else if (v === null) return null;
  else if (typeof v === 'number') cell = { t: 'n', v, ...(o.fmt ? { z: o.fmt } : {}) };
  else if (typeof v === 'boolean') cell = { t: 'b', v };
  else cell = { t: 's', v };
  if (o.f) cell.f = o.f;
  return cell;
}

export function toWorkbook(f: MessyFile): XLSX.WorkBook {
  const wb = XLSX.utils.book_new();
  for (const sp of f.sheets) {
    const ws: XLSX.WorkSheet = {};
    let maxC = 0;
    sp.rows.forEach((row, r) => {
      row.forEach((spec, c) => {
        const cell = toCell(spec);
        if (!cell) return;
        ws[XLSX.utils.encode_cell({ r, c })] = cell;
        maxC = Math.max(maxC, c);
      });
    });
    ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: Math.max(0, sp.rows.length - 1), c: maxC } });
    if (sp.merges?.length) ws['!merges'] = sp.merges.map((m) => ({ s: { r: m.r1 - 1, c: m.c1 - 1 }, e: { r: m.r2 - 1, c: m.c2 - 1 } }));
    if (f.hiddenCols?.length && sp === f.sheets[0]) {
      const cols: XLSX.ColInfo[] = [];
      for (const c of f.hiddenCols) cols[c - 1] = { hidden: true };
      for (let i = 0; i < cols.length; i++) cols[i] ??= {};
      ws['!cols'] = cols;
    }
    XLSX.utils.book_append_sheet(wb, ws, sp.name);
  }
  if (f.sheets.some((sp) => sp.state && sp.state !== 'visible')) {
    wb.Workbook = { Sheets: f.sheets.map((sp) => ({ Hidden: sp.state === 'hidden' ? 1 : sp.state === 'veryHidden' ? 2 : 0 })) };
  }
  return wb;
}

export function toXlsxBytes(f: MessyFile): Uint8Array {
  const out = XLSX.write(toWorkbook(f), { type: 'array', bookType: 'xlsx', compression: true }) as ArrayBuffer;
  return new Uint8Array(out);
}
