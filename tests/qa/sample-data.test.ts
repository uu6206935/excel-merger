// Invariants of the deterministic sample data (tools/lib/sample-data.mjs) — what each sample is meant to show.
import { describe, expect, it } from 'vitest';
import {
  buildSpec,
  buildTruth,
  foldName,
  fromSerial,
  serial,
  toCsv,
  toFullDigits,
  toHalfKana,
  INTERVIEW,
  REF_DATE,
  ROSTER,
} from '../../tools/lib/sample-data.mjs';

const spec = buildSpec();
const students = spec.students;
const truth = buildTruth(spec);

describe('sample data helpers', () => {
  it('serial dates match Excel (1900 system)', () => {
    expect(serial(1900, 3, 1)).toBe(61);
    expect(serial(2026, 4, 1)).toBe(46113);
    expect(fromSerial(46113)).toEqual({ y: 2026, m: 4, d: 1 });
  });
  it('converts to half-width kana and full-width digits', () => {
    expect(toHalfKana('サトウ　ショウタ')).toBe('ｻﾄｳ ｼｮｳﾀ');
    expect(toHalfKana('ガッコウ ヴィ')).toBe('ｶﾞｯｺｳ ｳﾞｨ');
    expect(toHalfKana('パ')).toBe('ﾊﾟ');
    expect(toFullDigits('000123')).toBe('０００１２３');
  });
  it('folds 異体字 and spaces for the uniqueness check', () => {
    expect(foldName('髙橋　優斗')).toBe(foldName('高橋優斗'));
    expect(foldName('齋藤 美咲')).toBe('斎藤美咲');
  });
  it('writes CSV with CRLF and quoting', () => {
    expect(toCsv(['a', 'b'], [['x,y', 'q"q'], [null, 1]])).toBe('a,b\r\n"x,y","q""q"\r\n,1\r\n');
  });
  it('is deterministic', () => {
    const again = buildSpec();
    expect(JSON.stringify(again)).toBe(JSON.stringify(spec));
  });
});

describe('01_学生名簿 (base)', () => {
  it('has 60 students with unique 6-digit text IDs (leading zeros)', () => {
    expect(students).toHaveLength(60);
    const ids = students.map((s) => s.id);
    expect(new Set(ids).size).toBe(60);
    for (const id of ids) expect(id).toMatch(/^0\d{5}$/);
    // column A of the roster rows is the raw ID (the builder formats the column as text first)
    expect(spec.base.rows[0][0]).toBe(ids[0]);
  });
  it('has exactly one 同姓同名 pair, at different universities', () => {
    const byName = new Map<string, typeof students>();
    for (const s of students) {
      const k = foldName(s.name);
      byName.set(k, [...(byName.get(k) ?? []), s]);
    }
    const dups = [...byName.values()].filter((v) => v.length > 1);
    expect(dups).toHaveLength(1);
    expect(dups[0].map((s) => s.uni).sort()).toEqual(['慶應義塾大学', '早稲田大学']);
    expect(truth.base.sameNamePair).toEqual(dups[0].map((s) => ROSTER.firstDataRow + s.index));
  });
  it('contains 髙橋 and 高橋, 齋藤, 渡邉, 山﨑 and a hiragana given name', () => {
    const names = students.map((s) => s.name);
    expect(names.some((n) => n.startsWith('髙橋'))).toBe(true);
    expect(names.some((n) => n.startsWith('高橋'))).toBe(true);
    expect(names.some((n) => n.startsWith('齋藤'))).toBe(true);
    expect(names.some((n) => n.startsWith('渡邉'))).toBe(true);
    expect(names.some((n) => n.startsWith('山﨑'))).toBe(true);
    expect(names).toContain('小野　さくら');
  });
  it('ages agree with DATEDIF(生年月日, 基準日, "Y")', () => {
    const ref = Date.UTC(REF_DATE.y, REF_DATE.m - 1, REF_DATE.d);
    for (const s of students) {
      const b = fromSerial(s.birth);
      let age = REF_DATE.y - b.y;
      if (Date.UTC(REF_DATE.y, b.m - 1, b.d) > ref) age--;
      expect(age).toBe(s.age);
    }
  });
  it('writes formulas for 年齢 and the hidden helper column', () => {
    const r = spec.base.rows[0];
    expect(r[9]).toBe('=DATEDIF(I4,基準日,"Y")');
    expect(r[11]).toBe('=COUNTIF($B$4:$B$63,B4)');
    expect(ROSTER.headers).toHaveLength(12);
  });
  it('集計 uses whole-column, bounded, name, SUMPRODUCT, VLOOKUP and INDEX/MATCH formulas', () => {
    const s = spec.base.summary;
    const all = [...s.uniRows.flat(), ...s.total, ...s.gradeRows.flat(), ...s.others.flat()].filter((x): x is string => typeof x === 'string' && x.startsWith('='));
    const text = all.join('\n');
    for (const needle of ['名簿!$D:$D', '名簿!$D$4:$D$63', 'COUNTIFS(', 'AVERAGEIF(', 'SUMPRODUCT(', '志望度列', 'VLOOKUP(', 'INDEX(', 'COUNTBLANK(']) {
      expect(text).toContain(needle);
    }
  });
});

describe('02_面接評価', () => {
  const meta = spec.interview.meta;
  it('misses 2 base students, has 3 unknown IDs and 2 resubmissions', () => {
    const base = new Set(meta.filter((m) => m.baseIndex != null).map((m) => m.baseIndex));
    expect(base.size).toBe(58);
    for (const i of INTERVIEW.missing) expect(base.has(i)).toBe(false);
    expect(meta.filter((m) => m.extra)).toHaveLength(3);
    const resub = meta.filter((m) => m.resubmitted).map((m) => m.baseIndex);
    expect(resub.sort()).toEqual([...INTERVIEW.resubmitted].sort());
    expect(spec.interview.rows).toHaveLength(58 + 2 + 3);
    const baseIds = new Set(students.map((s) => s.id));
    for (const x of INTERVIEW.extras) expect(baseIds.has(x.id)).toBe(false);
  });
  it('stores IDs as zero-padded text, plain numbers and full-width text', () => {
    const kinds = new Set(meta.map((m) => m.idKind));
    expect([...kinds].sort()).toEqual(['fullwidth', 'number', 'text']);
    const numRow = spec.interview.rows[meta.findIndex((m) => m.idKind === 'number')];
    expect(typeof numRow[0]).toBe('number');
    const fwRow = spec.interview.rows[meta.findIndex((m) => m.idKind === 'fullwidth')];
    expect(String(fwRow[0])).toMatch(/^'[０-９]{6}$/);
  });
  it('uses 異体字 and varied spacing in 氏名', () => {
    const names = meta.map((m) => m.name);
    expect(names.some((n) => n.startsWith('高橋') && n.includes('優斗'))).toBe(true); // base 髙橋
    expect(names.some((n) => n.startsWith('髙橋') && n.includes('健太'))).toBe(true); // base 高橋
    expect(names.some((n) => n.startsWith('斎藤'))).toBe(true);
    expect(names.some((n) => /^\S+ \S+$/.test(n))).toBe(true);
    expect(names.some((n) => /^[^\s　]+$/.test(n))).toBe(true);
    expect(names.some((n) => n.includes('　'))).toBe(true);
  });
  it('has 平均 and 判定 formulas on every row', () => {
    spec.interview.rows.forEach((r, k) => {
      expect(r[6]).toBe(`=AVERAGE(D${k + 2}:F${k + 2})`);
      expect(r[7]).toBe(`=IF(G${k + 2}>=3.5,"合格","保留")`);
    });
  });
});

describe('03 / 04 / 05', () => {
  it('03_連絡先 has no ID column and both 同姓同名 students', () => {
    expect(spec.contacts.headers).toEqual(['氏名', 'メール', '電話']);
    const pair = truth.base.sameNamePair;
    const mapped = truth.contacts.rows.map((r) => r.baseRow);
    for (const p of pair) expect(mapped).toContain(p);
    expect(truth.contacts.expected.extraRows).toHaveLength(2);
  });
  it('04_適性検査 has kana only (full-width, half-width, no space) and unrelated 学籍番号', () => {
    const kinds = new Set(spec.aptitude.meta.map((m) => m.kanaKind));
    expect([...kinds].sort()).toEqual(['fullwidth', 'halfwidth', 'nospace']);
    const ids = new Set(students.map((s) => Number(s.id)));
    for (const r of spec.aptitude.rows) {
      expect(typeof r[0]).toBe('number');
      expect(ids.has(r[0] as number)).toBe(false);
      expect(String(r[0])).toMatch(/^\d{8}$/);
    }
  });
  it('05_研修出欠 needs 氏名＋大学 (both 佐藤 翔太 present with different universities)', () => {
    const rows = spec.attendance.rows;
    const sato = rows.filter((r) => r[0] === '佐藤 翔太');
    expect(sato).toHaveLength(2);
    expect(new Set(sato.map((r) => r[1])).size).toBe(2);
    expect(spec.attendance.headers.slice(0, 2)).toEqual(['氏名', '大学']);
    // some universities are written with abbreviations (表記ゆれ)
    expect(rows.some((r) => ['東大', '早大', '慶應大学', '京大', '阪大', '一橋', '明大', '上智', '東北大', '名大'].includes(String(r[1])))).toBe(true);
  });
  it('truth.json expectations are consistent', () => {
    expect(truth.interview.expected.matchedBaseRows + truth.interview.expected.unmatchedBaseRows.length).toBe(60);
    expect(truth.attendance.expected.matchedBaseRows + truth.attendance.expected.unmatchedBaseRows.length).toBe(60);
    expect(truth.base.lastDataRow).toBe(63);
    expect(truth.base.summaryRefs.length).toBeGreaterThan(40);
  });
});
