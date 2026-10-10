// tools/lib/sheetnames.mjs is a plain-JavaScript COPY of the contract file src/core/xlsx/sheetnames.ts (the QA tools run without a TypeScript loader).
// This test keeps the two identical: the same names for the same inputs, on hand-picked and on random inputs.
import { describe, expect, it } from 'vitest';
import { planSheetNames as portPlan, joinSheetName as portJoin, fileBaseName as portBase } from '../../tools/lib/sheetnames.mjs';
import { planSheetNames, joinSheetName, fileBaseName } from '../../src/core/xlsx/sheetnames';

const FILES = ['01_学生名簿.xlsx', 'C:\\Users\\x\\01_学生名簿.xlsx', '営業部門の月次実績報告書（2026年度第3四半期確定版）.xlsx', '売上[確定]:v2*?.xlsx', "田中's 出欠.csv", 'a.b.c.xls', '.xlsx', '', 'ＡＢＣ.xlsx', 'abc.xlsm', '😀😀😀😀😀😀😀😀😀😀😀😀😀😀😀😀.xlsx'];
const SHEETS = ['名簿', 'Sheet1', 'データ', "it's", 'a/b?c', '[x]', '', "'", '集計表の長い名前その１２３４５６７８９', 'ＡＢＣ', 'abc', 'ABC', '😀'.repeat(20), 'x'.repeat(40)];

function rnd(seed: number): () => number {
  let s = seed >>> 0 || 1;
  return () => ((s = (Math.imul(s, 1664525) + 1013904223) >>> 0) / 2 ** 32);
}

describe('tools/lib/sheetnames.mjs = src/core/xlsx/sheetnames.ts', () => {
  it('fileBaseName / joinSheetName agree on every pair of the pools', () => {
    for (const f of FILES) {
      expect(portBase(f), f).toBe(fileBaseName(f));
      for (const s of SHEETS) expect(portJoin(portBase(f), s), `${f} / ${s}`).toBe(joinSheetName(fileBaseName(f), s));
    }
  });

  it('planSheetNames agrees on 400 random sets of files (collisions, shortening, suffixes, existing names)', () => {
    const r = rnd(20261010);
    const pick = <T,>(a: T[]): T => a[Math.floor(r() * a.length)];
    for (let i = 0; i < 400; i++) {
      const existing = Array.from({ length: Math.floor(r() * 4) }, () => pick(SHEETS.filter(Boolean)));
      const requests = Array.from({ length: 1 + Math.floor(r() * 4) }, () => ({ file: pick(FILES), sheets: Array.from({ length: 1 + Math.floor(r() * 4) }, () => pick(SHEETS)) }));
      const want = planSheetNames(existing, requests);
      const got = portPlan(existing, requests);
      expect(got, JSON.stringify({ existing, requests })).toEqual(want);
    }
  });
});
