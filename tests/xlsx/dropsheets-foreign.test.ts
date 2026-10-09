/**
 * Single-sheet output (PremergeOptions.dropOtherSheets) on packages that are spelled the way other writers spell them
 * (tests/fixtures/foreign/mutate.ts producer mutations + the generic respellings of tests/golden/package-tools.ts) and on books that
 * were read from legacy formats (xls / ods / xlsb: the "package" is generated from the values). The result must be the same as for
 * the original spelling and a clean one-sheet package.
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { readBook } from '../../src/core/xlsx';
import type { SourceBook } from '../../src/core/types';
import { MUTATIONS, mutate } from '../fixtures/foreign/mutate';
import { GOLDEN_FIX } from '../golden/premerge-scenario';
import { edit, prefixMain, text, utf16le, type Parts } from '../golden/package-tools';
import { baseSource, files, lintSingleSheet, lintText, premergeOne, readFixture } from './dropsheets-helpers';

const here = path.dirname(fileURLToPath(import.meta.url));
const XLSX_FIX = path.resolve(here, '../fixtures/xlsx');
const DS_FIX = path.resolve(here, '../fixtures/dropsheets');

/** what has to be the same whatever the spelling: cells, names, rules, links, tables, the report */
function digest(book: SourceBook, report: { droppedSheets?: string[]; formulasDetached?: number; items: Array<{ code: string; message: string }> }) {
  const s = book.sheets[0];
  // quotes a sheet name does not need, and the merged sheet's own name in front of its own references, are spelling only
  const unq = (t: string) => t.replace(/'([^' !]+)'!/g, '$1!').split(`${s.name}!`).join('');
  const cells: string[] = [];
  s.rows.forEach((row, r) => row?.forEach((c, i) => { if (c) cells.push(`${r + 1},${i + 1}|${unq(c.f ?? '')}|${c.t}|${String(c.v)}|${c.arr ?? ''}|${c.dyn ? 'd' : ''}`); }));
  return {
    sheet: s.name,
    cells,
    names: book.names.map((n) => `${n.scope ?? ''}!${n.name}=${unq(n.refersTo)}`).sort(),
    rules: s.features.rules.map((r) => `${r.kind}|${r.ext ? 'x14' : ''}|${r.type}|${r.sqref}|${r.formulas.map(unq).join(',')}`).sort(),
    links: s.features.hyperlinks.map((h) => `${h.ref}|${h.target ?? ''}|${unq(h.location ?? '')}`).sort(),
    tables: s.features.tables.map((t) => `${t.name}|${t.ref}`),
    merges: s.features.merges.length,
    dropped: report.droppedSheets,
    detached: report.formulasDetached,
    items: report.items.map((i) => `${i.code}:${i.message}`),
  };
}

async function run(bytes: Uint8Array, fileName: string, sheet?: string) {
  const base = await baseSource({ bytes, fileName, sheet });
  const res = await premergeOne(base);
  return { res, book: await readBook(res.bytes, 'o.xlsx') };
}

const SPELLINGS: Array<[string, (p: Parts) => void]> = [
  ['x: prefix on every part', (p) => { for (const n of [...p.keys()]) if (/\.(xml|rels)$/i.test(n) && typeof p.get(n) === 'string') p.set(n, prefixMain(p.get(n) as string, 'x')); }],
  ['BOM on every xml part', (p) => { for (const n of [...p.keys()]) if (/\.(xml|rels)$/i.test(n) && typeof p.get(n) === 'string') p.set(n, '﻿' + (p.get(n) as string)); }],
  ['UTF-16 workbook and sheets', (p) => { for (const n of [...p.keys()]) if (/^xl\/(workbook|worksheets\/[^/]+)\.xml$/.test(n)) p.set(n, utf16le(p.get(n) as string)); }],
];

const subjects: Array<{ label: string; file: string; fileName: string; sheet: string }> = [
  { label: 'torture.xlsx', file: path.join(GOLDEN_FIX, 'torture.xlsx'), fileName: 'torture.xlsx', sheet: 'Data' },
  { label: 'ds-features.xlsx', file: path.join(DS_FIX, 'ds-features.xlsx'), fileName: 'ds-features.xlsx', sheet: '名簿' },
  { label: 'samples/01_学生名簿.xlsx', file: path.resolve(here, '../../samples/01_学生名簿.xlsx'), fileName: '01_学生名簿.xlsx', sheet: '名簿' },
].filter((s) => fs.existsSync(s.file));

describe('producer-shaped spellings give the same single-sheet output', () => {
  for (const sub of subjects) {
    describe(sub.label, () => {
      const bytes = readFixture(sub.file);
      let expected: ReturnType<typeof digest>;
      it('original', async () => {
        const { res, book } = await run(bytes, sub.fileName, sub.sheet);
        expect(lintText(await lintSingleSheet(res.bytes, (res.report.droppedSheets ?? [])))).toBe('');
        expected = digest(book, res.report);
        expect(expected.dropped!.length).toBeGreaterThan(0);
      });
      for (const m of MUTATIONS) {
        it(m.name, async () => {
          const { res, book } = await run(await mutate(bytes, m), sub.fileName, sub.sheet);
          expect(lintText(await lintSingleSheet(res.bytes, (res.report.droppedSheets ?? [])))).toBe('');
          expect(digest(book, res.report)).toEqual(expected);
        });
      }
      for (const [name, respell] of SPELLINGS) {
        it(name, async () => {
          const { res, book } = await run(await edit(bytes, respell), sub.fileName, sub.sheet);
          expect(lintText(await lintSingleSheet(res.bytes, (res.report.droppedSheets ?? [])))).toBe('');
          expect(digest(book, res.report)).toEqual(expected);
        });
      }
    });
  }
});

describe('legacy formats as the base (the package is generated from the values)', () => {
  const cases: Array<{ file: string; sheet: string; removed: string[] }> = [
    { file: 'torture.xls', sheet: 'Data', removed: ['集計 シート', 'Tbl', 'Pivot', 'Lists'] },
    { file: 'legacy.ods', sheet: 'Data', removed: ['隠し'] },
    { file: 'legacy.xlsb', sheet: 'Data', removed: ['隠し'] },
  ];
  for (const c of cases) {
    it(`${c.file}: ${c.sheet} stays, ${c.removed.join(' / ')} go`, async () => {
      const p = path.join(XLSX_FIX, c.file);
      if (!fs.existsSync(p)) return;
      const { res, book } = await run(readFixture(p), c.file, c.sheet);
      expect(res.report.droppedSheets).toEqual(c.removed);
      expect(book.sheets.map((s) => s.name)).toEqual([c.sheet]);
      expect(lintText(await lintSingleSheet(res.bytes, c.removed))).toBe('');
      expect(sheetCount(await files(res.bytes))).toBe(1);
    });
  }
});

const sheetCount = (f: Map<string, string>) => [...f.keys()].filter((k) => /^xl\/worksheets\/[^/]+\.xml$/.test(k)).length;

void text;
