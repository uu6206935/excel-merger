/**
 * The formula-import corpus (tests/fixtures/formulaimport) through the whole pipeline WITH includeSourceSheets (round 9, task F2).
 * expected-copies.json holds the counters of every file under four configurations (tests/oracle/import-values/make-expected.ts --copies
 * rewrites it); expected.json is the same without the copies. The values are verified in real Excel by tests/oracle/import-values
 * (run.ps1 -Sources and -Random).
 */
import fs from 'node:fs';
import path from 'node:path';
import { describe, expect, it } from 'vitest';
import { tokenize } from '../../src/core/formula/excelRef';
import { readBook } from '../../src/core/xlsx';
import { CONFIGS } from '../oracle/import-values/make-expected';
import { loadManifest, runCorpusFile, sheetOf, type Counts, type CorpusEntry } from '../oracle/import-values/lib';
import { snapshot } from '../golden/snapshot';

const DIR = path.resolve(__dirname, '../fixtures/formulaimport');
const round8 = JSON.parse(fs.readFileSync(path.join(DIR, 'expected.json'), 'utf8')) as Record<string, Counts>;
const expected = JSON.parse(fs.readFileSync(path.join(DIR, 'expected-copies.json'), 'utf8')) as Record<string, Counts>;
const manifest: CorpusEntry[] = loadManifest(DIR);
const OPT = { includeSourceSheets: true };
/** what the copies cannot help: positions, other workbooks, 3-D references, spills, dynamic names, odd syntax */
const REMAINING = new Set(['positional', 'unsafe', 'array', 'unsupported', 'unresolvedName', 'otherSheet', 'unresolvedTable']);

describe('copies corpus: counters per file and configuration', () => {
  for (const e of manifest) {
    for (const cfg of CONFIGS) {
      it(`${e.file} | ${cfg.name}`, async () => {
        const r = await runCorpusFile(DIR, e.file, cfg, e.base ?? 'base.xlsx', OPT);
        expect(r.counts).toEqual(expected[`${e.file}|${cfg.name}`]);
        const before = round8[`${e.file}|${cfg.name}`];
        // nothing that round 8 kept is lost, and what is not kept any more is only what the copies cannot help
        expect(r.counts.kept).toBeGreaterThanOrEqual(before.kept);
        expect(r.counts.value + r.counts.empty).toBeLessThanOrEqual(before.value + before.empty);
        for (const why of Object.keys(r.counts.reasons)) expect(REMAINING.has(why), `${e.file}: ${why}`).toBe(true);
        // the report agrees with the trace
        const st = r.result.report.formulaStats ?? [];
        const sum = (k: 'kept' | 'toValues' | 'empty') => st.reduce((a, x) => a + x[k], 0);
        expect([sum('kept'), sum('toValues'), sum('empty')]).toEqual([r.counts.kept, r.counts.value, r.counts.empty]);
        expect(st.reduce((a, x) => a + (x.reasons.viaCopy ?? 0), 0)).toBe(r.counts.viaCopy ?? 0);
        for (const why of Object.keys(r.counts.reasons)) expect(st.reduce((a, x) => a + (x.reasons[why] ?? 0), 0)).toBe(r.counts.reasons[why]);
        // 'viaCopy' is informational: the reasons of the formulas that were NOT kept add up to toValues + empty without it
        for (const x of st) expect(Object.entries(x.reasons).filter(([k]) => k !== 'viaCopy').reduce((a, [, n]) => a + n, 0)).toBe(x.toValues + x.empty);
        expect(r.result.report.items.some((i) => i.code === 'formulasViaCopy')).toBe((r.counts.viaCopy ?? 0) > 0);
        expect(r.result.report.formulasKept).toBe(r.counts.kept);
        expect(r.result.report.formulasToValues).toBe(r.counts.value + r.counts.empty);
        for (const it of r.result.report.items) expect(it.message).toMatch(/[ぁ-んァ-ン一-龥]/);
      });
    }
  }
});

describe('copies corpus: what round 8 kept is kept exactly as it was', () => {
  it.each(manifest.filter((e) => e.order === 'shuf').map((e) => [e.file, e.base ?? 'base.xlsx']))('%s: the same formula for every formula round 8 kept, never read through a copy', async (file, baseFile) => {
    const cfg = { results: 'cached', imports: 'default', join: 'left' } as const;
    const a = await runCorpusFile(DIR, file, cfg, baseFile);
    const b = await runCorpusFile(DIR, file, cfg, baseFile, OPT);
    const key = (t: { srcRow: number; srcCol: number }) => `${t.srcRow}:${t.srcCol}`;
    const mine = new Map(b.trace.map((t) => [key(t), t]));
    let n = 0;
    for (const t of a.trace.filter((x) => x.outcome === 'kept')) {
      const u = mine.get(key(t))!;
      expect(u.outcome).toBe('kept');
      expect(u.formula).toBe(t.formula);
      expect(u.viaCopy ?? false).toBe(false);
      expect(u.cached).toBe(t.cached);
      n++;
    }
    expect(n).toBe(a.counts.kept);
  });
});

describe('copies corpus: the premerged packages are well-formed', () => {
  it.each(manifest.filter((e) => e.order === 'shuf').map((e) => [e.file, e.base ?? 'base.xlsx']))('%s: every kept formula is a formula of the merged sheet, tokenizes cleanly and only reads sheets that are in the output', async (file, baseFile) => {
    for (const join of ['left', 'full'] as const) {
      const r = await runCorpusFile(DIR, file, { results: 'nocache', imports: 'default', join }, baseFile, OPT);
      const book = await readBook(r.result.bytes, 'm.xlsx', { id: 'm' });
      const sheet = sheetOf(book, r.result.layout.sheet)!;
      const names = new Set(book.sheets.map((s) => s.name.toLowerCase()));
      // the sheets of the source are in the output (planSheetNames: "<file>_<sheet>")
      const added = r.result.report.addedSheets ?? [];
      expect(added.length).toBeGreaterThan(0);
      for (const n of added) expect(names.has(n.toLowerCase()), n).toBe(true);
      let viaCopy = 0;
      for (const t of r.trace.filter((x) => x.outcome === 'kept')) {
        const c = sheet.rows[t.row - 1]?.[t.col - 1];
        expect(c?.f, `${file} ${t.source}`).toBeTruthy();
        expect(c!.f).toBe(t.formula);
        const toks = tokenize(c!.f!);
        expect(toks.some((x) => x.type === 'unknown'), t.formula).toBe(false);
        expect(c!.f).not.toMatch(/#REF!/);
        let reads = false;
        for (const tk of toks) {
          if (tk.type !== 'ref' && tk.type !== 'name' && tk.type !== 'sref') continue;
          if (!tk.prefix) continue;
          reads = true;
          expect(tk.prefix.sheet2, t.formula).toBeUndefined();
          expect(tk.prefix.book, t.formula).toBeUndefined();
          expect(names.has((tk.prefix.sheet ?? '').toLowerCase()), `${t.formula} reads ${tk.prefix.sheet}`).toBe(true);
          // the merged sheet itself is never read through a prefix
          expect((tk.prefix.sheet ?? '').toLowerCase()).not.toBe(r.result.layout.sheet.toLowerCase());
        }
        expect(reads, t.formula).toBe(!!t.viaCopy);
        if (reads) viaCopy++;
      }
      expect(viaCopy).toBe(r.counts.viaCopy ?? 0);
      expect((await snapshot(r.result.bytes)).fullCalcOnLoad).toBe(true);
    }
  });
});

describe('copies corpus: what the findings say', () => {
  const cfg = { results: 'cached', imports: 'default', join: 'left' } as const;
  it('the sample workbook: the running totals that could not follow the merged order read the copy', async () => {
    const r = await runCorpusFile(DIR, 'g01_samples__shuf.xlsx', cfg, 'base_samples.xlsx', OPT);
    expect(r.counts).toMatchObject({ kept: 570, value: 0, empty: 0, viaCopy: 58 });
    expect(r.trace.filter((t) => t.viaCopy).every((t) => /^g01_samples__shuf_評価!/.test(t.formula!.replace(/^.*?(?=g01_samples)/, '')))).toBe(true);
    // the values are the source's own: the cached result stays next to the formula
    expect(r.trace.filter((t) => t.viaCopy).every((t) => t.cached && t.exact)).toBe(true);
  });

  it('a lookup into the master sheet of the source, an index/match over it and a rate on another sheet all stay formulas', async () => {
    const r = await runCorpusFile(DIR, 'f07_othersheet__shuf.xlsx', cfg, 'base.xlsx', OPT);
    expect(r.counts).toMatchObject({ kept: 162, value: 0, empty: 0, viaCopy: 108 });
    const f = r.trace.filter((t) => t.viaCopy).map((t) => t.formula!);
    expect(f.some((x) => /VLOOKUP\(A\d+,f07_othersheet__shuf_Master!\$A\$2:\$B\$100,2,FALSE\)/.test(x))).toBe(true);
    expect(f.some((x) => /INDEX\(f07_othersheet__shuf_Master!B:B,MATCH\(A\d+,f07_othersheet__shuf_Master!A:A,0\)\)/.test(x))).toBe(true);
    expect(f.some((x) => /f07_othersheet__shuf_Sheet2!\$B\$3\*E\d+/.test(x))).toBe(true);
  });

  it('a rate in a side cell and a footer total', async () => {
    const f12 = await runCorpusFile(DIR, 'f12_footer__shuf.xlsx', cfg, 'base.xlsx', OPT);
    expect(f12.counts).toMatchObject({ kept: 108, value: 0, viaCopy: 54 });
    expect(f12.trace.some((t) => /^E\d+\/f12_footer__shuf_評価!\$C\$34$/.test(t.formula!))).toBe(true);
    const f06 = await runCorpusFile(DIR, 'f06_running__shuf.xlsx', cfg, 'base.xlsx', OPT);
    expect(f06.trace.some((t) => /^E\d+\*f06_running__shuf_評価!\$P\$1$/.test(t.formula!))).toBe(true);
    // ROW()-1 is host-dependent: it stays a value
    expect(f06.counts.reasons).toEqual({ positional: 27 });
  });

  it('defined names: a name of a side cell or of another sheet is inlined as a reference to the copy; a dynamic (OFFSET) name stays a value', async () => {
    const r = await runCorpusFile(DIR, 'f08_names__shuf.xlsx', cfg, 'base.xlsx', OPT);
    expect(r.counts.reasons).toEqual({ unresolvedName: 27 });
    const f = r.trace.filter((t) => t.viaCopy).map((t) => t.formula!);
    expect(f.some((x) => /^E\d+\*f08_names__shuf_評価!\$P\$1$/.test(x))).toBe(true);
    expect(f.some((x) => /VLOOKUP\(A\d+,f08_names__shuf_Master!\$A\$2:\$B\$100,2,FALSE\)/.test(x))).toBe(true);
  });

  it('tables: a column that is not imported and a whole table read the copy', async () => {
    const r = await runCorpusFile(DIR, 'f09_tables__shuf.xlsx', cfg, 'base.xlsx', OPT);
    expect(r.counts).toMatchObject({ value: 0, empty: 0 });
    expect(r.trace.some((t) => /^ROWS\(f09_tables__shuf_評価!\$A\$2:\$N\$32\)$/.test(t.formula!))).toBe(true);
  });

  it('positional / unsafe / 3-D formulas stay values with the same reasons as before', async () => {
    const r = await runCorpusFile(DIR, 'f16_unsafe__shuf.xlsx', cfg, 'base.xlsx', OPT);
    expect(r.counts.reasons).toEqual({ otherSheet: 27, positional: 54, unsafe: 27 });
    // OFFSET(ref, …) reads the copy; INDIRECT / ROW() stay values
    expect(r.trace.some((t) => /^OFFSET\(f16_unsafe__shuf_評価!C\d+,0,1\)$/.test(t.formula!))).toBe(true);
  });

  it('a file without sheets of its own besides the data sheet: formulas that need a copy read the copy of the data sheet', async () => {
    const r = await runCorpusFile(DIR, 'f03_keycol__shuf.xlsx', cfg, 'base.xlsx', OPT);
    expect(r.counts).toMatchObject({ kept: 189, value: 0, viaCopy: 27 });
    expect(r.trace.find((t) => t.viaCopy)!.formula).toMatch(/^VLOOKUP\(A\d+,f03_keycol__shuf_評価!\$A\$2:\$E\$32,3,FALSE\)$/);
  });
});
