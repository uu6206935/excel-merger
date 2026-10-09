/**
 * The same workbook spelled the way OTHER writers spell it (tests/fixtures/foreign/mutate.ts: LibreOffice booleans, explicit t="n",
 * inline strings, pretty printed XML with comments, reversed attribute order, numeric character references, lower-case cell references,
 * no dimension / spans / views, "85.0" numbers, a stored zip with data descriptors and [Content_Types].xml last) must read, premerge and
 * replay to exactly the same workbook as the original spelling: the pipeline runs on both with ONE op log (taken from a random run on
 * the original) and every sheet of the two outputs is compared (values, formulas, styles, widths, merges, tables, rules, names).
 */
import fs from 'node:fs';
import path from 'node:path';
import { describe, expect, it } from 'vitest';
import { readBook, replayLog } from '../../src/core/xlsx';
import { MUTATIONS, mutate } from '../fixtures/foreign/mutate';
import { CORPUS_DIR, PARTNER, bookDiff, load, loadBytes, premergeFiles, runPipeline } from '../fixtures/foreign/harness';

/** files that show different features: tables, notes, charts, names, CF / DV, 1904 dates, two-row headers, shared formulas */
const FILES = [
  'hc-01-google-sheets.xlsx', 'hc-02-libreoffice.xlsx', 'hc-05-mac-excel-2011-1904.xlsx', 'hc-10-shared-formulas.xlsx', 'hc-28-table-awkward-names.xlsx',
  'hc-35-names-cf-dv.xlsx', 'hc-30-libreoffice-hidden-filter.xlsx', 'py-01-openpyxl-styled.xlsx', 'py-04-openpyxl-table.xlsx', 'py-11-openpyxl-chart-names-links.xlsx',
  'py-15-xlsxwriter-basic.xlsx', 'js-04-exceljs-basic.xlsx', 'hc-29-formula-variety.xlsx',
];

describe('the same workbook in other producers\' spelling gives the same result', () => {
  for (const f of FILES) {
    for (const role of ['base', 'other'] as const) {
      it(`${f} as ${role}: ${MUTATIONS.length} spellings`, async () => {
        const partner = await load(CORPUS_DIR, PARTNER, 'p');
        const original = await load(CORPUS_DIR, f, 'f');
        // the op log comes from a random run on the original; both packages replay it
        const run = await runPipeline(original, partner, { role, join: 'full', seed: 3, steps: 12 });
        expect(run.verdict, `${run.stage}: ${run.problems.join(' | ')}`).toBe('ok');
        const log = run.log!;
        const base0 = role === 'base' ? original : partner;
        const other0 = role === 'base' ? partner : original;
        const ref = await premergeFiles(base0, other0, 'full');
        const refOut = await readBook(await replayLog(ref.bytes, log), f);
        const failures: string[] = [];
        for (const m of MUTATIONS) {
          const mutated = await loadBytes(f, await mutate(original.bytes, m), 'f');
          // the reader must see the same sheet before any patching
          const readDiff = bookDiff(original.book, mutated.book);
          if (readDiff.length) failures.push(`[${m.name}] readBook differs: ${readDiff.slice(0, 3).join(' ; ')}`);
          if (JSON.stringify(mutated.ks.region) !== JSON.stringify(original.ks.region)) { failures.push(`[${m.name}] region ${JSON.stringify(mutated.ks.region)} vs ${JSON.stringify(original.ks.region)}`); continue; }
          try {
            const pm = await premergeFiles(role === 'base' ? mutated : partner, role === 'base' ? partner : mutated, 'full');
            const out = await readBook(await replayLog(pm.bytes, log), f);
            const d = bookDiff(refOut, out);
            if (d.length) failures.push(`[${m.name}] (${m.like}) output differs: ${d.slice(0, 3).join(' ; ')}`);
          } catch (e) {
            failures.push(`[${m.name}] (${m.like}) ${(e as Error).message.slice(0, 200)}`);
          }
        }
        expect(failures).toEqual([]);
      });
    }
  }

  it('keeps the corpus honest: the mutations really change the package', async () => {
    const bytes = new Uint8Array(fs.readFileSync(path.join(CORPUS_DIR, 'hc-02-libreoffice.xlsx')));
    for (const m of MUTATIONS) expect(Buffer.compare(Buffer.from(await mutate(bytes, m)), Buffer.from(bytes)), m.name).not.toBe(0);
  });
});
