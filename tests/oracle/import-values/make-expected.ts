/**
 * Writes tests/fixtures/formulaimport/expected.json: the counters of every fixture file under four configurations. The corpus test
 * (tests/xlsx/formula-import-corpus.test.ts) recomputes them, so a change of behaviour shows up as a diff of this file.
 *   npx tsx tests/oracle/import-values/make-expected.ts            (expected.json)
 *   npx tsx tests/oracle/import-values/make-expected.ts --copies   (expected-copies.json: the same with PremergeOptions.includeSourceSheets,
 *                                                                  checked by tests/xlsx/formula-import-copies-corpus.test.ts)
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { loadManifest, runCorpusFile, type RunCfg } from './lib';

export const CONFIGS: Array<RunCfg & { name: string }> = [
  { name: 'cached-default-left', results: 'cached', imports: 'default', join: 'left' },
  { name: 'cached-all-left', results: 'cached', imports: 'all', join: 'left' },
  { name: 'nocache-default-left', results: 'nocache', imports: 'default', join: 'left' },
  { name: 'nocache-default-full', results: 'nocache', imports: 'default', join: 'full' },
];

async function main(): Promise<void> {
  const dir = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../../fixtures/formulaimport');
  const out: Record<string, unknown> = {};
  const copies = process.argv.includes('--copies');
  for (const e of loadManifest(dir)) {
    for (const cfg of CONFIGS) {
      const r = await runCorpusFile(dir, e.file, cfg, e.base ?? 'base.xlsx', copies ? { includeSourceSheets: true } : {});
      out[`${e.file}|${cfg.name}`] = r.counts;
    }
  }
  fs.writeFileSync(path.join(dir, copies ? 'expected-copies.json' : 'expected.json'), JSON.stringify(out, null, 1));
  console.log(`wrote ${Object.keys(out).length} entries`);
}
if (process.argv[1] && /make-expected\.ts$/.test(process.argv[1])) void main();
