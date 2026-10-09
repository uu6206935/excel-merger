// Writes the .xlsx packages the reader GENERATES for legacy inputs (xls / ods / xlsb / XML 2003 / HTML / csv / tsv) into
// tests/fixtures/xlsx/generated/, so com/verify-open.ps1 can check that Excel opens them without repair.
// Run: npx tsx tests/fixtures/xlsx/gen/make-generated.ts
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { readBook } from '../../../../src/core/xlsx/read';

const here = dirname(fileURLToPath(import.meta.url));
const dir = join(here, '..');
const out = join(dir, 'generated');
mkdirSync(out, { recursive: true });

const inputs = ['legacy.xls', 'legacy.ods', 'legacy.xlsb', 'legacy2003.xls', 'html-sjis.xls', 'torture.xls', 'date1904.xls', 'sjis.csv', 'utf8bom.csv', 'plain.tsv'];
for (const f of inputs) {
  const book = await readBook(new Uint8Array(readFileSync(join(dir, f))), f);
  const name = f.replace(/\./g, '_') + '.xlsx';
  writeFileSync(join(out, name), book.bytes);
  console.log('wrote', name, book.bytes.length);
}
