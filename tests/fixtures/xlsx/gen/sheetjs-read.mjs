// Child-process helper for tests/xlsx/read-legacy.test.ts (owner: reader R): runs SheetJS on a file so the test can
// detect a HANG with a timeout (a synchronous infinite loop inside the test worker could never be interrupted).
// Usage: node sheetjs-read.mjs <file> '<json ParsingOptions>'  → prints {"sheets":n,"cells":n} or {"error":"..."}
import { readFileSync } from 'node:fs';
import * as XLSX from 'xlsx';

const bytes = new Uint8Array(readFileSync(process.argv[2]));
const opts = { type: 'array', ...JSON.parse(process.argv[3] ?? '{}') };
try {
  const wb = XLSX.read(bytes, opts);
  let cells = 0;
  for (const n of wb.SheetNames) cells += Object.keys(wb.Sheets[n] ?? {}).filter((k) => k[0] !== '!').length;
  console.log(JSON.stringify({ sheets: wb.SheetNames.length, cells }));
} catch (e) {
  console.log(JSON.stringify({ error: String(e && e.message ? e.message : e).slice(0, 200) }));
}
