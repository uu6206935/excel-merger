#!/usr/bin/env node
// node tools/inspect.mjs <file...> [--recalc] [--json <dir>] [--max-cells N]
// Opens each workbook in our own hidden Excel (tools/inspect.ps1) and prints a short summary; with --json the
// full dumps are written to <dir>/<file>.json (owner: T).
import fs from 'node:fs';
import path from 'node:path';
import { inspectFiles } from './lib/inspect.mjs';
import { summarize } from './lib/compare-core.mjs';

const args = process.argv.slice(2);
const files = [];
let jsonDir = null;
let recalc = false;
let maxCells;
for (let i = 0; i < args.length; i++) {
  const a = args[i];
  if (a === '--recalc') recalc = true;
  else if (a === '--json') jsonDir = args[++i];
  else if (a === '--max-cells') maxCells = Number(args[++i]);
  else files.push(a);
}
if (!files.length) {
  process.stderr.write('usage: node tools/inspect.mjs <file...> [--recalc] [--json <dir>] [--max-cells N]\n');
  process.exit(2);
}
const results = await inspectFiles(files, { recalc, maxCells });
let bad = 0;
results.forEach((r) => {
  if (jsonDir) {
    fs.mkdirSync(jsonDir, { recursive: true });
    const out = path.join(jsonDir, path.basename(r.path ?? r.file) + '.json');
    fs.writeFileSync(out, JSON.stringify(r, null, 1), 'utf8');
  }
  if (r.openMode !== 'normal') bad++;
  process.stdout.write(summarize(r) + '\n');
});
process.exit(bad ? 1 : 0);
