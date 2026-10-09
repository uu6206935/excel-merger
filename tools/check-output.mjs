#!/usr/bin/env node
// Verdict for an exported workbook vs the base input it was made from (owner: T).
//
//   node tools/check-output.mjs <output.xlsx> --base <base.xlsx> [--sheet 名簿] [--header-row 3] [--json] [--no-recalc] [--self-check]
//
// Opens both files in our own hidden Excel (tools/inspect.ps1, recalculated), then checks:
//   * the output opens WITHOUT repair
//   * no new #REF! (formula text, names, CF, DV, chart series) and no new #REF! / #NAME? values
//   * sheets, defined names, charts, pivots, tables still exist
//   * formulas (other sheets + the merged sheet) still reference columns with the same HEADER TEXT
//   * CF / DV of the merged sheet still sit on the same header columns; freeze panes / filter kept
//   * what Excel reports is cross-checked against the package itself (tools/lib/dump-check.mjs); a disagreement means Excel answered
//     while busy and the file is inspected again in a fresh Excel instance (the old "one good output flagged FAIL" flake)
//   * --self-check: both files are inspected a second time in another Excel instance and the two readings must be identical
//     (exit 2 when they are not: the inspection itself is unreliable on this machine right now)
// Exit code: 0 = PASS, 1 = FAIL, 2 = could not run.
import path from 'node:path';
import fs from 'node:fs';
import { inspectFiles } from './lib/inspect.mjs';
import { dumpDifferences } from './lib/dump-check.mjs';
import { checkOutput, formatVerdict } from './lib/verdict.mjs';

function usage(msg) {
  if (msg) process.stderr.write(msg + '\n');
  process.stderr.write('usage: node tools/check-output.mjs <output.xlsx> --base <base.xlsx> [--sheet <name>] [--header-row N] [--json] [--no-recalc] [--self-check]\n');
  process.exit(2);
}

const args = process.argv.slice(2);
let output = null;
let base = null;
let sheet;
let headerRow;
let asJson = false;
let recalc = true;
let dumpDir = null;
let selfCheck = false;
for (let i = 0; i < args.length; i++) {
  const a = args[i];
  if (a === '--base') base = args[++i];
  else if (a === '--sheet') sheet = args[++i];
  else if (a === '--header-row') headerRow = Number(args[++i]);
  else if (a === '--json') asJson = true;
  else if (a === '--no-recalc') recalc = false;
  else if (a === '--dump') dumpDir = args[++i];
  else if (a === '--self-check') selfCheck = true;
  else if (a.startsWith('--')) usage(`unknown option ${a}`);
  else if (!output) output = a;
  else usage(`unexpected argument ${a}`);
}
if (!output || !base) usage();
for (const f of [output, base]) if (!fs.existsSync(f)) usage(`ファイルがありません: ${f}`);

try {
  const files = [path.resolve(base), path.resolve(output)];
  const [b, o] = await inspectFiles(files, { recalc });
  let self = null;
  if (selfCheck) {
    // a second, independent reading (fresh Excel instance): the two must agree in every detail
    const [b2, o2] = await inspectFiles(files, { recalc });
    self = { base: dumpDifferences(b, b2), output: dumpDifferences(o, o2) };
    if (self.base.length || self.output.length) {
      process.stderr.write(`--self-check: Excel の読み取りが 2 回で一致しませんでした（検査が不安定です）
  基準: ${self.base.join(', ') || '同じ'}
  出力: ${self.output.join(', ') || '同じ'}
`);
      process.exit(2);
    }
  }
  if (dumpDir) {
    fs.mkdirSync(dumpDir, { recursive: true });
    fs.writeFileSync(path.join(dumpDir, 'base.json'), JSON.stringify(b, null, 1), 'utf8');
    fs.writeFileSync(path.join(dumpDir, 'output.json'), JSON.stringify(o, null, 1), 'utf8');
  }
  const v = checkOutput(b, o, { sheet, headerRow });
  if (asJson) process.stdout.write(JSON.stringify({ output: path.resolve(output), base: path.resolve(base), openMode: o.openMode, ...(self ? { selfCheck: 'identical' } : {}), ...v }, null, 2) + '\n');
  else process.stdout.write(formatVerdict(v, path.basename(output)) + (self ? '\n(--self-check: 2 回の読み取りは同じでした)' : '') + '\n');
  process.exit(v.ok ? 0 : 1);
} catch (e) {
  process.stderr.write(`チェックを実行できませんでした: ${e instanceof Error ? e.stack ?? e.message : String(e)}\n`);
  process.exit(2);
}
