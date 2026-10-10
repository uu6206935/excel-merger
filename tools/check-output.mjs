#!/usr/bin/env node
// Verdict for an exported workbook vs the base input it was made from (owner: T).
//
//   node tools/check-output.mjs <output.xlsx> --base <base.xlsx> [--sheet 名簿] [--header-row 3] [--json] [--no-recalc] [--self-check]
//        [--single-sheet] [--with-copies --source <other.xlsx> [--source-name <name>] ...]
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
//
// Output shapes (the app's 「他のシートも出力に含める」 / 「他のファイルのシートも含める」 settings):
//   (default)        the output keeps the base file's sheets: a sheet that is gone is a FAILURE (what the outputs of the older rounds looked like)
//   --single-sheet   the output holds ONLY the merged sheet (the app's default): the base's other sheets, and the names / charts / validation lists /
//                    conditional formats that belonged to them are INFO, not FAIL; a formula that still reads a sheet that is not in the output FAILS
//   --with-copies    the output also holds the sheets of the other files, named <file>_<sheet> ("(1)" "(2)" … when a name is taken): every one of them
//                    must be there under the name the rule gives and must equal its source sheet. The source files are given with --source, in the
//                    order of the cards (--source-name <name> right after a --source = the name the app saw; default: the file's own name)
//   the two flags combine (merged sheet + the copies only). The naming rule is tools/lib/sheetnames.mjs (= src/core/xlsx/sheetnames.ts).
//
// Exit code: 0 = PASS, 1 = FAIL, 2 = could not run.
import path from 'node:path';
import fs from 'node:fs';
import { inspectFiles } from './lib/inspect.mjs';
import { dumpDifferences } from './lib/dump-check.mjs';
import { checkOutput, formatVerdict } from './lib/verdict.mjs';
import { isStrictOoxml } from './lib/strict.mjs';

function usage(msg) {
  if (msg) process.stderr.write(msg + '\n');
  process.stderr.write(
    'usage: node tools/check-output.mjs <output.xlsx> --base <base.xlsx> [--sheet <name>] [--header-row N] [--json] [--no-recalc] [--self-check]\n' +
      '       [--single-sheet] [--with-copies --source <file> [--source-name <name>] ...]\n',
  );
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
let singleSheet = false;
let withCopies = false;
/** @type {{path: string, name: string}[]} */
const sourceArgs = [];
for (let i = 0; i < args.length; i++) {
  const a = args[i];
  if (a === '--base') base = args[++i];
  else if (a === '--sheet') sheet = args[++i];
  else if (a === '--header-row') headerRow = Number(args[++i]);
  else if (a === '--json') asJson = true;
  else if (a === '--no-recalc') recalc = false;
  else if (a === '--dump') dumpDir = args[++i];
  else if (a === '--self-check') selfCheck = true;
  else if (a === '--single-sheet') singleSheet = true;
  else if (a === '--with-copies') withCopies = true;
  else if (a === '--source') {
    const p = args[++i];
    if (!p) usage('--source needs a file');
    sourceArgs.push({ path: p, name: path.basename(p) });
  } else if (a === '--source-name') {
    if (!sourceArgs.length) usage('--source-name must follow a --source');
    sourceArgs[sourceArgs.length - 1].name = args[++i] ?? sourceArgs[sourceArgs.length - 1].name;
  } else if (a.startsWith('--')) usage(`unknown option ${a}`);
  else if (!output) output = a;
  else usage(`unexpected argument ${a}`);
}
if (!output || !base) usage();
if (sourceArgs.length && !withCopies) usage('--source は --with-copies と一緒に使います');
if (withCopies && !sourceArgs.length) {
  process.stderr.write('(--with-copies だけでは、コピーの名前の確認はできても元のシートとの比較ができません。--source <元のファイル> を付けてください)\n');
}
for (const f of [output, base, ...sourceArgs.map((s) => s.path)]) if (!fs.existsSync(f)) usage(`ファイルがありません: ${f}`);

try {
  const files = [path.resolve(base), path.resolve(output), ...sourceArgs.map((s) => path.resolve(s.path))];
  const [b, o, ...srcDumps] = await inspectFiles(files, { recalc });
  let self = null;
  if (selfCheck) {
    // a second, independent reading (fresh Excel instance): the two must agree in every detail
    const [b2, o2] = await inspectFiles(files.slice(0, 2), { recalc });
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
    srcDumps.forEach((d, i) => fs.writeFileSync(path.join(dumpDir, `source-${i + 1}.json`), JSON.stringify(d, null, 1), 'utf8'));
  }
  // (a Strict Open XML source / output: the app cannot copy charts / pictures between it and a normal workbook, and says so)
  const strictOf = async (f) => (withCopies ? isStrictOoxml(path.resolve(f)) : false);
  const sources = [];
  for (let i = 0; i < sourceArgs.length; i++) sources.push({ name: sourceArgs[i].name, dump: srcDumps[i], strict: await strictOf(sourceArgs[i].path) });
  const v = checkOutput(b, o, { sheet, headerRow, singleSheet, withCopies, sources, outStrict: await strictOf(output) });
  if (asJson) {
    process.stdout.write(
      JSON.stringify(
        {
          output: path.resolve(output),
          base: path.resolve(base),
          openMode: o.openMode,
          ...(singleSheet ? { singleSheet: true } : {}),
          ...(withCopies ? { withCopies: true, sources: sourceArgs.map((s) => s.name) } : {}),
          ...(self ? { selfCheck: 'identical' } : {}),
          ...v,
        },
        null,
        2,
      ) + '\n',
    );
  } else process.stdout.write(formatVerdict(v, path.basename(output)) + (self ? '\n(--self-check: 2 回の読み取りは同じでした)' : '') + '\n');
  process.exit(v.ok ? 0 : 1);
} catch (e) {
  process.stderr.write(`チェックを実行できませんでした: ${e instanceof Error ? e.stack ?? e.message : String(e)}\n`);
  process.exit(2);
}
