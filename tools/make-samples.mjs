#!/usr/bin/env node
// `npm run samples` — (re)generates the sample / fixture files in samples/ (owner: T).
//
//   node tools/make-samples.mjs            all files (needs Microsoft Excel for the .xlsx/.xls files)
//   node tools/make-samples.mjs --no-excel only the CSV and truth.json (no Excel needed)
//   node tools/make-samples.mjs --out <dir> write somewhere else than samples/
//
// Data: tools/lib/sample-data.mjs (deterministic). Workbooks: tools/make-samples.ps1 in our own hidden Excel
// (ASCII temp paths, PID-scoped watchdog; Excel is never killed by name). CSV: Shift_JIS via encoding-japanese.
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import Encoding from 'encoding-japanese';
import { buildSpec, buildTruth, toCsv, FILES } from './lib/sample-data.mjs';
import { runPs } from './lib/ps.mjs';

const here = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(here, '..');
const args = process.argv.slice(2);
const noExcel = args.includes('--no-excel');
const outIdx = args.indexOf('--out');
const outDir = outIdx >= 0 ? path.resolve(args[outIdx + 1]) : path.join(root, 'samples');

function log(msg) {
  process.stdout.write(msg + '\n');
}

/** Shift_JIS (cp932) bytes of a string; throws when a character cannot be encoded. */
export function encodeSjis(text) {
  const codes = Encoding.convert(Encoding.stringToCode(text), { to: 'SJIS', from: 'UNICODE' });
  const bytes = Uint8Array.from(codes);
  const back = Encoding.codeToString(Encoding.convert(Array.from(bytes), { to: 'UNICODE', from: 'SJIS' }));
  if (back !== text) {
    const i = [...text].findIndex((ch, k) => ch !== [...back][k]);
    throw new Error(`Shift_JIS に変換できない文字があります: ${JSON.stringify([...text][i])}`);
  }
  return bytes;
}

function copyInto(src, dest) {
  try {
    fs.copyFileSync(src, dest);
  } catch (e) {
    if (e && (e.code === 'EBUSY' || e.code === 'EPERM')) {
      throw new Error(`${path.basename(dest)} を上書きできません（Excelで開いていませんか？閉じてから再実行してください）`);
    }
    throw e;
  }
}

async function main() {
  fs.mkdirSync(outDir, { recursive: true });
  const spec = buildSpec();

  // 03_連絡先.csv (Shift_JIS, CRLF) and truth.json never need Excel
  fs.writeFileSync(path.join(outDir, FILES.contacts), encodeSjis(toCsv(spec.contacts.headers, spec.contacts.rows)));
  log(`wrote ${FILES.contacts}`);
  fs.writeFileSync(path.join(outDir, 'truth.json'), JSON.stringify(buildTruth(spec), null, 2) + '\n', 'utf8');
  log('wrote truth.json');
  if (noExcel) return;

  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'em-samples-'));
  if (/[^\x20-\x7e]/.test(tmp)) log(`warning: temp path is not ASCII (${tmp}); Excel COM may be less reliable`);
  try {
    const specPath = path.join(tmp, 'spec.json');
    fs.writeFileSync(specPath, JSON.stringify(spec), 'utf8');
    const res = await runPs(path.join(here, 'make-samples.ps1'), ['-SpecJson', specPath, '-OutDir', tmp], {
      timeoutMs: 6 * 60 * 1000,
      onLine: (o) => {
        if (o.type === 'excel') log(`excel pid ${o.pid} (version ${o.version})`);
        else if (o.type === 'step') log(`  ${o.message}`);
        else if (o.type === 'saved') log(`  saved ${o.file}`);
        else if (o.type === 'error') log(`  ERROR ${o.message}\n${o.at ?? ''}`);
        else if (o.type === 'text') log(`  ${o.text}`);
      },
    });
    if (res.timedOut) throw new Error('Excel の処理がタイムアウトしました（自分で起動した Excel だけを終了しました）');
    if (res.code !== 0 || !res.lines.some((l) => l.type === 'done')) {
      throw new Error(`make-samples.ps1 failed (exit ${res.code})\n${res.stderr}`);
    }
    const pairs = [
      ['s01.xlsx', FILES.base],
      ['s02.xlsx', FILES.interview],
      ['s04.xls', FILES.aptitude],
      ['s05.xlsx', FILES.attendance],
    ];
    for (const [from, to] of pairs) {
      copyInto(path.join(tmp, from), path.join(outDir, to));
      log(`wrote ${to}`);
    }
  } finally {
    fs.rmSync(tmp, { recursive: true, force: true });
  }
}

const isMain = process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url);
if (isMain) {
  main().then(
    () => log('done'),
    (e) => {
      process.stderr.write(`サンプルの生成に失敗しました: ${e instanceof Error ? e.message : String(e)}\n`);
      process.exit(1);
    },
  );
}
