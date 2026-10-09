#!/usr/bin/env node
/**
 * Fake excel-runner for the E2E fallback specs (owner: T). The second E2E server (playwright.config.ts) starts with
 *   EXCEL_DETECT=installed   (the UI offers 「Excel本体で出力」)
 *   EXCEL_RUNNER=<this file> (run-job.mjs starts it as `node fake-excel-runner.mjs <job.json>`)
 *   E2E_FAKE_CONTROL=<json>  (what to do; written by the spec before every export, see setFakeExcel in e2e/round5.ts)
 * so that the real UI, the real server and the real job protocol can be driven through every way the Excel engine fails
 * without starting Excel. Control file: {"mode":"fail","code":"MOVE_VERIFY_FAILED","message":"…"} |
 *   {"mode":"crash"}                     exits with code 3 and no final line
 *   {"mode":"slow","ms":60000}           progress lines for `ms` (lets the spec press 中止), then fails
 *   {"mode":"noout"}                     reports success without writing the output file
 * Anything else (or no control file): fails with OPEN_FAILED.
 */
import fs from 'node:fs';

const emit = (o) => process.stdout.write(JSON.stringify({ ...o, t: Date.now() }) + '\n');
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

let ctl = {};
try {
  ctl = JSON.parse(fs.readFileSync(process.env.E2E_FAKE_CONTROL ?? '', 'utf8'));
} catch {
  /* no control file: default failure */
}

const jobPath = process.argv[2];
let job = {};
try {
  job = JSON.parse(fs.readFileSync(jobPath, 'utf8'));
} catch {
  /* the server always passes a job file; tolerate a probe call */
}

emit({ type: 'start', runner: 'fake-e2e' });
emit({ type: 'progress', phase: 'excel', version: '16.0', pid: 0 });
emit({ type: 'progress', phase: 'open', ms: 1 });

if (ctl.mode === 'crash') process.exit(3);

if (ctl.mode === 'slow') {
  const until = Date.now() + Number(ctl.ms || 60_000);
  let i = 0;
  while (Date.now() < until) {
    emit({ type: 'progress', phase: 'op', i: i++, n: 1000, op: 'moveCols', ms: 1 });
    await sleep(250);
  }
}

if (ctl.mode === 'noout') {
  emit({ type: 'progress', phase: 'save', ms: 1 });
  emit({ type: 'result', ok: true, output: job.output, warnings: [] });
  process.exit(0);
}

emit({
  type: 'error',
  ok: false,
  code: ctl.code || 'OPEN_FAILED',
  message: ctl.message || 'Excel でブックを開けませんでした（E2E のテスト用の失敗です）',
  opIndex: 0,
  warnings: [],
});
process.exit(5);
