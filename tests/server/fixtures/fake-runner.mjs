#!/usr/bin/env node
/**
 * Fake excel-runner for the server tests (EXCEL_RUNNER=<this file>): speaks the runner's JSON-lines protocol
 * without Excel. run-job.mjs starts it as `node fake-runner.mjs <job.json> [0]`.
 *   - copies job.input to job.output
 *   - op {op:'debugWait', seconds}  sleeps (to test the queue / cancel / timeouts)
 *   - op {op:'debugWait', seconds, afterResult:true}  keeps the PROCESS alive after the final result line (a powershell.exe that cannot
 *     exit because its Excel needs ~60 s to shut down: the round 5 performance bug)
 *   - op with sheet 'FAIL'          ends with {"type":"error","code":"EXCEL_REFUSED"} (exit 5)
 *   - op with sheet 'VERIFY'        ends with {"code":"MOVE_VERIFY_FAILED"} and a message that names the built-in engine (exit 5)
 *   - op with sheet 'VERIFYOLD'     same code, but a message WITHOUT the hint (an older runner): the server adds it
 *   - op with sheet 'CRASH'         exits with code 3 and no final line
 *   - op with sheet 'NOOUT'         reports success without writing the output
 *   - op with sheet 'CORRUPT'       reports success with an output that is not a complete package
 *   - always warns "input=<ext>" (lets tests see which extension the server gave the input copy)
 *   - env FAKE_DUMP_JOB=<file>      copies the job file there (tests look at what the server really hands to the runner)
 *   - job.fakeExcel=<EXCEL.EXE copy of node>  (direct runRunner tests only) starts that process, reports it like the real runner reports
 *     its Excel ({"type":"excel","pid","started"}) and leaves it running: the watchdog must kill exactly this PID
 */
import fs from 'node:fs';
import path from 'node:path';
import { spawn } from 'node:child_process';

const emit = (o) => process.stdout.write(JSON.stringify({ ...o, t: Date.now() }) + '\n');
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function main() {
  const jobPath = process.argv[2];
  const job = JSON.parse(fs.readFileSync(jobPath, 'utf8'));
  if (process.env.FAKE_DUMP_JOB) fs.copyFileSync(jobPath, process.env.FAKE_DUMP_JOB);
  emit({ type: 'start', runner: 'fake' });
  if (job.fakeExcel) {
    const c = spawn(job.fakeExcel, ['-e', 'setTimeout(() => {}, 120000)'], { stdio: 'ignore', detached: true, windowsHide: true });
    c.unref();
    await sleep(300);
    emit({ type: 'excel', pid: c.pid, started: Date.now() - 300 });
  }
  emit({ type: 'progress', phase: 'excel', version: '16.0', pid: 0 });
  emit({ type: 'warn', code: 'INPUT_EXT', message: `input=${path.extname(job.input)}` });
  emit({ type: 'progress', phase: 'open', ms: 1 });
  const ops = job.ops || [];
  let noOut = false;
  let corrupt = false;
  for (let i = 0; i < ops.length; i++) {
    const op = ops[i];
    if (op.op === 'debugWait' && !op.afterResult) await sleep(Number(op.seconds || 1) * 1000);
    if (op.sheet === 'FAIL') {
      emit({ type: 'error', ok: false, code: 'EXCEL_REFUSED', message: 'Excel が「列の移動」を実行できませんでした (テスト)', opIndex: i, warnings: [] });
      process.exit(5);
    }
    if (op.sheet === 'VERIFY' || op.sheet === 'VERIFYOLD') {
      const hint = op.sheet === 'VERIFY' ? '「内蔵エンジン」で出力してください（画面の出力設定で選べます）。' : '';
      emit({ type: 'error', ok: false, code: 'MOVE_VERIFY_FAILED', message: `テーブルの中の列の移動を Excel本体では正しく行えなかったため中止しました（保存していません）。${hint}`, opIndex: i, warnings: [] });
      process.exit(5);
    }
    if (op.sheet === 'CRASH') process.exit(3);
    if (op.sheet === 'NOOUT') noOut = true;
    if (op.sheet === 'CORRUPT') corrupt = true;
    emit({ type: 'progress', phase: 'op', i, n: ops.length, op: op.op, ms: 1 });
  }
  emit({ type: 'progress', phase: 'calc', ms: 1 });
  if (corrupt) fs.writeFileSync(job.output, Buffer.concat([Buffer.from([0x50, 0x4b, 0x03, 0x04]), Buffer.alloc(200, 7)]));
  else if (!noOut) fs.copyFileSync(job.input, job.output);
  emit({ type: 'progress', phase: 'save', ms: 1 });
  emit({ type: 'result', ok: true, output: job.output, warnings: [{ code: 'TEST', message: 'テスト警告' }] });
  // op {op:'debugWait', seconds, afterResult:true}: the process cannot exit after its result (a powershell.exe waiting for an Excel
  // that needs ~60 s to shut down: the round 5 performance bug) - the server must deliver the result anyway and end us later
  const stall = ops.find((o) => o && o.op === 'debugWait' && o.afterResult);
  if (stall) await sleep(Number(stall.seconds || 60) * 1000);
}

main().catch((e) => {
  emit({ type: 'error', ok: false, code: 'INTERNAL', message: String(e && e.message) });
  process.exit(9);
});
