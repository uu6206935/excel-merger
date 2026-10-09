/**
 * Excel side of the consistency oracle (task V1), driven from Node:
 *  - runBatch(): ONE hidden Excel of our own per batch of 30-50 workbooks (excel-batch.ps1: open, CalculateFull, dump values + formulas of
 *    every sheet / "reduced" copies / resave) — a repair prompt is a failure;
 *  - runRunnerJob(): the app's "Excel本体" engine (server/excel/excel-runner.ps1 through server/excel/run-job.mjs) replaying a LoggedOp[].
 * Both give the Excel process its own TEMP so that the ~80 MB of Office diagnostics every start writes never fill the drive: the folder is
 * removed when that Excel has exited. Only the PID reported by our own instance is ever killed.
 */
import { spawn } from 'node:child_process';
import fs from 'node:fs';
import path from 'node:path';
import { POWERSHELL, killExcelPid, runRunner, type RunnerResult } from '../../../server/excel/run-job.mjs';
import { logToRunnerJob } from '../../../src/core/export/comJob';
import type { LoggedOp } from '../../../src/core/types';
import { HERE, REPO } from './lib';

export type XVal = null | boolean | number | string | { e: number };

export interface BatchJob {
  id: string;
  file: string;
  kind: 'dump' | 'reduced' | 'resave';
  sheet?: string;
  deleteRows?: number[];
  saveAs?: string;
}

export interface SheetDump {
  name: string;
  rows: number;
  cols: number;
  /** address of a circular reference Excel reports on this sheet ('' = none) */
  circ?: string;
  v: XVal[][];
  f: string[][];
}
export interface BookDump {
  open: 'normal' | 'repair' | 'failed';
  log: string;
  sheets: SheetDump[];
}

export const ERR: Record<number, string> = {
  [-2146826281]: '#DIV/0!', [-2146826246]: '#N/A', [-2146826259]: '#NAME?', [-2146826288]: '#NULL!', [-2146826252]: '#NUM!', [-2146826265]: '#REF!',
  [-2146826273]: '#VALUE!', [-2146826233]: '#GETTING_DATA', [-2146826238]: '#CALC!', [-2146826239]: '#SPILL!', [-2146826240]: '#FIELD!',
};

/** the value of a dumped cell with errors as text */
export function norm(v: XVal | undefined): null | boolean | number | string {
  if (v === undefined || v === null) return null;
  if (typeof v === 'object') return ERR[v.e] ?? `#E${v.e}`;
  return v;
}

const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms));

/** free bytes of the drive of `dir` (Infinity when unknown) */
export function freeBytes(dir: string): number {
  try {
    const s = fs.statfsSync(dir);
    return Number(s.bavail) * Number(s.bsize);
  } catch {
    return Number.POSITIVE_INFINITY;
  }
}

export function ensureDisk(dir: string, minGB = 1.5): void {
  const free = freeBytes(dir);
  if (free < minGB * 1024 ** 3) throw new Error(`Only ${(free / 1048576).toFixed(0)} MB free on the drive of ${dir}: the Excel steps need ${minGB} GB (every Excel start writes ~80 MB of Office diagnostics).`);
}

async function rmrf(dir: string): Promise<void> {
  for (let i = 0; i < 6; i++) {
    try {
      fs.rmSync(dir, { recursive: true, force: true });
      if (!fs.existsSync(dir)) return;
    } catch {
      /* Excel may still hold a log file */
    }
    await sleep(700);
  }
}

export interface BatchOptions {
  /** where the per-Excel TEMP folders are created (must be on a roomy drive) */
  tempRoot: string;
  batchSize?: number;
  log?: (s: string) => void;
  noRetry?: boolean;
}

/**
 * Runs the jobs in batches — one hidden Excel per batch. Dumps land in <outDir>/<id>.values.json (a job whose dump exists is skipped, so a run
 * can be resumed).
 */
export async function runBatch(jobs: BatchJob[], outDir: string, opts: BatchOptions): Promise<void> {
  fs.mkdirSync(outDir, { recursive: true });
  fs.mkdirSync(opts.tempRoot, { recursive: true });
  const log = opts.log ?? (() => {});
  const todo = jobs.filter((j) => !fs.existsSync(path.join(outDir, `${j.id}.values.json`)));
  const size = opts.batchSize ?? 40;
  for (let b = 0; b * size < todo.length; b++) {
    const chunk = todo.slice(b * size, (b + 1) * size);
    ensureDisk(opts.tempRoot);
    const tmp = path.join(opts.tempRoot, `xlb-${process.pid}-${b}`);
    fs.mkdirSync(tmp, { recursive: true });
    const jobsFile = path.join(tmp, 'jobs.json');
    fs.writeFileSync(jobsFile, JSON.stringify(chunk), 'utf8');
    const ps1 = path.join(HERE, 'excel-batch.ps1');
    const t0 = Date.now();
    let pid = 0;
    let stderr = '';
    await new Promise<void>((resolve, reject) => {
      const child = spawn(POWERSHELL, ['-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', ps1, '-JobsFile', jobsFile, '-OutDir', outDir], {
        windowsHide: true,
        stdio: ['ignore', 'pipe', 'pipe'],
        env: { ...process.env, TEMP: tmp, TMP: tmp },
      });
      let buf = '';
      child.stdout.setEncoding('utf8');
      child.stdout.on('data', (d: string) => {
        buf += d;
        let nl;
        while ((nl = buf.indexOf('\n')) >= 0) {
          const line = buf.slice(0, nl).trim();
          buf = buf.slice(nl + 1);
          try {
            const o = JSON.parse(line);
            if (o.type === 'excel') pid = o.pid;
          } catch {
            /* log line */
          }
        }
      });
      child.stderr.on('data', (d: Buffer) => (stderr += d.toString('latin1')));
      // chunk * 30 s + 5 min: a hung dialog must not hang the run forever
      const timer = setTimeout(async () => {
        if (pid) await killExcelPid(pid);
        child.kill();
      }, chunk.length * 30_000 + 300_000);
      child.on('error', reject);
      child.on('close', () => {
        clearTimeout(timer);
        resolve();
      });
    });
    // anything the batch did not produce (a crash / kill) is reported; the caller sees the missing dumps
    const missing = chunk.filter((j) => !fs.existsSync(path.join(outDir, `${j.id}.values.json`)));
    log(`excel batch ${b + 1}: ${chunk.length - missing.length}/${chunk.length} dumps in ${((Date.now() - t0) / 1000).toFixed(0)} s${missing.length ? ` (MISSING ${missing.length}: ${missing.slice(0, 3).map((m) => m.id).join(', ')})` : ''}${stderr.trim() ? ` stderr: ${stderr.trim().slice(0, 300)}` : ''}`);
    await rmrf(tmp);
  }
  // a job that failed three times inside a batch gets one more chance in a fresh Excel (a COM hiccup of a long-running instance)
  const failed = todo.filter((j) => {
    const d = readDump(outDir, j.id);
    return d === null || d.open === 'failed';
  });
  if (failed.length && !opts.noRetry) {
    log(`retrying ${failed.length} failed jobs in a fresh Excel`);
    for (const j of failed) fs.rmSync(path.join(outDir, `${j.id}.values.json`), { force: true });
    await runBatch(failed, outDir, { ...opts, noRetry: true });
  }
}

export function readDump(outDir: string, id: string): BookDump | null {
  const p = path.join(outDir, `${id}.values.json`);
  if (!fs.existsSync(p)) return null;
  return JSON.parse(fs.readFileSync(p, 'utf8').replace(/^﻿/, '')) as BookDump;
}

// ---------------------------------------------------------------------------------------------------------------------
// the "Excel本体" engine
// ---------------------------------------------------------------------------------------------------------------------

export const RUNNER = path.join(REPO, 'server', 'excel', 'excel-runner.ps1');

export interface RunnerOutcome {
  ok: boolean;
  code?: string;
  message?: string;
  warnings: Array<{ code: string; message: string }>;
  ms: number;
  refErrors?: { cells: number; names: number; samples: string[] };
}

/**
 * Replays `log` on the premerged package `input` with the app's Excel engine and saves the result to `output`. One Excel start per call
 * (that is how the runner works); its TEMP is private and removed afterwards.
 */
export async function runRunnerJob(input: string, log: LoggedOp[], output: string, sheet: string, opts: { tempRoot: string; cacheDir: string; timeoutMs?: number }): Promise<RunnerOutcome> {
  ensureDisk(opts.tempRoot);
  const tmp = path.join(opts.tempRoot, `xlr-${process.pid}-${Date.now().toString(36)}`);
  fs.mkdirSync(tmp, { recursive: true });
  fs.mkdirSync(opts.cacheDir, { recursive: true });
  const job = logToRunnerJob(log, { format: 'xlsx', sheet });
  const work = path.join(tmp, 'work');
  const jobFile = path.join(tmp, 'job.json');
  fs.mkdirSync(path.dirname(output), { recursive: true });
  fs.writeFileSync(jobFile, JSON.stringify({ ...job, input, output, workDir: work }), 'utf8');
  const warnings: Array<{ code: string; message: string }> = [];
  const t0 = Date.now();
  let res: RunnerResult | null = null;
  try {
    res = await runRunner({
      runner: RUNNER,
      jobFile,
      env: { XLRUNNER_CACHE: opts.cacheDir, TEMP: tmp, TMP: tmp },
      timeoutMs: opts.timeoutMs ?? 5 * 60_000,
      idleMs: 3 * 60_000,
      earlyResult: true,
      exitGraceMs: 8000,
      onEvent: (e) => {
        if (e.type === 'warn') warnings.push({ code: String(e.code), message: String(e.message) });
      },
    });
    const out: RunnerOutcome = { ok: !!res.ok, code: res.code, message: res.message, warnings, ms: Date.now() - t0, refErrors: res.refErrors };
    return out;
  } finally {
    // wait for powershell.exe / our Excel to be gone before the folder (Office diagnostics) is removed
    try {
      if (res?.closed) await Promise.race([res.closed, sleep(60_000)]);
    } catch {
      /* ignore */
    }
    await rmrf(tmp);
  }
}
