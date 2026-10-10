/**
 * Excel side of the copied-sheets oracle (task V9), driven from Node:
 *  - runBooks(): ONE hidden Excel of our own per batch of workbooks (excel-book.ps1: open, CalculateFull, dump values + formulas of every sheet
 *    and, for the sheets named in `deep`, how Excel renders them: formats, widths, objects ...). A repair prompt is a failure.
 *  - runRunner(): the app's "Excel本体" engine (server/excel/excel-runner.ps1 through server/excel/run-job.mjs) replaying a LoggedOp[] —
 *    one Excel start per call, like the app does.
 * Both give the Excel process its own TEMP so that the ~80 MB of Office diagnostics every start writes never fill the drive: the folder is
 * removed when that Excel has exited. Only the PID reported by our own instance is ever killed.
 */
import { spawn } from 'node:child_process';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';
import { POWERSHELL, killExcelPid, runRunner as runRunnerScript, type RunnerResult } from '../../../server/excel/run-job.mjs';
import { logToRunnerJob } from '../../../src/core/export/comJob';
import type { LoggedOp } from '../../../src/core/types';
import { ensureDisk, type XVal } from '../consistency/excel';

export { ensureDisk };

export interface BookJob {
  id: string;
  file: string;
  /** 'all' | 'none' | the names of the sheets that also get the deep dump */
  deep: 'all' | 'none' | string[];
}

export interface DeepLook {
  r: number;
  c: number;
  nf: string;
  fn: string;
  fs: number | null;
  fb: boolean | null;
  fi: boolean | null;
  fc: number | null;
  ic: number | null;
  ip: number | null;
  ha: number | null;
  va: number | null;
  wr: boolean | null;
  mg: boolean | null;
  bb: number;
  bl: number;
  bt: number;
  br: number;
}

export interface Deep {
  tab: number;
  widths?: number[];
  hiddenCols?: number[];
  hiddenRows?: number[];
  look?: DeepLook[];
  cf?: number;
  dv?: string;
  links?: number;
  shapes?: Array<{ t: number; n: string; x: string }>;
  shapesOther?: number;
  charts?: Array<{ type: string; series: Array<{ v: string; x: string }> }>;
  tables?: Array<{ n: string; a: string; cols: string; totals: boolean | null }>;
  notes?: Array<{ a: string; t: string }>;
  pivots?: Array<{ n: string; a: string; src: string; v: string }>;
  chartSheet?: { type: string; series: Array<{ v: string }> };
}

export interface BookSheet {
  name: string;
  /** XlSheetType: -4167 worksheet, -4109 chart sheet, 3 / 4 macro sheet, -4116 dialog sheet */
  type: number;
  /** -1 visible, 0 hidden, 2 very hidden */
  visible: number;
  rows: number;
  cols: number;
  used?: [number, number];
  trunc?: boolean;
  circ?: string;
  v: XVal[][];
  f: string[][];
  deep?: Deep;
  deepError?: string;
}

export interface BookDump {
  open: 'normal' | 'repair' | 'failed';
  log: string;
  date1904: boolean;
  names: Array<{ name: string; refers: string; visible: boolean | null }>;
  sheets: BookSheet[];
}

const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms));

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

export const HERE = path.dirname(fileURLToPath(import.meta.url));

export interface BatchOptions {
  /** where the per-Excel TEMP folders are created (must be on a roomy drive) */
  tempRoot: string;
  batchSize?: number;
  log?: (s: string) => void;
  noRetry?: boolean;
}

/** an Excel that answered while it was busy: no dump, a dump that says "failed", a sheet without a name, a deep dump that threw */
export function isBadDump(d: BookDump | null): boolean {
  return d === null || d.open === 'failed' || d.sheets.some((s) => !s.name || !!s.deepError);
}

/**
 * The last row and column that hold a value / formula in every worksheet of an .xlsx / .xlsm package (by sheet name); null when it cannot be read.
 * (a cheap scan from the end of the sheet data: a row counts when one of its cells has <v>, <f> or <is>)
 */
export async function packageLastCells(file: string): Promise<Map<string, { row: number; col: number }> | null> {
  if (!/\.(xlsx|xlsm)$/i.test(file)) return null;
  try {
    const zip = await JSZip.loadAsync(fs.readFileSync(file));
    const wb = await zip.file('xl/workbook.xml')?.async('string');
    const rels = await zip.file('xl/_rels/workbook.xml.rels')?.async('string');
    if (!wb || !rels) return null;
    const target = new Map<string, string>();
    for (const m of rels.matchAll(/<Relationship\b[^>]*>/g)) {
      const id = /\bId="([^"]*)"/.exec(m[0])?.[1];
      const t = /\bTarget="([^"]*)"/.exec(m[0])?.[1];
      if (id && t) target.set(id, t.replace(/^\//, '').replace(/^(?!xl\/)/, 'xl/'));
    }
    const out = new Map<string, { row: number; col: number }>();
    for (const m of wb.matchAll(/<sheet\b[^>]*>/g)) {
      const name = /\bname="([^"]*)"/.exec(m[0])?.[1];
      const rid = /\br:id="([^"]*)"/.exec(m[0])?.[1] ?? /\bid="([^"]*)"/.exec(m[0])?.[1];
      const part = rid ? target.get(rid) : undefined;
      const xml = part ? await zip.file(part)?.async('string') : undefined;
      if (!name || !xml || !/<worksheet\b/.test(xml)) continue;
      const decoded = name.replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'");
      let row = 0;
      let col = 0;
      let at = xml.lastIndexOf('</sheetData>');
      for (let guard = 0; at > 0 && guard < 2000; guard++) {
        const start = xml.lastIndexOf('<row ', at - 1);
        if (start < 0) break;
        const chunk = xml.slice(start, at);
        at = start;
        if (!/<c\b[^>]*>\s*<(?:v|f|is)\b/.test(chunk)) continue;
        row = Number(/\br="(\d+)"/.exec(chunk)?.[1] ?? 0);
        break;
      }
      for (const c of xml.matchAll(/<c\b[^>]*\br="([A-Z]+)\d+"[^>]*>\s*<(?:v|f|is)\b/g)) {
        let n = 0;
        for (const ch of c[1]) n = n * 26 + (ch.charCodeAt(0) - 64);
        if (n > col) col = n;
      }
      out.set(decoded, { row, col });
    }
    return out;
  } catch {
    return null;
  }
}

/**
 * A dump whose used range ends BEFORE the last cell that holds something in the package: Excel answered before it had finished (the file is fine; seen once in
 * every few hundred dumps as a used range one row short). The dump is read again.
 */
export async function isShortDump(d: BookDump | null, file: string): Promise<boolean> {
  if (!d || !d.sheets) return false;
  const last = await packageLastCells(file);
  if (!last) return false;
  return d.sheets.some((s) => {
    const l = last.get(s.name);
    return !!l && s.type === -4167 && !s.trunc && !!s.used && (s.used[0] < l.row || s.used[1] < l.col);
  });
}

/** isBadDump, and a used range that is shorter than the package says */
export async function isBadDumpOf(d: BookDump | null, file: string): Promise<boolean> {
  return isBadDump(d) || (await isShortDump(d, file));
}

export function dumpPath(outDir: string, id: string): string {
  return path.join(outDir, `${id}.book.json`);
}

export function readBook(outDir: string, id: string): BookDump | null {
  const p = dumpPath(outDir, id);
  if (!fs.existsSync(p)) return null;
  return JSON.parse(fs.readFileSync(p, 'utf8').replace(/^﻿/, '')) as BookDump;
}

/**
 * Runs the jobs in batches — one hidden Excel per batch. Dumps land in <outDir>/<id>.book.json (a job whose dump exists is skipped, so a run
 * can be resumed; a dump that says "failed" is tried once more in a fresh Excel).
 */
export async function runBooks(jobs: BookJob[], outDir: string, opts: BatchOptions): Promise<void> {
  fs.mkdirSync(outDir, { recursive: true });
  fs.mkdirSync(opts.tempRoot, { recursive: true });
  const log = opts.log ?? (() => {});
  // (a dump whose sheets came back without a name — Excel answering while it was busy — is as good as none)
  const todo: BookJob[] = [];
  for (const j of jobs) if (!fs.existsSync(dumpPath(outDir, j.id)) || (await isBadDumpOf(readBook(outDir, j.id), j.file))) todo.push(j);
  for (const j of todo) fs.rmSync(dumpPath(outDir, j.id), { force: true });
  const size = opts.batchSize ?? 24;
  for (let b = 0; b * size < todo.length; b++) {
    const chunk = todo.slice(b * size, (b + 1) * size);
    ensureDisk(opts.tempRoot);
    const tmp = path.join(opts.tempRoot, `xlb-${process.pid}-${b}-${Date.now().toString(36)}`);
    fs.mkdirSync(tmp, { recursive: true });
    const jobsFile = path.join(tmp, 'jobs.json');
    fs.writeFileSync(jobsFile, JSON.stringify(chunk), 'utf8');
    const ps1 = path.join(HERE, 'excel-book.ps1');
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
      // chunk * 90 s + 5 min: a hung dialog must not hang the run forever (a deep dump of a big sheet takes a while)
      const timer = setTimeout(async () => {
        if (pid) await killExcelPid(pid);
        child.kill();
      }, chunk.length * 90_000 + 300_000);
      child.on('error', reject);
      child.on('close', () => {
        clearTimeout(timer);
        resolve();
      });
    });
    const missing = chunk.filter((j) => !fs.existsSync(dumpPath(outDir, j.id)));
    log(`excel batch ${b + 1}: ${chunk.length - missing.length}/${chunk.length} dumps in ${((Date.now() - t0) / 1000).toFixed(0)} s${missing.length ? ` (MISSING ${missing.length}: ${missing.slice(0, 3).map((m) => m.id).join(', ')})` : ''}${stderr.trim() ? ` stderr: ${stderr.trim().slice(0, 300)}` : ''}`);
    await rmrf(tmp);
  }
  // a job that failed three times inside a batch gets one more chance in a fresh Excel (a COM hiccup of a long-running instance)
  const failed: BookJob[] = [];
  for (const j of todo) if (await isBadDumpOf(readBook(outDir, j.id), j.file)) failed.push(j);
  if (failed.length && !opts.noRetry) {
    log(`retrying ${failed.length} failed jobs in a fresh Excel`);
    for (const j of failed) fs.rmSync(dumpPath(outDir, j.id), { force: true });
    await runBooks(failed, outDir, { ...opts, noRetry: true });
  }
}

// ---------------------------------------------------------------------------------------------------------------------
// the "Excel本体" engine
// ---------------------------------------------------------------------------------------------------------------------

const RUNNER = path.resolve(HERE, '..', '..', '..', 'server', 'excel', 'excel-runner.ps1');

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
 * (that is how the runner works); its TEMP is private and removed afterwards. `format` is the app's outputFormatOf(base): 'xlsm' keeps macros.
 */
export async function runRunnerJob(input: string, log: LoggedOp[], output: string, sheet: string, opts: { tempRoot: string; cacheDir: string; format?: 'xlsx' | 'xlsm'; timeoutMs?: number }): Promise<RunnerOutcome> {
  ensureDisk(opts.tempRoot);
  const tmp = path.join(opts.tempRoot, `xlr-${process.pid}-${Date.now().toString(36)}`);
  fs.mkdirSync(tmp, { recursive: true });
  fs.mkdirSync(opts.cacheDir, { recursive: true });
  const job = logToRunnerJob(log, { format: opts.format ?? 'xlsx', sheet });
  const work = path.join(tmp, 'work');
  const jobFile = path.join(tmp, 'job.json');
  fs.mkdirSync(path.dirname(output), { recursive: true });
  fs.writeFileSync(jobFile, JSON.stringify({ ...job, input, output, workDir: work }), 'utf8');
  const warnings: Array<{ code: string; message: string }> = [];
  const t0 = Date.now();
  let res: RunnerResult | null = null;
  try {
    res = await runRunnerScript({
      runner: RUNNER,
      jobFile,
      env: { XLRUNNER_CACHE: opts.cacheDir, TEMP: tmp, TMP: tmp },
      timeoutMs: opts.timeoutMs ?? 6 * 60_000,
      idleMs: 4 * 60_000,
      earlyResult: true,
      exitGraceMs: 8000,
      onEvent: (e) => {
        if (e.type === 'warn') warnings.push({ code: String(e.code), message: String(e.message) });
      },
    });
    return { ok: !!res.ok, code: res.code, message: res.message, warnings, ms: Date.now() - t0, refErrors: res.refErrors };
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
