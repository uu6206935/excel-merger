// Node wrapper around tools/inspect.ps1 (owner: T): opens workbooks in our own hidden Excel and returns what
// Excel sees (see the JSON shape in tools/lib/inspect.d.mts).
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { runPs } from './ps.mjs';
import { dumpIssues, packageFacts } from './dump-check.mjs';

const here = path.dirname(fileURLToPath(import.meta.url));
export const INSPECT_PS1 = path.resolve(here, '..', 'inspect.ps1');

/**
 * true when Microsoft Excel is registered for COM automation on this machine.
 * QA_NO_EXCEL=1 pretends it is not (runs the checks that need no Excel only: every Excel start writes ~80 MB of
 * Office diagnostics into %TEMP% — set TEMP/TMP to another drive when the system drive is small).
 */
export function excelAvailable() {
  if (process.platform !== 'win32') return false;
  if (process.env.QA_NO_EXCEL === '1') return false;
  try {
    const r = spawnSync('reg', ['query', 'HKCR\\Excel.Application\\CurVer'], { windowsHide: true, encoding: 'latin1', timeout: 10000 });
    return r.status === 0 && /Excel\.Application/i.test(r.stdout || '');
  } catch {
    return false;
  }
}

/**
 * Inspect several workbooks with ONE Excel instance.
 *
 * Every dump of an .xlsx / .xlsm is cross-checked against the package it came from (tools/lib/dump-check.mjs): defined names,
 * formula cells and the ranges they point to, CF / DV / chart / pivot / table presence, reads Excel reported as "busy". A dump
 * that disagrees with its file is an Excel that answered while it was busy (the old check-output flake), so it is inspected again
 * in a FRESH Excel instance (up to `maxAttempts` = 3 in all; an inspection that crashed is retried the same way). The result
 * carries `attempts`, and `suspicious` (the reasons) when the last attempt still disagreed.
 *
 * @param {string[]} files
 * @param {{recalc?: boolean, maxCells?: number, timeoutMs?: number, noRetry?: boolean, maxAttempts?: number, noCrossCheck?: boolean, probeCells?: string[], onLine?: (o: any) => void}} [opts]
 *        probeCells: 'Sheet!A1' cells to look at in detail (result.cellProbes)
 * @returns {Promise<any[]>} one inspect result per file (same order); a file that could not be inspected yields
 *          {file, path, openMode: 'failed', error}
 */
export async function inspectFiles(files, opts = {}) {
  const maxAttempts = opts.noRetry ? 1 : Math.max(1, opts.maxAttempts ?? 3);
  const results = await inspectOnce(files, opts);
  const facts = await Promise.all(files.map((f) => (opts.noCrossCheck ? null : packageFacts(f))));
  const judge = (r, i) => {
    r.attempts = 1;
    const issues = r.error ? [] : dumpIssues(facts[i], r);
    if (issues.length) r.suspicious = issues;
  };
  results.forEach(judge);
  // an inspection that crashed (not a file Excel refuses to open) or disagrees with its package is repeated in a fresh Excel
  // instance: COM calls into a busy Excel fail transiently now and then
  for (let attempt = 2; attempt <= maxAttempts; attempt++) {
    const retry = results.map((r, i) => (r.error || r.suspicious?.length ? i : -1)).filter((i) => i >= 0);
    if (!retry.length) break;
    const again = await inspectOnce(retry.map((i) => files[i]), opts);
    retry.forEach((i, k) => {
      const prev = results[i];
      const next = again[k];
      next.attempts = attempt;
      const issues = next.error ? [] : dumpIssues(facts[i], next);
      if (issues.length) next.suspicious = issues;
      if (next.error) {
        results[i] = { ...next, error: `${next.error}（再試行後も失敗: ${prev.error ?? prev.suspicious?.join(' / ')}）` };
      } else {
        results[i] = next;
      }
    });
  }
  return results;
}

async function inspectOnce(files, opts) {
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'em-inspect-'));
  try {
    const job = {
      files: files.map((f, i) => ({ in: path.resolve(f), out: path.join(tmp, `out-${i}.json`) })),
      recalc: !!opts.recalc,
      maxCells: opts.maxCells ?? 300000,
      ...(opts.probeCells?.length ? { probeCells: opts.probeCells } : {}),
    };
    const jobPath = path.join(tmp, 'job.json');
    fs.writeFileSync(jobPath, JSON.stringify(job), 'utf8');
    const res = await runPs(INSPECT_PS1, ['-JobJson', jobPath], {
      timeoutMs: opts.timeoutMs ?? 10 * 60 * 1000,
      onLine: opts.onLine,
    });
    const errors = new Map();
    for (const l of res.lines) {
      if (l.type === 'error' && l.file) errors.set(path.resolve(l.file), `${l.message}${l.at ? ` @ ${String(l.at).split(/\r?\n/)[0]}` : ''}`);
    }
    return job.files.map((f) => {
      if (fs.existsSync(f.out)) return JSON.parse(fs.readFileSync(f.out, 'utf8'));
      const msg =
        errors.get(f.in) ??
        (res.timedOut ? 'Excel の処理がタイムアウトしました' : `inspect.ps1 failed (exit ${res.code}) ${res.stderr.slice(0, 500)}`);
      return { tool: 'excel-merger/inspect', version: 1, file: path.basename(f.in), path: f.in, openMode: 'failed', error: msg };
    });
  } finally {
    fs.rmSync(tmp, { recursive: true, force: true });
  }
}

export async function inspectFile(file, opts = {}) {
  return (await inspectFiles([file], opts))[0];
}
