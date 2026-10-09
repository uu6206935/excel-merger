/**
 * Round 5 hardening of the Excel本体 engine against the REAL Excel (skipped unless EXCEL_TESTS=1; ~8 Excel starts, ~1.5 min):
 *   1. Performance: the result of a job is delivered when the runner's final line arrives, not when powershell.exe / Excel have
 *      exited (measured on this machine: 6 of 40 jobs took ~64 s because powershell.exe could not exit while Excel's shutdown stalled).
 *      Every Excel we started is gone shortly afterwards (our own PID, killed by the watchdog if necessary).
 *   2. Injection: typed / imported text that looks like a DDE call, an external workbook link or a web request is written as plain
 *      text (quote prefix), a normal formula stays a formula. A marker file proves that nothing was executed.
 *   3. Plumbing: a job dir with Japanese characters, spaces and typographic apostrophes; the upload opened in place; Excel's private
 *      TEMP (Office diagnostics) deleted afterwards; the server's job API end to end.
 */
import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import fs from 'node:fs';
import path from 'node:path';
import JSZip from 'jszip';
import { pidAlive, runRunner } from '../../server/excel/run-job.mjs';
import { EXCEL_TESTS, rawCells, tempDir } from './lib';
import { makeSandbox, raw, sleep, startServer, waitJob, type RunningServer, type Sandbox } from '../server/helpers';

const ROOT = path.resolve(__dirname, '..', '..');
const RUNNER = path.join(ROOT, 'server', 'excel', 'excel-runner.ps1');
const SAMPLE = path.join(ROOT, 'samples', '01_学生名簿.xlsx');
const SHEET = '名簿';

async function until(check: () => boolean, ms: number): Promise<boolean> {
  const end = Date.now() + ms;
  while (Date.now() < end) {
    if (check()) return true;
    await sleep(200);
  }
  return check();
}

/** A job folder like the server's (work.xlsx + output.xlsx + job.json in one folder), under a hostile name. */
function makeJobDir(root: string, name: string, ops: unknown[]) {
  const dir = path.join(root, name);
  fs.mkdirSync(dir, { recursive: true });
  const input = path.join(dir, 'work.xlsx');
  fs.copyFileSync(SAMPLE, input);
  const output = path.join(dir, 'output.xlsx');
  const jobFile = path.join(dir, 'job.json');
  fs.writeFileSync(jobFile, JSON.stringify({ version: 1, format: 'xlsx', input, output, workDir: dir, options: { recalc: true, clipboard: 'restoreEachOp' }, ops }));
  return { dir, input, output, jobFile };
}

describe.skipIf(!EXCEL_TESTS)('Excel本体 engine hardening (real Excel)', () => {
  let root: string;
  beforeAll(() => {
    root = tempDir('hard');
  });
  afterAll(() => {
    try {
      if (root) fs.rmSync(root, { recursive: true, force: true });
    } catch {
      /* an Excel of a failed test may still hold a file: the test's own failure is what matters */
    }
  });

  it('text that would become a DDE call / external link / web request is written as plain text; normal formulas stay live; nothing is executed', async () => {
    const marker = path.join(root, 'DDE-EXECUTED.txt');
    const dde = `=cmd|' /C echo pwned > "${marker}"'!A0`;
    const rows: unknown[][] = [
      [dde], // N5: bare text (what Value2 = "..." would enter as a formula)
      [{ f: dde }], // N6: a typed formula
      [{ f: "='[x.xlsx]Sheet1'!A1" }], // N7: external workbook link
      [{ f: '=WEBSERVICE("http://127.0.0.1:9/x")' }], // N8
      [{ f: '=SUM(1,2)' }], // N9: stays a formula
      [{ f: '=HYPERLINK("http://example.com","l")' }], // N10: stays a formula
      ['+cmd|\' /C calc\'!A0'], // N11: Lotus-style sign
      [{ t: 's', v: dde, quote: true }], // N12: already quoted text
      [{ f: "='\\\\127.0.0.1\\share\\[x.xlsx]S'!A1" }], // N13: UNC external link
    ];
    const hostile = `日本語 フォルダ'’ ‘x’; New-Item -ItemType File -Path PWNED.txt; $(1+1)`; // quotes of every kind, PowerShell syntax
    const j = makeJobDir(root, hostile, [{ op: 'setCells', sheet: SHEET, at: 'N5', rows }]);
    const warnings: string[] = [];
    const r = await runRunner({
      runner: RUNNER,
      jobFile: j.jobFile,
      earlyResult: true,
      exitGraceMs: 6000,
      extraArgs: [{ name: 'QuitWaitMs', value: '0' }],
      env: { XLRUNNER_CACHE: path.join(root, 'cache') },
      onEvent: (e) => {
        if (e.type === 'warn') warnings.push(String(e.code));
      },
      timeoutMs: 240_000,
    });
    expect(r.ok, JSON.stringify(r)).toBe(true);
    expect(r.earlyResult).toBe(true);
    expect(warnings.filter((c) => c === 'UNSAFE_FORMULA'), `warnings: ${warnings.join(',')}`).toHaveLength(6); // N5 N6 N7 N8 N11 N13
    await r.closed;
    expect(pidAlive(r.excelPid), 'our Excel is gone after the shutdown / watchdog').toBe(false);
    expect(fs.existsSync(marker), 'the DDE payload must never run').toBe(false);
    expect(fs.existsSync(path.join(process.cwd(), 'PWNED.txt'))).toBe(false);
    expect(fs.existsSync(path.join(j.dir, 'PWNED.txt'))).toBe(false);
    expect(fs.statSync(j.output).size).toBeGreaterThan(10_000);
    // the upload opened in place is the server's private copy (it deletes the folder); Excel saved straight into output.xlsx, no saved.xlsx
    expect(fs.readdirSync(j.dir).sort()).toEqual(['job.json', 'output.xlsx', 'work.xlsx']);

    const bytes = fs.readFileSync(j.output);
    const cells = await rawCells(bytes, SHEET);
    for (const a of ['N5', 'N6', 'N7', 'N8', 'N11', 'N12', 'N13']) {
      const c = cells.get(a);
      expect(c, a).toBeDefined();
      expect(c!.formula, `${a} must be text, not a formula`).toBeUndefined();
      expect(c!.quotePrefix, a).toBe(true);
    }
    expect(cells.get('N5')!.value).toBe(dde);
    expect(cells.get('N6')!.value).toBe(dde);
    expect(cells.get('N7')!.value).toBe("='[x.xlsx]Sheet1'!A1");
    expect(cells.get('N11')!.value).toBe("+cmd|' /C calc'!A0");
    expect(cells.get('N9')!.formula).toBe('SUM(1,2)');
    expect(cells.get('N10')!.formula).toContain('HYPERLINK');
    const zip = await JSZip.loadAsync(bytes);
    expect(Object.keys(zip.files).filter((f) => /externalLinks|ddeLink|oleObject/i.test(f))).toEqual([]);
  }, 300_000);

  it('10-job soak (same code path as the server): every result arrives in seconds, every Excel started is gone afterwards', async () => {
    const N = 10;
    const walls: number[] = [];
    const lags: number[] = []; // final line arrived -> runRunner resolved (used to be: until powershell.exe + Excel had exited, up to ~60 s)
    const gaps: number[] = []; // 'save' progress -> final line (the runner's own end game)
    const closes: Array<Promise<unknown>> = [];
    const pids: number[] = [];
    const sumStart = Date.now();
    for (let i = 0; i < N; i++) {
      const j = makeJobDir(root, `soak ${i}`, [
        { op: 'unmerge', sheet: SHEET, refs: ['A1:K1'], centerAcross: true }, // the sample's merged title blocks column moves
        { op: 'moveCols', sheet: SHEET, start: 6, count: 2, to: 2, method: 'clipFree' },
        { op: 'insertCols', sheet: SHEET, at: 3, count: 1, formatFrom: 'left' },
        { op: 'setCells', sheet: SHEET, at: 'C4', rows: [['メモ'], ['a']] },
      ]);
      const t0 = Date.now();
      let saveAt = 0;
      let resultAt = 0;
      const r = await runRunner({
        runner: RUNNER,
        jobFile: j.jobFile,
        earlyResult: true,
        exitGraceMs: 6000,
        extraArgs: [{ name: 'QuitWaitMs', value: '0' }],
        env: { XLRUNNER_CACHE: path.join(root, 'cache') },
        onEvent: (e) => {
          if (e.type === 'progress' && e.phase === 'save') saveAt = Date.now();
          if (e.type === 'result') resultAt = Date.now();
        },
        timeoutMs: 240_000,
      });
      const resolvedAt = Date.now();
      walls.push(resolvedAt - t0);
      lags.push(resolvedAt - resultAt);
      gaps.push(resultAt - saveAt);
      expect(r.ok, JSON.stringify(r)).toBe(true);
      expect(fs.statSync(j.output).size).toBeGreaterThan(10_000);
      pids.push(r.excelPid);
      closes.push(r.closed!);
    }
    const total = Date.now() - sumStart;
    // eslint-disable-next-line no-console
    console.log(`soak: walls(ms) ${walls.join(' ')}  max ${Math.max(...walls)}  total ${total}; result line -> delivered ${lags.join(' ')}; save -> result line ${gaps.join(' ')}`);
    // machine-load independent: the result is delivered the moment its line is out, and that line follows 'save' at once
    expect(Math.max(...lags), `lags: ${lags.join(' ')}`).toBeLessThan(1500);
    expect(Math.max(...gaps), `gaps: ${gaps.join(' ')}`).toBeLessThan(8000);
    expect(Math.max(...walls), `walls: ${walls.join(' ')}`).toBeLessThan(45_000); // used to be ~64 s in ~15 % of the jobs (and ~6 s otherwise)
    const closed = (await Promise.all(closes)) as Array<{ excelKilled?: boolean; wallMs: number }>;
    // eslint-disable-next-line no-console
    console.log(`soak: process shutdown took ${closed.map((c) => c.wallMs).join(' ')} ms; watchdog had to kill Excel in ${closed.filter((c) => c.excelKilled).length} of ${N} jobs`);
    for (const pid of pids) expect(pidAlive(pid), `Excel ${pid} still running`).toBe(false);
  }, 600_000);

  it('a stalled shutdown (powershell.exe kept alive after its result): result at once, then exactly OUR Excel is ended; another Excel (the user\'s) is untouched', async () => {
    // A bystander = a second REAL Excel of this test (held alive like the user's own Excel): a snapshot of "all EXCEL.EXE" would be fooled by
    // other tests / programs that start and end Excel meanwhile. The bystander is never given to the watchdog of the stalled job.
    const by = makeJobDir(root, 'bystander', [{ op: 'setCells', sheet: SHEET, at: 'N5', rows: [['x']] }]);
    const bystander = await runRunner({
      runner: RUNNER,
      jobFile: by.jobFile,
      earlyResult: true,
      exitGraceMs: 10 * 60_000,
      extraArgs: [{ name: 'QuitWaitMs', value: '0' }],
      env: { XLRUNNER_CACHE: path.join(root, 'cache'), XLRUNNER_ALLOW_DEBUG: '1', XLRUNNER_HOLD_AFTER_RESULT_MS: '300000' },
      timeoutMs: 240_000,
    });
    expect(bystander.ok, JSON.stringify(bystander)).toBe(true);
    expect(pidAlive(bystander.excelPid)).toBe(true);
    const j = makeJobDir(root, 'stall', [{ op: 'setCells', sheet: SHEET, at: 'N5', rows: [['x']] }]);
    const t0 = Date.now();
    const r = await runRunner({
      runner: RUNNER,
      jobFile: j.jobFile,
      earlyResult: true,
      exitGraceMs: 1500,
      extraArgs: [{ name: 'QuitWaitMs', value: '0' }],
      env: { XLRUNNER_CACHE: path.join(root, 'cache'), XLRUNNER_ALLOW_DEBUG: '1', XLRUNNER_HOLD_AFTER_RESULT_MS: '120000' },
      timeoutMs: 240_000,
    });
    const early = Date.now() - t0;
    expect(r.ok, JSON.stringify(r)).toBe(true);
    expect(r.exitCode, 'powershell.exe is still held alive').toBeNull();
    expect(pidAlive(r.excelPid), 'our Excel is still there right after the result').toBe(true);
    const closed = await r.closed!;
    expect(closed.excelKilled, 'the watchdog had to end our Excel').toBe(true);
    expect(pidAlive(r.excelPid)).toBe(false);
    expect(Date.now() - t0).toBeLessThan(early + 15_000);
    expect(early, 'the result did not wait for the stalled shutdown').toBeLessThan(25_000);
    expect(pidAlive(bystander.excelPid), 'the other Excel (PID ' + bystander.excelPid + ') must still be running').toBe(true);
    expect(fs.statSync(j.output).size).toBeGreaterThan(10_000);
    // clean up the bystander: only its own PID, through its own watchdog
    await bystander.reap!();
    const bc = await bystander.closed!;
    expect(bc.excelKilled).toBe(true);
    expect(pidAlive(bystander.excelPid)).toBe(false);
  }, 300_000);

  describe('through the real server (job API, private TEMP, upload opened in place)', () => {
    let sb: Sandbox;
    let srv: RunningServer;
    beforeAll(async () => {
      sb = makeSandbox('hard');
      srv = await startServer(sb, { EXCEL_RUNNER: RUNNER, EXCEL_DETECT: '', XLRUNNER_ALLOW_DEBUG: '' });
    }, 120_000);
    afterAll(async () => {
      await srv?.stop();
      if (sb) fs.rmSync(sb.dir, { recursive: true, force: true });
    });

    it('create -> input -> done -> result; Excel\'s private TEMP and the job folder are cleaned up; the result is a valid workbook', async () => {
      const ops = [
        { op: 'insertCols', sheet: SHEET, at: 3, count: 1, formatFrom: 'left' },
        { op: 'setCells', sheet: SHEET, at: 'C4', rows: [['メモ'], [{ f: '=1+1' }], [{ t: 's', v: '001', quote: true }]] },
      ];
      const created = await raw(srv.port, 'POST', '/api/jobs', { headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ version: 1, format: 'xlsx', ops, options: { recalc: true, clipboard: 'restoreEachOp' } }) });
      expect(created.status).toBe(200);
      const { id } = created.json();
      const t0 = Date.now();
      const put = await raw(srv.port, 'PUT', `/api/jobs/${id}/input`, { body: fs.readFileSync(SAMPLE) });
      expect(put.status).toBe(200);
      const v = await waitJob(srv.port, id, 240_000);
      const took = Date.now() - t0;
      expect(v.state, JSON.stringify(v)).toBe('done');
      expect(took, `job took ${took} ms`).toBeLessThan(30_000);
      const res = await raw(srv.port, 'GET', `/api/jobs/${id}/result`);
      expect(res.status).toBe(200);
      const cells = await rawCells(res.body, SHEET);
      expect(cells.get('C4')?.value).toBe('メモ');
      expect(cells.get('C5')?.formula).toBe('1+1');
      expect(cells.get('C6')?.value).toBe('001');
      expect(cells.get('C6')?.quotePrefix).toBe(true);
      // the upload was opened in place and Excel saved straight into output.xlsx: nothing but the output is left in the job folder
      const jobDirs = () => (fs.existsSync(sb.temp) ? fs.readdirSync(sb.temp).filter((d) => d.startsWith('job-')) : []);
      const only = (d: string) => fs.readdirSync(path.join(sb.temp, d));
      expect(await until(() => jobDirs().length === 1 && JSON.stringify(only(jobDirs()[0])) === JSON.stringify(['output.xlsx']), 10_000)).toBe(true);
      // Excel's own TEMP (Office diagnostics, ~80 MB per start on other machines) is deleted once that Excel has exited
      const xl = () => (fs.existsSync(sb.temp) ? fs.readdirSync(sb.temp).filter((d) => d.startsWith('xl-')) : []);
      expect(await until(() => xl().length === 0, 30_000), `left: ${xl().join(', ')}`).toBe(true);
      expect((await raw(srv.port, 'DELETE', `/api/jobs/${id}`)).status).toBe(200);
      expect(await until(() => jobDirs().length === 0, 10_000)).toBe(true);
      expect(srv.out()).not.toContain('失敗');
    }, 300_000);
  });
});
