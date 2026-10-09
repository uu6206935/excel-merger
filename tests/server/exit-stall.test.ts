/**
 * Round 5 performance bug: the UI saw the Excel本体 export finish its work in ~5 s ("保存しました") but the job only COMPLETED after ~64 s
 * (4 of 6 UI runs, 6 of 40 jobs in a soak test of the real runner). Root cause, measured: the runner's final {"type":"result"} line was out
 * after ~5 s every time, but powershell.exe could not EXIT for ~58 s - a hidden Excel that has been told to Quit() lives on until
 * powershell.exe releases it (the COM binder keeps references), then shuts down by talking to Microsoft's cloud services (8-9 open
 * HTTPS connections seen), and when that stalls it needs ~60 s; powershell.exe cannot exit before it (COM rundown). The server used to
 * wait for the process ('close') and then for the Excel PID. Now the result is delivered as soon as the final line has arrived and the
 * output is verified; the shutdown continues in the background under a watchdog that kills ONLY our own Excel PID.
 *
 * Everything here runs without Excel: a fake runner that keeps its process alive after the result (op debugWait + afterResult) and
 * a copy of node.exe named EXCEL.EXE that plays "our Excel" (and a second one that plays the user's Excel and must survive).
 */
import { afterAll, describe, expect, it } from 'vitest';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { spawn, type ChildProcess } from 'node:child_process';
import { isExcelPid, killExcelPid, pidAlive, processStartMs, runRunner } from '../../server/excel/run-job.mjs';
import { FAKE_RUNNER, makePackage, makeSandbox, raw, sleep, startServer, waitJob } from './helpers';

const WIN = process.platform === 'win32';
const cleanup: Array<() => void> = [];
afterAll(() => {
  for (const f of cleanup.reverse()) {
    try {
      f();
    } catch {
      /* best effort */
    }
  }
});

function scratch(label: string): string {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), `em-${label}-`));
  cleanup.push(() => fs.rmSync(dir, { recursive: true, force: true }));
  return dir;
}

function writeJob(dir: string, ops: unknown[], extra: Record<string, unknown> = {}) {
  const input = path.join(dir, 'in.xlsx');
  fs.writeFileSync(input, 'PK');
  const jobFile = path.join(dir, 'job.json');
  fs.writeFileSync(jobFile, JSON.stringify({ version: 1, format: 'xlsx', input, output: path.join(dir, 'out.xlsx'), ops, ...extra }));
  return jobFile;
}

/** A node.exe copy named EXCEL.EXE: tasklist shows it like an Excel process, nothing else about it is Excel. */
let fakeExeDir = '';
function fakeExcelExe(): string {
  if (!fakeExeDir) fakeExeDir = scratch('fakexl');
  const exe = path.join(fakeExeDir, 'EXCEL.EXE');
  if (!fs.existsSync(exe)) fs.copyFileSync(process.execPath, exe);
  return exe;
}
function startFakeExcel(): ChildProcess {
  const c = spawn(fakeExcelExe(), ['-e', 'setTimeout(() => {}, 120000)'], { stdio: 'ignore', windowsHide: true });
  cleanup.push(() => {
    try {
      c.kill();
    } catch {
      /* gone */
    }
  });
  return c;
}
async function until(check: () => boolean | Promise<boolean>, ms = 10_000): Promise<boolean> {
  const end = Date.now() + ms;
  while (Date.now() < end) {
    if (await check()) return true;
    await sleep(100);
  }
  return check();
}

describe('runRunner earlyResult (no Excel)', () => {
  it('delivers the result at once while the process is still running; the watchdog ends the process after the grace time', async () => {
    const dir = scratch('early');
    const jobFile = writeJob(dir, [{ op: 'debugWait', seconds: 120, afterResult: true }]);
    const t0 = Date.now();
    const events: string[] = [];
    const r = await runRunner({ runner: FAKE_RUNNER, jobFile, earlyResult: true, exitGraceMs: 800, onEvent: (e) => events.push(e.type) });
    const early = Date.now() - t0;
    expect(r.ok).toBe(true);
    expect(r.earlyResult).toBe(true);
    expect(r.exitCode, 'the process has not ended yet').toBeNull();
    expect(early, 'the result is not held back by the stalled process').toBeLessThan(5000);
    expect(events[events.length - 1]).toBe('result');
    const closed = await r.closed!;
    const total = Date.now() - t0;
    expect(total, 'the stalled runner was ended by the watchdog').toBeLessThan(12_000);
    expect(total).toBeGreaterThanOrEqual(early + 700); // not before the grace time
    expect(closed.wallMs).toBeGreaterThan(0);
  }, 30_000);

  it('without earlyResult the old behaviour remains: the promise waits for the process (CLI / tests)', async () => {
    const dir = scratch('late');
    const jobFile = writeJob(dir, [{ op: 'debugWait', seconds: 2.5, afterResult: true }]);
    const t0 = Date.now();
    const r = await runRunner({ runner: FAKE_RUNNER, jobFile });
    expect(Date.now() - t0).toBeGreaterThanOrEqual(2400);
    expect(r.ok).toBe(true);
    expect(r.earlyResult).toBeUndefined();
    expect(r.exitCode).toBe(0);
  }, 30_000);

  it('an error result is delivered early as well (the UI must not wait a minute to learn that a job failed)', async () => {
    const dir = scratch('err');
    const jobFile = writeJob(dir, [{ op: 'moveCols', sheet: 'FAIL', start: 1, count: 1, to: 3 }]);
    const r = await runRunner({ runner: FAKE_RUNNER, jobFile, earlyResult: true, exitGraceMs: 800 });
    expect(r.ok).toBe(false);
    expect(r.code).toBe('EXCEL_REFUSED');
    await r.closed;
  }, 30_000);
});

describe.skipIf(!WIN)('the watchdog kills ONLY our own Excel PID (fake EXCEL.EXE processes)', () => {
  it('reap after the grace time: our reported EXCEL.EXE dies, another EXCEL.EXE (the user\'s) is untouched', async () => {
    const dir = scratch('reap');
    const bystander = startFakeExcel(); // "the user's Excel"
    await sleep(500);
    expect(await isExcelPid(bystander.pid!)).toBe(true);
    const jobFile = writeJob(dir, [{ op: 'debugWait', seconds: 120, afterResult: true }], { fakeExcel: fakeExcelExe() });
    const t0 = Date.now();
    const r = await runRunner({ runner: FAKE_RUNNER, jobFile, earlyResult: true, exitGraceMs: 1000 });
    expect(Date.now() - t0).toBeLessThan(8000);
    expect(r.ok).toBe(true);
    const ours = r.excelPid;
    expect(ours).toBeGreaterThan(0);
    expect(ours).not.toBe(bystander.pid);
    expect(r.excelStartedMs).toBeGreaterThan(0);
    expect(pidAlive(ours), 'our Excel is still alive right after the result (it is shutting down in the background)').toBe(true);
    const closed = await r.closed!;
    expect(closed.excelKilled).toBe(true);
    expect(await until(() => !pidAlive(ours), 5000)).toBe(true);
    expect(pidAlive(bystander.pid!), 'the other EXCEL.EXE must not be touched').toBe(true);
  }, 60_000);

  it('a recycled PID is not killed: the start time must match what the runner reported', async () => {
    const x = startFakeExcel();
    await sleep(500);
    const pid = x.pid!;
    const start = await processStartMs(pid);
    expect(start).not.toBeNull();
    expect(Math.abs(start! - (Date.now() - process.uptime() * 0)), 'sane epoch ms').toBeLessThan(60_000);
    // a different start time (= a different process that got the same PID): refused, process untouched
    expect(await killExcelPid(pid, start! - 60_000)).toBe(false);
    expect(pidAlive(pid)).toBe(true);
    // the real start time: killed
    expect(await killExcelPid(pid, start!)).toBe(true);
    expect(await until(() => !pidAlive(pid), 5000)).toBe(true);
  }, 60_000);

  it('never kills a process that is not an EXCEL.EXE, whatever PID it is given', async () => {
    const c = spawn(process.execPath, ['-e', 'setTimeout(() => {}, 30000)'], { stdio: 'ignore', windowsHide: true });
    cleanup.push(() => c.kill());
    await sleep(300);
    expect(await isExcelPid(c.pid!)).toBe(false);
    expect(await killExcelPid(c.pid!)).toBe(false);
    expect(pidAlive(c.pid!)).toBe(true);
    expect(await killExcelPid(0)).toBe(false);
    expect(await killExcelPid(-5)).toBe(false);
    expect(await killExcelPid(1.5 as number)).toBe(false);
    c.kill();
  }, 30_000);
});

describe('the server delivers the job when the result line is out, not when the process has ended', () => {
  it('a runner that cannot exit after its result: done within seconds, result downloadable, leftovers cleaned later, shutdown quick', async () => {
    const sb = makeSandbox('stall');
    const srv = await startServer(sb, { EXCEL_EXIT_GRACE_MS: '1500' });
    try {
      const pkg = await makePackage();
      const { id } = (await raw(srv.port, 'POST', '/api/jobs', { body: JSON.stringify({ version: 1, format: 'xlsx', ops: [{ op: 'debugWait', seconds: 120, afterResult: true }] }) })).json();
      const t0 = Date.now();
      expect((await raw(srv.port, 'PUT', `/api/jobs/${id}/input`, { body: pkg })).status).toBe(200);
      const v = await waitJob(srv.port, id, 20_000);
      const took = Date.now() - t0;
      expect(v.state).toBe('done');
      expect(took, `done after ${took} ms (used to be ~64 s with a real Excel that stalls)`).toBeLessThan(6000);
      const res = await raw(srv.port, 'GET', `/api/jobs/${id}/result`);
      expect(res.status).toBe(200);
      expect(Buffer.compare(res.body, pkg)).toBe(0);
      // the next job is not held up by the lingering runner either
      const b = (await raw(srv.port, 'POST', '/api/jobs', { body: JSON.stringify({ version: 1, format: 'xlsx', ops: [] }) })).json().id;
      await raw(srv.port, 'PUT', `/api/jobs/${b}/input`, { body: pkg });
      expect((await waitJob(srv.port, b, 20_000)).state).toBe('done');
      // Excel's private TEMP folder (xl-*) exists while the runner lingers and is gone after it was ended
      const xl = () => (fs.existsSync(sb.temp) ? fs.readdirSync(sb.temp).filter((d) => d.startsWith('xl-')) : []);
      expect(await until(() => xl().length === 0, 15_000), `leftover: ${xl().join(', ')}`).toBe(true);
      expect((await raw(srv.port, 'DELETE', `/api/jobs/${id}`)).status).toBe(200);
      expect((await raw(srv.port, 'DELETE', `/api/jobs/${b}`)).status).toBe(200);
    } finally {
      await srv.stop();
      fs.rmSync(sb.dir, { recursive: true, force: true });
    }
  }, 90_000);

  it('the server ends lingering runners when it shuts down (nothing of ours stays behind)', async () => {
    const sb = makeSandbox('stall-bye');
    const srv = await startServer(sb, { EXCEL_EXIT_GRACE_MS: '60000', EXIT_GRACE_MS: '300' });
    const pkg = await makePackage();
    const { id } = (await raw(srv.port, 'POST', '/api/jobs', { body: JSON.stringify({ version: 1, format: 'xlsx', ops: [{ op: 'debugWait', seconds: 120, afterResult: true }] }) })).json();
    await raw(srv.port, 'PUT', `/api/jobs/${id}/input`, { body: pkg });
    expect((await waitJob(srv.port, id, 20_000)).state).toBe('done');
    const t0 = Date.now();
    await raw(srv.port, 'POST', '/api/bye', { body: '' });
    const code = await Promise.race([srv.exited, sleep(25_000).then(() => 'timeout' as const)]);
    expect(code).toBe(0);
    expect(Date.now() - t0, 'bye -> exit with a lingering runner').toBeLessThan(20_000);
    await sleep(300);
    expect(fs.existsSync(sb.temp) ? fs.readdirSync(sb.temp) : []).toEqual([]);
    fs.rmSync(sb.dir, { recursive: true, force: true });
  }, 60_000);

  it('an output that is not a complete package is not delivered', async () => {
    const sb = makeSandbox('corrupt');
    const srv = await startServer(sb);
    try {
      const { id } = (await raw(srv.port, 'POST', '/api/jobs', { body: JSON.stringify({ version: 1, format: 'xlsx', ops: [{ op: 'setCells', sheet: 'CORRUPT', at: 'A1', rows: [] }] }) })).json();
      await raw(srv.port, 'PUT', `/api/jobs/${id}/input`, { body: await makePackage() });
      const v = await waitJob(srv.port, id, 20_000);
      expect(v.state).toBe('error');
      expect(v.error).toMatch(/出力ファイル/);
    } finally {
      await srv.stop();
      fs.rmSync(sb.dir, { recursive: true, force: true });
    }
  }, 60_000);
});
