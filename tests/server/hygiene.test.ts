/**
 * Disk hygiene and error hand-over of the job server (round 3), fake runner (no Excel needed):
 *   - the temp job folder (%TEMP%\excel-merger\job-<id>-<pid>, usually on the SYSTEM drive) is deleted after EVERY job: a finished
 *     job keeps only its output for the download (input copy, job file and Excel's work folder go at once), a failed / crashed /
 *     cancelled job loses the whole folder, DELETE removes the rest;
 *   - stale job folders (server gone for > 1 hour, or ANY job folder older than 1 day) are removed at server start, everything else
 *     in the temp root (fresh folders, folders of other programs) is left alone;
 *   - a MOVE_VERIFY_FAILED job reaches the client as errorCode + a Japanese message that names the built-in engine + fallbackToJs
 *     (the UI retries with the built-in engine; src/state treats every non-cancel error that way).
 */
import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import fs from 'node:fs';
import path from 'node:path';
import { spawn } from 'node:child_process';
import { makePackage, makeSandbox, openEvents, raw, sleep, startServer, waitJob, type RunningServer, type Sandbox } from './helpers';

let sb: Sandbox;
let srv: RunningServer;

beforeAll(async () => {
  sb = makeSandbox('hyg');
  srv = await startServer(sb);
});
afterAll(async () => {
  await srv?.stop();
  fs.rmSync(sb.dir, { recursive: true, force: true });
});

const createJob = (job: unknown) => raw(srv.port, 'POST', '/api/jobs', { headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(job) });
const baseJob = (ops: unknown[] = []) => ({ version: 1, format: 'xlsx', ops, options: { recalc: true, clipboard: 'restoreEachOp' } });
const jobDirs = (id: string, temp = sb.temp) => (fs.existsSync(temp) ? fs.readdirSync(temp).filter((d) => d.startsWith(`job-${id.slice(0, 8)}-`)) : []);

/** Polls until `check()` holds (the cleanup runs asynchronously right after the state changes). */
async function until(check: () => boolean, ms = 8000): Promise<boolean> {
  const end = Date.now() + ms;
  while (Date.now() < end) {
    if (check()) return true;
    await sleep(100);
  }
  return check();
}

async function submit(ops: unknown[], pkg?: Buffer): Promise<string> {
  const { id } = (await createJob(baseJob(ops))).json();
  const put = await raw(srv.port, 'PUT', `/api/jobs/${id}/input`, { body: pkg ?? (await makePackage()) });
  expect(put.status).toBe(200);
  return id;
}

describe('temp job folders are deleted after every job', () => {
  it('a finished job keeps only its output; DELETE removes the folder', async () => {
    const id = await submit([{ op: 'moveCols', sheet: 'S', start: 2, count: 1, to: 5, method: 'clipFree' }]);
    expect((await waitJob(srv.port, id)).state).toBe('done');
    const [dir] = jobDirs(id);
    expect(dir, 'the job folder exists while the result waits for its download').toBeTruthy();
    const full = path.join(sb.temp, dir);
    expect(await until(() => JSON.stringify(fs.readdirSync(full)) === JSON.stringify(['output.xlsx']))).toBe(true);
    // the result is still there
    const res = await raw(srv.port, 'GET', `/api/jobs/${id}/result`);
    expect(res.status).toBe(200);
    expect(res.body.length).toBeGreaterThan(100);
    expect((await raw(srv.port, 'DELETE', `/api/jobs/${id}`)).status).toBe(200);
    expect(await until(() => jobDirs(id).length === 0)).toBe(true);
  });

  it('failed jobs (runner error, crash, no output file) leave no folder', async () => {
    const ids = [
      await submit([{ op: 'moveCols', sheet: 'FAIL', start: 1, count: 1, to: 3 }]),
      await submit([{ op: 'setCells', sheet: 'CRASH', at: 'A1', rows: [] }]),
      await submit([{ op: 'setCells', sheet: 'NOOUT', at: 'A1', rows: [] }]),
      await submit([{ op: 'moveCols', sheet: 'VERIFY', start: 1, count: 1, to: 3 }]),
    ];
    for (const id of ids) {
      expect((await waitJob(srv.port, id)).state, id).toBe('error');
      expect(await until(() => jobDirs(id).length === 0), `folder of ${id}`).toBe(true);
    }
  });

  it('a cancelled running job and a cancelled queued job leave no folder', async () => {
    const running = await submit([{ op: 'debugWait', seconds: 60 }]);
    const queued = await submit([]);
    await sleep(600);
    expect((await raw(srv.port, 'GET', `/api/jobs/${running}`)).json().state).toBe('running');
    expect(jobDirs(running).length).toBe(1);
    expect(jobDirs(queued).length).toBe(1);
    expect((await raw(srv.port, 'DELETE', `/api/jobs/${queued}`)).status).toBe(200);
    expect(await until(() => jobDirs(queued).length === 0)).toBe(true);
    expect((await raw(srv.port, 'DELETE', `/api/jobs/${running}`)).status).toBe(200);
    expect(await until(() => jobDirs(running).length === 0)).toBe(true);
  });
});

describe('stale job folders are removed at server start', () => {
  /** a pid that is gone: a short-lived child of this test */
  async function deadPid(): Promise<number> {
    const child = spawn(process.execPath, ['-e', '0'], { stdio: 'ignore', windowsHide: true });
    const pid = child.pid!;
    await new Promise((r) => child.on('exit', r));
    return pid;
  }
  const age = (p: string, ms: number) => {
    const t = new Date(Date.now() - ms);
    fs.utimesSync(p, t, t);
  };
  const HOUR = 3600_000;

  it('removes folders of a dead server after 1 hour and ANY job folder after 1 day, keeps everything else', async () => {
    const sb2 = makeSandbox('sweep');
    try {
      fs.mkdirSync(sb2.temp, { recursive: true });
      const dead = await deadPid();
      const alive = process.pid; // this test process: "alive"
      const mk = (name: string, ageMs: number, file = 'input.xlsx') => {
        const d = path.join(sb2.temp, name);
        fs.mkdirSync(path.join(d, 'work'), { recursive: true });
        fs.writeFileSync(path.join(d, file), Buffer.alloc(5000, 1));
        fs.writeFileSync(path.join(d, 'work', 'work.xlsx'), Buffer.alloc(5000, 2));
        age(d, ageMs);
        return d;
      };
      const staleDead = mk(`job-aaaaaaaa-${dead}`, 2 * HOUR);
      const staleAlive = mk(`job-bbbbbbbb-${alive}`, 25 * HOUR); // a live pid proves nothing after a day (pid reuse)
      const veryOld = mk('job-cccccccc-1', 30 * HOUR);
      const freshDead = mk(`job-dddddddd-${dead}`, 10 * 60_000); // server gone but only 10 minutes ago: kept for now
      const freshAlive = mk(`job-eeeeeeee-${alive}`, 2 * HOUR); // its server is alive: kept (a long job of a parallel instance)
      const foreign = path.join(sb2.temp, 'somebody-elses-folder'); // not a job folder: never touched
      fs.mkdirSync(foreign);
      fs.writeFileSync(path.join(foreign, 'keep.txt'), 'x');
      age(foreign, 40 * HOUR);
      const notDir = path.join(sb2.temp, 'job-ffffffff-1'); // a FILE with a job-like name
      fs.writeFileSync(notDir, 'x');
      age(notDir, 40 * HOUR);

      const s2 = await startServer(sb2);
      try {
        expect(await until(() => !fs.existsSync(staleDead) && !fs.existsSync(staleAlive) && !fs.existsSync(veryOld), 15000)).toBe(true);
        expect(fs.existsSync(freshDead)).toBe(true);
        expect(fs.existsSync(freshAlive)).toBe(true);
        expect(fs.existsSync(path.join(foreign, 'keep.txt'))).toBe(true);
        expect(fs.existsSync(notDir)).toBe(true);
        expect(s2.out()).toMatch(/古い一時フォルダを 3 件削除しました/);
      } finally {
        await s2.stop();
      }
    } finally {
      fs.rmSync(sb2.dir, { recursive: true, force: true });
    }
  });

  it("also sweeps Excel's private TEMP folders (xl-*: Office diagnostics, ~80 MB per Excel start) of a dead server, like job folders", async () => {
    const sb2 = makeSandbox('sweep-xl');
    try {
      fs.mkdirSync(sb2.temp, { recursive: true });
      const dead = await deadPid();
      const mkXl = (name: string, ageMs: number) => {
        const d = path.join(sb2.temp, name);
        fs.mkdirSync(path.join(d, 'Diagnostics', 'EXCEL'), { recursive: true });
        fs.writeFileSync(path.join(d, 'Diagnostics', 'EXCEL', 'Primary1790000000000000000_X.log'), Buffer.alloc(200_000, 3));
        age(d, ageMs);
        return d;
      };
      const staleDead = mkXl(`xl-aaaaaaaa-${dead}`, 2 * HOUR);
      const veryOld = mkXl('xl-bbbbbbbb-1', 30 * HOUR);
      const freshDead = mkXl(`xl-cccccccc-${dead}`, 10 * 60_000);
      const aliveOwner = mkXl(`xl-dddddddd-${process.pid}`, 2 * HOUR);
      const foreign = mkXl('xl-notahexid-1', 40 * HOUR); // the name must match xl-<8 hex>-<pid> exactly
      const s2 = await startServer(sb2);
      try {
        expect(await until(() => !fs.existsSync(staleDead) && !fs.existsSync(veryOld), 15000)).toBe(true);
        expect(fs.existsSync(freshDead)).toBe(true);
        expect(fs.existsSync(aliveOwner)).toBe(true);
        expect(fs.existsSync(foreign)).toBe(true);
        expect(s2.out()).toMatch(/古い一時フォルダを 2 件削除しました/);
      } finally {
        await s2.stop();
      }
    } finally {
      fs.rmSync(sb2.dir, { recursive: true, force: true });
    }
  });

  it('the server removes its own job folders when the window closes (graceful shutdown)', async () => {
    const sb2 = makeSandbox('shutdown');
    const s2 = await startServer(sb2, { EXIT_GRACE_MS: '300', IDLE_EXIT_MS: '600000' });
    try {
      const ev = await openEvents(s2.port);
      const { id } = (await raw(s2.port, 'POST', '/api/jobs', { body: JSON.stringify(baseJob([])) })).json();
      await raw(s2.port, 'PUT', `/api/jobs/${id}/input`, { body: await makePackage() });
      expect((await waitJob(s2.port, id)).state).toBe('done');
      expect(jobDirs(id, sb2.temp).length).toBe(1); // the finished job still waits for its download
      await raw(s2.port, 'POST', '/api/bye', { body: '' });
      ev.close();
      const code = await Promise.race([s2.exited, sleep(8000).then(() => 'timeout' as const)]);
      expect(code).toBe(0);
      expect(jobDirs(id, sb2.temp)).toEqual([]);
      expect(fs.readdirSync(sb2.temp)).toEqual([]);
    } finally {
      await s2.stop();
      fs.rmSync(sb2.dir, { recursive: true, force: true });
    }
  });
});

describe('MOVE_VERIFY_FAILED reaches the client with the hint to use the built-in engine', () => {
  it('errorCode, Japanese message naming the built-in engine, fallbackToJs', async () => {
    const id = await submit([{ op: 'moveCols', sheet: 'VERIFY', start: 1, count: 1, to: 3 }]);
    const v = await waitJob(srv.port, id);
    expect(v.state).toBe('error');
    expect(v.errorCode).toBe('MOVE_VERIFY_FAILED');
    expect(v.fallbackToJs).toBe(true);
    expect(v.error).toMatch(/内蔵エンジン/);
    expect(v.error).toMatch(/テーブル/);
  });

  it('an older runner message without the hint gets it from the server', async () => {
    const id = await submit([{ op: 'moveCols', sheet: 'VERIFYOLD', start: 1, count: 1, to: 3 }]);
    const v = await waitJob(srv.port, id);
    expect(v.errorCode).toBe('MOVE_VERIFY_FAILED');
    expect(v.fallbackToJs).toBe(true);
    expect(v.error).toMatch(/内蔵エンジン/);
    expect(v.error).toMatch(/保存していません/);
  });

  it('other errors do not ask for the fallback', async () => {
    const id = await submit([{ op: 'moveCols', sheet: 'FAIL', start: 1, count: 1, to: 3 }]);
    const v = await waitJob(srv.port, id);
    expect(v.errorCode).toBe('EXCEL_REFUSED');
    expect(v.fallbackToJs).toBeUndefined();
  });
});
