/**
 * src/api.ts against the real server (fake runner): standalone fallback, status, runExcelJob ({bytes, warnings},
 * progress, errors, abort), saveToOutput, openFile, openOutputFolder.
 */
import { afterAll, beforeAll, describe, expect, it, vi } from 'vitest';
import fs from 'node:fs';
import { makePackage, makeSandbox, raw, startServer, type RunningServer, type Sandbox } from './helpers';
import { canFallBackToJs, getServerStatus, openFile, openOutputFolder, runExcelJob, saveToOutput, startHeartbeat } from '../../src/api';
import type { RunnerJob } from '../../src/core/contracts';

let sb: Sandbox;
let srv: RunningServer;
const realFetch = globalThis.fetch;

function asPage(base: string | null) {
  // the client uses relative URLs like a page served by the server would
  vi.stubGlobal('location', { protocol: base ? 'http:' : 'file:', origin: base ?? 'null' });
  vi.stubGlobal('fetch', (u: string | URL | Request, init?: RequestInit) => realFetch(typeof u === 'string' && base ? new URL(u, base) : u, init));
}

beforeAll(async () => {
  sb = makeSandbox('api');
  srv = await startServer(sb);
});
afterAll(async () => {
  vi.unstubAllGlobals();
  await srv?.stop();
  fs.rmSync(sb.dir, { recursive: true, force: true });
});

const job = (ops: Array<Record<string, unknown>> = []): RunnerJob => ({ version: 1, format: 'xlsx', ops, options: { recalc: true, clipboard: 'restoreEachOp' } });

describe('standalone (file://)', () => {
  it('getServerStatus() is null and the other calls fail with Japanese messages', async () => {
    asPage(null);
    expect(await getServerStatus()).toBeNull();
    expect(() => startHeartbeat()).not.toThrow();
    await expect(saveToOutput(new Uint8Array([1]), 'a.xlsx')).rejects.toThrow(/スタンドアロン/);
    await expect(runExcelJob(new Uint8Array([1]), job())).rejects.toThrow(/スタンドアロン/);
  });

  it('getServerStatus() is null when nothing answers or a foreign server answers', async () => {
    asPage('http://localhost:5198'); // nothing listens there
    expect(await getServerStatus()).toBeNull();
    await expect(saveToOutput(new Uint8Array([1]), 'a.xlsx')).rejects.toThrow(/ローカルサーバーに接続できません/);
  });
});

describe('with the server', () => {
  it('status', async () => {
    asPage(srv.base);
    const s = await getServerStatus();
    expect(s).toMatchObject({ app: 'excel-merger', excel: { installed: true } });
    expect(s!.outputDir.length).toBeGreaterThan(0);
  });

  it('runExcelJob returns the result bytes and reports progress', async () => {
    asPage(srv.base);
    const pkg = new Uint8Array(await makePackage());
    const progress: Array<{ ratio: number | null; message: string }> = [];
    const out = await runExcelJob(pkg, job([{ op: 'moveCols', sheet: 'S', start: 1, count: 1, to: 3 }]), (p) => progress.push(p));
    expect(Buffer.compare(Buffer.from(out.bytes), Buffer.from(pkg))).toBe(0);
    // the runner's warnings (streamed "warn" lines + the final result's list), deduplicated, as Japanese text
    expect(out.warnings).toEqual(['input=.xlsx', 'テスト警告']);
    expect(progress.length).toBeGreaterThan(1);
    expect(progress[progress.length - 1]).toEqual({ ratio: 1, message: '完了しました' });
    for (const p of progress) expect(typeof p.message).toBe('string');
    // the finished job was removed from the server
    await new Promise((r) => setTimeout(r, 200));
    expect((await raw(srv.port, 'GET', '/api/status')).json().jobs).toEqual({ running: 0, queued: 0 });
  });

  it('runExcelJob works on a subarray view (copies only what it must)', async () => {
    asPage(srv.base);
    const pkg = await makePackage();
    const big = new Uint8Array(pkg.length + 20);
    big.set(pkg, 10);
    const view = big.subarray(10, 10 + pkg.length);
    const out = await runExcelJob(view, job());
    expect(Buffer.compare(Buffer.from(out.bytes), pkg)).toBe(0);
    expect(out.warnings.every((w) => typeof w === 'string' && w.length > 0)).toBe(true);
  });

  it('runExcelJob rejects with the runner error message', async () => {
    asPage(srv.base);
    const err = await runExcelJob(new Uint8Array(await makePackage()), job([{ op: 'moveCols', sheet: 'FAIL', start: 1, count: 1, to: 3 }])).catch((e) => e);
    expect(err).toBeInstanceOf(Error);
    expect(err.message).toContain('列の移動');
    expect(err.code).toBe('EXCEL_REFUSED');
    const bad = await runExcelJob(new Uint8Array([1, 2, 3]), job()).catch((e) => e);
    expect(bad.message).toMatch(/Excel ブック/);
  });

  it('a MOVE_VERIFY_FAILED job rejects with the code, a message naming the built-in engine and fallbackToJs (the UI retries with the built-in engine)', async () => {
    asPage(srv.base);
    const pkg = new Uint8Array(await makePackage());
    for (const sheet of ['VERIFY', 'VERIFYOLD']) {
      const err = await runExcelJob(pkg, job([{ op: 'moveCols', sheet, start: 1, count: 1, to: 3 }])).catch((e) => e);
      expect(err).toBeInstanceOf(Error);
      expect(err.code).toBe('MOVE_VERIFY_FAILED');
      expect(err.fallbackToJs).toBe(true);
      expect(err.message).toMatch(/内蔵エンジン/);
      expect(canFallBackToJs(err)).toBe(true);
    }
    // other failures and user cancels do not ask for the fallback through the flag (the UI still falls back for every non-cancel error)
    const refused = await runExcelJob(pkg, job([{ op: 'moveCols', sheet: 'FAIL', start: 1, count: 1, to: 3 }])).catch((e) => e);
    expect(refused.fallbackToJs).toBeUndefined();
    expect(canFallBackToJs(refused)).toBe(false);
    expect(canFallBackToJs(Object.assign(new Error('x'), { code: 'CANCELLED', fallbackToJs: true }))).toBe(false);
    expect(canFallBackToJs(Object.assign(new Error('x'), { code: 'MOVE_VERIFY_FAILED' }))).toBe(true); // an older server without the flag
    expect(canFallBackToJs(null)).toBe(false);
  });

  it('runExcelJob can be aborted (the server job is cancelled)', async () => {
    asPage(srv.base);
    const ctrl = new AbortController();
    const p = runExcelJob(new Uint8Array(await makePackage()), job([{ op: 'debugWait', seconds: 60 }]), undefined, ctrl.signal);
    await new Promise((r) => setTimeout(r, 800));
    ctrl.abort();
    const err = await p.catch((e) => e);
    expect(err.name).toBe('AbortError');
    expect(err.message).toBe('中止しました');
    let jobs = { running: 1, queued: 0 };
    for (let i = 0; i < 50 && jobs.running; i++) {
      await new Promise((r) => setTimeout(r, 100));
      jobs = (await raw(srv.port, 'GET', '/api/status')).json().jobs;
    }
    expect(jobs).toEqual({ running: 0, queued: 0 });
  });

  it('saveToOutput / openFile / openOutputFolder', async () => {
    asPage(srv.base);
    const a = await saveToOutput(new Uint8Array([80, 75, 1]), '名簿_統合_20260930.xlsx');
    expect(a.name).toBe('名簿_統合_20260930.xlsx');
    const b = await saveToOutput(new Uint8Array([80, 75, 2]), '名簿_統合_20260930.xlsx');
    expect(b.name).toBe('名簿_統合_20260930 (2).xlsx');
    expect(fs.readFileSync(b.path)).toEqual(Buffer.from([80, 75, 2]));
    await expect(saveToOutput(new Uint8Array([1]), '../x.xlsx')).rejects.toThrow(/フォルダ/);
    await expect(openFile(a.path)).resolves.toBeUndefined();
    await expect(openFile('C:\\Windows\\win.ini')).rejects.toThrow(/出力フォルダ/);
    await expect(openOutputFolder()).resolves.toBeUndefined();
  });
});

describe('disk space (round 5)', () => {
  it('getServerStatus() passes the disk fields on; a full temp drive makes runExcelJob fail with DISK_LOW, which asks for the built-in engine', async () => {
    const sb2 = makeSandbox('api-disk');
    const s2 = await startServer(sb2, { EM_DISK_LOW_MB: '999999999', EM_MIN_FREE_MB: '999999999' });
    try {
      asPage(s2.base);
      const s = await getServerStatus();
      expect(s).not.toBeNull();
      if (s!.diskFreeMB === undefined) return; // Node without statfs: the server cannot tell
      expect(s!.diskLow).toBe(true);
      expect(typeof s!.diskFreeMB).toBe('number');
      expect(s!.diskMessage).toMatch(/空き容量が少なくなっています/);
      expect(s!.outputDiskLow).toBe(true);
      let err: unknown;
      try {
        await runExcelJob(new Uint8Array(await makePackage()), job());
      } catch (e) {
        err = e;
      }
      expect(err).toBeDefined();
      expect((err as { code?: string }).code).toBe('DISK_LOW');
      expect((err as Error).message).toMatch(/空き容量が足りません/);
      expect(canFallBackToJs(err)).toBe(true);
    } finally {
      asPage(srv.base);
      await s2.stop();
      fs.rmSync(sb2.dir, { recursive: true, force: true });
    }
  }, 60_000);

  it('a normal server reports diskLow as a boolean and no message unless it is low', async () => {
    asPage(srv.base);
    const s = await getServerStatus();
    expect(typeof s!.diskLow === 'boolean' || s!.diskLow === undefined).toBe(true);
    if (s!.diskLow === false) expect(s!.diskMessage).toBeUndefined();
  });
});
