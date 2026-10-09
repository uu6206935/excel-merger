/**
 * Server lifetime: auto-shutdown (SSE + bye beacon), idle timeout, single instance, port fallback, --detach launcher.
 * Fixed ports 5190-5198 are reserved for these tests.
 */
import { afterAll, describe, expect, it } from 'vitest';
import { spawn } from 'node:child_process';
import fs from 'node:fs';
import http from 'node:http';
import { makeSandbox, openEvents, portAlive, raw, sleep, startServer, SERVER, type Sandbox } from './helpers';

const sandboxes: Sandbox[] = [];
const sandbox = (label: string) => {
  const s = makeSandbox(label);
  sandboxes.push(s);
  return s;
};
afterAll(() => {
  for (const s of sandboxes) fs.rmSync(s.dir, { recursive: true, force: true });
});

const within = <T>(p: Promise<T>, ms: number) =>
  Promise.race([p, new Promise<'timeout'>((r) => setTimeout(() => r('timeout'), ms))]);

describe('auto-shutdown', () => {
  it('exits a few seconds after the window closed (bye + SSE closed)', async () => {
    const s = await startServer(sandbox('bye'), { EXIT_GRACE_MS: '400', IDLE_EXIT_MS: '600000' });
    try {
      const ev = await openEvents(s.port);
      expect(ev.firstChunk).toContain('retry:');
      await sleep(200);
      expect(await portAlive(s.port)).toBe(true);
      const bye = await raw(s.port, 'POST', '/api/bye', { headers: { Origin: `http://localhost:${s.port}`, 'Sec-Fetch-Site': 'same-origin' }, body: '' });
      expect(bye.status).toBe(204);
      ev.close();
      const code = await within(s.exited, 5000);
      expect(code).toBe(0);
      expect(s.out()).toContain('ウィンドウが閉じられました');
    } finally {
      await s.stop();
    }
  });

  it('survives a reload (bye, then a new SSE connection right away)', async () => {
    const s = await startServer(sandbox('reload'), { EXIT_GRACE_MS: '500', IDLE_EXIT_MS: '600000' });
    try {
      const ev1 = await openEvents(s.port);
      await raw(s.port, 'POST', '/api/bye', { body: '' });
      ev1.close();
      const ev2 = await openEvents(s.port); // the reloaded page reconnects within milliseconds
      await sleep(1200);
      expect(await portAlive(s.port)).toBe(true);
      ev2.close();
    } finally {
      await s.stop();
    }
  });

  it('a connection lost WITHOUT bye (sleeping tab, crash) only arms the long idle timer', async () => {
    const s = await startServer(sandbox('nobye'), { EXIT_GRACE_MS: '300', IDLE_EXIT_MS: '2000' });
    try {
      const ev = await openEvents(s.port);
      await sleep(100);
      ev.close();
      await sleep(900);
      expect(await portAlive(s.port)).toBe(true);
      const code = await within(s.exited, 6000);
      expect(code).toBe(0);
    } finally {
      await s.stop();
    }
  });

  it('exits when no window ever connects (IDLE_EXIT_MS)', async () => {
    const s = await startServer(sandbox('never'), { IDLE_EXIT_MS: '800' });
    const code = await within(s.exited, 5000);
    expect(code).toBe(0);
  });

  it('keeps running while a window is connected', async () => {
    const s = await startServer(sandbox('conn'), { IDLE_EXIT_MS: '600', EXIT_GRACE_MS: '300' });
    try {
      const ev = await openEvents(s.port);
      await sleep(1500);
      expect(await portAlive(s.port)).toBe(true);
      ev.close();
    } finally {
      await s.stop();
    }
  });
});

describe('single instance / ports', () => {
  it('a second start on the same port just reuses the running server and exits 0', async () => {
    const a = await startServer(sandbox('inst-a'), { PORT: '5191' });
    try {
      expect(a.port).toBe(5191);
      const sb = sandbox('inst-b');
      const b = spawn(process.execPath, [SERVER], {
        env: { ...process.env, NO_OPEN: '1', PORT: '5191', DIST_DIR: sb.dist, OUTPUT_DIR: sb.output, EM_DATA_DIR: sb.data, EM_TEMP_DIR: sb.temp },
        stdio: ['ignore', 'pipe', 'pipe'],
        windowsHide: true,
      });
      let out = '';
      b.stdout!.on('data', (d) => (out += d));
      b.stderr!.on('data', (d) => (out += d));
      const code = await new Promise<number | null>((r) => b.on('exit', r));
      expect(code).toBe(0);
      expect(out).toContain('既に起動しています');
      expect(out).toContain('http://localhost:5191/');
      expect(await portAlive(5191)).toBe(true);
    } finally {
      await a.stop();
    }
  });

  it('skips a port used by another program', async () => {
    const foreign = http.createServer((_req, res) => {
      res.writeHead(404);
      res.end('not ours');
    });
    await new Promise<void>((r) => foreign.listen(5193, '127.0.0.1', () => r()));
    try {
      const s = await startServer(sandbox('foreign'), { PORT: '5193' });
      try {
        expect(s.port).toBe(5194);
        expect(s.out()).toContain('5193');
      } finally {
        await s.stop();
      }
    } finally {
      foreign.close();
    }
  });
});

describe('--detach launcher', () => {
  it('starts a hidden server, prints a Japanese line and exits; the server stops after bye', async () => {
    const sb = sandbox('detach');
    const p = spawn(process.execPath, [SERVER, '--detach'], {
      env: {
        ...process.env,
        NO_OPEN: '1',
        PORT: '5196',
        DIST_DIR: sb.dist,
        OUTPUT_DIR: sb.output,
        EM_DATA_DIR: sb.data,
        EM_TEMP_DIR: sb.temp,
        EXIT_GRACE_MS: '300',
        IDLE_EXIT_MS: '60000',
        EM_FROM_BAT: '1',
      },
      stdio: ['ignore', 'pipe', 'pipe'],
      windowsHide: true,
    });
    let out = '';
    p.stdout!.on('data', (d) => (out += d));
    p.stderr!.on('data', (d) => (out += d));
    const code = await new Promise<number | null>((r) => p.on('exit', r));
    expect(code, out).toBe(0);
    expect(out).toContain('起動しました: http://localhost:5196/');
    expect(await portAlive(5196)).toBe(true);
    const log = fs.readFileSync(`${sb.data}/logs/server.log`, 'utf8');
    expect(log).toContain('READY http://localhost:5196/');
    // a second --detach start reuses the running instance
    const p2 = spawn(process.execPath, [SERVER, '--detach'], {
      env: { ...process.env, NO_OPEN: '1', PORT: '5196', EM_DATA_DIR: sb.data, EM_FROM_BAT: '1' },
      stdio: ['ignore', 'pipe', 'pipe'],
      windowsHide: true,
    });
    let out2 = '';
    p2.stdout!.on('data', (d) => (out2 += d));
    expect(await new Promise<number | null>((r) => p2.on('exit', r))).toBe(0);
    expect(out2).toContain('既に起動しています');
    // window closed: bye without any SSE connection -> exits after the grace period
    await raw(5196, 'POST', '/api/bye', { body: '' });
    let alive = true;
    for (let i = 0; i < 30 && alive; i++) {
      await sleep(200);
      alive = await portAlive(5196);
    }
    expect(alive).toBe(false);
  });
});
