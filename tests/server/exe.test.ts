/**
 * The built ExcelMerger.exe (npm run build:exe; Node SEA): skipped when the exe has not been built.
 *   - starts (READY), answers /api/ping, serves the embedded UI and the embedded runner, exits after the window closed
 *   - start.bat's exe branch: `start.bat` -> `ExcelMerger.exe --detach` -> hidden server, Japanese line, exit 0
 * Uses random ports / 5191 and a temp data folder; every server it starts is stopped (only its own PID is ever killed).
 */
import { afterAll, describe, expect, it } from 'vitest';
import { spawn } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { ROOT, openEvents, portAlive, raw, sleep } from './helpers';

const EXE = path.join(ROOT, 'ExcelMerger.exe');
const WIN = process.platform === 'win32';
const HAVE_EXE = WIN && fs.existsSync(EXE);

const dirs: string[] = [];
const tempData = (label: string) => {
  const d = fs.mkdtempSync(path.join(os.tmpdir(), `em-exe-${label}-`));
  dirs.push(d);
  return d;
};
afterAll(() => {
  for (const d of dirs) fs.rmSync(d, { recursive: true, force: true });
});

const within = <T>(p: Promise<T>, ms: number) => Promise.race([p, new Promise<'timeout'>((r) => setTimeout(() => r('timeout'), ms))]);

describe.skipIf(!HAVE_EXE)('ExcelMerger.exe', () => {
  it('starts, answers /api/ping, serves the embedded UI, reports sea + Excel, and exits after the window closed', async () => {
    const data = tempData('direct');
    const env = { ...process.env, NO_OPEN: '1', PORT: '0', EM_DATA_DIR: data, EM_TEMP_DIR: path.join(data, 'temp'), EXIT_GRACE_MS: '400', IDLE_EXIT_MS: '600000' } as Record<string, string>;
    delete env.DIST_DIR;
    delete env.EXCEL_RUNNER;
    const proc = spawn(EXE, [], { env, stdio: ['ignore', 'pipe', 'pipe'], windowsHide: true });
    let out = '';
    proc.stdout!.on('data', (d) => (out += d.toString('utf8')));
    proc.stderr!.on('data', (d) => (out += d.toString('utf8')));
    const exited = new Promise<number | null>((r) => proc.on('exit', r));
    try {
      let port = 0;
      for (let i = 0; i < 300 && !port; i++) {
        const m = /READY http:\/\/localhost:(\d+)\//.exec(out);
        if (m) port = Number(m[1]);
        else await sleep(100);
      }
      expect(port, out).toBeGreaterThan(0);
      const ping = await raw(port, 'GET', '/api/ping');
      expect(ping.status).toBe(200);
      expect(ping.json()).toMatchObject({ app: 'excel-merger', pid: proc.pid });
      // the UI is the dist/index.html embedded into the exe (single file)
      const page = await raw(port, 'GET', '/');
      expect(page.status).toBe(200);
      expect(page.headers['content-type']).toContain('text/html');
      const html = page.text();
      expect(html).toContain('<div id="root"');
      expect(html.length).toBeGreaterThan(100_000);
      expect((await raw(port, 'GET', '/some/route')).status).toBe(200); // SPA fallback
      const status = (await raw(port, 'GET', '/api/status')).json();
      expect(status).toMatchObject({ app: 'excel-merger', sea: true });
      expect(typeof status.excel.installed).toBe('boolean');
      // a foreign Host header is refused (DNS rebinding guard)
      expect((await raw(port, 'GET', '/api/ping', { headers: { Host: 'evil.example:80' } })).status).toBe(403);
      // the embedded runner: a bad job is reported by the runner itself (PowerShell starts, Excel does not)
      const created = await raw(port, 'POST', '/api/jobs', { body: JSON.stringify({ version: 1, format: 'xlsx', ops: [{ op: 'moveCols', sheet: 'S', start: 1, count: 1, to: 3, method: 'clipFree' }] }) });
      expect(created.status).toBe(200);
      const { id } = created.json();
      const bad = Buffer.from('PK\u0003\u0004 not a workbook but a zip-looking blob with xl/workbook.xml');
      const put = await raw(port, 'PUT', `/api/jobs/${id}/input`, { body: bad });
      expect(put.status).toBe(400); // rejected by the package check before any Excel is started
      expect((await raw(port, 'DELETE', `/api/jobs/${id}`)).status).toBe(200);
      // window closed -> exit 0 shortly after
      const ev = await openEvents(port);
      await sleep(100);
      await raw(port, 'POST', '/api/bye', { headers: { Origin: `http://localhost:${port}`, 'Sec-Fetch-Site': 'same-origin' }, body: '' });
      ev.close();
      const code = await within(exited, 10_000);
      expect(code, out).toBe(0);
      expect(out).toContain('ウィンドウが閉じられました');
    } finally {
      if (proc.exitCode === null) proc.kill();
      await exited;
    }
  }, 120_000);

  it('start.bat (exe branch) starts a hidden ExcelMerger.exe, prints a Japanese line and exits 0; the server stops after bye', async () => {
    const PORT = 5191;
    expect(await portAlive(PORT), 'port 5191 is reserved for this test').toBe(false);
    const data = tempData('bat');
    const env = {
      ...process.env,
      NO_OPEN: '1',
      PORT: String(PORT),
      EM_DATA_DIR: data,
      EM_TEMP_DIR: path.join(data, 'temp'),
      EXIT_GRACE_MS: '400',
      IDLE_EXIT_MS: '120000',
    } as Record<string, string>;
    delete env.DIST_DIR;
    delete env.EXCEL_RUNNER;
    const p = spawn(process.env.ComSpec || 'cmd.exe', ['/d', '/c', path.join(ROOT, 'start.bat')], { cwd: ROOT, env, stdio: ['ignore', 'pipe', 'pipe'], windowsHide: true });
    let out = '';
    p.stdout!.on('data', (d) => (out += d.toString('utf8')));
    p.stderr!.on('data', (d) => (out += d.toString('utf8')));
    let pid = 0;
    try {
      const code = await within(new Promise<number | null>((r) => p.on('exit', r)), 60_000);
      expect(code, out).toBe(0);
      expect(out).toContain(`起動しました: http://localhost:${PORT}/`);
      const ping = await raw(PORT, 'GET', '/api/ping');
      expect(ping.status).toBe(200);
      pid = ping.json().pid;
      expect(pid).toBeGreaterThan(0);
      const log = fs.readFileSync(path.join(data, 'logs', 'server.log'), 'utf8');
      expect(log).toContain(`READY http://localhost:${PORT}/`);
      expect(log).toContain('(exe)');
      // the window closes: bye without any connection -> the hidden server exits after the grace period
      await raw(PORT, 'POST', '/api/bye', { body: '' });
      let alive = true;
      for (let i = 0; i < 50 && alive; i++) {
        await sleep(200);
        alive = await portAlive(PORT);
      }
      expect(alive, 'the hidden server stopped after bye').toBe(false);
      pid = 0;
    } finally {
      if (pid > 0) {
        try {
          process.kill(pid); // only the server this test started
        } catch {
          /* gone */
        }
      }
      if (p.exitCode === null) p.kill();
    }
  }, 120_000);
});
