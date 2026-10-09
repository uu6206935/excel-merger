/**
 * start.bat on a PC where the first choice cannot run (round 5 packaging dry run, Windows only, nothing is really started):
 *   ExcelMerger.exe blocked by Smart App Control / antivirus (simulated by a file that is not an executable) -> runtime\node.exe
 *   (simulated by a tool that fails) -> the installed Node.js + server\server.mjs (a stub that records its arguments) -> ... and the
 *   last resort prints a clear message. "start.bat debug" never hands over (the log of the first method must stay visible).
 * The standalone branch (Edge --app with dist\index.html) is not exercised here: it would open a real browser window.
 */
import { afterAll, describe, expect, it } from 'vitest';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { spawn } from 'node:child_process';

const WIN = process.platform === 'win32';
const ROOT = path.resolve(__dirname, '..', '..');
const SYS32 = path.join(process.env.SystemRoot || 'C:\\Windows', 'System32');
const made: string[] = [];
afterAll(() => {
  for (const d of made) fs.rmSync(d, { recursive: true, force: true });
});

interface Layout {
  fakeExe?: boolean;
  runtimeFails?: boolean;
  nodeStub?: boolean;
  dist?: boolean;
}

function makeFolder(l: Layout): string {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'em-ツール フォルダ-'));
  made.push(dir);
  fs.copyFileSync(path.join(ROOT, 'start.bat'), path.join(dir, 'start.bat'));
  if (l.fakeExe) fs.writeFileSync(path.join(dir, 'ExcelMerger.exe'), 'this is not an executable (a blocked / quarantined exe)');
  if (l.runtimeFails) {
    fs.mkdirSync(path.join(dir, 'runtime'));
    fs.mkdirSync(path.join(dir, 'app'));
    fs.copyFileSync(path.join(SYS32, 'where.exe'), path.join(dir, 'runtime', 'node.exe')); // "node.exe" that refuses the arguments: exit code 2
    fs.writeFileSync(path.join(dir, 'app', 'server.cjs'), '');
  }
  if (l.nodeStub) {
    fs.mkdirSync(path.join(dir, 'server'));
    fs.writeFileSync(path.join(dir, 'server', 'server.mjs'), "import fs from 'node:fs'; fs.writeFileSync(new URL('../marker-node.json', import.meta.url), JSON.stringify(process.argv.slice(2)));\n");
  }
  if (l.dist) {
    fs.mkdirSync(path.join(dir, 'dist'));
    fs.writeFileSync(path.join(dir, 'dist', 'index.html'), '<!doctype html><title>x</title>');
  }
  return dir;
}

function runBat(dir: string, args: string[] = [], pathEnv?: string): Promise<{ code: number | null; out: string }> {
  return new Promise((resolve) => {
    const env: NodeJS.ProcessEnv = { ...process.env, EM_DATA_DIR: path.join(dir, 'data'), NO_OPEN: '1' };
    if (pathEnv !== undefined) env.PATH = pathEnv;
    const p = spawn(process.env.ComSpec || 'cmd.exe', ['/d', '/c', path.join(dir, 'start.bat'), ...args], { cwd: dir, env, stdio: ['ignore', 'pipe', 'pipe'], windowsHide: true });
    let out = '';
    p.stdout!.on('data', (d) => (out += d.toString('latin1')));
    p.stderr!.on('data', (d) => (out += d.toString('latin1')));
    const timer = setTimeout(() => p.kill(), 90_000);
    p.on('exit', (code) => {
      clearTimeout(timer);
      resolve({ code, out });
    });
  });
}

const marker = (dir: string) => path.join(dir, 'marker-node.json');

describe.skipIf(!WIN)('start.bat hands over to the next method when one cannot be started', () => {
  it('blocked exe -> failing runtime node.exe -> installed Node.js: the server stub is started with --detach, exit 0', async () => {
    const dir = makeFolder({ fakeExe: true, runtimeFails: true, nodeStub: true, dist: true });
    const r = await runBat(dir);
    expect(r.code, r.out).toBe(0);
    expect(r.out).toContain('[WARN] ExcelMerger.exe could not be started');
    expect(r.out).toContain('[WARN] The bundled node.exe could not be started');
    expect(JSON.parse(fs.readFileSync(marker(dir), 'utf8'))).toEqual(['--detach']);
    expect(r.out).not.toContain('[ERROR]');
  }, 120_000);

  it('blocked exe, no runtime folder -> installed Node.js', async () => {
    const dir = makeFolder({ fakeExe: true, nodeStub: true, dist: true });
    const r = await runBat(dir);
    expect(r.code, r.out).toBe(0);
    expect(r.out).toContain('[WARN] ExcelMerger.exe could not be started');
    expect(r.out).not.toContain('bundled node.exe');
    expect(JSON.parse(fs.readFileSync(marker(dir), 'utf8'))).toEqual(['--detach']);
  }, 120_000);

  it('no exe, no runtime: Node.js is used directly (no warning)', async () => {
    const dir = makeFolder({ nodeStub: true, dist: true });
    const r = await runBat(dir);
    expect(r.code, r.out).toBe(0);
    expect(r.out).not.toContain('[WARN]');
    expect(JSON.parse(fs.readFileSync(marker(dir), 'utf8'))).toEqual(['--detach']);
  }, 120_000);

  it('nothing usable at all (blocked exe, no Node.js on the PATH, no dist): a clear message and exit code 1, no loop', async () => {
    const dir = makeFolder({ fakeExe: true });
    const r = await runBat(dir, [], `${SYS32};${path.dirname(SYS32)}`);
    expect(r.code, r.out).toBe(1);
    expect(r.out).toContain('[WARN] ExcelMerger.exe could not be started');
    expect(r.out).toContain('[ERROR] ExcelMerger.exe, Node.js and dist\\index.html are all missing or not usable');
    expect(r.out).toContain('runtime');
  }, 120_000);

  it('start.bat debug keeps the first method visible: no hand-over, the error text and exit code 1', async () => {
    const dir = makeFolder({ fakeExe: true, nodeStub: true, dist: true });
    const r = await runBat(dir, ['debug']);
    expect(r.code, r.out).toBe(1);
    expect(r.out).toContain('[ERROR] Excel Merger could not start');
    expect(fs.existsSync(marker(dir))).toBe(false);
  }, 120_000);

  it('the launch checks compare %ERRORLEVEL% with 0: a killed / crashed process exits with a NEGATIVE code (-1073741515 ...) that "if errorlevel 1" would not see', async () => {
    const dir = makeFolder({ nodeStub: true, dist: true });
    fs.mkdirSync(path.join(dir, 'runtime'));
    fs.mkdirSync(path.join(dir, 'app'));
    fs.writeFileSync(path.join(dir, 'app', 'server.cjs'), '');
    fs.copyFileSync(path.join(SYS32, 'where.exe'), path.join(dir, 'runtime', 'node.exe')); // exits with 2 here; a negative code takes the same branch
    const src = fs.readFileSync(path.join(ROOT, 'start.bat'), 'latin1');
    expect(src).toContain('if not "%ERRORLEVEL%"=="0" goto :exefailed');
    expect(src).toContain('if not "%ERRORLEVEL%"=="0" goto :runtimefailed');
    expect(src).toContain('if not "%ERRORLEVEL%"=="0" goto :nodefailed');
    expect(/if errorlevel 1 goto :(exe|runtime|node)failed/.test(src)).toBe(false);
    const r = await runBat(dir);
    expect(r.code, r.out).toBe(0);
    expect(r.out).toContain('[WARN] The bundled node.exe could not be started');
    expect(JSON.parse(fs.readFileSync(marker(dir), 'utf8'))).toEqual(['--detach']);
  }, 120_000);
});
