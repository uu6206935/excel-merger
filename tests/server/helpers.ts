/** Helpers for the server tests: spawn server/server.mjs on a test port with a temp fake dist. */
import { spawn, type ChildProcess } from 'node:child_process';
import fs from 'node:fs';
import http from 'node:http';
import os from 'node:os';
import path from 'node:path';
import JSZip from 'jszip';

export const ROOT = path.resolve(__dirname, '..', '..');
export const SERVER = path.join(ROOT, 'server', 'server.mjs');
export const FAKE_RUNNER = path.join(__dirname, 'fixtures', 'fake-runner.mjs');

export interface Sandbox {
  dir: string;
  dist: string;
  output: string;
  data: string;
  temp: string;
}

/** Temp folder with a Japanese name + space (the real app lives in such folders too). */
export function makeSandbox(label = 'srv'): Sandbox {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), `em-${label}-テスト フォルダ-`));
  const dist = path.join(dir, 'dist');
  fs.mkdirSync(dist, { recursive: true });
  fs.writeFileSync(path.join(dist, 'index.html'), '<!DOCTYPE html><html lang="ja"><head><meta charset="utf-8"><title>Excel 統合ツール</title></head><body>fake dist</body></html>');
  fs.writeFileSync(path.join(dist, 'extra.js'), 'console.log(1)');
  return { dir, dist, output: path.join(dir, 'output'), data: path.join(dir, 'data'), temp: path.join(dir, 'temp') };
}

export interface RunningServer {
  proc: ChildProcess;
  port: number;
  base: string;
  out: () => string;
  exited: Promise<number | null>;
  stop: () => Promise<void>;
}

export function startServer(sb: Sandbox, env: Record<string, string> = {}, args: string[] = []): Promise<RunningServer> {
  return new Promise((resolve, reject) => {
    const proc = spawn(process.execPath, [SERVER, ...args], {
      env: {
        ...process.env,
        NO_OPEN: '1',
        PORT: '0',
        DIST_DIR: sb.dist,
        OUTPUT_DIR: sb.output,
        EM_DATA_DIR: sb.data,
        EM_TEMP_DIR: sb.temp,
        EXCEL_RUNNER: FAKE_RUNNER,
        EXCEL_DETECT: 'installed',
        XLRUNNER_ALLOW_DEBUG: '1',
        IDLE_EXIT_MS: '0',
        ...env,
      },
      stdio: ['ignore', 'pipe', 'pipe'],
      windowsHide: true,
    });
    let buf = '';
    let settled = false;
    const exited = new Promise<number | null>((r) => proc.on('exit', (code) => r(code)));
    const timer = setTimeout(() => {
      if (settled) return;
      settled = true;
      proc.kill();
      reject(new Error(`server did not start:\n${buf}`));
    }, 20000);
    const onData = (d: Buffer) => {
      buf += d.toString('utf8');
      const m = /READY http:\/\/localhost:(\d+)\//.exec(buf);
      if (m && !settled) {
        settled = true;
        clearTimeout(timer);
        const port = Number(m[1]);
        resolve({
          proc,
          port,
          base: `http://localhost:${port}`,
          out: () => buf,
          exited,
          stop: async () => {
            if (proc.exitCode === null) proc.kill();
            await exited;
          },
        });
      }
    };
    proc.stdout!.on('data', onData);
    proc.stderr!.on('data', onData);
    proc.on('exit', (code) => {
      if (!settled) {
        settled = true;
        clearTimeout(timer);
        reject(new Error(`server exited (${code}) before READY:\n${buf}`));
      }
    });
  });
}

export interface RawResponse {
  status: number;
  headers: http.IncomingHttpHeaders;
  body: Buffer;
  json: () => any;
  text: () => string;
}

/** Raw HTTP request (unlike fetch, any Host / Origin header can be sent). */
export function raw(
  port: number,
  method: string,
  pathName: string,
  opts: { headers?: Record<string, string | number>; body?: Buffer | string } = {},
): Promise<RawResponse> {
  return new Promise((resolve, reject) => {
    // agent: false = one connection per request: a server that answers early (403 / 413) and drops the rest of a body may close the socket, and a
    // reused keep-alive socket would then fail with ECONNRESET for the NEXT request (a test artefact, not a server problem)
    const req = http.request({ host: '127.0.0.1', port, method, path: pathName, agent: false, headers: { Host: `localhost:${port}`, ...(opts.headers || {}) } }, (res) => {
      const chunks: Buffer[] = [];
      res.on('data', (c: Buffer) => chunks.push(c));
      res.on('end', () => {
        const body = Buffer.concat(chunks);
        resolve({ status: res.statusCode || 0, headers: res.headers, body, json: () => JSON.parse(body.toString('utf8')), text: () => body.toString('utf8') });
      });
    });
    req.on('error', reject);
    if (opts.body !== undefined) req.write(opts.body);
    req.end();
  });
}

/** Minimal workbook package the server accepts (zip with xl/workbook.xml + [Content_Types].xml). */
export async function makePackage(kind: 'xlsx' | 'xlsm' = 'xlsx'): Promise<Buffer> {
  const zip = new JSZip();
  const main = kind === 'xlsm' ? 'application/vnd.ms-excel.sheet.macroEnabled.main+xml' : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml';
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/xl/workbook.xml" ContentType="${main}"/></Types>`,
  );
  zip.file('xl/workbook.xml', '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheets/></workbook>');
  zip.file('xl/worksheets/sheet1.xml', '<worksheet/>'.repeat(50));
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}

export const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms));

/** Polls GET /api/jobs/:id until the state is final. */
export async function waitJob(port: number, id: string, timeoutMs = 20000): Promise<any> {
  const end = Date.now() + timeoutMs;
  for (;;) {
    const r = await raw(port, 'GET', `/api/jobs/${id}`);
    const v = r.json();
    if (v.state === 'done' || v.state === 'error') return v;
    if (Date.now() > end) throw new Error(`job ${id} not finished: ${JSON.stringify(v)}`);
    await sleep(100);
  }
}

/** Opens an SSE connection to /api/events; returns a closer. */
export function openEvents(port: number): Promise<{ close: () => void; firstChunk: string }> {
  return new Promise((resolve, reject) => {
    const req = http.get({ host: '127.0.0.1', port, path: '/api/events', headers: { Host: `localhost:${port}` } }, (res) => {
      res.once('data', (c: Buffer) => resolve({ close: () => req.destroy(), firstChunk: c.toString('utf8') }));
    });
    req.on('error', (e) => reject(e));
  });
}

export async function portAlive(port: number): Promise<boolean> {
  try {
    const r = await raw(port, 'GET', '/api/ping');
    return r.status === 200;
  } catch {
    return false;
  }
}
