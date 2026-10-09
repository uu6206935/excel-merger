/**
 * server/server.mjs: static serving, guards (Host / Origin / Sec-Fetch-Site), save / open path rules,
 * jobs API with a fake runner (EXCEL_RUNNER override). The server runs with NO_OPEN=1 on a random port.
 */
import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import fs from 'node:fs';
import path from 'node:path';
import { makePackage, makeSandbox, raw, startServer, waitJob, sleep, type RunningServer, type Sandbox } from './helpers';

let sb: Sandbox;
let srv: RunningServer;
let port: number;

beforeAll(async () => {
  sb = makeSandbox('main');
  srv = await startServer(sb);
  port = srv.port;
});
afterAll(async () => {
  await srv?.stop();
  fs.rmSync(sb.dir, { recursive: true, force: true });
});

const saveReq = (name: string, body: Buffer | string = 'PKdata', headers: Record<string, string> = {}) =>
  raw(port, 'POST', '/api/save', { headers: { 'x-file-name': encodeURIComponent(name), ...headers }, body });

describe('basic endpoints', () => {
  it('ping / status', async () => {
    const p = await raw(port, 'GET', '/api/ping');
    expect(p.status).toBe(200);
    expect(p.json()).toMatchObject({ app: 'excel-merger' });
    expect(typeof p.json().version).toBe('string');
    const s = (await raw(port, 'GET', '/api/status')).json();
    expect(s.app).toBe('excel-merger');
    expect(s.excel).toMatchObject({ installed: true });
    expect(path.resolve(s.outputDir)).toBe(path.resolve(sb.output));
  });

  it('serves index.html for / and SPA routes, files from dist, 404 for missing assets / APIs', async () => {
    const root = await raw(port, 'GET', '/');
    expect(root.status).toBe(200);
    expect(root.headers['content-type']).toContain('text/html');
    expect(root.text()).toContain('fake dist');
    expect((await raw(port, 'GET', '/some/route')).text()).toContain('fake dist');
    const js = await raw(port, 'GET', '/extra.js');
    expect(js.status).toBe(200);
    expect(js.headers['content-type']).toContain('javascript');
    expect((await raw(port, 'GET', '/missing.js')).status).toBe(404);
    expect((await raw(port, 'GET', '/%2e%2e/%2e%2e/package.json')).status).toBe(404);
    const api = await raw(port, 'GET', '/api/nothing');
    expect(api.status).toBe(404);
    expect(api.json().error).toMatch(/API/);
    const head = await raw(port, 'HEAD', '/');
    expect(head.status).toBe(200);
    expect(head.body.length).toBe(0);
  });
});

describe('guards', () => {
  it('refuses foreign Host headers (DNS rebinding)', async () => {
    for (const host of ['evil.example', `evil.example:${port}`, 'localhost.evil.example']) {
      const r = await raw(port, 'GET', '/api/ping', { headers: { Host: host } });
      expect(r.status).toBe(403);
      expect(r.json().error).toContain('localhost');
    }
    for (const host of [`localhost:${port}`, `127.0.0.1:${port}`, `[::1]:${port}`, 'LOCALHOST']) {
      expect((await raw(port, 'GET', '/api/ping', { headers: { Host: host } })).status).toBe(200);
    }
    expect((await raw(port, 'GET', '/', { headers: { Host: 'evil.example' } })).status).toBe(403);
  });

  it('refuses cross-site writes (Origin / Sec-Fetch-Site)', async () => {
    const bad: Array<Record<string, string>> = [
      { Origin: 'http://evil.example' },
      { Origin: `http://evil.example:${port}` },
      { Origin: `http://localhost:${port + 1}` },
      { Origin: 'null' },
      { Origin: `https://localhost:${port}` },
      { 'Sec-Fetch-Site': 'cross-site' },
      { 'Sec-Fetch-Site': 'same-site' },
    ];
    for (const h of bad) {
      const r = await saveReq('x.xlsx', 'PK', h);
      expect(r.status, JSON.stringify(h)).toBe(403);
      expect(r.json().error).toContain('拒否');
    }
    const del = await raw(port, 'DELETE', '/api/jobs/00000000-0000-0000-0000-000000000000', { headers: { Origin: 'http://evil.example' } });
    expect(del.status).toBe(403);
    const ok = await saveReq('guard-ok.xlsx', 'PK', { Origin: `http://localhost:${port}`, 'Sec-Fetch-Site': 'same-origin' });
    expect(ok.status).toBe(200);
    const ok2 = await saveReq('guard-ok2.xlsx', 'PK', { Origin: `http://127.0.0.1:${port}` });
    expect(ok2.status).toBe(200);
  });

  it('refuses OPTIONS (no CORS preflight is ever granted)', async () => {
    const r = await raw(port, 'OPTIONS', '/api/save', { headers: { 'Access-Control-Request-Method': 'POST' } });
    expect(r.status).toBe(405);
    expect(r.headers['access-control-allow-origin']).toBeUndefined();
  });
});

describe('POST /api/save', () => {
  it('saves into the output folder and never overwrites', async () => {
    const a = await saveReq('統合 結果.xlsx', 'PK-first');
    expect(a.status).toBe(200);
    const ja = a.json();
    expect(ja.name).toBe('統合 結果.xlsx');
    expect(path.dirname(ja.path)).toBe(path.resolve(sb.output));
    expect(fs.readFileSync(ja.path, 'utf8')).toBe('PK-first');
    const b = (await saveReq('統合 結果.xlsx', 'PK-second')).json();
    expect(b.name).toBe('統合 結果 (2).xlsx');
    const c = (await saveReq('統合 結果.xlsx', 'PK-third')).json();
    expect(c.name).toBe('統合 結果 (3).xlsx');
    expect(fs.readFileSync(ja.path, 'utf8')).toBe('PK-first');
    expect(fs.readFileSync(b.path, 'utf8')).toBe('PK-second');
  });

  it('rejects folders, traversal, other extensions, empty bodies; cleans unsafe characters', async () => {
    for (const name of ['../evil.xlsx', '..\\evil.xlsx', 'sub/evil.xlsx', 'C:\\x.xlsx', 'c:evil.xlsx']) {
      const r = await saveReq(name);
      expect(r.status, name).toBe(400);
      expect(r.json().error).toMatch(/フォルダ/);
    }
    for (const name of ['a.txt', 'a.exe', 'a.xlsx.bat', 'noext']) {
      const r = await saveReq(name);
      expect(r.status, name).toBe(400);
      expect(r.json().error).toContain('.xlsx');
    }
    const missing = await raw(port, 'POST', '/api/save', { body: 'PK' });
    expect(missing.status).toBe(400);
    expect(missing.json().error).toContain('ファイル名');
    expect((await saveReq('empty.xlsx', '')).status).toBe(400);
    expect((await saveReq('report:b?c*.xlsx')).json().name).toBe('report_b_c_.xlsx');
    expect((await saveReq('CON.xlsx')).json().name).toBe('_CON.xlsx');
    expect((await saveReq('macro.XLSM')).json().name).toBe('macro.xlsm');
    for (const f of fs.readdirSync(sb.output)) expect(isInside(sb.output, path.join(sb.output, f))).toBe(true);
  });

  it('accepts a raw UTF-8 header (mojibake repair) as well as the URI-encoded one', async () => {
    const r = await raw(port, 'POST', '/api/save', { headers: { 'x-file-name': Buffer.from('生データ.xlsx', 'utf8').toString('latin1') }, body: 'PK' });
    expect(r.status).toBe(200);
    expect(r.json().name).toBe('生データ.xlsx');
  });

  it('refuses bodies larger than 200 MB up front (413)', async () => {
    const r = await new Promise<number>((resolve, reject) => {
      const http = require('node:http') as typeof import('node:http');
      const req = http.request(
        { host: '127.0.0.1', port, method: 'POST', path: '/api/save', headers: { Host: `localhost:${port}`, 'x-file-name': 'big.xlsx', 'Content-Length': 300 * 1024 * 1024 } },
        (res) => {
          resolve(res.statusCode || 0);
          res.resume();
          req.destroy();
        },
      );
      req.on('error', (e) => (String(e).includes('socket hang up') || String(e).includes('ECONNRESET') ? undefined : reject(e)));
      req.write(Buffer.alloc(1024));
    });
    expect(r).toBe(413);
  });
});

function isInside(dir: string, f: string) {
  const rel = path.relative(path.resolve(dir), path.resolve(f));
  return rel !== '' && !rel.startsWith('..') && !path.isAbsolute(rel);
}

describe('POST /api/open and /api/open-folder', () => {
  it('opens only existing Excel files inside the output folder', async () => {
    const saved = (await saveReq('open-me.xlsx', 'PK')).json();
    const ok = await raw(port, 'POST', '/api/open', { headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ path: saved.path }) });
    expect(ok.status).toBe(200);
    expect(ok.json()).toMatchObject({ ok: true, opened: false }); // NO_OPEN: validated but not launched

    const outside = path.join(sb.dir, 'secret.xlsx');
    fs.writeFileSync(outside, 'PK');
    const cases: Array<[string, number]> = [
      [outside, 403],
      [path.join(sb.output, '..', 'secret.xlsx'), 403],
      [path.join(sb.output), 403],
      ['C:\\Windows\\System32\\notepad.exe', 403],
      [path.join(sb.output, 'nothing-here.xlsx'), 404],
    ];
    fs.writeFileSync(path.join(sb.output, 'script.bat'), 'echo hi');
    cases.push([path.join(sb.output, 'script.bat'), 403]);
    for (const [p, status] of cases) {
      const r = await raw(port, 'POST', '/api/open', { headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ path: p }) });
      expect(r.status, p).toBe(status);
      expect(typeof r.json().error).toBe('string');
    }
    const noPath = await raw(port, 'POST', '/api/open', { body: '{}' });
    expect(noPath.status).toBe(400);
    const badJson = await raw(port, 'POST', '/api/open', { body: '{' });
    expect(badJson.status).toBe(400);
    const get = await raw(port, 'GET', '/api/open');
    expect(get.status).toBe(405);
  });

  it('open-folder reports the output folder', async () => {
    const r = await raw(port, 'POST', '/api/open-folder', { body: '' });
    expect(r.status).toBe(200);
    expect(path.resolve(r.json().path)).toBe(path.resolve(sb.output));
  });
});

describe('jobs API (fake runner)', () => {
  const createJob = (job: unknown) => raw(port, 'POST', '/api/jobs', { headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(job) });
  const baseJob = (ops: unknown[] = [], format = 'xlsx') => ({ version: 1, format, ops, options: { recalc: true, clipboard: 'restoreEachOp' } });

  it('validates the job', async () => {
    const cases: Array<[unknown, RegExp]> = [
      [{ version: 2, format: 'xlsx', ops: [] }, /version/],
      [{ version: 1, format: 'csv', ops: [] }, /xlsx/],
      [{ version: 1, format: 'xlsx', ops: {} }, /配列/],
      [{ version: 1, format: 'xlsx', ops: [{ op: 'formatDisk' }] }, /未知の操作/],
      [{ version: 1, format: 'xlsx', ops: [5] }, /オブジェクト/],
      [[1, 2], /オブジェクト/],
    ];
    for (const [job, re] of cases) {
      const r = await createJob(job);
      expect(r.status, JSON.stringify(job)).toBe(400);
      expect(r.json().error).toMatch(re);
    }
    const bad = await raw(port, 'POST', '/api/jobs', { body: 'not json' });
    expect(bad.status).toBe(400);
    expect((await raw(port, 'GET', '/api/jobs/00000000-0000-0000-0000-000000000000')).status).toBe(404);
  });

  it('accepts every op kind the runner knows (round 2: setColumnWidth with a stored width, unmerge)', async () => {
    const ops = [
      { op: 'moveCols', sheet: 'S', start: 2, count: 1, to: 5, method: 'clipFree' },
      { op: 'moveRows', sheet: 'S', start: 2, count: 1, to: 5, method: 'clipFree' },
      { op: 'insertCols', sheet: 'S', at: 2, count: 1, formatFrom: 'left' },
      { op: 'insertRows', sheet: 'S', at: 2, count: 1, formatFrom: 'above' },
      { op: 'deleteCols', sheet: 'S', at: 2, count: 1 },
      { op: 'deleteRows', sheet: 'S', at: 2, count: 1 },
      { op: 'setCells', sheet: 'S', at: 'A1', rows: [[{ t: 'd', v: 46000, numFmt: 'yyyy/m/d' }, { f: '=SORT(B1:B3)', dyn: true }]], toValues: true },
      { op: 'setColumnWidth', sheet: 'S', col: 2, width: 20, unit: 'stored' },
      { op: 'unmerge', sheet: 'S', refs: ['A1:F1'], centerAcross: true },
    ];
    const c = await createJob(baseJob(ops));
    expect(c.status, c.text()).toBe(200);
    const { id } = c.json();
    expect((await raw(port, 'DELETE', `/api/jobs/${id}`)).status).toBe(200);
    // input / output / workDir are always chosen by the server: paths sent by a client are ignored (the runner would
    // otherwise copy C:\Windows\win.ini to a path of the client's choice)
    const evil = await createJob({ ...baseJob([]), input: 'C:\\Windows\\win.ini', output: path.join(sb.dir, 'evil.xlsx'), workDir: sb.dir });
    expect(evil.status).toBe(200);
    const eid = evil.json().id;
    const pkg = await makePackage('xlsx');
    expect((await raw(port, 'PUT', `/api/jobs/${eid}/input`, { body: pkg })).status).toBe(200);
    const done = await waitJob(port, eid);
    expect(done.state, JSON.stringify(done)).toBe('done');
    expect(Buffer.compare((await raw(port, 'GET', `/api/jobs/${eid}/result`)).body, pkg)).toBe(0);
    expect(fs.existsSync(path.join(sb.dir, 'evil.xlsx'))).toBe(false);
    expect((await raw(port, 'DELETE', `/api/jobs/${eid}`)).status).toBe(200);
  });

  it('runs a job end to end: create -> input -> done -> result -> delete', async () => {
    const pkg = await makePackage('xlsx');
    const c = await createJob(baseJob([{ op: 'moveCols', sheet: 'S', start: 2, count: 1, to: 5, method: 'clipFree' }]));
    expect(c.status).toBe(200);
    const { id } = c.json();
    expect(id).toMatch(/^[0-9a-f-]{36}$/);
    const waiting = (await raw(port, 'GET', `/api/jobs/${id}`)).json();
    expect(waiting).toMatchObject({ state: 'queued', waitingForInput: true });
    expect((await raw(port, 'GET', `/api/jobs/${id}/result`)).status).toBe(409);

    const notZip = await raw(port, 'PUT', `/api/jobs/${id}/input`, { body: 'hello' });
    expect(notZip.status).toBe(400);
    expect(notZip.json().error).toMatch(/Excel/);
    const cfb = Buffer.alloc(512);
    cfb.writeUInt32LE(0xe011cfd0, 0);
    const enc = await raw(port, 'PUT', `/api/jobs/${id}/input`, { body: cfb });
    expect(enc.status).toBe(400);
    expect(enc.json().error).toMatch(/パスワード/);

    const put = await raw(port, 'PUT', `/api/jobs/${id}/input`, { body: pkg });
    expect(put.status).toBe(200);
    expect(['queued', 'running', 'done']).toContain(put.json().state);
    expect((await raw(port, 'PUT', `/api/jobs/${id}/input`, { body: pkg })).status).toBe(409);

    const done = await waitJob(port, id);
    expect(done.state).toBe('done');
    expect(done.progress).toBe(1);
    expect(done.warnings).toContain('input=.xlsx');
    expect(done.warnings).toContain('テスト警告');
    const res = await raw(port, 'GET', `/api/jobs/${id}/result`);
    expect(res.status).toBe(200);
    expect(res.headers['content-type']).toContain('spreadsheetml');
    expect(Buffer.compare(res.body, pkg)).toBe(0);

    const tempDirs = fs.readdirSync(sb.temp);
    expect(tempDirs.length).toBeGreaterThan(0);
    expect((await raw(port, 'DELETE', `/api/jobs/${id}`)).status).toBe(200);
    expect((await raw(port, 'GET', `/api/jobs/${id}`)).status).toBe(404);
    await sleep(200);
    expect(fs.readdirSync(sb.temp).filter((d) => d.startsWith(`job-${id.slice(0, 8)}`))).toEqual([]);
  });

  it('keeps a macro package as .xlsm for Excel', async () => {
    const { id } = (await createJob(baseJob([], 'xlsm'))).json();
    await raw(port, 'PUT', `/api/jobs/${id}/input`, { body: await makePackage('xlsm') });
    const v = await waitJob(port, id);
    expect(v.state).toBe('done');
    expect(v.warnings).toContain('input=.xlsm');
    expect((await raw(port, 'GET', `/api/jobs/${id}/result`)).headers['content-type']).toContain('macroEnabled');
  });

  it('reports runner errors in Japanese', async () => {
    const { id } = (await createJob(baseJob([{ op: 'moveCols', sheet: 'FAIL', start: 1, count: 1, to: 3 }]))).json();
    await raw(port, 'PUT', `/api/jobs/${id}/input`, { body: await makePackage() });
    const v = await waitJob(port, id);
    expect(v.state).toBe('error');
    expect(v.errorCode).toBe('EXCEL_REFUSED');
    expect(v.error).toContain('列の移動');
    expect((await raw(port, 'GET', `/api/jobs/${id}/result`)).status).toBe(409);

    const crash = (await createJob(baseJob([{ op: 'setCells', sheet: 'CRASH', at: 'A1', rows: [] }]))).json();
    await raw(port, 'PUT', `/api/jobs/${crash.id}/input`, { body: await makePackage() });
    const vc = await waitJob(port, crash.id);
    expect(vc.state).toBe('error');
    expect(vc.errorCode).toBe('RUNNER_CRASHED');
    expect(vc.error).toMatch(/異常終了/);

    const noout = (await createJob(baseJob([{ op: 'setCells', sheet: 'NOOUT', at: 'A1', rows: [] }]))).json();
    await raw(port, 'PUT', `/api/jobs/${noout.id}/input`, { body: await makePackage() });
    const vn = await waitJob(port, noout.id);
    expect(vn.state).toBe('error');
    expect(vn.error).toMatch(/出力ファイル/);
  });

  it('runs one job at a time (queue) and cancels running / queued jobs', async () => {
    const pkg = await makePackage();
    const a = (await createJob(baseJob([{ op: 'debugWait', seconds: 1.5 }]))).json().id;
    const b = (await createJob(baseJob([]))).json().id;
    const c = (await createJob(baseJob([{ op: 'debugWait', seconds: 60 }]))).json().id;
    await raw(port, 'PUT', `/api/jobs/${a}/input`, { body: pkg });
    await raw(port, 'PUT', `/api/jobs/${b}/input`, { body: pkg });
    await sleep(300);
    const va = (await raw(port, 'GET', `/api/jobs/${a}`)).json();
    const vb = (await raw(port, 'GET', `/api/jobs/${b}`)).json();
    expect(va.state).toBe('running');
    expect(vb.state).toBe('queued');
    expect(vb.message).toContain('順番待ち');
    const status = (await raw(port, 'GET', '/api/status')).json();
    expect(status.jobs).toEqual({ running: 1, queued: 1 });
    expect((await waitJob(port, b)).state).toBe('done');
    expect((await raw(port, 'GET', `/api/jobs/${a}`)).json().state).toBe('done');

    await raw(port, 'PUT', `/api/jobs/${c}/input`, { body: pkg });
    await sleep(500);
    expect((await raw(port, 'GET', `/api/jobs/${c}`)).json().state).toBe('running');
    const t = Date.now();
    const del = await raw(port, 'DELETE', `/api/jobs/${c}`);
    expect(del.status).toBe(200);
    expect(Date.now() - t).toBeLessThan(10000);
    expect((await raw(port, 'GET', `/api/jobs/${c}`)).status).toBe(404);
    expect((await raw(port, 'GET', '/api/status')).json().jobs).toEqual({ running: 0, queued: 0 });

    // a job that never got its input can be deleted too
    const d = (await createJob(baseJob([]))).json().id;
    expect((await raw(port, 'DELETE', `/api/jobs/${d}`)).status).toBe(200);
  });

  it('times out a silent runner (JOB_IDLE_MS) with a Japanese message', async () => {
    const sb2 = makeSandbox('idle');
    const s2 = await startServer(sb2, { JOB_IDLE_MS: '1000' });
    try {
      const { id } = (await raw(s2.port, 'POST', '/api/jobs', { body: JSON.stringify(baseJob([{ op: 'debugWait', seconds: 30 }])) })).json();
      await raw(s2.port, 'PUT', `/api/jobs/${id}/input`, { body: await makePackage() });
      const v = await waitJob(s2.port, id, 15000);
      expect(v.state).toBe('error');
      expect(v.errorCode).toBe('TIMEOUT');
      expect(v.error).toMatch(/応答しなかった/);
    } finally {
      await s2.stop();
      fs.rmSync(sb2.dir, { recursive: true, force: true });
    }
  });
});

describe('missing build', () => {
  it('shows a Japanese build hint when dist/index.html is missing', async () => {
    const sb2 = makeSandbox('nodist');
    fs.rmSync(path.join(sb2.dist, 'index.html'));
    const s2 = await startServer(sb2);
    try {
      const r = await raw(s2.port, 'GET', '/');
      expect(r.status).toBe(503);
      expect(r.text()).toContain('build.bat');
      expect((await raw(s2.port, 'GET', '/api/ping')).status).toBe(200);
    } finally {
      await s2.stop();
      fs.rmSync(sb2.dir, { recursive: true, force: true });
    }
  });
});
