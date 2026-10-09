#!/usr/bin/env node
/**
 * Excel 統合ツール - local server (Node >= 18, built-in modules only).
 *
 * Runs three ways from this ONE source file (docs/research/R5-packaging.md):
 *   node server/server.mjs              dev / "system node" mode (files on disk)
 *   runtime\node.exe app\server.cjs     esbuild CJS bundle next to the signed node.exe ("runtime" layout)
 *   ExcelMerger.exe                     Node SEA (bundle + embedded index.html / excel-runner.ps1)
 * => no top-level await, no dynamic import(), import.meta.url only through the esbuild define.
 *
 * API (docs/ARCHITECTURE.md §4; JSON unless noted):
 *   GET    /api/ping                 {app:'excel-merger', version, pid}
 *   GET    /api/status               {app, version, excel:{installed, version?}, outputDir, jobs:{running, queued}, windows, sea,
 *                                     diskLow, diskFreeMB, diskMessage?, outputDiskLow}  (diskLow: < 2 GB free on the temp drive)
 *   POST   /api/jobs                 JSON RunnerJob -> {id}
 *   PUT    /api/jobs/:id/input       premerged .xlsx/.xlsm bytes -> queues the job -> {id, state}
 *   GET    /api/jobs/:id             {id, state:'queued'|'running'|'done'|'error', progress, message, error?, errorCode?, fallbackToJs?, warnings}
 *                                    (fallbackToJs: true = the built-in engine can do what Excel本体 could not, e.g. MOVE_VERIFY_FAILED)
 *   GET    /api/jobs/:id/result      output bytes
 *   DELETE /api/jobs/:id             cancel (kills only that job's own Excel PID) + clean up
 *   POST   /api/save                 bytes + header x-file-name (URI-encoded) -> {path, name}; never overwrites
 *   POST   /api/open                 {path} -> opens an output file (only inside the output folder)
 *   POST   /api/open-folder          [{path}] -> opens the output folder in Explorer (selecting path)
 *   GET    /api/events               SSE: the page keeps it open while the window lives
 *   POST   /api/bye                  pagehide beacon: the window is closing (or reloading)
 *
 * Security (round 5 review, tests/server/security.test.ts; pure helpers in ./safe.mjs): only requests whose Host header names this
 * machine are served (DNS rebinding); every API request (and every write) with a foreign Origin or a Sec-Fetch-Site other than
 * same-origin / none is refused; saves take the name from a custom header only (forces a CORS preflight, which this server never
 * grants) and clean it (device names incl. CON.tar.xlsx, invisible characters, alternate data streams); /api/open only opens regular
 * .xlsx / .xlsm files whose REAL path (junctions, symlinks, 8.3 names resolved) lies inside the output folder, and launches them
 * without any shell string; the op log of a job is whitelisted field by field before it reaches PowerShell. The server listens on
 * 127.0.0.1 only; the UI is opened as http://localhost:PORT/ (R5: Edge marks downloads from 127.0.0.1 with Mark-of-the-Web, which
 * would open exports in Protected View).
 *
 * Excel本体 jobs (round 5): the result is delivered as soon as the runner's final line has arrived and the output file is verified;
 * the shutdown of the hidden Excel (which can take ~60 s when its cloud services stall) goes on in the background under a watchdog
 * that kills ONLY that job's own Excel PID after EXCEL_EXIT_GRACE_MS. Excel's own TEMP (Office diagnostics, ~80 MB per start) is a
 * private folder that is deleted afterwards; the upload is streamed to disk and opened in place (no extra copies).
 *
 * Environment (all optional):
 *   PORT (5180; busy -> next free, 0 = random) · EM_HOST (127.0.0.1 or ::1 only) · NO_OPEN=1 (no window) ·
 *   APP_MODE=auto|edge|chrome|browser|none · DIST_DIR · OUTPUT_DIR · EM_DATA_DIR (%LOCALAPPDATA%\ExcelMerger) ·
 *   EM_TEMP_DIR (%TEMP%\excel-merger) · EXCEL_RUNNER (runner path; a .mjs fake in tests) · EXCEL_DETECT=none|installed ·
 *   EXIT_GRACE_MS (4000) · IDLE_EXIT_MS (30 min; 0 = never) · JOB_TIMEOUT_MS · JOB_IDLE_MS · XLRUNNER_INVOKE ·
 *   EXCEL_EXIT_GRACE_MS (6000) · EM_MIN_FREE_MB (disk check of Excel jobs, default 300 + 4 x the workbook size) ·
 *   EM_DISK_LOW_MB (status.diskLow threshold, 2048) · EM_MAX_BODY_MB (upload / save limit, 200)
 * Flags: --detach (launcher: start a hidden copy, print a line, exit; used by start.bat) · --foreground
 */
import http from 'node:http';
import fs from 'node:fs';
import fsp from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import crypto from 'node:crypto';
import zlib from 'node:zlib';
import { spawn, execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { detectExcel, runRunner } from './excel/run-job.mjs';
import { buildRunnerFile, runnerIsStale } from './excel/merge-runner.mjs';
import { OUTPUT_EXTS, buildOpenCommand, cleanOutputName, isInside, safeDecode, textCheckOpenPath, validateRunnerOps } from './safe.mjs';

const T0 = Date.now();
const APP_ID = 'excel-merger';
const APP_TITLE = 'Excel 統合ツール';

const NODE_MAJOR = Number(String(process.versions.node).split('.')[0]);
if (!(NODE_MAJOR >= 18)) {
  console.error(`[エラー] Node.js 18 以上が必要です（現在 v${process.versions.node}）。https://nodejs.org/ から LTS 版を入れてください。`);
  process.exit(1);
}

// ---------------------------------------------------------------------------------------------------------
// Where are we? SEA detection without a static import of node:sea (works in ESM, in the CJS bundle, in SEA).
// ---------------------------------------------------------------------------------------------------------
const sea = (() => {
  try {
    return typeof process.getBuiltinModule === 'function' ? process.getBuiltinModule('node:sea') : null;
  } catch {
    return null;
  }
})();
const IS_SEA = !!(sea && typeof sea.isSea === 'function' && sea.isSea());
// In the CJS bundle import.meta.url is defined from __filename (scripts/build-exe.mjs).
const SCRIPT_DIR = IS_SEA ? path.dirname(process.execPath) : path.dirname(fileURLToPath(import.meta.url));
/** App folder: next to the exe (SEA), the parent of server/ (dev) or of app/ (runtime layout). */
const APP_DIR = IS_SEA ? path.dirname(process.execPath) : path.resolve(SCRIPT_DIR, '..');

const envStr = (name) => String(process.env[name] ?? '').trim();
const envFlag = (name) => /^(1|true|yes|on)$/i.test(envStr(name));
function envInt(name, fallback, min = 0) {
  const v = envStr(name);
  if (!v) return fallback;
  const n = Number(v);
  return Number.isFinite(n) && n >= min ? Math.floor(n) : fallback;
}

// The server only ever listens on the loopback interface. A generic HOST variable (tcsh, some shells and tools export the computer's
// name in it) must NOT move it to a LAN address: that would hand the file-saving / Excel-running API to every PC of the network.
const WANTED_HOST = (envStr('EM_HOST') || envStr('HOST') || '127.0.0.1').toLowerCase();
const HOST = WANTED_HOST === '::1' || WANTED_HOST === '[::1]' ? '::1' : '127.0.0.1';
const HOST_IGNORED = ['127.0.0.1', 'localhost', '::1', '[::1]'].includes(WANTED_HOST) ? '' : WANTED_HOST;
const LOCAL_HOSTS = new Set(['localhost', '127.0.0.1', '::1', '[::1]', HOST.toLowerCase(), `[${HOST.toLowerCase()}]`]);
const BASE_PORT = envInt('PORT', 5180);
const MAX_PORT_TRIES = 10;
const MAX_BODY_BYTES = envInt('EM_MAX_BODY_MB', 200, 1) * 1024 * 1024; // upload / save limit (EM_MAX_BODY_MB lowers it for tests)
const MAX_JOB_JSON_BYTES = Math.min(64 * 1024 * 1024, MAX_BODY_BYTES);
const MAX_LIVE_JOBS = 20; // created-but-unfinished jobs kept in memory (each holds its op log)
const MAX_EVENT_CLIENTS = 64; // SSE connections
const DISK_LOW_BYTES = envInt('EM_DISK_LOW_MB', 2048, 1) * 1024 * 1024; // status.diskLow below this (2 GB)
const EXCEL_EXIT_GRACE_MS = envInt('EXCEL_EXIT_GRACE_MS', 6000, 500);
const EXIT_GRACE_MS = envInt('EXIT_GRACE_MS', 4000);
const IDLE_EXIT_MS = envInt('IDLE_EXIT_MS', 30 * 60 * 1000);
const APP_MODE = (envStr('APP_MODE') || (envFlag('NO_OPEN') || process.argv.includes('--no-open') ? 'none' : 'auto')).toLowerCase();
const LOCALAPPDATA = envStr('LOCALAPPDATA') || path.join(os.homedir(), 'AppData', 'Local');
const DATA_DIR = envStr('EM_DATA_DIR') ? path.resolve(envStr('EM_DATA_DIR')) : path.join(LOCALAPPDATA, 'ExcelMerger');
const DIST_DIR = envStr('DIST_DIR') ? path.resolve(envStr('DIST_DIR')) : path.join(APP_DIR, 'dist');
const PREFERRED_OUTPUT_DIR = envStr('OUTPUT_DIR') ? path.resolve(envStr('OUTPUT_DIR')) : path.join(APP_DIR, 'output');
const FALLBACK_OUTPUT_DIR = path.join(os.homedir(), 'Documents', 'ExcelMerger', 'output');
const TEMP_ROOT = pickTempRoot();
const JOB_KEEP_MS = 15 * 60 * 1000; // finished jobs (and their temp files) are dropped after this
/**
 * Runner error codes after which the built-in (JS) engine is the right answer (nothing was saved by Excel):
 * MOVE_VERIFY_FAILED = a move Excel本体 could not do faithfully (a column moved inside an Excel table with cells above / below
 * it, the clipboard kept busy by another program ...), EXCEL_UNAVAILABLE / PASSWORD_PROTECTED / OPEN_FAILED = Excel cannot open it,
 * DISK_LOW = too little free space on the temp drive for Excel (the built-in engine works in memory).
 */
const FALLBACK_TO_JS_CODES = new Set(['MOVE_VERIFY_FAILED', 'EXCEL_UNAVAILABLE', 'PASSWORD_PROTECTED', 'OPEN_FAILED', 'RUNNER_MISSING', 'DISK_LOW']);
const BUILT_IN_HINT = '「内蔵エンジン」で出力してください。';
// eslint-disable-next-line no-undef
const VERSION = typeof __APP_VERSION__ === 'string' ? __APP_VERSION__ : readPackageVersion();

let listeningPort = 0;
let outputDir = PREFERRED_OUTPUT_DIR;

function readPackageVersion() {
  for (const dir of [APP_DIR, path.resolve(SCRIPT_DIR, '..')]) {
    try {
      const pkg = JSON.parse(fs.readFileSync(path.join(dir, 'package.json'), 'utf8'));
      if (pkg && pkg.version) return String(pkg.version);
    } catch {
      /* next */
    }
  }
  return '1.0.0';
}

/**
 * Temp root for jobs: %TEMP%excel-merger (the user's own temp folder: only that user can write there). Earlier builds fell back to an
 * ASCII folder under ProgramData when %TEMP% had Japanese characters; that folder is shared by every user of the PC (another user
 * could pre-create it and read or swap the job files), and Excel + PowerShell work with Unicode paths (verified, round 5).
 */
function pickTempRoot() {
  if (envStr('EM_TEMP_DIR')) return path.resolve(envStr('EM_TEMP_DIR'));
  return path.join(os.tmpdir(), 'excel-merger');
}

// ---------------------------------------------------------------------------------------------------------
// Logging (Japanese for the user, plain timestamps; in --detach mode stdout goes to the log file)
// ---------------------------------------------------------------------------------------------------------
function log(line) {
  const t = new Date();
  const p = (n) => String(n).padStart(2, '0');
  console.log(`[${p(t.getHours())}:${p(t.getMinutes())}:${p(t.getSeconds())}] ${line}`);
}

// ---------------------------------------------------------------------------------------------------------
// Assets: embedded (SEA) or on disk
// ---------------------------------------------------------------------------------------------------------
function seaAsset(key) {
  if (!IS_SEA) return null;
  try {
    return Buffer.from(sea.getAsset(key));
  } catch {
    return null;
  }
}

let indexCache = null; // { buf, mtimeMs } (disk) or { buf } (SEA)
async function readIndexHtml() {
  if (IS_SEA && !envStr('DIST_DIR')) {
    if (!indexCache) {
      const buf = seaAsset('index.html');
      if (buf) indexCache = { buf, mtimeMs: 0 };
    }
    if (indexCache) return indexCache.buf;
  }
  const file = path.join(DIST_DIR, 'index.html');
  const st = await statOrNull(file);
  if (!st || !st.isFile()) return null;
  if (!indexCache || indexCache.mtimeMs !== st.mtimeMs) indexCache = { buf: await fsp.readFile(file), mtimeMs: st.mtimeMs };
  return indexCache.buf;
}

/** Path of excel-runner.ps1: env override, the built file on disk (rebuilt from sources in dev), or the SEA asset. */
let runnerPathCache = null;
function resolveRunner() {
  if (envStr('EXCEL_RUNNER')) return path.resolve(envStr('EXCEL_RUNNER'));
  if (runnerPathCache && fs.existsSync(runnerPathCache)) return runnerPathCache;
  const candidates = [path.join(SCRIPT_DIR, 'excel'), path.join(APP_DIR, 'server', 'excel'), APP_DIR];
  if (!IS_SEA) {
    for (const dir of candidates) {
      const hasSources = fs.existsSync(path.join(dir, 'excel-runner.src.ps1')) && fs.existsSync(path.join(dir, 'XlRunnerHelper.cs'));
      if (hasSources && runnerIsStale(dir)) {
        try {
          buildRunnerFile(dir);
          log('excel-runner.ps1 をソースから作り直しました');
        } catch (e) {
          log(`excel-runner.ps1 を作れませんでした: ${e && e.message}`);
        }
      }
      const f = path.join(dir, 'excel-runner.ps1');
      if (fs.existsSync(f)) return (runnerPathCache = f);
    }
    return null;
  }
  // SEA: write the embedded script once to %LOCALAPPDATA%\ExcelMerger\ps\<sha12>-excel-runner.ps1 (bytes as-is,
  // so the UTF-8 BOM that Windows PowerShell 5.1 needs stays).
  const buf = seaAsset('excel-runner.ps1');
  if (!buf) return null;
  const sha = crypto.createHash('sha256').update(buf).digest('hex').slice(0, 12);
  const dir = path.join(DATA_DIR, 'ps');
  const file = path.join(dir, `${sha}-excel-runner.ps1`);
  try {
    // The file name carries the hash of the content, but an existing file is only trusted after comparing its bytes (a half-written
    // copy from a crash, or a planted file with the same name, would otherwise be executed).
    let intact = false;
    try {
      intact = Buffer.compare(fs.readFileSync(file), buf) === 0;
    } catch {
      intact = false;
    }
    if (!intact) {
      fs.mkdirSync(dir, { recursive: true });
      const tmp = `${file}.${process.pid}.tmp`;
      fs.writeFileSync(tmp, buf);
      fs.renameSync(tmp, file);
    }
    return (runnerPathCache = file);
  } catch (e) {
    log(`excel-runner.ps1 を展開できませんでした: ${e && e.message}`);
    return null;
  }
}

// ---------------------------------------------------------------------------------------------------------
// Small helpers
// ---------------------------------------------------------------------------------------------------------
async function statOrNull(file) {
  try {
    return await fsp.stat(file);
  } catch {
    return null;
  }
}
const headerValue = (raw) => (Array.isArray(raw) ? raw[0] : raw);

/** DNS-rebinding guard: only Host headers naming this machine (a missing Host is tolerated). */
function hostAllowed(raw) {
  const v = headerValue(raw);
  if (v === undefined || v === '') return true;
  if (typeof v !== 'string') return false;
  const value = v.trim().toLowerCase();
  const host = value.startsWith('[') ? value.replace(/\]:\d+$/, ']') : value.replace(/:\d+$/, '');
  return LOCAL_HOSTS.has(host);
}
/** Cross-site write guard: an Origin must be this server (host + port); no Origin (curl, tests) passes. */
function originAllowed(raw) {
  const v = headerValue(raw);
  if (v === undefined || v === '') return true;
  if (typeof v !== 'string') return false;
  let origin;
  try {
    origin = new URL(v.trim());
  } catch {
    return false; // "null" (file:// pages, sandboxed frames)
  }
  if (origin.protocol !== 'http:') return false;
  const port = origin.port ? Number(origin.port) : 80;
  if (listeningPort !== 0 && port !== listeningPort) return false;
  return LOCAL_HOSTS.has(origin.hostname.toLowerCase());
}
function fetchSiteAllowed(raw) {
  const v = headerValue(raw);
  if (v === undefined || v === '') return true;
  const s = String(v).trim().toLowerCase();
  return s === 'same-origin' || s === 'none';
}

/**
 * Every API request (reads included) and every write must come from this server's own page: a page of another site can neither
 * trigger nor (with CORS never granted) read anything. Browsers send Sec-Fetch-Site on all requests; Origin on cross-origin ones.
 */
function crossSiteBlocked(req) {
  return !originAllowed(req.headers.origin) || !fetchSiteAllowed(req.headers['sec-fetch-site']);
}

const COMMON_HEADERS = { 'X-Content-Type-Options': 'nosniff', 'Referrer-Policy': 'no-referrer', 'Cross-Origin-Resource-Policy': 'same-origin' };
/** pages: nobody may frame the app (clickjacking of the export / open buttons) */
const HTML_HEADERS = { ...COMMON_HEADERS, 'X-Frame-Options': 'DENY', 'Content-Security-Policy': "frame-ancestors 'none'" };

function sendJson(res, status, body) {
  const text = JSON.stringify(body);
  res.writeHead(status, {
    'Content-Type': 'application/json; charset=utf-8',
    'Content-Length': Buffer.byteLength(text),
    'Cache-Control': 'no-store',
    ...COMMON_HEADERS,
  });
  res.end(text);
}
const sendError = (res, status, message, extra = {}) => sendJson(res, status, { error: message, ...extra });

class HttpError extends Error {
  constructor(status, message, extra) {
    super(message);
    this.status = status;
    this.extra = extra || {};
  }
}

const tooLarge = (limit) => new HttpError(413, `データが大きすぎます（上限 ${Math.round(limit / 1048576)} MB）`);
const DISK_FULL_MESSAGE = 'ディスクの空き容量が足りないため保存できませんでした。不要なファイルを削除してから、もう一度お試しください。';

/** Reads a request body into memory (limit in bytes). */
function readBody(req, limit) {
  return new Promise((resolve, reject) => {
    const declared = Number(req.headers['content-length']);
    if (Number.isFinite(declared) && declared > limit) {
      req.resume();
      reject(tooLarge(limit));
      return;
    }
    const chunks = [];
    let size = 0;
    let over = false;
    req.on('data', (c) => {
      size += c.length;
      if (over) return;
      if (size > limit) {
        over = true;
        chunks.length = 0;
        return;
      }
      chunks.push(c);
    });
    req.on('end', () => (over ? reject(tooLarge(limit)) : resolve(Buffer.concat(chunks))));
    req.on('error', () => reject(new HttpError(400, 'データを受け取れませんでした')));
    req.on('aborted', () => reject(new HttpError(400, 'データの送信が中断されました')));
  });
}

async function readJson(req, limit) {
  const buf = await readBody(req, limit);
  try {
    return JSON.parse(buf.toString('utf8'));
  } catch {
    throw new HttpError(400, 'JSON の形式が正しくありません');
  }
}

/**
 * Streams a request body into a file (nothing is held in memory: a 200 MB workbook no longer costs 200 MB of RAM, and the disk
 * is written once). flags 'wx' = the file must not exist. The limit is enforced while streaming; on ANY failure (too large,
 * aborted, disk full) the partial file is removed. Resolves with the byte count. The rest of a refused body is drained (not
 * piped: destroying the request would close the socket before the error answer can be sent).
 */
function streamToFile(req, file, limit, flags = 'wx') {
  return new Promise((resolve, reject) => {
    const declared = Number(req.headers['content-length']);
    if (Number.isFinite(declared) && declared > limit) {
      req.resume();
      reject(tooLarge(limit));
      return;
    }
    const ws = fs.createWriteStream(file, { flags });
    let size = 0;
    let error = null;
    let settled = false;
    const finish = () => {
      if (settled) return;
      settled = true;
      if (error) {
        fsp
          .unlink(file)
          .catch(() => {})
          .then(() => reject(error));
      } else resolve(size);
    };
    const fail = (e) => {
      if (error) return;
      error = e;
      ws.destroy();
      req.resume();
    };
    ws.on('error', (e) => fail(e && e.code === 'ENOSPC' ? new HttpError(507, DISK_FULL_MESSAGE, { errorCode: 'DISK_LOW' }) : new HttpError(500, 'ファイルを書き込めませんでした')));
    ws.on('close', finish);
    req.on('data', (c) => {
      if (error) return;
      size += c.length;
      if (size > limit) return fail(tooLarge(limit));
      if (!ws.write(c)) {
        req.pause();
        ws.once('drain', () => req.resume());
      }
    });
    req.on('end', () => {
      if (!error) ws.end();
    });
    req.on('error', () => fail(new HttpError(400, 'データを受け取れませんでした')));
    req.on('aborted', () => fail(new HttpError(400, 'データの送信が中断されました')));
  });
}

/**
 * Looks inside a workbook package WITHOUT reading all of it (a tail read for the ZIP directory, a few small reads for
 * [Content_Types].xml): must be a ZIP with xl/workbook.xml; tells .xlsx from .xlsm (Excel refuses a macro package saved as .xlsx).
 * `read(pos, len)` returns a Buffer. Returns { ext: '.xlsx' | '.xlsm' } or { error } (Japanese).
 */
async function inspectZip(size, read) {
  const broken = { error: 'Excel ブックが途中までしか届いていないか、壊れています' };
  try {
    if (size < 30) return { error: 'Excel ブックのデータが空か、壊れています' };
    const sig = (await read(0, 4)).readUInt32LE(0);
    if (sig === 0xe011cfd0) return { error: 'パスワード付き、または古い形式 (.xls) のブックは Excel本体モードで処理できません' };
    if (sig !== 0x04034b50) return { error: 'Excel ブック (.xlsx) ではないデータが送られました' };
    const tailLen = Math.min(size, 65557);
    const tail = await read(size - tailLen, tailLen);
    let eocd = -1;
    for (let i = tail.length - 22; i >= 0; i--) {
      if (tail.readUInt32LE(i) === 0x06054b50) {
        eocd = i;
        break;
      }
    }
    if (eocd < 0) return broken;
    const entries = tail.readUInt16LE(eocd + 10);
    const cdSize = tail.readUInt32LE(eocd + 12);
    const cdOff = tail.readUInt32LE(eocd + 16);
    if (entries === 0xffff || cdOff === 0xffffffff || cdSize === 0xffffffff) return { ext: '.xlsx' }; // ZIP64 (> 4 GB): not inspected
    if (cdOff + cdSize > size || cdSize > 64 * 1024 * 1024) return broken;
    const cd = await read(cdOff, cdSize);
    let p = 0;
    let hasWorkbook = false;
    let contentTypes = null;
    let hasVba = false;
    for (let n = 0; n < entries && p + 46 <= cd.length; n++) {
      if (cd.readUInt32LE(p) !== 0x02014b50) break;
      const method = cd.readUInt16LE(p + 10);
      const compSize = cd.readUInt32LE(p + 20);
      const nameLen = cd.readUInt16LE(p + 28);
      const extraLen = cd.readUInt16LE(p + 30);
      const commentLen = cd.readUInt16LE(p + 32);
      const local = cd.readUInt32LE(p + 42);
      const name = cd.toString('utf8', p + 46, p + 46 + nameLen);
      if (/^xl\/workbook\.xml$/i.test(name)) hasWorkbook = true;
      if (/^xl\/vbaProject\.bin$/i.test(name)) hasVba = true;
      if (name === '[Content_Types].xml' && compSize <= 4 * 1024 * 1024 && local + 30 <= size) {
        const lh = await read(local, 30);
        if (lh.length === 30 && lh.readUInt32LE(0) === 0x04034b50) {
          const start = local + 30 + lh.readUInt16LE(26) + lh.readUInt16LE(28);
          const data = await read(start, compSize);
          try {
            // a few KB in every real workbook: a decompression bomb ([Content_Types].xml that inflates to gigabytes) is cut off at 4 MB
            contentTypes = (method === 8 ? zlib.inflateRawSync(data, { maxOutputLength: 4 * 1024 * 1024 }) : data).toString('utf8');
          } catch {
            contentTypes = null;
          }
        }
      }
      p += 46 + nameLen + extraLen + commentLen;
    }
    if (!hasWorkbook) return { error: 'Excel ブック (.xlsx) ではありません（xl/workbook.xml がありません）' };
    if (contentTypes) return { ext: /sheet\.macroEnabled\.main\+xml/i.test(contentTypes) ? '.xlsm' : '.xlsx' };
    return { ext: hasVba ? '.xlsm' : '.xlsx' };
  } catch {
    return broken;
  }
}

async function inspectPackageFile(file) {
  let fh;
  try {
    fh = await fsp.open(file, 'r');
    const { size } = await fh.stat();
    return await inspectZip(size, async (pos, len) => {
      const b = Buffer.alloc(len);
      const { bytesRead } = await fh.read(b, 0, len, pos);
      return bytesRead === len ? b : b.subarray(0, bytesRead);
    });
  } catch {
    return { error: 'Excel ブックを読み取れませんでした' };
  } finally {
    if (fh) await fh.close().catch(() => {});
  }
}

// ---------------------------------------------------------------------------------------------------------
// Disk space (the system drive that holds %TEMP% is often nearly full)
// ---------------------------------------------------------------------------------------------------------
/** Free bytes on the drive of `dir` (an existing parent is used when it does not exist yet); null = unknown (Node < 18.15 has no statfs). */
async function freeBytes(dir) {
  if (typeof fsp.statfs !== 'function') return null;
  let d = path.resolve(dir);
  for (let i = 0; i < 16; i++) {
    try {
      const st = await fsp.statfs(d);
      return Number(st.bavail) * Number(st.bsize);
    } catch {
      const up = path.dirname(d);
      if (up === d) return null;
      d = up;
    }
  }
  return null;
}
const fmtSize = (mb) => (mb >= 1024 ? `${(mb / 1024).toFixed(1)} GB` : `${mb} MB`);
const driveOf = (p) => (/^[a-zA-Z]:/.test(p) ? `${p.slice(0, 2).toUpperCase()} ドライブ` : '一時フォルダのドライブ');

/** Fields added to /api/status: diskLow = less than 2 GB free on the drive that holds the temp folder. */
async function diskReport() {
  const tempFree = await freeBytes(TEMP_ROOT);
  const outFree = await freeBytes(outputDir);
  const rep = {
    diskLow: tempFree !== null && tempFree < DISK_LOW_BYTES,
    diskFreeMB: tempFree === null ? null : Math.floor(tempFree / 1048576),
    outputDiskLow: outFree !== null && outFree < DISK_LOW_BYTES,
  };
  if (rep.diskLow) {
    rep.diskMessage = `${driveOf(TEMP_ROOT)}の空き容量が少なくなっています（残り ${fmtSize(rep.diskFreeMB)}）。Excel本体での出力は一時ファイルを作るため、空きが足りないと失敗することがあります。不要なファイルを削除するか、「内蔵エンジン」で出力してください。`;
  }
  return rep;
}

/** Free space an Excel本体 job of a workbook of `bytes` needs on the temp drive: Excel's own files (~100-300 MB of diagnostics / caches) + 4 copies. */
function minFreeForJob(bytes) {
  const fixed = envInt('EM_MIN_FREE_MB', 300) * 1048576;
  return fixed + 4 * (Number.isFinite(bytes) ? bytes : 0);
}
/** HttpError (507, errorCode DISK_LOW: the UI falls back to the built-in engine) when the temp drive cannot hold an Excel job of that size. */
async function diskLowError(bytes) {
  const free = await freeBytes(TEMP_ROOT);
  if (free === null || free >= minFreeForJob(bytes)) return null;
  return new HttpError(
    507,
    `${driveOf(TEMP_ROOT)}の空き容量が足りません（残り ${fmtSize(Math.floor(free / 1048576))}）。Excel本体での出力には一時ファイルの領域が必要です。${BUILT_IN_HINT}`,
    { errorCode: 'DISK_LOW', fallbackToJs: true },
  );
}

// ---------------------------------------------------------------------------------------------------------
// Output folder / save / open
// ---------------------------------------------------------------------------------------------------------
async function ensureOutputDir() {
  try {
    await fsp.mkdir(outputDir, { recursive: true });
    return outputDir;
  } catch (e) {
    if (outputDir !== FALLBACK_OUTPUT_DIR && !envStr('OUTPUT_DIR')) {
      log(`出力フォルダを作れませんでした (${e && e.code}): ${outputDir} -> ${FALLBACK_OUTPUT_DIR} を使います`);
      outputDir = FALLBACK_OUTPUT_DIR;
      await fsp.mkdir(outputDir, { recursive: true });
      return outputDir;
    }
    throw e;
  }
}

/**
 * Receives the request body as `name` in the output folder; an existing file is never replaced: "name (2).xlsx", ... The name
 * is claimed first (exclusive create), then the body is streamed into it (nothing is held in memory); a failed upload leaves no file.
 */
async function saveUnique(req, name, limit) {
  const dir = await ensureOutputDir();
  const ext = path.extname(name);
  const stem = name.slice(0, name.length - ext.length);
  for (let i = 1; i <= 9999; i++) {
    const candidate = i === 1 ? name : `${stem} (${i})${ext}`;
    const full = path.join(dir, candidate);
    try {
      const fh = await fsp.open(full, 'wx'); // exclusive create: fails with EEXIST instead of overwriting
      await fh.close();
    } catch (e) {
      if (e && e.code === 'EEXIST') continue;
      if (e && (e.code === 'EPERM' || e.code === 'EACCES') && fs.existsSync(full)) continue; // locked existing file
      throw e;
    }
    try {
      const size = await streamToFile(req, full, limit, 'w');
      if (size === 0) throw new HttpError(400, '保存するデータが空です');
    } catch (e) {
      await fsp.unlink(full).catch(() => {});
      throw e;
    }
    return { path: full, name: candidate };
  }
  throw new Error('同じ名前のファイルが多すぎます');
}

function outputDirs() {
  return [...new Set([outputDir, PREFERRED_OUTPUT_DIR, FALLBACK_OUTPUT_DIR])];
}

/**
 * The file /api/open (or the selection of /api/open-folder) may act on: an existing regular .xlsx / .xlsm file whose REAL path
 * lies inside one of the output folders. Text checks first (UNC / device paths, alternate data streams, control characters),
 * then the real path: a junction or symbolic link inside the output folder that leads elsewhere, 8.3 short names and trailing
 * dots / spaces are resolved before the containment test, and the resolved path is what gets opened.
 * Returns { file } or { status, error } (Japanese).
 */
async function resolveOutputFile(raw) {
  const bad = textCheckOpenPath(raw);
  if (bad) return bad;
  const target = path.resolve(raw);
  if (!outputDirs().some((d) => isInside(d, target))) return { status: 403, error: '出力フォルダの中のファイルだけを開けます' };
  if (!OUTPUT_EXTS.has(path.extname(target).toLowerCase())) return { status: 403, error: 'Excel ファイル (.xlsx / .xlsm) だけを開けます' };
  let real;
  try {
    const link = await fsp.lstat(target);
    if (link.isSymbolicLink()) return { status: 403, error: 'ショートカット（リンク）は開けません' };
    real = await fsp.realpath(target);
  } catch {
    return { status: 404, error: `ファイルが見つかりません: ${path.basename(target)}` };
  }
  let inside = false;
  for (const d of outputDirs()) {
    try {
      if (isInside(await fsp.realpath(d), real)) {
        inside = true;
        break;
      }
    } catch {
      /* this output folder does not exist (yet) */
    }
  }
  if (!inside) return { status: 403, error: '出力フォルダの外のファイルは開けません' };
  if (!OUTPUT_EXTS.has(path.extname(real).toLowerCase())) return { status: 403, error: 'Excel ファイル (.xlsx / .xlsm) だけを開けます' };
  const st = await statOrNull(real);
  if (!st || !st.isFile()) return { status: 404, error: `ファイルが見つかりません: ${path.basename(target)}` };
  return { file: real };
}

/**
 * Opens a file (with its default program = Excel) or a folder (Explorer), or selects a file in Explorer. The command comes from
 * safe.mjs buildOpenCommand: never a shell string (no cmd.exe / start), and the path is handed over in an environment variable.
 */
function shellOpen(target, select = false) {
  if (APP_MODE === 'none' || process.platform !== 'win32') return false;
  let cmd;
  try {
    cmd = buildOpenCommand(target, select, envStr('SystemRoot') || 'C:\\Windows');
  } catch {
    return false;
  }
  const c = spawn(cmd.cmd, cmd.args, { stdio: 'ignore', detached: true, windowsHide: !select, windowsVerbatimArguments: cmd.verbatim, env: { ...process.env, ...cmd.env } });
  c.on('error', (e) => log(`開けませんでした: ${e && e.message}`));
  c.unref();
  return true;
}

// ---------------------------------------------------------------------------------------------------------
// Jobs (Excel本体 engine): one at a time, FIFO
// ---------------------------------------------------------------------------------------------------------
/** @type {Map<string, any>} */
const jobs = new Map();
const queue = [];
let runningJob = null;

function jobView(j) {
  const v = { id: j.id, state: j.state, progress: j.progress, message: j.message, warnings: j.warnings };
  if (j.state === 'error') {
    v.error = j.error;
    v.errorCode = j.errorCode;
    // the built-in engine can do what Excel本体 could not: the UI retries with it by itself (src/api.ts ApiError.fallbackToJs)
    if (FALLBACK_TO_JS_CODES.has(j.errorCode)) v.fallbackToJs = true;
  }
  if (j.state === 'queued' && !j.inputReady) v.waitingForInput = true;
  if (j.state === 'queued' && j.inputReady) {
    const pos = queue.indexOf(j.id);
    if (pos >= 0) v.queuePosition = pos + 1 + (runningJob ? 1 : 0);
  }
  return v;
}

function validateJob(body) {
  if (!body || typeof body !== 'object' || Array.isArray(body)) throw new HttpError(400, 'ジョブは JSON オブジェクトで送ってください');
  if (body.version !== 1) throw new HttpError(400, `未対応のジョブ形式です (version ${body.version})`);
  const format = body.format === 'xlsm' ? 'xlsm' : body.format === 'xlsx' || body.format === undefined ? 'xlsx' : null;
  if (!format) throw new HttpError(400, `出力形式は xlsx か xlsm にしてください: ${body.format}`);
  const ops = body.ops === undefined ? [] : body.ops;
  if (!Array.isArray(ops)) throw new HttpError(400, 'ops は配列にしてください');
  if (ops.length > 200000) throw new HttpError(400, '操作が多すぎます');
  // every op is rebuilt from its known, typed fields (safe.mjs): what reaches PowerShell / Excel is never the raw request
  const checked = validateRunnerOps(ops, { allowDebug: envFlag('XLRUNNER_ALLOW_DEBUG') });
  if (checked.error) throw new HttpError(400, checked.error);
  const o = body.options && typeof body.options === 'object' ? body.options : {};
  const options = {
    recalc: o.recalc !== false,
    clipboard: o.clipboard === 'restoreAtEnd' ? 'restoreAtEnd' : 'restoreEachOp',
  };
  // input / output / workDir are always chosen by the server (paths inside its private job folder).
  return { version: 1, format, ops: checked.ops, options };
}

/** Drops finished jobs (and their folders) at once: makes room when too many jobs were created and never finished with. */
function purgeFinishedJobs() {
  for (const j of jobs.values()) {
    if (!j.finishedAt) continue;
    jobs.delete(j.id);
    removeDir(j.dir).catch(() => {});
  }
}

async function createJob(req, res) {
  const body = await readJson(req, MAX_JOB_JSON_BYTES);
  const job = validateJob(body);
  if (jobs.size >= MAX_LIVE_JOBS) purgeFinishedJobs();
  if (jobs.size >= MAX_LIVE_JOBS) throw new HttpError(429, '処理中のジョブが多すぎます。しばらく待ってから、もう一度お試しください。');
  let id;
  let dir;
  do {
    id = crypto.randomUUID();
    dir = path.join(TEMP_ROOT, `job-${id.slice(0, 8)}-${process.pid}`);
  } while ([...jobs.values()].some((x) => x.dir === dir));
  jobs.set(id, {
    id,
    dir,
    job,
    state: 'queued',
    inputReady: false,
    receiving: false,
    progress: 0,
    message: '入力データを待っています',
    warnings: [],
    createdAt: Date.now(),
    finishedAt: 0,
    controller: new AbortController(),
    inputFile: null,
    outputFile: null,
    xlTemp: null,
  });
  log(`ジョブ ${id.slice(0, 8)} を受け付けました（操作 ${job.ops.length} 件, ${job.format}）`);
  return sendJson(res, 200, { id });
}

/**
 * The uploaded workbook is streamed straight into the job folder as work.xlsx (the runner opens it in place, Excel saves into
 * output.xlsx next to it): two files on disk instead of four, and no memory spike for a big workbook.
 */
async function putJobInput(req, res, j) {
  if (j.inputReady) throw new HttpError(409, 'このジョブの入力は既に受け取っています');
  if (j.state !== 'queued') throw new HttpError(409, 'このジョブは既に終了しています');
  if (j.receiving) throw new HttpError(409, 'このジョブの入力を受信中です');
  const low = await diskLowError(Number(req.headers['content-length']));
  if (low) {
    req.resume();
    throw low;
  }
  j.receiving = true;
  try {
    await fsp.mkdir(j.dir, { recursive: true });
    const part = path.join(j.dir, 'upload.part');
    await streamToFile(req, part, MAX_BODY_BYTES);
    const info = await inspectPackageFile(part);
    if (info.error) {
      await fsp.unlink(part).catch(() => {});
      throw new HttpError(400, info.error);
    }
    if (!jobs.has(j.id) || j.state !== 'queued') {
      // cancelled while uploading
      await removeDir(j.dir);
      throw new HttpError(409, 'このジョブは中止されました');
    }
    j.inputFile = path.join(j.dir, `work${info.ext}`);
    j.outputFile = path.join(j.dir, `output.${j.job.format}`);
    await fsp.rename(part, j.inputFile);
  } finally {
    j.receiving = false;
  }
  j.inputReady = true;
  j.message = runningJob ? '順番待ちです（前の処理が終わると始まります）' : 'まもなく開始します';
  queue.push(j.id);
  setImmediate(pump);
  return sendJson(res, 200, jobView(j));
}

const PHASE_TEXT = {
  excel: 'Excel を起動しました',
  open: 'ブックを開きました',
  calc: '再計算しました',
  save: '保存しました',
};
const OP_TEXT = {
  moveCols: '列を移動',
  moveRows: '行を移動',
  insertCols: '列を挿入',
  insertRows: '行を挿入',
  deleteCols: '列を削除',
  deleteRows: '行を削除',
  setCells: 'セルを書き込み',
  setColumnWidth: '列幅を変更',
  unmerge: 'セル結合を解除',
  debugWait: '待機',
};

function onRunnerEvent(j, ev) {
  if (!ev || typeof ev !== 'object') return;
  if (ev.type === 'start') {
    j.progress = 0.03;
    j.message = 'Excel を起動しています…';
  } else if (ev.type === 'excel') {
    j.excelPid = ev.pid;
    j.progress = 0.06;
    j.message = 'Excel を起動しています…';
  } else if (ev.type === 'progress') {
    if (ev.phase === 'excel') {
      j.progress = 0.1;
      j.message = 'ブックを開いています…';
    } else if (ev.phase === 'open') {
      j.progress = 0.15;
      j.message = j.job.ops.length ? `操作を再現しています（0/${j.job.ops.length}）…` : '再計算しています…';
    } else if (ev.phase === 'op') {
      const n = Number(ev.n) || j.job.ops.length || 1;
      const i = Number(ev.i) + 1;
      j.progress = 0.15 + 0.7 * Math.min(1, i / n);
      j.message = `操作を再現しています（${i}/${n}: ${OP_TEXT[ev.op] || ev.op}）…`;
      if (i >= n) j.message = '再計算しています…';
    } else if (ev.phase === 'calc') {
      j.progress = 0.9;
      j.message = '保存しています…';
    } else if (ev.phase === 'save') {
      j.progress = 0.97;
      j.message = PHASE_TEXT.save;
    }
  } else if (ev.type === 'warn') {
    if (ev.message && !j.warnings.includes(String(ev.message))) j.warnings.push(String(ev.message));
  }
}

/** Runners (powershell.exe + our Excel) that are still shutting down after their job was delivered: { reap, done }. */
const lingering = new Set();

/**
 * The job is over but powershell.exe / Excel may still be shutting down (seconds normally; ~60 s when Excel's cloud services
 * stall - the reason results are delivered early). Afterwards: delete Excel's private TEMP (Office diagnostics) and, if the job
 * was dropped meanwhile, its folder. shutdown() kills whatever is still here.
 */
function trackLingering(j, r) {
  const cleanup = async () => {
    await removeDir(j.xlTemp);
    if (!jobs.has(j.id)) await removeDir(j.dir);
  };
  if (!r || !r.closed) {
    cleanup().catch(() => {});
    return;
  }
  const entry = { reap: () => (typeof r.reap === 'function' ? r.reap() : Promise.resolve()), done: null };
  lingering.add(entry);
  entry.done = r.closed
    .then((c) => {
      if (c && c.excelKilled) log(`ジョブ ${j.id.slice(0, 8)}: Excel の終了が遅れたため、このジョブの Excel を終了させました（出力は保存済みです）`);
    })
    .catch(() => {})
    .then(cleanup)
    .catch(() => {})
    .finally(() => lingering.delete(entry));
}

/** The output of a finished runner: present, not empty, a complete package (the ZIP directory at its end is readable). */
async function verifyOutput(file) {
  const st = await statOrNull(file);
  if (!st || st.size === 0) return { error: 'Excel本体エンジンが出力ファイルを作りませんでした' };
  const info = await inspectPackageFile(file);
  if (info.error) return { error: `Excel本体エンジンの出力ファイルが正しくありません（${info.error}）` };
  return { ok: true, size: st.size };
}

async function pump() {
  if (runningJob || shuttingDown) return;
  const id = queue.shift();
  if (!id) return;
  const j = jobs.get(id);
  if (!j || j.state !== 'queued') return pump();
  runningJob = j;
  j.state = 'running';
  j.progress = 0.01;
  j.message = 'Excel を起動しています…';
  j.startedAt = Date.now();
  for (const other of queue) {
    const q = jobs.get(other);
    if (q) q.message = '順番待ちです（前の処理が終わると始まります）';
  }
  log(`ジョブ ${id.slice(0, 8)}: Excel本体で処理を開始します`);
  let result = null;
  try {
    const runner = resolveRunner();
    if (!runner) throw Object.assign(new Error('Excel本体エンジン (excel-runner.ps1) が見つかりません'), { code: 'RUNNER_MISSING' });
    const disk = await diskReport();
    if (disk.diskLow) log(`警告: ${disk.diskMessage}`);
    // Excel's own TEMP (Office diagnostics, ~80 MB per start, never cleaned by Office) is a private folder of this job, deleted afterwards
    j.xlTemp = path.join(TEMP_ROOT, `xl-${id.slice(0, 8)}-${process.pid}`);
    await fsp.mkdir(j.xlTemp, { recursive: true });
    const jobFile = path.join(j.dir, 'job.json');
    // input = <job folder>\work.xlsx (opened in place), output = <job folder>\output.xlsx (Excel saves straight into it): workDir = the job folder
    const full = { ...j.job, input: j.inputFile, output: j.outputFile, workDir: j.dir };
    await fsp.writeFile(jobFile, JSON.stringify(full), 'utf8');
    const r = await runRunner({
      runner,
      jobFile,
      signal: j.controller.signal,
      onEvent: (ev) => onRunnerEvent(j, ev),
      timeoutMs: envInt('JOB_TIMEOUT_MS', 30 * 60_000, 1000),
      idleMs: envInt('JOB_IDLE_MS', 5 * 60_000, 1000),
      invoke: envStr('XLRUNNER_INVOKE') || 'auto',
      earlyResult: true,
      exitGraceMs: EXCEL_EXIT_GRACE_MS,
      extraArgs: [{ name: 'QuitWaitMs', value: '0' }],
      env: { XLRUNNER_CACHE: path.join(DATA_DIR, 'cache'), TEMP: j.xlTemp, TMP: j.xlTemp },
    });
    result = r;
    for (const w of r.warnings || []) {
      const m = w && typeof w === 'object' ? w.message : w;
      if (m && !j.warnings.includes(String(m))) j.warnings.push(String(m));
    }
    if (j.controller.signal.aborted || r.cancelled) {
      j.state = 'error';
      j.errorCode = 'CANCELLED';
      j.error = '中止しました';
      j.message = '中止しました';
    } else if (r.ok) {
      const v = await verifyOutput(j.outputFile);
      if (v.error) throw Object.assign(new Error(v.error), { code: 'NO_OUTPUT' });
      j.state = 'done';
      j.progress = 1;
      j.message = '完了しました';
      log(`ジョブ ${id.slice(0, 8)}: 完了（${((Date.now() - j.startedAt) / 1000).toFixed(1)} 秒）`);
    } else {
      j.state = 'error';
      j.errorCode = r.code || 'RUNNER_FAILED';
      j.error = friendlyRunnerError(r);
      j.message = j.error;
      if (['SAVE_FAILED', 'OPEN_FAILED', 'EXCEL_REFUSED', 'OP_FAILED', 'INTERNAL', 'RUNNER_CRASHED'].includes(j.errorCode)) {
        const low = await diskLowError(0);
        if (low) {
          j.errorCode = 'DISK_LOW';
          j.error = low.message;
          j.message = j.error;
        }
      }
      log(`ジョブ ${id.slice(0, 8)}: 失敗 ${j.errorCode}: ${j.error}`);
    }
  } catch (e) {
    j.state = 'error';
    j.errorCode = (e && e.code) || 'SERVER_ERROR';
    j.error = e && e.code === 'ENOENT' ? 'PowerShell を起動できませんでした（Windows PowerShell が見つかりません）' : String((e && e.message) || e);
    j.message = j.error;
    log(`ジョブ ${id.slice(0, 8)}: 失敗 ${j.errorCode}: ${j.error}`);
  } finally {
    j.finishedAt = Date.now();
    runningJob = null;
    // Disk hygiene (%TEMP% is usually the system drive): a failed / cancelled job loses its whole folder now; a finished one
    // keeps only the output for the result download (upload copy and job file go at once). Excel's private TEMP goes when it has exited.
    if (j.state === 'done') cleanJobFolder(j, [j.outputFile]).catch(() => {});
    else removeDir(j.dir).catch(() => {});
    trackLingering(j, result);
    if (runningDoneWaiters.length) runningDoneWaiters.splice(0).forEach((f) => f());
    setImmediate(pump);
  }
}
const runningDoneWaiters = [];

function friendlyRunnerError(r) {
  const msg = String(r.message || '').trim();
  switch (r.code) {
    case 'EXCEL_UNAVAILABLE':
      return `Excel を起動できませんでした。内蔵エンジンで出力してください。${msg ? `（${msg}）` : ''}`;
    case 'PASSWORD_PROTECTED':
      return 'パスワード付きのブックは Excel本体モードで処理できません。内蔵エンジンで出力してください。';
    case 'MOVE_VERIFY_FAILED':
      // the runner already says so; make sure the hint is there whatever runner version produced the message
      return !msg ? `この列の移動は Excel本体では正しく行えませんでした（保存していません）。${BUILT_IN_HINT}` : /内蔵エンジン/.test(msg) ? msg : `${msg}${BUILT_IN_HINT}`;
    case 'TIMEOUT':
      return msg || 'Excel の処理が時間内に終わらなかったため中止しました。';
    default:
      return msg || `Excel本体エンジンでエラーが発生しました (${r.code || '不明'})`;
  }
}

/** Removes a folder inside TEMP_ROOT (Excel may keep a file open for a moment after Quit(): retried for ~5 s). */
async function removeDir(dir) {
  if (!dir || !isInside(TEMP_ROOT, dir)) return;
  for (let i = 0; i < 10; i++) {
    try {
      await fsp.rm(dir, { recursive: true, force: true });
      return;
    } catch {
      await new Promise((r) => setTimeout(r, 500));
    }
  }
}

/** Deletes everything in the job folder except the listed files (the output of a finished job). */
async function cleanJobFolder(j, keep) {
  let entries;
  try {
    entries = await fsp.readdir(j.dir);
  } catch {
    return;
  }
  const keepSet = new Set((keep || []).map((p) => path.resolve(p).toLowerCase()));
  for (const name of entries) {
    const p = path.join(j.dir, name);
    if (keepSet.has(path.resolve(p).toLowerCase())) continue;
    await removeDir(p);
  }
}

async function cancelJob(j) {
  if (j.state === 'queued') {
    const pos = queue.indexOf(j.id);
    if (pos >= 0) queue.splice(pos, 1);
    j.state = 'error';
    j.errorCode = 'CANCELLED';
    j.error = '中止しました';
    j.message = '中止しました';
    j.finishedAt = Date.now();
  } else if (j.state === 'running') {
    j.controller.abort();
    await new Promise((resolve) => {
      const t = setTimeout(resolve, 15000);
      runningDoneWaiters.push(() => {
        clearTimeout(t);
        resolve();
      });
    });
  }
}

// drop finished jobs (and their temp folders) after a while
setInterval(() => {
  const now = Date.now();
  for (const j of jobs.values()) {
    const stale = (j.finishedAt && now - j.finishedAt > JOB_KEEP_MS) || (!j.inputReady && !j.receiving && j.state === 'queued' && now - j.createdAt > JOB_KEEP_MS);
    if (stale) {
      jobs.delete(j.id);
      removeDir(j.dir).catch(() => {});
    }
  }
}, 60_000).unref();

/**
 * Removes stale job folders at server start (crash / killed server / a job whose cleanup failed): a folder whose server process is
 * gone after 1 hour, ANY job folder after 1 day (a live server drops its jobs after 15 minutes; a pid can also be reused by an
 * unrelated program, so "alive" proves nothing after a day). Only folders named job-<8 hex>-<pid> (a job) or xl-<8 hex>-<pid>
 * (Excel's private TEMP of a job) inside TEMP_ROOT are touched. Returns the names it removed.
 */
const DEAD_OWNER_DIR_MS = 60 * 60 * 1000;
const STALE_JOB_DIR_MS = 24 * 60 * 60 * 1000;
async function sweepTempRoot() {
  const removed = [];
  let entries;
  try {
    entries = await fsp.readdir(TEMP_ROOT, { withFileTypes: true });
  } catch {
    return removed;
  }
  for (const e of entries) {
    if (!e.isDirectory() || !/^(job|xl)-[0-9a-f]{8}-\d+$/.test(e.name)) continue;
    const pid = Number(e.name.split('-').pop());
    if (pid === process.pid) continue;
    let alive = false;
    try {
      process.kill(pid, 0);
      alive = true;
    } catch (err) {
      alive = !!(err && err.code === 'EPERM'); // exists, but belongs to someone else
    }
    const dir = path.join(TEMP_ROOT, e.name);
    const st = await statOrNull(dir);
    if (!st) continue;
    const age = Date.now() - st.mtimeMs;
    if (age > STALE_JOB_DIR_MS || (!alive && age > DEAD_OWNER_DIR_MS)) {
      await removeDir(dir);
      removed.push(e.name);
    }
  }
  if (removed.length) log(`古い一時フォルダを ${removed.length} 件削除しました`);
  return removed;
}

// ---------------------------------------------------------------------------------------------------------
// Excel detection (registry only, cached)
// ---------------------------------------------------------------------------------------------------------
let excelInfoPromise = null;
function excelInfo() {
  if (!excelInfoPromise) {
    excelInfoPromise = detectExcel()
      .then((r) => {
        const out = { installed: !!r.installed };
        if (r.installed && r.version) out.version = String(r.version);
        return out;
      })
      .catch(() => ({ installed: false }));
  }
  return excelInfoPromise;
}

// ---------------------------------------------------------------------------------------------------------
// App window (Edge --app with its own profile -> Chrome --app -> default browser)
// ---------------------------------------------------------------------------------------------------------
function regDefault(key) {
  try {
    const out = execFileSync(path.join(envStr('SystemRoot') || 'C:\\Windows', 'System32', 'reg.exe'), ['query', key, '/ve'], {
      encoding: 'utf8',
      windowsHide: true,
      stdio: ['ignore', 'pipe', 'ignore'],
      timeout: 5000,
    });
    const m = out.match(/REG_(?:EXPAND_)?SZ\s+(.+?)\s*$/m);
    if (!m) return null;
    return m[1].trim().replace(/^"|"$/g, '').replace(/%([^%]+)%/g, (_, v) => process.env[v] ?? `%${v}%`);
  } catch {
    return null;
  }
}

function findChromium(kind) {
  const exe = kind === 'edge' ? 'msedge.exe' : 'chrome.exe';
  const sub = kind === 'edge' ? ['Microsoft', 'Edge', 'Application', exe] : ['Google', 'Chrome', 'Application', exe];
  const candidates = [
    regDefault(`HKCU\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\${exe}`),
    regDefault(`HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\${exe}`),
    regDefault(`HKLM\\SOFTWARE\\WOW6432Node\\Microsoft\\Windows\\CurrentVersion\\App Paths\\${exe}`),
    process.env['ProgramFiles(x86)'] && path.join(process.env['ProgramFiles(x86)'], ...sub),
    process.env.ProgramFiles && path.join(process.env.ProgramFiles, ...sub),
    process.env.LOCALAPPDATA && path.join(process.env.LOCALAPPDATA, ...sub),
  ];
  for (const c of candidates) if (c && fs.existsSync(c)) return c;
  return null;
}

/** Returns the kind of window opened ('edge' | 'chrome' | 'browser') or null. */
function openAppWindow(url) {
  if (APP_MODE === 'none' || process.platform !== 'win32') return null;
  const order = APP_MODE === 'auto' ? ['edge', 'chrome', 'browser'] : [APP_MODE];
  for (const kind of order) {
    try {
      if (kind === 'browser') {
        const c = spawn(path.join(envStr('SystemRoot') || 'C:\\Windows', 'System32', 'cmd.exe'), ['/d', '/s', '/c', `start "" "${url}"`], {
          stdio: 'ignore',
          detached: true,
          windowsHide: true,
          windowsVerbatimArguments: true,
        });
        c.on('error', () => log('ブラウザを開けませんでした。上の URL を手で開いてください。'));
        c.unref();
        return kind;
      }
      const exe = findChromium(kind);
      if (!exe) continue;
      // Own profile: the user's tabs/extensions are untouched, the window size is honoured and the app's
      // localStorage stays in one place (%LOCALAPPDATA%\ExcelMerger\edge-profile).
      const profile = path.join(DATA_DIR, `${kind}-profile`);
      const c = spawn(
        exe,
        [`--app=${url}`, `--user-data-dir=${profile}`, '--no-first-run', '--no-default-browser-check', '--disable-sync', '--window-size=1400,900'],
        { stdio: 'ignore', detached: true, windowsHide: false },
      );
      c.on('error', (e) => log(`${kind} を起動できませんでした: ${e && e.message}`));
      c.unref();
      return kind;
    } catch (e) {
      log(`${kind} を起動できませんでした: ${e && e.message}`);
    }
  }
  return null;
}

// ---------------------------------------------------------------------------------------------------------
// Lifetime: SSE connections + pagehide beacon (R5 §3.5)
// ---------------------------------------------------------------------------------------------------------
const clients = new Set();
let byeAt = 0;
let exitTimer = null;
let shuttingDown = false;

function armExitTimer(ms, why, closedOnPurpose = false) {
  clearTimeout(exitTimer);
  exitTimer = null;
  if (!(ms > 0)) return;
  exitTimer = setTimeout(() => {
    if (clients.size > 0) return;
    // Without a bye (no window, API-only use) a job in progress keeps the server alive.
    if (!closedOnPurpose && (runningJob || queue.length > 0)) return armExitTimer(ms, why, false);
    shutdown(why);
  }, ms);
}

function scheduleExitIfIdle(reason) {
  if (clients.size > 0 || shuttingDown) return;
  const closedOnPurpose = Date.now() - byeAt < 15000;
  const wait = closedOnPurpose ? EXIT_GRACE_MS : IDLE_EXIT_MS;
  if (closedOnPurpose) log(`ウィンドウが閉じられました。${Math.round(wait / 1000)} 秒以内に再接続がなければ終了します（${reason}）`);
  else if (wait > 0) log(`ウィンドウとの接続が切れました。${Math.round(wait / 60000)} 分間再接続がなければ終了します（${reason}）`);
  armExitTimer(wait, closedOnPurpose ? 'ウィンドウが閉じられたため終了します' : '長時間接続がなかったため終了します', closedOnPurpose);
}

function onEvents(req, res) {
  if (clients.size >= MAX_EVENT_CLIENTS) return sendError(res, 429, '接続が多すぎます（ウィンドウを閉じてからやり直してください）');
  res.writeHead(200, {
    'Content-Type': 'text/event-stream; charset=utf-8',
    'Cache-Control': 'no-store',
    Connection: 'keep-alive',
    'X-Accel-Buffering': 'no',
    ...COMMON_HEADERS,
  });
  res.write('retry: 1000\n\n');
  res.write(`event: hello\ndata: ${JSON.stringify({ app: APP_ID, version: VERSION, pid: process.pid })}\n\n`);
  clients.add(res);
  clearTimeout(exitTimer);
  exitTimer = null;
  const ping = setInterval(() => {
    try {
      res.write(': ping\n\n');
    } catch {
      /* closed */
    }
  }, 20000);
  req.on('close', () => {
    clearInterval(ping);
    clients.delete(res);
    scheduleExitIfIdle('接続終了');
  });
}

async function shutdown(why) {
  if (shuttingDown) return;
  shuttingDown = true;
  log(why || '終了します');
  // Cancel queued jobs; stop a running one (kills only that job's own Excel PID) and wait for it.
  queue.length = 0;
  if (runningJob) {
    runningJob.controller.abort();
    await new Promise((resolve) => {
      const t = setTimeout(resolve, 12000);
      runningDoneWaiters.push(() => {
        clearTimeout(t);
        resolve();
      });
    });
  }
  // Runners whose job was already delivered but whose powershell.exe / Excel are still shutting down: end them now (our own PIDs only)
  for (const l of lingering) l.reap().catch(() => {});
  if (lingering.size) await Promise.race([Promise.all([...lingering].map((l) => l.done)), new Promise((resolve) => setTimeout(resolve, 8000))]);
  for (const j of jobs.values()) await removeDir(j.dir).catch(() => {});
  process.exit(0);
}

// ---------------------------------------------------------------------------------------------------------
// HTTP routing
// ---------------------------------------------------------------------------------------------------------
function sendBuildHint(res) {
  const html = `<!DOCTYPE html><html lang="ja"><head><meta charset="utf-8"><title>${APP_TITLE}</title>
<style>body{font-family:"Yu Gothic UI",Meiryo,sans-serif;margin:40px;line-height:1.8;color:#222}code{background:#eee;padding:2px 6px;border-radius:4px}</style></head>
<body><h1>画面ファイル (dist/index.html) が見つかりません</h1>
<p>アプリがまだビルドされていません。フォルダ内の <code>build.bat</code> を実行するか、<code>npm install</code> → <code>npm run build</code> を実行してから、もう一度 <code>start.bat</code> を起動してください。</p></body></html>`;
  res.writeHead(503, { 'Content-Type': 'text/html; charset=utf-8', 'Content-Length': Buffer.byteLength(html), 'Cache-Control': 'no-store', ...HTML_HEADERS });
  res.end(html);
}

async function handleStatic(req, res, pathname) {
  const method = req.method || 'GET';
  if (method !== 'GET' && method !== 'HEAD') return sendError(res, 405, `許可されていないメソッドです: ${method}`);
  const headOnly = method === 'HEAD';
  const rel = safeDecode(pathname);
  if (rel !== '/' && rel !== '/index.html' && !rel.includes('\0') && !(IS_SEA && !envStr('DIST_DIR'))) {
    const full = path.resolve(DIST_DIR, `.${rel}`);
    if (isInside(DIST_DIR, full)) {
      const st = await statOrNull(full);
      if (st && st.isFile()) {
        res.writeHead(200, { 'Content-Type': contentTypeOf(full), 'Content-Length': st.size, 'Cache-Control': 'no-cache', ...(/\.html?$/i.test(full) ? HTML_HEADERS : COMMON_HEADERS) });
        if (headOnly) return res.end();
        const s = fs.createReadStream(full);
        s.on('error', () => res.destroy());
        return s.pipe(res);
      }
    }
    if (/\.[a-z0-9]{1,8}$/i.test(rel)) return sendError(res, 404, `見つかりません: ${rel}`);
  }
  const body = await readIndexHtml();
  if (!body) return sendBuildHint(res);
  res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8', 'Content-Length': body.length, 'Cache-Control': 'no-cache', ...HTML_HEADERS });
  return res.end(headOnly ? undefined : body);
}

const CONTENT_TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.svg': 'image/svg+xml',
  '.png': 'image/png',
  '.ico': 'image/x-icon',
  '.woff2': 'font/woff2',
  '.txt': 'text/plain; charset=utf-8',
};
const contentTypeOf = (f) => CONTENT_TYPES[path.extname(f).toLowerCase()] || 'application/octet-stream';
const XLSX_TYPES = {
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  xlsm: 'application/vnd.ms-excel.sheet.macroEnabled.12',
};

function needMethod(req, ...methods) {
  const m = req.method || 'GET';
  if (!methods.includes(m)) throw new HttpError(405, `${methods.join(' / ')} のみ対応しています`);
}

async function route(req, res) {
  let url;
  try {
    url = new URL(req.url || '/', 'http://localhost');
  } catch {
    return sendError(res, 400, 'URL が不正です');
  }
  const method = req.method || 'GET';
  const p = url.pathname.replace(/\/+$/, '') || '/';
  const isApi = p === '/api' || p.startsWith('/api/');

  if (!hostAllowed(req.headers.host)) {
    req.resume();
    return sendError(res, 403, `このサーバーは http://localhost:${listeningPort || BASE_PORT}/ からのみ利用できます`);
  }
  // writes of any kind, and every API read, must come from this server's own page (static pages may be navigated to from anywhere)
  if ((isApi || (method !== 'GET' && method !== 'HEAD')) && crossSiteBlocked(req)) {
    req.resume(); // read (and drop) the body: the client gets the answer instead of a reset connection
    return sendError(res, 403, '他のサイトのページからの操作は拒否しました');
  }
  if (method === 'OPTIONS') {
    req.resume();
    return sendError(res, 405, 'OPTIONS には対応していません');
  }

  if (p === '/api/ping') {
    needMethod(req, 'GET', 'HEAD');
    return sendJson(res, 200, { app: APP_ID, version: VERSION, pid: process.pid });
  }
  if (p === '/api/status') {
    needMethod(req, 'GET', 'HEAD');
    return sendJson(res, 200, {
      app: APP_ID,
      version: VERSION,
      excel: await excelInfo(),
      outputDir,
      jobs: { running: runningJob ? 1 : 0, queued: queue.length },
      windows: clients.size,
      sea: IS_SEA,
      ...(await diskReport()),
    });
  }
  if (p === '/api/events') {
    needMethod(req, 'GET');
    return onEvents(req, res);
  }
  if (p === '/api/bye') {
    needMethod(req, 'POST');
    byeAt = Date.now();
    req.resume();
    res.writeHead(204, { 'Cache-Control': 'no-store', ...COMMON_HEADERS });
    res.end();
    scheduleExitIfIdle('bye');
    return;
  }
  if (p === '/api/save') {
    needMethod(req, 'POST', 'PUT');
    const cleaned = cleanOutputName(headerValue(req.headers['x-file-name']));
    if (cleaned.error) {
      req.resume();
      return sendError(res, 400, cleaned.error);
    }
    if (Number(req.headers['content-length']) > MAX_BODY_BYTES) {
      req.resume();
      throw tooLarge(MAX_BODY_BYTES);
    }
    try {
      let saved;
      try {
        saved = await saveUnique(req, cleaned.name, MAX_BODY_BYTES);
      } catch (e) {
        // app folder not writable (Program Files, read-only share ...): use Documents\ExcelMerger\output
        const denied = e && ['EPERM', 'EACCES', 'EROFS'].includes(e.code);
        if (!denied || outputDir === FALLBACK_OUTPUT_DIR || envStr('OUTPUT_DIR')) throw e;
        log(`出力フォルダに書き込めません (${e.code}): ${outputDir} -> ${FALLBACK_OUTPUT_DIR} を使います`);
        outputDir = FALLBACK_OUTPUT_DIR;
        saved = await saveUnique(req, cleaned.name, MAX_BODY_BYTES);
      }
      log(`保存しました: ${saved.path}`);
      return sendJson(res, 200, saved);
    } catch (e) {
      if (e instanceof HttpError) throw e;
      if (e && e.code === 'ENOSPC') throw new HttpError(507, DISK_FULL_MESSAGE);
      req.resume();
      return sendError(res, 500, `保存できませんでした (${(e && e.code) || (e && e.message) || e})。出力フォルダに書き込めるか確認してください: ${outputDir}`);
    }
  }
  if (p === '/api/open') {
    needMethod(req, 'POST');
    const body = await readJson(req, 64 * 1024);
    const found = await resolveOutputFile(body && body.path);
    if (found.error) return sendError(res, found.status, found.error);
    const opened = shellOpen(found.file);
    return sendJson(res, 200, { ok: true, opened });
  }
  if (p === '/api/open-folder') {
    needMethod(req, 'POST');
    const body = await readBody(req, 64 * 1024);
    let select = null;
    if (body.length) {
      try {
        const j = JSON.parse(body.toString('utf8'));
        if (j && typeof j.path === 'string') {
          const found = await resolveOutputFile(j.path);
          if (found.file) select = found.file;
        }
      } catch {
        /* no selection */
      }
    }
    const dir = await ensureOutputDir();
    const opened = select ? shellOpen(select, true) : shellOpen(dir);
    return sendJson(res, 200, { ok: true, opened, path: dir });
  }
  if (p === '/api/jobs') {
    needMethod(req, 'POST');
    return createJob(req, res);
  }
  const m = /^\/api\/jobs\/([0-9a-f-]{36})(?:\/(input|result))?$/.exec(p);
  if (m) {
    const j = jobs.get(m[1]);
    if (!j) {
      req.resume();
      return sendError(res, 404, 'ジョブが見つかりません（期限切れの可能性があります）');
    }
    if (m[2] === 'input') {
      needMethod(req, 'PUT', 'POST');
      return putJobInput(req, res, j);
    }
    if (m[2] === 'result') {
      needMethod(req, 'GET');
      if (j.state !== 'done') return sendError(res, 409, j.state === 'error' ? j.error || '失敗しました' : 'まだ終わっていません');
      const st = await statOrNull(j.outputFile);
      if (!st) return sendError(res, 410, '出力ファイルがありません（期限切れ）');
      res.writeHead(200, { 'Content-Type': XLSX_TYPES[j.job.format], 'Content-Length': st.size, 'Cache-Control': 'no-store', ...COMMON_HEADERS });
      const s = fs.createReadStream(j.outputFile);
      s.on('error', () => res.destroy());
      return s.pipe(res);
    }
    if (method === 'DELETE') {
      await cancelJob(j);
      jobs.delete(j.id);
      await removeDir(j.dir);
      return sendJson(res, 200, { ok: true });
    }
    needMethod(req, 'GET', 'HEAD');
    return sendJson(res, 200, jobView(j));
  }
  if (isApi) return sendError(res, 404, `API が見つかりません: ${p}`);
  return handleStatic(req, res, url.pathname);
}

function onRequest(req, res) {
  route(req, res).catch((err) => {
    if (err instanceof HttpError) {
      if (!res.headersSent) sendError(res, err.status, err.message, err.extra);
      else res.destroy();
      req.resume();
      return;
    }
    log(`内部エラー: ${(err && err.stack) || err}`);
    if (!res.headersSent) sendError(res, 500, 'サーバー内部エラーが発生しました');
    else res.destroy();
  });
}

// ---------------------------------------------------------------------------------------------------------
// Start: single instance, port selection, window
// ---------------------------------------------------------------------------------------------------------
function pingExisting(port) {
  return new Promise((resolve) => {
    const req = http.get({ host: '127.0.0.1', port, path: '/api/ping', timeout: 2000, headers: { Host: `localhost:${port}` } }, (res) => {
      const chunks = [];
      res.on('data', (d) => chunks.push(d));
      res.on('end', () => {
        try {
          resolve(JSON.parse(Buffer.concat(chunks).toString('utf8')).app === APP_ID);
        } catch {
          resolve(false);
        }
      });
    });
    req.on('error', () => resolve(false));
    req.on('timeout', () => {
      req.destroy();
      resolve(false);
    });
  });
}

function notifyParent(msg) {
  if (typeof process.send === 'function') {
    try {
      process.send(msg);
      process.disconnect();
    } catch {
      /* launcher gone */
    }
  }
}

function listen(port, attempt) {
  const server = http.createServer(onRequest);
  server.requestTimeout = 0; // large uploads on slow disks; SSE is long-lived anyway
  server.maxConnections = 512; // the window opens a handful (SSE + polling); this only bounds a runaway local client
  server.headersTimeout = 60_000;
  server.on('error', async (err) => {
    if (err && err.code === 'EADDRINUSE' && BASE_PORT !== 0 && attempt < MAX_PORT_TRIES) {
      if (await pingExisting(port)) {
        const url = `http://localhost:${port}/`;
        console.log(`${APP_TITLE} は既に起動しています: ${url}`);
        const kind = openAppWindow(url);
        if (kind) console.log('新しいウィンドウを開きました。');
        notifyParent({ ready: true, url, reused: true, window: kind });
        setTimeout(() => process.exit(0), 300);
        return;
      }
      log(`ポート ${port} は別のプログラムが使用中です。${port + 1} を試します…`);
      listen(port + 1, attempt + 1);
      return;
    }
    console.error(`[エラー] 起動できませんでした（${HOST}:${port}: ${(err && (err.code || err.message)) || err}）`);
    notifyParent({ ready: false, error: String((err && (err.code || err.message)) || err) });
    process.exit(1);
  });
  server.listen(port, HOST, () => {
    listeningPort = server.address().port;
    const url = `http://localhost:${listeningPort}/`;
    console.log(`READY ${url}`);
    if (HOST_IGNORED) log(`環境変数 HOST=${HOST_IGNORED} は無視しました（このツールは、このパソコンの中（${HOST}）でしか待ち受けません）`);
    log(`${APP_TITLE} v${VERSION} を起動しました: ${url}（${Date.now() - T0} ms${IS_SEA ? ', exe' : ''}）`);
    log(`出力フォルダ: ${outputDir}`);
    if (process.platform === 'win32') process.title = `${APP_TITLE} - ${url}`;
    const kind = openAppWindow(url);
    if (kind) log(kind === 'browser' ? '既定のブラウザで開きました' : `${kind === 'edge' ? 'Microsoft Edge' : 'Google Chrome'} のアプリウィンドウで開きました`);
    else if (APP_MODE === 'none') log('ウィンドウは自動で開きません（NO_OPEN）。上の URL を開いてください。');
    log('ウィンドウを閉じると、このツールも自動で終了します。');
    notifyParent({ ready: true, url, window: kind });
    // Nobody ever connects (browser failed to open, window closed before loading): do not linger forever.
    armExitTimer(IDLE_EXIT_MS, 'ウィンドウが接続されなかったため終了します');
    excelInfo().then((x) => log(x.installed ? `Excel本体: 利用できます${x.version ? `（${x.version}）` : ''}` : 'Excel本体: 見つかりません（内蔵エンジンで出力します）'));
    // clear Japanese warning (server.log / console; the same text is in /api/status for the UI) when the drive of the temp folder is nearly full
    diskReport()
      .then((d) => {
        if (d.diskLow) log(`警告: ${d.diskMessage}`);
      })
      .catch(() => {});
    sweepTempRoot().catch(() => {});
  });
}

// --detach: the visible console process starts a hidden copy of itself (log file), waits for READY, exits.
function launcherMain() {
  const logDir = path.join(DATA_DIR, 'logs');
  let fd = 'ignore';
  let logFile = '';
  try {
    fs.mkdirSync(logDir, { recursive: true });
    logFile = path.join(logDir, 'server.log');
    try {
      if (fs.statSync(logFile).size > 1024 * 1024) fs.renameSync(logFile, `${logFile}.1`);
    } catch {
      /* no log yet */
    }
    fd = fs.openSync(logFile, 'a');
  } catch {
    fd = 'ignore';
  }
  console.log(`${APP_TITLE} を起動しています…`);
  const args = IS_SEA ? [] : [process.argv[1]];
  let child;
  try {
    child = spawn(process.execPath, args, {
      detached: true,
      windowsHide: true,
      stdio: ['ignore', fd, fd, 'ipc'],
      env: { ...process.env, EM_CHILD: '1' },
    });
  } catch (e) {
    fail(e && e.message);
    return;
  }
  const t = Date.now();
  let done = false;
  function fail(why) {
    if (done) return;
    done = true;
    console.error(`[エラー] 起動できませんでした（${why}）。`);
    if (logFile) console.error(`        ログ: ${logFile}`);
    console.error('        start.bat debug で起動すると、詳しいメッセージがこのウィンドウに表示されます。');
    process.exitCode = 1;
    if (process.stdin.isTTY && process.env.EM_FROM_BAT !== '1') {
      console.error('Enter キーで閉じます。');
      process.stdin.once('data', () => process.exit(1));
    } else setTimeout(() => process.exit(1), 100);
  }
  const timer = setTimeout(() => fail('60 秒以内に起動しませんでした'), 60000);
  child.on('message', (m) => {
    if (!m || done) return;
    if (!m.ready) {
      clearTimeout(timer);
      fail(m.error || '不明なエラー');
      return;
    }
    done = true;
    clearTimeout(timer);
    console.log(m.reused ? `既に起動しています: ${m.url}` : `起動しました: ${m.url}（${Date.now() - t} ms）`);
    console.log('このウィンドウは自動で閉じます。アプリのウィンドウを閉じると、ツールも終了します。');
    try {
      child.unref();
      child.disconnect();
    } catch {
      /* already */
    }
    setTimeout(() => process.exit(0), 1500);
  });
  child.on('exit', (code) => {
    clearTimeout(timer);
    fail(`終了コード ${code}`);
  });
  child.on('error', (e) => {
    clearTimeout(timer);
    fail(e && e.message);
  });
}

for (const sig of ['SIGINT', 'SIGTERM', 'SIGBREAK', 'SIGHUP']) {
  try {
    process.on(sig, () => shutdown('終了の指示を受けたため終了します'));
  } catch {
    /* signal not supported here */
  }
}
process.on('uncaughtException', (e) => {
  log(`予期しないエラー: ${(e && e.stack) || e}`);
});

if (process.argv.includes('--detach') && process.env.EM_CHILD !== '1') {
  launcherMain();
} else {
  log(`${APP_TITLE} v${VERSION}  node ${process.version}${IS_SEA ? '  (exe)' : ''}`);
  listen(BASE_PORT, 0);
}
