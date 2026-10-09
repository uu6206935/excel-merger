/**
 * Client for the local server (server/server.mjs). Owner: server/COM (S).
 *
 * Every function fails soft in standalone mode (dist/index.html opened from file://, or no server):
 *   getServerStatus() -> null, startHeartbeat() -> no-op; the others throw an Error with a Japanese message.
 * Errors from the server carry its Japanese message (`{ error }` bodies); `code` is set when known
 * (e.g. 'CANCELLED', 'EXCEL_UNAVAILABLE', 'TIMEOUT', 'NO_SERVER').
 */
import type { ExcelJobResult, ProgressFn, RunnerJob } from './core/contracts';

export interface ServerStatus {
  app: 'excel-merger';
  version: string;
  excel: { installed: boolean; version?: string };
  outputDir: string;
  /** true when less than 2 GB are free on the drive of the server's temp folder (Excel本体 exports need temporary space) */
  diskLow?: boolean;
  /** free space of that drive in MB (undefined = unknown) */
  diskFreeMB?: number;
  /** Japanese warning text to show when diskLow is true */
  diskMessage?: string;
  /** true when less than 2 GB are free on the drive of the output folder */
  outputDiskLow?: boolean;
}

export interface ApiError extends Error {
  code?: string;
  status?: number;
  /**
   * The Excel本体 engine could not do this job but the built-in engine can (e.g. code MOVE_VERIFY_FAILED: a column moved inside
   * an Excel table with cells above / below it). The UI retries with the built-in engine for these (see canFallBackToJs).
   */
  fallbackToJs?: boolean;
}

/** Codes of a failed Excel job after which the built-in engine is the right answer (mirrors the server's FALLBACK_TO_JS_CODES). */
const FALLBACK_TO_JS_CODES = ['MOVE_VERIFY_FAILED', 'EXCEL_UNAVAILABLE', 'PASSWORD_PROTECTED', 'OPEN_FAILED', 'RUNNER_MISSING', 'DISK_LOW'];

/** true when an error from runExcelJob says "use the built-in engine instead" (never for a user cancel). */
export function canFallBackToJs(e: unknown): boolean {
  if (!e || typeof e !== 'object') return false;
  const x = e as ApiError;
  if (x.code === 'CANCELLED' || x.name === 'AbortError') return false;
  return x.fallbackToJs === true || (typeof x.code === 'string' && FALLBACK_TO_JS_CODES.includes(x.code));
}

const APP_ID = 'excel-merger';

function apiError(message: string, code?: string, status?: number): ApiError {
  const e = new Error(message) as ApiError;
  if (code) e.code = code;
  if (status) e.status = status;
  if (code === 'CANCELLED') e.name = 'AbortError';
  return e;
}

/** true when the page was served by an http(s) origin (i.e. possibly by our server). */
function hasOrigin(): boolean {
  try {
    return typeof location !== 'undefined' && (location.protocol === 'http:' || location.protocol === 'https:') && typeof fetch === 'function';
  } catch {
    return false;
  }
}

const NO_SERVER = 'ローカルサーバーに接続できません。start.bat からツールを起動し直してください。';

async function request(path: string, init: RequestInit = {}, timeoutMs = 0): Promise<Response> {
  if (!hasOrigin()) throw apiError('スタンドアロンモードではこの機能は使えません（start.bat から起動してください）', 'NO_SERVER');
  const ctrl = new AbortController();
  const outer = init.signal;
  const onAbort = () => ctrl.abort();
  if (outer) {
    if (outer.aborted) ctrl.abort();
    else outer.addEventListener('abort', onAbort, { once: true });
  }
  const timer = timeoutMs > 0 ? setTimeout(() => ctrl.abort(), timeoutMs) : null;
  try {
    return await fetch(path, { cache: 'no-store', credentials: 'same-origin', ...init, signal: ctrl.signal });
  } catch (e) {
    if (outer?.aborted) throw apiError('中止しました', 'CANCELLED');
    if (ctrl.signal.aborted) throw apiError('ローカルサーバーが応答しません。しばらく待ってからやり直してください。', 'TIMEOUT');
    throw apiError(NO_SERVER, 'NO_SERVER');
  } finally {
    if (timer) clearTimeout(timer);
    outer?.removeEventListener('abort', onAbort);
  }
}

/** Throws the server's Japanese `{ error }` message for a non-2xx response. */
async function ensureOk(res: Response, fallback: string): Promise<Response> {
  if (res.ok) return res;
  let msg = '';
  let code: string | undefined;
  try {
    const j = (await res.json()) as { error?: unknown; errorCode?: unknown };
    if (typeof j.error === 'string') msg = j.error;
    if (typeof j.errorCode === 'string') code = j.errorCode;
  } catch {
    /* not JSON */
  }
  throw apiError(msg || `${fallback}（HTTP ${res.status}）`, code ?? `HTTP_${res.status}`, res.status);
}

async function readJsonBody<T>(res: Response): Promise<T> {
  try {
    return (await res.json()) as T;
  } catch {
    throw apiError('ローカルサーバーから不正な応答がありました', 'BAD_RESPONSE');
  }
}

function bytesBody(bytes: Uint8Array, type: string): Blob {
  // Copy only when the view does not cover a plain ArrayBuffer exactly (Blob wants ArrayBuffer-backed parts).
  const plain = bytes.buffer instanceof ArrayBuffer && bytes.byteOffset === 0 && bytes.byteLength === bytes.buffer.byteLength;
  const part = plain ? (bytes.buffer as ArrayBuffer) : bytes.slice().buffer;
  return new Blob([part], { type });
}

// ---------------------------------------------------------------------------------------------------------

/** null = standalone (file:// or no server). Never throws. */
export async function getServerStatus(): Promise<ServerStatus | null> {
  if (!hasOrigin()) return null;
  try {
    const res = await request('/api/status', {}, 5000);
    if (!res.ok) return null;
    const j = (await res.json()) as Partial<ServerStatus> | null;
    if (!j || j.app !== APP_ID) return null; // e.g. the Vite dev server answering with index.html
    return {
      app: APP_ID,
      version: String(j.version ?? ''),
      excel: { installed: !!j.excel?.installed, ...(j.excel?.version ? { version: String(j.excel.version) } : {}) },
      outputDir: String(j.outputDir ?? ''),
      ...(typeof j.diskLow === 'boolean' ? { diskLow: j.diskLow } : {}),
      ...(typeof j.diskFreeMB === 'number' ? { diskFreeMB: j.diskFreeMB } : {}),
      ...(typeof j.diskMessage === 'string' && j.diskMessage ? { diskMessage: j.diskMessage } : {}),
      ...(typeof j.outputDiskLow === 'boolean' ? { outputDiskLow: j.outputDiskLow } : {}),
    };
  } catch {
    return null;
  }
}

interface JobView {
  id: string;
  state: 'queued' | 'running' | 'done' | 'error';
  progress: number | null;
  message: string;
  error?: string;
  errorCode?: string;
  fallbackToJs?: boolean;
  warnings?: string[];
}

const sleep = (ms: number, signal?: AbortSignal) =>
  new Promise<void>((resolve) => {
    const t = setTimeout(done, ms);
    function done() {
      clearTimeout(t);
      signal?.removeEventListener('abort', done);
      resolve();
    }
    signal?.addEventListener('abort', done, { once: true });
  });

/** Japanese runner warnings of a finished job (deduplicated, non-empty strings only). */
function jobWarnings(v: JobView | null): string[] {
  const out: string[] = [];
  const list: unknown[] = v && Array.isArray(v.warnings) ? v.warnings : [];
  for (const w of list) {
    const m = typeof w === 'string' ? w.trim() : '';
    if (m && !out.includes(m)) out.push(m);
  }
  return out;
}

/**
 * Runs the op log in the user's own Excel (Excel本体 engine) on the premerged package and returns the saved
 * workbook plus the runner's Japanese warnings (MACROS_DROPPED, REF_ERRORS, MOVE_RETRY, CLIPBOARD_*, BAD_FORMULA …).
 * Cancelling via `signal` stops the job on the server (its own Excel is closed) and rejects with an AbortError
 * ('中止しました').
 */
export async function runExcelJob(premerged: Uint8Array, job: RunnerJob, onProgress?: ProgressFn, signal?: AbortSignal): Promise<ExcelJobResult> {
  const report = (ratio: number | null, message: string) => {
    try {
      onProgress?.({ ratio, message });
    } catch {
      /* UI callback errors must not break the job */
    }
  };
  if (signal?.aborted) throw apiError('中止しました', 'CANCELLED');
  report(0, 'Excel本体エンジンを準備しています…');
  const { input: _in, output: _out, ...clean } = job;
  void _in;
  void _out;
  const created = await ensureOk(
    await request('/api/jobs', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(clean), signal }, 60000),
    'Excel本体エンジンにジョブを登録できませんでした',
  );
  const { id } = await readJsonBody<{ id: string }>(created);
  const jobUrl = `/api/jobs/${encodeURIComponent(id)}`;
  const cancel = () => {
    // fire and forget; keepalive lets it through even while the page is going away
    fetch(jobUrl, { method: 'DELETE', keepalive: true }).catch(() => {});
  };
  try {
    report(0.01, 'ブックを送っています…');
    await ensureOk(
      await request(`${jobUrl}/input`, { method: 'PUT', body: bytesBody(premerged, 'application/octet-stream'), signal }),
      'ブックを Excel本体エンジンに送れませんでした',
    );
    let delay = 300;
    let lastMsg = '';
    let final: JobView | null = null;
    for (;;) {
      if (signal?.aborted) throw apiError('中止しました', 'CANCELLED');
      const res = await ensureOk(await request(jobUrl, { signal }, 15000), 'Excel本体エンジンの状態を取得できませんでした');
      const v = await readJsonBody<JobView>(res);
      if (v.state === 'done') {
        final = v;
        break;
      }
      if (v.state === 'error') {
        const err = apiError(v.error || v.message || 'Excel本体エンジンでエラーが発生しました', v.errorCode || 'JOB_FAILED');
        if (v.fallbackToJs) err.fallbackToJs = true;
        throw err;
      }
      if (v.message !== lastMsg || v.state === 'running') report(typeof v.progress === 'number' ? v.progress : null, v.message);
      lastMsg = v.message;
      await sleep(delay, signal);
      delay = Math.min(1000, delay + 100);
    }
    report(0.99, '結果を受け取っています…');
    const res = await ensureOk(await request(`${jobUrl}/result`, { signal }), '結果を受け取れませんでした');
    const bytes = new Uint8Array(await res.arrayBuffer());
    report(1, '完了しました');
    return { bytes, warnings: jobWarnings(final) };
  } catch (e) {
    const err = e as ApiError;
    if (signal?.aborted || err?.code === 'CANCELLED') {
      cancel();
      throw apiError('中止しました', 'CANCELLED');
    }
    throw e;
  } finally {
    // the server drops finished jobs by itself after a while; removing it now frees the temp files
    if (!signal?.aborted) cancel();
  }
}

/** Saves bytes into the server's output folder. Never overwrites: 'name (2).xlsx' … Returns the final path/name. */
export async function saveToOutput(bytes: Uint8Array, fileName: string): Promise<{ path: string; name: string }> {
  const res = await ensureOk(
    await request('/api/save', {
      method: 'POST',
      headers: { 'x-file-name': encodeURIComponent(fileName) },
      body: bytesBody(bytes, 'application/octet-stream'),
    }),
    '保存できませんでした',
  );
  const j = await readJsonBody<{ path: string; name: string }>(res);
  return { path: String(j.path), name: String(j.name) };
}

/** Opens an output file with Excel (only files inside the output folder are allowed by the server). */
export async function openFile(path: string): Promise<void> {
  await ensureOk(
    await request('/api/open', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ path }) }, 15000),
    'ファイルを開けませんでした',
  );
}

/** Opens the output folder in Explorer. */
export async function openOutputFolder(): Promise<void> {
  await ensureOk(await request('/api/open-folder', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: '{}' }, 15000), 'フォルダを開けませんでした');
}

let heartbeat: EventSource | null = null;

/**
 * Keeps the server alive while this window is open (SSE /api/events) and tells it when the window closes
 * (pagehide beacon /api/bye), so the server exits a few seconds after the window is closed. Idempotent.
 * Fires `window` event 'excel-merger-server' with detail { connected: boolean } on connection changes.
 */
export function startHeartbeat(): void {
  if (heartbeat || !hasOrigin() || typeof EventSource === 'undefined') return;

  const notify = (connected: boolean) => {
    try {
      window.dispatchEvent(new CustomEvent('excel-merger-server', { detail: { connected } }));
    } catch {
      /* no window */
    }
  };
  try {
    heartbeat = new EventSource('/api/events');
    heartbeat.addEventListener('open', () => notify(true));
    heartbeat.addEventListener('error', () => notify(false)); // EventSource reconnects by itself (retry: 1000)
  } catch {
    heartbeat = null;
    return;
  }
  window.addEventListener('pagehide', (e: PageTransitionEvent) => {
    if (e.persisted) return;
    try {
      navigator.sendBeacon('/api/bye', '');
    } catch {
      /* ignore */
    }
  });
}
