/**
 * safe.mjs - the input-hygiene helpers of server.mjs as pure functions (no I/O, no side effects), so that tests can throw
 * hostile input at them directly (tests/server/security.test.ts). Node >= 18, built-in modules only; bundled with the server.
 *
 *   cleanOutputName(header)        x-file-name -> a safe base name (.xlsx / .xlsm) or a Japanese error
 *   textCheckOpenPath(raw)         first, textual screening of the path given to /api/open and /api/open-folder
 *   isInside(dir, file)            file lies strictly inside dir (Windows: case-insensitive)
 *   buildOpenCommand(target, ...)  how a file / folder is opened: never through a shell string (no cmd.exe, no path in a command line)
 *   validateRunnerOps(ops)         whitelist of the op log a job may carry: only known ops, only known fields, typed
 */
import path from 'node:path';

export const OUTPUT_EXTS = new Set(['.xlsx', '.xlsm']);

export function safeDecode(s) {
  try {
    return decodeURIComponent(s);
  } catch {
    return s;
  }
}

/**
 * Windows device names. "CON.tar.xlsx", "NUL .xlsx" and "COM1.anything" are the device as well (only the part before the first dot
 * counts, trailing spaces are ignored), the superscript digits COM¹ COM² COM³ LPT¹ ... are reserved too, and so are CONIN$ / CONOUT$.
 */
const WIN_RESERVED = /^(con|prn|aux|nul|clock\$|conin\$|conout\$|com[0-9\u00b9\u00b2\u00b3]|lpt[0-9\u00b9\u00b2\u00b3])$/i;
export function isReservedDeviceName(stem) {
  return WIN_RESERVED.test(String(stem).split('.')[0].replace(/[ \u3000]+$/, ''));
}

/**
 * Output file name from the x-file-name header: a plain base name (no folders), Windows-safe characters,
 * .xlsx / .xlsm only. Returns { name } or { error } (Japanese).
 */
export function cleanOutputName(raw) {
  let s = typeof raw === 'string' ? raw : '';
  s = safeDecode(s);
  // A raw UTF-8 name that travelled through a Latin-1 header comes out as mojibake: undo it.
  if (/[\x80-\xff]/.test(s) && !/[^\x00-\xff]/.test(s)) {
    const utf8 = Buffer.from(s, 'latin1').toString('utf8');
    if (!utf8.includes('\ufffd')) s = utf8;
  }
  s = s.normalize('NFC').trim();
  if (!s) return { error: 'ファイル名がありません（x-file-name ヘッダー）' };
  if (/[\\/]/.test(s) || s.includes('..') || /^[a-zA-Z]:/.test(s)) return { error: 'ファイル名にフォルダ（\\ / ..）は指定できません' };
  s = s
    .replace(/[\x00-\x1f\x7f<>:"|?*]/g, '_') // illegal in Windows names (':' also closes alternate data streams: name.xlsx:stream)
    .replace(/[\u00ad\u200b-\u200f\u2028-\u202e\u2060-\u2069\ufeff]/g, '') // invisible / bidi-control characters (name spoofing)
    .replace(/[\ud800-\udbff](?![\udc00-\udfff])|(?<![\ud800-\udbff])[\udc00-\udfff]/g, '_') // lone surrogates
    .replace(/[. \u3000]+$/, ''); // Windows drops trailing dots / spaces
  const ext = path.extname(s).toLowerCase();
  if (!OUTPUT_EXTS.has(ext)) return { error: `保存できるのは .xlsx / .xlsm ファイルだけです: ${s}` };
  let stem = s.slice(0, s.length - ext.length).replace(/^[ \u3000]+/, '');
  if (!stem.trim()) stem = '出力';
  if (stem.length > 150) stem = stem.slice(0, 150).replace(/[\ud800-\udbff]$/, '');
  if (isReservedDeviceName(stem)) stem = `_${stem}`;
  return { name: stem + ext };
}

/** true when `file` is inside `dir` (case-insensitive on Windows). */
export function isInside(dir, file) {
  const rel = path.relative(path.resolve(dir), path.resolve(file));
  return rel !== '' && !rel.startsWith('..') && !path.isAbsolute(rel);
}

/**
 * Textual screening of a path that /api/open or /api/open-folder should act on. Returns null when it may go on to the
 * real-path checks, else { status, error } (Japanese). Refused here: not a string, control characters, UNC / device paths
 * (\\server\share, \\?\C:\..., \\.\pipe\...), alternate data streams (any ':' but the drive letter's), over-long paths.
 */
export function textCheckOpenPath(raw) {
  if (typeof raw !== 'string' || raw.length === 0) return { status: 400, error: 'path がありません' };
  if (raw.length > 1000 || /[\x00-\x1f\x7f]/.test(raw)) return { status: 400, error: 'path が不正です' };
  if (/^[\\/]{2}/.test(raw)) return { status: 403, error: 'ネットワークパス・デバイスパスは開けません（出力フォルダの中のファイルだけを開けます）' };
  if (raw.replace(/^[a-zA-Z]:/, '').includes(':')) return { status: 403, error: 'このパスは開けません（出力フォルダの中のファイルだけを開けます）' };
  return null;
}

/**
 * Command line that opens `target` (an existing file inside the output folder, or the output folder itself) with its default
 * program, or selects it in Explorer. NEVER a shell string:
 *   - select:   explorer.exe /select,"<path>" (verbatim; the path cannot contain a double quote on Windows and is refused here if it does)
 *   - otherwise powershell.exe ... -Command "Invoke-Item -LiteralPath $env:EM_OPEN_TARGET": the path travels in an ENVIRONMENT
 *     VARIABLE, so no quoting rule (' ` $ ( ) ; & and the typographic quotes U+2018 / U+2019 PowerShell also treats as apostrophes)
 *     can ever turn a file name into code.
 * Returns { cmd, args, env, verbatim }.
 */
export function buildOpenCommand(target, select, sysRoot) {
  const root = sysRoot || 'C:\\Windows';
  const t = String(target);
  if (/[\x00-\x1f"]/.test(t)) throw new Error('unsafe path');
  if (select) return { cmd: path.join(root, 'explorer.exe'), args: [`/select,"${t}"`], env: {}, verbatim: true };
  return {
    cmd: path.join(root, 'System32', 'WindowsPowerShell', 'v1.0', 'powershell.exe'),
    args: ['-NoLogo', '-NoProfile', '-NonInteractive', '-Command', 'Invoke-Item -LiteralPath $env:EM_OPEN_TARGET'],
    env: { EM_OPEN_TARGET: t },
    verbatim: false,
  };
}

// ---------------------------------------------------------------------------------------------------------
// Job op whitelist (the op log the UI sends is what excel-runner.ps1 replays: nothing else may get through)
// ---------------------------------------------------------------------------------------------------------
const A1 = /^\$?[A-Za-z]{1,3}\$?[0-9]{1,7}$/;
const A1_RANGE = /^\$?[A-Za-z]{1,3}\$?[0-9]{1,7}(:\$?[A-Za-z]{1,3}\$?[0-9]{1,7})?$/;
const MAX_LINE = 1048576;
const MAX_CELLS = 4_000_000;
const CELL_KEYS = new Set(['f', 'dyn', 'e', 't', 'v', 's', 'n', 'quote', 'numFmt', 'skip']);
const METHODS = new Set(['clipFree', 'cutInsert', 'insertCut', 'auto']);
const FORMAT_FROM = new Set(['left', 'right', 'above', 'below', 'none']);
const TEXT_MODES = new Set(['auto', 'prefixAll', 'raw']);
const UNITS = new Set(['chars', 'stored']);

const isInt = (v, min = 1, max = MAX_LINE) => typeof v === 'number' && Number.isInteger(v) && v >= min && v <= max;
const isStr = (v, max) => typeof v === 'string' && v.length <= max;

function validateCell(v) {
  if (v === null || typeof v === 'string' || typeof v === 'boolean') return true;
  if (typeof v === 'number') return Number.isFinite(v);
  if (typeof v !== 'object' || Array.isArray(v)) return false;
  for (const [k, x] of Object.entries(v)) {
    if (!CELL_KEYS.has(k)) return false;
    if (x !== null && typeof x === 'object') return false;
    if (typeof x === 'number' && !Number.isFinite(x)) return false;
    if ((k === 'numFmt' || k === 't' || k === 'e') && x !== null && !isStr(x, 255)) return false;
  }
  return true;
}

/** One op -> a clean copy (only the known fields) or { error }. Positions are checked like the runner does (1-based, sheet limits). */
function cleanOp(op, i, allowDebug) {
  const at = `ops[${i}]`;
  const kind = op.op;
  const sheet = op.sheet;
  // test hook (XLRUNNER_ALLOW_DEBUG=1 only); afterResult: the fake runner keeps running after its result, like a powershell.exe that cannot exit
  if (kind === 'debugWait') return allowDebug ? { op: 'debugWait', seconds: Number(op.seconds) || 1, ...(op.afterResult === true ? { afterResult: true } : {}) } : { error: `未知の操作です: ${String(kind)} (${at})` };
  const needSheet = () => (isStr(sheet, 255) && sheet.length > 0 && !/[\x00-\x1f]/.test(sheet) ? null : `シート名が不正です (${at})`);
  let err;
  switch (kind) {
    case 'moveCols':
    case 'moveRows': {
      if ((err = needSheet())) return { error: err };
      const method = op.method === undefined ? 'clipFree' : op.method;
      if (!isInt(op.start) || !isInt(op.count) || !isInt(op.to)) return { error: `${kind} の start / count / to が不正です (${at})` };
      if (!METHODS.has(method)) return { error: `${kind} の method が不正です (${at})` };
      return { op: kind, sheet, start: op.start, count: op.count, to: op.to, method };
    }
    case 'insertCols':
    case 'insertRows': {
      if ((err = needSheet())) return { error: err };
      const ff = op.formatFrom === undefined ? (kind === 'insertCols' ? 'left' : 'above') : op.formatFrom;
      if (!isInt(op.at) || !isInt(op.count)) return { error: `${kind} の at / count が不正です (${at})` };
      if (!FORMAT_FROM.has(ff)) return { error: `${kind} の formatFrom が不正です (${at})` };
      return { op: kind, sheet, at: op.at, count: op.count, formatFrom: ff };
    }
    case 'deleteCols':
    case 'deleteRows': {
      if ((err = needSheet())) return { error: err };
      if (!isInt(op.at) || !isInt(op.count)) return { error: `${kind} の at / count が不正です (${at})` };
      return { op: kind, sheet, at: op.at, count: op.count };
    }
    case 'setCells': {
      if ((err = needSheet())) return { error: err };
      if (typeof op.at !== 'string' || !A1.test(op.at)) return { error: `setCells の at はセル番地（A1 形式）にしてください (${at})` };
      const rows = op.rows === undefined ? [] : op.rows;
      if (!Array.isArray(rows)) return { error: `setCells の rows が配列ではありません (${at})` };
      let cells = 0;
      for (const row of rows) {
        if (row === null) continue;
        if (!Array.isArray(row)) return { error: `setCells の rows が二重配列ではありません (${at})` };
        cells += row.length;
        if (cells > MAX_CELLS) return { error: `setCells のセル数が多すぎます (${at})` };
        for (const c of row) if (!validateCell(c)) return { error: `setCells に書き込めない形式のセルがあります (${at})` };
      }
      const out = { op: 'setCells', sheet, at: op.at, rows };
      if (op.textMode !== undefined) {
        if (!TEXT_MODES.has(op.textMode)) return { error: `setCells の textMode が不正です (${at})` };
        out.textMode = op.textMode;
      }
      if (op.toValues === true) out.toValues = true;
      return out;
    }
    case 'setColumnWidth': {
      if ((err = needSheet())) return { error: err };
      const unit = op.unit === undefined ? 'chars' : op.unit;
      const count = op.count === undefined ? 1 : op.count;
      if (!isInt(op.col, 1, 16384) || !isInt(count, 1, 16384)) return { error: `setColumnWidth の col / count が不正です (${at})` };
      if (typeof op.width !== 'number' || !Number.isFinite(op.width) || op.width < 0 || op.width > 256) return { error: `setColumnWidth の width が不正です (${at})` };
      if (!UNITS.has(unit)) return { error: `setColumnWidth の unit が不正です (${at})` };
      return { op: 'setColumnWidth', sheet, col: op.col, count, width: op.width, unit };
    }
    case 'unmerge': {
      if ((err = needSheet())) return { error: err };
      if (!Array.isArray(op.refs) || op.refs.length > 100000 || !op.refs.every((r) => typeof r === 'string' && A1_RANGE.test(r))) return { error: `unmerge の refs はセル範囲（A1:B2 形式）の配列にしてください (${at})` };
      return { op: 'unmerge', sheet, refs: op.refs, centerAcross: op.centerAcross === true };
    }
    default:
      return { error: `未知の操作です: ${String(kind)} (${at})` };
  }
}

/** Returns { ops } (clean copies) or { error } (Japanese). */
export function validateRunnerOps(ops, { allowDebug = false } = {}) {
  const out = [];
  for (let i = 0; i < ops.length; i++) {
    const op = ops[i];
    if (!op || typeof op !== 'object' || Array.isArray(op)) return { error: `ops[${i}] がオブジェクトではありません` };
    const c = cleanOp(op, i, allowDebug);
    if (c.error) return { error: c.error };
    out.push(c);
  }
  return { ops: out };
}
