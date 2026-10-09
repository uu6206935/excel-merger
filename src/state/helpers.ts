/**
 * Small pure helpers used by the UI (no React, no DOM access at import time).
 */
import type { ServerStatus } from '../api';
import type { NormalizeOptions } from '../core/contracts';
import type { Cell, SheetGrid, SourceBook, SourceFormat, TableRegion } from '../core/types';

// ---------------------------------------------------------------------------------------------
// colours

/** Source colours (card strip + grid header band). White text must stay readable on them. */
export const SOURCE_COLORS = ['#2563eb', '#15803d', '#c2410c', '#7c3aed', '#be185d', '#0f766e', '#a16207', '#4338ca'] as const;
/** Header band colour of columns added in the app. */
export const APP_SOURCE_COLOR = '#6b7280';

/** First palette colour not used yet (cycles when all are taken). */
export function pickColor(used: readonly string[]): string {
  for (const c of SOURCE_COLORS) if (!used.includes(c)) return c;
  return SOURCE_COLORS[used.length % SOURCE_COLORS.length];
}

// ---------------------------------------------------------------------------------------------
// files

export const ACCEPTED_EXTENSIONS = ['.xlsx', '.xlsm', '.xls', '.xlsb', '.csv', '.ods'] as const;
export const ACCEPT_ATTR = ACCEPTED_EXTENSIONS.join(',');
export const UNSUPPORTED_FILE_MESSAGE = '対応していない形式です（.xlsx / .xlsm / .xls / .xlsb / .csv / .ods を選んでください）';

export function extOf(name: string): string {
  const i = name.lastIndexOf('.');
  return i >= 0 ? name.slice(i).toLowerCase() : '';
}

export function isAcceptedFileName(name: string): boolean {
  return (ACCEPTED_EXTENSIONS as readonly string[]).includes(extOf(name));
}

/** Excel's lock file of an open workbook ("~$名簿.xlsx") — never a real input. */
export function isExcelLockFile(name: string): boolean {
  return name.startsWith('~$');
}

export function stripExt(name: string): string {
  const i = name.lastIndexOf('.');
  return i > 0 ? name.slice(0, i) : name;
}

export const FORMAT_LABEL: Record<SourceFormat, string> = {
  xlsx: 'xlsx',
  xlsm: 'xlsm',
  xls: 'xls',
  xlsb: 'xlsb',
  csv: 'csv',
  ods: 'ods',
  other: 'その他',
};

/** Display name of a detected text encoding ('SJIS' / 'CP932' -> 'Shift_JIS', 'UTF8' -> 'UTF-8' …). */
export function encodingLabel(enc: string | undefined): string | null {
  if (!enc || !enc.trim()) return null;
  const k = enc.replace(/[-_\s]/g, '').toUpperCase();
  if (k === 'SJIS' || k === 'SHIFTJIS' || k === 'CP932' || k === 'MS932' || k === 'WINDOWS31J' || k === 'XSJIS') return 'Shift_JIS';
  if (k === 'UTF8') return 'UTF-8';
  if (k === 'UTF8BOM') return 'UTF-8（BOM付き）';
  if (k === 'EUCJP') return 'EUC-JP';
  if (k === 'UTF16' || k === 'UTF16LE' || k === 'UTF16BE') return 'UTF-16';
  if (k === 'JIS' || k === 'ISO2022JP') return 'JIS';
  if (k === 'ASCII') return 'ASCII';
  return enc.trim();
}

/** Human readable byte size ('12.3 KB'). */
export function formatSize(bytes: number): string {
  if (!Number.isFinite(bytes) || bytes < 0) return '';
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

/** Error -> message text (falls back to String()). */
export function errorText(e: unknown): string {
  if (e instanceof Error) return e.message || e.name;
  if (typeof e === 'string') return e;
  if (e === null || e === undefined) return '不明なエラー';
  if (typeof e === 'object') {
    const m = (e as { message?: unknown }).message;
    return typeof m === 'string' && m ? m : '不明なエラー';
  }
  try {
    return String(e);
  } catch {
    return '不明なエラー';
  }
}

/** True when the text contains hiragana, katakana or kanji (a message written for the user, not a library's English one). */
export function hasJapanese(s: string): boolean {
  return /[぀-ヿ一-鿿]/.test(s);
}

/**
 * Unexpected errors come from libraries / the browser and are usually English ('Cannot read properties of undefined').
 * The user sees a Japanese sentence; the raw text (for a bug report) is returned separately as `detail`.
 * Messages that are already Japanese (the modules' and the server's) pass through unchanged.
 */
export function friendlyError(e: unknown): { text: string; detail?: string } {
  const raw = errorText(e).trim();
  if (!raw) return { text: '原因のわからないエラーが発生しました' };
  if (hasJapanese(raw)) return { text: raw };
  if (/out of memory|allocation failed|invalid (array|string|typed array) length|array buffer allocation/i.test(raw)) {
    return { text: 'メモリが足りません。ファイルが大きすぎるか、同時に開いているファイルが多すぎる可能性があります', detail: raw };
  }
  if (/failed to fetch|networkerror|network request failed|load failed/i.test(raw)) {
    return { text: 'ローカルサーバーと通信できませんでした', detail: raw };
  }
  return { text: '予期しないエラーが発生しました', detail: raw };
}

/** One-line form of friendlyError for places without a tooltip (card / dialog error text): the raw text stays as a short technical note. */
export function errorLine(e: unknown): string {
  const f = friendlyError(e);
  return f.detail ? `${f.text}［技術情報: ${f.detail.length > 120 ? f.detail.slice(0, 120) + '…' : f.detail}］` : f.text;
}

/** Library errors are often English; keep the original but lead with a Japanese sentence. */
export function readErrorMessage(fileName: string, e: unknown): string {
  const raw = errorText(e);
  if (/password|encrypt|暗号|パスワード/i.test(raw)) return `${fileName}: パスワードで保護されているため読み込めません。Excelで保護を解除して保存し直してください`;
  if (hasJapanese(raw)) return `${fileName}: ${raw}`;
  return `${fileName}: 読み込めませんでした（ファイルが壊れているか、対応していない形式です）［${raw}］`;
}

/** 'YYYYMMDD' in local time. */
export function dateStamp(now: Date = new Date()): string {
  const p = (n: number) => String(n).padStart(2, '0');
  return `${now.getFullYear()}${p(now.getMonth() + 1)}${p(now.getDate())}`;
}

/** Characters Windows does not allow in file names -> '_'. */
export function sanitizeFileName(name: string): string {
  // eslint-disable-next-line no-control-regex
  let s = name.replace(/[\\/:*?"<>|\u0000-\u001f]/g, '_').replace(/[. ]+$/, '').trim();
  if (/^\.+$/.test(s)) s = '';
  // same rules as the server (cleanOutputName): reserved device names get a leading '_', at most 150 characters
  if (/^(con|prn|aux|nul|com[0-9]|lpt[0-9])$/i.test(s)) s = `_${s}`;
  if (s.length > 150) s = s.slice(0, 150).replace(/[. ]+$/, '');
  return s || '統合結果';
}

/** Default export name: <基準ファイル名>_統合_<yyyyMMdd>.xlsx */
export function exportFileName(baseFileName: string, format: 'xlsx' | 'xlsm', now: Date = new Date()): string {
  return `${sanitizeFileName(stripExt(baseFileName))}_統合_${dateStamp(now)}.${format}`;
}

/** Normalises the name the user typed: sanitised, extension forced to the output format. */
export function normalizeExportName(input: string, format: 'xlsx' | 'xlsm'): string {
  let s = input.trim();
  const ext = extOf(s);
  if (ext === '.xlsx' || ext === '.xlsm' || ext === '.xls') s = s.slice(0, -ext.length);
  return `${sanitizeFileName(s)}.${format}`;
}

// ---------------------------------------------------------------------------------------------
// options

export const DEFAULT_OPTIONS: NormalizeOptions = {
  itaiji: true,
  leadingZeros: true,
  honorifics: true,
  maidenName: true,
  universityAbbrev: true,
};

export const OPTION_LABELS: Array<{ key: keyof NormalizeOptions; label: string; title: string }> = [
  { key: 'itaiji', label: '異体字を同一視', title: '髙/高・﨑/崎・邊/辺 などを同じ字として照合します' },
  { key: 'leadingZeros', label: '先頭ゼロを無視', title: '000123 と 123 を同じIDとして照合します' },
  { key: 'honorifics', label: '敬称を除去', title: '「様」「さん」「殿」などを取り除いて照合します' },
  { key: 'maidenName', label: '旧姓の括弧を除去', title: '「山田 花子（旧姓：佐藤）」の括弧部分を取り除いて照合します' },
  { key: 'universityAbbrev', label: '大学名の表記ゆれを統一', title: '「早大」「早稲田大」などの略称・表記ゆれを統一して照合します' },
];

// ---------------------------------------------------------------------------------------------
// columns / widths

/** 1-based column number -> Excel letters (1 -> A, 27 -> AA). */
export function colLetter(n: number): string {
  let s = '';
  let x = Math.floor(n);
  while (x > 0) {
    const r = (x - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    x = Math.floor((x - 1) / 26);
  }
  return s;
}

// ---------------------------------------------------------------------------------------------
// sheets / regions

export function sheetByName(book: SourceBook, name: string | undefined): SheetGrid | undefined {
  return name === undefined ? undefined : book.sheets.find((s) => s.name === name);
}

export function cellAt(sheet: SheetGrid, row: number, col: number): Cell | undefined {
  return sheet.rows[row - 1]?.[col - 1];
}

export function isBlankCell(c: Cell | undefined): boolean {
  if (!c) return true;
  if (c.f) return false;
  return c.v === null || c.v === '' || c.t === 'z';
}

function rowIsBlank(sheet: SheetGrid, row: number, c1: number, c2: number): boolean {
  const r = sheet.rows[row - 1];
  if (!r) return true;
  for (let c = c1; c <= c2; c++) if (!isBlankCell(r[c - 1])) return false;
  return true;
}

/** Choices for the 見出し行 select: 1 .. min(maxRow, 30), always including the current one. */
export function headerRowChoices(sheet: SheetGrid | undefined, current?: number): number[] {
  const max = Math.max(1, Math.min(sheet?.maxRow ?? 1, 30));
  const out: number[] = [];
  for (let r = 1; r <= max; r++) out.push(r);
  if (current !== undefined && current > max) out.push(current);
  return out;
}

/**
 * Local FALLBACK for the data region when the user picks another header row (the controller asks
 * keys.detectRegion(book, sheet, { headerRow }) first and uses this only when that throws / ignores the row).
 * Columns = first..last non-blank header cell (else the detected ones); data rows run from the next
 * row down to the detected last data row when it lies below, else until two blank rows in a row.
 */
export function regionWithHeaderRow(book: SourceBook, detected: TableRegion, headerRow: number): TableRegion {
  const sheet = sheetByName(book, detected.sheet);
  if (!sheet || headerRow === detected.headerRow) return { ...detected };
  const hdr = sheet.rows[headerRow - 1] ?? [];
  let firstCol = 0;
  let lastCol = 0;
  const maxCol = Math.max(sheet.maxCol, hdr.length);
  for (let c = 1; c <= maxCol; c++) {
    if (!isBlankCell(hdr[c - 1])) {
      if (!firstCol) firstCol = c;
      lastCol = c;
    }
  }
  if (!firstCol) {
    firstCol = detected.firstCol;
    lastCol = detected.lastCol;
  }
  const firstDataRow = headerRow + 1;
  let lastDataRow: number;
  if (detected.lastDataRow > headerRow) {
    lastDataRow = detected.lastDataRow;
  } else {
    lastDataRow = headerRow;
    let blankRun = 0;
    for (let r = firstDataRow; r <= sheet.maxRow; r++) {
      if (rowIsBlank(sheet, r, firstCol, lastCol)) {
        blankRun++;
        if (blankRun >= 2) break;
      } else {
        blankRun = 0;
        lastDataRow = r;
      }
    }
  }
  return { sheet: detected.sheet, headerRow, firstDataRow, lastDataRow: Math.max(lastDataRow, headerRow), firstCol, lastCol };
}

/**
 * A table without a header row (only numbers, e.g. raw survey data) is read with its first row as the header: that row is not
 * matched / joined (it stays in the output). True when the header row holds at least two cells and every one is a number / date,
 * so the card can say so.
 */
export function headerRowIsNumeric(book: SourceBook, region: TableRegion | undefined): boolean {
  if (!region) return false;
  const sheet = sheetByName(book, region.sheet);
  const row = sheet?.rows[region.headerRow - 1];
  if (!row) return false;
  let n = 0;
  for (let c = region.firstCol; c <= region.lastCol; c++) {
    const cell = row[c - 1];
    if (isBlankCell(cell)) continue;
    if (!cell || (cell.t !== 'n' && cell.t !== 'd')) return false;
    n++;
  }
  return n >= 2;
}

export function regionRowCount(region: TableRegion | undefined): number {
  if (!region) return 0;
  return Math.max(0, region.lastDataRow - region.firstDataRow + 1);
}

/** Allowed values of the 最終行 input: the first data row .. the last used row of the sheet. */
export function lastRowBounds(sheet: SheetGrid | undefined, region: TableRegion | undefined): { min: number; max: number } | null {
  if (!sheet || !region) return null;
  const min = region.firstDataRow;
  const max = Math.max(min, sheet.maxRow, region.lastDataRow);
  return { min, max };
}

/** Local fallback of detectRegion(…, { lastDataRow }): the same table, cut (or extended) at the chosen row. */
export function regionWithLastRow(region: TableRegion, lastDataRow: number, maxRow: number): TableRegion {
  const hi = Math.max(region.firstDataRow, maxRow);
  const last = Math.max(region.firstDataRow, Math.min(hi, Math.round(lastDataRow)));
  return { ...region, lastDataRow: last };
}

// ---------------------------------------------------------------------------------------------
// messages

/**
 * Short, friendly toast text for an operation the editor refused (doc reasons are long and exact): the first
 * sentence without the 「（Excel と同じ制限です）」 style remarks, 「…することはできません」 → 「…できません」.
 * The full reason is returned as `detail` (tooltip) when it says more.
 */
export function refusalMessage(reason: string | undefined, fallback: string): { text: string; detail?: string } {
  const full = (reason ?? '').trim();
  if (!full) return { text: fallback };
  const cleaned = full
    .replace(/（Excel[^（）]*）/g, '')
    .replace(/\(Excel[^()]*\)/g, '')
    .trim();
  let first = cleaned.split('。')[0].trim() || cleaned;
  first = first.replace(/することはできません$/, 'できません');
  return first === full ? { text: first } : { text: first, detail: full };
}

/**
 * Toast for the warnings of an operation that succeeded (doc OpResult.warnings, often several long sentences:
 * 「タイトル行の結合（A1:K1）を解除し…」「VLOOKUP の列番号は…」「集計!B28: =VLOOKUP(…)」): the first sentence of
 * the first warning + 「ほか N 件の注意」; every warning, one per line, as `detail` (「詳しく」).
 */
export function opWarningsMessage(warnings: readonly string[] | undefined): { text: string; detail?: string } | null {
  const list = (warnings ?? []).map((w) => (typeof w === 'string' ? w.trim() : '')).filter(Boolean);
  if (!list.length) return null;
  const first = list[0];
  const cut = first.indexOf('。');
  let text = cut >= 0 && cut < first.length - 1 ? first.slice(0, cut + 1) : first;
  if (list.length > 1) text += `（ほか ${list.length - 1} 件の注意）`;
  const detail = list.join('\n');
  return text === detail ? { text } : { text, detail };
}

// ---------------------------------------------------------------------------------------------
// keyboard

export interface KeyLike {
  key: string;
  ctrlKey: boolean;
  metaKey: boolean;
  shiftKey: boolean;
  altKey: boolean;
}

/** 'undo' for Ctrl/Cmd+Z, 'redo' for Ctrl/Cmd+Y or Ctrl/Cmd+Shift+Z, else null. */
export function undoRedoAction(e: KeyLike): 'undo' | 'redo' | null {
  if (!(e.ctrlKey || e.metaKey) || e.altKey) return null;
  const k = e.key.toLowerCase();
  if (k === 'z') return e.shiftKey ? 'redo' : 'undo';
  if (k === 'y' && !e.shiftKey) return 'redo';
  return null;
}

/** True for inputs where Ctrl+Z must keep its native text-editing meaning. */
export function isTextEditingTarget(target: unknown): boolean {
  if (typeof HTMLElement === 'undefined' || !(target instanceof HTMLElement)) return false;
  const tag = target.tagName;
  if (tag === 'INPUT' || tag === 'TEXTAREA' || tag === 'SELECT') return true;
  return target.isContentEditable;
}

// ---------------------------------------------------------------------------------------------
// low disk space (GET /api/status: diskLow / diskFreeMB / diskMessage / outputDiskLow)

/** Text of the small amber chip in the header. */
export const DISK_LOW_CHIP_TEXT = '空き容量が少ない';

/** What the UI shows when the server reports that a drive is nearly full (see diskNotice). */
export interface DiskNotice {
  /** the drive of the server's temp folder is nearly full: Excel本体 exports may fail, the built-in engine is the safe choice */
  tempLow: boolean;
  /** the drive of the output folder is nearly full */
  outputLow: boolean;
  /** one text for the header chip (tooltip + aria-label): the server's own message, plus a sentence for the output drive */
  summary: string;
  /** the temp-drive part of the export dialog's notice (the server's message), null when only the output drive is low */
  tempMessage: string | null;
  /** the 内蔵エンジン recommendation (only with tempLow) */
  recommendation: string | null;
  /** the output-drive part of the export dialog's notice, null when the output drive is fine */
  outputMessage: string | null;
}

/** '1.5 GB' / '800 MB' (same format as the server's message). */
export function formatFreeMB(mb: number): string {
  if (!Number.isFinite(mb) || mb < 0) return '';
  return mb >= 1024 ? `${(mb / 1024).toFixed(1)} GB` : `${Math.round(mb)} MB`;
}

/** 'D: ドライブ' for a path that starts with a drive letter (the server's wording), else null. */
export function driveLabel(path: string | undefined): string | null {
  const m = /^\s*([a-zA-Z]):/.exec(path ?? '');
  return m ? `${m[1].toUpperCase()}: ドライブ` : null;
}

/**
 * The low-disk notice for the header chip and the export dialog, or null when there is nothing to show: no server
 * (standalone / the server stopped answering: its numbers are stale and exports use the browser) or both drives fine.
 * The Japanese text of the temp drive comes from the server (`diskMessage`); a server without it gets the same sentence made here.
 */
export function diskNotice(server: { status: ServerStatus | null; connected: boolean }): DiskNotice | null {
  const st = server.status;
  if (!st || !server.connected) return null;
  const tempLow = st.diskLow === true;
  const outputLow = st.outputDiskLow === true;
  if (!tempLow && !outputLow) return null;
  let tempMessage: string | null = null;
  if (tempLow) {
    const given = typeof st.diskMessage === 'string' ? st.diskMessage.trim() : '';
    if (given) tempMessage = given;
    else {
      const free = typeof st.diskFreeMB === 'number' ? formatFreeMB(st.diskFreeMB) : '';
      tempMessage = `一時ファイルを置くドライブの空き容量が少なくなっています${free ? `（残り ${free}）` : ''}。Excel本体での出力は一時ファイルを作るため、空きが足りないと失敗することがあります。不要なファイルを削除するか、「内蔵エンジン」で出力してください。`;
    }
  }
  const drive = driveLabel(st.outputDir);
  const outputMessage = outputLow
    ? `出力フォルダのドライブ${drive ? `（${drive}）` : ''}の空き容量が少なくなっています。出力したファイルを保存できないことがあります。不要なファイルを削除してから出力してください。`
    : null;
  const recommendation = tempLow ? 'Excel本体エンジンは空き容量が足りないと使えないことがあります。内蔵エンジンでの出力をおすすめします。' : null;
  return { tempLow, outputLow, summary: [tempMessage, outputMessage].filter((t): t is string => !!t).join(' '), tempMessage, recommendation, outputMessage };
}

/** Same status (every field the UI reads)? A refresh that changes nothing must not re-render the app. */
export function sameServerStatus(a: ServerStatus | null, b: ServerStatus | null): boolean {
  if (a === b) return true;
  if (!a || !b) return false;
  return (
    a.version === b.version &&
    a.excel.installed === b.excel.installed &&
    a.excel.version === b.excel.version &&
    a.outputDir === b.outputDir &&
    a.diskLow === b.diskLow &&
    a.diskFreeMB === b.diskFreeMB &&
    a.diskMessage === b.diskMessage &&
    a.outputDiskLow === b.outputDiskLow
  );
}

// ---------------------------------------------------------------------------------------------
// values

/** Promise that resolves after the browser had a chance to paint (never hangs in a background tab). */
export function yieldToUI(): Promise<void> {
  return new Promise((resolve) => {
    let done = false;
    const fin = () => {
      if (!done) {
        done = true;
        resolve();
      }
    };
    if (typeof requestAnimationFrame === 'function') requestAnimationFrame(() => setTimeout(fin, 0));
    setTimeout(fin, 40);
  });
}
