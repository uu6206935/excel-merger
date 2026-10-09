/**
 * Legacy / non-OOXML inputs (owner: reader R): .xls (BIFF), .ods, .xlsb, HTML/XML-2003 "xls" and CSV/TSV/TXT.
 *
 * convertLegacy() reads the file with SheetJS (values + number formats only) and generates a minimal, valid .xlsx
 * package holding those values (buildValuesPackage). read.ts then reads that package with the normal OOXML reader,
 * so the SourceBook always describes exactly the bytes every later step (premerge / export) works on.
 * Excel opens every generated package without repair (tests/fixtures/xlsx/generated/verify-open.json).
 *
 * Hang protection (SheetJS CE loops forever on some corrupt inputs, found by fuzzing — it would freeze the app):
 *  - .xls: SheetJS only sees a sanitized BIFF8 stream (sanitizeBiff8: value records only, formula tokens removed)
 *    and runs without cellStyles (no Theme-zip inflate); widths / number formats / 1904 flag come from scanBiff.
 *  - .ods / .xlsb: every zip entry is inflated by JSZip and re-packed STORED (restoreZip), so SheetJS never inflates.
 *
 * CSV / TSV / TXT: bytes decoded by decodeText (BOM, declared charset, strict UTF-8, else encoding-japanese detection;
 * Shift_JIS = CP932), then parsed by our own parseDelimited (Excel quoting rules; SheetJS would sniff the content
 * first: "ID;…" is SYLK to it). Every field is text except unambiguous values (csvCell: plain numbers without
 * leading zero, thousands-separated numbers, TRUE/FALSE, error literals) — so '00123' stays text; dates and
 * percentages stay text too (safer for key matching than Excel's own CSV guessing). HTML tables use the same rule.
 */
import * as XLSX from 'xlsx';
import Encoding from 'encoding-japanese';
import JSZip from 'jszip';
import type { RangeBox } from '../types';
import { JA_BUILTIN_NUMFMT } from '../display/numfmt';
import { MSG_CORRUPT, MSG_PASSWORD, ReadBookError, assertExpandable, readEntryBytes, resolveTarget } from './zip';
import { attr, encodeOoxmlEscapes, escapeXmlAttr, escapeXmlText, kids, parseXml, rootOf } from './xml';
import { dateModeOf } from './styles';

export interface VCell {
  t: 's' | 'n' | 'b' | 'e';
  v: string | number | boolean;
  /** number format code (own BIFF scan, else as SheetJS reports it; en-US built-in codes map back to their ids) */
  fmt?: string;
}

export interface ValuesSheet {
  name: string;
  /** 0 visible, 1 hidden, 2 very hidden */
  hidden?: 0 | 1 | 2;
  /** rows[r-1][c-1] */
  rows: Array<Array<VCell | undefined> | undefined>;
  merges: RangeBox[];
  /** Excel character widths per column (index = col - 1) */
  colWidths: Array<number | undefined>;
  /**
   * Columns (index = col - 1) whose width is an AUTO width (BIFF ColInfo without fUserSet or with fBestFit, an
   * auto-fitted column): written without customWidth, so the reader reports ColProps.customWidth undefined like for
   * the xlsx twin (customWidth without bestFit = set by the user). Absent = every width in colWidths was set by the
   * user (customWidth="1").
   */
  autoWidth?: ReadonlySet<number>;
}

export interface LegacyConversion {
  /** generated .xlsx package */
  bytes: Uint8Array;
  warnings: string[];
  /** detected text encoding for CSV inputs ('SJIS', 'UTF8', …) */
  encoding?: string;
}

// ---------------------------------------------------------------------------------------------
// Text decoding (CSV)
// ---------------------------------------------------------------------------------------------

/** encoding-japanese name of a charset label declared in HTML / XML ('Shift_JIS', 'windows-31j', 'euc-jp' …). */
function japaneseEncodingOf(label: string): 'SJIS' | 'EUCJP' | 'JIS' | 'UTF8' | undefined {
  const l = label.trim().toLowerCase().replace(/_/g, '-');
  if (/^(shift-jis|sjis|x-sjis|ms-kanji|csshiftjis|windows-31j|cp932|ms932|x-ms-cp932)$/.test(l)) return 'SJIS';
  if (/^(euc-jp|x-euc-jp|cseucpkdfmtjapanese|eucjp)$/.test(l)) return 'EUCJP';
  if (/^(iso-2022-jp|csiso2022jp|jis)$/.test(l)) return 'JIS';
  if (/^(utf-8|utf8|unicode-1-1-utf-8)$/.test(l)) return 'UTF8';
  return undefined;
}

/**
 * Charset declared by a markup document in its first 2 KB: <meta charset="…">, <meta http-equiv content="…;
 * charset=…"> or <?xml … encoding="…"?>. The head of such documents is ASCII in every Japanese encoding.
 */
export function declaredCharset(bytes: Uint8Array): string | undefined {
  const n = Math.min(bytes.length, 2048);
  let head = '';
  for (let i = 0; i < n; i++) head += String.fromCharCode(bytes[i]);
  const m = /<\?xml[^>]*\bencoding\s*=\s*["']([\w.:-]+)["']/i.exec(head) ?? /<meta\b[^>]*\bcharset\s*=\s*["']?([\w.:-]+)/i.exec(head);
  return m ? m[1] : undefined;
}

/**
 * Decode text bytes (CSV / TSV / TXT / HTML / XML) to a string: BOM first, then a charset the document declares
 * (HTML meta / XML declaration), then a strict UTF-8 check, else encoding-japanese detection (Shift_JIS/CP932, EUC-JP,
 * JIS).
 */
export function decodeText(bytes: Uint8Array): { text: string; encoding: string } {
  if (bytes.length >= 3 && bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) return { text: new TextDecoder('utf-8').decode(bytes.subarray(3)), encoding: 'UTF8' };
  if (bytes.length >= 2 && bytes[0] === 0xff && bytes[1] === 0xfe) return { text: new TextDecoder('utf-16le').decode(bytes.subarray(2)), encoding: 'UTF16LE' };
  if (bytes.length >= 2 && bytes[0] === 0xfe && bytes[1] === 0xff) return { text: new TextDecoder('utf-16be').decode(bytes.subarray(2)), encoding: 'UTF16BE' };
  const declared = declaredCharset(bytes);
  const decl = declared ? japaneseEncodingOf(declared) : undefined;
  if (decl && decl !== 'UTF8') return { text: decodeJapanese(bytes, decl), encoding: decl };
  // detection on a bounded sample keeps big files fast; a valid UTF-8 check on the whole file decides UTF-8
  try {
    const text = new TextDecoder('utf-8', { fatal: true }).decode(bytes);
    // ISO-2022-JP (JIS) is 7-bit, hence also valid UTF-8: recognise it by its escape sequences
    if (hasJisEscape(bytes)) return { text: decodeJapanese(bytes, 'JIS'), encoding: 'JIS' };
    return { text, encoding: 'UTF8' };
  } catch {
    /* not UTF-8 (even when declared so: files re-saved in another encoding keep a stale declaration) */
  }
  const sample = bytes.length > 1 << 20 ? bytes.subarray(0, 1 << 20) : bytes;
  let enc = Encoding.detect(sample) || 'SJIS';
  if (enc === 'UTF8' || enc === 'ASCII' || enc === 'BINARY' || enc === 'UNICODE' || enc === 'UTF16' || enc === 'UTF32') enc = 'SJIS';
  const text = enc === 'SJIS' || enc === 'EUCJP' || enc === 'JIS' ? decodeJapanese(bytes, enc) : Encoding.convert(bytes, { to: 'UNICODE', from: enc as Encoding.Encoding, type: 'string' });
  return { text, encoding: enc };
}

/**
 * True when 7-bit text contains an ISO-2022-JP escape into a Japanese set (ESC $ @, ESC $ B: JIS X 0208;
 * ESC ( I: half-width katakana; ESC $ ( D: JIS X 0212). Only called for text that is valid UTF-8; any byte >= 0x80
 * means it is not JIS.
 */
function hasJisEscape(bytes: Uint8Array): boolean {
  let i = bytes.indexOf(0x1b);
  if (i < 0) return false;
  const n = Math.min(bytes.length, 1 << 20);
  for (let k = 0; k < n; k++) if (bytes[k] >= 0x80) return false;
  for (; i >= 0 && i + 2 < bytes.length; i = bytes.indexOf(0x1b, i + 1)) {
    const a = bytes[i + 1];
    const b = bytes[i + 2];
    if (a === 0x24 && (b === 0x40 || b === 0x42 || (b === 0x28 && bytes[i + 3] === 0x44))) return true;
    if (a === 0x28 && b === 0x49) return true;
  }
  return false;
}

const WHATWG_LABEL = { SJIS: 'shift_jis', EUCJP: 'euc-jp', JIS: 'iso-2022-jp' } as const;

/**
 * Shift_JIS (= Windows CP932) / EUC-JP / ISO-2022-JP bytes → string. The platform TextDecoder (WHATWG Encoding
 * Standard: its Shift_JIS decoder is CP932 incl. NEC / IBM extensions, and maps the user-defined area F040-F9FC to
 * U+E000.. like Windows) is ~40x faster than encoding-japanese and decodes every assigned code identically
 * (checked over all 11,280 double-byte codes); encoding-japanese is the fallback where TextDecoder lacks the label.
 */
function decodeJapanese(bytes: Uint8Array, enc: 'SJIS' | 'EUCJP' | 'JIS'): string {
  try {
    return new TextDecoder(WHATWG_LABEL[enc]).decode(bytes);
  } catch {
    return Encoding.convert(bytes, { to: 'UNICODE', from: enc, type: 'string' });
  }
}

export type TextKind = 'delimited' | 'html' | 'xml2003' | 'xml';

/**
 * What a text input is: an Excel 2003 XML spreadsheet (SpreadsheetML), an HTML table (what many business systems
 * export as "*.xls"), or delimited text (CSV / TSV / TXT).
 */
export function classifyText(text: string): TextKind {
  const head = text.slice(0, 4096).replace(/^﻿/, '').trimStart().toLowerCase();
  if (!head.startsWith('<')) return 'delimited';
  if (/urn:schemas-microsoft-com:office:spreadsheet|progid\s*=\s*"excel\.sheet"/.test(head) && /<(ss:)?workbook\b/.test(text.slice(0, 1 << 16).toLowerCase())) return 'xml2003';
  if (/<(html|table|!doctype\s+html|head|body|meta|style|tr|td|th)\b/.test(head)) return 'html';
  // other XML: not a spreadsheet we know; SheetJS gets a try, an empty result means "not an Excel file"
  if (head.startsWith('<?xml')) return 'xml';
  return 'delimited';
}

// ---------------------------------------------------------------------------------------------
// SheetJS → ValuesSheet[]
// ---------------------------------------------------------------------------------------------

const ERR_CODES: Record<number, string> = {
  0x00: '#NULL!', 0x07: '#DIV/0!', 0x0f: '#VALUE!', 0x17: '#REF!', 0x1d: '#NAME?', 0x24: '#NUM!', 0x2a: '#N/A', 0x2b: '#GETTING_DATA',
};

const RE_PLAIN_NUMBER = /^-?(0|[1-9]\d*)(\.\d+)?$/;
const RE_THOUSANDS = /^-?[1-9]\d{0,2}(,\d{3})+(\.\d+)?$/;
const CSV_ERRORS: ReadonlySet<string> = new Set(['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A']);

function significantDigits(s: string): number {
  return s.replace(/^-/, '').replace('.', '').replace(/^0+/, '').length;
}

/**
 * CSV / HTML cell text → VCell. Only unambiguous values are typed: plain numbers (no leading zero, <= 15
 * significant digits), numbers with thousands separators ('1,234' → '#,##0', '1,234.5' → '#,##0.0'), TRUE / FALSE
 * (any case, as Excel) and Excel's error literals (#N/A …) — exactly what Excel writes when it saves a CSV.
 * Leading-zero codes, dates, percentages and everything else stay text.
 */
export function csvCell(text: string): VCell {
  if (RE_PLAIN_NUMBER.test(text) && significantDigits(text) <= 15) {
    const n = Number(text);
    if (Number.isFinite(n)) return { t: 'n', v: n };
  }
  if (RE_THOUSANDS.test(text) && significantDigits(text.replace(/,/g, '')) <= 15) {
    const n = Number(text.replace(/,/g, ''));
    const dot = text.indexOf('.');
    if (Number.isFinite(n)) return { t: 'n', v: n, fmt: dot < 0 ? '#,##0' : '#,##0.' + '0'.repeat(text.length - dot - 1) };
  }
  if (text.length <= 7) {
    const up = text.toUpperCase();
    if (up === 'TRUE' || up === 'FALSE') return { t: 'b', v: up === 'TRUE' };
    if (CSV_ERRORS.has(up)) return { t: 'e', v: up };
  }
  return { t: 's', v: text };
}

/** Rows / columns of a worksheet. */
const MAX_ROWS = 1048576;
const MAX_COLS = 16384;

/**
 * Delimiter of delimited text: a leading "sep=X" line wins (Excel honours it); .tsv = TAB; .txt = TAB unless the
 * first line has commas but no tabs (Excel's text import default is TAB); otherwise (.csv, mis-named files) comma,
 * unless the first line has no comma but tabs (→ TAB) or semicolons (→ ';', European CSV).
 */
export function detectDelimiter(text: string, fileName: string): string {
  const ext = fileName.slice(fileName.lastIndexOf('.') + 1).toLowerCase();
  if (ext === 'tsv') return '\t';
  let commas = 0;
  let tabs = 0;
  let semis = 0;
  let inQ = false;
  for (let i = 0; i < text.length && i < 1 << 16; i++) {
    const ch = text.charCodeAt(i);
    if (ch === 34) inQ = !inQ;
    else if (!inQ) {
      if (ch === 10 || ch === 13) break;
      if (ch === 44) commas++;
      else if (ch === 9) tabs++;
      else if (ch === 59) semis++;
    }
  }
  if (ext === 'txt') return tabs === 0 && commas > 0 ? ',' : '\t';
  if (commas === 0 && tabs > 0) return '\t';
  if (commas === 0 && semis > 0) return ';';
  return ',';
}

/**
 * Parse CSV / TSV text like Excel does (RFC 4180 quoting: "" inside quotes, delimiters and line breaks inside quoted
 * fields; text after a closing quote is appended; CR LF / LF / CR end a record; CR LF inside a quoted field becomes
 * LF, Excel's in-cell line break). Every field goes through csvCell (so '00123' stays text). Empty fields are no
 * cells. The sheet is named after the file, as Excel does. Content beyond 1,048,576 rows / 16,384 columns is
 * dropped (`truncated`).
 */
export function parseDelimited(text: string, fileName: string): ValuesSheet & { truncated?: boolean } {
  let s = text.charCodeAt(0) === 0xfeff ? text.slice(1) : text;
  let delim: string | undefined;
  const sep = /^sep=([^\r\n])(\r\n|\n|\r|$)/i.exec(s);
  if (sep) {
    delim = sep[1];
    s = s.slice(sep[0].length);
  }
  delim ??= detectDelimiter(s, fileName);
  const d = delim.charCodeAt(0);
  const sheet: ValuesSheet & { truncated?: boolean } = { name: sanitizeSheetName(baseName(fileName), new Set()), hidden: 0, rows: [], merges: [], colWidths: [] };
  const n = s.length;
  let r = 0;
  let c = 0;
  let row: Array<VCell | undefined> = [];
  let any = false;
  let i = 0;
  const endRow = () => {
    if (any) {
      if (r < MAX_ROWS) sheet.rows[r] = row;
      else sheet.truncated = true;
    }
    r++;
    c = 0;
    row = [];
    any = false;
  };
  const fieldEnd = (from: number): number => {
    let k = from;
    while (k < n) {
      const ch = s.charCodeAt(k);
      if (ch === d || ch === 10 || ch === 13) break;
      k++;
    }
    return k;
  };
  for (;;) {
    let value: string;
    if (s.charCodeAt(i) === 34 /* " */) {
      let out = '';
      let j = i + 1;
      for (;;) {
        const q = s.indexOf('"', j);
        if (q < 0) {
          out += s.slice(j);
          i = n;
          break;
        }
        out += s.slice(j, q);
        if (s.charCodeAt(q + 1) === 34) {
          out += '"';
          j = q + 2;
          continue;
        }
        i = q + 1;
        break;
      }
      const k = fieldEnd(i);
      value = (out + s.slice(i, k)).replace(/\r\n?/g, '\n');
      i = k;
    } else {
      const k = fieldEnd(i);
      value = s.slice(i, k);
      i = k;
    }
    if (value !== '') {
      if (c < MAX_COLS) {
        row[c] = csvCell(value);
        any = true;
      } else sheet.truncated = true;
    }
    if (i >= n) {
      endRow();
      break;
    }
    const ch = s.charCodeAt(i);
    if (ch === d) {
      c++;
      i++;
      continue;
    }
    i += ch === 13 && s.charCodeAt(i + 1) === 10 ? 2 : 1;
    endRow();
    if (i >= n) break;
  }
  return sheet;
}

function sanitizeSheetName(name: string, used: Set<string>): string {
  let s = name.replace(/[\\/?*[\]:]/g, '_').replace(/^'+|'+$/g, '').trim();
  if (!s) s = 'Sheet';
  if (s.length > 31) s = s.slice(0, 31);
  let out = s;
  for (let i = 2; used.has(out.toLowerCase()); i++) {
    const suffix = ` (${i})`;
    out = s.slice(0, 31 - suffix.length) + suffix;
  }
  used.add(out.toLowerCase());
  return out;
}

function baseName(fileName: string): string {
  const f = fileName.replace(/^.*[\\/]/, '');
  const i = f.lastIndexOf('.');
  return i > 0 ? f.slice(0, i) : f;
}

/** Marker for ODS error cells (see prepareOds): invisible separators around "ERR", never found in real data. */
const ODS_ERR_MARK = '⁣ERR⁣';

/** Pixels (96 dpi) of an ODF length ("2.258cm", "0.889in", "64pt" …); undefined when unparsable. */
function odfLengthPx(v: string | undefined): number | undefined {
  if (!v) return undefined;
  const m = /^\s*([0-9]*\.?[0-9]+)\s*(in|cm|mm|pt|pc|px)\s*$/i.exec(v);
  if (!m) return undefined;
  const n = parseFloat(m[1]);
  const perInch: Record<string, number> = { in: 1, cm: 2.54, mm: 25.4, pt: 72, pc: 6, px: 96 };
  const px = (n / perInch[m[2].toLowerCase()]) * 96;
  return Number.isFinite(px) ? px : undefined;
}

/** <style:style …/> or <style:style …>…</style:style> (attributes in m[1], body in m[2]). */
const RE_ODS_STYLE = /<style:style\b([^>]*?)(?:\/>|>([\s\S]*?)<\/style:style>)/g;

/** Maximum digit width (px) of 游ゴシック 11pt, the font of the generated package: OOXML width = px / MDW. */
const GEN_MDW = 8;

/** Column widths of one ODS table (OOXML units, index = col - 1) and the columns among them that are auto widths. */
interface OdsTableColumns {
  widths: Array<number | undefined>;
  auto: Set<number>;
}

/**
 * Column widths of every table in an ODS content.xml (index = table order = SheetJS sheet order), in OOXML width
 * units for the generated package's font. Only columns that differ from the table's default (most repeated)
 * column width are returned. Verified: Excel's ODS export of 14.625 / 11.375 wide columns reads back exactly.
 * `auto` = the columns whose style says style:use-optimal-column-width="true" (Excel writes it for an auto-fitted
 * column, the twin of <col bestFit="1">): an auto width, not one set by the user.
 */
function odsColumnWidths(xml: string): OdsTableColumns[] {
  const stylePx = new Map<string, number>();
  const styleOptimal = new Set<string>();
  for (const m of xml.matchAll(RE_ODS_STYLE)) {
    if (!/style:family="table-column"/.test(m[1])) continue;
    const name = /style:name="([^"]*)"/.exec(m[1])?.[1];
    const px = odfLengthPx(/style:column-width="([^"]*)"/.exec(m[2] ?? '')?.[1]);
    if (name !== undefined && px !== undefined) stylePx.set(name, px);
    if (name !== undefined && /style:use-optimal-column-width="true"/.test(m[2] ?? '')) styleOptimal.add(name);
  }
  const out: OdsTableColumns[] = [];
  // (?=[\s>]) — \b would also match <table:table-row>, <table:table-cell> …
  for (const t of xml.matchAll(/<table:table(?=[\s>])[^>]*>/g)) {
    const start = (t.index ?? 0) + t[0].length;
    let end = xml.indexOf('<table:table-row', start);
    const close = xml.indexOf('</table:table>', start);
    if (end < 0 || (close >= 0 && close < end)) end = close < 0 ? xml.length : close;
    const cols: Array<{ px: number | undefined; n: number; auto: boolean }> = [];
    for (const c of xml.slice(start, end).matchAll(/<table:table-column(?=[\s/>])([^>]*?)\/?>/g)) {
      const style = /table:style-name="([^"]*)"/.exec(c[1])?.[1];
      const rep = Math.max(1, Math.min(16384, parseInt(/table:number-columns-repeated="(\d+)"/.exec(c[1])?.[1] ?? '1', 10) || 1));
      cols.push({ px: style !== undefined ? stylePx.get(style) : undefined, n: rep, auto: style !== undefined && styleOptimal.has(style) });
    }
    let dflt: number | undefined;
    let most = 0;
    for (const c of cols) if (c.n > most) [most, dflt] = [c.n, c.px];
    const widths: Array<number | undefined> = [];
    const auto = new Set<number>();
    let col = 0;
    for (const c of cols) {
      if (col >= 16384) break;
      const w = c.px !== undefined && (dflt === undefined || Math.abs(c.px - dflt) >= 0.5) ? Math.min(255, Math.trunc((Math.round(c.px) / GEN_MDW) * 256) / 256) : undefined;
      if (w !== undefined && w > 0) {
        for (let k = 0; k < c.n && col + k < 16384; k++) {
          widths[col + k] = w;
          if (c.auto) auto.add(col + k);
        }
      }
      col += c.n;
    }
    out.push({ widths, auto });
  }
  return out;
}

/** Visit the records of an XLSB part (record id 1-2 bytes, size 1-4 bytes, 7 bits each); `visit` returns false to stop. */
function forEachXlsbRecord(b: Uint8Array, visit: (id: number, start: number, len: number) => boolean | void): void {
  let p = 0;
  while (p < b.length) {
    let id = b[p++];
    if (id & 0x80) {
      if (p >= b.length) return;
      id = (id & 0x7f) | ((b[p++] & 0x7f) << 7);
    }
    let len = 0;
    for (let i = 0, mul = 1; i < 4; i++, mul *= 128) {
      if (p >= b.length) return;
      const c = b[p++];
      len += (c & 0x7f) * mul;
      if (!(c & 0x80)) break;
    }
    if (p + len > b.length) return;
    if (visit(id, p, len) === false) return;
    p += len;
  }
}

/**
 * Auto-width columns of an .xlsb (per sheet NAME; index = col - 1): BrtColInfo (record 60) grbit fUserSet = 0 or
 * fBestFit = 1 — the same rule as BIFF8 ColInfo (scanBiff). SheetJS returns the widths but not these flags. The
 * sheet parts come from BrtBundleSh (record 156: hsState, iTabID, strRelID, strName) + xl/_rels/workbook.bin.rels.
 * Bounds-checked: garbage yields fewer results, never an exception (the zip was already inflated by restoreZip).
 */
async function scanXlsbAutoWidths(zip: JSZip): Promise<Map<string, Set<number>>> {
  const out = new Map<string, Set<number>>();
  try {
    const wbFile = zip.file('xl/workbook.bin');
    const relsFile = zip.file('xl/_rels/workbook.bin.rels');
    if (!wbFile || !relsFile) return out;
    const targets = new Map<string, string>();
    for (const r of kids(rootOf(parseXml(await relsFile.async('string'))), 'Relationship')) {
      const id = attr(r, 'Id');
      const target = attr(r, 'Target');
      if (id !== undefined && target !== undefined && (attr(r, 'TargetMode') ?? '').toLowerCase() !== 'external') targets.set(id, resolveTarget('xl', target));
    }
    const wb = await wbFile.async('uint8array');
    const sheets: Array<{ name: string; path: string }> = [];
    forEachXlsbRecord(wb, (id, start, len) => {
      if (id !== 156 || len < 16) return;
      const dv = new DataView(wb.buffer, wb.byteOffset + start, len);
      // XLNullableWideString / XLWideString: u32 character count (0xFFFFFFFF = null) + UTF-16LE characters
      const wide = (at: number): { text: string; next: number } | undefined => {
        if (at + 4 > len) return undefined;
        const cch = dv.getUint32(at, true);
        if (cch === 0xffffffff) return { text: '', next: at + 4 };
        if (cch > 255 || at + 4 + cch * 2 > len) return undefined;
        let text = '';
        for (let i = 0; i < cch; i++) text += String.fromCharCode(dv.getUint16(at + 4 + i * 2, true));
        return { text, next: at + 4 + cch * 2 };
      };
      const rel = wide(8);
      const name = rel && wide(rel.next);
      const path = rel && targets.get(rel.text);
      if (name && path !== undefined) sheets.push({ name: name.text, path });
    });
    for (const sh of sheets) {
      const f = zip.file(sh.path);
      if (!f) continue;
      const b = await f.async('uint8array');
      const auto = new Set<number>();
      let records = 0;
      forEachXlsbRecord(b, (id, start, len) => {
        if (id === 145) return false; // BrtBeginSheetData: the column infos precede the cells
        if (id !== 60 || len < 18 || ++records > 16384) return;
        const dv = new DataView(b.buffer, b.byteOffset + start, len);
        const first = dv.getUint32(0, true);
        const last = Math.min(dv.getUint32(4, true), 16383);
        const grbit = dv.getUint16(16, true);
        const userSet = (grbit & 0x0002) !== 0 && (grbit & 0x0004) === 0;
        for (let c = first; c <= last; c++) {
          if (userSet) auto.delete(c);
          else auto.add(c);
        }
      });
      if (auto.size) out.set(sh.name, auto);
    }
  } catch {
    /* no flags: every width is taken as set by the user (the conservative reading) */
  }
  return out;
}

/**
 * ODS / XLSB pre-pass (SheetJS CE gaps):
 *  - SheetJS throws "Unsupported value type error" on error cells (office:value-type="error", which Excel writes
 *    for #N/A, #DIV/0! …): they are rewritten as marked string cells; sheetJsToValues turns them back into errors.
 *  - SheetJS ignores hidden tables (table:display="false" in the table's automatic style): returned in `hidden`.
 *  - SheetJS returns no column widths for ODS: read from the column styles (`colWidths`, per table index).
 *  - Which widths are AUTO widths (not set by the user): ODS `autoWidths` per table index (use-optimal-column-width),
 *    XLSB `autoWidthsByName` per sheet name (BrtColInfo flags); see ValuesSheet.autoWidth.
 * Every zip input (.ods, .xlsb) is also re-packed with STORED entries (see restoreZip).
 */
export async function prepareOds(bytes: Uint8Array): Promise<{
  bytes: Uint8Array;
  hidden: Set<string>;
  colWidths: Array<Array<number | undefined>>;
  autoWidths: Array<Set<number> | undefined>;
  autoWidthsByName: Map<string, Set<number>>;
  isOds: boolean;
}> {
  const hidden = new Set<string>();
  const autoWidthsByName = new Map<string, Set<number>>();
  if (bytes.length < 4 || bytes[0] !== 0x50 || bytes[1] !== 0x4b) return { bytes, hidden, colWidths: [], autoWidths: [], autoWidthsByName, isOds: false };
  const zip = await restoreZip(bytes);
  const f = zip.file('content.xml');
  if (!f) {
    for (const [k, v] of await scanXlsbAutoWidths(zip)) autoWidthsByName.set(k, v);
    return { bytes: await zip.generateAsync({ type: 'uint8array', compression: 'STORE' }), hidden, colWidths: [], autoWidths: [], autoWidthsByName, isOds: false };
  }
  const xml = await f.async('string');
  // hidden tables
  const hiddenStyles = new Set<string>();
  for (const m of xml.matchAll(RE_ODS_STYLE)) {
    const name = /style:name="([^"]+)"/.exec(m[1])?.[1];
    if (name && /style:family="table"/.test(m[1]) && /table:display="false"/.test(m[2] ?? '')) hiddenStyles.add(name);
  }
  for (const m of xml.matchAll(/<table:table(?=[\s>])([^>]*)>/g)) {
    const name = /table:name="([^"]*)"/.exec(m[1])?.[1];
    const style = /table:style-name="([^"]*)"/.exec(m[1])?.[1];
    if (name !== undefined && style && hiddenStyles.has(style)) hidden.add(decodeAttr(name));
  }
  const tables = odsColumnWidths(xml);
  const colWidths = tables.map((t) => t.widths);
  const autoWidths = tables.map((t) => (t.auto.size ? t.auto : undefined));
  if (!xml.includes('office:value-type="error"')) return { bytes: await zip.generateAsync({ type: 'uint8array', compression: 'STORE' }), hidden, colWidths, autoWidths, autoWidthsByName, isOds: true };
  const out = xml.replace(
    /<table:table-cell\b([^>]*?)office:value-type="error"([^>]*)(?<!\/)>([\s\S]*?)<\/table:table-cell>/g,
    (_m, a: string, b: string, inner: string) => {
      const attrs = (a + 'office:value-type="string"' + b).replace(/office:string-value="([^"]*)"/, (_x, v: string) => `office:string-value="${ODS_ERR_MARK}${v}"`);
      return `<table:table-cell${attrs}>${inner.replace(/<text:p>/, '<text:p>' + ODS_ERR_MARK)}</table:table-cell>`;
    },
  );
  zip.file('content.xml', out);
  return { bytes: await zip.generateAsync({ type: 'uint8array', compression: 'STORE' }), hidden, colWidths, autoWidths, autoWidthsByName, isOds: true };
}

/**
 * Inflate every entry of a zip with JSZip (pako: rejects corrupt deflate data; CRC checked) into a new zip whose
 * entries are STORED. SheetJS then never runs its own inflate, which loops forever on some corrupt deflate
 * streams (found by fuzzing; it would freeze the app). Zip-bomb safe: the declared total is checked first
 * (assertExpandable → ReadBookError 'tooLarge') and no entry inflates beyond its declared size (readEntryBytes).
 * Throws ReadBookError('corrupt') for a broken zip.
 */
export async function restoreZip(bytes: Uint8Array): Promise<JSZip> {
  assertExpandable(bytes);
  let src: JSZip;
  try {
    src = await JSZip.loadAsync(bytes, { checkCRC32: false });
  } catch (e) {
    throw new ReadBookError('corrupt', MSG_CORRUPT, { cause: e });
  }
  assertExpandable(bytes, src);
  try {
    const out = new JSZip();
    for (const name of Object.keys(src.files)) {
      const e = src.files[name];
      if (e.dir) continue;
      out.file(name, await readEntryBytes(e, true));
    }
    return out;
  } catch (e) {
    throw new ReadBookError('corrupt', MSG_CORRUPT, { cause: e });
  }
}

// ---------------------------------------------------------------------------------------------
// BIFF (.xls) column widths
// ---------------------------------------------------------------------------------------------

interface CfbEntryLike {
  content?: ArrayLike<number>;
}
interface CfbLike {
  read(data: Uint8Array, o: { type: 'array' }): unknown;
  find(cfb: unknown, path: string): CfbEntryLike | null | undefined;
}

/** The BIFF Workbook ('Workbook' BIFF8 / 'Book' BIFF5) stream of a CFB .xls; undefined when unreadable. */
export function biffWorkbookStream(bytes: Uint8Array): Uint8Array | undefined {
  try {
    const CFB = (XLSX as unknown as { CFB: CfbLike }).CFB;
    const cfb = CFB.read(bytes, { type: 'array' });
    const e = CFB.find(cfb, 'Workbook') ?? CFB.find(cfb, 'Book');
    if (!e?.content || e.content.length < 8) return undefined;
    return e.content instanceof Uint8Array ? e.content : Uint8Array.from(e.content);
  } catch {
    return undefined;
  }
}

/** BIFF8 records SheetJS needs for values, sheet names / visibility, number formats and merges. */
const BIFF8_KEEP: ReadonlySet<number> = new Set([
  0x0809, // BOF
  0x000a, // EOF
  0x0042, // CodePage
  0x0022, // Date1904
  0x041e, // Format
  0x00e0, // XF
  0x00fc, // SST (+ Continue)
  0x0085, // BoundSheet8
  0x0200, // Dimensions
  0x0203, // Number
  0x027e, // RK
  0x00bd, // MulRk
  0x00fd, // LabelSst
  0x0204, // Label
  0x00d6, // RString
  0x0205, // BoolErr
  0x0006, // Formula (rgce removed, cached value kept)
  0x0207, // String (cached text result of the preceding Formula)
  0x00e5, // MergeCells
  0x007d, // ColInfo
]);

/**
 * Rebuild a BIFF8 Workbook stream with only the records needed for values (BIFF8_KEEP), formula token arrays
 * removed (cce = 0; the cached result stays), embedded chart / object substreams dropped, and BoundSheet8 offsets
 * re-pointed to the moved sheet BOFs. SheetJS parses every formula token array even when formulas are not asked
 * for, and a corrupt record (a Style record whose type byte flipped into a formula …) makes its parse_Rgce /
 * record loop spin forever — found by fuzzing, it would freeze the app. Values-only reading never needs those
 * records. Returns undefined when the stream is not BIFF8 or is encrypted (FilePass: SheetJS reports the
 * password itself from the original bytes).
 */
export function sanitizeBiff8(st: Uint8Array): Uint8Array | undefined {
  const n = st.length;
  const u16 = (p: number) => st[p] | (st[p + 1] << 8);
  if (n < 8 || u16(0) !== 0x0809 || u16(2) < 4 || u16(4) !== 0x0600) return undefined;
  const kept: Array<{ off: number; type: number; data: Uint8Array }> = [];
  let depth = 0;
  let lastKept = false;
  for (let p = 0, guard = 0; p + 4 <= n && guard < 2_000_000; guard++) {
    const type = u16(p);
    const len = u16(p + 2);
    const d = p + 4;
    if (d + len > n) break; // truncated: keep what is complete
    if (type === 0x002f) return undefined; // FilePass
    let keep = false;
    let data = st.subarray(d, d + len);
    if (type === 0x0809) {
      depth++;
      keep = depth === 1;
    } else if (type === 0x000a) {
      keep = depth === 1;
      depth = Math.max(0, depth - 1);
    } else if (depth === 1) {
      if (type === 0x003c) keep = lastKept;
      else if (type === 0x0006) {
        if (len >= 22) {
          const f = new Uint8Array(22);
          f.set(st.subarray(d, d + 20));
          data = f; // cce = 0
          keep = true;
        }
      } else keep = BIFF8_KEEP.has(type);
    }
    if (type !== 0x003c) lastKept = keep;
    if (keep) kept.push({ off: p, type, data });
    p = d + len;
  }
  // new offsets, then re-point BoundSheet8.lbPlyPos
  const newOff = new Map<number, number>();
  let size = 0;
  for (const r of kept) {
    newOff.set(r.off, size);
    size += 4 + r.data.length;
  }
  const out = new Uint8Array(size);
  let w = 0;
  for (const r of kept) {
    out[w] = r.type & 0xff;
    out[w + 1] = r.type >> 8;
    out[w + 2] = r.data.length & 0xff;
    out[w + 3] = r.data.length >> 8;
    out.set(r.data, w + 4);
    if (r.type === 0x0085 && r.data.length >= 4) {
      const old = (r.data[0] | (r.data[1] << 8) | (r.data[2] << 16)) + r.data[3] * 0x1000000;
      const nw = newOff.get(old);
      if (nw !== undefined) {
        out[w + 4] = nw & 0xff;
        out[w + 5] = (nw >>> 8) & 0xff;
        out[w + 6] = (nw >>> 16) & 0xff;
        out[w + 7] = (nw >>> 24) & 0xff;
      }
    }
    w += 4 + r.data.length;
  }
  return out;
}

export interface BiffSheetScan {
  /** BIFF8 only */
  name?: string;
  /** OOXML column widths (ColInfo), index = col - 1 */
  widths: Array<number | undefined>;
  /**
   * columns (index = col - 1) of `widths` whose ColInfo is an auto width, not one set by the user: fUserSet = 0, or
   * fBestFit = 1 (an auto-fitted column, grbit 0x0006, the twin of <col bestFit="1" customWidth="1">)
   */
  autoWidths: Set<number>;
  /** BIFF8 only: number format code of numeric / formula cells, key = (row - 1) * 16384 + (col - 1) */
  numFmts: Map<number, string>;
}

export interface BiffScan {
  date1904: boolean;
  sheets: BiffSheetScan[];
}

/**
 * Own pass over a BIFF5/8 Workbook stream for what SheetJS does not return reliably, per sheet in BoundSheet order:
 *  - column widths (ColInfo). The .xls is read by SheetJS WITHOUT cellStyles, because with cellStyles SheetJS
 *    inflates the zip embedded in the Theme record with its own inflate, which loops forever on corrupt data
 *    (found by fuzzing); SheetJS only returns widths with cellStyles.
 *  - BIFF8: the number format of every numeric / formula cell (XF → Format record, else the ja-JP built-in code).
 *    SheetJS gives formula cells no format and has no code for the ja-JP built-in ids 27-36 / 50-58
 *    (yyyy"年"m"月"d"日" …), so dates computed by formulas would lose their format.
 *  - the 1904 date system flag (Date1904).
 * Bounds-checked: garbage yields fewer results, never an exception or a hang.
 */
export function scanBiff(stream: Uint8Array): BiffScan {
  const st = stream;
  const n = st.length;
  const u16 = (p: number) => st[p] | (st[p + 1] << 8);
  const u32 = (p: number) => (st[p] | (st[p + 1] << 8) | (st[p + 2] << 16)) + st[p + 3] * 0x1000000;
  const biff8 = u16(0) === 0x0809 && u16(2) >= 4 && u16(4) === 0x0600;
  const xlString = (p: number, end: number, cchBytes: 1 | 2): string | undefined => {
    const cch = cchBytes === 2 ? u16(p) : st[p];
    const f = p + cchBytes;
    const hi = (st[f] & 1) === 1;
    if (f + 1 + cch * (hi ? 2 : 1) > end) return undefined;
    let s = '';
    for (let i = 0; i < cch; i++) s += String.fromCharCode(hi ? u16(f + 1 + i * 2) : st[f + 1 + i]);
    return s;
  };
  // globals substream: BoundSheet8, Format, XF, Date1904 records up to the globals EOF
  const sheets: Array<{ pos: number; name?: string }> = [];
  const formats = new Map<number, string>();
  const xfFmt: number[] = [];
  let date1904 = false;
  for (let p = 0, guard = 0; p + 4 <= n && guard < 1_000_000; guard++) {
    const type = u16(p);
    const len = u16(p + 2);
    const d = p + 4;
    if (d + len > n) break;
    if (type === 0x0022 && len >= 2) date1904 = u16(d) === 1;
    else if (biff8 && type === 0x041e && len >= 5) {
      const code = xlString(d + 2, d + len, 2);
      if (code !== undefined) formats.set(u16(d), code);
    } else if (biff8 && type === 0x00e0 && len >= 4) xfFmt.push(u16(d + 2));
    else if (type === 0x0085 && len >= 6) {
      let name: string | undefined;
      if (biff8 && len >= 8) {
        const cch = st[d + 6];
        const hi = (st[d + 7] & 1) === 1;
        if (8 + cch * (hi ? 2 : 1) <= len) {
          name = '';
          for (let i = 0; i < cch; i++) name += String.fromCharCode(hi ? u16(d + 8 + i * 2) : st[d + 8 + i]);
        }
      }
      sheets.push({ pos: u32(d), name });
    } else if (type === 0x000a) break;
    p = d + len;
  }
  const codeOfXf = (ixfe: number): string | undefined => {
    const id = xfFmt[ixfe];
    return id === undefined ? undefined : formats.get(id) ?? JA_BUILTIN_NUMFMT[id];
  };
  const out = sheets.map((sh): BiffSheetScan => {
    const widths: Array<number | undefined> = [];
    const autoWidths = new Set<number>();
    const numFmts = new Map<number, string>();
    const res = { name: sh.name, widths, autoWidths, numFmts };
    let p = sh.pos;
    if (p + 4 > n || (u16(p) !== 0x0809 && u16(p) !== 0x0409 && u16(p) !== 0x0209)) return res;
    const setFmt = (row: number, col: number, ixfe: number) => {
      if (row > 65535 || col > 255) return;
      const code = codeOfXf(ixfe);
      if (code !== undefined) numFmts.set(row * 16384 + col, code);
    };
    let depth = 0;
    for (let guard = 0; p + 4 <= n && guard < 2_000_000; guard++) {
      const type = u16(p);
      const len = u16(p + 2);
      const d = p + 4;
      if (d + len > n) break;
      if (type === 0x0809 || type === 0x0409 || type === 0x0209) depth++;
      else if (type === 0x000a && --depth <= 0) break;
      else if (depth === 1) {
        if (type === 0x007d && len >= 10) {
          const first = u16(d);
          const last = Math.min(u16(d + 2), 255);
          const w = u16(d + 4) / 256;
          const grbit = u16(d + 8);
          // a width set by the USER = fUserSet (0x0002, Excel: customWidth) without fBestFit (0x0004, Excel: bestFit;
          // an auto-fitted column is saved as 0x0006). Anything else is an auto width (see BiffSheetScan.autoWidths)
          const userSet = (grbit & 0x0002) !== 0 && (grbit & 0x0004) === 0;
          if (w > 0) {
            for (let c = first; c <= last; c++) {
              widths[c] = Math.min(255, Math.round(w * 1000) / 1000);
              if (userSet) autoWidths.delete(c);
              else autoWidths.add(c);
            }
          }
        } else if (biff8 && (type === 0x0203 || type === 0x027e || type === 0x0006) && len >= 6) {
          setFmt(u16(d), u16(d + 2), u16(d + 4)); // Number / RK / Formula: rw, col, ixfe
        } else if (biff8 && type === 0x00bd && len >= 12) {
          const row = u16(d);
          const first = u16(d + 2);
          const count = Math.floor((len - 6) / 6);
          for (let i = 0; i < count; i++) setFmt(row, first + i, u16(d + 4 + i * 6)); // MulRk: RkRec = ixfe + RK
        }
      }
      p = d + len;
    }
    return res;
  });
  return { date1904, sheets: out };
}

function decodeAttr(s: string): string {
  return s.replace(/&(amp|lt|gt|quot|apos);/g, (_m, e: string) => ({ amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" })[e] ?? _m);
}

function isPasswordError(e: unknown): boolean {
  const m = e instanceof Error ? e.message : String(e);
  return /password|encrypt|パスワード/i.test(m);
}

/**
 * Read with SheetJS and convert to the values model.
 * `extra.numFmts[i]`: own number formats of sheet i (scanBiff), preferred over SheetJS' `z`.
 * `extra.xml2003`: an Excel 2003 XML spreadsheet; a <Column> without ss:AutoFitWidth="0" is an auto width (Excel writes
 * AutoFitWidth="0" for the widths the user set), see ValuesSheet.autoWidth.
 * `extra.date1904`: the source uses the 1904 date system; date-formatted serials are converted to 1900 serials like
 * the OOXML reader does (calendar dates always, time formats only for values >= 1, durations never).
 */
export function sheetJsToValues(
  wb: XLSX.WorkBook,
  csv: boolean,
  fileName: string,
  extra: { numFmts?: Array<Map<number, string> | undefined>; date1904?: boolean; xml2003?: boolean } = {},
): ValuesSheet[] {
  const used = new Set<string>();
  const out: ValuesSheet[] = [];
  wb.SheetNames.forEach((sn, idx) => {
    const ws = wb.Sheets[sn];
    const name = sanitizeSheetName(csv && wb.SheetNames.length === 1 ? baseName(fileName) : sn, used);
    const hiddenRaw = wb.Workbook?.Sheets?.[idx]?.Hidden;
    const sheet: ValuesSheet = { name, hidden: hiddenRaw === 1 || hiddenRaw === 2 ? hiddenRaw : 0, rows: [], merges: [], colWidths: [] };
    if (ws) {
      const addCell = (r: number, c: number, cell: XLSX.CellObject | undefined): void => {
        if (!cell || cell.t === 'z') return;
        if (r < 1 || c < 1 || r > 1048576 || c > 16384) return;
        let vc: VCell | undefined;
        if (cell.t === 's') {
          const text = String(cell.v ?? '');
          if (text === '') return;
          if (!csv && text.startsWith(ODS_ERR_MARK)) vc = { t: 'e', v: text.slice(ODS_ERR_MARK.length) || '#N/A' };
          else vc = csv ? csvCell(text) : { t: 's', v: text };
        } else if (cell.t === 'n') {
          const n = Number(cell.v);
          if (!Number.isFinite(n)) return;
          vc = { t: 'n', v: n };
          const own = extra.numFmts?.[idx]?.get((r - 1) * 16384 + (c - 1));
          const fmt = own !== undefined ? own : cell.z;
          if (fmt && fmt !== 'General') vc.fmt = String(fmt);
          if (extra.date1904 && vc.fmt) {
            const mode = dateModeOf(vc.fmt);
            if (mode === 'date' || (mode === 'time' && n >= 1)) vc.v = n + 1462;
          }
        } else if (cell.t === 'b') {
          vc = { t: 'b', v: !!cell.v };
        } else if (cell.t === 'e') {
          const w = typeof cell.w === 'string' && cell.w.startsWith('#') ? cell.w : ERR_CODES[Number(cell.v)] ?? '#N/A';
          vc = { t: 'e', v: w };
        } else if (cell.t === 'd') {
          const d = cell.v instanceof Date ? cell.v : new Date(String(cell.v));
          if (Number.isNaN(d.getTime())) return;
          const serial = (Date.UTC(d.getFullYear(), d.getMonth(), d.getDate(), d.getHours(), d.getMinutes(), d.getSeconds()) / 86400000) + 25569;
          vc = { t: 'n', v: serial, fmt: cell.z ? String(cell.z) : 'yyyy/m/d' };
        }
        if (!vc) return;
        let row = sheet.rows[r - 1];
        if (!row) sheet.rows[r - 1] = row = [];
        row[c - 1] = vc;
      };
      const dense = (ws as { '!data'?: Array<Array<XLSX.CellObject | undefined> | undefined> })['!data'];
      if (Array.isArray(dense)) {
        // dense mode (all reads use it): no per-cell address decoding
        for (let ri = 0; ri < dense.length; ri++) {
          const row = dense[ri];
          if (!row) continue;
          for (let ci = 0; ci < row.length; ci++) addCell(ri + 1, ci + 1, row[ci]);
        }
      } else {
        for (const key of Object.keys(ws)) {
          if (key.charCodeAt(0) === 33 /* ! */) continue;
          const addr = XLSX.utils.decode_cell(key);
          addCell(addr.r + 1, addr.c + 1, ws[key] as XLSX.CellObject);
        }
      }
      for (const m of (ws['!merges'] as XLSX.Range[] | undefined) ?? []) {
        sheet.merges.push({ r1: m.s.r + 1, c1: m.s.c + 1, r2: m.e.r + 1, c2: m.e.c + 1 });
      }
      const cols = (ws['!cols'] as XLSX.ColInfo[] | undefined) ?? [];
      const auto = new Set<number>();
      cols.forEach((ci, i) => {
        if (!ci) return;
        const w = ci.width ?? ci.wch;
        if (typeof w === 'number' && Number.isFinite(w) && w > 0) {
          sheet.colWidths[i] = Math.min(255, Math.round(w * 1000) / 1000);
          // XML 2003 keeps the <Column> attributes on the entry (SheetJS' own parse)
          if (extra.xml2003 && String((ci as { AutoFitWidth?: unknown }).AutoFitWidth) !== '0') auto.add(i);
        }
      });
      if (auto.size) sheet.autoWidth = auto;
    }
    out.push(sheet);
  });
  return out;
}

/**
 * Convert a non-OOXML input to a values-only .xlsx package.
 *  - `text`: the bytes are text (CSV / TSV / TXT, an HTML table or an Excel 2003 XML spreadsheet — business systems
 *    often export those two as "*.xls"). Decoded with decodeText (BOM / declared charset / UTF-8 / Shift_JIS …).
 *    HTML tables are read like CSV (raw text; only unambiguous numbers become numbers, so 00123 stays text; the
 *    sheet is named after the file like Excel does); XML 2003 cells keep their declared types.
 *    `delimited` forces CSV parsing (.csv / .tsv / .txt names).
 *  - otherwise binary: .xls (BIFF), .xlsb, .ods via SheetJS.
 */
export async function convertLegacy(
  bytes: Uint8Array,
  fileName: string,
  opts: { text?: boolean; delimited?: boolean } = {},
): Promise<LegacyConversion & { kind: 'binary' | TextKind }> {
  const warnings: string[] = [];
  let wb: XLSX.WorkBook | undefined;
  let delimitedSheet: ValuesSheet | undefined;
  let encoding: string | undefined;
  let kind: 'binary' | TextKind = 'binary';
  let extraWidths: Array<Array<number | undefined>> = [];
  let extraAuto: Array<Set<number> | undefined> = [];
  let autoByName = new Map<string, Set<number>>(); // .xlsb: auto-width columns per sheet name
  let numFmts: Array<Map<number, string> | undefined> | undefined;
  let date1904 = false;
  try {
    if (opts.text) {
      const dec = decodeText(bytes);
      encoding = dec.encoding;
      kind = opts.delimited ? 'delimited' : classifyText(dec.text);
      if (kind === 'delimited') {
        // own parser: SheetJS sniffs the content first ("ID;…" = SYLK, "<…" = XML/HTML) whatever the caller wants
        const parsed = parseDelimited(dec.text, fileName);
        if (parsed.truncated) warnings.push('Excelの上限（1,048,576行・16,384列）を超える部分は読み込みませんでした。');
        delimitedSheet = { name: parsed.name, hidden: parsed.hidden, rows: parsed.rows, merges: parsed.merges, colWidths: parsed.colWidths };
      } else if (kind === 'html') {
        wb = XLSX.read(dec.text, { type: 'string', raw: true, dense: true, cellDates: false, cellHTML: false });
      } else {
        wb = XLSX.read(dec.text, { type: 'string', cellNF: true, cellDates: false, cellFormula: false, cellHTML: false, dense: true });
      }
    } else if (!(bytes[0] === 0x50 && bytes[1] === 0x4b)) {
      // .xls (CFB / raw BIFF): a sanitized BIFF8 stream (see sanitizeBiff8), no cellStyles (see
      // scanBiff); widths, number formats and the 1904 flag from our own record scan
      const stream = bytes[0] === 0xd0 ? biffWorkbookStream(bytes) : undefined;
      const clean = stream ? sanitizeBiff8(stream) : undefined;
      const w = XLSX.read(clean ?? bytes, { type: 'array', cellNF: true, cellDates: false, cellFormula: false, cellHTML: false, cellStyles: false, sheetStubs: false, dense: true });
      wb = w;
      const biff = stream ? scanBiff(stream) : undefined;
      if (biff && w.SheetNames) {
        const byName = new Map(biff.sheets.filter((b) => b.name !== undefined).map((b) => [b.name!, b] as const));
        const per = w.SheetNames.map((sn, i) => byName.get(sn) ?? (biff.sheets.length === w.SheetNames.length ? biff.sheets[i] : undefined));
        extraWidths = per.map((b) => b?.widths ?? []);
        extraAuto = per.map((b) => b?.autoWidths);
        numFmts = per.map((b) => b?.numFmts);
      }
      date1904 = biff ? biff.date1904 : !!w.Workbook?.WBProps?.date1904;
    } else {
      const ods = await prepareOds(bytes);
      extraWidths = ods.colWidths;
      extraAuto = ods.autoWidths;
      autoByName = ods.autoWidthsByName;
      const w = XLSX.read(ods.bytes, { type: 'array', cellNF: true, cellDates: false, cellFormula: false, cellHTML: false, cellStyles: true, sheetStubs: false, dense: true });
      wb = w;
      // .xlsb keeps raw 1904 serials (ODS dates are converted by SheetJS)
      date1904 = !ods.isOds && !!w.Workbook?.WBProps?.date1904;
      if (ods.hidden.size && w.SheetNames) {
        w.Workbook ??= {};
        w.Workbook.Sheets ??= w.SheetNames.map((name) => ({ name }));
        w.SheetNames.forEach((name, i) => {
          const entry = (w.Workbook!.Sheets![i] ??= { name });
          if (ods.hidden.has(name)) entry.Hidden = 1;
        });
      }
    }
  } catch (e) {
    if (e instanceof ReadBookError) throw e; // our own verdicts (restoreZip: corrupt / tooLarge) stand
    if (isPasswordError(e)) throw new ReadBookError('password', MSG_PASSWORD, { cause: e });
    throw new ReadBookError('corrupt', MSG_CORRUPT, { cause: e });
  }
  let sheets: ValuesSheet[];
  if (delimitedSheet) sheets = [delimitedSheet];
  else {
    if (!wb || !wb.SheetNames || !wb.SheetNames.length) throw new ReadBookError('corrupt', MSG_CORRUPT);
    sheets = sheetJsToValues(wb, kind === 'html', fileName, { numFmts, date1904, xml2003: kind === 'xml2003' });
  }
  sheets.forEach((s, i) => {
    if (!s.colWidths.length && extraWidths[i]?.length) s.colWidths = extraWidths[i];
    // which of the widths are auto widths: .xls / .ods by sheet index, .xlsb by the sheet's own name (SheetJS' name)
    const auto = extraAuto[i]?.size ? extraAuto[i] : wb?.SheetNames && autoByName.get(wb.SheetNames[i]);
    if (auto && s.colWidths.length && !s.autoWidth) s.autoWidth = auto;
  });
  if (kind === 'xml' && sheets.every((s) => !s.rows.length)) {
    // SheetJS "reads" any XML; a document without a single cell is not a spreadsheet we understand
    throw new ReadBookError('corrupt', MSG_CORRUPT);
  }
  const pkg = await buildValuesPackage(sheets);
  return { bytes: pkg, warnings, encoding, kind };
}

// ---------------------------------------------------------------------------------------------
// Minimal .xlsx package writer (values + number formats + merges + widths)
// ---------------------------------------------------------------------------------------------

const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS_PKG_R = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CT_PREFIX = 'application/vnd.openxmlformats-officedocument.spreadsheetml';

/** SheetJS reports built-in number formats with its en-US codes; map those codes back to their ids. */
function builtinIdOf(code: string): number | undefined {
  const table = XLSX.SSF.get_table() as Record<number, string>;
  for (let id = 0; id <= 49; id++) if (table[id] === code) return id;
  for (const [id, c] of Object.entries(JA_BUILTIN_NUMFMT)) if (c === code) return Number(id);
  return undefined;
}

function colName(n: number): string {
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function boxRef(b: RangeBox): string {
  const a = colName(b.c1) + b.r1;
  const z = colName(b.c2) + b.r2;
  return a === z ? a : `${a}:${z}`;
}

function numText(n: number): string {
  // shortest round-trip representation; Excel accepts xsd:double lexical forms (1E+21, 1.5E-7)
  return String(n).replace('e', 'E');
}

function strXml(s: string): string {
  const esc = escapeXmlText(encodeOoxmlEscapes(s));
  return /^[\s]|[\s]$|[\n\t]/.test(s) ? `<t xml:space="preserve">${esc}</t>` : `<t>${esc}</t>`;
}

export async function buildValuesPackage(sheets: ValuesSheet[]): Promise<Uint8Array> {
  if (!sheets.length) sheets = [{ name: 'Sheet1', rows: [], merges: [], colWidths: [] }];
  // number formats → xf indices
  const fmtXf = new Map<string, number>();
  const xfs: string[] = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'];
  const customFmts: string[] = [];
  let nextFmtId = 176;
  const xfOf = (fmt: string | undefined): number => {
    if (!fmt || fmt === 'General') return 0;
    let xf = fmtXf.get(fmt);
    if (xf !== undefined) return xf;
    let id = builtinIdOf(fmt);
    if (id === undefined) {
      id = nextFmtId++;
      customFmts.push(`<numFmt numFmtId="${id}" formatCode="${escapeXmlAttr(fmt)}"/>`);
    }
    xf = xfs.length;
    xfs.push(`<xf numFmtId="${id}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>`);
    fmtXf.set(fmt, xf);
    return xf;
  };

  const sst: string[] = [];
  const sstIndex = new Map<string, number>();
  let sstCount = 0;
  const zip = new JSZip();
  const firstVisible = Math.max(0, sheets.findIndex((s) => !s.hidden));
  const sheetEntries: string[] = [];
  const wbRels: string[] = [];
  const overrides: string[] = [];

  sheets.forEach((sh, i) => {
    const hidden = i === firstVisible ? 0 : sh.hidden ?? 0;
    let maxR = 0;
    let maxC = 0;
    const rowXml: string[] = [];
    for (let r = 1; r <= sh.rows.length; r++) {
      const row = sh.rows[r - 1];
      if (!row) continue;
      let cells = '';
      for (let c = 1; c <= row.length; c++) {
        const v = row[c - 1];
        if (!v) continue;
        const ref = colName(c) + r;
        const s = xfOf(v.fmt);
        const sAttr = s ? ` s="${s}"` : '';
        if (v.t === 's') {
          const text = String(v.v);
          let idx = sstIndex.get(text);
          if (idx === undefined) {
            idx = sst.length;
            sst.push(`<si>${strXml(text)}</si>`);
            sstIndex.set(text, idx);
          }
          sstCount++;
          cells += `<c r="${ref}"${sAttr} t="s"><v>${idx}</v></c>`;
        } else if (v.t === 'n') {
          cells += `<c r="${ref}"${sAttr}><v>${numText(v.v as number)}</v></c>`;
        } else if (v.t === 'b') {
          cells += `<c r="${ref}"${sAttr} t="b"><v>${v.v ? 1 : 0}</v></c>`;
        } else {
          cells += `<c r="${ref}"${sAttr} t="e"><v>${escapeXmlText(String(v.v))}</v></c>`;
        }
        if (c > maxC) maxC = c;
      }
      if (cells) {
        rowXml.push(`<row r="${r}">${cells}</row>`);
        maxR = r;
      }
    }
    let cols = '';
    sh.colWidths.forEach((w, ci) => {
      if (w !== undefined) cols += `<col min="${ci + 1}" max="${ci + 1}" width="${w}"${sh.autoWidth?.has(ci) ? '' : ' customWidth="1"'}/>`;
    });
    const dim = maxR && maxC ? `A1:${colName(maxC)}${maxR}` : 'A1';
    const merges = sh.merges.filter((m) => m.r2 > m.r1 || m.c2 > m.c1);
    const xml =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
      `<worksheet xmlns="${NS_MAIN}" xmlns:r="${NS_R}">` +
      `<dimension ref="${dim}"/>` +
      `<sheetViews><sheetView${i === firstVisible ? ' tabSelected="1"' : ''} workbookViewId="0"/></sheetViews>` +
      `<sheetFormatPr defaultRowHeight="18.75"/>` +
      (cols ? `<cols>${cols}</cols>` : '') +
      `<sheetData>${rowXml.join('')}</sheetData>` +
      (merges.length ? `<mergeCells count="${merges.length}">${merges.map((m) => `<mergeCell ref="${boxRef(m)}"/>`).join('')}</mergeCells>` : '') +
      `<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>` +
      `</worksheet>`;
    zip.file(`xl/worksheets/sheet${i + 1}.xml`, xml);
    sheetEntries.push(`<sheet name="${escapeXmlAttr(sh.name)}" sheetId="${i + 1}"${hidden === 1 ? ' state="hidden"' : hidden === 2 ? ' state="veryHidden"' : ''} r:id="rId${i + 1}"/>`);
    wbRels.push(`<Relationship Id="rId${i + 1}" Type="${NS_R}/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`);
    overrides.push(`<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="${CT_PREFIX}.worksheet+xml"/>`);
  });

  const n = sheets.length;
  wbRels.push(`<Relationship Id="rId${n + 1}" Type="${NS_R}/styles" Target="styles.xml"/>`);
  if (sst.length) wbRels.push(`<Relationship Id="rId${n + 2}" Type="${NS_R}/sharedStrings" Target="sharedStrings.xml"/>`);

  const workbook =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<workbook xmlns="${NS_MAIN}" xmlns:r="${NS_R}"><workbookPr/>` +
    `<bookViews><workbookView activeTab="${firstVisible}"/></bookViews>` +
    `<sheets>${sheetEntries.join('')}</sheets><calcPr calcId="191029"/></workbook>`;

  const styles =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<styleSheet xmlns="${NS_MAIN}">` +
    (customFmts.length ? `<numFmts count="${customFmts.length}">${customFmts.join('')}</numFmts>` : '') +
    `<fonts count="1"><font><sz val="11"/><color rgb="FF000000"/><name val="游ゴシック"/><family val="3"/><charset val="128"/></font></fonts>` +
    `<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>` +
    `<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>` +
    `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"><alignment vertical="center"/></xf></cellStyleXfs>` +
    `<cellXfs count="${xfs.length}">${xfs.map((x) => x.replace('/>', '><alignment vertical="center"/></xf>')).join('')}</cellXfs>` +
    `<cellStyles count="1"><cellStyle name="標準" xfId="0" builtinId="0"/></cellStyles>` +
    `<dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>` +
    `</styleSheet>`;

  const contentTypes =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
    `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
    `<Default Extension="xml" ContentType="application/xml"/>` +
    `<Override PartName="/xl/workbook.xml" ContentType="${CT_PREFIX}.sheet.main+xml"/>` +
    overrides.join('') +
    `<Override PartName="/xl/styles.xml" ContentType="${CT_PREFIX}.styles+xml"/>` +
    (sst.length ? `<Override PartName="/xl/sharedStrings.xml" ContentType="${CT_PREFIX}.sharedStrings+xml"/>` : '') +
    `</Types>`;

  zip.file('[Content_Types].xml', contentTypes);
  zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${NS_PKG_R}"><Relationship Id="rId1" Type="${NS_R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
  zip.file('xl/workbook.xml', workbook);
  zip.file('xl/_rels/workbook.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${NS_PKG_R}">${wbRels.join('')}</Relationships>`);
  zip.file('xl/styles.xml', styles);
  if (sst.length) {
    zip.file(
      'xl/sharedStrings.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<sst xmlns="${NS_MAIN}" count="${sstCount}" uniqueCount="${sst.length}">${sst.join('')}</sst>`,
    );
  }
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 6 } });
}
