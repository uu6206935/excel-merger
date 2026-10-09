/**
 * OOXML package access (owner: reader R): ZIP loading with Japanese error messages, encryption / format sniffing,
 * part-path resolution, relationships and [Content_Types].xml. Pure TS (JSZip), browser + Node.
 *
 * ZIP-BOMB SAFETY FOR THE PATCHER (patch/pkg.ts, owner P) — the small API to call, in this order:
 *
 *   1. `const zip = await loadZip(bytes)` instead of `JSZip.loadAsync(bytes)`. It refuses, before anything is
 *      inflated, a package whose central directory declares more than MAX_UNCOMPRESSED_TOTAL (1.5 GB) in total
 *      (overlapping entries, ZIP64 sizes, >65,535 entries and 32-bit sizes >= 2 GB included — see
 *      zipDirectoryInfo) and throws `ReadBookError`: `.code === 'tooLarge'` (`.message` is the Japanese
 *      user-facing text, msgTooLarge) or `.code === 'corrupt'` (MSG_CORRUPT). Map it to your own error, e.g.
 *      `catch (e) { throw new PatchError(e instanceof ReadBookError ? e.message : …) }`. (If you keep your own
 *      JSZip.loadAsync, call `assertExpandable(bytes)` before it and `assertExpandable(bytes, zip)` after it: same
 *      ReadBookError, JSZip's own view of the sizes added.)
 *   2. Read every part through `readEntryBytes(zip.files[name], checkCrc?, maxBytes?)` instead of
 *      `file.async('uint8array')`. It never yields more bytes than the entry declares (a lying entry — small
 *      declared size, huge deflate stream — stops right after the declared size instead of filling memory),
 *      rejects with an Error for bad deflate data / size mismatch / CRC mismatch (`checkCrc`), and rejects at once,
 *      before inflating, when the declared size exceeds `maxBytes`. Pass `maxBytes = MAX_PART_BYTES` for parts that
 *      are decoded to a string (every XML part): V8 cannot hold a longer string, so a bigger part is refused up
 *      front instead of failing inside TextDecoder. Works on entries JSZip holds in memory too (added with
 *      `zip.file(name, string | Uint8Array)`): no declared size is known for them, the 1.5 GB total is the cap.
 *   3. Text of a part: `decodePartBytes(u8)` (UTF-8 / UTF-16 BOM aware) if you want the reader's decoding; P's own
 *      decoder is fine too.
 *
 * Nothing else is needed: the JS engine's output is produced by `zip.generateAsync`, which only re-compresses what
 * was loaded under the guards above.
 */
import JSZip from 'jszip';
import { kids, parseXml, rootOf, attr } from './xml';

/** User-facing error messages (Japanese). Exported so UI / tests can match them. */
export const MSG_PASSWORD = 'パスワード付きのファイルは読み込めません。Excelでパスワードを解除して保存し直してください。';
export const MSG_CORRUPT = 'Excelファイルとして読み込めませんでした（ファイルが壊れているか、Excel形式ではありません）';
export const MSG_EMPTY = 'ファイルが空です（0バイト）。';

/**
 * Zip-bomb guard: a package whose parts would expand to more than this many bytes in total is refused before
 * anything is inflated (1.5 GB = 1.5 × 2^30 bytes).
 */
export const MAX_UNCOMPRESSED_TOTAL = 1.5 * 1024 * 1024 * 1024;

/**
 * Largest XML part (bytes) that is decoded to a string: V8's maximum string length (2^29 − 24 UTF-16 units, the
 * same in Node, Chrome and Edge). A UTF-8 part of at most this many bytes always fits in one string; larger parts
 * are refused up front instead of failing inside TextDecoder (a failure there would silently read as "missing").
 */
export const MAX_PART_BYTES = 0x1fffffe8;

/** "約 1.8 GB" / "約 620 MB" (1 GB = 2^30 bytes, as Explorer shows sizes). */
export function sizeText(bytes: number): string {
  if (!Number.isFinite(bytes)) return '2 GB 以上';
  const gb = bytes / 1024 ** 3;
  if (gb >= 1) return `約 ${gb.toFixed(1)} GB`;
  return `約 ${Math.max(1, Math.round(bytes / 1024 ** 2))} MB`;
}

/** Message for a package that would expand beyond MAX_UNCOMPRESSED_TOTAL. */
export function msgTooLarge(total: number): string {
  return `ファイルを展開すると大きすぎるため読み込めません（展開後 ${sizeText(total)}、上限 1.5 GB）。不要なシート・行・画像を削除するか、ファイルを分けて保存し直してください。`;
}

/** Message for a text input (CSV / TSV / TXT / HTML) too large to decode (see MAX_PART_BYTES). */
export function msgTextTooLarge(bytes: number): string {
  return `ファイルが大きすぎるため読み込めません（${sizeText(bytes)}、テキスト形式の上限 約 512 MB）。ファイルを分けて保存し直してください。`;
}

/** Message for a shared-strings part too large to read (see MAX_PART_BYTES). */
export function msgSharedStringsTooLarge(bytes: number): string {
  return `文字データ（共有文字列）が大きすぎるため読み込めません（展開後 ${sizeText(bytes)}、上限 約 512 MB）。不要なシートや行を削除するか、ファイルを分けて保存し直してください。`;
}

/** Error thrown by the reader; `message` is the Japanese user-facing text. */
export class ReadBookError extends Error {
  readonly code: 'password' | 'corrupt' | 'empty' | 'unsupported' | 'tooLarge';
  constructor(code: ReadBookError['code'], message: string, options?: { cause?: unknown }) {
    super(message);
    this.name = 'ReadBookError';
    this.code = code;
    if (options?.cause !== undefined) (this as { cause?: unknown }).cause = options.cause;
  }
}

export type Sniffed = 'zip' | 'cfb' | 'text' | 'empty' | 'unknown';

const CFB_MAGIC = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];

/** Identify the container by magic bytes. */
export function sniff(bytes: Uint8Array): Sniffed {
  if (bytes.length === 0) return 'empty';
  if (bytes.length >= 4 && bytes[0] === 0x50 && bytes[1] === 0x4b && (bytes[2] === 3 || bytes[2] === 5 || bytes[2] === 7) && (bytes[3] === 4 || bytes[3] === 6 || bytes[3] === 8)) return 'zip';
  if (bytes.length >= 8 && CFB_MAGIC.every((b, i) => bytes[i] === b)) return 'cfb';
  // Heuristic: text if the first 4 KB contain no NUL bytes (UTF-16 with BOM is also text)
  if ((bytes[0] === 0xff && bytes[1] === 0xfe) || (bytes[0] === 0xfe && bytes[1] === 0xff)) return 'text';
  const n = Math.min(bytes.length, 4096);
  for (let i = 0; i < n; i++) if (bytes[i] === 0) return 'unknown';
  return 'text';
}

/**
 * True when a CFB (OLE2) container holds an encrypted OOXML package (Excel "パスワードを使用して暗号化").
 * Directory entry names are UTF-16LE, so search for "EncryptedPackage" / "EncryptionInfo" encoded that way.
 */
export function isEncryptedCfb(bytes: Uint8Array): boolean {
  return containsUtf16le(bytes, 'EncryptedPackage') || containsUtf16le(bytes, 'EncryptionInfo');
}

function containsUtf16le(bytes: Uint8Array, word: string): boolean {
  const pat: number[] = [];
  for (let i = 0; i < word.length; i++) pat.push(word.charCodeAt(i), 0);
  const first = pat[0];
  const last = bytes.length - pat.length;
  outer: for (let i = 0; i <= last; i++) {
    if (bytes[i] !== first) continue;
    for (let k = 1; k < pat.length; k++) if (bytes[i + k] !== pat[k]) continue outer;
    return true;
  }
  return false;
}

// ---------------------------------------------------------------------------------------------
// Zip-bomb guard
// ---------------------------------------------------------------------------------------------

export interface ZipDirectoryInfo {
  /** central directory entries (files and folders) */
  entries: number;
  /** sum of the uncompressed sizes the central directory declares (ZIP64 sizes included) */
  uncompressedTotal: number;
}

/**
 * Own read of the zip central directory (EOCD, ZIP64 EOCD, per-entry sizes incl. the ZIP64 extra field), used for
 * the zip-bomb guard before JSZip inflates anything. JSZip reads 4/8-byte sizes with 32-bit shifts, so sizes of
 * 2 GB and more come back negative or wrapped (a 5 GB ZIP64 entry reads as 1 GB); this parser does 53-bit math.
 * Overlapping entries (several directory records pointing at the same data — the "non-recursive zip bomb") are
 * each counted with their declared size, like an unzipper would inflate them. Returns undefined when no
 * well-formed central directory is found (JSZip then decides whether the file is a zip at all).
 */
export function zipDirectoryInfo(bytes: Uint8Array): ZipDirectoryInfo | undefined {
  const n = bytes.length;
  const u16 = (p: number) => bytes[p] | (bytes[p + 1] << 8);
  const u32 = (p: number) => (bytes[p] | (bytes[p + 1] << 8) | (bytes[p + 2] << 16)) + bytes[p + 3] * 0x1000000;
  const u64 = (p: number) => u32(p) + u32(p + 4) * 0x100000000;
  const sigAt = (p: number, sig: number) => p >= 0 && p + 4 <= n && u32(p) === sig;
  // End of central directory record: the last one within the final 22 + 65535 bytes
  let eocd = -1;
  for (let p = n - 22; p >= 0 && p >= n - 22 - 0xffff; p--) {
    if (bytes[p] === 0x50 && sigAt(p, 0x06054b50)) {
      eocd = p;
      break;
    }
  }
  if (eocd < 0) return undefined;
  let count = u16(eocd + 10);
  let cdSize = u32(eocd + 12);
  let cdOffset = u32(eocd + 16);
  let dirEnd = eocd;
  // ZIP64: locator right before the EOCD → ZIP64 EOCD record
  if (sigAt(eocd - 20, 0x07064b50)) {
    const z = u64(eocd - 20 + 8);
    const at = sigAt(z, 0x06064b50) ? z : eocd - 20 - 56; // prepended data shifts offsets: try just before the locator
    if (sigAt(at, 0x06064b50) && at + 56 <= n) {
      count = u64(at + 32);
      cdSize = u64(at + 40);
      cdOffset = u64(at + 48);
      dirEnd = at;
    }
  }
  // prepended data (self-extractor stubs …): the directory then ends right before the (ZIP64) EOCD
  let p = cdOffset;
  if (!sigAt(p, 0x02014b50)) p = dirEnd - cdSize;
  if (!sigAt(p, 0x02014b50)) return count > 0 ? undefined : { entries: 0, uncompressedTotal: 0 };
  let total = 0;
  let entries = 0;
  // every record up to the EOCD, whatever the (16-bit, possibly wrapped) entry count says — as JSZip reads them
  for (; sigAt(p, 0x02014b50) && p + 46 <= dirEnd; entries++) {
    let usize = u32(p + 24);
    const nameLen = u16(p + 28);
    const extraLen = u16(p + 30);
    const commentLen = u16(p + 32);
    if (usize === 0xffffffff) {
      // ZIP64 extended information extra field (id 0x0001): its first field is the 8-byte uncompressed size
      let e = p + 46 + nameLen;
      const eEnd = Math.min(n, e + extraLen);
      while (e + 4 <= eEnd) {
        const id = u16(e);
        const len = u16(e + 2);
        if (id === 0x0001 && len >= 8 && e + 12 <= eEnd) {
          usize = u64(e + 4);
          break;
        }
        e += 4 + len;
      }
    }
    total += usize;
    p += 46 + nameLen + extraLen + commentLen;
  }
  return { entries, uncompressedTotal: total };
}

/** The declared uncompressed size of a zip entry loaded by JSZip (internal field); undefined when unknown. */
export function entrySize(file: JSZip.JSZipObject): number | undefined {
  const d = (file as unknown as { _data?: { uncompressedSize?: unknown } })._data;
  const n = d && typeof d === 'object' ? d.uncompressedSize : undefined;
  return typeof n === 'number' && Number.isFinite(n) ? n : undefined;
}

/**
 * Refuse (ReadBookError 'tooLarge') a zip whose parts would expand beyond MAX_UNCOMPRESSED_TOTAL, from the central
 * directory, before any part is inflated. `zip` (already loaded) adds JSZip's own view of the sizes, in case its
 * directory search found other entries (a negative JSZip size means ≥ 2 GB).
 */
export function assertExpandable(bytes: Uint8Array, zip?: JSZip): void {
  let total = zipDirectoryInfo(bytes)?.uncompressedTotal ?? 0;
  if (zip) {
    let jt = 0;
    for (const name of Object.keys(zip.files)) {
      const s = entrySize(zip.files[name]);
      if (s !== undefined) jt += s < 0 ? Infinity : s;
    }
    total = Math.max(total, jt);
  }
  if (total > MAX_UNCOMPRESSED_TOTAL) throw new ReadBookError('tooLarge', msgTooLarge(total));
}

export async function loadZip(bytes: Uint8Array): Promise<JSZip> {
  assertExpandable(bytes);
  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(bytes, { checkCRC32: false });
  } catch (e) {
    throw new ReadBookError('corrupt', MSG_CORRUPT, { cause: e });
  }
  assertExpandable(bytes, zip);
  normalizeEntryNames(zip);
  return zip;
}

/**
 * Entry names written with backslashes (`xl\worksheets\sheet1.xml`: .NET Framework's ZipFile and Windows PowerShell 5.1's
 * Compress-Archive store Windows path separators; Excel opens such a file) are renamed to the '/' form every part lookup uses.
 * The entry object is kept (its compressed stream is reused by generateAsync), only its name changes. An entry whose
 * normalized name already exists is left alone.
 */
function normalizeEntryNames(zip: JSZip): void {
  const files = zip.files as Record<string, JSZip.JSZipObject & { name: string }>;
  for (const name of Object.keys(files)) {
    if (name.indexOf('\\') < 0) continue;
    const entry = files[name];
    const norm = name.replace(/\\/g, '/').replace(/^\/+/, '');
    if (!norm || files[norm]) continue;
    entry.name = norm;
    files[norm] = entry;
    delete files[name];
  }
}

/** CRC-32 (IEEE 802.3, the zip checksum) of bytes, unsigned. */
const CRC_TABLE = (() => {
  const t = new Int32Array(256);
  for (let i = 0; i < 256; i++) {
    let c = i;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[i] = c;
  }
  return t;
})();

export function crc32(bytes: Uint8Array): number {
  let c = -1;
  for (let i = 0; i < bytes.length; i++) c = CRC_TABLE[(c ^ bytes[i]) & 0xff] ^ (c >>> 8);
  return (c ^ -1) >>> 0;
}

/** Why a part could not be read: absent, broken (bad deflate data, size mismatch, CRC), or too large to decode. */
export type PartProblem = 'missing' | 'corrupt' | 'tooLarge';

export type PartRead = { ok: true; text: string } | { ok: false; problem: PartProblem };

/**
 * Inflate one zip entry, never producing more bytes than it declares (a lying entry — small declared size, huge
 * deflate stream — is stopped at its declared size instead of filling memory; JSZip itself would only notice the
 * mismatch after inflating everything). `checkCrc`: verify the CRC-32 the directory declares. `maxBytes`: refuse
 * (reject before inflating) an entry that declares more, and never inflate more than that (pass MAX_PART_BYTES
 * for parts that are decoded to a string).
 * Rejects with an Error for broken or oversized entries.
 */
export function readEntryBytes(file: JSZip.JSZipObject, checkCrc = false, maxBytes?: number): Promise<Uint8Array> {
  const declared = entrySize(file);
  if (maxBytes !== undefined && declared !== undefined && (declared < 0 || declared > maxBytes)) {
    return Promise.reject(new Error(`zip entry ${file.name}: larger than the allowed ${maxBytes} bytes`));
  }
  const declaredCap = declared !== undefined && declared >= 0 ? declared : MAX_UNCOMPRESSED_TOTAL;
  const cap = maxBytes !== undefined ? Math.min(declaredCap, maxBytes) : declaredCap;
  return new Promise<Uint8Array>((resolve, reject) => {
    const chunks: Uint8Array[] = [];
    let total = 0;
    let settled = false;
    const fail = (why: string): void => {
      if (settled) return;
      settled = true;
      chunks.length = 0;
      reject(new Error(`zip entry ${file.name}: ${why}`));
    };
    let helper: JSZip.JSZipStreamHelper<Uint8Array>;
    try {
      helper = (file as unknown as { internalStream(t: 'uint8array'): JSZip.JSZipStreamHelper<Uint8Array> }).internalStream('uint8array');
    } catch (e) {
      fail(String(e));
      return;
    }
    helper
      .on('data', (chunk) => {
        if (settled) return;
        total += chunk.length;
        if (total > cap) {
          helper.pause();
          fail(maxBytes !== undefined && cap === maxBytes && maxBytes < declaredCap ? `larger than the allowed ${maxBytes} bytes` : 'inflates beyond its declared size');
          return;
        }
        chunks.push(chunk);
      })
      .on('error', (e) => fail(String(e)))
      .on('end', () => {
        if (settled) return;
        if (declared !== undefined && total !== declared) return fail('size mismatch');
        let out: Uint8Array;
        if (chunks.length === 1) out = chunks[0];
        else {
          out = new Uint8Array(total);
          let o = 0;
          for (const c of chunks) {
            out.set(c, o);
            o += c.length;
          }
        }
        if (checkCrc) {
          const want = (file as unknown as { _data?: { crc32?: unknown } })._data?.crc32;
          if (typeof want === 'number' && crc32(out) !== want >>> 0) return fail('CRC mismatch');
        }
        settled = true;
        resolve(out);
      })
      .resume();
  });
}

/** Case-insensitive part lookup (OPC part names are case-insensitive; some writers differ in case). */
export class Package {
  readonly zip: JSZip;
  private readonly lower = new Map<string, string>();
  private readonly cache = new Map<string, Promise<PartRead>>();

  constructor(zip: JSZip) {
    this.zip = zip;
    for (const name of Object.keys(zip.files)) {
      if (!zip.files[name].dir) this.lower.set(name.toLowerCase(), name);
    }
  }

  /** All part names (no leading '/'), directories excluded. */
  get partNames(): string[] {
    return [...this.lower.values()];
  }

  resolveName(path: string): string | undefined {
    const p = path.replace(/^\/+/, '');
    if (this.zip.files[p] && !this.zip.files[p].dir) return p;
    return this.lower.get(p.toLowerCase());
  }

  has(path: string): boolean {
    return this.resolveName(path) !== undefined;
  }

  /** Declared uncompressed size of a part (bytes), undefined when missing / unknown. */
  size(path: string): number | undefined {
    const name = this.resolveName(path);
    return name ? entrySize(this.zip.files[name]) : undefined;
  }

  /** Part text (UTF-8/UTF-16 decoded, BOM stripped) or null when missing / unreadable. Cached. */
  text(path: string): Promise<string | null> {
    return this.load(path).then((r) => (r.ok ? r.text : null));
  }

  /**
   * Part text, or why it could not be read: 'missing', 'corrupt' (bad deflate data, size mismatch, undecodable),
   * 'tooLarge' (declared size above MAX_PART_BYTES: refused before inflating). Cached.
   */
  load(path: string): Promise<PartRead> {
    const name = this.resolveName(path);
    if (!name) return Promise.resolve({ ok: false, problem: 'missing' });
    let p = this.cache.get(name);
    if (!p) {
      const file = this.zip.files[name];
      const declared = entrySize(file);
      if (declared !== undefined && (declared < 0 || declared > MAX_PART_BYTES)) p = Promise.resolve({ ok: false, problem: 'tooLarge' });
      else
        p = readEntryBytes(file).then(
          (u8): PartRead => {
            if (u8.length > MAX_PART_BYTES) return { ok: false, problem: 'tooLarge' };
            try {
              return { ok: true, text: decodePartBytes(u8) };
            } catch {
              return { ok: false, problem: 'corrupt' };
            }
          },
          (): PartRead => ({ ok: false, problem: 'corrupt' }),
        );
      this.cache.set(name, p);
    }
    return p;
  }

  /** Drop a cached part text (large sheets are read once). */
  forget(path: string): void {
    const name = this.resolveName(path);
    if (name) this.cache.delete(name);
  }
}

/** Decode an XML part: honour UTF-16 BOMs, default UTF-8, strip BOM. */
export function decodePartBytes(u8: Uint8Array): string {
  if (u8.length >= 2 && u8[0] === 0xff && u8[1] === 0xfe) return new TextDecoder('utf-16le').decode(u8.subarray(2));
  if (u8.length >= 2 && u8[0] === 0xfe && u8[1] === 0xff) return new TextDecoder('utf-16be').decode(u8.subarray(2));
  const s = new TextDecoder('utf-8').decode(u8);
  return s.charCodeAt(0) === 0xfeff ? s.slice(1) : s;
}

// ---------------------------------------------------------------------------------------------
// Paths and relationships
// ---------------------------------------------------------------------------------------------

/** Directory part of a part name ("xl/worksheets/sheet1.xml" -> "xl/worksheets"). */
export function dirOf(path: string): string {
  const i = path.lastIndexOf('/');
  return i < 0 ? '' : path.slice(0, i);
}

/** Resolve a relationship target against the source part's directory ("../tables/t1.xml", "/xl/x.xml"). */
export function resolveTarget(baseDir: string, target: string): string {
  let t = target;
  try {
    t = decodeURIComponent(target);
  } catch {
    /* keep raw */
  }
  const parts = t.startsWith('/') ? [] : baseDir.split('/').filter(Boolean);
  for (const seg of t.replace(/^\/+/, '').split('/')) {
    if (seg === '' || seg === '.') continue;
    if (seg === '..') parts.pop();
    else parts.push(seg);
  }
  return parts.join('/');
}

/** rels part of a part ("xl/workbook.xml" -> "xl/_rels/workbook.xml.rels"). */
export function relsPathOf(partPath: string): string {
  const d = dirOf(partPath);
  const f = partPath.slice(d.length ? d.length + 1 : 0);
  return (d ? d + '/' : '') + '_rels/' + f + '.rels';
}

export interface Rel {
  id: string;
  /** full type URI */
  type: string;
  /** last path segment of the type ("worksheet", "styles", "hyperlink" ...) — Strict and Transitional agree on it */
  kind: string;
  /** raw Target attribute */
  target: string;
  external: boolean;
  /** resolved part path for internal targets (no leading '/') */
  path: string;
}

export async function readRels(pkg: Package, partPath: string): Promise<Rel[]> {
  const xml = await pkg.text(relsPathOf(partPath));
  if (!xml) return [];
  const root = rootOf(parseXml(xml));
  const base = dirOf(partPath);
  const out: Rel[] = [];
  for (const r of kids(root, 'Relationship')) {
    const type = attr(r, 'Type') ?? '';
    const target = attr(r, 'Target') ?? '';
    const external = (attr(r, 'TargetMode') ?? '').toLowerCase() === 'external';
    out.push({
      id: attr(r, 'Id') ?? '',
      type,
      kind: type.slice(type.lastIndexOf('/') + 1),
      target,
      external,
      path: external ? '' : resolveTarget(base, target),
    });
  }
  return out;
}

export interface ContentTypes {
  defaults: Map<string, string>;
  overrides: Map<string, string>;
}

export async function readContentTypes(pkg: Package): Promise<ContentTypes> {
  const ct: ContentTypes = { defaults: new Map(), overrides: new Map() };
  const xml = await pkg.text('[Content_Types].xml');
  if (!xml) return ct;
  const root = rootOf(parseXml(xml));
  for (const d of kids(root, 'Default')) ct.defaults.set((attr(d, 'Extension') ?? '').toLowerCase(), attr(d, 'ContentType') ?? '');
  for (const o of kids(root, 'Override')) ct.overrides.set((attr(o, 'PartName') ?? '').replace(/^\/+/, '').toLowerCase(), attr(o, 'ContentType') ?? '');
  return ct;
}

export function contentTypeOf(ct: ContentTypes, partPath: string): string {
  const o = ct.overrides.get(partPath.toLowerCase());
  if (o) return o;
  const dot = partPath.lastIndexOf('.');
  return dot < 0 ? '' : ct.defaults.get(partPath.slice(dot + 1).toLowerCase()) ?? '';
}
