/**
 * Native deflate + CRC-32 for the big parts of an output package (owner: patcher P).
 *
 * JSZip compresses with pako (JavaScript): about 20 MB/s, so the sheets of a workbook with ten copied 20,000 x 30 sheets (250 MB of XML)
 * need 12 s of the export. Node's zlib (a native thread pool: several parts at once) and the browsers' CompressionStream do the same work
 * 3 - 10 times faster. Pkg.generate hands the finished entries to JSZip as ready-made compressed entries (what JSZip itself does for the
 * untouched parts of a loaded package: a CompressedObject); anything unexpected falls back to JSZip's own compression.
 */
import type JSZip from 'jszip';

export interface NativeCodec {
  /** raw deflate (RFC 1951), no zlib / gzip header */
  deflate(bytes: Uint8Array): Promise<Uint8Array>;
  /** raw inflate that never produces more than `max` bytes (rejects); only where Node's zlib is available */
  inflate?(bytes: Uint8Array, max: number): Promise<Uint8Array>;
  crc32(bytes: Uint8Array): number;
}

let codec: NativeCodec | null | undefined;

export function nativeCodec(): NativeCodec | null {
  if (codec === undefined) codec = detect();
  return codec;
}

type Zlib = {
  deflateRaw(buf: Uint8Array, opts: { level: number }, cb: (e: Error | null, out: Uint8Array) => void): void;
  inflateRaw?(buf: Uint8Array, opts: { maxOutputLength: number }, cb: (e: Error | null, out: Uint8Array) => void): void;
  crc32?: (data: Uint8Array, value?: number) => number;
};

function detect(): NativeCodec | null {
  try {
    const proc = (globalThis as { process?: { versions?: { node?: string }; getBuiltinModule?: (id: string) => unknown } }).process;
    if (proc?.versions?.node && typeof proc.getBuiltinModule === 'function') {
      const z = proc.getBuiltinModule('node:zlib') as Zlib | undefined;
      if (z && typeof z.deflateRaw === 'function') {
        return {
          // level 3 like the JSZip path: a 50 MB sheet compresses in a fraction of a second, 12 % larger than level 6
          deflate: (b) => new Promise((res, rej) => z.deflateRaw(b, { level: 3 }, (e, out) => (e ? rej(e) : res(new Uint8Array(out.buffer, out.byteOffset, out.length))))),
          inflate: typeof z.inflateRaw === 'function' ? (b, max) => new Promise((res, rej) => z.inflateRaw!(b, { maxOutputLength: max }, (e, out) => (e ? rej(e) : res(new Uint8Array(out.buffer, out.byteOffset, out.length))))) : undefined,
          crc32: typeof z.crc32 === 'function' ? (b) => z.crc32!(b) >>> 0 : jsCrc32,
        };
      }
    }
    const g = globalThis as { CompressionStream?: new (format: string) => unknown; Blob?: new (parts: unknown[]) => { stream(): { pipeThrough(t: unknown): ReadableStream } }; Response?: new (body: unknown) => { arrayBuffer(): Promise<ArrayBuffer> } };
    if (typeof g.CompressionStream === 'function' && typeof g.Blob === 'function' && typeof g.Response === 'function') {
      const CS = g.CompressionStream, B = g.Blob, R = g.Response;
      return {
        deflate: async (b) => new Uint8Array(await new R(new B([b]).stream().pipeThrough(new CS('deflate-raw'))).arrayBuffer()),
        crc32: jsCrc32,
      };
    }
  } catch {
    /* no native codec: JSZip compresses */
  }
  return null;
}

let table: Int32Array | null = null;

/** CRC-32 (IEEE 802.3), table driven */
export function jsCrc32(bytes: Uint8Array): number {
  if (!table) {
    table = new Int32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      table[n] = c;
    }
  }
  const t = table;
  let c = -1;
  for (let i = 0; i < bytes.length; i++) c = t[(c ^ bytes[i]) & 0xff] ^ (c >>> 8);
  return (c ^ -1) >>> 0;
}

interface Template {
  /** the prototype of JSZip's CompressedObject (its `constructor` property is Object: the class itself is not reachable from an instance) */
  proto: object;
  compression: unknown;
}

/**
 * The CompressedObject prototype and the DEFLATE compression object of this JSZip, taken from an entry that was loaded from a package (so
 * no deep import of JSZip's internals is needed); null when the zip has no deflated entry or JSZip looks different.
 */
export function deflateTemplate(zip: JSZip): Template | null {
  for (const f of Object.values(zip.files)) {
    const d = (f as unknown as { _data?: { compression?: { magic?: string }; compressedSize?: unknown } })._data;
    if (d && d.compression && d.compression.magic === '\x08\x00' && typeof d.compressedSize === 'number') {
      const proto = Object.getPrototypeOf(d) as object | null;
      if (proto && proto !== Object.prototype && typeof (proto as { getCompressedWorker?: unknown }).getCompressedWorker === 'function') return { proto, compression: d.compression };
    }
  }
  return null;
}

export interface Ready {
  path: string;
  entry: unknown;
}

/**
 * Compress `parts` (path, text) natively, up to `concurrency` at a time; the entries come back in the order of `parts`. A part that fails is
 * missing from the result (the caller adds it as plain text).
 */
export async function compressParts(zip: JSZip, parts: Array<{ path: string; text: string }>, concurrency = 4): Promise<Ready[] | null> {
  const c = nativeCodec();
  const tpl = deflateTemplate(zip);
  if (!c || !tpl) return null;
  const enc = new TextEncoder();
  const out: Array<Ready | undefined> = new Array(parts.length);
  let next = 0;
  const run = async () => {
    for (;;) {
      const i = next++;
      if (i >= parts.length) return;
      try {
        const bytes = enc.encode(parts[i].text);
        const comp = await c.deflate(bytes);
        const entry = Object.create(tpl.proto) as Record<string, unknown>;
        entry.compressedSize = comp.length;
        entry.uncompressedSize = bytes.length;
        entry.crc32 = c.crc32(bytes);
        entry.compression = tpl.compression;
        entry.compressedContent = comp;
        out[i] = { path: parts[i].path, entry };
      } catch {
        out[i] = undefined;
      }
    }
  };
  await Promise.all(Array.from({ length: Math.min(concurrency, parts.length) }, run));
  return out.filter((x): x is Ready => !!x);
}

/** entries smaller than this are inflated by JSZip (the thread-pool round trip costs more than it saves) */
const NATIVE_INFLATE_MIN = 1 << 18;

/**
 * Inflate a big deflated entry with Node's zlib (about 4 times faster than pako, and off the main thread) under the same guard as
 * zip.ts readEntryBytes: nothing beyond the size the entry declares, an entry that declares more than `maxBytes` is left to readEntryBytes
 * (which refuses it). null = not applicable or anything unexpected (wrong size, broken data): the caller inflates with readEntryBytes, which
 * reports the problem the usual way.
 */
export async function nativeInflateEntry(file: JSZip.JSZipObject, maxBytes?: number): Promise<Uint8Array | null> {
  const c = nativeCodec();
  if (!c || !c.inflate) return null;
  const d = (file as unknown as { _data?: { compression?: { magic?: string }; uncompressedSize?: unknown; compressedContent?: unknown } })._data;
  if (!d || !d.compression || d.compression.magic !== '\x08\x00') return null;
  const declared = d.uncompressedSize;
  const content = d.compressedContent;
  if (typeof declared !== 'number' || !Number.isFinite(declared) || declared < NATIVE_INFLATE_MIN || (maxBytes !== undefined && declared > maxBytes)) return null;
  if (!(content instanceof Uint8Array)) return null;
  try {
    const out = await c.inflate(content, declared);
    return out.length === declared ? out : null;
  } catch {
    return null;
  }
}
