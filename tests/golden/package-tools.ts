/**
 * Helpers for the package-level robustness tests of the patcher (round 4): take a valid package apart, damage or
 * re-express parts of it, put it back together, and check outputs for well-formedness.
 */
import JSZip from 'jszip';
import { xmlProblem } from '../../src/core/xlsx/patch/wellformed';
import { decodeBytes } from '../../src/core/xlsx/patch/xml';

export type Parts = Map<string, string | Uint8Array>;

/** all files of a package (text for xml / rels — UTF-8 or UTF-16, BOM removed —, bytes for everything else) */
export async function unpack(bytes: Uint8Array): Promise<Parts> {
  const zip = await JSZip.loadAsync(bytes);
  const out: Parts = new Map();
  for (const name of Object.keys(zip.files)) {
    const f = zip.files[name];
    if (f.dir) continue;
    const raw = await f.async('uint8array');
    out.set(name, /\.(xml|rels)$/i.test(name) ? decodeBytes(raw).text : raw);
  }
  return out;
}

/** UTF-16LE with a BOM and a matching encoding declaration (what .NET writers produce) */
export function utf16le(text: string): Uint8Array {
  const t = text.replace(/^(<\?xml[^>]*encoding=")[^"]*(")/i, '$1UTF-16$2');
  const out = new Uint8Array(2 + t.length * 2);
  out[0] = 0xff;
  out[1] = 0xfe;
  for (let i = 0; i < t.length; i++) {
    const c = t.charCodeAt(i);
    out[2 + 2 * i] = c & 255;
    out[3 + 2 * i] = c >> 8;
  }
  return out;
}

export async function pack(parts: Parts): Promise<Uint8Array> {
  const zip = new JSZip();
  for (const [name, data] of parts) zip.file(name, data);
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

/** unpack, edit, pack */
export async function edit(bytes: Uint8Array, fn: (parts: Parts) => unknown): Promise<Uint8Array> {
  const parts = await unpack(bytes);
  await fn(parts);
  return pack(parts);
}

export function text(parts: Parts, name: string): string {
  const v = parts.get(name);
  if (typeof v !== 'string') throw new Error(`no text part ${name}`);
  return v;
}

/** the SpreadsheetML elements of a default-namespace part written with a prefix (<x:worksheet xmlns:x=…><x:row>…) */
export function prefixMain(xml: string, prefix = 'x'): string {
  const main = /\sxmlns="(http:\/\/schemas\.openxmlformats\.org\/spreadsheetml\/2006\/main)"/.exec(xml)?.[1];
  if (!main) return xml;
  const named = xml.replace(/\sxmlns="[^"]*main"/, ` xmlns:${prefix}="${main}"`);
  return named.replace(/<(\/?)([A-Za-z_][\w.-]*)(?=[\s/>])/g, (_m, slash: string, name: string) => `<${slash}${prefix}:${name}`);
}

/** XML problems of every xml / rels part of a package ('' list = all well-formed) */
export async function packageProblems(bytes: Uint8Array): Promise<string[]> {
  const out: string[] = [];
  for (const [name, data] of await unpack(bytes)) {
    if (typeof data !== 'string') continue;
    const why = xmlProblem(data.replace(/^﻿/, ''));
    if (why) out.push(`${name}: ${why}`);
  }
  return out;
}

/** small deterministic PRNG (mulberry32) for the fuzz tests */
export function rng(seed: number): () => number {
  let a = seed >>> 0;
  return () => {
    a = (a + 0x6d2b79f5) >>> 0;
    let t = a;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}
