/**
 * OOXML package wrapper over JSZip (owner: patcher P).
 * Parts are read lazily and written back only when changed, so every untouched part keeps its bytes
 * (JSZip re-uses the compressed stream of unchanged entries).
 *
 * Zip-bomb safety (reader R's API, see the header of ../zip.ts): the package is opened with loadZip (refuses a
 * package that would expand beyond MAX_UNCOMPRESSED_TOTAL before anything is inflated) and every part is inflated
 * through readEntryBytes (never more bytes than the entry declares; XML parts capped at MAX_PART_BYTES because they
 * are decoded to a string). The JS engine's output is produced by generateAsync over what was loaded under these
 * guards, so nothing else can expand.
 */
import type JSZip from 'jszip';
import { MAX_PART_BYTES, ReadBookError, loadZip, readEntryBytes } from '../zip';
import { xmlProblem } from './wellformed';
import { decodeBytes, getAttr, normalizeAttrQuotes, normalizeRelPrefix, stripMainPrefix, unesc, escAttr } from './xml';

export interface Rel {
  id: string;
  type: string;
  target: string;
  external: boolean;
  raw: string;
}

/**
 * Every failure of the patcher is a PatchError whose `message` is the Japanese user-facing text. `internal` is set
 * when the cause was not a condition the patcher knows about (an unexpected exception, see toPatchError): the text is
 * still Japanese, the original error is kept in `cause` for the log.
 */
export class PatchError extends Error {
  internal = false;
  constructor(message: string) {
    super(message);
    this.name = 'PatchError';
  }
}

/** Anything thrown inside the patcher becomes a PatchError (a hostile or damaged file must never surface as a TypeError). */
export function toPatchError(e: unknown, context: string): PatchError {
  if (e instanceof PatchError) return e;
  const why = e instanceof Error ? e.message : String(e);
  const tooBig = e instanceof RangeError && /string length|array length|Invalid|allocation|memory/i.test(why);
  const pe = new PatchError(tooBig
    ? `${context}: ファイルが大きすぎるため処理できません。不要なシート・行・画像を削除するか、ファイルを分けて保存し直してください。`
    : `${context}: 想定外の内容のため処理できませんでした（ファイルが壊れているか、特殊な形式の可能性があります）。内部情報: ${why.slice(0, 160)}`);
  pe.internal = true;
  (pe as { cause?: unknown }).cause = e;
  return pe;
}

export function relsPathOf(partPath: string): string {
  const i = partPath.lastIndexOf('/');
  const dir = i >= 0 ? partPath.slice(0, i) : '';
  const base = i >= 0 ? partPath.slice(i + 1) : partPath;
  return (dir ? dir + '/' : '') + '_rels/' + base + '.rels';
}

export function dirOf(partPath: string): string {
  const i = partPath.lastIndexOf('/');
  return i >= 0 ? partPath.slice(0, i) : '';
}

/** Resolve a relationship target relative to the directory of the source part. */
export function resolveTarget(baseDir: string, target: string): string {
  let t = target;
  try { t = decodeURIComponent(target); } catch { /* keep raw */ }
  // some generators write Windows separators ("worksheets\sheet1.xml"); Excel reads them as '/'
  if (t.indexOf('\\') >= 0) t = t.replace(/\\/g, '/');
  if (t.startsWith('/')) return t.slice(1);
  const parts = (baseDir ? baseDir.split('/') : []).concat(t.split('/'));
  const out: string[] = [];
  for (const p of parts) {
    if (p === '..') out.pop();
    else if (p && p !== '.') out.push(p);
  }
  return out.join('/');
}

/** Relative target from a source dir to a part path (used when adding rels). */
export function relativeTarget(fromDir: string, partPath: string): string {
  const a = fromDir ? fromDir.split('/') : [];
  const b = partPath.split('/');
  let i = 0;
  while (i < a.length && i < b.length - 1 && a[i] === b[i]) i++;
  return '../'.repeat(a.length - i) + b.slice(i).join('/');
}

export function parseRels(xml: string): Rel[] {
  const out: Rel[] = [];
  for (const m of xml.matchAll(/<Relationship\b[^>]*>/g)) {
    const raw = m[0];
    out.push({
      raw,
      id: getAttr(raw, 'Id') ?? '',
      type: getAttr(raw, 'Type') ?? '',
      target: getAttr(raw, 'Target') ?? '',
      external: /TargetMode\s*=\s*"External"/.test(raw),
    });
  }
  return out;
}

/** relationship type check by its last segment (transitional and Strict OOXML namespaces) */
export function isRel(type: string, kind: string): boolean {
  return type.endsWith('/' + kind);
}

export const REL_TYPE = {
  worksheet: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
  sharedStrings: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings',
  styles: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
  theme: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
  calcChain: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain',
};

const CT_SST = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml';

export class Pkg {
  zip: JSZip;
  private cache = new Map<string, string>();
  private dirty = new Set<string>();
  private utf16 = new Set<string>();
  /** entries that were inflated successfully (read by the patcher, or checked by verifyUntouched) */
  private verified = new Set<string>();
  /** lower-case part name -> the name the zip / cache really uses (OPC part names are case-insensitive, writers differ in case) */
  private lower = new Map<string, string>();

  private constructor(zip: JSZip) {
    this.zip = zip;
    this.zip.forEach((p, f) => { if (!f.dir) this.lower.set(p.toLowerCase(), p); });
  }

  /**
   * The name a part is stored under: an exact match first, else the entry whose name differs only in case (Excel reads
   * "XL/Worksheets/Sheet1.xml" for a relationship target "worksheets/sheet1.xml"; a write must overwrite that entry,
   * not add a second one), else the name as given.
   */
  private canon(path: string): string {
    if (this.cache.has(path) || this.zip.file(path)) return path;
    return this.lower.get(path.toLowerCase()) ?? path;
  }

  static async load(bytes: Uint8Array): Promise<Pkg> {
    let zip: JSZip;
    try {
      zip = await loadZip(bytes);
    } catch (e) {
      // ReadBookError carries the Japanese user-facing text ('tooLarge': expands beyond 1.5 GB, 'corrupt': not a zip)
      if (e instanceof ReadBookError) throw new PatchError(e.message);
      throw new PatchError('ファイルを開けませんでした（Excel ブック .xlsx 形式ではないか、ファイルが壊れています）。');
    }
    if (!zip.file('[Content_Types].xml') || !zip.file('xl/workbook.xml')) {
      // workbook part may live elsewhere; resolved through the root rels in workbookPath()
      if (!zip.file('[Content_Types].xml')) throw new PatchError('Excel ブック（.xlsx）の構造が見つかりません。');
    }
    return new Pkg(zip);
  }

  has(path: string): boolean {
    const p = this.canon(path);
    return this.cache.has(p) || !!this.zip.file(p);
  }

  list(): string[] {
    const s = new Set<string>();
    this.zip.forEach((p, f) => { if (!f.dir) s.add(p); });
    for (const p of this.cache.keys()) s.add(p);
    return [...s];
  }

  async read(path0: string): Promise<string | null> {
    const path = this.canon(path0);
    const c = this.cache.get(path);
    if (c !== undefined) return c;
    const f = this.zip.file(path);
    if (!f) return null;
    const d = decodeBytes(await this.inflate(f, path, MAX_PART_BYTES));
    if (d.utf16) this.utf16.add(path);
    let text = d.text;
    if (/\.(xml|rels)$/i.test(path)) {
      // attributes in single quotes are read (and written back) in double quotes: the scanners know only name="value"
      text = normalizeAttrQuotes(text);
      if (/\.xml$/i.test(path)) {
        // SpreadsheetML with a namespace prefix (<x:row>…) is read in its default-namespace form and written back that way
        const n = stripMainPrefix(text);
        if (n === null) throw new PatchError(`ファイル内のデータ（${path}）の XML 名前空間の書き方には対応していません。Excel で開いて保存し直してから、もう一度お試しください。`);
        text = normalizeRelPrefix(n) ?? n;
      }
    }
    this.cache.set(path, text);
    return text;
  }

  async readBytes(path0: string): Promise<Uint8Array | null> {
    const path = this.canon(path0);
    const f = this.zip.file(path);
    return f ? this.inflate(f, path) : null;
  }

  /** Inflate one entry under reader R's guards (declared size cap, optional `maxBytes`); a broken entry is a PatchError. */
  private async inflate(f: JSZip.JSZipObject, path: string, maxBytes?: number): Promise<Uint8Array> {
    try {
      const b = await readEntryBytes(f, false, maxBytes);
      this.verified.add(path);
      return b;
    } catch {
      throw new PatchError(`ファイル内のデータ（${path}）を読み込めませんでした。ファイルが壊れているか、大きすぎる可能性があります。`);
    }
  }

  /**
   * The entries the patcher never touches are carried into the output as their compressed bytes (JSZip re-uses the
   * compressed stream). A damaged entry (broken deflate data, a size that does not match the directory) would come out
   * as a package Excel has to repair, so every entry that was not inflated yet is inflated once here (one at a time,
   * result dropped). A CRC-only mismatch is not refused: the data decodes to the declared size and the reader (zip.ts
   * loadZip) does not check CRCs either.
   */
  private async verifyUntouched(): Promise<void> {
    const names: string[] = [];
    this.zip.forEach((p, f) => { if (!f.dir && !this.dirty.has(p) && !this.verified.has(p)) names.push(p); });
    for (const p of names) {
      const f = this.zip.file(p);
      if (!f) continue;
      try {
        await readEntryBytes(f, false);
      } catch {
        throw new PatchError(`ファイル内のデータ（${p}）が壊れているため、出力できません。元のファイルが壊れている可能性があります。Excel でファイルを開いて保存し直してから、もう一度お試しください。`);
      }
      this.verified.add(p);
    }
  }

  /** Write a part (only marked dirty when the text differs from what was read). */
  write(path0: string, text: string): void {
    const path = this.canon(path0);
    const old = this.cache.get(path);
    if (old === text && this.zip.file(path)) return;
    this.cache.set(path, text);
    this.lower.set(path.toLowerCase(), path);
    this.dirty.add(path);
  }

  remove(path0: string): void {
    const path = this.canon(path0);
    this.cache.delete(path);
    this.dirty.delete(path);
    if (this.zip.file(path)) this.zip.remove(path);
    if (this.lower.get(path.toLowerCase()) === path) this.lower.delete(path.toLowerCase());
  }

  isDirty(path: string): boolean {
    return this.dirty.has(this.canon(path));
  }

  async rels(partPath: string): Promise<Rel[]> {
    const x = await this.read(relsPathOf(partPath));
    return x ? parseRels(x) : [];
  }

  /** Resolved internal targets of a part's relationships (path + type). */
  async relTargets(partPath: string): Promise<Array<Rel & { path: string }>> {
    const dir = dirOf(partPath);
    return (await this.rels(partPath)).filter((r) => !r.external).map((r) => ({ ...r, path: resolveTarget(dir, r.target) }));
  }

  async removeRel(partPath: string, pred: (r: Rel) => boolean): Promise<void> {
    const rp = relsPathOf(partPath);
    const x = await this.read(rp);
    if (!x) return;
    let y = x;
    for (const r of parseRels(x)) if (pred(r)) y = y.replace(r.raw, '');
    if (y !== x) this.write(rp, y);
  }

  async addRel(partPath: string, type: string, targetPath: string): Promise<string> {
    const rp = relsPathOf(partPath);
    let x = await this.read(rp);
    if (!x) x = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
    const ids = new Set(parseRels(x).map((r) => r.id));
    let n = ids.size + 1;
    while (ids.has('rId' + n)) n++;
    const id = 'rId' + n;
    const rel = `<Relationship Id="${id}" Type="${escAttr(type)}" Target="${escAttr(relativeTarget(dirOf(partPath), targetPath))}"/>`;
    x = x.replace(/<\/Relationships>\s*$/, rel + '</Relationships>');
    if (!x.includes(rel)) x = x.replace(/<Relationships\b([^>]*)\/>/, `<Relationships$1>${rel}</Relationships>`);
    this.write(rp, x);
    return id;
  }

  async setContentTypeOverride(partPath: string, contentType: string | null): Promise<void> {
    let x = (await this.read('[Content_Types].xml')) ?? '';
    const name = '/' + partPath;
    const re = new RegExp(`<Override\\b[^>]*PartName="${name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"[^>]*/>`, 'i');
    if (contentType === null) {
      x = x.replace(re, '');
    } else if (!re.test(x)) {
      x = x.replace(/<\/Types>\s*$/, `<Override PartName="${escAttr(name)}" ContentType="${escAttr(contentType)}"/></Types>`);
    }
    this.write('[Content_Types].xml', x);
  }

  async workbookPath(): Promise<string> {
    const rels = await this.read('_rels/.rels');
    if (rels) {
      for (const r of parseRels(rels)) {
        if (/\/officeDocument$/.test(r.type) && !r.external) return resolveTarget('', r.target);
      }
    }
    return 'xl/workbook.xml';
  }

  async ensureSharedStrings(wbPath: string): Promise<string> {
    const t = (await this.relTargets(wbPath)).find((r) => isRel(r.type, 'sharedStrings'));
    if (t) return t.path;
    const path = dirOf(wbPath) + '/sharedStrings.xml';
    this.write(path, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0"></sst>');
    await this.addRel(wbPath, REL_TYPE.sharedStrings, path);
    await this.setContentTypeOverride(path, CT_SST);
    return path;
  }

  /**
   * The package bytes. Every XML part the patcher wrote is checked for well-formedness first (wellformed.ts): damaged
   * input that the text patching cannot carry through ends as a PatchError, never as a package Excel has to repair.
   */
  async generate(): Promise<Uint8Array> {
    await this.verifyUntouched();
    for (const p of this.dirty) {
      let t = this.cache.get(p)!;
      if (/\.(xml|rels)$/i.test(p)) {
        const why = xmlProblem(t);
        if (why) throw new PatchError(`出力ファイルの検査で問題が見つかりました（${p}）。元のファイルの一部が壊れている可能性があります。${why}`);
      }
      // the text is stored as UTF-8 (JSZip): a UTF-16 part's declaration must say so (either quote style)
      if (this.utf16.has(p)) t = t.replace(/^(<\?xml[^>]*?\sencoding\s*=\s*)(["'])[^"']*\2/i, '$1$2UTF-8$2');
      this.zip.file(p, t);
    }
    this.dirty.clear();
    this.utf16.clear();
    // level 3: a 50,000 x 30 sheet (50 MB of XML) deflates in 2.8 s instead of 5.7 s (level 6) for a 12 % larger file (measured with
    // tests/fixtures/foreign/perf-big.ts); Excel does not care about the level
    return this.zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 3 } });
  }
}

export { unesc };
