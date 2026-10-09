/**
 * Files dropped onto the window. A dropped FOLDER arrives in `dataTransfer.files` as one useless entry (no extension, size 0 / 4096), so
 * the drop handler hands over the entries of `DataTransferItem.webkitGetAsEntry()` (which are only valid DURING the drop event: take
 * them synchronously) and this module walks the folders for the Excel files inside.
 * Pure TypeScript with minimal structural types, so it runs in vitest with fake entries.
 */
import { isAcceptedFileName, isExcelLockFile } from './helpers';

/** The part of FileSystemEntry / FileSystemFileEntry / FileSystemDirectoryEntry used here. */
export interface EntryLike {
  isFile: boolean;
  isDirectory: boolean;
  name: string;
  file?: (success: (file: File) => void, failure?: (e: unknown) => void) => void;
  createReader?: () => { readEntries: (success: (entries: EntryLike[]) => void, failure?: (e: unknown) => void) => void };
}

/** What the synchronous part of the drop handler collected. */
export interface DropPayload {
  /** plain files (folders removed) */
  files: File[];
  /** folder entries (valid only while the drop event is being handled) */
  folders: EntryLike[];
}

export interface FolderScan {
  files: File[];
  /** Excel files found (before the limit) */
  found: number;
  /** files inside the folders that are not Excel files (ignored) */
  ignored: number;
  /** folders that could not be read */
  unreadable: number;
  /** more Excel files than `maxFiles`: the rest was left out */
  truncated: boolean;
  /** folders nested deeper than `maxDepth` were not opened */
  tooDeep: boolean;
  /** the folder holds more entries than `maxEntries` (e.g. a whole Documents folder): the search stopped early */
  stopped: boolean;
}

export const MAX_FOLDER_FILES = 30;
export const MAX_FOLDER_DEPTH = 4;
/** entries (files + folders) looked at in total before the search gives up */
export const MAX_FOLDER_ENTRIES = 3000;

interface ItemLike {
  kind: string;
  getAsFile(): File | null;
  webkitGetAsEntry?: () => EntryLike | null;
}

/**
 * Synchronous part of the drop: separates the dropped folders from the plain files.
 * Without `items` (old browsers / tests) the plain `files` list is used as it is.
 */
export function readDropPayload(dt: { files?: ArrayLike<File> | null; items?: ArrayLike<ItemLike> | null } | null | undefined): DropPayload {
  const files: File[] = [];
  const folders: EntryLike[] = [];
  if (!dt) return { files, folders };
  const items = dt.items ? Array.from(dt.items) : [];
  if (items.length) {
    for (const it of items) {
      if (it.kind !== 'file') continue;
      let entry: EntryLike | null = null;
      try {
        entry = it.webkitGetAsEntry ? it.webkitGetAsEntry() : null;
      } catch {
        entry = null;
      }
      if (entry?.isDirectory) {
        folders.push(entry);
        continue;
      }
      const f = it.getAsFile();
      if (f) files.push(f);
    }
    // some browsers expose fewer items than files
    if (files.length + folders.length === 0 && dt.files) files.push(...Array.from(dt.files));
    return { files, folders };
  }
  return { files: dt.files ? Array.from(dt.files) : [], folders };
}

function readAll(reader: ReturnType<NonNullable<EntryLike['createReader']>>): Promise<EntryLike[]> {
  // readEntries returns the entries in batches (Chromium: 100 at a time) until an empty batch
  return new Promise((resolve, reject) => {
    const all: EntryLike[] = [];
    const next = () =>
      reader.readEntries((batch) => {
        if (!batch.length) resolve(all);
        else {
          all.push(...batch);
          next();
        }
      }, reject);
    next();
  });
}

const fileOf = (entry: EntryLike): Promise<File> =>
  new Promise((resolve, reject) => {
    if (!entry.file) reject(new Error('not a file entry'));
    else entry.file(resolve, reject);
  });

/** Natural, case-insensitive order (like the Explorer): 'a2' before 'a10'. */
const byName = (a: EntryLike, b: EntryLike) => a.name.localeCompare(b.name, 'ja', { numeric: true, sensitivity: 'base' });

/** Walks the dropped folders for Excel files (other files are counted and ignored). */
export async function scanFolders(folders: readonly EntryLike[], opts: { maxFiles?: number; maxDepth?: number; maxEntries?: number } = {}): Promise<FolderScan> {
  const maxFiles = opts.maxFiles ?? MAX_FOLDER_FILES;
  const maxDepth = opts.maxDepth ?? MAX_FOLDER_DEPTH;
  const maxEntries = opts.maxEntries ?? MAX_FOLDER_ENTRIES;
  const out: FolderScan = { files: [], found: 0, ignored: 0, unreadable: 0, truncated: false, tooDeep: false, stopped: false };
  let seen = 0;
  const walk = async (dir: EntryLike, depth: number): Promise<void> => {
    if (out.stopped) return;
    if (!dir.createReader) {
      out.unreadable++;
      return;
    }
    let entries: EntryLike[];
    try {
      entries = await readAll(dir.createReader());
    } catch {
      out.unreadable++;
      return;
    }
    entries.sort(byName);
    for (const e of entries) {
      if (++seen > maxEntries) {
        out.stopped = true;
        return;
      }
      if (e.isDirectory) {
        if (e.name.startsWith('.')) continue;
        if (depth >= maxDepth) out.tooDeep = true;
        else await walk(e, depth + 1);
        continue;
      }
      if (!e.isFile) continue;
      if (isExcelLockFile(e.name) || e.name.startsWith('.') || /^(thumbs\.db|desktop\.ini)$/i.test(e.name)) continue;
      if (!isAcceptedFileName(e.name)) {
        out.ignored++;
        continue;
      }
      out.found++;
      if (out.files.length >= maxFiles) {
        out.truncated = true;
        continue;
      }
      try {
        out.files.push(await fileOf(e));
      } catch {
        out.unreadable++;
      }
    }
  };
  for (const f of folders) await walk(f, 0);
  return out;
}

/** Japanese one-line summary of a folder scan for a toast, or null when there is nothing to say. */
export function folderScanMessage(folderNames: readonly string[], scan: FolderScan): { text: string; kind: 'info' | 'error' } | null {
  const what = folderNames.length === 1 ? `フォルダ「${folderNames[0]}」` : `${folderNames.length} 個のフォルダ`;
  if (scan.files.length === 0) {
    if (scan.unreadable > 0 && scan.found === 0) return { text: `${what}を読み込めませんでした。フォルダの中のExcelファイルを選んでドロップしてください`, kind: 'error' };
    return { text: `${what}にExcelファイル（.xlsx / .xlsm / .xls / .xlsb / .csv / .ods）が見つかりませんでした${scan.stopped ? '（中身が非常に多いため、途中までしか調べていません）' : ''}`, kind: 'error' };
  }
  const parts = [`${what}から ${scan.files.length} 個のファイルを読み込みます`];
  if (scan.truncated) parts.push(`多すぎるため先頭の ${scan.files.length} 個だけです（全部で ${scan.found} 個）`);
  if (scan.tooDeep) parts.push('深い階層のフォルダは開いていません');
  if (scan.stopped) parts.push('中身が非常に多いため、途中までしか調べていません');
  if (scan.unreadable > 0) parts.push(`${scan.unreadable} 件は読み込めませんでした`);
  if (scan.ignored > 0) parts.push(`Excel以外の ${scan.ignored} 個は無視しました`);
  return { text: parts.join('。'), kind: scan.truncated || scan.unreadable > 0 || scan.stopped ? 'error' : 'info' };
}
