/**
 * Sheet names for sheets copied from the other files into the merged workbook (plan A, round 9).
 *
 * Rule (user's request): the name of a copied sheet is "<name of the original Excel file without extension>_<original sheet name>".
 * When that name is already taken in the workbook, the sheet that is ADDED LATER gets a suffix "(1)", "(2)", ... (the earlier one keeps
 * its name; the merged sheet and the base file's own sheets are never renamed).
 *
 * Excel's limits are respected: at most 31 characters (UTF-16 code units, no surrogate pair is cut), none of \ / ? * [ ] : (replaced
 * by "_"), no leading / trailing apostrophe, not empty, compared case-insensitively (and ignoring full-width / half-width forms, which is
 * stricter than Excel on purpose). Too long names shorten the FILE part first (keeping at least 6 characters of it), then the sheet part.
 *
 * Pure functions: no DOM, no Node APIs.
 */

export const SHEET_NAME_MAX = 31;

const EXT = /\.(xlsx|xlsm|xltx|xltm|xlsb|xls|ods|csv|tsv|txt)$/i;
const INVALID = /[\\/?*[\]:]/g;

/** the file name without folder and known extension ("C:\x\02_面接評価.xlsx" -> "02_面接評価") */
export function fileBaseName(fileName: string): string {
  const leaf = String(fileName ?? '').split(/[\\/]/).pop() ?? '';
  return leaf.replace(EXT, '');
}

/** key used to compare names (Excel ignores case; full / half width folded as well) */
export function sheetNameKey(name: string): string {
  return name.normalize('NFKC').toLowerCase();
}

/** cut to at most `units` UTF-16 code units without splitting a surrogate pair */
export function cutUnits(s: string, units: number): string {
  if (units <= 0) return '';
  if (s.length <= units) return s;
  let end = units;
  const c = s.charCodeAt(end - 1);
  if (c >= 0xd800 && c <= 0xdbff) end--; // high surrogate left alone at the end
  return s.slice(0, end);
}

/** replace what Excel does not allow in a sheet name; never returns '' */
export function sanitizeSheetName(raw: string, fallback = 'Sheet'): string {
  let s = String(raw ?? '').replace(INVALID, '_').replace(/[\u0000-\u001f]/g, '');
  s = s.replace(/^'+/, '').replace(/'+$/, '');
  return s.length ? s : fallback;
}

/** "<file>_<sheet>" shortened to the Excel limit (file part first, at least 6 characters of it are kept when it has that many) */
export function joinSheetName(file: string, sheet: string, max = SHEET_NAME_MAX): string {
  const f = sanitizeSheetName(file, 'Book');
  const s = sanitizeSheetName(sheet, 'Sheet');
  if (f.length + 1 + s.length <= max) return f + '_' + s;
  const minFile = Math.min(f.length, 6);
  const sheetKeep = Math.min(s.length, max - 1 - minFile);
  const fileKeep = Math.min(f.length, max - 1 - sheetKeep);
  return sanitizeSheetName(cutUnits(f, fileKeep) + '_' + cutUnits(s, sheetKeep));
}

/** `name` + "(n)" (n >= 1) within the Excel limit */
export function withSuffix(name: string, n: number, max = SHEET_NAME_MAX): string {
  const suffix = `(${n})`;
  return cutUnits(name, max - suffix.length) + suffix;
}

export interface SheetNameRequest {
  /** the original file name (any folder, with or without extension) */
  file: string;
  /** the sheets of that file in workbook order */
  sheets: string[];
}

export interface PlannedSheetName {
  file: string;
  original: string;
  name: string;
  /** true when "(n)" had to be appended because the plain name was already taken */
  suffixed: boolean;
}

/**
 * Names for every sheet of every request, in request order (= the order the files are merged in, the order of the cards).
 * `existing`: the sheet names that are in the workbook already (the merged sheet + the base file's other sheets that are kept).
 */
export function planSheetNames(existing: readonly string[], requests: readonly SheetNameRequest[]): PlannedSheetName[] {
  const taken = new Set(existing.map(sheetNameKey));
  const out: PlannedSheetName[] = [];
  for (const rq of requests) {
    const base = fileBaseName(rq.file);
    for (const original of rq.sheets) {
      const plain = joinSheetName(base, original);
      let name = plain;
      let n = 0;
      while (taken.has(sheetNameKey(name))) {
        n++;
        name = withSuffix(plain, n);
      }
      taken.add(sheetNameKey(name));
      out.push({ file: rq.file, original, name, suffixed: n > 0 });
    }
  }
  return out;
}
