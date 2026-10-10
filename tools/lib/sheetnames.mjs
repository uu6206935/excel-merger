// Sheet names of the sheets copied from the other files (plan A): a plain-JavaScript port of src/core/xlsx/sheetnames.ts (planSheetNames), so that
// the QA tools (node tools/check-output.mjs, no TypeScript loader) can compute what the copies must be called. It is a COPY of a contract file:
// tests/qa/sheetnames-port.test.ts compares it with the TypeScript original on hundreds of random inputs — change both or neither.
//
// Rule: "<file name without extension>_<sheet name>", at most 31 characters (the file part is shortened first, at least 6 characters of it are
// kept), none of \ / ? * [ ] : (replaced by "_"), "(1)", "(2)" … appended to the sheet that is added LATER when a name is taken
// (case-insensitive, full / half width folded).

export const SHEET_NAME_MAX = 31;

const EXT = /\.(xlsx|xlsm|xltx|xltm|xlsb|xls|ods|csv|tsv|txt)$/i;
const INVALID = /[\\/?*[\]:]/g;

/** the file name without folder and known extension ("C:\x\02_面接評価.xlsx" -> "02_面接評価") */
export function fileBaseName(fileName) {
  const leaf = String(fileName ?? '').split(/[\\/]/).pop() ?? '';
  return leaf.replace(EXT, '');
}

/** key used to compare names (Excel ignores case; full / half width folded as well) */
export function sheetNameKey(name) {
  return name.normalize('NFKC').toLowerCase();
}

/** cut to at most `units` UTF-16 code units without splitting a surrogate pair */
export function cutUnits(s, units) {
  if (units <= 0) return '';
  if (s.length <= units) return s;
  let end = units;
  const c = s.charCodeAt(end - 1);
  if (c >= 0xd800 && c <= 0xdbff) end--;
  return s.slice(0, end);
}

/** replace what Excel does not allow in a sheet name; never returns '' */
export function sanitizeSheetName(raw, fallback = 'Sheet') {
  let s = String(raw ?? '').replace(INVALID, '_').replace(/[\u0000-\u001f]/g, '');
  s = s.replace(/^'+/, '').replace(/'+$/, '');
  return s.length ? s : fallback;
}

/** "<file>_<sheet>" shortened to the Excel limit */
export function joinSheetName(file, sheet, max = SHEET_NAME_MAX) {
  const f = sanitizeSheetName(file, 'Book');
  const s = sanitizeSheetName(sheet, 'Sheet');
  if (f.length + 1 + s.length <= max) return f + '_' + s;
  const minFile = Math.min(f.length, 6);
  const sheetKeep = Math.min(s.length, max - 1 - minFile);
  const fileKeep = Math.min(f.length, max - 1 - sheetKeep);
  return sanitizeSheetName(cutUnits(f, fileKeep) + '_' + cutUnits(s, sheetKeep));
}

/** `name` + "(n)" (n >= 1) within the Excel limit */
export function withSuffix(name, n, max = SHEET_NAME_MAX) {
  const suffix = `(${n})`;
  return cutUnits(name, max - suffix.length) + suffix;
}

/**
 * Names for every sheet of every request, in request order (= the order of the cards).
 * @param {readonly string[]} existing sheet names already in the workbook (the merged sheet + the base's other sheets that are kept)
 * @param {ReadonlyArray<{file: string, sheets: string[]}>} requests
 * @returns {Array<{file: string, original: string, name: string, suffixed: boolean}>}
 */
export function planSheetNames(existing, requests) {
  const taken = new Set(existing.map(sheetNameKey));
  const out = [];
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
