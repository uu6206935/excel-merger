/**
 * readBook (owner: reader R): bytes → SourceBook. See ./index.ts for the contract.
 *
 * .xlsx / .xlsm / .xltx / .xltm (Transitional and Strict OOXML): own reader — JSZip + string scanning, no DOM.
 * .xls / .ods / .xlsb / .csv / .tsv / .txt / HTML-or-XML "xls": SheetJS via ./legacy (fidelity 'values', bytes
 * replaced by a generated minimal .xlsx package; format 'xls' / 'ods' / 'xlsb' / 'csv' / 'other'; text inputs also
 * report `encoding`: 'Shift_JIS', 'UTF-8', 'UTF-16LE', 'UTF-16BE', 'EUC-JP', 'ISO-2022-JP').
 * Errors are ReadBookError with Japanese user-facing messages (MSG_* / msg*() in ./zip). Guards: zip bombs (declared
 * total > 1.5 GB → 'tooLarge' before inflating; no part inflates beyond its declared size), parts / text inputs
 * beyond V8's string limit (sharedStrings → 'tooLarge' error, a sheet → empty sheet + warning, text → 'tooLarge').
 */
import type { BookFeatureCounts, DefinedName, SheetGrid, SourceBook, SourceFormat, TableInfo } from '../types';
import type { ProgressFn } from '../contracts';
import { defaultColumnPx, mdwFor, pxToExcelWidth } from '../display/styles';
import { buildDisplayStyles, paletteOf, parseStyleSheet } from './styles';
import { parseThemeColors, type Palette } from './theme';
import { EMPTY_SST, parseSharedStrings, rootPrefix, type SharedStrings } from './strings';
import { parseSheetRest, scanSheetData, type ScannedSheetData } from './sheet';
import {
  MAX_PART_BYTES,
  MSG_CORRUPT,
  MSG_EMPTY,
  MSG_PASSWORD,
  Package,
  ReadBookError,
  contentTypeOf,
  isEncryptedCfb,
  loadZip,
  msgSharedStringsTooLarge,
  msgTextTooLarge,
  readContentTypes,
  readRels,
  sizeText,
  sniff,
  type Rel,
} from './zip';
import { boolAttr, decodeOoxmlEscapes, innerText, kid, kids, numAttr, parseXml, relId, rootOf, type XEl } from './xml';
import { convertLegacy } from './legacy';

export { ReadBookError, MSG_CORRUPT, MSG_EMPTY, MSG_PASSWORD };

let idCounter = 0;
function makeId(): string {
  idCounter++;
  const rnd = Math.random().toString(36).slice(2, 8);
  return `book-${Date.now().toString(36)}-${idCounter}-${rnd}`;
}

function extOf(fileName: string): string {
  const i = fileName.lastIndexOf('.');
  return i < 0 ? '' : fileName.slice(i + 1).toLowerCase();
}

/** BIFF2..BIFF5 BOF record at offset 0 (record type 0x0009 / 0x0209 / 0x0409 / 0x0809, little endian). */
function isRawBiff(bytes: Uint8Array): boolean {
  return bytes.length >= 8 && bytes[0] === 0x09 && (bytes[1] === 0x00 || bytes[1] === 0x02 || bytes[1] === 0x04 || bytes[1] === 0x08);
}

/** Yield to the event loop so progress callbacks can paint (no-op cost in Node). */
function tick(): Promise<void> {
  return new Promise((r) => setTimeout(r, 0));
}

export async function readBook(bytes: Uint8Array, fileName: string, opts?: { id?: string; onProgress?: ProgressFn }): Promise<SourceBook> {
  const id = opts?.id ?? makeId();
  const progress = opts?.onProgress;
  const ext = extOf(fileName);
  const kind = sniff(bytes);
  if (kind === 'empty') throw new ReadBookError('empty', MSG_EMPTY);
  if (kind === 'cfb') {
    if (isEncryptedCfb(bytes)) throw new ReadBookError('password', MSG_PASSWORD);
    return readLegacy(bytes, fileName, { id, format: 'xls', label: 'xls', onProgress: progress });
  }
  if (kind === 'text') {
    // CSV / TSV / TXT, or markup (HTML table / Excel 2003 XML) that business systems save as "*.xls";
    // a .csv / .tsv / .txt name always means delimited text. Text beyond V8's string limit cannot be decoded.
    if (bytes.length > MAX_PART_BYTES) throw new ReadBookError('tooLarge', msgTextTooLarge(bytes.length));
    return readLegacy(bytes, fileName, { id, format: 'csv', text: true, delimited: ext === 'csv' || ext === 'tsv' || ext === 'txt', onProgress: progress });
  }
  if (kind !== 'zip') {
    // unknown binary: only raw BIFF2-5 streams (Excel 2.x-5.0 .xls without a CFB container) go to SheetJS;
    // anything else would be "parsed" as garbage text, so it is rejected.
    if (isRawBiff(bytes)) return readLegacy(bytes, fileName, { id, format: 'xls', label: 'xls', onProgress: progress });
    throw new ReadBookError('corrupt', MSG_CORRUPT);
  }
  return readPackage(bytes, fileName, id, progress, true);
}

/** Legacy inputs: SheetJS → generated values-only .xlsx → normal package reader. */
async function readLegacy(
  bytes: Uint8Array,
  fileName: string,
  o: { id: string; format: SourceFormat; label?: string; text?: boolean; delimited?: boolean; onProgress?: ProgressFn },
): Promise<SourceBook> {
  o.onProgress?.({ ratio: 0, message: 'ファイルを変換中…' });
  const conv = await convertLegacy(bytes, fileName, { text: o.text, delimited: o.delimited });
  const book = await readPackage(conv.bytes, fileName, o.id, o.onProgress, false);
  let format = o.format;
  let label = o.label ?? 'この';
  if (conv.kind === 'delimited') [format, label] = ['csv', extOf(fileName) === 'tsv' ? 'TSV' : 'CSV'];
  else if (conv.kind === 'html') [format, label] = [extOf(fileName) === 'xls' ? 'xls' : 'other', 'HTML'];
  else if (conv.kind === 'xml2003' || conv.kind === 'xml') [format, label] = [extOf(fileName) === 'xls' ? 'xls' : 'other', 'XMLスプレッドシート'];
  const enc = conv.encoding ? ENCODING_LABEL[conv.encoding] ?? conv.encoding : undefined;
  const first =
    conv.kind === 'delimited'
      ? `${label}ファイルを読み込みました（文字コード: ${enc ?? '不明'}）。先頭が0のコード（00123など）や日付は文字のまま扱います。`
      : `${label}形式のファイルは値（と表示形式）だけを読み込みました。数式・色・罫線などは引き継がれません。`;
  const out: SourceBook = {
    ...book,
    size: bytes.length,
    format,
    fidelity: 'values',
    warnings: [first, ...conv.warnings, ...book.warnings],
  };
  // every text input (CSV / TSV / TXT, HTML / XML "xls") was decoded: report the encoding by its standard name
  if (conv.encoding) out.encoding = ENCODING_NAME[conv.encoding] ?? conv.encoding;
  return out;
}

/** Labels used in the warning text (kept as they were: UTF-16 without byte order). */
const ENCODING_LABEL: Record<string, string> = { SJIS: 'Shift_JIS', UTF8: 'UTF-8', UTF16LE: 'UTF-16', UTF16BE: 'UTF-16', EUCJP: 'EUC-JP', JIS: 'JIS' };

/**
 * SourceBook.encoding: WHATWG / IANA names of the detected encodings (decodeText / encoding-japanese codes →
 * names). Shift_JIS here means Windows CP932 (what Excel writes and what the decoder uses).
 */
const ENCODING_NAME: Record<string, string> = {
  SJIS: 'Shift_JIS',
  UTF8: 'UTF-8',
  UTF16LE: 'UTF-16LE',
  UTF16BE: 'UTF-16BE',
  EUCJP: 'EUC-JP',
  JIS: 'ISO-2022-JP',
};

async function readPackage(bytes: Uint8Array, fileName: string, id: string, progress: ProgressFn | undefined, allowLegacy: boolean): Promise<SourceBook> {
  const ext = extOf(fileName);
  progress?.({ ratio: 0, message: 'ファイルを展開中…' });
  const zip = await loadZip(bytes);
  const pkg = new Package(zip);

  // locate the workbook part through the package relationships
  const rootRels = await readRels(pkg, '');
  const officeRel = rootRels.find((r) => r.kind === 'officeDocument' && !r.external);
  let wbPath = officeRel?.path && pkg.has(officeRel.path) ? pkg.resolveName(officeRel.path)! : pkg.resolveName('xl/workbook.xml');
  if (!wbPath || wbPath.toLowerCase().endsWith('.bin')) {
    if (allowLegacy) {
      if (pkg.has('xl/workbook.bin') || officeRel?.path?.toLowerCase().endsWith('.bin')) return readLegacy(bytes, fileName, { id, format: 'xlsb', label: 'xlsb', onProgress: progress });
      if (pkg.has('content.xml') && (pkg.has('mimetype') || ext === 'ods')) return readLegacy(bytes, fileName, { id, format: 'ods', label: 'ODS', onProgress: progress });
    }
    throw new ReadBookError('corrupt', MSG_CORRUPT);
  }

  const ct = await readContentTypes(pkg);
  const wbType = contentTypeOf(ct, wbPath);
  let format: SourceFormat = ext === 'xlsm' || ext === 'xltm' || /macroEnabled/i.test(wbType) ? 'xlsm' : 'xlsx';

  const wbXml = await pkg.text(wbPath);
  if (!wbXml) throw new ReadBookError('corrupt', MSG_CORRUPT);
  const wbRoot = rootOf(parseXml(wbXml));
  if (!wbRoot || wbRoot.local !== 'workbook') throw new ReadBookError('corrupt', MSG_CORRUPT);
  const warnings: string[] = [];
  const wbRels = await readRels(pkg, wbPath);
  const relById = new Map(wbRels.map((r) => [r.id, r] as const));
  const date1904 = boolAttr(kid(wbRoot, 'workbookPr')?.attrs.date1904);

  // theme, styles, shared strings
  const themeRel = wbRels.find((r) => r.kind === 'theme' && !r.external);
  const themeColors = parseThemeColors(themeRel ? await pkg.text(themeRel.path) : null);
  const stylesRel = wbRels.find((r) => r.kind === 'styles' && !r.external);
  const styleSheet = parseStyleSheet(stylesRel ? await pkg.text(stylesRel.path) : await pkg.text('xl/styles.xml'));
  const pal = paletteOf(styleSheet, themeColors);
  const { styles, isDate, dateMode } = buildDisplayStyles(styleSheet, pal);
  const sst = await readSharedStrings(pkg, wbRels, pal);

  // sheets
  const sheetEls = kids(kid(wbRoot, 'sheets'), 'sheet');
  const sheets: SheetGrid[] = [];
  const counts = { dataTable: 0, sharedMissing: 0, badStyle: 0, badValue: 0, missing: 0, noResult: 0 };
  const tooLarge: string[] = [];
  let sparklineGroups = 0;
  let macroSheets = 0;
  let chartSheets = 0;
  const usedNames = new Set<string>();
  for (let si = 0; si < sheetEls.length; si++) {
    const el = sheetEls[si];
    let name = el.attrs.name ?? `Sheet${si + 1}`;
    if (usedNames.has(name.toLowerCase())) name = `${name} (${si + 1})`;
    usedNames.add(name.toLowerCase());
    const stateAttr = el.attrs.state;
    const state: SheetGrid['state'] = stateAttr === 'hidden' ? 'hidden' : stateAttr === 'veryHidden' ? 'veryHidden' : 'visible';
    progress?.({ ratio: 0.05 + (0.9 * si) / Math.max(1, sheetEls.length), message: `シートを読み込み中…（${si + 1}/${sheetEls.length}）${name}` });
    const rid = relId(el);
    const rel = rid ? relById.get(rid) : undefined;
    const grid = emptyGrid(name, state);
    if (!rel || rel.external || !pkg.has(rel.path)) {
      counts.missing++;
      sheets.push(grid);
      continue;
    }
    if (rel.kind === 'chartsheet') {
      chartSheets++;
      sheets.push(grid);
      continue;
    }
    if (rel.kind !== 'worksheet') {
      // macrosheet (Excel 4.0), dialogsheet …
      macroSheets++;
      sheets.push(grid);
      continue;
    }
    const part = await pkg.load(rel.path);
    pkg.forget(rel.path);
    if (!part.ok) {
      // a sheet part beyond the string limit cannot be read at all: say so by name (not just "missing")
      if (part.problem === 'tooLarge') tooLarge.push(`シート「${name}」は大きすぎるため読み込めませんでした（展開後 ${sizeText(pkg.size(rel.path) ?? Infinity)}）。空のシートとして扱います。`);
      else counts.missing++;
      sheets.push(grid);
      continue;
    }
    const xml = part.text;
    const sheetRels = await readRels(pkg, rel.path);
    const r = readWorksheet(xml, sheetRels, { sst, isDate, dateMode, date1904, styleCount: styles.length, pal });
    counts.dataTable += r.scan.counts.dataTable;
    counts.sharedMissing += r.scan.counts.sharedMissing;
    counts.badStyle += r.scan.counts.badStyle;
    counts.badValue += r.scan.counts.badValue;
    counts.noResult += r.scan.counts.noResult;
    sparklineGroups += r.sparklineGroups;
    grid.rows = r.scan.rows;
    grid.rowProps = r.scan.rowProps;
    grid.maxRow = r.scan.maxRow;
    grid.maxCol = r.scan.maxCol;
    grid.colProps = r.colProps;
    grid.features = { ...r.features, tables: [], arrays: r.scan.arrays };
    const dcw = sheetDefaultColWidth(r, styles[0]);
    if (dcw !== undefined) grid.defaultColWidth = dcw;
    if (r.defaultRowHeight !== undefined) grid.defaultRowHeight = r.defaultRowHeight;
    for (const tp of r.tableParts) {
      const t = await readTable(pkg, tp);
      if (t) grid.features.tables.push(t);
    }
    sheets.push(grid);
    if (si % 2 === 1) await tick();
  }

  // defined names
  const names: DefinedName[] = [];
  for (const dn of kids(kid(wbRoot, 'definedNames'), 'definedName')) {
    const nm = dn.attrs.name;
    if (!nm) continue;
    const lsi = numAttr(dn.attrs.localSheetId);
    const scope = lsi !== undefined && sheets[lsi] ? sheets[lsi].name : null;
    const n: DefinedName = { name: nm, scope, refersTo: decodeOoxmlEscapes(innerText(dn)).trim() };
    if (boolAttr(dn.attrs.hidden)) n.hidden = true;
    names.push(n);
  }

  const features = await countFeatures(pkg, ct, sparklineGroups);
  if (features.vba && format === 'xlsx') format = 'xlsm';

  if (counts.dataTable) warnings.push(`データテーブル（What-If分析）の数式 ${counts.dataTable} 個は値として読み込みました。`);
  if (counts.sharedMissing) warnings.push(`共有数式の元の数式が見つからないセルが ${counts.sharedMissing} 個あります（値のみ読み込みました）。`);
  if (counts.badStyle) warnings.push(`書式の番号が不正なセルが ${counts.badStyle} 個あります（標準の書式で表示します）。`);
  if (counts.badValue) warnings.push(`読み取れない値のセルが ${counts.badValue} 個あります。`);
  if (counts.noResult) warnings.push(`計算結果が保存されていない数式が ${counts.noResult} 個あります（プログラムで作られたファイルに多い形式です）。結合すると値は空欄になることがあります。Excel で開いて保存し直すと値が入ります。`);
  if (counts.missing) warnings.push(`読み込めなかったシートが ${counts.missing} 個あります（空のシートとして扱います）。`);
  warnings.push(...tooLarge);
  if (chartSheets) warnings.push(`グラフシート ${chartSheets} 枚はデータを持たないため空のシートとして表示します。`);
  if (macroSheets) warnings.push(`マクロシート／ダイアログシート ${macroSheets} 枚は内容を表示できません。`);
  if (!sheets.length) throw new ReadBookError('corrupt', MSG_CORRUPT);

  progress?.({ ratio: 1, message: '読み込み完了' });
  return {
    id,
    fileName,
    size: bytes.length,
    format,
    bytes,
    fidelity: 'full',
    styles,
    sheets,
    names,
    features,
    date1904,
    warnings,
  };
}

/**
 * The shared-strings table. Size check first: a part too large to decode (MAX_PART_BYTES, V8's string limit) is
 * refused with a Japanese message BEFORE it is inflated; an existing part that cannot be read (broken deflate
 * data, a part that inflates beyond its declared size) is a corrupt file. Both used to read as "no shared
 * strings", i.e. every text cell silently empty — wrong keys and data with only a vague warning.
 */
async function readSharedStrings(pkg: Package, wbRels: Rel[], pal: Palette): Promise<SharedStrings> {
  const sstRel = wbRels.find((r) => r.kind === 'sharedStrings' && !r.external);
  const sstPath = sstRel?.path ?? 'xl/sharedStrings.xml';
  if (!pkg.has(sstPath)) return EMPTY_SST;
  const declared = pkg.size(sstPath);
  if (declared !== undefined && (declared < 0 || declared > MAX_PART_BYTES)) throw new ReadBookError('tooLarge', msgSharedStringsTooLarge(declared < 0 ? Infinity : declared));
  const part = await pkg.load(sstPath);
  pkg.forget(sstPath);
  if (!part.ok) {
    if (part.problem === 'tooLarge') throw new ReadBookError('tooLarge', msgSharedStringsTooLarge(declared ?? Infinity));
    throw new ReadBookError('corrupt', MSG_CORRUPT);
  }
  return parseSharedStrings(part.text, pal);
}

/**
 * SheetGrid.defaultColWidth (stored <col width> unit) of a sheet from its <sheetFormatPr>, or undefined = the standard
 * 8 characters of the default font.
 *
 *  - defaultColWidth > 0: as stored. A non-positive value is ignored, like the patcher does (it takes defaultColWidth
 *    only when > 0 and otherwise derives the width from baseColWidth), so the doc and the patcher stay on one number.
 *  - else baseColWidth other than 8 (Mac Excel writes baseColWidth="10" and no defaultColWidth, in every file): the
 *    default column is baseColWidth digits of the default font wide + 5 px padding, rounded up to a multiple of 8 px
 *    (= `defaultColumnPx(mdw, baseColWidth)`, the formula the patcher's auto-widening uses). Measured in Excel (own hidden
 *    instance, Range.Width of a column without <col>, 2026-10): base 1, 5, 9, 10, 11, 12, 15, 20, 100 for 游ゴシック 11,
 *    Calibri 11, ＭＳ Ｐゴシック 11, メイリオ 9; base 10 for Calibri 12 (the Mac default), Aptos Narrow 11 / 12, 游ゴシック
 *    10, ＭＳ Ｐゴシック 10, Calibri 9, 游明朝 11: the same pixel width everywhere, e.g. Calibri 12 / base 10 = 88 px (10.38
 *    chars; Excel saves defaultColWidth="11"). Excel stores width = px / mdw truncated to 1/256; this one is derived through
 *    pxToExcelWidth and can sit one 1/256 step lower — the same pixels (excelWidthToPx gives the same number). Fonts whose
 *    digit width mdwFor() gets wrong (Arial, Times New Roman, Calibri 10, メイリオ 11 …) are off by the same amount with and
 *    without baseColWidth: that is the font table of display/styles, not this rule.
 *  - baseColWidth absent / 8 / not positive: undefined (8 is the schema default; Excel treats 0 as 8 as well).
 *  - a derived width above 255 characters (Excel itself ignores baseColWidth 255 and 65536 and falls back to the standard
 *    width; it refuses to open the file for 256, 300 and 1000; base 254 = exactly 255 characters is the last one that
 *    counts, by the same rule — not measured) is not used, so a hostile value cannot make a 500 000 px column.
 *
 * `defaultFont` = cellXfs[0]'s font (what the patcher measures with: fontOf(0)).
 */
export function sheetDefaultColWidth(fmt: { defaultColWidth?: number; baseColWidth?: number }, defaultFont?: { font?: string; size?: number }): number | undefined {
  if (fmt.defaultColWidth !== undefined && fmt.defaultColWidth > 0) return fmt.defaultColWidth;
  const base = fmt.baseColWidth;
  if (base === undefined || !(base > 0) || base === 8) return undefined;
  const mdw = mdwFor(defaultFont?.font, defaultFont?.size && defaultFont.size > 0 ? defaultFont.size : 11);
  const px = defaultColumnPx(mdw, base);
  return px <= 255 * mdw ? pxToExcelWidth(px, mdw) : undefined; // 255 characters: the widest column Excel has
}

export function emptyGrid(name: string, state: SheetGrid['state'] = 'visible'): SheetGrid {
  return {
    name,
    state,
    maxRow: 0,
    maxCol: 0,
    rows: [],
    colProps: [],
    rowProps: [],
    features: { merges: [], rules: [], tables: [], arrays: [], hyperlinks: [] },
  };
}

interface WorksheetRead {
  scan: ScannedSheetData;
  features: ReturnType<typeof parseSheetRest>['features'];
  colProps: SheetGrid['colProps'];
  defaultColWidth?: number;
  defaultRowHeight?: number;
  baseColWidth?: number;
  tableParts: string[];
  sparklineGroups: number;
}

/** Read one worksheet part (exported for tests). */
export function readWorksheet(xml: string, rels: Rel[], ctx: Parameters<typeof scanSheetData>[4]): WorksheetRead {
  const p = rootPrefix(xml);
  const openTag = `<${p}sheetData`;
  let sdStart = -1;
  let search = 0;
  for (;;) {
    const i = xml.indexOf(openTag, search);
    if (i < 0) break;
    const ch = xml.charCodeAt(i + openTag.length);
    if (ch === 62 || ch === 32 || ch === 47 || ch === 9 || ch === 10 || ch === 13) {
      sdStart = i;
      break;
    }
    search = i + openTag.length;
  }
  let rest = xml;
  let scan: ScannedSheetData;
  if (sdStart >= 0) {
    const gt = xml.indexOf('>', sdStart);
    const selfClose = gt > 0 && xml.charCodeAt(gt - 1) === 47;
    if (selfClose) {
      scan = scanSheetData(xml, gt + 1, gt + 1, p, ctx);
      rest = xml.slice(0, sdStart) + xml.slice(gt + 1);
    } else {
      const closeTag = `</${p}sheetData>`;
      // the end tag sits near the end of the part: search backwards (a forward search scans the whole sheetData)
      let sdEnd = xml.lastIndexOf(closeTag);
      if (sdEnd < gt) sdEnd = -1;
      const contentEnd = sdEnd < 0 ? xml.length : sdEnd;
      sdEnd = sdEnd < 0 ? xml.length : sdEnd + closeTag.length;
      scan = scanSheetData(xml, gt + 1, contentEnd, p, ctx);
      rest = xml.slice(0, sdStart) + xml.slice(sdEnd);
    }
  } else {
    scan = scanSheetData('', 0, 0, p, ctx);
  }
  let root: XEl | undefined;
  try {
    root = rootOf(parseXml(rest));
  } catch {
    root = undefined;
  }
  const r = parseSheetRest(root, rels);
  return {
    scan,
    features: r.features,
    colProps: r.colProps,
    defaultColWidth: r.defaultColWidth,
    defaultRowHeight: r.defaultRowHeight,
    baseColWidth: r.baseColWidth,
    tableParts: r.tableParts,
    sparklineGroups: r.sparklineGroups,
  };
}

/** xl/tables/tableN.xml → TableInfo (headerRowCount absent = 1; totalsRowCount absent = 0). */
async function readTable(pkg: Package, path: string): Promise<TableInfo | null> {
  const xml = await pkg.text(path);
  if (!xml) return null;
  return parseTableXml(xml);
}

export function parseTableXml(xml: string): TableInfo | null {
  const t = rootOf(parseXml(xml));
  if (!t || t.local !== 'table') return null;
  const ref = t.attrs.ref;
  if (!ref) return null;
  const hdr = numAttr(t.attrs.headerRowCount);
  const tot = numAttr(t.attrs.totalsRowCount);
  return {
    name: t.attrs.displayName || t.attrs.name || '',
    ref,
    headerRow: hdr === undefined ? true : hdr > 0,
    totalsRow: (tot ?? 0) > 0,
    columns: kids(kid(t, 'tableColumns'), 'tableColumn').map((c) => decodeOoxmlEscapes(c.attrs.name ?? '')),
  };
}

const IMAGE_EXT = /\.(png|jpe?g|gif|bmp|emf|wmf|tiff?|svg|ico|webp)$/i;

async function countFeatures(pkg: Package, ct: Awaited<ReturnType<typeof readContentTypes>>, sparklineGroups: number): Promise<BookFeatureCounts> {
  const f: BookFeatureCounts = {
    charts: 0,
    pivotTables: 0,
    vba: false,
    drawings: 0,
    slicers: 0,
    sparklineGroups,
    externalLinks: 0,
    threadedComments: 0,
    formControls: 0,
    tables: 0,
    images: 0,
  };
  for (const name of pkg.partNames) {
    const lower = name.toLowerCase();
    const type = contentTypeOf(ct, name);
    if (/\/charts\/chart(ex)?\d*\.xml$/.test(lower) || /drawingml\.chart\+xml|chartex\+xml/.test(type)) f.charts++;
    else if (/\/pivottables\/pivottable\d*\.xml$/.test(lower) || /pivotTable\+xml/.test(type)) f.pivotTables++;
    else if (/vbaproject\.bin$/.test(lower)) f.vba = true;
    else if (/\/drawings\/drawing\d*\.xml$/.test(lower) || /drawing\+xml/.test(type)) f.drawings++;
    else if (/\/slicers\/slicer\d*\.xml$/.test(lower) || /slicer\+xml/.test(type)) {
      const x = await pkg.text(name);
      const n = x ? (x.match(/<(?:[A-Za-z0-9]+:)?slicer[\s>]/g) ?? []).length : 0;
      f.slicers += n || 1;
    } else if (/\/externallinks\/externallink\d*\.xml$/.test(lower)) f.externalLinks++;
    else if (/\/threadedcomments\/[^/]+\.xml$/.test(lower)) {
      const x = await pkg.text(name);
      f.threadedComments += x ? (x.match(/<(?:[A-Za-z0-9]+:)?threadedComment[\s>]/g) ?? []).length : 0;
    } else if (/\/ctrlprops\/[^/]+\.xml$/.test(lower)) f.formControls++;
    else if (/\/tables\/table\d*\.xml$/.test(lower) || /table\+xml/.test(type)) f.tables++;
    else if (/(^|\/)media\//.test(lower) && IMAGE_EXT.test(lower)) f.images++;
  }
  return f;
}
