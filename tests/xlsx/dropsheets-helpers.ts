/**
 * Helpers of the single-sheet output tests (PremergeOptions.dropOtherSheets): a flexible package builder for hand-made
 * workbooks, a one-call premerge of a base file, and the package linter that every output has to pass.
 */
import fs from 'node:fs';
import path from 'node:path';
import JSZip from 'jszip';
import type { KeySource, PremergeInput, PremergeResult } from '../../src/core/contracts';
import type { SourceBook, TableRegion } from '../../src/core/types';
import { tokenize, type Prefix } from '../../src/core/formula/excelRef';
import { detectRegion, pickDataSheet } from '../../src/core/keys';
import { buildPremerged, readBook } from '../../src/core/xlsx';
import { STYLES } from '../golden/build';
import { fakeBook } from '../golden/premerge-scenario';

export const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
export const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

export const CT = {
  drawing: 'application/vnd.openxmlformats-officedocument.drawing+xml',
  chart: 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
  table: 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml',
  comments: 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml',
  pivotTable: 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml',
  pivotCacheDef: 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml',
  pivotCacheRec: 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml',
  chartsheet: 'application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml',
  dialogsheet: 'application/vnd.openxmlformats-officedocument.spreadsheetml.dialogsheet+xml',
  slicerCache: 'application/vnd.ms-excel.slicerCache+xml',
  slicer: 'application/vnd.ms-excel.slicer+xml',
  timelineCache: 'application/vnd.ms-excel.timelineCache+xml',
  app: 'application/vnd.openxmlformats-officedocument.extended-properties+xml',
  core: 'application/vnd.openxmlformats-package.core-properties+xml',
};

export interface RelDef {
  id: string;
  /** relationship type: a full URL, or a suffix that is appended to the office relationships namespace */
  type: string;
  target: string;
  external?: boolean;
}

export interface PartDef {
  path: string;
  xml?: string;
  /** a binary part (images) */
  data?: Uint8Array;
  /** content type override (default by extension otherwise: xml = application/xml) */
  ct?: string;
}

export interface SheetDef {
  name: string;
  state?: 'hidden' | 'veryHidden';
  kind?: 'worksheet' | 'chartsheet' | 'dialogsheet';
  /** inner xml of <sheetData> */
  data?: string;
  /** between <sheetFormatPr> and <sheetData> (<cols> ...) */
  head?: string;
  /** after </sheetData> */
  tail?: string;
  rels?: RelDef[];
  /** the whole part, when the sheet is not a plain worksheet */
  xml?: string;
}

export interface PkgDef {
  sheets: SheetDef[];
  sst?: string[];
  /** inner xml of <definedNames> */
  names?: string;
  /** xml after <calcPr> (<pivotCaches>, <extLst> ...) */
  wbExtra?: string;
  bookViews?: string;
  parts?: PartDef[];
  /** relationships of any part (key = the owner part path) */
  partRels?: Record<string, RelDef[]>;
  wbRels?: RelDef[];
  calcChain?: boolean;
  app?: string;
  styles?: string;
}

const relType = (t: string) => (t.includes('://') ? t : `${R}/${t}`);
const relXml = (rels: RelDef[]) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.map((r) => `<Relationship Id="${r.id}" Type="${relType(r.type)}" Target="${r.target}"${r.external ? ' TargetMode="External"' : ''}/>`).join('')}</Relationships>`;

/** hand-made package: worksheets / chart sheets, workbook parts, extra parts with their relationships */
export async function pkgOf(def: PkgDef): Promise<Uint8Array> {
  const zip = new JSZip();
  const overrides: string[] = [];
  const wbRels: RelDef[] = [];
  const names: string[] = [];
  def.sheets.forEach((s, i) => {
    const kind = s.kind ?? 'worksheet';
    const dir = kind === 'worksheet' ? 'worksheets' : kind === 'chartsheet' ? 'chartsheets' : 'dialogsheets';
    const p = `xl/${dir}/sheet${i + 1}.xml`;
    const xml = s.xml ?? `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet ${NS}><dimension ref="A1"/><sheetViews><sheetView${i === 0 ? ' tabSelected="1"' : ''} workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="18.75"/>${s.head ?? ''}<sheetData>${s.data ?? ''}</sheetData>${s.tail ?? ''}</worksheet>`;
    zip.file(p, xml);
    overrides.push(`<Override PartName="/${p}" ContentType="${kind === 'worksheet' ? 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml' : kind === 'chartsheet' ? CT.chartsheet : CT.dialogsheet}"/>`);
    wbRels.push({ id: `rId${i + 1}`, type: kind, target: `${dir}/sheet${i + 1}.xml` });
    if (s.rels?.length) zip.file(`xl/${dir}/_rels/sheet${i + 1}.xml.rels`, relXml(s.rels));
    names.push(`<sheet name="${s.name.replace(/&/g, '&amp;')}" sheetId="${i + 1}"${s.state ? ` state="${s.state}"` : ''} r:id="rId${i + 1}"/>`);
  });
  const n = def.sheets.length;
  wbRels.push({ id: `rId${n + 1}`, type: 'styles', target: 'styles.xml' });
  zip.file('xl/styles.xml', def.styles ?? STYLES);
  overrides.push('<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>');
  wbRels.push({ id: `rId${n + 2}`, type: 'theme', target: 'theme/theme1.xml' });
  zip.file('xl/theme/theme1.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements/></a:theme>');
  overrides.push('<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>');
  if (def.sst) {
    wbRels.push({ id: `rId${n + 3}`, type: 'sharedStrings', target: 'sharedStrings.xml' });
    zip.file('xl/sharedStrings.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<sst ${NS} count="${def.sst.length}" uniqueCount="${def.sst.length}">${def.sst.map((t) => `<si><t>${t.replace(/&/g, '&amp;')}</t></si>`).join('')}</sst>`);
    overrides.push('<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>');
  }
  if (def.calcChain) {
    wbRels.push({ id: `rId${n + 4}`, type: 'calcChain', target: 'calcChain.xml' });
    zip.file('xl/calcChain.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<calcChain ${NS}><c r="A1" i="1"/></calcChain>`);
    overrides.push('<Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>');
  }
  for (const r of def.wbRels ?? []) wbRels.push(r);
  for (const p of def.parts ?? []) {
    zip.file(p.path, p.data ?? p.xml ?? '');
    if (p.ct) overrides.push(`<Override PartName="/${p.path}" ContentType="${p.ct}"/>`);
  }
  for (const [owner, rels] of Object.entries(def.partRels ?? {})) {
    const i = owner.lastIndexOf('/');
    zip.file(`${owner.slice(0, i + 1)}_rels/${owner.slice(i + 1)}.rels`, relXml(rels));
  }
  zip.file('xl/workbook.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook ${NS}>${def.bookViews ?? ''}<sheets>${names.join('')}</sheets>${def.names ? `<definedNames>${def.names}</definedNames>` : ''}<calcPr calcId="191029"/>${def.wbExtra ?? ''}</workbook>`);
  zip.file('xl/_rels/workbook.xml.rels', relXml(wbRels));
  const root: RelDef[] = [{ id: 'rId1', type: 'officeDocument', target: 'xl/workbook.xml' }];
  if (def.app) {
    zip.file('docProps/app.xml', def.app);
    overrides.push(`<Override PartName="/docProps/app.xml" ContentType="${CT.app}"/>`);
    root.push({ id: 'rId2', type: 'extended-properties', target: 'docProps/app.xml' });
  }
  zip.file('_rels/.rels', relXml(root));
  zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>${overrides.join('')}</Types>`);
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

// ---------------------------------------------------------------------------------------------------
// cells
// ---------------------------------------------------------------------------------------------------

export const num = (ref: string, v: number, s = 0) => `<c r="${ref}"${s ? ` s="${s}"` : ''}><v>${v}</v></c>`;
export const str = (ref: string, text: string, s = 0) => `<c r="${ref}"${s ? ` s="${s}"` : ''} t="inlineStr"><is><t>${text.replace(/&/g, '&amp;')}</t></is></c>`;
/** a formula cell with its cached result (type: n number, s string, b boolean, e error, '' no cached value) */
export const fml = (ref: string, f: string, v: string | number | boolean | null, type: 'n' | 's' | 'b' | 'e' | '' = 'n', extra = '') => {
  const t = type === 's' ? ' t="str"' : type === 'b' ? ' t="b"' : type === 'e' ? ' t="e"' : '';
  const body = v === null ? '' : `<v>${type === 'b' ? (v ? 1 : 0) : String(v)}</v>`;
  return `<c r="${ref}"${t}${extra}><f>${f.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')}</f>${body}</c>`;
};
export const row = (r: number, cells: string) => `<row r="${r}">${cells}</row>`;

// ---------------------------------------------------------------------------------------------------
// premerge of one base file (no sources): what the app does when the user only has one file
// ---------------------------------------------------------------------------------------------------

export interface OneFile {
  bytes: Uint8Array;
  fileName: string;
  /** merged sheet (default: the reader's pickDataSheet) */
  sheet?: string;
  /** TableRegion (default: detectRegion) */
  region?: TableRegion;
}

export async function baseSource(f: OneFile): Promise<KeySource> {
  const book = await readBook(f.bytes, f.fileName, { id: 'b' });
  const sheet = f.sheet ?? pickDataSheet(book);
  return { id: 'b', book, region: f.region ?? detectRegion(book, sheet), label: f.fileName.replace(/\.[^.]+$/, '') };
}

/** a base package whose reader view is a bare book (hand-made packages): the region has to be given */
export function bareBase(bytes: Uint8Array, region: TableRegion, label = 'base'): KeySource {
  return { id: 'b', label, region, book: fakeBook('b', `${label}.xlsx`, bytes) };
}

export const baseOptions = (extra: Partial<PremergeInput['options']> = {}): PremergeInput['options'] => ({ formulaPolicy: 'keep-if-possible', extendTable: true, headerConflict: 'suffix', ...extra });

/** premerge one base file without imports / appended rows */
export async function premergeOne(base: KeySource, extra: Partial<PremergeInput['options']> = { dropOtherSheets: true }): Promise<PremergeResult> {
  const rows = [];
  for (let r = base.region.firstDataRow; r <= base.region.lastDataRow; r++) rows.push({ baseRow: r, matches: {} });
  return buildPremerged({ base, sources: [], rows, imports: [], keyFill: {}, options: baseOptions(extra) });
}

export const readFixture = (...p: string[]) => new Uint8Array(fs.readFileSync(path.join(...p)));

// ---------------------------------------------------------------------------------------------------
// reading a package
// ---------------------------------------------------------------------------------------------------

export type Files = Map<string, string>;

/** every part of a package: text for xml / rels / vml, a placeholder for binary parts */
export async function files(bytes: Uint8Array): Promise<Files> {
  const zip = await JSZip.loadAsync(bytes);
  const out: Files = new Map();
  for (const n of Object.keys(zip.files)) {
    if (zip.files[n].dir) continue;
    out.set(n, /\.(xml|rels|vml)$/i.test(n) ? await zip.files[n].async('string') : '<binary>');
  }
  return out;
}

export const sheetParts = (f: Files) => [...f.keys()].filter((p) => /^xl\/(worksheets|chartsheets|dialogsheets)\/[^/]+\.xml$/i.test(p));

// ---------------------------------------------------------------------------------------------------
// the linter
// ---------------------------------------------------------------------------------------------------

const attr = (tag: string, name: string): string | undefined => new RegExp(`\\s${name}\\s*=\\s*"([^"]*)"`).exec(tag)?.[1];
const dirName = (p: string) => (p.includes('/') ? p.slice(0, p.lastIndexOf('/')) : '');
function resolve(owner: string, target: string): string {
  let t = target;
  try { t = decodeURIComponent(t); } catch { /* keep */ }
  if (t.startsWith('/')) return t.slice(1);
  const parts = dirName(owner) ? dirName(owner).split('/') : [];
  for (const seg of t.split('/')) { if (seg === '..') parts.pop(); else if (seg && seg !== '.') parts.push(seg); }
  return parts.join('/');
}

export interface DropLint {
  code: string;
  detail: string;
}

/**
 * Package-level rules of a single-sheet output: every relationship target exists, every content-type override has a part, every
 * part has a content type and is reachable from the package root (no orphans), no calcChain, exactly one worksheet and nothing
 * else of the sheet kinds, names scoped to sheet 0 only, bookViews pointing at the sheet, no pivot cache nobody uses, app.xml
 * without the removed sheet names.
 */
export async function lintSingleSheet(bytes: Uint8Array, removed: string[] = []): Promise<DropLint[]> {
  const f = await files(bytes);
  const out: DropLint[] = [];
  const add = (code: string, detail: string) => out.push({ code, detail });
  const lower = new Map([...f.keys()].map((k) => [k.toLowerCase(), k]));
  // part names are case-insensitive (a writer may upper-case them)
  const get = (p: string) => f.get(lower.get(p.toLowerCase()) ?? p);
  const ct = get('[Content_Types].xml') ?? '';
  const defaults = new Set([...ct.matchAll(/<Default\b[^>]*>/g)].map((m) => (attr(m[0], 'Extension') ?? '').toLowerCase()));
  const overrides = new Map([...ct.matchAll(/<Override\b[^>]*>/g)].map((m) => [(attr(m[0], 'PartName') ?? '').replace(/^\//, '').toLowerCase(), attr(m[0], 'ContentType') ?? '']));
  for (const k of overrides.keys()) if (!lower.has(k)) add('CT_DANGLING', k);
  for (const p of f.keys()) {
    if (p.toLowerCase() === '[content_types].xml') continue;
    const ext = p.includes('.') ? p.slice(p.lastIndexOf('.') + 1).toLowerCase() : '';
    if (!overrides.has(p.toLowerCase()) && !defaults.has(ext)) add('CT_NONE', p);
  }
  // relationships + reachability
  const relsOf = (owner: string) => {
    const i = owner.lastIndexOf('/');
    const rp = `${owner.slice(0, i + 1)}_rels/${owner.slice(i + 1)}.rels`;
    const x = get(rp);
    return x === undefined ? [] : [...x.matchAll(/<Relationship\b[^>]*>/g)].map((m) => ({ id: attr(m[0], 'Id') ?? '', type: attr(m[0], 'Type') ?? '', target: attr(m[0], 'Target') ?? '', external: /TargetMode\s*=\s*"External"/.test(m[0]) }));
  };
  for (const p of f.keys()) {
    if (!/\.rels$/i.test(p)) continue;
    const owner = p.replace(/(^|\/)_rels\/([^/]*)\.rels$/i, (_m, d: string, n: string) => d + n);
    const isRoot = p.toLowerCase() === '_rels/.rels';
    if (!isRoot && !lower.has(owner.toLowerCase())) add('RELS_ORPHAN', `${p} belongs to a part that does not exist`);
    for (const r of relsOf(isRoot ? '' : owner)) {
      if (r.external) continue;
      const t = resolve(isRoot ? '' : owner, r.target);
      if (!lower.has(t.toLowerCase())) add('REL_DANGLING', `${p}: ${r.id} -> ${t}`);
    }
  }
  const seen = new Set<string>();
  const queue = [''];
  while (queue.length) {
    const o = queue.pop()!;
    for (const r of relsOf(o)) {
      if (r.external) continue;
      const t = lower.get(resolve(o, r.target).toLowerCase());
      if (!t || seen.has(t)) continue;
      seen.add(t);
      queue.push(t);
    }
  }
  for (const p of f.keys()) {
    if (p.toLowerCase() === '[content_types].xml' || /\.rels$/i.test(p) || seen.has(p)) continue;
    add('ORPHAN', p);
  }
  if (lower.has('xl/calcchain.xml')) add('CALCCHAIN', 'xl/calcChain.xml');
  // sheets
  const wbPath = 'xl/workbook.xml';
  const wb = get(wbPath) ?? '';
  const sheets = [...wb.matchAll(/<sheet\b[^>]*>/g)].map((m) => m[0]);
  if (sheets.length !== 1) add('SHEET_COUNT', `${sheets.length} sheets in workbook.xml`);
  const parts = sheetParts(f);
  if (parts.length !== 1 || !/worksheets\//i.test(parts[0] ?? '')) add('SHEET_PARTS', parts.join(','));
  if (sheets.some((s) => /\sstate="(hidden|veryHidden)"/.test(s))) add('SHEET_HIDDEN', sheets.join(''));
  for (const m of wb.matchAll(/<definedName\b[^>]*>/g)) {
    const ls = attr(m[0], 'localSheetId');
    if (ls !== undefined && ls !== '0') add('NAME_SCOPE', m[0]);
  }
  const view = /<workbookView\b[^>]*>/.exec(wb)?.[0] ?? '';
  const at = Number(attr(view, 'activeTab') ?? 0);
  const fs0 = Number(attr(view, 'firstSheet') ?? 0);
  if (at !== 0 || fs0 !== 0) add('BOOKVIEW', view);
  // pivot caches nobody uses
  const pivotDefs = [...f.keys()].filter((p) => /pivotTables\/[^/]+\.xml$/i.test(p));
  for (const c of [...f.keys()].filter((p) => /pivotCacheDefinition[^/]*\.xml$/i.test(p))) {
    const used = pivotDefs.some((pt) => relsOf(pt).some((r) => resolve(pt, r.target).toLowerCase() === c.toLowerCase()));
    if (!used) add('PIVOT_CACHE_UNUSED', c);
  }
  // <pivotCaches> entries have to resolve
  for (const m of wb.matchAll(/<pivotCache\b[^>]*>/g)) {
    const rid = attr(m[0], 'r:id');
    if (!relsOf(wbPath).some((r) => r.id === rid)) add('PIVOTCACHE_REL', m[0]);
  }
  // app.xml names no removed sheet
  const app = get('docProps/app.xml');
  if (app) {
    const titles = [...(/<TitlesOfParts>([\s\S]*?)<\/TitlesOfParts>/.exec(app)?.[1] ?? '').matchAll(/<vt:lpstr>([\s\S]*?)<\/vt:lpstr>/g)].map((m) => m[1].replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>'));
    for (const n of removed) if (titles.includes(n) || titles.some((t) => t.startsWith(`${n}!`) || t.startsWith(`'${n}'!`))) add('APP_STALE', `docProps/app.xml still lists ${n}`);
    const hp = /<HeadingPairs>([\s\S]*?)<\/HeadingPairs>/.exec(app)?.[1];
    if (hp !== undefined) {
      const sum = [...hp.matchAll(/<vt:i4>\s*(\d+)\s*<\/vt:i4>/g)].reduce((a, m) => a + Number(m[1]), 0);
      if (sum !== titles.length) add('APP_COUNT', `HeadingPairs count ${sum} vs ${titles.length} titles`);
    }
  }
  // names of the removed sheets nowhere in workbook-level formulas
  for (const n of removed) {
    const q = [`${n}!`, `'${n.replace(/'/g, "''")}'!`];
    for (const [pn, x] of f) {
      if (!/\.xml$/i.test(pn) || /^docProps\//i.test(pn)) continue;
      if (/<f\b|<formula\b|<xm:f\b|<c:f\b|<definedName\b|location=/.test(x) && q.some((s) => x.includes(s.replace(/&/g, '&amp;')))) add('REMOVED_SHEET_NAMED', `${pn} still mentions ${n}`);
    }
  }
  return out;
}

export const lintText = (l: DropLint[]) => l.map((x) => `${x.code}: ${x.detail}`).join('\n');

/** a formula (FILE form) with a reference qualified by one of these sheets (tokens, 3D ends included) */
export function readsAnySheet(formula: string, sheets: string[]): boolean {
  const lower = new Set(sheets.map((n) => n.toLowerCase()));
  for (const t of tokenize(formula)) {
    const p = t.prefix as (Prefix & { sheet1?: string }) | undefined;
    if (!p || p.book !== undefined) continue;
    for (const n of [p.sheet, p.sheet1, p.sheet2]) if (n && lower.has(n.toLowerCase())) return true;
  }
  return false;
}

/** the merged workbook as the reader sees it */
export async function readOut(bytes: Uint8Array, name = 'out.xlsx'): Promise<SourceBook> {
  return readBook(bytes, name, { id: 'out' });
}
