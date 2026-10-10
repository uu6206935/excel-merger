/**
 * Helpers of the sheet-copier tests (PremergeOptions.includeSourceSheets, round 9):
 *  - lintPackage: the package linter every output has to pass (also used by task A2's tests: "the package linter from A1's tests must pass"):
 *    relationships / content types / orphans / unique sheet ids and names / defined names / style, string, dxf and metadata indexes in range /
 *    worksheet element order / counts / x14 links / app.xml consistency / well-formed XML.
 *  - premergeCopies: the app's pipeline reduced to what the copier needs (a base file, other files, options) -> PremergeResult.
 *  - compareSheets: copy vs source, cell by cell through the reader (values, formulas with renamed sheet references, resolved styles),
 *    plus columns / rows / merges / rules / hyperlinks / freeze.
 *  - plannedNames: the expected names (planSheetNames) of a set of files.
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';
import type { KeySource, PremergeInput, PremergeResult } from '../../src/core/contracts';
import type { Cell, SheetGrid, SourceBook, TableRegion } from '../../src/core/types';
import { tokenize, quoteSheetName } from '../../src/core/formula/excelRef';
import { excelWidthToPx, mdwFor } from '../../src/core/display/styles';
import { knownMdw } from '../../src/core/xlsx/patch/bookinfo';
import { detectRegion, pickDataSheet } from '../../src/core/keys';
import { buildPremerged, readBook } from '../../src/core/xlsx';
import { xmlProblem } from '../../src/core/xlsx/patch/wellformed';
import { planSheetNames } from '../../src/core/xlsx/sheetnames';

export const REPO = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../..');
export const readFix = (...p: string[]): Uint8Array => new Uint8Array(fs.readFileSync(path.join(REPO, ...p)));

// ---------------------------------------------------------------------------------------------------------------------
// package reading
// ---------------------------------------------------------------------------------------------------------------------

export type Files = Map<string, string>;

/** every part: text for xml / rels / vml, a placeholder for binary parts */
export async function readFiles(bytes: Uint8Array): Promise<Files> {
  const zip = await JSZip.loadAsync(bytes);
  const out: Files = new Map();
  for (const n of Object.keys(zip.files)) {
    if (zip.files[n].dir) continue;
    out.set(n, /\.(xml|rels|vml)$/i.test(n) ? await zip.files[n].async('string') : '<binary>');
  }
  return out;
}

const attr = (tag: string, name: string): string | undefined => new RegExp(`\\s${name}\\s*=\\s*"([^"]*)"`).exec(tag)?.[1];
const unq = (s: string) => s.replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&');
const dirName = (p: string) => (p.includes('/') ? p.slice(0, p.lastIndexOf('/')) : '');
function resolve(owner: string, target: string): string {
  let t = target;
  try { t = decodeURIComponent(t); } catch { /* keep */ }
  if (t.startsWith('/')) return t.slice(1);
  const parts = dirName(owner) ? dirName(owner).split('/') : [];
  for (const seg of t.split('/')) { if (seg === '..') parts.pop(); else if (seg && seg !== '.') parts.push(seg); }
  return parts.join('/');
}

export interface SheetRef {
  name: string;
  sheetId: string;
  state: string;
  rid: string;
  part: string;
}

/** the sheets of workbook.xml in order with their parts */
export function sheetsOf(f: Files): SheetRef[] {
  const wb = f.get('xl/workbook.xml') ?? '';
  const rels = new Map([...(f.get('xl/_rels/workbook.xml.rels') ?? '').matchAll(/<Relationship\b[^>]*>/g)].map((m) => [attr(m[0], 'Id') ?? '', attr(m[0], 'Target') ?? '']));
  return [...wb.matchAll(/<sheet\b[^>]*>/g)].map((m) => ({
    name: unq(attr(m[0], 'name') ?? ''),
    sheetId: attr(m[0], 'sheetId') ?? '',
    state: attr(m[0], 'state') ?? 'visible',
    rid: attr(m[0], 'r:id') ?? '',
    part: resolve('xl/workbook.xml', rels.get(attr(m[0], 'r:id') ?? '') ?? ''),
  }));
}

// ---------------------------------------------------------------------------------------------------------------------
// the linter
// ---------------------------------------------------------------------------------------------------------------------

export interface Lint {
  code: string;
  detail: string;
}

/** worksheet children in schema order */
export const WS_ORDER = ['sheetPr', 'dimension', 'sheetViews', 'sheetFormatPr', 'cols', 'sheetData', 'sheetCalcPr', 'sheetProtection', 'protectedRanges', 'scenarios', 'autoFilter', 'sortState', 'dataConsolidate', 'customSheetViews', 'mergeCells', 'phoneticPr', 'conditionalFormatting', 'dataValidations', 'hyperlinks', 'printOptions', 'pageMargins', 'pageSetup', 'headerFooter', 'rowBreaks', 'colBreaks', 'customProperties', 'cellWatches', 'ignoredErrors', 'smartTags', 'drawing', 'legacyDrawing', 'legacyDrawingHF', 'drawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst'];

function topLevel(xml: string): Array<{ tag: string; text: string }> {
  // the children of <worksheet>: skip the open tag, walk siblings
  const open = /<worksheet\b[^>]*>/.exec(xml);
  if (!open) return [];
  const out: Array<{ tag: string; text: string }> = [];
  let i = open.index + open[0].length;
  const re = /<([A-Za-z_][\w.:-]*)(?=[\s/>])[^>]*?(\/?)>/y;
  while (i < xml.length) {
    const lt = xml.indexOf('<', i);
    if (lt < 0 || xml.startsWith('</', lt)) break;
    re.lastIndex = lt;
    const m = re.exec(xml);
    if (!m) break;
    if (m[2] === '/') { out.push({ tag: m[1], text: m[0] }); i = lt + m[0].length; continue; }
    const close = `</${m[1]}>`;
    const e = xml.indexOf(close, lt);
    if (e < 0) break;
    out.push({ tag: m[1], text: xml.slice(lt, e + close.length) });
    i = e + close.length;
  }
  return out;
}

export interface LintOptions {
  /** the number of worksheets the output must hold (undefined: not checked) */
  worksheets?: number;
}

/**
 * Package-level rules of an output workbook (every output of the app has to pass them):
 *  parts / relationships: every relationship target exists, every content-type override has a part and every part a content type, no part
 *  is unreachable from the package root, no calcChain, part names unique ignoring case;
 *  workbook: sheetIds and sheet names unique (case-insensitive), names legal (31 characters, none of \ / ? * [ ] :), every <sheet> r:id
 *  leads to a worksheet / chartsheet part, definedName localSheetId in range and (name, scope) unique, a visible sheet exists;
 *  worksheets: children in schema order, rows and cells strictly increasing and inside their row, style ids / shared-string ids / dxf ids /
 *  cell metadata ids in range, every r:id of the sheet is a relationship of it, counts equal the children, x14 rule links resolve;
 *  styles: every xf points at existing fonts / fills / borders / number formats, counts equal the children; shared strings: uniqueCount;
 *  app.xml: HeadingPairs agree with TitlesOfParts; every XML part is well-formed.
 */
export async function lintPackage(bytes: Uint8Array, opts: LintOptions = {}): Promise<Lint[]> {
  const f = await readFiles(bytes);
  const out: Lint[] = [];
  const add = (code: string, detail: string) => out.push({ code, detail });
  const lower = new Map<string, string>();
  for (const k of f.keys()) {
    if (lower.has(k.toLowerCase())) add('PART_DUPLICATE_CASE', k);
    lower.set(k.toLowerCase(), k);
  }
  const get = (p: string) => f.get(lower.get(p.toLowerCase()) ?? p);
  // ---- content types
  const ct = get('[Content_Types].xml') ?? '';
  const defaults = new Set([...ct.matchAll(/<Default\b[^>]*>/g)].map((m) => (attr(m[0], 'Extension') ?? '').toLowerCase()));
  const overrides = new Map([...ct.matchAll(/<Override\b[^>]*>/g)].map((m) => [(attr(m[0], 'PartName') ?? '').replace(/^\//, '').toLowerCase(), attr(m[0], 'ContentType') ?? '']));
  for (const k of overrides.keys()) if (!lower.has(k)) add('CT_DANGLING', k);
  for (const p of f.keys()) {
    if (p.toLowerCase() === '[content_types].xml') continue;
    const ext = p.includes('.') ? p.slice(p.lastIndexOf('.') + 1).toLowerCase() : '';
    if (!overrides.has(p.toLowerCase()) && !defaults.has(ext)) add('CT_NONE', p);
  }
  // ---- relationships + reachability
  const relsOf = (owner: string) => {
    const i = owner.lastIndexOf('/');
    const x = get(`${owner.slice(0, i + 1)}_rels/${owner.slice(i + 1)}.rels`);
    return x === undefined ? [] : [...x.matchAll(/<Relationship\b[^>]*>/g)].map((m) => ({ id: attr(m[0], 'Id') ?? '', type: attr(m[0], 'Type') ?? '', target: unq(attr(m[0], 'Target') ?? ''), external: /TargetMode\s*=\s*"External"/.test(m[0]) }));
  };
  for (const p of f.keys()) {
    if (!/\.rels$/i.test(p)) continue;
    const owner = p.replace(/(^|\/)_rels\/([^/]*)\.rels$/i, (_m, d: string, n: string) => d + n);
    const isRoot = p.toLowerCase() === '_rels/.rels';
    if (!isRoot && !lower.has(owner.toLowerCase())) add('RELS_ORPHAN', `${p} belongs to a part that does not exist`);
    const rs = relsOf(isRoot ? '' : owner);
    const seenIds = new Set<string>();
    for (const r of rs) {
      if (seenIds.has(r.id)) add('REL_ID_DUPLICATE', `${p}: ${r.id}`);
      seenIds.add(r.id);
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
  for (const [p, x] of f) if (/\.(xml|rels)$/i.test(p)) { const why = xmlProblem(x); if (why) add('XML_NOT_WELLFORMED', `${p}: ${why}`); }
  // ---- workbook
  const wb = f.get('xl/workbook.xml') ?? '';
  const sheets = sheetsOf(f);
  const ids = new Set<string>();
  const names = new Set<string>();
  for (const s of sheets) {
    if (ids.has(s.sheetId)) add('SHEETID_DUPLICATE', s.sheetId);
    ids.add(s.sheetId);
    const l = s.name.toLowerCase();
    if (names.has(l)) add('SHEETNAME_DUPLICATE', s.name);
    names.add(l);
    if (s.name.length > 31 || s.name.length === 0 || /[\\/?*[\]:]/.test(s.name) || /^'|'$/.test(s.name)) add('SHEETNAME_ILLEGAL', s.name);
    if (!lower.has(s.part.toLowerCase())) add('SHEET_PART_MISSING', `${s.name}: ${s.part}`);
  }
  if (!sheets.some((s) => s.state === 'visible')) add('NO_VISIBLE_SHEET', '');
  const view = /<workbookView\b[^>]*>/.exec(wb)?.[0] ?? '';
  const activeTab = Number(attr(view, 'activeTab') ?? 0);
  if (activeTab >= sheets.length || (sheets[activeTab] && sheets[activeTab].state !== 'visible')) add('ACTIVE_TAB', view);
  const nameKeys = new Set<string>();
  for (const m of wb.matchAll(/<definedName\b[^>]*>/g)) {
    const ls = attr(m[0], 'localSheetId');
    if (ls !== undefined && !(Number(ls) >= 0 && Number(ls) < sheets.length)) add('NAME_SCOPE', m[0]);
    const key = `${(attr(m[0], 'name') ?? '').toLowerCase()}|${ls ?? ''}`;
    if (nameKeys.has(key)) add('NAME_DUPLICATE', m[0]);
    nameKeys.add(key);
  }
  if (opts.worksheets !== undefined) {
    const n = [...f.keys()].filter((p) => /^xl\/worksheets\/[^/]+\.xml$/i.test(p)).length;
    if (n !== opts.worksheets) add('WORKSHEET_COUNT', `${n} != ${opts.worksheets}`);
  }
  // ---- styles / strings / metadata limits
  const styles = get('xl/styles.xml') ?? '';
  const count = (tag: string, child: string) => {
    const m = new RegExp(`<${tag}\\b[^>]*?(?:/>|>([\\s\\S]*?)</${tag}>)`).exec(styles);
    if (!m) return { n: 0, declared: undefined as number | undefined };
    const open = /^<[^>]*>/.exec(m[0])![0];
    const n = m[1] ? (m[1].match(new RegExp(`<${child}(?=[\\s/>])`, 'g')) ?? []).length : 0;
    const declared = attr(open, 'count');
    return { n, declared: declared === undefined ? undefined : Number(declared) };
  };
  const nXf = count('cellXfs', 'xf'), nFont = count('fonts', 'font'), nFill = count('fills', 'fill'), nBorder = count('borders', 'border');
  const nDxf = count('dxfs', 'dxf'), nNumFmt = count('numFmts', 'numFmt');
  for (const [tag, c] of [['cellXfs', nXf], ['fonts', nFont], ['fills', nFill], ['borders', nBorder], ['dxfs', nDxf], ['numFmts', nNumFmt]] as const) {
    if (c.declared !== undefined && c.declared !== c.n) add('STYLES_COUNT', `${tag} count=${c.declared} but ${c.n} children`);
  }
  if (nXf.n > 64000) add('XF_LIMIT', String(nXf.n));
  const customFmts = new Set([...styles.matchAll(/<numFmt\b[^>]*>/g)].map((m) => Number(attr(m[0], 'numFmtId'))));
  const xfBlock = /<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(styles)?.[1] ?? '';
  let xi = 0;
  for (const m of xfBlock.matchAll(/<xf\b[^>]*>/g)) {
    const o = m[0];
    const fo = Number(attr(o, 'fontId') ?? 0), fi = Number(attr(o, 'fillId') ?? 0), bo = Number(attr(o, 'borderId') ?? 0), nf = Number(attr(o, 'numFmtId') ?? 0);
    if (fo >= nFont.n) add('XF_FONT', `xf ${xi}: font ${fo}`);
    if (fi >= nFill.n) add('XF_FILL', `xf ${xi}: fill ${fi}`);
    if (bo >= nBorder.n) add('XF_BORDER', `xf ${xi}: border ${bo}`);
    if (nf >= 164 && !customFmts.has(nf)) add('XF_NUMFMT', `xf ${xi}: numFmt ${nf}`);
    xi++;
  }
  const sst = get('xl/sharedStrings.xml');
  const nSst = sst === undefined ? 0 : (sst.match(/<si(?=[\s/>])/g) ?? []).length;
  if (sst !== undefined) {
    const u = attr(/<sst\b[^>]*>/.exec(sst)?.[0] ?? '', 'uniqueCount');
    if (u !== undefined && Number(u) !== nSst) add('SST_UNIQUE', `uniqueCount=${u} but ${nSst} items`);
  }
  const metadata = get('xl/metadata.xml');
  const nCellMeta = metadata === undefined ? 0 : (/<cellMetadata\b[^>]*>([\s\S]*?)<\/cellMetadata>/.exec(metadata)?.[1].match(/<bk>/g) ?? []).length;
  // ---- worksheets
  for (const s of sheets) {
    const x = get(s.part);
    if (x === undefined || !/\/worksheets\//i.test(s.part)) continue;
    const kids = topLevel(x);
    let prev = -1;
    for (const k of kids) {
      const i = WS_ORDER.indexOf(k.tag);
      if (i < 0) add('WS_UNKNOWN_ELEMENT', `${s.name}: <${k.tag}>`);
      else if (i < prev) add('WS_ORDER', `${s.name}: <${k.tag}> after ${WS_ORDER[prev]}`);
      else prev = i;
    }
    const sheetRels = relsOf(s.part);
    for (const m of x.matchAll(/\sr:id="([^"]*)"/g)) if (!sheetRels.some((r) => r.id === m[1])) add('WS_RID_DANGLING', `${s.name}: r:id ${m[1]}`);
    for (const m of x.matchAll(/\sr:(?!id=)[a-z]+="/g)) add('WS_RATTR', `${s.name}: ${m[0]}`);
    // rows / cells
    let lastRow = 0;
    for (const rm of x.matchAll(/<row\b([^>]*?)(?:\/>|>([\s\S]*?)<\/row>)/g)) {
      const r = Number(attr(rm[1], 'r') ?? lastRow + 1);
      if (r <= lastRow) add('ROW_ORDER', `${s.name}: row ${r} after ${lastRow}`);
      lastRow = r;
      const rs = attr(rm[1], 's');
      if (rs !== undefined && Number(rs) >= nXf.n) add('ROW_STYLE', `${s.name}: row ${r} s=${rs}`);
      let lastC = 0;
      for (const cm of (rm[2] ?? '').matchAll(/<c\b([^>]*?)(\/>|>([\s\S]*?)<\/c>)/g)) {
        const ref = attr(cm[1], 'r') ?? '';
        const mm = /^([A-Z]+)(\d+)$/.exec(ref);
        if (!mm || Number(mm[2]) !== r) { add('CELL_REF', `${s.name}: ${ref} in row ${r}`); continue; }
        let col = 0;
        for (const ch of mm[1]) col = col * 26 + ch.charCodeAt(0) - 64;
        if (col <= lastC) add('CELL_ORDER', `${s.name}: ${ref}`);
        lastC = col;
        const st = attr(cm[1], 's');
        if (st !== undefined && Number(st) >= nXf.n) add('CELL_STYLE', `${s.name}: ${ref} s=${st}`);
        const t = attr(cm[1], 't');
        if (t === 's') {
          const v = /<v>\s*(\d+)\s*<\/v>/.exec(cm[3] ?? '');
          if (!v || Number(v[1]) >= nSst) add('CELL_SST', `${s.name}: ${ref} -> ${v?.[1]}`);
        }
        const cmI = attr(cm[1], 'cm');
        if (cmI !== undefined && (Number(cmI) < 1 || Number(cmI) > nCellMeta)) add('CELL_CM', `${s.name}: ${ref} cm=${cmI}`);
        if (attr(cm[1], 'vm') !== undefined) add('CELL_VM', `${s.name}: ${ref} has vm`);
        if (cm[3] && /<f\b[^>]*t="shared"/.test(cm[3])) add('CELL_SHARED', `${s.name}: ${ref} has a shared formula`);
      }
    }
    for (const m of x.matchAll(/<col\b[^>]*>/g)) {
      const st = attr(m[0], 'style');
      if (st !== undefined && Number(st) >= nXf.n) add('COL_STYLE', `${s.name}: ${m[0]}`);
    }
    for (const m of x.matchAll(/\sdxfId="(\d+)"/g)) if (Number(m[1]) >= nDxf.n) add('DXF_RANGE', `${s.name}: dxfId=${m[1]}`);
    // counts
    for (const [container, child] of [['mergeCells', 'mergeCell'], ['dataValidations', 'dataValidation'], ['hyperlinks', 'hyperlink']] as const) {
      const c = new RegExp(`<${container}\\b([^>]*)>([\\s\\S]*?)</${container}>`).exec(x);
      if (!c) continue;
      const n = (c[2].match(new RegExp(`<${child}(?=[\\s/>])`, 'g')) ?? []).length;
      const d = attr(c[1], 'count');
      if (d !== undefined && Number(d) !== n) add('COUNT', `${s.name}: ${container} count=${d} but ${n}`);
      if (n === 0) add('EMPTY_CONTAINER', `${s.name}: ${container}`);
    }
    if (/<(dataValidations|mergeCells|hyperlinks|conditionalFormatting)\b[^>]*\/>/.test(x)) add('EMPTY_CONTAINER', `${s.name}: self-closed container`);
    // x14 links
    const classic = new Set([...x.matchAll(/<x14:id>([^<]*)<\/x14:id>/g)].map((m) => m[1]));
    const ext = new Set([...x.matchAll(/<x14:cfRule\b[^>]*\sid="([^"]*)"/g)].map((m) => m[1]));
    for (const g of classic) if (!ext.has(g)) add('X14_LINK', `${s.name}: classic rule links ${g} but there is no x14 rule`);
    // priorities
    const prios = [...x.matchAll(/<(?:x14:)?cfRule\b[^>]*\spriority="(\d+)"/g)].map((m) => Number(m[1]));
    if (new Set(prios).size !== prios.length) add('CF_PRIORITY', `${s.name}: duplicate priorities`);
    // sheet views: at most one selected tab in the whole workbook is checked below
  }
  const selected = sheets.filter((s) => /<sheetView\b[^>]*\stabSelected="(1|true)"/.test(get(s.part) ?? '')).length;
  if (selected > 1) add('TAB_SELECTED', `${selected} sheets are selected (grouped)`);
  // ---- app.xml
  const app = get('docProps/app.xml');
  if (app) {
    const titles = [...(/<TitlesOfParts>([\s\S]*?)<\/TitlesOfParts>/.exec(app)?.[1] ?? '').matchAll(/<vt:lpstr>([\s\S]*?)<\/vt:lpstr>/g)].map((m) => unq(m[1]));
    const hp = /<HeadingPairs>([\s\S]*?)<\/HeadingPairs>/.exec(app)?.[1];
    if (hp !== undefined) {
      const sum = [...hp.matchAll(/<vt:i4>\s*(\d+)\s*<\/vt:i4>/g)].reduce((a, m) => a + Number(m[1]), 0);
      if (sum !== titles.length) add('APP_COUNT', `HeadingPairs count ${sum} vs ${titles.length} titles`);
      const wsCount = Number(/<vt:i4>\s*(\d+)\s*<\/vt:i4>/.exec(hp)?.[1] ?? -1);
      const nWs = sheets.filter((s) => /\/worksheets\//i.test(s.part)).length;
      if (wsCount !== nWs) add('APP_SHEETS', `HeadingPairs lists ${wsCount} worksheets, the workbook has ${nWs}`);
      const first = titles.slice(0, wsCount);
      const expect = sheets.filter((s) => /\/worksheets\//i.test(s.part)).map((s) => s.name);
      if (first.join('\u0000') !== expect.join('\u0000')) add('APP_TITLES', `${first.join('|')} != ${expect.join('|')}`);
    }
  }
  return out;
}

export const lintText = (l: Lint[]): string => l.map((x) => `${x.code}: ${x.detail}`).join('\n');

// ---------------------------------------------------------------------------------------------------------------------
// premerge with copies
// ---------------------------------------------------------------------------------------------------------------------

export interface FileIn {
  bytes: Uint8Array;
  fileName: string;
  /** the data sheet (default: pickDataSheet) */
  sheet?: string;
  /** region override */
  region?: TableRegion;
}

export async function keySource(f: FileIn, id: string): Promise<KeySource> {
  const book = await readBook(f.bytes, f.fileName, { id });
  const sheet = f.sheet ?? pickDataSheet(book);
  return { id, book, region: f.region ?? detectRegion(book, sheet), label: f.fileName.replace(/\.[^.]+$/, '') };
}

/** premerge `base` with the sheets of `others` copied (no imported column, no appended row: the merged sheet is the base data sheet as it is) */
export async function premergeCopies(base: FileIn, others: FileIn[], extra: Partial<PremergeInput['options']> = {}, mutate?: (input: PremergeInput) => void): Promise<PremergeResult & { input: PremergeInput }> {
  const b = await keySource(base, 'b');
  const sources = await Promise.all(others.map((o, i) => keySource(o, `s${i + 2}`)));
  const rows = [];
  for (let r = b.region.firstDataRow; r <= b.region.lastDataRow; r++) rows.push({ baseRow: r, matches: {} });
  const input: PremergeInput = {
    base: b, sources, rows, imports: [], keyFill: {},
    options: { formulaPolicy: 'keep-if-possible', extendTable: true, headerConflict: 'suffix', includeSourceSheets: true, dropOtherSheets: false, ...extra },
  };
  mutate?.(input);
  const res = await buildPremerged(input);
  return { ...res, input };
}

/** the names planSheetNames gives (the contract) */
export function plannedNames(existing: string[], files: Array<{ fileName: string; sheets: string[] }>): string[] {
  return planSheetNames(existing, files.map((f) => ({ file: f.fileName, sheets: f.sheets }))).map((p) => p.name);
}

/** the worksheet names of a package in workbook order */
export async function sheetNamesOf(bytes: Uint8Array): Promise<string[]> {
  return sheetsOf(await readFiles(bytes)).map((s) => s.name);
}

// ---------------------------------------------------------------------------------------------------------------------
// copy vs source
// ---------------------------------------------------------------------------------------------------------------------

/** every sheet reference of a formula renamed by `map` (lower-cased old -> new), simultaneous (a chain A->B, B->C works), token based */
export function renameRefs(formula: string, map: ReadonlyMap<string, string>): string {
  const refText = /INDIRECT|HYPERLINK/i.test(formula);
  return tokenize(formula).map((t) => {
    if (t.type === 'string' && refText) {
      // INDIRECT("Sheet2!A1") / HYPERLINK("#Sheet2!A1"): a string that is exactly one reference
      const body = t.text.slice(1, -1).replace(/""/g, '"');
      const hash = body.startsWith('#') ? '#' : '';
      const inner = tokenize(body.slice(hash.length));
      const q = inner.length === 1 ? inner[0] : undefined;
      if (q?.prefix && q.prefix.book === undefined && q.prefix.sheet != null && (q.type === 'ref' || q.type === 'name')) {
        const nn = map.get(q.prefix.sheet.toLowerCase());
        if (nn !== undefined) return '"' + (hash + quoteSheetName(nn) + '!' + q.text.slice(q.prefix.text.length)).replace(/"/g, '""') + '"';
      }
      return t.text;
    }
    const p = t.prefix;
    if (!p || p.book !== undefined || p.sheet == null) return t.text;
    const nn = map.get(p.sheet.toLowerCase());
    return nn === undefined ? t.text : quoteSheetName(nn) + '!' + t.text.slice(p.text.length);
  }).join('');
}

export interface Diff {
  where: string;
  detail: string;
}

const styleKey = (book: SourceBook, c: Cell | undefined, grid: SheetGrid, row: number, col: number): string => {
  void grid; void row; void col;
  return JSON.stringify(book.styles[c?.s ?? 0] ?? {});
};

export interface CompareOptions {
  /** lower-cased source sheet name -> copy name (formulas of the source are renamed with it before the comparison) */
  rename: ReadonlyMap<string, string>;
  /** formulas the copy may hold as values (reason text) */
  allowValueFormulas?: boolean;
  /** compare the resolved styles (fill / font / border / alignment / number format) */
  styles?: boolean;
  /** date serial difference the copy may have (1904 source: +1462) */
  dateShift?: number;
  /** skip the compare of rich text runs */
  noRich?: boolean;
}

/** Compare a copied sheet with its source through the reader's view of both books; returns the differences (empty = identical). */
export function compareSheets(src: SourceBook, srcName: string, out: SourceBook, outName: string, o: CompareOptions): Diff[] {
  const diffs: Diff[] = [];
  const a = src.sheets.find((s) => s.name === srcName);
  const b = out.sheets.find((s) => s.name === outName);
  if (!a || !b) return [{ where: 'sheet', detail: `missing ${a ? '' : srcName} ${b ? '' : outName}` }];
  const add = (where: string, detail: string) => { if (diffs.length < 40) diffs.push({ where, detail }); };
  if (a.state !== b.state) add('state', `${a.state} != ${b.state}`);
  if (a.maxRow !== b.maxRow || a.maxCol !== b.maxCol) add('extent', `${a.maxRow}x${a.maxCol} != ${b.maxRow}x${b.maxCol}`);
  const rows = Math.max(a.maxRow, b.maxRow), cols = Math.max(a.maxCol, b.maxCol);
  for (let r = 1; r <= rows; r++) {
    for (let c = 1; c <= cols; c++) {
      const ca = a.rows[r - 1]?.[c - 1], cb = b.rows[r - 1]?.[c - 1];
      const ref = `${colName(c)}${r}`;
      const emptyA = !ca || (ca.t === 'z' && ca.v == null && !ca.f), emptyB = !cb || (cb.t === 'z' && cb.v == null && !cb.f);
      // a formula that could not be copied and stored no result is an empty cell in the copy (the warning says so)
      if (o.allowValueFormulas && ca?.f !== undefined && ca.v == null && emptyB) continue;
      if (emptyA && emptyB) { if (o.styles && ca && cb && styleKey(src, ca, a, r, c) !== styleKey(out, cb, b, r, c)) add(ref, `style ${styleKey(src, ca, a, r, c)} != ${styleKey(out, cb, b, r, c)}`); continue; }
      if (emptyA !== emptyB) { add(ref, `${emptyA ? 'empty' : JSON.stringify(ca)} != ${emptyB ? 'empty' : JSON.stringify(cb)}`); continue; }
      const ea = ca!.f !== undefined ? renameRefs(ca!.f, o.rename) : undefined;
      if (ea !== cb!.f) {
        if (!(o.allowValueFormulas && cb!.f === undefined)) add(ref, `formula ${ea} != ${cb!.f}`);
      }
      const shift = o.dateShift && ca!.t === 'n' ? o.dateShift : 0;
      const va = typeof ca!.v === 'number' && shift ? ca!.v + shift : ca!.v;
      if (ca!.t !== cb!.t && !(ca!.t === 'd' && cb!.t === 'n') && !(ca!.t === 'n' && cb!.t === 'd')) add(ref, `type ${ca!.t} != ${cb!.t}`);
      else if (!sameValue(va, cb!.v)) add(ref, `value ${JSON.stringify(va)} != ${JSON.stringify(cb!.v)}`);
      if (!o.noRich && JSON.stringify(ca!.rich?.map((x) => x.text)) !== JSON.stringify(cb!.rich?.map((x) => x.text))) add(ref, 'rich text runs differ');
      if (o.styles && styleKey(src, ca, a, r, c) !== styleKey(out, cb, b, r, c)) add(ref, `style ${styleKey(src, ca, a, r, c)} != ${styleKey(out, cb, b, r, c)}`);
      if ((ca!.arr ?? '') !== (cb!.arr ?? '') || !!ca!.dyn !== !!cb!.dyn) add(ref, `array ${ca!.arr}/${ca!.dyn} != ${cb!.arr}/${cb!.dyn}`);
    }
  }
  // columns / rows
  const mdwA = knownMdw(src.styles[0]?.font, src.styles[0]?.size) ?? mdwFor(src.styles[0]?.font, src.styles[0]?.size), mdwB = knownMdw(out.styles[0]?.font, out.styles[0]?.size) ?? mdwFor(out.styles[0]?.font, out.styles[0]?.size);
  const nCols = Math.max(a.colProps.length, b.colProps.length);
  for (let c = 0; c < nCols; c++) {
    const pa = a.colProps[c], pb = b.colProps[c];
    // column widths are stored in characters of each workbook's default font: a copy keeps either the stored width (default fonts alike or not
    // known: scheme fonts and theme fonts resolve differently in the reader) or the PIXEL width (to 1 px); addsheets-width.test.ts checks the rule itself
    const wa = pa?.width, wb = pb?.width;
    const sameW = wa === wb || (wa !== undefined && wb !== undefined && Math.abs(excelWidthToPx(wa, mdwA) - excelWidthToPx(wb, mdwB)) <= 1);
    const ka = JSON.stringify({ h: !!pa?.hidden, o: pa?.outlineLevel ?? 0 });
    const kb = JSON.stringify({ h: !!pb?.hidden, o: pb?.outlineLevel ?? 0 });
    if (ka !== kb || !sameW) add(`col ${colName(c + 1)}`, `${wa}${ka} != ${wb}${kb}`);
    if (o.styles && ((pa?.s ?? 0) || (pb?.s ?? 0)) && JSON.stringify(src.styles[pa?.s ?? 0]) !== JSON.stringify(out.styles[pb?.s ?? 0])) add(`col ${colName(c + 1)}`, 'column style');
  }
  const nRows = Math.max(a.rowProps.length, b.rowProps.length);
  for (let r = 0; r < nRows; r++) {
    const pa = a.rowProps[r], pb = b.rowProps[r];
    const ka = JSON.stringify({ h: pa?.height, x: !!pa?.hidden, o: pa?.outlineLevel ?? 0 });
    const kb = JSON.stringify({ h: pb?.height, x: !!pb?.hidden, o: pb?.outlineLevel ?? 0 });
    if (ka !== kb) add(`row ${r + 1}`, `${ka} != ${kb}`);
  }
  // features
  const fa = a.features, fb = b.features;
  const box = (m: { r1: number; c1: number; r2: number; c2: number }) => `${m.r1},${m.c1},${m.r2},${m.c2}`;
  if (fa.merges.map(box).sort().join(';') !== fb.merges.map(box).sort().join(';')) add('merges', `${fa.merges.length} != ${fb.merges.length}`);
  if (JSON.stringify(fa.freeze ?? null) !== JSON.stringify(fb.freeze ?? null)) add('freeze', `${JSON.stringify(fa.freeze)} != ${JSON.stringify(fb.freeze)}`);
  if ((fa.autoFilter ?? '') !== (fb.autoFilter ?? '')) add('autoFilter', `${fa.autoFilter} != ${fb.autoFilter}`);
  const rule = (x: SheetGrid['features']['rules'][number]) => JSON.stringify({ s: x.sqref, t: x.type, k: x.kind, e: !!x.ext, f: x.formulas.map((f) => renameRefs(f, o.rename)) });
  const ra = fa.rules.map(rule).sort(), rb = fb.rules.map((x) => JSON.stringify({ s: x.sqref, t: x.type, k: x.kind, e: !!x.ext, f: x.formulas })).sort();
  if (ra.join('\n') !== rb.join('\n')) add('rules', `${ra.length} != ${rb.length}`);
  const hl = (h: { ref: string; target?: string; location?: string }) => JSON.stringify({ r: h.ref, t: h.target, l: h.location === undefined ? undefined : renameRefs(h.location, o.rename) });
  if (fa.hyperlinks.map(hl).sort().join('\n') !== fb.hyperlinks.map((h) => JSON.stringify({ r: h.ref, t: h.target, l: h.location })).sort().join('\n')) add('hyperlinks', `${JSON.stringify(fa.hyperlinks)} != ${JSON.stringify(fb.hyperlinks)}`);
  return diffs;
}

function sameValue(a: Cell['v'], b: Cell['v']): boolean {
  if (a === b) return true;
  if (typeof a === 'number' && typeof b === 'number') return Math.abs(a - b) <= 1e-9 * Math.max(1, Math.abs(a));
  return false;
}

export function colName(n: number): string {
  let s = '';
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

/** the defined names of a package: "name|scope sheet name or ''" -> text */
export async function definedNamesOf(bytes: Uint8Array): Promise<Map<string, string>> {
  const f = await readFiles(bytes);
  const sheets = sheetsOf(f);
  const out = new Map<string, string>();
  for (const m of (f.get('xl/workbook.xml') ?? '').matchAll(/<definedName\b([^>]*?)(?:\/>|>([\s\S]*?)<\/definedName>)/g)) {
    const ls = attr(m[1], 'localSheetId');
    out.set(`${unq(attr(m[1], 'name') ?? '')}|${ls === undefined ? '' : sheets[Number(ls)]?.name ?? `#${ls}`}`, unq(m[2] ?? ''));
  }
  return out;
}
