/**
 * Helpers of the object-phase tests of the sheet copier (task A2; objects = drawings, charts, pictures, shapes, notes, tables, pivot tables):
 *  - objectsOf: what every sheet of a package owns (counts and names), read from the package parts (not through the app's reader), so that a
 *    copy can be compared with its source;
 *  - lintObjects: addsheets-helpers' lintPackage (A1) plus the rules of the object parts: every r:id / r:embed of every part is a relationship
 *    of it, drawings / charts / tables / pivot tables / comments are owned by exactly one owner, ids are unique (table ids, table names,
 *    pivot cache ids, pivot table names, shape ids of a drawing, VML ids), the workbook lists every pivot cache, chart formulas read sheets that
 *    exist, tables lie on cells of their sheet header strings, hyperlink targets "#Sheet!A1" are not parts.
 */
import JSZip from 'jszip';
import { lintPackage, type Lint } from './addsheets-helpers';

export interface SheetObjects {
  name: string;
  part: string;
  kind: 'worksheet' | 'chartsheet';
  charts: number;
  pictures: number;
  /** <xdr:sp> that are not placeholders of controls / slicers */
  shapes: number;
  connectors: number;
  groups: number;
  slicers: number;
  /** anchors that are placeholders of form controls / ActiveX / OLE objects */
  controlPlaceholders: number;
  notes: number;
  threaded: number;
  /** VML shapes that are not notes */
  vmlControls: number;
  tables: Array<{ name: string; ref: string; style?: string; totals: number; columns: string[] }>;
  pivots: Array<{ name: string; cacheId: string; source: string }>;
  background: boolean;
  /** formulas of the charts (<c:f> text, in document order) */
  chartFormulas: string[];
  /** the cached values of the charts (<c:v> text, in document order) */
  chartValues: string[];
  /** per chart (anchor order): a series has no colour of its own (it takes the colour of the workbook's theme) */
  chartAuto: boolean[];
}

const attr = (tag: string, name: string): string | undefined => new RegExp(`\\s${name}\\s*=\\s*"([^"]*)"`).exec(tag)?.[1];
const unq = (s: string): string => s.replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&');
const dirName = (p: string): string => (p.includes('/') ? p.slice(0, p.lastIndexOf('/')) : '');
function resolve(owner: string, target: string): string {
  let t = target;
  try { t = decodeURIComponent(t); } catch { /* keep */ }
  if (t.startsWith('/')) return t.slice(1);
  const parts = dirName(owner) ? dirName(owner).split('/') : [];
  for (const seg of t.split('/')) { if (seg === '..') parts.pop(); else if (seg && seg !== '.') parts.push(seg); }
  return parts.join('/');
}

export interface Pkg {
  files: Map<string, string>;
  names: string[];
  zip: JSZip;
}

export async function openPkg(bytes: Uint8Array): Promise<Pkg> {
  const zip = await JSZip.loadAsync(bytes);
  const files = new Map<string, string>();
  const names: string[] = [];
  for (const n of Object.keys(zip.files)) {
    if (zip.files[n].dir) continue;
    names.push(n);
    if (/\.(xml|rels|vml)$/i.test(n)) files.set(n, await zip.files[n].async('string'));
  }
  return { files, names, zip };
}

export interface RelInfo {
  id: string;
  type: string;
  target: string;
  external: boolean;
  path: string;
}

export function relsOf(p: Pkg, owner: string): RelInfo[] {
  const i = owner.lastIndexOf('/');
  const x = p.files.get(`${owner.slice(0, i + 1)}_rels/${owner.slice(i + 1)}.rels`);
  if (x === undefined) return [];
  return [...x.matchAll(/<Relationship\b[^>]*>/g)].map((m) => {
    const target = unq(attr(m[0], 'Target') ?? '');
    const external = /TargetMode\s*=\s*"External"/.test(m[0]);
    return { id: attr(m[0], 'Id') ?? '', type: attr(m[0], 'Type') ?? '', target, external, path: external || target.startsWith('#') ? '' : resolve(owner, target) };
  });
}

const kindOf = (r: { type: string }, kind: string): boolean => r.type.endsWith('/' + kind);
const count = (x: string, re: RegExp): number => (x.match(re) ?? []).length;

export async function objectsOf(bytes: Uint8Array): Promise<SheetObjects[]> {
  const p = await openPkg(bytes);
  const wb = p.files.get('xl/workbook.xml') ?? '';
  const wbRels = relsOf(p, 'xl/workbook.xml');
  const persons = new Map<string, string>();
  void persons;
  const out: SheetObjects[] = [];
  for (const m of wb.matchAll(/<sheet\b[^>]*>/g)) {
    const name = unq(attr(m[0], 'name') ?? '');
    const rel = wbRels.find((r) => r.id === attr(m[0], 'r:id'));
    if (!rel || !p.files.has(rel.path)) continue;
    const kind = kindOf(rel, 'chartsheet') ? 'chartsheet' : kindOf(rel, 'worksheet') ? 'worksheet' : null;
    if (!kind) continue;
    const x = p.files.get(rel.path)!;
    const rels = relsOf(p, rel.path);
    const so: SheetObjects = {
      name, part: rel.path, kind, charts: 0, pictures: 0, shapes: 0, connectors: 0, groups: 0, slicers: 0, controlPlaceholders: 0, notes: 0, threaded: 0,
      vmlControls: 0, tables: [], pivots: [], background: /<picture\b/.test(x), chartFormulas: [], chartValues: [], chartAuto: [],
    };
    const dr = rels.find((r) => kindOf(r, 'drawing'));
    const dx = dr ? p.files.get(dr.path) : undefined;
    if (dx && dr) {
      so.pictures = count(dx, /<(?:\w+:)?pic[\s>]/g);
      so.connectors = count(dx, /<(?:\w+:)?cxnSp[\s>]/g);
      so.groups = count(dx, /<(?:\w+:)?grpSp[\s>]/g);
      so.slicers = count(dx, /<(?:\w+:)?slicer\b|<(?:\w+:)?timeslicer\b/g);
      const placeholders = count(dx, /<a14:compatExt\b/g);
      so.controlPlaceholders = placeholders;
      so.shapes = count(dx, /<(?:\w+:)?sp[\s>]/g) - placeholders - count(dx, /<mc:Fallback>\s*<xdr:sp[\s>]/g);
      const drels = relsOf(p, dr.path);
      // in the order of the anchors
      const crs = [...dx.matchAll(/<(?:[\w.-]+:)?chart\b[^>]*\sr:id="([^"]*)"/g)].map((m) => drels.find((r) => r.id === m[1])).filter((r): r is RelInfo => !!r && (kindOf(r, 'chart') || r.type.endsWith('/chartEx')));
      for (const cr of crs) {
        so.charts++;
        const cx = p.files.get(cr.path) ?? '';
        so.chartAuto.push((cx.match(/<(?:[\w.-]+:)?ser>[\s\S]*?<\/(?:[\w.-]+:)?ser>/g) ?? []).some((s) => !/<(?:[\w.-]+:)?(?:spPr|dPt)>[\s\S]*?<(?:[\w.-]+:)?(?:solidFill|gradFill|pattFill|blipFill)\b/.test(s)));
        for (const f of cx.matchAll(/<(?:[\w.-]+:)?f(?:\s[^>]*)?>([^<]*)<\/(?:[\w.-]+:)?f>/g)) so.chartFormulas.push(unq(f[1]));
        for (const v of cx.matchAll(/<(?:[\w.-]+:)?v>([^<]*)<\/(?:[\w.-]+:)?v>/g)) so.chartValues.push(unq(v[1]));
      }
    }
    const cr = rels.find((r) => kindOf(r, 'comments'));
    if (cr) so.notes = count(p.files.get(cr.path) ?? '', /<comment\b/g);
    const tr = rels.find((r) => r.type.endsWith('/threadedComment'));
    if (tr) so.threaded = count(p.files.get(tr.path) ?? '', /<threadedComment\b(?![^>]*\sparentId=)/g);
    for (const v of rels.filter((r) => kindOf(r, 'vmlDrawing'))) so.vmlControls += count(p.files.get(v.path) ?? '', /ObjectType="(?!Note")[^"]*"/g);
    for (const t of rels.filter((r) => kindOf(r, 'table'))) {
      const tx = p.files.get(t.path) ?? '';
      const open = /<table\b[^>]*>/.exec(tx)?.[0] ?? '';
      so.tables.push({
        name: unq(attr(open, 'displayName') ?? ''), ref: attr(open, 'ref') ?? '', style: attr(/<tableStyleInfo\b[^>]*>/.exec(tx)?.[0] ?? '', 'name'),
        totals: Number(attr(open, 'totalsRowCount') ?? 0), columns: [...tx.matchAll(/<tableColumn\b[^>]*>/g)].map((c) => unq(attr(c[0], 'name') ?? '')),
      });
    }
    for (const pv of rels.filter((r) => kindOf(r, 'pivotTable'))) {
      const px = p.files.get(pv.path) ?? '';
      const open = /<pivotTableDefinition\b[^>]*>/.exec(px)?.[0] ?? '';
      const def = relsOf(p, pv.path).find((r) => kindOf(r, 'pivotCacheDefinition'));
      const ws = /<worksheetSource\b[^>]*>/.exec(p.files.get(def?.path ?? '') ?? '')?.[0] ?? '';
      so.pivots.push({ name: unq(attr(open, 'name') ?? ''), cacheId: attr(open, 'cacheId') ?? '', source: `${unq(attr(ws, 'sheet') ?? '')}!${unq(attr(ws, 'ref') ?? attr(ws, 'name') ?? '')}` });
    }
    out.push(so);
  }
  return out;
}

// ---------------------------------------------------------------------------------------------------------------------
// the object linter
// ---------------------------------------------------------------------------------------------------------------------

const R_ATTRS = /\sr:(?:id|embed|link|pict|dm|lo|qs|cs)="([^"]*)"/g;

/** lintPackage + the rules of the object parts; the in-document hyperlink targets of shapes ("#Sheet1!A1") are not parts */
export async function lintObjects(bytes: Uint8Array, opts: { worksheets?: number } = {}): Promise<Lint[]> {
  // (the sheets of a base file are not touched: its own mc:AlternateContent elements, e.g. form controls, are not the copier's)
  const out = (await lintPackage(bytes, opts)).filter((l) => !(l.code === 'REL_DANGLING' && /->\s.*#/.test(l.detail)) && !(l.code === 'WS_UNKNOWN_ELEMENT' && /<mc:AlternateContent>/.test(l.detail)));
  const p = await openPkg(bytes);
  const add = (code: string, detail: string) => out.push({ code, detail });
  // every relationship id used inside a part exists in its relationships
  for (const [path, x] of p.files) {
    if (/\.rels$/i.test(path) || /\[Content_Types\]/.test(path) || /\.vml$/i.test(path)) continue;
    const rels = relsOf(p, path);
    for (const m of x.matchAll(R_ATTRS)) if (!rels.some((r) => r.id === m[1])) add('OBJ_RID_DANGLING', `${path}: ${m[1]}`);
  }
  // owners
  const owners = new Map<string, string[]>();
  for (const [path] of p.files) {
    if (!/\.rels$/i.test(path)) continue;
    const owner = path.replace(/(^|\/)_rels\/([^/]*)\.rels$/i, (_m, d: string, n: string) => d + n);
    for (const r of relsOf(p, owner)) {
      if (!r.path || !/\/(drawing|chart|chartEx|table|pivotTable|pivotCacheDefinition|pivotCacheRecords|comments|vmlDrawing|chartUserShapes|threadedComment)$/i.test(r.type)) continue;
      const l = owners.get(r.path.toLowerCase()) ?? [];
      l.push(`${owner}:${r.id}`);
      owners.set(r.path.toLowerCase(), l);
    }
  }
  for (const [part, list] of owners) {
    // a pivot cache is shared by its pivot tables and the workbook; every other part has exactly one owner
    if (/pivotCache/i.test(part)) continue;
    if (list.length !== 1) add('OBJ_OWNERS', `${part}: ${list.join(', ')}`);
  }
  // tables
  const tableIds = new Set<string>(), tableNames = new Set<string>();
  const definedNames = new Set([...(p.files.get('xl/workbook.xml') ?? '').matchAll(/<definedName\b[^>]*\sname="([^"]*)"/g)].map((m) => m[1].toLowerCase()));
  for (const [path, x] of p.files) {
    if (!/^xl\/tables\/[^/]+\.xml$/i.test(path)) continue;
    const open = /<table\b[^>]*>/.exec(x)?.[0] ?? '';
    const id = attr(open, 'id') ?? '', nm = (attr(open, 'displayName') ?? '').toLowerCase(), nm2 = (attr(open, 'name') ?? '').toLowerCase();
    if (tableIds.has(id)) add('TABLE_ID_DUPLICATE', `${path}: ${id}`);
    tableIds.add(id);
    for (const n of new Set([nm, nm2])) {
      if (tableNames.has(n)) add('TABLE_NAME_DUPLICATE', `${path}: ${n}`);
      tableNames.add(n);
      if (definedNames.has(n)) add('TABLE_NAME_IS_NAME', `${path}: ${n}`);
    }
    const stylesX = p.files.get('xl/styles.xml') ?? '';
    const nDxfT = count(/<dxfs\b[^>]*?(?:\/>|>[\s\S]*?<\/dxfs>)/.exec(stylesX)?.[0] ?? '', /<dxf\b/g);
    for (const m of x.matchAll(/\s(?:data|headerRow|totalsRow|headerRowBorder|tableBorder|totalsRowBorder)DxfId="(\d+)"/g)) if (Number(m[1]) >= nDxfT) add('TABLE_DXF', `${path}: ${m[1]}`);
    const colIds = [...x.matchAll(/<tableColumn\b[^>]*\sid="([^"]*)"/g)].map((m) => m[1]);
    if (new Set(colIds).size !== colIds.length) add('TABLE_COLUMN_ID', path);
    const declared = Number(attr(/<tableColumns\b[^>]*>/.exec(x)?.[0] ?? '', 'count'));
    if (declared !== colIds.length) add('TABLE_COLUMN_COUNT', `${path}: ${declared} vs ${colIds.length}`);
    if (/\sxr\d*:uid="/.test(x) && new Set([...x.matchAll(/\sxr:uid="([^"]*)"/g)].map((m) => m[1])).size < (x.match(/\sxr:uid="/g) ?? []).length - 1) add('TABLE_UID', path);
  }
  // pivot caches
  const wb = p.files.get('xl/workbook.xml') ?? '';
  const wbRels = relsOf(p, 'xl/workbook.xml');
  const caches = [...wb.matchAll(/<pivotCache\b[^>]*>/g)].map((m) => ({ id: attr(m[0], 'cacheId') ?? '', rid: attr(m[0], 'r:id') ?? '' }));
  if (new Set(caches.map((c) => c.id)).size !== caches.length) add('PIVOT_CACHEID_DUPLICATE', caches.map((c) => c.id).join(','));
  for (const c of caches) if (!wbRels.some((r) => r.id === c.rid && kindOf(r, 'pivotCacheDefinition'))) add('PIVOT_CACHE_REL', `${c.id}: ${c.rid}`);
  const pivotNames = new Set<string>();
  for (const [path, x] of p.files) {
    if (!/^xl\/pivotTables\/[^/]+\.xml$/i.test(path)) continue;
    const open = /<pivotTableDefinition\b[^>]*>/.exec(x)?.[0] ?? '';
    const id = attr(open, 'cacheId') ?? '';
    if (!caches.some((c) => c.id === id)) add('PIVOT_CACHEID_MISSING', `${path}: cacheId ${id}`);
    const n = (attr(open, 'name') ?? '').toLowerCase();
    if (pivotNames.has(n)) add('PIVOT_NAME_DUPLICATE', `${path}: ${n}`);
    pivotNames.add(n);
    const def = relsOf(p, path).find((r) => kindOf(r, 'pivotCacheDefinition'));
    const c = caches.find((q) => q.id === id);
    if (def && c && wbRels.find((r) => r.id === c.rid)?.path !== def.path) add('PIVOT_CACHE_MISMATCH', `${path}: cacheId ${id} is ${wbRels.find((r) => r.id === c.rid)?.path}, the part says ${def.path}`);
    // dxf / numFmt ids in range
    const styles = p.files.get('xl/styles.xml') ?? '';
    const nDxf = count(/<dxfs\b[^>]*?(?:\/>|>[\s\S]*?<\/dxfs>)/.exec(styles)?.[0] ?? '', /<dxf\b/g);
    for (const m of x.matchAll(/\sdxfId="(\d+)"/g)) if (Number(m[1]) >= nDxf) add('PIVOT_DXF', `${path}: ${m[1]}`);
    const fmts = new Set([...styles.matchAll(/<numFmt\b[^>]*\snumFmtId="(\d+)"/g)].map((m) => Number(m[1])));
    for (const m of x.matchAll(/\snumFmtId="(\d+)"/g)) if (Number(m[1]) >= 164 && !fmts.has(Number(m[1]))) add('PIVOT_NUMFMT', `${path}: ${m[1]}`);
  }
  // charts: formulas read sheets that exist
  const sheetNames = new Set([...wb.matchAll(/<sheet\b[^>]*>/g)].map((m) => unq(attr(m[0], 'name') ?? '').toLowerCase()));
  for (const [path, x] of p.files) {
    if (!/^xl\/charts\/chart[^/]*\.xml$/i.test(path)) continue;
    for (const f of x.matchAll(/<(?:[\w.-]+:)?f(?:\s[^>]*)?>([^<]*)<\/(?:[\w.-]+:)?f>/g)) {
      for (const s of unq(f[1]).matchAll(/(?:'((?:[^']|'')+)'|([^\s'!(),:]+))!/g)) {
        const nm = (s[1] ?? s[2]).replace(/''/g, "'").toLowerCase();
        if (!sheetNames.has(nm)) add('CHART_SHEET_MISSING', `${path}: ${nm}`);
      }
    }
  }
  // drawings: ids unique, comments VML ids unique across the workbook
  for (const [path, x] of p.files) {
    if (!/^xl\/drawings\/drawing[^/]*\.xml$/i.test(path)) continue;
    const ids = [...x.matchAll(/<(?:\w+:)?cNvPr\b[^>]*\sid="(\d+)"/g)].map((m) => m[1]);
    const dup = ids.filter((v, i) => ids.indexOf(v) !== i);
    // the Fallback shapes of a slicer / control have id 0
    if (dup.filter((v) => v !== '0').length) add('DRAWING_SHAPE_ID', `${path}: ${dup.join(',')}`);
  }
  const vmlIds = new Map<string, string>();
  for (const [path, x] of p.files) {
    if (!/\.vml$/i.test(path)) continue;
    for (const m of x.matchAll(/<v:shape\b[^>]*\sid="([^"]*)"/g)) {
      if (vmlIds.has(m[1])) add('VML_ID_DUPLICATE', `${path} and ${vmlIds.get(m[1])}: ${m[1]}`);
      vmlIds.set(m[1], path);
    }
  }
  return out;
}

export const lintObjectsText = (l: Lint[]): string => l.map((x) => `${x.code}: ${x.detail}`).join('\n');
