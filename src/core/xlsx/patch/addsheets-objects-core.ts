/**
 * Shared machinery of the object phase of the sheet copier (task A2; see addsheets-objects.ts for the overview): the per-workbook state,
 * relationship reading / writing, unique part names, content types, and a staging area so that an object that cannot be copied leaves
 * nothing behind (no orphan part, no dangling relationship).
 */
import type { CheckItem } from '../../types';
import { convertNamespaces, isStrict } from './bookinfo';
import type { Engine } from './engine';
import { PatchError, dirOf, isRel, relativeTarget, relsPathOf, resolveTarget, type Rel } from './pkg';
import { xmlProblem } from './wellformed';
import { escAttr, findAll, findFirst, getAttr } from './xml';

const DML = 'http://schemas.openxmlformats.org/drawingml/2006/';
const DML_STRICT = 'http://purl.oclc.org/ooxml/drawingml/';

/** the URI of a DrawingML namespace in the flavour of the package ("main", "chart", "spreadsheetDrawing" ...) */
export const dmlUri = (name: string, strict: boolean): string => (strict ? DML_STRICT : DML) + name;

/**
 * An XML part of a package written in the Strict OOXML namespaces (purl.oclc.org: SpreadsheetML, relationships) read for a package written
 * in the transitional ones, or the other way round; the same flavour comes back unchanged. Only for SpreadsheetML parts (comments, tables,
 * pivot tables and caches): the values of DrawingML parts differ as well (percentages, angles), so drawings and charts are not converted — they
 * are copied between packages of the same flavour only (addsheets-objects.ts).
 */
export function partFlavour(xml: string, srcStrict: boolean, tgtStrict: boolean): string {
  return srcStrict === tgtStrict ? xml : convertNamespaces(xml, tgtStrict);
}

export const REL_BASE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
/** relationship types of the Microsoft extensions the copier knows */
export const REL_MS = {
  chartEx: 'http://schemas.microsoft.com/office/2014/relationships/chartEx',
  chartStyle: 'http://schemas.microsoft.com/office/2011/relationships/chartStyle',
  chartColors: 'http://schemas.microsoft.com/office/2011/relationships/chartColorStyle',
  threadedComment: 'http://schemas.microsoft.com/office/2017/10/relationships/threadedComment',
  person: 'http://schemas.microsoft.com/office/2017/10/relationships/person',
};

export const CT = {
  drawing: 'application/vnd.openxmlformats-officedocument.drawing+xml',
  chart: 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
  chartEx: 'application/vnd.ms-office.chartex+xml',
  chartStyle: 'application/vnd.ms-office.chartstyle+xml',
  chartColors: 'application/vnd.ms-office.chartcolorstyle+xml',
  chartShapes: 'application/vnd.openxmlformats-officedocument.drawingml.chartshapes+xml',
  themeOverride: 'application/vnd.openxmlformats-officedocument.themeOverride+xml',
  table: 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml',
  comments: 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml',
  vml: 'application/vnd.openxmlformats-officedocument.vmlDrawing',
  pivotTable: 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml',
  pivotCacheDef: 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml',
  pivotCacheRec: 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml',
  chartsheet: 'application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml',
};

/** a relationship of a part of the SOURCE package */
export interface SrcRel extends Rel {
  /** the resolved part path ('' for an external target and for an in-document link such as "#Sheet1!A1") */
  path: string;
  /** an in-document link ("#Sheet1!A1") written as a relationship (hyperlinks of shapes) */
  fragment: boolean;
}

export async function readSrcRels(eng: Engine, owner: string): Promise<SrcRel[]> {
  const dir = dirOf(owner);
  return (await eng.pkg.rels(owner)).map((r) => {
    const fragment = !r.external && r.target.startsWith('#');
    return { ...r, fragment, path: r.external || fragment ? '' : resolveTarget(dir, r.target) };
  });
}

/** a relationship of a part that is created in the target package */
export interface OutRel {
  id: string;
  type: string;
  /** a part path ("xl/media/image3.png") or, with `external` / `raw`, the target text */
  target: string;
  external?: boolean;
  /** the target is written as it is (an in-document hyperlink target "#Sheet1!A1") */
  raw?: boolean;
}

export function relsXml(owner: string, rels: readonly OutRel[]): string {
  const dir = dirOf(owner);
  const body = rels.map((r) => (r.external
    ? `<Relationship Id="${r.id}" Type="${escAttr(r.type)}" Target="${escAttr(r.target)}" TargetMode="External"/>`
    : `<Relationship Id="${r.id}" Type="${escAttr(r.type)}" Target="${escAttr(r.raw ? r.target : relativeTarget(dir, r.target))}"/>`)).join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${body}</Relationships>`;
}

// ---------------------------------------------------------------------------------------------------------------------
// per-workbook state
// ---------------------------------------------------------------------------------------------------------------------

export interface PivotCacheCopy {
  /** cacheId in the target workbook */
  id: number;
  /** the part of the copied definition */
  path: string;
}

/** What the object phase has done to ONE target workbook so far (every sheet of every source shares it). */
export interface ObjState {
  eng: Engine;
  /** the target is written in the Strict OOXML namespaces */
  strict: boolean;
  /** part paths handed out but not written yet (staged) */
  reserved: Set<string>;
  next: Map<string, number>;
  /** highest table id of the workbook */
  tableId: number;
  /** highest cacheId of <pivotCaches> */
  cacheId: number;
  /** lower-cased names of the pivot tables of the workbook */
  pivotNames: Set<string>;
  /** the highest <o:idmap data> of the VML parts of the workbook */
  vmlMap: number;
  /** source engine -> (lower-cased source part path -> new part path) of the parts that several objects share (images, embedded workbooks) */
  shared: Map<Engine, Map<string, string>>;
  /** source engine -> (lower-cased pivot cache definition path -> what became of it; null = not usable) */
  caches: Map<Engine, Map<string, PivotCacheCopy | null>>;
  /** source engine -> (source custom table style -> its name in the target) */
  tableStyles: Map<Engine, Map<string, string>>;
  /** source engine -> cached content types of its package */
  srcTypes: Map<Engine, { defaults: Map<string, string>; overrides: Map<string, string> }>;
}

const STATES = new WeakMap<Engine, ObjState>();

export async function getState(eng: Engine): Promise<ObjState> {
  const ex = STATES.get(eng);
  if (ex) return ex;
  const pkg = eng.pkg;
  const st: ObjState = {
    eng, strict: isStrict(eng.wbXml), reserved: new Set(), next: new Map(), tableId: 0, cacheId: 0, pivotNames: new Set(), vmlMap: 0,
    shared: new Map(), caches: new Map(), tableStyles: new Map(), srcTypes: new Map(),
  };
  for (const p of pkg.list()) {
    if (/^xl\/tables\/[^/]+\.xml$/i.test(p)) {
      const x = await pkg.read(p);
      const id = Number(getAttr(/<table\b[^>]*>/.exec(x ?? '')?.[0] ?? '', 'id'));
      if (Number.isFinite(id)) st.tableId = Math.max(st.tableId, id);
    } else if (/^xl\/pivotTables\/[^/]+\.xml$/i.test(p)) {
      const x = await pkg.read(p);
      const n = getAttr(/<pivotTableDefinition\b[^>]*>/.exec(x ?? '')?.[0] ?? '', 'name');
      if (n) st.pivotNames.add(n.toLowerCase());
    } else if (/\.vml$/i.test(p)) {
      const x = await pkg.read(p);
      for (const m of (x ?? '').matchAll(/<o:idmap\b[^>]*\bdata="([^"]*)"/g)) for (const d of m[1].split(',')) st.vmlMap = Math.max(st.vmlMap, Number(d) || 0);
    }
  }
  for (const m of eng.wbXml.matchAll(/<pivotCache\b[^>]*>/g)) st.cacheId = Math.max(st.cacheId, Number(getAttr(m[0], 'cacheId')) || 0);
  STATES.set(eng, st);
  return st;
}

/** the lower-cased names of the tables and defined names of a workbook (a new table must not take one of them) */
export async function collectTakenNames(eng: Engine): Promise<Set<string>> {
  const out = new Set<string>();
  for (const p of eng.pkg.list()) {
    if (!/^xl\/tables\/[^/]+\.xml$/i.test(p)) continue;
    const open = /<table\b[^>]*>/.exec((await eng.pkg.read(p)) ?? '')?.[0] ?? '';
    for (const a of ['name', 'displayName']) { const v = getAttr(open, a); if (v) out.add(v.toLowerCase()); }
  }
  const dn = findFirst(eng.wbXml, 'definedNames');
  if (dn && !dn.selfClosing) for (const d of findAll(dn.inner, 'definedName')) { const n = getAttr(d.open, 'name'); if (n) out.add(n.toLowerCase()); }
  return out;
}

/** a part name that is free in the target: <dir><base><n><ext> */
export function freshPart(st: ObjState, dir: string, base: string, ext: string): string {
  const key = dir + base + ext;
  let n = st.next.get(key) ?? 1;
  for (;; n++) {
    const p = `${dir}${base}${n}${ext}`;
    if (!st.reserved.has(p.toLowerCase()) && !st.eng.pkg.has(p)) {
      st.next.set(key, n + 1);
      st.reserved.add(p.toLowerCase());
      return p;
    }
  }
}

/** the relationship type URL of a standard kind ("drawing", "chart" ...) in the flavour of the target */
export function relType(st: ObjState, kind: string): string {
  // (the VML of the notes keeps the transitional type in a Strict package, as Excel writes it)
  return kind === 'vmlDrawing' ? `${REL_BASE}/${kind}` : convertNamespaces(`${REL_BASE}/${kind}`, st.strict);
}

// ---------------------------------------------------------------------------------------------------------------------
// content types of the source
// ---------------------------------------------------------------------------------------------------------------------

export async function srcContentType(st: ObjState, src: Engine, path: string): Promise<string | undefined> {
  let t = st.srcTypes.get(src);
  if (!t) {
    const x = (await src.pkg.read('[Content_Types].xml')) ?? '';
    t = { defaults: new Map(), overrides: new Map() };
    for (const m of x.matchAll(/<Default\b[^>]*>/g)) t.defaults.set((getAttr(m[0], 'Extension') ?? '').toLowerCase(), getAttr(m[0], 'ContentType') ?? '');
    for (const m of x.matchAll(/<Override\b[^>]*>/g)) t.overrides.set((getAttr(m[0], 'PartName') ?? '').replace(/^\//, '').toLowerCase(), getAttr(m[0], 'ContentType') ?? '');
    st.srcTypes.set(src, t);
  }
  const ext = /\.([^./]+)$/.exec(path)?.[1]?.toLowerCase() ?? '';
  return t.overrides.get(path.toLowerCase()) ?? t.defaults.get(ext) ?? GUESS[ext];
}

/** the types of the parts a package of another writer may leave without a content type (the package is invalid without one) */
const GUESS: Record<string, string> = {
  png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', tif: 'image/tiff', tiff: 'image/tiff', emf: 'image/x-emf', wmf: 'image/x-wmf', svg: 'image/svg+xml',
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', xlsm: 'application/vnd.ms-excel.sheet.macroEnabled.12', bin: 'application/vnd.openxmlformats-officedocument.oleObject',
};

/** <Default Extension="ext"/> in [Content_Types].xml (nothing when the extension has one) */
async function ensureDefault(eng: Engine, ext: string, type: string): Promise<void> {
  const x = (await eng.pkg.read('[Content_Types].xml')) ?? '';
  if (new RegExp(`<Default\\b[^>]*\\bExtension="${ext.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"`, 'i').test(x)) return;
  const el = `<Default Extension="${escAttr(ext)}" ContentType="${escAttr(type)}"/>`;
  const last = [...x.matchAll(/<Default\b[^>]*\/>/g)].pop();
  if (last) eng.pkg.write('[Content_Types].xml', x.slice(0, last.index + last[0].length) + el + x.slice(last.index + last[0].length));
  else eng.pkg.write('[Content_Types].xml', x.replace(/(<Types\b[^>]*>)/, `$1${el}`));
}

// ---------------------------------------------------------------------------------------------------------------------
// staging
// ---------------------------------------------------------------------------------------------------------------------

/**
 * What one object (a drawing with its charts and pictures, a table, a pivot table ...) wants to add to the target. Nothing is written until
 * commit(): an object that fails half way is dropped without leaving a part, a relationship or a content type behind.
 */
export class Stage {
  private texts = new Map<string, string>();
  private bytes = new Map<string, Uint8Array>();
  private overrides = new Map<string, string>();
  private defaults = new Map<string, string>();
  private rels = new Map<string, OutRel[]>();
  /** work to do at commit (workbook-level registrations) */
  private after: Array<() => Promise<void>> = [];

  constructor(readonly st: ObjState) {}

  text(path: string, xml: string, contentType?: string): void {
    // a part that was damaged in the source is damaged in the copy: the object is left out instead of the whole output being refused at the end
    if (/\.(xml|rels)$/i.test(path)) {
      const why = xmlProblem(xml);
      if (why) throw new PatchError(`図・表などの部品（${path}）が壊れています: ${why}`);
    }
    this.texts.set(path, xml);
    if (contentType) this.overrides.set(path, contentType);
  }

  /** a binary part whose content type is an Override of its own (parts of XML types copied byte by byte) */
  binaryAs(path: string, data: Uint8Array, contentType: string): void {
    // an XML part copied byte by byte is checked like the ones written as text (a damaged part must not reach the output)
    if (/\.xml$/i.test(path) && data.length <= 64 * 1024 * 1024) {
      const why = xmlProblem(new TextDecoder('utf-8').decode(data));
      if (why) throw new PatchError(`図・表などの部品（${path}）が壊れています: ${why}`);
    }
    this.bytes.set(path, data);
    this.overrides.set(path, contentType);
  }

  /** the content type every part of this extension gets (<Default Extension=…>) */
  defaultType(ext: string, contentType: string): void {
    this.defaults.set(ext.toLowerCase(), contentType);
  }

  binary(path: string, data: Uint8Array, contentType: string | undefined): void {
    this.bytes.set(path, data);
    const ext = /\.([^./]+)$/.exec(path)?.[1]?.toLowerCase();
    if (contentType && ext) this.defaults.set(ext, contentType);
  }

  setRels(owner: string, rels: readonly OutRel[]): void {
    if (rels.length) this.rels.set(owner, [...rels]);
  }

  /** run `fn` when the stage is committed (after every part is written) */
  onCommit(fn: () => Promise<void>): void {
    this.after.push(fn);
  }

  /** shared parts this stage has registered (forgotten again by discard()) */
  private sharedAdded: Array<{ map: Map<string, string>; key: string }> = [];

  /** Copy a part of the source package byte by byte; returns the new path (shared parts are copied once per source). */
  async copyBinary(src: Engine, srcPath: string, dir: string, base: string, ext: string, share: boolean): Promise<string | null> {
    const key = srcPath.toLowerCase();
    let map = this.st.shared.get(src);
    if (share) {
      const hit = map?.get(key);
      if (hit) return hit;
    }
    const data = await src.pkg.readBytes(srcPath);
    if (!data) return null;
    const path = freshPart(this.st, dir, base, ext);
    this.binary(path, data, await srcContentType(this.st, src, srcPath));
    if (share) {
      if (!map) this.st.shared.set(src, (map = new Map()));
      map.set(key, path);
      this.sharedAdded.push({ map, key });
    }
    return path;
  }

  /** take over everything another (sub-)stage holds */
  adopt(o: Stage): void {
    for (const [k, v] of o.texts) this.texts.set(k, v);
    for (const [k, v] of o.bytes) this.bytes.set(k, v);
    for (const [k, v] of o.overrides) this.overrides.set(k, v);
    for (const [k, v] of o.defaults) this.defaults.set(k, v);
    for (const [k, v] of o.rels) this.rels.set(k, v);
    this.after.push(...o.after);
    this.sharedAdded.push(...o.sharedAdded);
    o.texts.clear(); o.bytes.clear(); o.overrides.clear(); o.defaults.clear(); o.rels.clear();
    o.after = [];
    o.sharedAdded = [];
  }

  /** the object failed: nothing of it is written, and the parts it shared with later objects are forgotten */
  discard(): void {
    for (const s of this.sharedAdded) s.map.delete(s.key);
    this.sharedAdded = [];
    this.texts.clear();
    this.bytes.clear();
    this.overrides.clear();
    this.defaults.clear();
    this.rels.clear();
    this.after = [];
  }

  async commit(): Promise<void> {
    const eng = this.st.eng;
    const pkg = eng.pkg;
    for (const [p, x] of this.texts) pkg.write(p, x);
    for (const [p, b] of this.bytes) pkg.zip.file(p, b);
    for (const [p, ct] of this.overrides) await pkg.setContentTypeOverride(p, ct);
    for (const [ext, ct] of this.defaults) await ensureDefault(eng, ext, ct);
    for (const [owner, list] of this.rels) pkg.write(relsPathOf(owner), relsXml(owner, list));
    for (const f of this.after) await f();
    this.sharedAdded = [];
  }
}

/** true when the source relationship is of the given standard kind (transitional or Strict spelling) */
export const relIs = (r: { type: string }, kind: string): boolean => isRel(r.type, kind);

/** String.replace with an async replacer (the replacements are computed in order, then the string is rebuilt) */
export async function replaceAsync(s: string, re: RegExp, fn: (m: RegExpMatchArray) => Promise<string>): Promise<string> {
  const ms = [...s.matchAll(re)];
  if (!ms.length) return s;
  let out = '';
  let pos = 0;
  for (const m of ms) {
    const at = m.index ?? 0;
    out += s.slice(pos, at) + (await fn(m));
    pos = at + m[0].length;
  }
  return out + s.slice(pos);
}

// ---------------------------------------------------------------------------------------------------------------------
// information for the user that is no limit of the copy (a table was renamed ...)
// ---------------------------------------------------------------------------------------------------------------------

const INFO = new WeakMap<Engine, CheckItem[]>();

/** an info item for the merge report ('tableRenamed' ...): it does not make the sheet "partly copied" */
export function addInfo(eng: Engine, item: CheckItem): void {
  const l = INFO.get(eng);
  if (l) l.push(item);
  else INFO.set(eng, [item]);
}

/** the info items added since the last call (addsheets.ts puts them into the report) */
export function takeInfo(eng: Engine): CheckItem[] {
  const l = INFO.get(eng) ?? [];
  INFO.delete(eng);
  return l;
}
