/**
 * Charts of a copied sheet (task A2, see addsheets-objects.ts): one chart part with the parts it owns (chart style / colours, a theme
 * override, an embedded workbook, user shapes, picture fills).
 *
 * Every formula of the chart (<c:f> of the series values, categories, names, the title, error bars, data-label ranges ...) is re-pointed at
 * the renamed copies of the source's sheets, token by token (the formula copier of the sheet; the same text rules as for cells). A series
 * whose reference cannot follow (a defined name of the workbook, an external workbook, a sheet that is not copied) keeps its cached values
 * as literal data instead of losing its data: the chart looks the same, it just no longer follows the cells.
 */
import { esc, findAll, findFirst, getAttr, nsPrefix, unesc } from './xml';
import { tokenize } from '../../formula/excelRef';
import { themeFont, themeXml } from './bookinfo';
import { PatchError } from './pkg';
import type { Rewrite } from './copyformula';
import type { SheetObjectsContext } from './copytypes';
import {
  CT, REL_MS, dmlUri, freshPart, readSrcRels, relIs, relType, replaceAsync,
  Stage, type ObjState, type OutRel, type SrcRel,
} from './addsheets-objects-core';

const SCHEME_INDEX: Record<string, number> = {
  lt1: 0, bg1: 0, dk1: 1, tx1: 1, lt2: 2, bg2: 2, dk2: 3, tx2: 3,
  accent1: 4, accent2: 5, accent3: 6, accent4: 7, accent5: 8, accent6: 9, hlink: 10, folHlink: 11,
};

// ---------------------------------------------------------------------------------------------------------------------
// theme: the colours / fonts a drawing or chart takes from the theme are the SOURCE theme's
// ---------------------------------------------------------------------------------------------------------------------

export interface ThemeConv {
  /** the source theme's colours in theme-index order */
  rgb: string[];
  minor?: string;
  major?: string;
}

const reEsc = (s: string): string => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

/** a drawing / chart of a source whose theme differs from the target's: theme colours become RGB, theme fonts become the fonts' names */
export function convertTheme(xml: string, tc: ThemeConv | null, dml: string): string {
  if (!tc) return xml;
  let x = xml;
  if (x.indexOf('schemeClr') >= 0) {
    const p = reEsc(dml);
    const colour = (name: string): string | undefined => {
      const i = SCHEME_INDEX[name];
      return i === undefined ? undefined : tc.rgb[i];
    };
    x = x.replace(new RegExp(`<${p}schemeClr val="(\\w+)"([^>]*?)(?<!/)>((?:(?!</${p}schemeClr>)[\\s\\S])*)</${p}schemeClr>`, 'g'), (m0, name: string, rest: string, inner: string) => {
      const c = colour(name);
      return c === undefined ? m0 : `<${dml}srgbClr val="${c}"${rest}>${inner}</${dml}srgbClr>`;
    });
    x = x.replace(new RegExp(`<${p}schemeClr val="(\\w+)"([^>]*?)/>`, 'g'), (m0, name: string, rest: string) => {
      const c = colour(name);
      return c === undefined ? m0 : `<${dml}srgbClr val="${c}"${rest}/>`;
    });
  }
  if (x.indexOf('typeface="+m') >= 0) {
    const face = (s: string) => esc(s).replace(/"/g, '&quot;');
    if (tc.minor) x = x.replace(/typeface="\+mn-(?:lt|ea|cs)"/g, `typeface="${face(tc.minor)}"`);
    if (tc.major) x = x.replace(/typeface="\+mj-(?:lt|ea|cs)"/g, `typeface="${face(tc.major)}"`);
  }
  return x;
}

/** null when the source and the target theme agree */
export async function themeConversion(ctx: SheetObjectsContext): Promise<ThemeConv | null> {
  const rgb = ctx.maps.themeRgb;
  if (!rgb) return null;
  const src = await themeXml(ctx.source.eng);
  const tgt = await themeXml(ctx.target.eng);
  const minor = themeFont(src, 'minor'), major = themeFont(src, 'major');
  return { rgb, minor: minor && minor !== themeFont(tgt, 'minor') ? minor : undefined, major: major && major !== themeFont(tgt, 'major') ? major : undefined };
}

// ---------------------------------------------------------------------------------------------------------------------
// formulas of a chart
// ---------------------------------------------------------------------------------------------------------------------

interface Points {
  fmt: string;
  count: string;
  pts: string;
  /** the text of the first point */
  first: string;
}

function cachePoints(block: string, p: string, level0 = false): Points {
  let b = block;
  if (level0) b = /<[\w.:-]*lvl\b[^>]*>([\s\S]*?)<\/[\w.:-]*lvl>/.exec(block)?.[1] ?? '';
  const fmt = new RegExp(`<${p}formatCode>[\\s\\S]*?</${p}formatCode>`).exec(block)?.[0] ?? '';
  const count = new RegExp(`<${p}ptCount\\b[^>]*/>`).exec(block)?.[0] ?? '';
  const pts = (b.match(new RegExp(`<${p}pt\\b[^>]*>[\\s\\S]*?</${p}pt>`, 'g')) ?? []).join('');
  const first = unesc(new RegExp(`<${p}v>([\\s\\S]*?)</${p}v>`).exec(pts)?.[1] ?? '');
  return { fmt, count, pts, first };
}

export interface ChartFormulaResult {
  xml: string;
  /** references that became literal data */
  literal: number;
  /** extension elements dropped because a reference in them could not follow */
  droppedExt: number;
  /** a formula that cannot be copied and has no literal form (chartEx): the chart is not copied */
  failed: boolean;
}

/** `p`: the prefix of the chart namespace ("c:" or ""), `dml`: of the DrawingML main namespace ("a:") */
export function rewriteChartFormulas(xml: string, rw: (f: string) => Rewrite, p: string, dml: string): ChartFormulaResult {
  let literal = 0;
  let droppedExt = 0;
  const saved: string[] = [];
  const token = (s: string) => { saved.push(s); return `\u0000${saved.length - 1}\u0000`; };
  const P = reEsc(p);
  const tag = (n: string) => new RegExp(`<${P}${n}>([\\s\\S]*?)</${P}${n}>`);
  // ---- the reference blocks (numRef / strRef / multiLvlStrRef) ----
  let out = xml.replace(new RegExp(`<${P}(numRef|strRef|multiLvlStrRef)>([\\s\\S]*?)</${P}\\1>`, 'g'), (whole: string, kind: string, inner: string, offset: number, all: string) => {
    const fm = tag('f').exec(inner);
    if (!fm) return token(whole);
    const r = rw(unesc(fm[1]));
    if (r.ok) return token(r.changed ? whole.replace(fm[0], `<${p}f>${esc(r.text)}</${p}f>`) : whole);
    literal++;
    const before = all.slice(Math.max(0, offset - 160), offset);
    const asTx = new RegExp(`<${P}tx>\\s*$`).test(before);
    const series = asTx && new RegExp(`<${P}order\\b[^>]*/>\\s*<${P}tx>\\s*$`).test(before);
    if (kind === 'numRef') {
      const c = cachePoints(inner, P);
      return token(`<${p}numLit>${c.fmt}${c.count}${c.pts}</${p}numLit>`);
    }
    if (kind === 'multiLvlStrRef') {
      const c = cachePoints(inner, P, true);
      return token(`<${p}strLit>${c.count}${c.pts}</${p}strLit>`);
    }
    const c = cachePoints(inner, P);
    if (series) return token(`<${p}v>${esc(c.first)}</${p}v>`);
    if (asTx) return token(`<${p}rich><${dml}bodyPr/><${dml}p><${dml}r><${dml}t>${esc(c.first)}</${dml}t></${dml}r></${dml}p></${p}rich>`);
    return token(`<${p}strLit>${c.count}${c.pts}</${p}strLit>`);
  });
  // ---- every other formula (c15 extensions: data label ranges, filtered series ...; chartEx data dimensions) ----
  const FAIL = '\u0001FAIL\u0001';
  out = out.replace(/<((?:[\w.-]+:)?(?:f|sqref))(\s[^>]*)?>([^<]*)<\/\1>/g, (m0, t: string, at: string | undefined, body: string) => {
    const r = rw(unesc(body));
    if (!r.ok) return FAIL;
    return r.changed ? `<${t}${at ?? ''}>${esc(r.text)}</${t}>` : m0;
  });
  if (out.indexOf(FAIL) >= 0) {
    out = out.replace(new RegExp(`<([\\w.-]+:ext|ext)\\b[^>]*>(?:(?!</\\1>)[\\s\\S])*?${FAIL}[\\s\\S]*?</\\1>`, 'g'), () => { droppedExt++; return ''; });
  }
  const failed = out.indexOf(FAIL) >= 0;
  out = out.replace(/\u0000(\d+)\u0000/g, (_m, i: string) => saved[Number(i)]);
  return { xml: out, literal, droppedExt, failed };
}

// ---------------------------------------------------------------------------------------------------------------------
// the chart part
// ---------------------------------------------------------------------------------------------------------------------

// ---------------------------------------------------------------------------------------------------------------------
// series that read a defined name
// ---------------------------------------------------------------------------------------------------------------------

const NAMES = new WeakMap<object, Array<{ name: string; local: number | null; text: string }>>();

function sourceNames(ctx: SheetObjectsContext): Array<{ name: string; local: number | null; text: string }> {
  const eng = ctx.source.eng;
  let l = NAMES.get(eng);
  if (!l) {
    l = [];
    const el = findFirst(eng.wbXml, 'definedNames');
    if (el && !el.selfClosing) {
      for (const d of findAll(el.inner, 'definedName')) {
        const name = getAttr(d.open, 'name');
        const ls = getAttr(d.open, 'localSheetId');
        if (name && !d.selfClosing) l.push({ name, local: ls === undefined || ls === '' ? null : Number(ls), text: unesc(d.inner).trim() });
      }
    }
    NAMES.set(eng, l);
  }
  return l;
}

/**
 * A series that reads a defined name ("[0]!Data", "Sheet1!Data") whose text is one plain reference on a sheet that is copied keeps reading the
 * cells: the name's range, renamed. Names that compute (OFFSET, INDEX, ...) or read several areas stay literal data.
 */
function inlineName(ctx: SheetObjectsContext, f: string): Rewrite | null {
  const m = /^(?:\[0\]|'((?:[^']|'')+)'|([^'!\s\[\]]+))!([^!\s]+)$/.exec(f.trim());
  if (!m) return null;
  const sheet = m[1] !== undefined ? m[1].replace(/''/g, "'") : m[2];
  const idx = sheet === undefined ? null : ctx.source.eng.sheets.findIndex((s) => s.name.toLowerCase() === sheet.toLowerCase());
  if (idx !== null && idx < 0) return null;
  const hit = sourceNames(ctx).find((n) => n.name.toLowerCase() === m[3].toLowerCase() && n.local === idx);
  if (!hit) return null;
  const toks = tokenize(hit.text).filter((x) => x.type !== 'ws');
  if (toks.length !== 1 || toks[0].type !== 'ref' || !toks[0].prefix || toks[0].prefix.sheet == null) return null;
  const r = ctx.names.rewrite(hit.text);
  return r.ok ? { ...r, changed: true } : null;
}

export interface ChartOptions {
  conv: ThemeConv | null;
  /** lower-cased "<source sheet>!<pivot name>" -> the pivot table's name in the target (pivot charts) */
  pivotNames: ReadonlyMap<string, string>;
  /** the date system of the target: the chart's own date1904 follows it */
  target1904: boolean;
}

export interface ChartCopy {
  path: string;
  literal: number;
  droppedExt: number;
  kind: 'chart' | 'chartEx';
}

const IMPLICIT_PARTS = [
  { type: REL_MS.chartStyle, dir: 'xl/charts/', base: 'style', ct: CT.chartStyle },
  { type: REL_MS.chartColors, dir: 'xl/charts/', base: 'colors', ct: CT.chartColors },
];

/**
 * Copies chart part `srcPath` into the stage. Returns null when the chart cannot be copied (its parts are unreadable, a formula of a
 * new-style chart cannot follow); the caller reports it and discards the stage.
 */
export async function copyChart(ctx: SheetObjectsContext, st: ObjState, stage: Stage, srcPath: string, kind: 'chart' | 'chartEx', o: ChartOptions): Promise<ChartCopy | null> {
  const sub = new Stage(st);
  let r: ChartCopy | null = null;
  try {
    r = await copyChartInto(ctx, st, sub, srcPath, kind, o);
  } catch (e) {
    // a damaged part of the chart itself: the chart is left out (the other objects of the drawing stay)
    if (!(e instanceof PatchError)) throw e;
  }
  if (!r) { sub.discard(); return null; }
  stage.adopt(sub);
  return r;
}

async function copyChartInto(ctx: SheetObjectsContext, st: ObjState, stage: Stage, srcPath: string, kind: 'chart' | 'chartEx', o: ChartOptions): Promise<ChartCopy | null> {
  const eng = ctx.source.eng;
  const raw = await eng.pkg.read(srcPath);
  if (!raw) return null;
  let x = raw;
  const p = nsPrefix(x, kind === 'chart' ? dmlUri('chart', st.strict) : 'http://schemas.microsoft.com/office/drawing/2014/chartex', kind === 'chart' ? 'c:' : 'cx:');
  const dml = nsPrefix(x, dmlUri('main', st.strict), 'a:');
  const P = reEsc(p);
  const rels = await readSrcRels(eng, srcPath);
  const out: OutRel[] = [];
  let n = 0;
  const newId = () => `rId${++n}`;

  // ---- formulas ----
  const fr = rewriteChartFormulas(x, (f) => ctx.names.rewrite(f).ok ? ctx.names.rewrite(f) : (inlineName(ctx, f) ?? ctx.names.rewrite(f)), p, dml);
  if (fr.failed) return null;
  if (kind === 'chartEx' && fr.literal > 0) return null;
  x = fr.xml;

  // ---- theme ----
  if (!/<[\w.:-]*clrMapOvr\b/.test(x)) x = convertTheme(x, o.conv, dml);

  // ---- pivot charts follow the pivot table copy; without one the chart is an ordinary chart ----
  x = x.replace(new RegExp(`<${P}pivotSource>\\s*<${P}name>([^<]*)</${P}name>[\\s\\S]*?</${P}pivotSource>`), (m0, name: string) => {
    const m = /^\[([^\]]*)\](.*)!([^!]*)$/.exec(unesc(name));
    if (!m) return '';
    const sheetSrc = m[2].replace(/^'|'$/g, '');
    const hit = o.pivotNames.get(`${sheetSrc.toLowerCase()}!${m[3].toLowerCase()}`);
    const sheet = ctx.names.rename.get(sheetSrc.toLowerCase());
    return hit === undefined || sheet === undefined ? '' : m0.replace(name, esc(`[${m[1]}]${sheet}!${hit}`));
  });
  if (!new RegExp(`<${P}pivotSource>`).test(x)) {
    x = x.replace(new RegExp(`<${P}pivotFmts>[\\s\\S]*?</${P}pivotFmts>`), '');
    x = x.replace(/<([\w.-]+:ext|ext)\b[^>]*>(?:(?!<\/\1>)[\s\S])*?pivotOptions[\s\S]*?<\/\1>/g, '');
    x = x.replace(new RegExp(`<${P}extLst>\\s*</${P}extLst>`, 'g'), '');
  }

  // ---- date system ----
  x = x.replace(new RegExp(`<${P}date1904\\s+val="[^"]*"/>`), `<${p}date1904 val="${o.target1904 ? 1 : 0}"/>`);

  // ---- the parts the chart owns, referenced by r:id / r:embed ----
  const rid = (m: string): SrcRel | undefined => rels.find((r) => r.id === /\sr:(?:id|embed)="([^"]*)"/.exec(m)?.[1]);
  const setRid = (m: string, id: string): string => m.replace(/(\sr:(?:id|embed)=")[^"]*(")/, `$1${id}$2`);
  // embedded workbook / linked data
  x = await replaceAsync(x, new RegExp(`<${P}externalData\\b[^>]*?(?:/>|>[\\s\\S]*?</${P}externalData>)`, 'g'), async (m) => {
    const rel = rid(m[0]);
    if (!rel) return '';
    let nr: OutRel | null = null;
    if (rel.external) nr = { id: newId(), type: relType(st, 'oleObject'), target: rel.target, external: true };
    else if (rel.path && (relIs(rel, 'package') || relIs(rel, 'oleObject'))) {
      const ext = /\.([^./]+)$/.exec(rel.path)?.[1] ?? 'xlsx';
      const path = await stage.copyBinary(eng, rel.path, 'xl/embeddings/', 'Microsoft_Excel_Sheet', `.${ext}`, false);
      if (path) nr = { id: newId(), type: relType(st, relIs(rel, 'package') ? 'package' : 'oleObject'), target: path };
    }
    if (!nr) return '';
    out.push(nr);
    return setRid(m[0], nr.id);
  });
  // shapes drawn on the chart
  x = await replaceAsync(x, new RegExp(`<${P}userShapes\\b[^>]*?/>`, 'g'), async (m) => {
    const rel = rid(m[0]);
    if (!rel || rel.external || !rel.path || !relIs(rel, 'chartUserShapes') || (await readSrcRels(eng, rel.path)).length) return '';
    const ux = await eng.pkg.read(rel.path);
    if (!ux) return '';
    const path = freshPart(st, 'xl/drawings/', 'drawing', '.xml');
    stage.text(path, ux, CT.chartShapes);
    const id = newId();
    out.push({ id, type: relType(st, 'chartUserShapes'), target: path });
    return setRid(m[0], id);
  });
  // pictures (blip fills)
  x = await replaceAsync(x, /<[\w.-]+:blip\b[^>]*\sr:embed="[^"]*"[^>]*?>|<blip\b[^>]*\sr:embed="[^"]*"[^>]*?>/g, async (m) => {
    const rel = rid(m[0]);
    if (!rel || rel.external || !rel.path || !relIs(rel, 'image')) return m[0].replace(/\sr:embed="[^"]*"/, '');
    const ext = (/\.([^./]+)$/.exec(rel.path)?.[1] ?? 'png').toLowerCase();
    const path = await stage.copyBinary(eng, rel.path, 'xl/media/', 'image', `.${ext}`, true);
    if (!path) return m[0].replace(/\sr:embed="[^"]*"/, '');
    let id = out.find((r) => r.target === path)?.id;
    if (!id) { id = newId(); out.push({ id, type: relType(st, 'image'), target: path }); }
    return setRid(m[0], id);
  });
  // style / colours / theme override: implicit parts of the chart (no r:id)
  for (const rel of rels) {
    if (rel.external || !rel.path) continue;
    const known = IMPLICIT_PARTS.find((c) => c.type === rel.type);
    const isTheme = relIs(rel, 'themeOverride');
    if (!known && !isTheme) continue;
    const path = known ? freshPart(st, known.dir, known.base, '.xml') : freshPart(st, 'xl/theme/', 'themeOverride', '.xml');
    const ct = known ? known.ct : CT.themeOverride;
    const bytes = await eng.pkg.readBytes(rel.path);
    if (!bytes) continue;
    try {
      stage.binaryAs(path, bytes, ct);
    } catch (e) {
      if (!(e instanceof PatchError)) throw e;
      continue; // a damaged style / colour / theme part: the chart does without it
    }
    out.push({ id: newId(), type: known ? known.type : relType(st, 'themeOverride'), target: path });
  }
  // an r:id left that points at something not copied would dangle: the chart is not copied
  for (const m of x.matchAll(/\sr:(?:id|embed|link|pict)="([^"]*)"/g)) if (!out.some((r) => r.id === m[1])) return null;

  const path = freshPart(st, 'xl/charts/', kind === 'chart' ? 'chart' : 'chartEx', '.xml');
  stage.text(path, x, kind === 'chart' ? CT.chart : CT.chartEx);
  stage.setRels(path, out);
  return { path, literal: fr.literal, droppedExt: fr.droppedExt, kind };
}
