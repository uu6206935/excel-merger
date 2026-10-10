/**
 * The drawing part of a copied sheet (task A2, see addsheets-objects.ts): charts, pictures, shapes, text boxes, group shapes, connectors.
 *
 * The drawing is read anchor by anchor (an anchor = one object and its position, written in cells, so the copy stays where it was because
 * the cells of the copy are in the same places). Each anchor is copied or left out as a whole:
 *   charts / new-style charts   copyChart (addsheets-objects-chart.ts)
 *   pictures                    the image part is copied (an image used by several pictures once per source file), linked pictures keep their link
 *   shapes, text boxes, groups, connectors   as they are; a shape that shows a cell (textlink) follows the renamed sheet, a macro name is removed,
 *                               a link to another sheet (hyperlink of a shape) follows the renamed copy
 *   slicers / timelines, SmartArt, ink, the hidden placeholders of form controls and OLE objects   left out (counted, the caller reports them)
 */
import { esc, nsPrefix, unesc } from './xml';
import type { SheetObjectsContext } from './copytypes';
import { copyChart, convertTheme, themeConversion, type ChartOptions } from './addsheets-objects-chart';
import {
  CT, REL_MS, Stage, dmlUri, freshPart, readSrcRels, relIs, relType, replaceAsync,
  type ObjState, type OutRel, type SrcRel,
} from './addsheets-objects-core';

export interface DrawingCopy {
  /** the new drawing part; null when no anchor was copied */
  path: string | null;
  charts: number;
  pictures: number;
  shapes: number;
  connectors: number;
  groups: number;
  /** left out */
  slicers: number;
  diagrams: number;
  controls: number;
  other: number;
  /** charts that could not be copied */
  chartsLost: number;
  /** series / titles of the copied charts that keep their values as literal data */
  literal: number;
}

/** the children of an element's inner text: complete elements (matching close tag found by depth), comments skipped */
export function splitChildren(inner: string): Array<{ tag: string; text: string }> {
  const out: Array<{ tag: string; text: string }> = [];
  const open = /<([A-Za-z_][\w.:-]*)(?=[\s/>])[^>]*?(\/?)>/y;
  let i = 0;
  while (i < inner.length) {
    const lt = inner.indexOf('<', i);
    if (lt < 0) break;
    if (inner.startsWith('<!--', lt)) { const e = inner.indexOf('-->', lt); i = e < 0 ? inner.length : e + 3; continue; }
    if (inner.startsWith('<?', lt)) { const e = inner.indexOf('?>', lt); i = e < 0 ? inner.length : e + 2; continue; }
    if (inner.startsWith('</', lt)) break;
    open.lastIndex = lt;
    const m = open.exec(inner);
    if (!m) break;
    if (m[2] === '/') { out.push({ tag: m[1], text: m[0] }); i = lt + m[0].length; continue; }
    const name = m[1].replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const re = new RegExp(`<(/?)${name}(?=[\\s/>])[^>]*?(/?)>`, 'g');
    re.lastIndex = lt + m[0].length;
    let depth = 1;
    let end = -1;
    for (let e = re.exec(inner); e; e = re.exec(inner)) {
      if (e[1] === '/') { if (--depth === 0) { end = e.index + e[0].length; break; } } else if (e[2] !== '/') depth++;
    }
    if (end < 0) break;
    out.push({ tag: m[1], text: inner.slice(lt, end) });
    i = end;
  }
  return out;
}

type Why = 'slicer' | 'diagram' | 'control' | 'ink' | 'other';

/** what an anchor is, by what it contains */
function classify(unit: string): Why | null {
  if (/drawing\/2010\/slicer|drawing\/2012\/timeslicer|<sle:|<tsle:/.test(unit)) return 'slicer';
  if (/drawingml\/2006\/diagram|purl\.oclc\.org\/ooxml\/drawingml\/diagram|<dgm:relIds/.test(unit)) return 'diagram';
  if (/<(?:[\w.-]+:)?contentPart\b/.test(unit)) return 'ink';
  if (/<a14:compatExt\b/.test(unit)) return 'control';
  return null;
}

const countOf = (unit: string, tag: string): number => (unit.match(new RegExp(`<(?:[\\w.-]+:)?${tag}[\\s>/]`, 'g')) ?? []).length;

export interface DrawingOptions {
  /** lower-cased "<source sheet>!<pivot name>" -> the pivot table's name in the target (pivot charts follow it) */
  pivotNames: ReadonlyMap<string, string>;
}

/** Copy the drawing part `srcPath` of a sheet into `stage` (nothing is written before the caller commits). */
export async function copyDrawing(ctx: SheetObjectsContext, st: ObjState, stage: Stage, srcPath: string, o: DrawingOptions): Promise<DrawingCopy> {
  const eng = ctx.source.eng;
  const res: DrawingCopy = { path: null, charts: 0, pictures: 0, shapes: 0, connectors: 0, groups: 0, slicers: 0, diagrams: 0, controls: 0, other: 0, chartsLost: 0, literal: 0 };
  const raw = await eng.pkg.read(srcPath);
  if (!raw) return res;
  const x = raw;
  const rootOpen = /<(?:[\w.-]+:)?wsDr\b[^>]*>/.exec(x);
  const rootClose = /<\/(?:[\w.-]+:)?wsDr>\s*$/.exec(x);
  if (!rootOpen || !rootClose) return res;
  const inner = x.slice(rootOpen.index + rootOpen[0].length, rootClose.index);
  const dml = nsPrefix(x, dmlUri('main', st.strict), 'a:');
  const srcRels = await readSrcRels(eng, srcPath);
  const conv = await themeConversion(ctx);
  const target1904 = /<workbookPr\b[^>]*\sdate1904="(1|true)"/.test(ctx.target.eng.wbXml);
  const chartOpts: ChartOptions = { conv, pivotNames: o.pivotNames, target1904 };
  const outRels: OutRel[] = [];
  let n = 0;
  const newId = () => `rId${++n}`;
  /** source relationship id -> relationship id in the new drawing */
  const done = new Map<string, string>();
  const kept: string[] = [];

  for (const child of splitChildren(inner)) {
    let unit = child.text;
    const why = classify(unit);
    if (why) {
      if (why === 'slicer') res.slicers++;
      else if (why === 'diagram') res.diagrams++;
      else if (why === 'control') res.controls++;
      else res.other++;
      continue;
    }
    // every relationship the anchor uses must be copyable, else the anchor is left out as a whole (with what was copied for it)
    const us = new Stage(st);
    const unitRels: OutRel[] = [];
    const unitDone = new Map<string, string>();
    const state: { lost: 'chart' | 'other' | null; literal: number } = { lost: null, literal: 0 };
    unit = await replaceAsync(unit, /(\sr:(?:id|embed|link|pict))="([^"]*)"/g, async (m) => {
      if (state.lost) return m[0];
      const key = m[2];
      if (!key) return m[0]; // an empty id (a shape with an action but no link)
      const hit = done.get(key) ?? unitDone.get(key);
      if (hit) return `${m[1]}="${hit}"`;
      const rel = srcRels.find((r) => r.id === key);
      const nr = rel ? await copyRel(ctx, st, us, rel, chartOpts, newId, (c) => { state.literal += c; }) : null;
      if (!nr) { state.lost = rel && (relIs(rel, 'chart') || rel.type === REL_MS.chartEx) ? 'chart' : 'other'; return m[0]; }
      if (!nr.rel) return `${m[1]}="${key}" data-drop`;
      unitRels.push(nr.rel);
      unitDone.set(key, nr.rel.id);
      return `${m[1]}="${nr.rel.id}"`;
    });
    if (state.lost) {
      us.discard();
      if (state.lost === 'chart') res.chartsLost++;
      else res.other++;
      continue;
    }
    stage.adopt(us);
    outRels.push(...unitRels);
    for (const [k, v] of unitDone) done.set(k, v);
    // hyperlinks that could not follow: the element goes, the shape stays
    unit = unit.replace(new RegExp(`<${dml}hlink(?:Click|Hover)\\b[^>]*\\sdata-drop[^>]*?(?:/>|>[\\s\\S]*?</${dml}hlink(?:Click|Hover)>)`, 'g'), '');
    unit = unit.replace(/ data-drop/g, '');
    res.literal += state.literal;
    // text that shows a cell, macro names
    unit = unit.replace(/(<(?:[\w.-]+:)?(?:sp|pic|graphicFrame|cxnSp|grpSp)\b[^>]*?)\smacro="[^"]+"/g, '$1 macro=""');
    unit = unit.replace(/\stextlink="([^"]*)"/g, (m0, f: string) => {
      const t = unesc(f);
      if (!t.trim()) return m0;
      const r = ctx.names.rewrite(t.replace(/^=/, ''));
      return r.ok ? ` textlink="${esc(r.text).replace(/"/g, '&quot;')}"` : ' textlink=""';
    });
    unit = convertTheme(unit, conv, dml);
    res.charts += countOf(unit, 'chart');
    res.pictures += countOf(unit, 'pic');
    res.shapes += countOf(unit, 'sp');
    res.connectors += countOf(unit, 'cxnSp');
    res.groups += countOf(unit, 'grpSp');
    kept.push(unit);
  }
  if (!kept.length) return res;
  const path = freshPart(st, 'xl/drawings/', 'drawing', '.xml');
  const prolog = /^[\s\S]*?(?=<(?:[\w.-]+:)?wsDr\b)/.exec(x)?.[0] ?? '';
  stage.text(path, `${prolog}${rootOpen[0]}${kept.join('')}${rootClose[0].trim()}`, CT.drawing);
  stage.setRels(path, outRels);
  res.path = path;
  return res;
}

/** one relationship of a drawing in the target: null = cannot be copied (the anchor is left out) */
async function copyRel(ctx: SheetObjectsContext, st: ObjState, stage: Stage, rel: SrcRel, o: ChartOptions, newId: () => string, addLiteral: (n: number) => void): Promise<{ rel: OutRel | null; drop?: boolean } | null> {
  const eng = ctx.source.eng;
  if (relIs(rel, 'hyperlink')) {
    if (rel.external) return { rel: { id: newId(), type: relType(st, 'hyperlink'), target: rel.target, external: true } };
    if (rel.fragment) {
      const r = ctx.names.rewrite(rel.target.slice(1));
      if (!r.ok) return { rel: null, drop: true };
      return { rel: { id: newId(), type: relType(st, 'hyperlink'), target: `#${r.text}`, raw: true } };
    }
    return { rel: null, drop: true };
  }
  if (relIs(rel, 'image')) {
    if (rel.external) return { rel: { id: newId(), type: relType(st, 'image'), target: rel.target, external: true } };
    if (!rel.path) return null;
    const ext = (/\.([^./]+)$/.exec(rel.path)?.[1] ?? 'png').toLowerCase();
    const path = await stage.copyBinary(eng, rel.path, 'xl/media/', 'image', `.${ext}`, true);
    return path ? { rel: { id: newId(), type: relType(st, 'image'), target: path } } : null;
  }
  if ((relIs(rel, 'chart') || rel.type === REL_MS.chartEx) && rel.path && !rel.external) {
    const kind = rel.type === REL_MS.chartEx ? 'chartEx' : 'chart';
    const c = await copyChart(ctx, st, stage, rel.path, kind, o);
    if (!c) return null;
    addLiteral(c.literal + c.droppedExt);
    return { rel: { id: newId(), type: kind === 'chart' ? relType(st, 'chart') : REL_MS.chartEx, target: c.path } };
  }
  return null;
}
