/**
 * Pivot tables of a copied sheet (task A2, see addsheets-objects.ts).
 *
 * A pivot table is three kinds of parts: the pivot table definition (per sheet), the pivot CACHE definition (shared by every pivot table
 * of the workbook that was made from the same data; its source range lives here) and the cache records (a snapshot of the data). A copy
 * gets all three, with a new cacheId and a workbook <pivotCaches> entry, when the cache reads a range of a sheet that is copied too (the
 * sheet name in <worksheetSource> follows the copy; a Table source follows the table's new name). Pivot tables made from external data,
 * the data model, several ranges (consolidation) or a defined name stay as plain cell values (nothing of the pivot is written, the cells
 * are there): the copied sheet is an ordinary sheet in that area, and the warning says so.
 *
 * The copy is refreshed by Excel on demand, not on open: the cells of the sheet are the copied cells and the cache is the copied snapshot.
 * (Only a source in the other date system is told to refresh when it is opened.)
 */
import { BUILTIN_NUMFMT_MAX } from './styles';
import { is1904, isStrict } from './bookinfo';
import { PatchError } from './pkg';
import { findFirst, getAttr, setAttr } from './xml';
import type { SheetObjectsContext } from './copytypes';
import {
  CT, Stage, addInfo, freshPart, partFlavour, readSrcRels, relIs, relType,
  type ObjState, type OutRel, type PivotCacheCopy,
} from './addsheets-objects-core';
import { ensureTableStyle, uniqueName } from './addsheets-objects-table';

export interface PivotCopy {
  rels: OutRel[];
  copied: number;
  /** pivot tables whose data cannot be reached: left as plain cells */
  values: number;
  /** lower-cased "<source sheet>!<pivot name>" -> the new pivot table name (pivot charts follow it) */
  names: Map<string, string>;
}

/** the number format id of the target for a format id of the source's styles.xml */
export function mapNumFmt(ctx: SheetObjectsContext, id: number): number {
  const code = ctx.source.styles.numFmtCode(id);
  if (code !== undefined) return ctx.target.styles.numFmtId(code);
  return id > BUILTIN_NUMFMT_MAX ? 0 : id;
}

function addPivotCacheEntry(wb: string, id: number, rid: string): string {
  const el = `<pivotCache cacheId="${id}" r:id="${rid}"/>`;
  const pc = findFirst(wb, 'pivotCaches');
  if (pc && !pc.selfClosing) return wb.slice(0, pc.end - '</pivotCaches>'.length) + el + '</pivotCaches>' + wb.slice(pc.end);
  const block = `<pivotCaches>${el}</pivotCaches>`;
  if (pc) return wb.slice(0, pc.start) + block + wb.slice(pc.end);
  for (const tag of ['smartTagPr', 'smartTagTypes', 'webPublishing', 'fileRecoveryPr', 'webPublishObjects', 'extLst']) {
    const m = new RegExp(`<${tag}\\b`).exec(wb);
    if (m) return wb.slice(0, m.index) + block + wb.slice(m.index);
  }
  return wb.replace(/<\/workbook>\s*$/, block + '</workbook>');
}

/** the cache definition of a pivot table: its own relationship, else the workbook's <pivotCache cacheId> entry */
async function cacheDefinitionPath(ctx: SheetObjectsContext, pivotPath: string, cacheId: string | undefined): Promise<string | null> {
  const eng = ctx.source.eng;
  const own = (await readSrcRels(eng, pivotPath)).find((r) => !r.external && r.path && relIs(r, 'pivotCacheDefinition'));
  if (own) return own.path;
  if (cacheId === undefined) return null;
  for (const m of eng.wbXml.matchAll(/<pivotCache\b[^>]*>/g)) {
    if (getAttr(m[0], 'cacheId') !== cacheId) continue;
    const rid = getAttr(m[0], 'r:id');
    const rel = (await readSrcRels(eng, eng.wbPath)).find((r) => r.id === rid && !r.external && r.path);
    if (rel) return rel.path;
  }
  return null;
}

/** copies a cache definition (+ records) into the target and registers it; null when its data cannot be reached from the copy */
async function ensureCache(ctx: SheetObjectsContext, st: ObjState, defPath: string): Promise<PivotCacheCopy | null> {
  const eng = ctx.source.eng;
  let map = st.caches.get(eng);
  if (!map) st.caches.set(eng, (map = new Map()));
  const key = defPath.toLowerCase();
  if (map.has(key)) return map.get(key) ?? null;
  map.set(key, null);
  const dxRaw = await eng.pkg.read(defPath);
  if (!dxRaw) return null;
  const srcStrict = isStrict(eng.wbXml);
  const dx = partFlavour(dxRaw, srcStrict, st.strict);
  const root = /<pivotCacheDefinition\b[^>]*>/.exec(dx)?.[0] ?? '';
  const source = /<cacheSource\b[^>]*>/.exec(dx)?.[0] ?? '';
  const ws = /<worksheetSource\b[^>]*>/.exec(dx)?.[0] ?? '';
  if (!root || getAttr(source, 'type') !== 'worksheet' || getAttr(source, 'connectionId') !== undefined || !ws || getAttr(ws, 'r:id') !== undefined) return null;
  if (/<cacheHierarchies\b|<tupleCache\b/.test(dx)) return null;
  // the source: a range of a copied sheet, or a table that is copied
  let newWs = ws;
  const sheet = getAttr(ws, 'sheet');
  const tname = getAttr(ws, 'name');
  if (sheet !== undefined) {
    const nn = ctx.names.rename.get(sheet.toLowerCase());
    if (nn === undefined) return null;
    newWs = setAttr(ws, 'sheet', nn);
  } else if (tname !== undefined) {
    const nt = ctx.names.copier.opts.tables?.get(tname.toLowerCase());
    if (nt === undefined) return null;
    newWs = setAttr(ws, 'name', nt);
  } else return null;

  const stage = new Stage(st);
  const recRel = (await readSrcRels(eng, defPath)).find((r) => !r.external && r.path && relIs(r, 'pivotCacheRecords'));
  const newDef = freshPart(st, 'xl/pivotCache/', 'pivotCacheDefinition', '.xml');
  let recPath: string | null = null;
  if (recRel) {
    // (the records are copied byte by byte unless the two packages use other namespaces)
    const bytes = srcStrict === st.strict ? await eng.pkg.readBytes(recRel.path) : null;
    const text = srcStrict === st.strict ? null : await eng.pkg.read(recRel.path);
    if (bytes || text) {
      const path = freshPart(st, 'xl/pivotCache/', 'pivotCacheRecords', '.xml');
      try {
        if (bytes) stage.binaryAs(path, bytes, CT.pivotCacheRec);
        else stage.text(path, partFlavour(text!, srcStrict, st.strict), CT.pivotCacheRec);
        recPath = path;
      } catch (e) {
        // damaged records: the cache is written without them and refreshed from the (copied) data when the file is opened
        if (!(e instanceof PatchError)) throw e;
      }
    }
  }
  const dateShift = is1904(eng) !== is1904(ctx.target.eng);
  let nd = dx.replace(ws, newWs);
  nd = nd.replace(/<pivotCacheDefinition\b[^>]*>/, (o) => {
    let t = o.replace(/\sxr\d*:uid="[^"]*"/g, '');
    t = setAttr(t, 'r:id', recPath ? 'rId1' : null);
    if (!recPath || dateShift) t = setAttr(t, 'refreshOnLoad', '1');
    return t;
  });
  nd = nd.replace(/<cacheField\b[^>]*>/g, (o) => {
    const v = getAttr(o, 'numFmtId');
    return v === undefined ? o : setAttr(o, 'numFmtId', String(mapNumFmt(ctx, Number(v))));
  });
  nd = nd.replace(/(<x14:pivotCacheDefinition\b[^>]*?)\spivotCacheId="[^"]*"/g, '$1');
  const id = (st.cacheId += 1);
  stage.text(newDef, nd, CT.pivotCacheDef);
  if (recPath) stage.setRels(newDef, [{ id: 'rId1', type: relType(st, 'pivotCacheRecords'), target: recPath }]);
  stage.onCommit(async () => {
    const rid = await st.eng.pkg.addRel(st.eng.wbPath, relType(st, 'pivotCacheDefinition'), newDef);
    st.eng.wbXml = addPivotCacheEntry(st.eng.wbXml, id, rid);
  });
  await stage.commit();
  const out = { id, path: newDef };
  map.set(key, out);
  return out;
}

export async function copyPivots(ctx: SheetObjectsContext, st: ObjState, stage: Stage): Promise<PivotCopy> {
  const out: PivotCopy = { rels: [], copied: 0, values: 0, names: new Map() };
  const eng = ctx.source.eng;
  for (const rel of ctx.source.rels.filter((r) => !r.external && r.path && relIs(r, 'pivotTable'))) {
    const pxRaw = await eng.pkg.read(rel.path);
    const px = pxRaw ? partFlavour(pxRaw, isStrict(eng.wbXml), st.strict) : pxRaw;
    const open = /<pivotTableDefinition\b[^>]*>/.exec(px ?? '')?.[0];
    if (!px || !open) { out.values++; continue; }
    const name = getAttr(open, 'name') ?? '';
    const defPath = await cacheDefinitionPath(ctx, rel.path, getAttr(open, 'cacheId'));
    const cache = defPath ? await ensureCache(ctx, st, defPath) : null;
    if (!cache) { out.values++; continue; }
    const newName = uniqueName(name, st.pivotNames);
    st.pivotNames.add(newName.toLowerCase());
    if (newName !== name) addInfo(ctx.target.eng, { severity: 'info', code: 'pivotRenamed', sourceId: ctx.source.ks.id, message: `シート「${ctx.target.name}」のピボットテーブル「${name}」は、名前が重なったため「${newName}」にしました。` });
    let nx = px.replace(/<pivotTableDefinition\b[^>]*>/, (o) => setAttr(setAttr(o.replace(/\sxr\d*:uid="[^"]*"/g, ''), 'cacheId', String(cache.id)), 'name', newName));
    nx = nx.replace(/<(pivotField|dataField)\b[^>]*>/g, (o) => {
      const v = getAttr(o, 'numFmtId');
      return v === undefined ? o : setAttr(o, 'numFmtId', String(mapNumFmt(ctx, Number(v))));
    });
    // formats (dxf ids of the source)
    nx = nx.replace(/<formats\b[^>]*>([\s\S]*?)<\/formats>/, (_m, inner: string) => {
      let n = 0;
      const body = inner.replace(/<format\b[^>]*?(?:\/>|>[\s\S]*?<\/format>)/g, (f) => {
        const d = getAttr(f, 'dxfId');
        if (d === undefined) { n++; return f; }
        const nd = ctx.maps.dxf(Number(d));
        if (nd === null) return '';
        n++;
        return f.replace(/^<format\b[^>]*>/, (o) => setAttr(o, 'dxfId', String(nd)));
      });
      return n ? `<formats count="${n}">${body}</formats>` : '';
    });
    nx = nx.replace(/<pivotTableStyleInfo\b[^>]*?\/>/, (el) => {
      const sn = getAttr(el, 'name');
      return sn === undefined ? el : setAttr(el, 'name', ensureTableStyle(ctx, st, sn) ?? null);
    });
    const path = freshPart(st, 'xl/pivotTables/', 'pivotTable', '.xml');
    stage.text(path, nx, CT.pivotTable);
    stage.setRels(path, [{ id: 'rId1', type: relType(st, 'pivotCacheDefinition'), target: cache.path }]);
    out.rels.push({ id: ctx.newRelId(), type: relType(st, 'pivotTable'), target: path });
    out.names.set(`${ctx.source.info.name.toLowerCase()}!${name.toLowerCase()}`, newName);
    out.copied++;
  }
  return out;
}
