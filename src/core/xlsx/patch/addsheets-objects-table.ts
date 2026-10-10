/**
 * Tables (ListObjects) of a copied sheet (task A2, see addsheets-objects.ts).
 *
 * Planning (before any formula is copied): which tables of a source are copied and under which name. A table keeps its name unless the
 * workbook has it already (a table of the merged sheet, of a kept base sheet, of an earlier copy, a defined name): then "_2", "_3" ... is
 * appended. The formula copier re-points every structured reference at the new name (cells, defined names, rules, validations, charts).
 *
 * Copy: new part with a workbook-unique id, unique column ids kept, dxf ids of the table (header / data / totals formats) re-mapped to the
 * target's dxfs, custom table styles copied into the target's <tableStyles> (renamed on a clash), calculated-column / totals formulas
 * rewritten for the renamed sheets and tables. Tables that are not plain worksheet tables (XML-mapped tables, query tables with an
 * external connection) are not copied: the cells stay, the table function is left out.
 */
import { esc, findAll, findFirst, getAttr, setAttr, unesc } from './xml';
import { canon } from './styles';
import type { SheetObjectsContext } from './copytypes';
import type { SourcePlan } from './copytypes';
import type { Engine } from './engine';
import { isRel } from './pkg';
import { isStrict } from './bookinfo';
import { CT, freshPart, partFlavour, readSrcRels, relType, type ObjState, type OutRel, type Stage } from './addsheets-objects-core';

/** what A2 decides about the tables of a source before the formulas are copied (see ObjectsPlan in addsheets.ts) */
export interface TablePlan {
  tables: Map<string, string>;
  allTables: Set<string>;
  sheetsWithTables: Set<string>;
}

/** a table part is a plain worksheet table that can be copied (no XML mapping, no external data connection) */
export function tableCopyable(xml: string, partRels: number): boolean {
  const open = /<table\b[^>]*>/.exec(xml)?.[0] ?? '';
  if (!open) return false;
  const type = getAttr(open, 'tableType');
  if (type !== undefined && type !== 'worksheet') return false;
  if (getAttr(open, 'connectionId') !== undefined) return false;
  if (partRels > 0) return false; // a query table part
  if (/<xmlColumnPr\b|\squeryTableFieldId=/.test(xml)) return false;
  return /<tableColumns\b/.test(xml) && !!getAttr(open, 'ref');
}

/** "<name>_2", "_3" ... until `used` has no such (lower-cased) name */
export function uniqueName(name: string, used: ReadonlySet<string>): string {
  if (!used.has(name.toLowerCase())) return name;
  for (let n = 2; ; n++) {
    const c = `${name}_${n}`;
    if (!used.has(c.toLowerCase())) return c;
  }
}

export async function planTables(source: SourcePlan, taken: ReadonlySet<string>): Promise<TablePlan> {
  const out: TablePlan = { tables: new Map(), allTables: new Set(), sheetsWithTables: new Set() };
  const used = new Set(taken);
  const eng: Engine = source.eng;
  for (const sh of source.sheets) {
    if (sh.name === null || !sh.info.worksheet || !sh.info.path) continue;
    for (const r of (await readSrcRels(eng, sh.info.path)).filter((x) => !x.external && !x.fragment && isRel(x.type, 'table'))) {
      const x = await eng.pkg.read(r.path);
      if (!x) continue;
      const open = /<table\b[^>]*>/.exec(x)?.[0] ?? '';
      const display = getAttr(open, 'displayName') ?? getAttr(open, 'name');
      if (!display) continue;
      const key = display.toLowerCase();
      out.allTables.add(key);
      const nm = getAttr(open, 'name');
      if (nm) out.allTables.add(nm.toLowerCase());
      if (out.tables.has(key)) continue;
      if (!tableCopyable(x, (await eng.pkg.rels(r.path)).length)) continue;
      const fresh = uniqueName(display, used);
      used.add(fresh.toLowerCase());
      out.tables.set(key, fresh);
      out.sheetsWithTables.add(sh.original.toLowerCase());
    }
  }
  return out;
}

// ---------------------------------------------------------------------------------------------------------------------
// custom table styles
// ---------------------------------------------------------------------------------------------------------------------

const BUILTIN_STYLE = /^(TableStyle|PivotStyle|SlicerStyle|TimeSlicerStyle)(Light|Medium|Dark)?\d*$/i;

function styleBlocks(xml: string): Array<{ name: string; text: string }> {
  const ts = findFirst(xml, 'tableStyles');
  if (!ts || ts.selfClosing) return [];
  return findAll(ts.inner, 'tableStyle').map((e) => ({ name: getAttr(e.open, 'name') ?? '', text: e.text }));
}

/**
 * The name a table (or pivot) style has in the target: built-in styles keep their name; a custom style of the source is copied into the
 * target's <tableStyles> (its dxfs mapped to the target's) under the same name, or under "<name>_2" when the target has another style of
 * that name. undefined = the style does not exist in the source (the table gets no style).
 */
export function ensureTableStyle(ctx: SheetObjectsContext, st: ObjState, name: string): string | undefined {
  if (BUILTIN_STYLE.test(name)) return name;
  const eng = ctx.source.eng;
  let m = st.tableStyles.get(eng);
  if (!m) st.tableStyles.set(eng, (m = new Map()));
  const done = m.get(name);
  if (done !== undefined) return done;
  const src = styleBlocks(ctx.source.styles.xml).find((b) => b.name === name);
  if (!src) return undefined;
  // dxf ids of the style elements: the target's
  let body = src.text.replace(/<tableStyleElement\b[^>]*?\/>/g, (el) => {
    const d = getAttr(el, 'dxfId');
    if (d === undefined) return el;
    const nd = ctx.maps.dxf(Number(d));
    return nd === null ? '' : setAttr(el, 'dxfId', String(nd));
  });
  body = body.replace(/\sxr\d*:uid="[^"]*"/g, '');
  const count = (body.match(/<tableStyleElement\b/g) ?? []).length;
  body = body.replace(/^<tableStyle\b[^>]*>/, (o) => setAttr(o, 'count', String(count)));
  const target = ctx.target.styles;
  const have = styleBlocks(target.xml);
  let final = name;
  const same = have.find((b) => b.name === name);
  if (same) {
    // the same style (identical elements; the uid and the element count of the writer do not count)
    const plain = (s: string): string => canon(s.replace(/\sxr\d*:uid="[^"]*"/g, ''));
    if (plain(same.text) === plain(body)) { m.set(name, name); return name; }
    const names = new Set(have.map((b) => b.name.toLowerCase()));
    final = uniqueName(name, names);
    body = body.replace(/^<tableStyle\b[^>]*>/, (o) => setAttr(o, 'name', final));
  }
  target.xml = insertTableStyle(target.xml, body);
  target.changed = true;
  m.set(name, final);
  return final;
}

function insertTableStyle(xml: string, el: string): string {
  const ts = findFirst(xml, 'tableStyles');
  if (ts && !ts.selfClosing) {
    const n = (ts.inner.match(/<tableStyle\b/g) ?? []).length + 1;
    const open = setAttr(ts.open, 'count', String(n));
    return xml.slice(0, ts.start) + open + ts.inner + el + '</tableStyles>' + xml.slice(ts.end);
  }
  if (ts) return xml.slice(0, ts.start) + `${setAttr(ts.open, 'count', '1').replace(/\/>$/, '>')}${el}</tableStyles>` + xml.slice(ts.end);
  const section = `<tableStyles count="1" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16">${el}</tableStyles>`;
  const dx = /<\/dxfs>|<dxfs\b[^>]*\/>/.exec(xml);
  if (dx) return xml.slice(0, dx.index + dx[0].length) + section + xml.slice(dx.index + dx[0].length);
  for (const tag of ['colors', 'extLst']) {
    const m = new RegExp(`<${tag}\\b`).exec(xml);
    if (m) return xml.slice(0, m.index) + section + xml.slice(m.index);
  }
  return xml.replace(/<\/styleSheet>\s*$/, section + '</styleSheet>');
}

// ---------------------------------------------------------------------------------------------------------------------
// the copy
// ---------------------------------------------------------------------------------------------------------------------

export interface TableCopy {
  /** <tablePart r:id> elements, relationship ids of the new sheet */
  parts: Array<{ rel: OutRel }>;
  copied: number;
  /** tables that stay plain cells (XML-mapped / query tables) */
  skipped: number;
  /** tables that got another name than they had */
  renamed: Array<{ from: string; to: string }>;
}

const DXF_ATTRS = ['dataDxfId', 'headerRowDxfId', 'totalsRowDxfId', 'headerRowBorderDxfId', 'tableBorderDxfId', 'totalsRowBorderDxfId'];

function mapDxfAttrs(el: string, ctx: SheetObjectsContext): string {
  let out = el;
  for (const a of DXF_ATTRS) {
    const v = getAttr(out, a);
    if (v === undefined) continue;
    const nd = ctx.maps.dxf(Number(v));
    out = setAttr(out, a, nd === null ? null : String(nd));
  }
  // named cell styles belong to the source's styles.xml
  for (const a of ['headerRowCellStyle', 'dataCellStyle', 'totalsRowCellStyle']) out = setAttr(out, a, null);
  return out;
}

/** The tables the sheet's <tableParts> lists, copied. */
export async function copyTables(ctx: SheetObjectsContext, st: ObjState, stage: Stage, rids: readonly string[]): Promise<TableCopy> {
  const out: TableCopy = { parts: [], copied: 0, skipped: 0, renamed: [] };
  const plan = ctx.names.copier.opts.tables;
  const srcRels = ctx.source.rels;
  for (const rid of rids) {
    const rel = srcRels.find((r) => r.id === rid);
    if (!rel || rel.external || !rel.path) { out.skipped++; continue; }
    const raw = await ctx.source.eng.pkg.read(rel.path);
    if (!raw) { out.skipped++; continue; }
    const x = partFlavour(raw, isStrict(ctx.source.eng.wbXml), st.strict);
    const open = /<table\b[^>]*>/.exec(x)?.[0] ?? '';
    const display = getAttr(open, 'displayName') ?? getAttr(open, 'name') ?? '';
    const newDisplay = plan?.get(display.toLowerCase());
    if (newDisplay === undefined) { out.skipped++; continue; }
    const tname = getAttr(open, 'name');
    const newName = tname === undefined || tname === display ? newDisplay : (tname.toLowerCase() === newDisplay.toLowerCase() ? tname : `${tname}${newDisplay.slice(display.length)}`);
    st.tableId += 1;
    let nx = x.replace(/<table\b[^>]*>/, (o) => {
      let t = setAttr(o, 'id', String(st.tableId));
      t = setAttr(t, 'name', newName);
      t = setAttr(t, 'displayName', newDisplay);
      return mapDxfAttrs(t, ctx);
    });
    nx = nx.replace(/\sxr\d*:uid="[^"]*"/g, '');
    // columns: formats, formulas
    nx = nx.replace(/<tableColumn\b[^>]*?(?:\/>|>[\s\S]*?<\/tableColumn>)/g, (col) => {
      let c = col.replace(/^<tableColumn\b[^>]*>/, (o) => setAttr(setAttr(mapDxfAttrs(o, ctx), 'queryTableFieldId', null), 'uniqueName', null));
      let totalsBroken = false;
      // a formula that cannot be re-pointed is left out: the column is an ordinary column / has no custom total (the cells keep what they hold)
      c = c.replace(/<(calculatedColumnFormula|totalsRowFormula)(\s[^>]*)?>([\s\S]*?)<\/\1>/g, (m0, tag: string, at: string | undefined, body: string) => {
        const r = ctx.names.rewrite(unesc(body));
        if (!r.ok) { if (tag === 'totalsRowFormula') totalsBroken = true; return ''; }
        return r.changed ? `<${tag}${at ?? ''}>${esc(r.text)}</${tag}>` : m0;
      });
      if (totalsBroken) c = c.replace(/^<tableColumn\b[^>]*>/, (o) => setAttr(o, 'totalsRowFunction', 'none'));
      return c;
    });
    // style
    nx = nx.replace(/<tableStyleInfo\b[^>]*?\/>/, (el) => {
      const sn = getAttr(el, 'name');
      if (sn === undefined) return el;
      const mapped = ensureTableStyle(ctx, st, sn);
      return setAttr(el, 'name', mapped ?? null);
    });
    const newPath = freshPart(st, 'xl/tables/', 'table', '.xml');
    stage.text(newPath, nx, CT.table);
    const id = ctx.newRelId();
    out.parts.push({ rel: { id, type: relType(st, 'table'), target: newPath } });
    if (newDisplay !== display) out.renamed.push({ from: display, to: newDisplay });
    out.copied++;
  }
  return out;
}
