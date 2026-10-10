/**
 * dropOtherSheets (owner: patcher P) — PremergeOptions.dropOtherSheets: the premerged package holds ONLY the merged sheet.
 *
 * Runs inside buildPremerged right after the merged sheet's state was loaded (before any merge step), on the engine's own
 * package: the merge steps, the replay and both export engines then only ever see a one-sheet workbook. What happens (see the
 * contract comment of PremergeOptions.dropOtherSheets in contracts.ts):
 *
 *  - every other sheet (worksheets, chart sheets, dialog / macro sheets) leaves <sheets>, the workbook relationships and the
 *    package; everything that was reachable ONLY through them goes with them (drawings, charts, chart styles / embedded
 *    workbooks, pivot tables + caches, comments, VML, tables, printer settings, control properties ...): the package is walked
 *    through its relationships before and after, and what became unreachable is removed (part + rels + content-type override).
 *    Parts that were already unreachable in the input are left alone. calcChain goes, workbook-level pivot / slicer caches that
 *    nothing uses any more go, bookViews / docProps/app.xml / definedNames are made consistent;
 *  - defined names scoped to a removed sheet, or reading one (directly, through another name or through a table of a removed
 *    sheet) are removed, the rest is re-indexed (the merged sheet becomes index 0);
 *  - formulas of the merged sheet that read a removed sheet (A1 / 3D / whole column refs, names, structured refs of tables of
 *    removed sheets, INDIRECT("Sheet!..") literals) are replaced by their cached values; table calculated-column formulas that do
 *    are dropped (the column keeps its values);
 *  - data validations / conditional formats / hyperlinks / charts / sparklines / shape text links of the merged sheet that read a
 *    removed sheet are converted to a literal list (list validations whose source is a constant range of a removed sheet) or
 *    removed.
 * Reading "a removed sheet" is decided from tokens (formula/excelRef.ts tokenize), never from regexes on formula text.
 */
import { parseStructuredRef, quoteSheetName, tokenize, type Area, type Prefix } from '../../formula/excelRef';
import { formatValue } from '../../display/numfmt';
import type { CheckItem } from '../../types';
import { normalizeEol } from '../xml';
import { rewriteCounted, type Engine, type SheetState } from './engine';
import { Pkg, isRel, relsPathOf } from './pkg';
import { getSlots, removeEmptyExt, renumberPriorities } from './rules';
import { SheetModel, findCell, parseF, type CellM, type ParsedF } from './sheet';
import type { Sst } from './sst';
import type { Styles } from './styles';
import { decodeXEsc, esc, findAll, findFirst, getAttr, nsPrefix, replaceElems, setAttr, unesc, type ElemMatch } from './xml';

export interface DropResult {
  /** names of the removed sheets, in workbook order (empty = the workbook had only the merged sheet: nothing was done) */
  dropped: string[];
  /** formulas of the merged sheet replaced by their cached values because they read a removed sheet */
  detached: number;
  /** the first rows (as of the time of the drop) that held such a formula */
  detachedRows: number[];
  /** list validations whose source was a removed sheet's range: now a literal list */
  dvToList: number;
  /** validations that had to go (other types / sources that cannot be written as a literal list) */
  dvDropped: number;
  cfDropped: number;
  linksDropped: number;
  chartsDropped: number;
  sparklinesDropped: number;
  /** pivot tables of the merged sheet that were fed by a removed sheet (kept as plain values) */
  pivotsToValues: number;
  /** slicers / timelines of the merged sheet that were bound to pivot tables of removed sheets */
  slicersDropped: number;
  /** defined names removed (scoped to a removed sheet, or reading one) */
  namesRemoved: string[];
}

interface NameDef {
  name: string;
  /** localSheetId, null = workbook scope */
  local: number | null;
  hidden: boolean;
  /** refersTo (unescaped) */
  text: string;
  el: ElemMatch;
}

const ERR_MAX_EXAMPLES = 20;
const MAX_LITERAL = 255;

// ---------------------------------------------------------------------------------------------------
// "does this formula read a removed sheet?"
// ---------------------------------------------------------------------------------------------------

class Scope {
  readonly mergedLower: string;
  /** lower-case names of the removed sheets */
  readonly removed: Set<string>;
  /** all sheets in workbook order (lower case), for 3D references */
  readonly order: string[];
  /** lower-case displayName / name of the tables on removed sheets */
  readonly removedTables: Set<string>;
  /** workbook names ('g:name') and names scoped to the merged sheet ('m:name'), lower case */
  readonly names = new Map<string, NameDef>();
  /** keys of `names` that read a removed sheet */
  readonly bad = new Set<string>();
  private fragments: string[] = [];

  constructor(merged: string, removed: string[], order: string[], tables: Set<string>) {
    this.mergedLower = merged.toLowerCase();
    this.removed = new Set(removed.map((s) => s.toLowerCase()));
    this.order = order.map((s) => s.toLowerCase());
    this.removedTables = tables;
  }

  /** after `bad` is final: lower-case fragments whose presence makes a formula worth tokenizing */
  freeze(): void {
    this.fragments = [...this.bad].map((k) => k.slice(2)).concat([...this.removedTables]);
  }

  private prefixRemoved(p: Prefix): boolean {
    if (p.sheet2 !== undefined) {
      const s1 = (p as Prefix & { sheet1?: string }).sheet1;
      const a = s1 !== undefined ? s1.toLowerCase() : '';
      const b = p.sheet2.toLowerCase();
      const i = this.order.indexOf(a), j = this.order.indexOf(b);
      if (i >= 0 && j >= 0) {
        for (let k = Math.min(i, j); k <= Math.max(i, j); k++) if (this.removed.has(this.order[k])) return true;
        return false;
      }
      return this.removed.has(a) || this.removed.has(b);
    }
    return !!p.sheet && this.removed.has(p.sheet.toLowerCase());
  }

  /** a text that names a removed sheet / table the way a reference would ("Sheet!", "'My Sheet'!", "Tbl[") */
  private mentions(low: string): boolean {
    for (const s of this.removed) if (low.includes(s + '!') || low.includes(s + "'!")) return true;
    for (const t of this.removedTables) if (low.includes(t + '[')) return true;
    return false;
  }

  /**
   * `hostMerged`: unqualified references / names resolve to the merged sheet (cell formulas, rules, names scoped to it);
   * false for workbook-scope names. `quick`: skip formulas that cannot name a removed sheet (the cell loop).
   */
  reads(formula: string, hostMerged: boolean, quick = false): boolean {
    if (!formula) return false;
    const low = formula.toLowerCase();
    if (quick && !(low.includes('!') || low.includes('[') || low.includes('indirect') || this.fragments.some((f) => low.includes(f)))) return false;
    let indirect = false;
    const toks = tokenize(formula);
    for (const t of toks) {
      if (t.type === 'unknown') {
        if (this.mentions(low)) return true;
        continue;
      }
      if (t.type === 'func') {
        if (/^(?:_xlfn\.)*INDIRECT$/i.test(t.text)) indirect = true;
        if (t.prefix && t.prefix.book === undefined && this.prefixRemoved(t.prefix)) return true;
        continue;
      }
      if (t.type !== 'ref' && t.type !== 'error' && t.type !== 'name' && t.type !== 'sref') continue;
      const p = t.prefix;
      if (p) {
        if (p.book !== undefined) continue;
        if (this.prefixRemoved(p)) return true;
      }
      if (t.type === 'name') {
        const key = (p ? t.text.slice(p.text.length) : t.text).toLowerCase();
        if (p) {
          if (p.sheet && p.sheet.toLowerCase() === this.mergedLower && this.bad.has('m:' + key)) return true;
        } else {
          if (hostMerged && this.names.has('m:' + key)) {
            if (this.bad.has('m:' + key)) return true;
          } else if (this.bad.has('g:' + key)) return true;
          if (this.removedTables.has(key)) return true;
        }
      } else if (t.type === 'sref') {
        const info = parseStructuredRef(p ? t.text.slice(p.text.length) : t.text);
        if (info?.table && this.removedTables.has(info.table.toLowerCase())) return true;
      }
    }
    if (indirect) {
      for (const t of toks) {
        if (t.type !== 'string') continue;
        const body = t.text.slice(1, -1).replace(/""/g, '"').trim().toLowerCase();
        if (this.mentions(body) || this.bad.has('g:' + body) || this.bad.has('m:' + body)) return true;
      }
    }
    return false;
  }
}

// ---------------------------------------------------------------------------------------------------
// package reachability
// ---------------------------------------------------------------------------------------------------

/** every part reachable from the package root through internal relationships: lower-case name -> the name the package uses */
async function reachable(pkg: Pkg): Promise<Map<string, string>> {
  const all = new Map<string, string>();
  for (const p of pkg.list()) all.set(p.toLowerCase(), p);
  const seen = new Map<string, string>();
  const queue: string[] = [''];
  while (queue.length) {
    const owner = queue.pop()!;
    for (const r of await pkg.relTargets(owner)) {
      const key = r.path.toLowerCase();
      const canon = all.get(key);
      if (!canon || seen.has(key)) continue;
      seen.set(key, canon);
      queue.push(canon);
    }
  }
  return seen;
}

async function removePart(pkg: Pkg, path: string): Promise<void> {
  pkg.remove(path);
  pkg.remove(relsPathOf(path));
  await pkg.setContentTypeOverride(path, null);
}

// ---------------------------------------------------------------------------------------------------
// the entry point
// ---------------------------------------------------------------------------------------------------

export async function dropOtherSheets(eng: Engine, st: SheetState): Promise<DropResult> {
  const res: DropResult = {
    dropped: [], detached: 0, detachedRows: [], dvToList: 0, dvDropped: 0, cfDropped: 0, linksDropped: 0, chartsDropped: 0,
    sparklinesDropped: 0, pivotsToValues: 0, slicersDropped: 0, namesRemoved: [],
  };
  const keep = st.info;
  const others = eng.sheets.filter((s) => s !== keep);
  if (!others.length) return res;
  const pkg = eng.pkg;
  const wbPath = eng.wbPath;
  const before = await reachable(pkg);
  res.dropped = others.map((s) => s.name);

  // ---- what the removed sheets own: table names (formulas may use structured references to them), sheet ids ------------
  const removedTables = new Set<string>();
  for (const o of others) {
    if (!o.worksheet || !o.path) continue;
    for (const r of await pkg.relTargets(o.path)) {
      if (!isRel(r.type, 'table')) continue;
      const open = /<table\b[^>]*>/.exec((await pkg.read(r.path)) ?? '')?.[0];
      if (!open) continue;
      for (const k of ['name', 'displayName']) {
        const v = getAttr(open, k);
        if (v) removedTables.add(v.toLowerCase());
      }
    }
  }
  const sheetIdOf = new Map<string, string>();
  for (const m of (findFirst(eng.wbXml, 'sheets')?.inner ?? '').matchAll(/<sheet\b[^>]*>/g)) {
    const rid = getAttr(m[0], 'r:id');
    const id = getAttr(m[0], 'sheetId');
    if (rid && id) sheetIdOf.set(rid, id);
  }

  // ---- defined names -----------------------------------------------------------------------------------------------
  const scope = new Scope(keep.name, res.dropped, eng.sheets.map((s) => s.name), removedTables);
  const nameContainer = findFirst(eng.wbXml, 'definedNames');
  const allNames: NameDef[] = [];
  if (nameContainer && !nameContainer.selfClosing) {
    for (const el of findAll(nameContainer.inner, 'definedName')) {
      const ls = getAttr(el.open, 'localSheetId');
      allNames.push({ name: getAttr(el.open, 'name') ?? '', local: ls === undefined || ls === '' ? null : Number(ls), hidden: /\shidden="(1|true)"/.test(el.open), text: unesc(el.inner), el });
    }
  }
  for (const d of allNames) {
    const k = d.name.toLowerCase();
    if (d.local === null) scope.names.set('g:' + k, d);
    else if (d.local === keep.index) scope.names.set('m:' + k, d);
  }
  for (let changed = true; changed;) {
    changed = false;
    for (const [key, def] of scope.names) {
      if (scope.bad.has(key) || !scope.reads(def.text, key.startsWith('m:'))) continue;
      scope.bad.add(key);
      changed = true;
    }
  }
  scope.freeze();

  // ---- the merged sheet: formulas -------------------------------------------------------------------------------------
  const model = st.model;
  const targets: Array<{ cell: CellM; row: number; f: ParsedF }> = [];
  for (const row of model.rows) {
    for (const cell of row.cells) {
      if (!cell.x || cell.x.indexOf('<f') < 0) continue;
      const f = parseF(cell.x);
      if (!f || !f.text || /\bt\s*=\s*"dataTable"/.test(f.attrs) || !scope.reads(f.text, true, true)) continue;
      targets.push({ cell, row: row.r, f });
    }
  }
  // a spill reference (A1#) to an array that became values would turn into #REF!: it follows its array
  const masterKeys = new Set<string>();
  const mark = () => {
    for (const t of targets) if (/\bt\s*=\s*"array"/.test(t.f.attrs)) masterKeys.add(`${t.row},${t.cell.c}`);
  };
  mark();
  for (let round = 0; round < 20 && masterKeys.size; round++) {
    const had = new Set(targets.map((t) => t.cell));
    const n0 = targets.length;
    const found: typeof targets = [];
    for (const row of model.rows) {
      for (const cell of row.cells) {
        if (had.has(cell) || !cell.x || !/ANCHORARRAY/i.test(cell.x)) continue;
        const f = parseF(cell.x);
        if (!f || !f.text || !followsDetachedArray(f.text, masterKeys)) continue;
        found.push({ cell, row: row.r, f });
      }
    }
    targets.push(...found);
    mark();
    if (targets.length === n0) break;
  }
  let sst: Sst | null = null;
  if (targets.some((t) => /\st="str"/.test(t.cell.a)) || targets.some((t) => /\bt\s*=\s*"array"/.test(t.f.attrs))) sst = await eng.sst();
  for (const t of targets) {
    detachCell(t.cell, t.f, sst);
    if (/\bt\s*=\s*"array"/.test(t.f.attrs)) freezeSpilled(model, t.row, t.cell.c, getAttr(t.f.attrs, 'ref'), sst);
    if (res.detachedRows.length < ERR_MAX_EXAMPLES && !res.detachedRows.includes(t.row)) res.detachedRows.push(t.row);
  }
  res.detached = targets.length;
  res.detachedRows.sort((a, b) => a - b);
  if (targets.length) { model.dirty = true; st.changed = true; st.arrays = null; }

  // table calculated-column / totals formulas that read a removed sheet or name: the column keeps its values
  for (const t of st.tables) {
    if (t.removed) continue;
    let x = t.xml.replace(/<(calculatedColumnFormula|totalsRowFormula)(\s[^>]*)?>([\s\S]*?)<\/\1>/g, (m0, _tag: string, _at: string | undefined, body: string) => (scope.reads(unesc(body), true) ? '' : m0));
    if (x === t.xml) continue;
    const cols = findAll(x, 'tableColumn');
    x = replaceElems(x, cols, (c) => (/\stotalsRowFunction="custom"/.test(c.open) && !/<totalsRowFormula\b/.test(c.text) ? c.text.replace(/\stotalsRowFunction="custom"/, '') : c.text));
    t.xml = x;
    t.dirty = true;
  }

  // ---- the merged sheet: validations, conditional formats, links, sparklines ----------------------------------------------
  let tail = model.tail;
  const literal = makeLiteralResolver(eng, scope);
  tail = await cleanRules(tail, scope, literal, res);
  tail = rewriteCounted(tail, 'hyperlinks', 'hyperlink', (el) => {
    const loc = getAttr(el, 'location');
    if (loc && !getAttr(el, 'r:id') && scope.reads(loc, true)) { res.linksDropped++; return ''; }
    return el;
  });
  if (tail.indexOf('x14:sparkline') >= 0) {
    tail = tail.replace(/<x14:sparkline>([\s\S]*?)<\/x14:sparkline>/g, (m0, inner: string) => {
      const f = /<xm:f>([\s\S]*?)<\/xm:f>/.exec(inner);
      if (f && scope.reads(unesc(f[1]), true)) { res.sparklinesDropped++; return ''; }
      return m0;
    });
    tail = tail.replace(/<x14:sparklineGroup\b[^>]*>[\s\S]*?<\/x14:sparklineGroup>/g, (g) => (/<x14:sparkline>/.test(g) ? g : ''));
    tail = tail.replace(/<x14:sparklineGroups\b[^>]*>\s*<\/x14:sparklineGroups>/g, '');
  }
  // a data consolidation that pulls from a removed sheet
  tail = tail.replace(/<dataConsolidate\b[\s\S]*?<\/dataConsolidate>/, (m0) => {
    for (const r of m0.matchAll(/<dataRef\b[^>]*>/g)) {
      const sh = getAttr(r[0], 'sheet');
      if (sh && scope.removed.has(sh.toLowerCase())) return '';
    }
    return m0;
  });

  // ---- pivot tables of the merged sheet that are fed by a removed sheet (their cells stay as values); slicers / timelines whose
  // pivot table or table is gone ------------------------------------------------------------------------------------------------------
  const removeRids = new Set<string>(others.map((s) => s.rid));
  let wb = eng.wbXml;
  const wbRels = await pkg.relTargets(wbPath);
  const cacheTarget = new Map<string, string>();
  for (const pc of findAll(findFirst(wb, 'pivotCaches')?.inner ?? '', 'pivotCache')) {
    const id = getAttr(pc.open, 'cacheId');
    const rel = wbRels.find((r) => r.id === getAttr(pc.open, 'r:id'));
    if (id !== undefined && rel) cacheTarget.set(id, rel.path);
  }
  const usedCaches = new Set<string>();
  const keptPivots = new Set<string>();
  const deadPivotRels = new Set<string>();
  for (const r of (await pkg.relTargets(keep.path)).filter((x) => isRel(x.type, 'pivotTable'))) {
    const open = /<pivotTableDefinition\b[^>]*>/.exec((await pkg.read(r.path)) ?? '')?.[0] ?? '';
    const cid = getAttr(open, 'cacheId');
    const src = cid !== undefined ? cacheTarget.get(cid) : undefined;
    if (pivotSourceRemoved(src ? (await pkg.read(src)) ?? '' : '', scope)) {
      deadPivotRels.add(r.id);
      res.pivotsToValues++;
      continue;
    }
    if (cid !== undefined) usedCaches.add(cid);
    keptPivots.add((getAttr(open, 'name') ?? '').toLowerCase());
  }
  if (deadPivotRels.size) {
    await pkg.removeRel(keep.path, (r) => deadPivotRels.has(r.id));
    model.head = model.head.replace(/<pivotSelection\b[^>]*?(?:\/>|>[\s\S]*?<\/pivotSelection>)/g, '');
  }
  const keptTableIds = new Set<string>();
  for (const t of st.tables) if (!t.removed) keptTableIds.add(getAttr(/<table\b[^>]*>/.exec(t.xml)?.[0] ?? '', 'id') ?? '');
  const deadCaches = new Set<string>();
  for (const r of wbRels) {
    if (!isRel(r.type, 'slicerCache') && !isRel(r.type, 'timelineCache')) continue;
    const x = (await pkg.read(r.path)) ?? '';
    const name = getAttr(/<(?:[\w.-]+:)?(?:slicerCacheDefinition|timelineCacheDefinition)\b[^>]*>/.exec(x)?.[0] ?? '', 'name') ?? '';
    const pivots = [...x.matchAll(/<(?:[\w.-]+:)?pivotTable\b[^>]*>/g)].map((m) => m[0]);
    const tbl = /<(?:[\w.-]+:)?tableSlicerCache\b[^>]*>/.exec(x)?.[0];
    const stays = (p: string) => getAttr(p, 'tabId') === sheetIdOf.get(keep.rid) && keptPivots.has((getAttr(p, 'name') ?? '').toLowerCase());
    const alive = pivots.length ? pivots.some(stays) : tbl ? keptTableIds.has(getAttr(tbl, 'tableId') ?? '') : true;
    if (alive) {
      // a cache shared by several pivot tables keeps only the ones that are left
      const gone = pivots.filter((p) => !stays(p));
      if (gone.length) pkg.write(r.path, gone.reduce((acc, p) => acc.replace(p, ''), x));
      continue;
    }
    removeRids.add(r.id);
    if (name) deadCaches.add(name.toLowerCase());
    wb = wb.replace(new RegExp(`<(?:[\\w.-]+:)?(?:slicerCache|timelineCacheRef)\\b[^>]*\\sr:id="${r.id}"[^>]*/>`, 'g'), '');
  }
  const deadSlicers = new Set<string>();
  if (deadCaches.size) {
    // (Excel writes pivot slicer caches into x14:slicerCaches and table slicer caches into x15:slicerCaches)
    wb = removeEmptyExt(wb.replace(/<(x14:slicerCaches|x15:slicerCaches|x15:timelineCacheRefs)\b[^>]*>\s*<\/\1>/g, ''));
    const deadRels = new Set<string>();
    for (const r of (await pkg.relTargets(keep.path)).filter((x) => isRel(x.type, 'slicer') || isRel(x.type, 'timeline'))) {
      const x = (await pkg.read(r.path)) ?? '';
      const tag = isRel(r.type, 'slicer') ? 'slicer' : 'timeline';
      const els = findAll(x, tag);
      const dead = els.filter((e) => deadCaches.has((getAttr(e.open, 'cache') ?? '').toLowerCase()));
      if (!dead.length) continue;
      res.slicersDropped += dead.length;
      for (const e of dead) { const n = getAttr(e.open, 'name'); if (n) deadSlicers.add(n); }
      if (dead.length === els.length) deadRels.add(r.id);
      else pkg.write(r.path, replaceElems(x, dead, () => ''));
    }
    if (deadRels.size) {
      await pkg.removeRel(keep.path, (r) => deadRels.has(r.id));
      for (const id of deadRels) tail = tail.replace(new RegExp(`<(?:[\\w.-]+:)?(?:slicer|timelineRef)\\b[^>]*\\sr:id="${id}"[^>]*/>`, 'g'), '');
      tail = tail.replace(/<(x14:slicerList|x15:timelineRefs)\b[^>]*>\s*<\/\1>/g, '');
    }
  }
  tail = removeEmptyExt(tail);
  if (tail !== model.tail) model.tail = tail;
  // the merged sheet is the only (and the selected) tab
  model.head = model.head.replace(/<sheetView\b[^>]*>/, (o) => setAttr(o, 'tabSelected', '1'));
  model.dirty = true;
  st.changed = true;

  // ---- the merged sheet: charts / shape links of its drawings ------------------------------------------------------------------
  await cleanDrawings(pkg, keep.path, scope, res, deadSlicers);
  await cleanControls(pkg, keep.path, scope);

  // ---- workbook.xml ------------------------------------------------------------------------------------------------------------
  // <sheets>: only the merged sheet, always visible
  {
    const el = findFirst(wb, 'sheets');
    if (el && !el.selfClosing) {
      const mine = [...el.inner.matchAll(/<sheet\b[^>]*?\/>|<sheet\b[^>]*?>[\s\S]*?<\/sheet>/g)].map((m) => m[0]).find((s) => getAttr(s, 'r:id') === keep.rid);
      if (mine) wb = wb.slice(0, el.start) + `${el.open}${setAttr(mine, 'state', null)}</sheets>` + wb.slice(el.end);
    }
  }
  // bookViews: the first tab is shown
  wb = wb.replace(/<workbookView\b[^>]*>/g, (o) => setAttr(setAttr(o, 'activeTab', null), 'firstSheet', null));
  // pivot caches: only those a remaining pivot table uses
  {
    const el = findFirst(wb, 'pivotCaches');
    if (el && !el.selfClosing) {
      const kept: string[] = [];
      for (const pc of findAll(el.inner, 'pivotCache')) {
        const id = getAttr(pc.open, 'cacheId');
        const rid = getAttr(pc.open, 'r:id');
        if (id !== undefined && usedCaches.has(id)) kept.push(pc.text);
        else if (rid) removeRids.add(rid);
      }
      wb = wb.slice(0, el.start) + (kept.length ? `${el.open}${kept.join('')}</pivotCaches>` : '') + wb.slice(el.end);
    }
  }
  // defined names
  const keptNames: NameDef[] = [];
  for (const d of allNames) {
    const k = d.name.toLowerCase();
    const dead = d.local !== null && d.local !== keep.index
      ? true
      : (d.local === null ? scope.bad.has('g:' + k) : scope.bad.has('m:' + k)) || deadCaches.has(k);
    if (dead) res.namesRemoved.push(d.local !== null && d.local !== keep.index ? `${d.name}（${eng.sheets[d.local]?.name ?? d.local}）` : d.name);
    else keptNames.push(d);
  }
  {
    const el = findFirst(wb, 'definedNames');
    if (el) {
      const body = keptNames.map((d) => (d.local === null ? d.el.text : setAttr(d.el.open, 'localSheetId', '0') + (d.el.selfClosing ? '' : `${d.el.inner}</definedName>`))).join('');
      wb = wb.slice(0, el.start) + (keptNames.length ? `${el.open}${body}</definedNames>` : '') + wb.slice(el.end);
    }
  }
  // the dynamic-array metadata is only needed by cells that use it
  if (!model.rows.some((row) => row.cells.some((c) => /\s(?:cm|vm)="/.test(c.a)))) {
    for (const r of wbRels) if (isRel(r.type, 'sheetMetadata')) removeRids.add(r.id);
  }
  for (const r of wbRels) if (isRel(r.type, 'calcChain')) removeRids.add(r.id);
  eng.wbXml = wb;
  await pkg.removeRel(wbPath, (r) => removeRids.has(r.id));

  // ---- parts nothing refers to any more ---------------------------------------------------------------------------------------------
  let after = await reachable(pkg);
  // the list of people of threaded comments goes with the last threaded comment
  if (![...after.keys()].some((p) => /(^|\/)threadedcomments\//.test(p))) {
    const persons = (await pkg.relTargets(wbPath)).filter((r) => isRel(r.type, 'person'));
    if (persons.length) {
      const ids = new Set(persons.map((r) => r.id));
      await pkg.removeRel(wbPath, (r) => ids.has(r.id));
      after = await reachable(pkg);
    }
  }
  for (const [key, path] of before) if (!after.has(key)) await removePart(pkg, path);

  // ---- docProps/app.xml ----------------------------------------------------------------------------------------------------------------
  await fixAppProps(pkg, keep.name, allNames, keptNames, eng.sheets);

  // ---- the engine now sees one sheet ---------------------------------------------------------------------------------------------------------
  keep.index = 0;
  eng.sheets = [keep];
  for (const k of [...eng.states.keys()]) if (k !== keep.path) eng.states.delete(k);
  return res;
}

// ---------------------------------------------------------------------------------------------------
// cells
// ---------------------------------------------------------------------------------------------------

/** A1# (_xlfn.ANCHORARRAY(B2)) of a cell in `masters` ("row,col" keys of arrays that became values) */
function followsDetachedArray(f: string, masters: Set<string>): boolean {
  const toks = tokenize(f);
  for (let i = 0; i < toks.length; i++) {
    const t = toks[i];
    if (t.type !== 'func' || !/^(?:_xlfn\.)*ANCHORARRAY$/i.test(t.text)) continue;
    for (let j = i + 1; j < toks.length && j < i + 4; j++) {
      const r = toks[j];
      if (r.type !== 'ref' || !r.area || r.prefix) continue;
      if (masters.has(`${r.area.r1},${r.area.c1}`)) return true;
    }
  }
  return false;
}

/** the formula of `cell` goes, its cached result stays as the value (a string result becomes a shared string); also used by the sheet copier */
export function detachCell(cell: CellM, f: ParsedF, sst: Sst | null): void {
  const x = cell.x ?? '';
  const rest = x.slice(0, f.start) + x.slice(f.end);
  let a = setAttr(cell.a, 'cm', null);
  const vm = /<v>([\s\S]*?)<\/v>/.exec(rest);
  if (getAttr(a, 't') === 'str') {
    if (vm && sst) {
      a = setAttr(a, 't', 's');
      cell.x = `<v>${sst.add(decodeXEsc(unesc(normalizeEol(vm[1]))))}</v>`;
    } else {
      a = setAttr(a, 't', null);
      cell.x = null;
    }
  } else if (!/<v[\s>]|<is>/.test(rest)) {
    a = setAttr(a, 't', null);
    cell.x = null;
  } else cell.x = rest;
  cell.a = a;
}

/** the cells an array formula spilled into hold their results as constants; a formula-string result becomes an ordinary string */
function freezeSpilled(model: SheetModel, row: number, col: number, ref: string | undefined, sst: Sst | null): void {
  const m = ref ? /^\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?$/.exec(ref) : null;
  if (!m || !sst) return;
  const colNum = (s: string) => { let n = 0; for (const ch of s) n = n * 26 + ch.charCodeAt(0) - 64; return n; };
  const c1 = colNum(m[1]), r1 = Number(m[2]), c2 = m[3] ? colNum(m[3]) : c1, r2 = m[4] ? Number(m[4]) : r1;
  for (const r of model.rows) {
    if (r.r < r1 || r.r > r2) continue;
    for (const c of r.cells) {
      if (c.c < c1 || c.c > c2 || (r.r === row && c.c === col) || !c.x || c.x.indexOf('<f') >= 0 || getAttr(c.a, 't') !== 'str') continue;
      const v = /<v>([\s\S]*?)<\/v>/.exec(c.x);
      c.a = setAttr(c.a, 't', v ? 's' : null);
      c.x = v ? `<v>${sst.add(decodeXEsc(unesc(normalizeEol(v[1]))))}</v>` : null;
    }
  }
}

// ---------------------------------------------------------------------------------------------------
// validations / conditional formats
// ---------------------------------------------------------------------------------------------------

type LiteralResolver = (formula1: string) => Promise<string | null>;

/**
 * A list validation's source as a literal list: the source (a range of a removed sheet, directly or through a name) must be one
 * column or one row of constant cells (cached values); the items must not contain commas / quotes and the whole list must fit
 * Excel's 255 characters. Anything else: null (the validation is dropped).
 */
function makeLiteralResolver(eng: Engine, scope: Scope): LiteralResolver {
  const models = new Map<string, SheetModel | null>();
  let sst: Sst | null = null;
  let styles: Styles | null = null;
  const sheetModel = async (name: string): Promise<SheetModel | null> => {
    const key = name.toLowerCase();
    if (models.has(key)) return models.get(key)!;
    const info = eng.sheets.find((s) => s.name.toLowerCase() === key && s.worksheet && s.path);
    let m: SheetModel | null = null;
    try {
      const xml = info ? await eng.pkg.read(info.path) : null;
      m = xml ? new SheetModel(xml) : null;
    } catch {
      m = null; // a damaged sheet that is dropped anyway: only a validation that reads it suffers (it is dropped)
    }
    models.set(key, m);
    return m;
  };
  const itemsOf = async (sheet: string, a: Area): Promise<string[] | null> => {
    const m = await sheetModel(sheet);
    if (!m) return null;
    const r1 = Math.min(a.r1, a.r2), r2 = Math.max(a.r1, a.r2), c1 = Math.min(a.c1, a.c2), c2 = Math.max(a.c1, a.c2);
    const wholeCols = a.kind === 'cols', wholeRows = a.kind === 'rows';
    const colVec = wholeCols ? c1 === c2 : !wholeRows && c1 === c2;
    const rowVec = wholeRows ? r1 === r2 : !wholeCols && r1 === r2;
    if (!colVec && !rowVec) return null;
    sst ??= await eng.sst();
    styles ??= await eng.styles();
    const out: string[] = [];
    const text = (cell: CellM): string | null => {
      const t = getAttr(cell.a, 't');
      const x = cell.x ?? '';
      const v = /<v>([\s\S]*?)<\/v>/.exec(x)?.[1];
      if (t === 'inlineStr') { const is = /<is>([\s\S]*?)<\/is>/.exec(x); return is ? siTextOf(is[1]) : ''; }
      if (v === undefined) return '';
      if (t === 's') { const i = Number(v); return Number.isInteger(i) ? sst!.text(i) : ''; }
      if (t === 'str') return decodeXEsc(unesc(normalizeEol(v)));
      if (t === 'b') return v.trim() === '1' || /^true$/i.test(v.trim()) ? 'TRUE' : 'FALSE';
      if (t === 'e') return null;
      const n = Number(v);
      if (!Number.isFinite(n)) return unesc(v);
      const xf = Number(getAttr(cell.a, 's') ?? 0) || 0;
      const id = Number(getAttr(/^<[^>]*>/.exec(styles!.xf(xf))?.[0] ?? '', 'numFmtId') ?? 0) || 0;
      return formatValue(n, styles!.numFmtCode(id) ?? id).text;
    };
    if (colVec) {
      for (const row of m.rows) {
        if (row.r < r1 || row.r > r2) continue;
        const cell = findCell(row, c1);
        if (!cell) continue;
        const s = text(cell);
        if (s === null) return null;
        if (s !== '') out.push(s);
      }
    } else {
      const row = m.row(r1);
      for (const cell of row?.cells ?? []) {
        if (cell.c < c1 || cell.c > c2) continue;
        const s = text(cell);
        if (s === null) return null;
        if (s !== '') out.push(s);
      }
    }
    return out;
  };
  const resolve = async (formula: string, depth: number): Promise<string[] | null> => {
    const toks = tokenize(formula).filter((t) => t.type !== 'ws');
    if (toks.length !== 1 || depth > 4) return null;
    const t = toks[0];
    if (t.type === 'name' && !t.prefix) {
      const k = t.text.toLowerCase();
      const def = scope.names.get('m:' + k) ?? scope.names.get('g:' + k);
      return def ? resolve(def.text, depth + 1) : null;
    }
    const p = t.prefix;
    if (t.type === 'ref' && t.area && p && p.book === undefined && p.sheet2 === undefined && p.sheet && scope.removed.has(p.sheet.toLowerCase())) return itemsOf(p.sheet, t.area);
    return null;
  };
  return async (formula1) => {
    const items = await resolve(formula1, 0);
    if (!items || !items.length) return null;
    if (items.some((s) => s.includes(',') || s.includes('"') || /[\r\n]/.test(s))) return null;
    const joined = items.join(',');
    return joined.length <= MAX_LITERAL ? joined : null;
  };
}

function siTextOf(inner: string): string {
  let out = '';
  for (const m of inner.replace(/<rPh\b[\s\S]*?<\/rPh>/g, '').matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g)) out += m[1];
  return decodeXEsc(unesc(normalizeEol(out)));
}

const AFTER_DV = /<(?:hyperlinks|printOptions|pageMargins|pageSetup|headerFooter|rowBreaks|colBreaks|customProperties|cellWatches|ignoredErrors|smartTags|drawing|legacyDrawing|legacyDrawingHF|drawingHF|picture|oleObjects|controls|webPublishItems|tableParts|extLst)(?=[\s/>])/;

const splice = (x: string, el: ElemMatch, repl: string) => x.slice(0, el.start) + repl + x.slice(el.end);

async function cleanRules(tail: string, scope: Scope, literal: LiteralResolver, res: DropResult): Promise<string> {
  let x = tail;
  // ---- data validations ---------------------------------------------------------------------------------------------
  const classicEl = findFirst(x, 'dataValidations');
  const x14El = findFirst(x, 'x14:dataValidations');
  const unitsOf = (el: ElemMatch | null, tag: string) => (el && !el.selfClosing ? findAll(el.inner, tag) : []);
  const keptClassic: string[] = [];
  const keptX14: string[] = [];
  const toClassic: string[] = [];
  let dvTouched = false;
  const decide = async (unit: ElemMatch, x14: boolean): Promise<string | null> => {
    const slots = getSlots(unit.text, x14 ? 'x14dv' : 'dv');
    if (!slots.some((s) => scope.reads(s, true))) return unit.text;
    dvTouched = true;
    if (getAttr(unit.open, 'type') !== 'list') { res.dvDropped++; return null; }
    const lit = await literal(slots[0] ?? '');
    if (lit === null) { res.dvDropped++; return null; }
    res.dvToList++;
    const f1 = `<formula1>${esc(`"${lit}"`)}</formula1>`;
    if (!x14) return unit.text.replace(/<formula1>[\s\S]*?<\/formula1>/, f1).replace(/<formula2>[\s\S]*?<\/formula2>/, '');
    const sq = /<xm:sqref>([\s\S]*?)<\/xm:sqref>/.exec(unit.text)?.[1] ?? '';
    const attrs = unit.open.replace(/^<x14:dataValidation/, '').replace(/\/?>$/, '').replace(/\sxr:uid="[^"]*"/g, '');
    toClassic.push(`<dataValidation${attrs} sqref="${sq}">${f1}</dataValidation>`);
    return null;
  };
  for (const u of unitsOf(classicEl, 'dataValidation')) { const r = await decide(u, false); if (r !== null) keptClassic.push(r); }
  for (const u of unitsOf(x14El, 'x14:dataValidation')) { const r = await decide(u, true); if (r !== null) keptX14.push(r); }
  if (dvTouched) {
    const classicAll = keptClassic.concat(toClassic);
    if (classicEl) {
      x = splice(x, classicEl, classicAll.length ? `${setAttr(classicEl.open, 'count', String(classicAll.length))}${classicAll.join('')}</dataValidations>` : '');
    } else if (classicAll.length) {
      const at = AFTER_DV.exec(x)?.index ?? Math.max(0, x.lastIndexOf('</worksheet>'));
      x = x.slice(0, at) + `<dataValidations count="${classicAll.length}">${classicAll.join('')}</dataValidations>` + x.slice(at);
    }
    const x14Now = findFirst(x, 'x14:dataValidations');
    if (x14Now) x = splice(x, x14Now, keptX14.length ? `${setAttr(x14Now.open, 'count', String(keptX14.length))}${keptX14.join('')}</x14:dataValidations>` : '');
  }

  // ---- conditional formatting (classic + x14); a classic rule and its x14 extension (<x14:id>) live and die together -----------
  interface RuleInfo { id: string | null; own: boolean }
  const classicRules: RuleInfo[] = [];
  const x14Rules: RuleInfo[] = [];
  for (const b of findAll(x, 'conditionalFormatting')) {
    for (const r of findAll(b.inner, 'cfRule')) {
      classicRules.push({ id: /<x14:id>([^<]*)<\/x14:id>/.exec(r.text)?.[1] ?? null, own: getSlots(r.text, 'cf').some((s) => scope.reads(s, true)) });
    }
  }
  for (const b of findAll(x, 'x14:conditionalFormatting')) {
    for (const r of findAll(b.inner, 'x14:cfRule')) {
      x14Rules.push({ id: getAttr(r.open, 'id') ?? null, own: getSlots(r.text, 'x14cf').some((s) => scope.reads(s, true)) });
    }
  }
  const ownBadX14 = new Set(x14Rules.filter((r) => r.own && r.id).map((r) => r.id!));
  const classicBad = classicRules.map((r) => r.own || (!!r.id && ownBadX14.has(r.id)));
  const linkedBad = new Set(classicRules.filter((_r, i) => classicBad[i]).map((r) => r.id).filter((s): s is string => !!s));
  const x14Bad = x14Rules.map((r) => r.own || (!!r.id && linkedBad.has(r.id)));
  if (classicBad.some(Boolean) || x14Bad.some(Boolean)) {
    res.cfDropped += classicBad.filter(Boolean).length + x14Rules.filter((r, i) => x14Bad[i] && !(r.id && linkedBad.has(r.id))).length;
    const dropBlocks = (xml: string, blockTag: string, ruleTag: string, flags: boolean[], counter: { i: number }): string => {
      const blocks = findAll(xml, blockTag);
      return replaceElems(xml, blocks, (b) => {
        const rules = findAll(b.inner, ruleTag);
        let left = 0;
        const inner = replaceElems(b.inner, rules, (r) => {
          if (flags[counter.i++]) return '';
          left++;
          return r.text;
        });
        return left ? `${b.open}${inner}</${blockTag}>` : '';
      });
    };
    x = dropBlocks(x, 'conditionalFormatting', 'cfRule', classicBad, { i: 0 });
    x = dropBlocks(x, 'x14:conditionalFormatting', 'x14:cfRule', x14Bad, { i: 0 });
    x = x.replace(/<x14:conditionalFormattings\b[^>]*>\s*<\/x14:conditionalFormattings>/g, '');
    x = renumberPriorities(removeEmptyExt(x));
  }
  return x;
}

// ---------------------------------------------------------------------------------------------------
// drawings of the merged sheet
// ---------------------------------------------------------------------------------------------------

const XDR_NS = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
const MC_NS = 'http://schemas.openxmlformats.org/markup-compatibility/2006';

/**
 * Remove the anchors of a drawing part whose text satisfies `dead` (a chart whose relationship id is dead ...), together with an
 * mc:AlternateContent wrapper that sits around anchors (slicers / timelines). Returns the new XML.
 */
export function removeDrawingObjects(xml: string, dead: (text: string) => boolean): string {
  const xp = nsPrefix(xml, XDR_NS, 'xdr:');
  const mp = nsPrefix(xml, MC_NS, 'mc:');
  const anchors = ['twoCellAnchor', 'oneCellAnchor', 'absoluteAnchor'].flatMap((n) => findAll(xml, xp + n));
  const inAnchor = (pos: number) => anchors.some((a) => pos >= a.start && pos < a.end);
  let out = xml;
  // wrappers around anchors (outside every anchor)
  const wrappers = findAll(out, mp + 'AlternateContent').filter((w) => !inAnchor(w.start) && dead(w.text));
  out = replaceElems(out, wrappers, () => '');
  for (const n of ['twoCellAnchor', 'oneCellAnchor', 'absoluteAnchor']) {
    out = replaceElems(out, findAll(out, xp + n).filter((a) => dead(a.text)), () => '');
  }
  return out;
}

async function cleanDrawings(pkg: Pkg, sheetPath: string, scope: Scope, res: DropResult, deadSlicers: Set<string>): Promise<void> {
  for (const dr of (await pkg.relTargets(sheetPath)).filter((r) => isRel(r.type, 'drawing'))) {
    const xml = await pkg.read(dr.path);
    if (xml == null) continue;
    const deadRids = new Set<string>();
    for (const cr of (await pkg.relTargets(dr.path)).filter((r) => isRel(r.type, 'chart') || isRel(r.type, 'chartEx'))) {
      const cx = await pkg.read(cr.path);
      if (cx == null) continue;
      for (const m of cx.matchAll(/<((?:[\w.-]+:)?f)(\s[^>]*)?>([^<]*)<\/\1>/g)) {
        if (scope.reads(unesc(m[3]), true)) { deadRids.add(cr.id); break; }
      }
    }
    let y = xml;
    if (deadRids.size) {
      y = removeDrawingObjects(y, (t) => [...t.matchAll(/\sr:id="([^"]*)"/g)].some((m) => deadRids.has(m[1])));
      await pkg.removeRel(dr.path, (r) => deadRids.has(r.id));
      res.chartsDropped += deadRids.size;
    }
    // slicers / timelines bound to a pivot table that is gone
    if (deadSlicers.size) y = removeDrawingObjects(y, (t) => [...t.matchAll(/<(?:[\w.-]+:)?(?:slicer|timeslicer)\b[^>]*\sname="([^"]*)"/g)].some((m) => deadSlicers.has(unesc(m[1]))));
    // text boxes linked to a cell of a removed sheet keep their text, the link goes
    y = y.replace(/\stextlink="([^"]*)"/g, (m0, v: string) => (scope.reads(unesc(v), true) ? ' textlink=""' : m0));
    if (y !== xml) pkg.write(dr.path, y);
  }
}

/** form controls of the merged sheet whose cell link / list range lies on a removed sheet are left unlinked */
async function cleanControls(pkg: Pkg, sheetPath: string, scope: Scope): Promise<void> {
  for (const r of await pkg.relTargets(sheetPath)) {
    if (isRel(r.type, 'ctrlProp')) {
      const x = await pkg.read(r.path);
      if (x == null) continue;
      const y = x.replace(/\s(fmlaLink|fmlaRange|fmlaTxbx|fmlaGroup)="([^"]*)"/g, (m0, _n: string, v: string) => (scope.reads(unesc(v), true) ? '' : m0));
      if (y !== x) pkg.write(r.path, y);
    } else if (isRel(r.type, 'vmlDrawing')) {
      const x = await pkg.read(r.path);
      if (x == null) continue;
      const xp = nsPrefix(x, 'urn:schemas-microsoft-com:office:excel', 'x:');
      if (x.indexOf(`<${xp}Fmla`) < 0) continue;
      const y = x.replace(new RegExp(`<${xp}(Fmla(?:Link|Range|Txbx|Group))>([\\s\\S]*?)</${xp}\\1>`, 'g'), (m0, _n: string, v: string) => (scope.reads(unesc(v), true) ? '' : m0));
      if (y !== x) pkg.write(r.path, y);
    }
  }
}

/** the data source of a pivot cache reads a removed sheet (a range, a table, a name, or a consolidation range set) */
function pivotSourceRemoved(cache: string, scope: Scope): boolean {
  const ws = /<worksheetSource\b[^>]*>/.exec(cache)?.[0];
  if (ws) {
    if (getAttr(ws, 'r:id')) return false;
    const sheet = getAttr(ws, 'sheet');
    if (sheet && scope.removed.has(sheet.toLowerCase())) return true;
    const name = getAttr(ws, 'name')?.toLowerCase();
    return !!name && (scope.removedTables.has(name) || scope.bad.has('g:' + name) || scope.bad.has('m:' + name));
  }
  for (const m of cache.matchAll(/<rangeSet\b[^>]*>/g)) {
    const sh = getAttr(m[0], 'sheet');
    if (sh && scope.removed.has(sh.toLowerCase())) return true;
  }
  return false;
}

// ---------------------------------------------------------------------------------------------------
// docProps/app.xml
// ---------------------------------------------------------------------------------------------------

const VT_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes';

function nameTitle(d: NameDef, sheets: Array<{ name: string }>): string {
  const disp = d.name.replace(/^_xlnm\./, '');
  if (d.local === null) return disp;
  const sh = sheets[d.local]?.name;
  return sh === undefined ? disp : `${quoteSheetName(sh)}!${disp}`;
}

const listedName = (d: NameDef) => !d.hidden && d.name !== '_xlnm._FilterDatabase';

/**
 * HeadingPairs / TitlesOfParts name the sheets and the named ranges: rewritten for the one sheet and the names that are left (the
 * group of the named ranges keeps its label and order; groups of other kinds — chart sheets, macros — belonged to removed sheets).
 * Anything unexpected: both lists are removed (they are optional; Excel rebuilds them on save).
 */
async function fixAppProps(pkg: Pkg, keepName: string, allNames: NameDef[], keptNames: NameDef[], sheets: Array<{ name: string }>): Promise<void> {
  const rel = (await pkg.relTargets('')).find((r) => isRel(r.type, 'extended-properties'));
  if (!rel) return;
  const xml = await pkg.read(rel.path);
  if (!xml) return;
  const hp = /<HeadingPairs>([\s\S]*?)<\/HeadingPairs>/.exec(xml);
  const tp = /<TitlesOfParts>([\s\S]*?)<\/TitlesOfParts>/.exec(xml);
  if (!hp && !tp) return;
  const strip = () => pkg.write(rel.path, xml.replace(/<HeadingPairs>[\s\S]*?<\/HeadingPairs>/, '').replace(/<TitlesOfParts>[\s\S]*?<\/TitlesOfParts>/, ''));
  const prefix = new RegExp(`xmlns:([\\w.-]+)="${VT_NS.replace(/\./g, '\\.')}"`).exec(xml)?.[1];
  if (!hp || !tp || !prefix) { strip(); return; }
  const vt = prefix + ':';
  const lp = new RegExp(`<${vt}lpstr>([\\s\\S]*?)</${vt}lpstr>`, 'g');
  const i4 = new RegExp(`<${vt}i4>\\s*(\\d+)\\s*</${vt}i4>`, 'g');
  const labels = [...hp[1].matchAll(lp)].map((m) => unesc(m[1]));
  const counts = [...hp[1].matchAll(i4)].map((m) => Number(m[1]));
  const titles = [...tp[1].matchAll(lp)].map((m) => unesc(m[1]));
  if (!labels.length || labels.length !== counts.length || counts.reduce((a, b) => a + b, 0) !== titles.length) { strip(); return; }
  const everyTitle = new Set(allNames.filter(listedName).map((d) => nameTitle(d, sheets)));
  const keptTitles = new Set(keptNames.filter(listedName).map((d) => nameTitle(d, sheets)));
  const groups: Array<{ label: string; titles: string[] }> = [{ label: labels[0], titles: [keepName] }];
  let at = counts[0];
  for (let g = 1; g < labels.length; g++) {
    const part = titles.slice(at, at + counts[g]);
    at += counts[g];
    if (!part.some((t) => everyTitle.has(t))) continue;
    const left = part.filter((t) => keptTitles.has(t));
    if (left.length) groups.push({ label: labels[g], titles: left });
  }
  const variants = groups.map((g) => `<${vt}variant><${vt}lpstr>${esc(g.label)}</${vt}lpstr></${vt}variant><${vt}variant><${vt}i4>${g.titles.length}</${vt}i4></${vt}variant>`).join('');
  const all = groups.flatMap((g) => g.titles);
  const heading = `<HeadingPairs><${vt}vector size="${groups.length * 2}" baseType="variant">${variants}</${vt}vector></HeadingPairs>`;
  const parts = `<TitlesOfParts><${vt}vector size="${all.length}" baseType="lpstr">${all.map((t) => `<${vt}lpstr>${esc(t)}</${vt}lpstr>`).join('')}</${vt}vector></TitlesOfParts>`;
  pkg.write(rel.path, xml.replace(/<HeadingPairs>[\s\S]*?<\/HeadingPairs>/, heading).replace(/<TitlesOfParts>[\s\S]*?<\/TitlesOfParts>/, parts));
}

// ---------------------------------------------------------------------------------------------------
// report
// ---------------------------------------------------------------------------------------------------

/**
 * The CheckItems (severity 'info', Japanese) that tell the user what the single-sheet output left out. `mapRow` converts a row of
 * the base sheet at the time of the drop to its row in the premerged sheet (appended rows move what lies below the table).
 */
export function dropCheckItems(r: DropResult, mapRow: (row: number) => number = (x) => x): CheckItem[] {
  const out: CheckItem[] = [];
  const info = (code: string, message: string, rows?: number[]) => out.push({ severity: 'info', code, message, ...(rows ? { rows } : {}) });
  if (r.dropped.length) {
    const names = r.dropped.slice(0, 5).map((n) => `「${n}」`).join('');
    info('sheetsDropped', `他のシート ${r.dropped.length} 枚（${names}${r.dropped.length > 5 ? ' ほか' : ''}）は出力に含めません。`);
  }
  if (r.detached) info('detachedFormulas', `他のシートを参照していた数式 ${r.detached} 個を値にしました`, r.detachedRows.map(mapRow));
  if (r.dvToList) info('dvToList', `入力規則 ${r.dvToList} 個を固定のリストにしました`);
  if (r.dvDropped) info('dvDropped', `他のシートを参照していた入力規則 ${r.dvDropped} 個を外しました`);
  if (r.cfDropped) info('cfDropped', `条件付き書式 ${r.cfDropped} 個を外しました`);
  if (r.linksDropped) info('linksDropped', `他のシートへのリンク ${r.linksDropped} 個を外しました`);
  if (r.chartsDropped) info('chartsDropped', `他のシートのデータを使うグラフ ${r.chartsDropped} 個は含めません`);
  if (r.sparklinesDropped) info('sparklinesDropped', `他のシートのデータを使うスパークライン ${r.sparklinesDropped} 個を外しました`);
  if (r.pivotsToValues) info('pivotsToValues', `他のシートのデータを使うピボットテーブル ${r.pivotsToValues} 個を値にしました`);
  if (r.slicersDropped) info('slicersDropped', `他のシートの表・ピボットテーブルに結び付いていたスライサー ${r.slicersDropped} 個を外しました`);
  return out;
}
