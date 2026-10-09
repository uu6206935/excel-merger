/**
 * Node-side XML snapshot of an .xlsx package, used to compare replayLog() output with the workbook Excel itself
 * saved after the same ops (golden fixtures). Everything is expressed in a form that does not depend on how a
 * writer numbers its records:
 *  - cell formulas (shared formulas expanded), non-formula cell values (shared strings resolved),
 *  - cell styles as resolved signatures (numFmt code + font + fill + border + alignment XML), not xf indexes,
 *  - CF / DV as per-cell effective rules (formula re-expressed relative to the cell: R1C1-like key),
 *  - names, tables, chart <c:f>, pivot sources, merges, column widths, frozen panes, autoFilter, hyperlinks,
 *    notes / threaded comments, drawing anchors.
 */
import JSZip from 'jszip';
import { colToNum, numToCol, quoteSheetName, relKey, sqrefTopLeft, tokenize, translateFormula } from '../../src/core/formula/excelRef';
import { parseSqref } from '../../src/core/xlsx/patch/formula';
import { SheetModel, parseF } from '../../src/core/xlsx/patch/sheet';
import { siText } from '../../src/core/xlsx/patch/sst';
import { decodeBytes, findAll, findFirst, getAttr, normalizeAttrQuotes, stripMainPrefix, unesc } from '../../src/core/xlsx/patch/xml';

export interface SheetSnap {
  name: string;
  formulas: Record<string, string>;
  values: Record<string, string>;
  styles: Record<string, string>;
  /** per-cell style signature with shared border edges merged (see edgeStyles); what diffSnap compares as `styles` */
  edgeStyles: Record<string, string>;
  merges: string[];
  cf: string[];
  dv: string[];
  cols: Record<string, string>;
  rows: Record<string, string>;
  pane: string;
  autoFilter: string;
  hyperlinks: string[];
  comments: string[];
  threaded: string[];
  anchors: string[];
  tables: string[];
  sparklines: string[];
  dimension: string;
}

export interface BookSnap {
  sheets: SheetSnap[];
  names: string[];
  charts: string[];
  pivots: string[];
  calcChain: boolean;
  fullCalcOnLoad: boolean;
}

/** limits for expanding sqrefs into cells */
const MAX_R = 400, MAX_C = 60;

function resolve(baseDir: string, target: string): string {
  target = target.replace(/\\/g, '/'); // some generators write Windows separators; Excel reads them as '/'
  if (target.startsWith('/')) return target.slice(1);
  const out: string[] = [];
  for (const p of (baseDir ? baseDir.split('/') : []).concat(target.split('/'))) {
    if (p === '..') out.pop(); else if (p && p !== '.') out.push(p);
  }
  return out.join('/');
}
const dirOf = (p: string) => p.slice(0, Math.max(0, p.lastIndexOf('/')));
const relsOf = (p: string) => `${dirOf(p)}/_rels/${p.slice(p.lastIndexOf('/') + 1)}.rels`;

export async function snapshot(bytes: Uint8Array): Promise<BookSnap> {
  const zip = await JSZip.loadAsync(bytes);
  const read = async (p: string) => {
    const f = zip.file(p);
    if (!f) return null;
    const t = decodeBytes(await f.async('uint8array')).text;
    // parts written with a namespace prefix (<x:row>) or with single-quoted attributes are compared in the ordinary form
    let q = /\.(xml|rels)$/i.test(p) ? normalizeAttrQuotes(t) : t;
    // (and pretty-printed ones: white space between elements is not content; the text of a <t> is, so it is left alone)
    if (/\.(xml|rels)$/i.test(p) && !/sharedStrings/i.test(p)) q = q.replace(/(?<!<t(?:\s[^<>]*)?)>\s+<(?!\/t>)/g, '><');
    return /\.xml$/i.test(p) ? stripMainPrefix(q) ?? q : q;
  };
  const rels = async (p: string) => {
    const x = await read(relsOf(p));
    if (!x) return [] as Array<{ id: string; type: string; path: string; external: boolean }>;
    return [...x.matchAll(/<Relationship\b[^>]*>/g)].map((m) => ({
      id: getAttr(m[0], 'Id') ?? '',
      type: (getAttr(m[0], 'Type') ?? '').split('/').pop() ?? '',
      path: resolve(dirOf(p), getAttr(m[0], 'Target') ?? ''),
      external: /TargetMode="External"/.test(m[0]),
    }));
  };
  const wbPath = 'xl/workbook.xml';
  const wb = (await read(wbPath)) ?? '';
  const wbRels = await rels(wbPath);
  const sstPath = wbRels.find((r) => r.type === 'sharedStrings')?.path;
  const sstXml = sstPath ? await read(sstPath) : null;
  const sst: string[] = [];
  if (sstXml) for (const m of sstXml.matchAll(/<si\s*\/>|<si(?:\s[^>]*)?>([\s\S]*?)<\/si>/g)) sst.push(m[1] === undefined ? '' : siText(m[1]));
  const stylesPath = wbRels.find((r) => r.type === 'styles')?.path;
  const stylesParts = styleParts((stylesPath ? await read(stylesPath) : null) ?? '');
  const styles = stylesParts.map((p) => `${p.pre}border=${p.borderXml}${p.post}`);

  const sheetNames: string[] = [];
  const snap: BookSnap = { sheets: [], names: [], charts: [], pivots: [], calcChain: !!zip.file('xl/calcChain.xml'), fullCalcOnLoad: /<calcPr\b[^>]*fullCalcOnLoad="1"/.test(wb) };
  for (const m of wb.matchAll(/<sheet\b[^>]*>/g)) sheetNames.push(getAttr(m[0], 'name') ?? '');
  for (const m of wb.matchAll(/<definedName\b([^>]*)>([\s\S]*?)<\/definedName>/g)) {
    const lsi = getAttr(m[1], 'localSheetId');
    const scope = lsi != null ? sheetNames[Number(lsi)] + '!' : '';
    snap.names.push(`${scope}${getAttr(m[1], 'name')} = ${normQuotes(unesc(m[2]))}`);
  }
  snap.names.sort();

  for (const m of wb.matchAll(/<sheet\b[^>]*>/g)) {
    const name = getAttr(m[0], 'name') ?? '';
    const rid = getAttr(m[0], 'r:id');
    const rel = wbRels.find((r) => r.id === rid);
    if (!rel || rel.type !== 'worksheet') continue;
    const xml = (await read(rel.path)) ?? '';
    const model = new SheetModel(xml);
    model.unshare();
    const s: SheetSnap = {
      name, formulas: {}, values: {}, styles: {}, edgeStyles: {}, merges: [], cf: [], dv: [], cols: {}, rows: {}, pane: '', autoFilter: '',
      hyperlinks: [], comments: [], threaded: [], anchors: [], tables: [], sparklines: [], dimension: getAttr(/<dimension\b[^>]*>/.exec(xml)?.[0] ?? '', 'ref') ?? '',
    };
    const dynRefs: Array<[string, string]> = [];
    const cellXf = new Map<number, number>();
    /** positions of the cells that hold a text constant (where a quote prefix is visible) */
    const textCells = new Set<number>();
    const cellPos: Array<[number, number]> = [];
    for (const row of model.rows) {
      const ht = getAttr(row.a, 'ht');
      const hid = getAttr(row.a, 'hidden');
      const ol = getAttr(row.a, 'outlineLevel');
      const custom = getAttr(row.a, 'customHeight');
      const rs = getAttr(row.a, 's');
      const parts = [custom === '1' && ht ? `ht=${Number(ht)}` : '', hid === '1' ? 'hidden' : '', ol ? `ol=${ol}` : '', getAttr(row.a, 'customFormat') === '1' && rs ? `s=${styles[Number(rs)] ?? rs}` : ''].filter(Boolean);
      if (parts.length) s.rows[row.r] = parts.join(' ');
      for (const c of row.cells) {
        const ref = numToCol(c.c) + row.r;
        const st = Number(getAttr(c.a, 's') ?? 0);
        if (st && styles[st] !== styles[0]) s.styles[ref] = styles[st] ?? `xf${st}`;
        if (st) cellXf.set(row.r * EDGE_KEY + c.c, st);
        cellPos.push([row.r, c.c]);
        const f = parseF(c.x);
        if (f) {
          const fref = getAttr(f.attrs, 'ref');
          const t = getAttr(f.attrs, 't');
          // a dynamic array's spill extent (ref) is a recalculation result: not compared
          const dyn = /\scm="/.test(c.a);
          s.formulas[ref] = normQuotes(f.text) + (t === 'array' ? (dyn ? ' {dynamic}' : ` {array ${fref}}`) : '');
          if (dyn && fref) dynRefs.push([ref, fref]);
          continue;
        }
        const t = getAttr(c.a, 't');
        const v = /<v>([\s\S]*?)<\/v>/.exec(c.x ?? '')?.[1];
        if (t === 's' && v != null) { s.values[ref] = 's:' + (sst[Number(v)] ?? '?'); textCells.add(row.r * EDGE_KEY + c.c); }
        else if (t === 'inlineStr') { s.values[ref] = 's:' + siText(/<is>([\s\S]*?)<\/is>/.exec(c.x ?? '')?.[1] ?? ''); textCells.add(row.r * EDGE_KEY + c.c); }
        // numbers compare as doubles (Excel writes 17 significant digits: 46295.395833333299 == 46295.3958333333)
        else if (v != null) s.values[ref] = `${t ?? 'n'}:${(t ?? 'n') === 'n' && v.trim() !== '' && Number.isFinite(Number(v)) ? String(Number(v)) : unesc(v)}`;
      }
    }
    // spilled cells (values of a dynamic array's range) depend on recalculation: drop values covered by arrays
    const spill = new Set<string>();
    for (const [ref, fref] of dynRefs) {
      for (const a of parseSqref(fref)) for (let r = a.r1; r <= a.r2; r++) for (let c = a.c1; c <= a.c2; c++) { const k = numToCol(c) + r; if (k !== ref) { delete s.values[k]; spill.add(k); } }
    }
    // shared border edges merged (a line between two cells may be stored on either of them)
    s.edgeStyles = edgeStyles(cellXf, stylesParts, textCells);
    // the <dimension> of a sheet with dynamic arrays covers the spill ranges Excel computes when it recalculates (ours holds
    // the master cell only, or cached cells of an array that became values): there it is compared as the extent of the
    // cells OUTSIDE the spill ranges
    if (dynRefs.length) {
      let r1 = Infinity, r2 = 0, c1 = Infinity, c2 = 0;
      for (const [r, c] of cellPos) {
        if (spill.has(numToCol(c) + r)) continue;
        r1 = Math.min(r1, r); r2 = Math.max(r2, r); c1 = Math.min(c1, c); c2 = Math.max(c2, c);
      }
      s.dimension = r2 ? `${numToCol(c1)}${r1}:${numToCol(c2)}${r2}` : 'A1';
    }
    // cols
    const cols = findFirst(model.head, 'cols');
    if (cols) {
      for (const cm of cols.inner.matchAll(/<col\b[^>]*>/g)) {
        const min = Number(getAttr(cm[0], 'min')), max = Number(getAttr(cm[0], 'max'));
        const w = getAttr(cm[0], 'width');
        const desc = [w != null && getAttr(cm[0], 'customWidth') === '1' ? `w=${Number(w)}` : '', getAttr(cm[0], 'hidden') === '1' ? 'hidden' : '', getAttr(cm[0], 'outlineLevel') ? `ol=${getAttr(cm[0], 'outlineLevel')}` : '',
          getAttr(cm[0], 'style') && getAttr(cm[0], 'style') !== '0' ? `s=${styles[Number(getAttr(cm[0], 'style'))]}` : ''].filter(Boolean).join(' ');
        if (!desc) continue;
        for (let c = min; c <= Math.min(max, MAX_C); c++) s.cols[numToCol(c)] = desc;
      }
    }
    const pane = /<pane\b[^>]*>/.exec(model.head)?.[0];
    if (pane && /state="frozen/.test(pane)) s.pane = `frozen x=${getAttr(pane, 'xSplit') ?? 0} y=${getAttr(pane, 'ySplit') ?? 0}`;
    const tail = model.tail;
    for (const mm of tail.matchAll(/<mergeCell\b[^>]*>/g)) s.merges.push(getAttr(mm[0], 'ref') ?? '');
    s.merges.sort();
    const af = /<autoFilter\b[^>]*?(?:\/>|>[\s\S]*?<\/autoFilter>)/.exec(tail);
    if (af) s.autoFilter = (getAttr(af[0], 'ref') ?? '') + [...af[0].matchAll(/<filterColumn\b[^>]*colId="(\d+)"/g)].map((x) => ` fc${x[1]}`).join('');
    // CF: per-cell effective rules
    const dxfs = dxfSigs((stylesPath ? await read(stylesPath) : null) ?? '');
    const x14ById = new Map<string, string>();
    for (const blk of findAll(tail, 'x14:conditionalFormatting')) {
      for (const r of findAll(blk.inner, 'x14:cfRule')) {
        const id = getAttr(r.open, 'id');
        if (id) x14ById.set(id, canon(r.text.replace(/\sid="[^"]*"/, '').replace(/\spriority="\d+"/, '')));
      }
    }
    const cfCells = new Map<string, string[]>();
    const addCf = (sq: string, rule: string, formulas: string[]) => {
      const areas = parseSqref(sq);
      if (!areas.length) return;
      // formulas are anchored at the top-left of the sqref bounding box
      const { row: ar, col: ac } = sqrefTopLeft(areas);
      for (const a of areas) {
        for (let r = Math.min(a.r1, a.r2); r <= Math.min(Math.max(a.r1, a.r2), MAX_R); r++) {
          for (let c = Math.min(a.c1, a.c2); c <= Math.min(Math.max(a.c1, a.c2), MAX_C); c++) {
            const fk = formulas.map((f) => relKey(translateFormula(f, r - ar, c - ac), r, c)).join(' ; ');
            const k = numToCol(c) + r;
            const l = cfCells.get(k) ?? [];
            l.push(rule + (fk ? ' [' + fk + ']' : ''));
            cfCells.set(k, l);
          }
        }
      }
    };
    for (const blk of findAll(tail, 'conditionalFormatting')) {
      const sq = getAttr(blk.open, 'sqref') ?? '';
      for (const r of findAll(blk.inner, 'cfRule')) {
        const formulas = [...r.inner.matchAll(/<formula>([\s\S]*?)<\/formula>/g)].map((x) => unesc(x[1]));
        const dxf = getAttr(r.open, 'dxfId');
        const x14id = /<x14:id>([^<]*)<\/x14:id>/.exec(r.inner)?.[1];
        const body = canon(r.inner.replace(/<formula>[\s\S]*?<\/formula>/g, '').replace(/<extLst>[\s\S]*?<\/extLst>/g, ''));
        const desc = `${getAttr(r.open, 'type')} ${getAttr(r.open, 'operator') ?? ''} dxf=${dxf != null ? dxfs[Number(dxf)] : ''} ${body}${x14id ? ' x14=' + (x14ById.get(x14id) ?? '?') : ''}`;
        addCf(sq, desc, formulas);
      }
    }
    for (const blk of findAll(tail, 'x14:conditionalFormatting')) {
      const sq = unesc(/<xm:sqref>([\s\S]*?)<\/xm:sqref>/.exec(blk.inner)?.[1] ?? '');
      for (const r of findAll(blk.inner, 'x14:cfRule')) {
        if (getAttr(r.open, 'id') && [...tail.matchAll(/<x14:id>([^<]*)<\/x14:id>/g)].some((x) => x[1] === getAttr(r.open, 'id'))) continue;
        const formulas = [...r.inner.matchAll(/<xm:f>([\s\S]*?)<\/xm:f>/g)].map((x) => unesc(x[1]));
        addCf(sq, `x14 ${getAttr(r.open, 'type')} ${r.inner.replace(/<xm:f>[\s\S]*?<\/xm:f>/g, '')}`, formulas);
      }
    }
    for (const [k, l] of cfCells) s.cf.push(`${k}: ${l.sort().join(' | ')}`);
    s.cf.sort(cellSort);
    // DV
    const dvCells = new Map<string, string[]>();
    const addDv = (sq: string, desc: string, formulas: string[]) => {
      const areas = parseSqref(sq);
      if (!areas.length) return;
      // formulas are anchored at the top-left of the sqref bounding box
      const { row: ar, col: ac } = sqrefTopLeft(areas);
      for (const a of areas) for (let r = Math.min(a.r1, a.r2); r <= Math.min(Math.max(a.r1, a.r2), MAX_R); r++) for (let c = Math.min(a.c1, a.c2); c <= Math.min(Math.max(a.c1, a.c2), MAX_C); c++) {
        const fk = formulas.map((f) => relKey(translateFormula(f, r - ar, c - ac), r, c)).join(' ; ');
        const k = numToCol(c) + r;
        const l = dvCells.get(k) ?? [];
        l.push(desc + ' [' + fk + ']');
        dvCells.set(k, l);
      }
    };
    for (const d of findAll(tail, 'dataValidation')) {
      const formulas = [...d.inner.matchAll(/<(formula1|formula2)>([\s\S]*?)<\/\1>/g)].map((x) => unesc(x[2]));
      addDv(getAttr(d.open, 'sqref') ?? '', `${getAttr(d.open, 'type')} ${getAttr(d.open, 'operator') ?? ''}`, formulas);
    }
    for (const d of findAll(tail, 'x14:dataValidation')) {
      const formulas = [...d.inner.matchAll(/<xm:f>([\s\S]*?)<\/xm:f>/g)].map((x) => unesc(x[1]));
      addDv(unesc(/<xm:sqref>([\s\S]*?)<\/xm:sqref>/.exec(d.inner)?.[1] ?? ''), `${getAttr(d.open, 'type')} ${getAttr(d.open, 'operator') ?? ''}`, formulas);
    }
    for (const [k, l] of dvCells) s.dv.push(`${k}: ${l.sort().join(' | ')}`);
    s.dv.sort(cellSort);
    const sheetRels = await rels(rel.path);
    for (const h of tail.matchAll(/<hyperlink\b[^>]*>/g)) {
      const rid2 = getAttr(h[0], 'r:id');
      const target = rid2 ? (await read(relsOf(rel.path)))?.match(new RegExp(`Id="${rid2}"[^>]*Target="([^"]*)"`))?.[1] ?? '' : '';
      s.hyperlinks.push(`${getAttr(h[0], 'ref')} loc=${getAttr(h[0], 'location') ?? ''} target=${target ? unesc(target) : ''}`);
    }
    s.hyperlinks.sort();
    for (const sp of tail.matchAll(/<x14:sparkline>([\s\S]*?)<\/x14:sparkline>/g)) s.sparklines.push(`${unesc(/<xm:sqref>([^<]*)/.exec(sp[1])?.[1] ?? '')} <- ${unesc(/<xm:f>([^<]*)/.exec(sp[1])?.[1] ?? '')}`);
    for (const r of sheetRels) {
      if (r.external) continue;
      const px = await read(r.path);
      if (!px) continue;
      if (r.type === 'comments') for (const c of px.matchAll(/<comment\b[^>]*>/g)) s.comments.push(getAttr(c[0], 'ref') ?? '');
      else if (r.type === 'threadedComment') for (const c of px.matchAll(/<threadedComment\b[^>]*>/g)) s.threaded.push(getAttr(c[0], 'ref') ?? '');
      else if (r.type === 'drawing') {
        for (const a of px.matchAll(/<xdr:(twoCellAnchor|oneCellAnchor)\b[^>]*>([\s\S]*?)<\/xdr:\1>/g)) {
          const pt = (tag: string) => {
            const b = new RegExp(`<xdr:${tag}>([\\s\\S]*?)</xdr:${tag}>`).exec(a[2])?.[1];
            if (!b) return '';
            const g = (t: string) => new RegExp(`<xdr:${t}>(-?\\d+)</xdr:${t}>`).exec(b)?.[1];
            return `${numToCol(Number(g('col')) + 1)}${Number(g('row')) + 1}`;
          };
          const nm = /<xdr:cNvPr\b[^>]*name="([^"]*)"/.exec(a[2])?.[1] ?? '';
          s.anchors.push(`${nm}: ${pt('from')}-${pt('to')}`);
        }
      } else if (r.type === 'table') {
        const open = /<table\b[^>]*>/.exec(px)?.[0] ?? '';
        const tcols = [...px.matchAll(/<tableColumn\b[^>]*name="([^"]*)"/g)].map((x) => unesc(x[1]));
        const calc = [...px.matchAll(/<calculatedColumnFormula>([\s\S]*?)<\/calculatedColumnFormula>/g)].map((x) => unesc(x[1]));
        const tot = [...px.matchAll(/<totalsRowFormula(?:\s[^>]*)?>([\s\S]*?)<\/totalsRowFormula>/g)].map((x) => unesc(x[1]));
        // totals row definition per column: function / label
        for (const tc of px.matchAll(/<tableColumn\b[^>]*>/g)) {
          const fn = getAttr(tc[0], 'totalsRowFunction'), lb = getAttr(tc[0], 'totalsRowLabel');
          if (fn || lb != null) tot.push(`${unesc(getAttr(tc[0], 'name') ?? '')}:${fn ?? ''}${lb != null ? ` label=${unesc(lb)}` : ''}`);
        }
        const taf = /<autoFilter\b[^>]*ref="([^"]*)"/.exec(px)?.[1] ?? '';
        // header / totals row counts are part of the description: a ref that shrank over the totals row while totalsRowCount stayed
        // 1 is a table Excel cannot load (found by the random op-log fuzz: Excel stopped responding on it)
        s.tables.push(`${getAttr(open, 'displayName')} ref=${getAttr(open, 'ref')} hdr=${getAttr(open, 'headerRowCount') ?? 1} tot=${getAttr(open, 'totalsRowCount') ?? 0} af=${taf} cols=[${tcols.join('|')}] calc=[${calc.join('|')}]${tot.length ? ` totals=[${tot.join('|')}]` : ''}`);
      }
    }
    s.comments.sort(cellSort);
    s.threaded.sort(cellSort);
    s.anchors.sort();
    snap.sheets.push(s);
  }
  for (const p of Object.keys(zip.files).filter((n) => /^xl\/charts\/chart[^/]*\.xml$/.test(n)).sort()) {
    const x = (await read(p)) ?? '';
    snap.charts.push([...x.matchAll(/<c:f>([^<]*)<\/c:f>/g)].map((m) => unesc(m[1])).join(' ; '));
  }
  for (const p of Object.keys(zip.files).filter((n) => /pivotCacheDefinition[^/]*\.xml$/.test(n)).sort()) {
    const x = (await read(p)) ?? '';
    const ws = /<worksheetSource\b[^>]*>/.exec(x)?.[0] ?? '';
    snap.pivots.push(`${getAttr(ws, 'sheet') ?? getAttr(ws, 'name') ?? ''}!${getAttr(ws, 'ref') ?? ''}`);
  }
  return snap;
}

/** Excel re-quotes sheet prefixes when it re-serializes a formula ('名簿'!A1 vs 名簿!A1): compare unquoted. */
export function normQuotes(f: string): string {
  if (f.indexOf("'") < 0) return f;
  return tokenize(f).map((t) => {
    const p = t.prefix;
    if (!p || p.book || p.sheet2 !== undefined || !p.sheet || !p.text.startsWith("'")) return t.text;
    return quoteSheetName(p.sheet) === p.sheet ? p.sheet + '!' + t.text.slice(p.text.length) : t.text;
  }).join('');
}

function cellSort(a: string, b: string): number {
  const pa = /^([A-Z]+)(\d+)/.exec(a), pb = /^([A-Z]+)(\d+)/.exec(b);
  if (!pa || !pb) return a < b ? -1 : a > b ? 1 : 0;
  return Number(pa[2]) - Number(pb[2]) || colToNum(pa[1]) - colToNum(pb[1]) || (a < b ? -1 : 1);
}

/** One xf, resolved: the signature is `pre + 'border=' + borderXml + post`. */
interface XfParts {
  pre: string;
  borderXml: string;
  post: string;
}

function styleParts(xml: string): XfParts[] {
  const numFmts = new Map<number, string>();
  for (const m of xml.matchAll(/<numFmt\b[^>]*>/g)) numFmts.set(Number(getAttr(m[0], 'numFmtId')), getAttr(m[0], 'formatCode') ?? '');
  const sect = (tag: string, child: string) => {
    const s = findFirst(xml, tag);
    return s ? findAll(s.inner, child).map((e) => canon(e.text)) : [];
  };
  const fonts = sect('fonts', 'font');
  const fills = sect('fills', 'fill');
  const borders = sect('borders', 'border');
  const xfs = findFirst(xml, 'cellXfs');
  if (!xfs) return [];
  return findAll(xfs.inner, 'xf').map((e) => {
    const nf = Number(getAttr(e.open, 'numFmtId') ?? 0);
    // Excel normalises codes set through Range.NumberFormat by escaping literal characters (yyyy-mm-dd is stored as
    // yyyy\-mm\-dd); an escaped and an unescaped literal display the same
    const code = (nf >= 164 ? numFmts.get(nf) ?? `#${nf}` : `#${nf}`).replace(/\\([-+/():!^&'~{}<>= ])/g, '$1');
    const align = /<alignment\b[^>]*>/.exec(e.inner)?.[0] ?? '';
    const prot = /<protection\b[^>]*>/.exec(e.inner)?.[0] ?? '';
    return {
      pre: `nf=${code} font=${fonts[Number(getAttr(e.open, 'fontId') ?? 0)] ?? ''} fill=${fills[Number(getAttr(e.open, 'fillId') ?? 0)] ?? ''} `,
      borderXml: borders[Number(getAttr(e.open, 'borderId') ?? 0)] ?? '',
      post: ` ${canon(align)} ${canon(prot)}${getAttr(e.open, 'quotePrefix') === '1' ? ' quote' : ''}`,
    };
  });
}

/** xf index -> resolved style signature. */
export function styleSigs(xml: string): string[] {
  return styleParts(xml).map((p) => `${p.pre}border=${p.borderXml}${p.post}`);
}

// ---------------------------------------------------------------------------------------------------------------
// Shared border edges. A line between two neighbouring cells can be stored on EITHER of them (or on both): Excel moves
// A1.right to B1.left (and H1.bottom to H2.top) when it rewrites a border, while a patcher that only copies records keeps
// them where they were — the drawn line is the same. The snapshot therefore also carries, per cell, a signature in
// which each of the four border sides is the UNION of the cell's own side and the facing side of its neighbour
// (`edgeStyles`); diffSnap compares those, so only a difference in what is DRAWN shows up.
// ---------------------------------------------------------------------------------------------------------------

/** cell position key of the edge merge: row * EDGE_KEY + column (EDGE_KEY > MAX_COL) */
const EDGE_KEY = 20000;
const SIDES = ['left', 'right', 'top', 'bottom'] as const;
type SideName = (typeof SIDES)[number];
interface Borders {
  l: string;
  r: string;
  t: string;
  b: string;
  /** everything else of the <border> element (attributes, diagonal ...) */
  rest: string;
}
const NO_BORDERS: Borders = { l: '', r: '', t: '', b: '', rest: '' };

function splitBorder(xml: string): Borders {
  if (!xml) return NO_BORDERS;
  const side = (s: SideName): string => {
    const el = new RegExp(`<${s}\\b[^>]*?(?:/>|>[\\s\\S]*?</${s}>)`).exec(xml)?.[0] ?? '';
    // the line itself (style + colour), without the name of the side it was stored on
    return /\sstyle="(?!none")/.test(el) ? el.replace(new RegExp(`^<${s}\\b`), '<line').replace(new RegExp(`</${s}>$`), '</line>') : '';
  };
  const rest = xml.replace(new RegExp(`<(?:${SIDES.join('|')})\\b[^>]*?(?:/>|>[\\s\\S]*?</(?:${SIDES.join('|')})>)`, 'g'), '');
  return { l: side('left'), r: side('right'), t: side('top'), b: side('bottom'), rest };
}

/** the line drawn on an edge from the two facing sides (equal or one empty = that side; two different ones = both) */
function unionSide(a: string, b: string): string {
  if (!a) return b;
  if (!b || a === b) return a;
  return [a, b].sort().join('+');
}

function edgeSig(p: XfParts, own: Borders, l: string, r: string, t: string, b: string, quote = true): string {
  // a quotePrefix acts only on a text constant: on a number / a formula / an empty cell it is invisible (Excel through COM keeps
  // the flag where typing resets it), so it counts only where it shows
  return `${p.pre}border=[L:${l}|R:${r}|T:${t}|B:${b}|${own.rest}]${quote ? p.post : p.post.replace(/ quote$/, '')}`;
}

/** cell ref -> signature with shared edges merged; cells that look like an unstyled cell are left out */
function edgeStyles(cellXf: Map<number, number>, parts: XfParts[], textCells: ReadonlySet<number>): Record<string, string> {
  const KEY = EDGE_KEY;
  const key = (r: number, c: number) => r * KEY + c;
  const borders = parts.map((p) => splitBorder(p.borderXml));
  const ownOf = (r: number, c: number): Borders => borders[cellXf.get(key(r, c)) ?? 0] ?? NO_BORDERS;
  const out: Record<string, string> = {};
  const def = parts[0] ? edgeSig(parts[0], borders[0], borders[0].l, borders[0].r, borders[0].t, borders[0].b) : '';
  const cand = new Set<number>();
  for (const [k, xf] of cellXf) {
    cand.add(k);
    const b = borders[xf] ?? NO_BORDERS;
    if (b.l || b.r || b.t || b.b) for (const d of [-KEY, KEY, -1, 1]) cand.add(k + d);
  }
  for (const k of cand) {
    const r = Math.floor(k / KEY), c = k - r * KEY;
    if (r < 1 || c < 1) continue;
    const xf = cellXf.get(k) ?? 0;
    const p = parts[xf] ?? parts[0];
    if (!p) continue;
    const own = borders[xf] ?? NO_BORDERS;
    const sig = edgeSig(p, own,
      unionSide(own.l, ownOf(r, c - 1).r), unionSide(own.r, ownOf(r, c + 1).l),
      unionSide(own.t, ownOf(r - 1, c).b), unionSide(own.b, ownOf(r + 1, c).t), textCells.has(k));
    if (sig !== def) out[numToCol(c) + r] = sig;
  }
  return out;
}

export function dxfSigs(xml: string): string[] {
  const s = findFirst(xml, 'dxfs');
  return s ? findAll(s.inner, 'dxf').map((e) => canon(e.inner)) : [];
}

/** attribute-order independent canonical form of a small XML fragment */
function canon(x: string): string {
  return x.replace(/<([\w:]+)((?:\s+[\w:]+="[^"]*")*)\s*(\/?)>/g, (_m, tag: string, attrs: string, sc: string) => {
    const list = [...attrs.matchAll(/([\w:]+)="([^"]*)"/g)].map((a) => `${a[1]}="${a[2]}"`).sort();
    return `<${tag}${list.length ? ' ' + list.join(' ') : ''}${sc}>`;
  }).replace(/>\s+</g, '><');
}

export interface SnapDiff {
  path: string;
  a: unknown;
  b: unknown;
}

/** Structural diff of two snapshots (a = ours, b = Excel). */
export function diffSnap(a: BookSnap, b: BookSnap, only?: Array<keyof SheetSnap | 'names' | 'charts' | 'pivots'>): SnapDiff[] {
  const out: SnapDiff[] = [];
  const want = (k: string) => !only || (only as string[]).includes(k);
  const cmpRec = (path: string, x: Record<string, string>, y: Record<string, string>) => {
    for (const k of new Set([...Object.keys(x), ...Object.keys(y)])) if (x[k] !== y[k]) out.push({ path: `${path}.${k}`, a: x[k], b: y[k] });
  };
  const cmpList = (path: string, x: string[], y: string[]) => {
    const sx = new Set(x), sy = new Set(y);
    for (const v of x) if (!sy.has(v)) out.push({ path, a: v, b: undefined });
    for (const v of y) if (!sx.has(v)) out.push({ path, a: undefined, b: v });
  };
  if (want('names')) cmpList('names', a.names, b.names);
  if (want('charts')) cmpList('charts', a.charts, b.charts);
  if (want('pivots')) cmpList('pivots', a.pivots, b.pivots);
  for (const sb of b.sheets) {
    const sa = a.sheets.find((s) => s.name === sb.name);
    if (!sa) { out.push({ path: `sheet ${sb.name}`, a: undefined, b: 'present' }); continue; }
    const p = `[${sb.name}]`;
    for (const k of ['formulas', 'values', 'cols', 'rows'] as const) if (want(k)) cmpRec(`${p}.${k}`, sa[k], sb[k]);
    // styles: what is DRAWN (shared border edges merged), reported under the old path `[sheet].styles.<ref>`
    if (want('styles')) cmpRec(`${p}.styles`, sa.edgeStyles, sb.edgeStyles);
    for (const k of ['merges', 'cf', 'dv', 'hyperlinks', 'comments', 'threaded', 'anchors', 'tables', 'sparklines'] as const) if (want(k)) cmpList(`${p}.${k}`, sa[k], sb[k]);
    for (const k of ['pane', 'autoFilter', 'dimension'] as const) if (want(k) && sa[k] !== sb[k]) out.push({ path: `${p}.${k}`, a: sa[k], b: sb[k] });
  }
  return out;
}
