// Structural linter for .xlsx packages (round 5, module H1): finds what makes Excel "repair" or refuse a file WITHOUT starting Excel.
// The patcher edits XML as text; every rule below is one that Excel enforces on load (child order of the schema sequences, ascending rows /
// cells / <col>s, style / shared-string / dxf indexes inside their tables, merges that do not overlap, a table header that says what the
// table columns say, relationship targets that exist, unique defined names / table names, a visible sheet, shared formulas with a master ...).
// Independent of the patcher's own parsing (JSZip + string scanning). Excel-written packages (tests/fixtures/golden/*.xlsx) must lint clean.
//   lintPackage(bytes) -> problems [{ code, part, detail }]
//   newProblems(before, after) -> the problems of `after` whose CODE does not occur in `before` (a damaged input stays damaged: only new kinds count)
import JSZip from 'jszip';
import { decodeBytes } from '../../../src/core/xlsx/patch/xml';

export interface LintProblem {
  code: string;
  part: string;
  detail: string;
}

const NS = '(?:[A-Za-z_][\\w.-]*:)?';
const local = (n: string) => n.replace(/^[^:]*:/, '');
const attr = (tag: string, name: string): string | undefined => {
  const m = new RegExp(`\\s(?:[\\w.-]+:)?${name}\\s*=\\s*(?:"([^"]*)"|'([^']*)')`).exec(tag);
  return m ? (m[1] ?? m[2]) : undefined;
};
const unesc = (s: string) => s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&#x([0-9a-fA-F]+);/g, (_m, h: string) => String.fromCodePoint(parseInt(h, 16))).replace(/&#(\d+);/g, (_m, d: string) => String.fromCodePoint(Number(d))).replace(/&amp;/g, '&');

const ORDER: Record<string, string[]> = {
  workbook: ['fileVersion', 'fileSharing', 'workbookPr', 'workbookProtection', 'bookViews', 'sheets', 'functionGroups', 'externalReferences', 'definedNames', 'calcPr', 'oleSize', 'customWorkbookViews', 'pivotCaches', 'smartTagPr', 'smartTagTypes', 'webPublishing', 'fileRecoveryPr', 'webPublishObjects', 'extLst'],
  worksheet: ['sheetPr', 'dimension', 'sheetViews', 'sheetFormatPr', 'cols', 'sheetData', 'sheetCalcPr', 'sheetProtection', 'protectedRanges', 'scenarios', 'autoFilter', 'sortState', 'dataConsolidate', 'customSheetViews', 'mergeCells', 'phoneticPr', 'conditionalFormatting', 'dataValidations', 'hyperlinks', 'printOptions', 'pageMargins', 'pageSetup', 'headerFooter', 'rowBreaks', 'colBreaks', 'customProperties', 'cellWatches', 'ignoredErrors', 'smartTags', 'drawing', 'legacyDrawing', 'legacyDrawingHF', 'drawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst'],
  styleSheet: ['numFmts', 'fonts', 'fills', 'borders', 'cellStyleXfs', 'cellXfs', 'cellStyles', 'dxfs', 'tableStyles', 'colors', 'extLst'],
  table: ['autoFilter', 'sortState', 'tableColumns', 'tableStyleInfo', 'extLst'],
  xf: ['alignment', 'protection', 'extLst'],
  border: ['start', 'left', 'end', 'right', 'top', 'bottom', 'diagonal', 'vertical', 'horizontal'],
  sheetView: ['pane', 'selection', 'pivotSelection', 'extLst'],
  sst: ['si', 'extLst'],
  comments: ['authors', 'commentList', 'extLst'],
  pageSetup: [],
};

interface Child {
  name: string;
  start: number;
  end: number;
  open: string;
}

/** Direct children of the root element of `xml` (local names, with the span of the child). Linear, comments / CDATA / PIs skipped. */
function topChildren(xml: string): { root: string; rootOpen: string; children: Child[] } | null {
  let i = 0;
  const n = xml.length;
  let depth = 0;
  let root = '';
  let rootOpen = '';
  const children: Child[] = [];
  let cur: Child | null = null;
  while (i < n) {
    const lt = xml.indexOf('<', i);
    if (lt < 0) break;
    if (xml.startsWith('<!--', lt)) { i = xml.indexOf('-->', lt + 4) + 3 || n; continue; }
    if (xml.startsWith('<![CDATA[', lt)) { i = xml.indexOf(']]>', lt + 9) + 3 || n; continue; }
    if (xml[lt + 1] === '?' || xml[lt + 1] === '!') { const e = xml.indexOf('>', lt); i = e < 0 ? n : e + 1; continue; }
    // end of the tag, quotes respected
    let j = lt + 1;
    let q = '';
    while (j < n) {
      const ch = xml[j];
      if (q) { if (ch === q) q = ''; } else if (ch === '"' || ch === "'") q = ch; else if (ch === '>') break;
      j++;
    }
    const tag = xml.slice(lt, j + 1);
    i = j + 1;
    if (tag[1] === '/') {
      depth--;
      if (depth === 1 && cur) { cur.end = i; children.push(cur); cur = null; }
      continue;
    }
    const nm = /^<([^\s/>]+)/.exec(tag)![1];
    const selfClose = tag[tag.length - 2] === '/';
    if (depth === 0) { root = local(nm); rootOpen = tag; }
    else if (depth === 1) {
      cur = { name: local(nm), start: lt, end: i, open: tag };
      if (selfClose) { children.push(cur); cur = null; }
    }
    if (!selfClose) depth++;
  }
  return { root, rootOpen, children };
}

const colNum = (s: string): number => { let v = 0; for (const ch of s) v = v * 26 + ch.charCodeAt(0) - 64; return v; };
const parseRef = (ref: string): { c1: number; r1: number; c2: number; r2: number } | null => {
  const m = /^\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?$/.exec(ref.trim());
  if (!m) return null;
  const c1 = colNum(m[1]), r1 = Number(m[2]);
  const c2 = m[3] ? colNum(m[3]) : c1, r2 = m[4] ? Number(m[4]) : r1;
  return { c1: Math.min(c1, c2), r1: Math.min(r1, r2), c2: Math.max(c1, c2), r2: Math.max(r1, r2) };
};
const overlap = (a: { c1: number; r1: number; c2: number; r2: number }, b: typeof a) => a.c1 <= b.c2 && b.c1 <= a.c2 && a.r1 <= b.r2 && b.r1 <= a.r2;

/** relationship target of `rel` (from a .rels file of `owner`) as a package part name, null for external / non-part targets */
function resolveTarget(owner: string, target: string): string {
  let t = target;
  try { t = decodeURIComponent(t); } catch { /* keep */ }
  if (t.startsWith('/')) return t.slice(1);
  const dir = owner.includes('/') ? owner.slice(0, owner.lastIndexOf('/')) : '';
  const parts = dir ? dir.split('/') : [];
  for (const seg of t.split('/')) { if (seg === '..') parts.pop(); else if (seg !== '.' && seg !== '') parts.push(seg); }
  return parts.join('/');
}

export async function lintPackage(bytes: Uint8Array): Promise<LintProblem[]> {
  const out: LintProblem[] = [];
  const add = (code: string, part: string, detail: string) => { if (out.length < 400) out.push({ code, part, detail }); };
  const zip = await JSZip.loadAsync(bytes);
  const names = Object.keys(zip.files).filter((x) => !zip.files[x].dir);
  const byLower = new Map<string, string>();
  for (const x of names) byLower.set(x.toLowerCase(), x);
  const cache = new Map<string, string>();
  const text = async (p: string): Promise<string | null> => {
    const real = byLower.get(p.toLowerCase());
    if (!real) return null;
    if (!cache.has(real)) cache.set(real, decodeBytes(await zip.file(real)!.async('uint8array')).text);
    return cache.get(real)!;
  };
  const checkOrder = (part: string, xml: string, kind: string, scope = 'root') => {
    const order = ORDER[kind];
    const t = topChildren(xml);
    if (!t || !order?.length) return t;
    let last = -1;
    let lastName = '';
    for (const c of t.children) {
      const idx = order.indexOf(c.name);
      if (idx < 0) continue; // mc:AlternateContent, x14ac extensions, unknown: ignored
      if (idx < last) add('ORDER', part, `<${c.name}> comes after <${lastName}> (${scope} ${kind})`);
      else { last = idx; lastName = c.name; }
    }
    return t;
  };

  // ---- content types
  const ct = await text('[Content_Types].xml');
  const defaults = new Map<string, string>();
  const overrides = new Map<string, string>();
  if (ct === null) add('CT_MISSING', '[Content_Types].xml', 'no [Content_Types].xml');
  else {
    for (const m of ct.matchAll(/<Default\b[^>]*>/g)) defaults.set((attr(m[0], 'Extension') ?? '').toLowerCase(), attr(m[0], 'ContentType') ?? '');
    for (const m of ct.matchAll(/<Override\b[^>]*>/g)) {
      const pn = (attr(m[0], 'PartName') ?? '').replace(/^\//, '');
      overrides.set(pn.toLowerCase(), attr(m[0], 'ContentType') ?? '');
      if (!byLower.has(pn.toLowerCase())) add('CT_DANGLING', '[Content_Types].xml', `Override for the missing part ${pn}`);
    }
    for (const x of names) {
      if (x === '[Content_Types].xml') continue;
      const ext = x.includes('.') ? x.slice(x.lastIndexOf('.') + 1).toLowerCase() : '';
      if (!overrides.has(x.toLowerCase()) && !defaults.has(ext)) add('CT_NONE', x, 'part without a content type');
    }
  }

  // ---- relationships
  const rels = new Map<string, Map<string, { type: string; target: string; part: string | null; external: boolean }>>();
  for (const x of names) {
    if (!/\.rels$/i.test(x)) continue;
    const xml = (await text(x))!;
    const owner = x.replace(/(^|\/)_rels\/([^/]*)\.rels$/i, (_m, d: string, f: string) => d + f);
    const ownerPart = x === '_rels/.rels' ? '' : owner;
    const map = new Map<string, { type: string; target: string; part: string | null; external: boolean }>();
    for (const m of xml.matchAll(/<Relationship\b[^>]*>/g)) {
      const id = attr(m[0], 'Id') ?? '';
      const target = unesc(attr(m[0], 'Target') ?? '');
      const external = /^external$/i.test(attr(m[0], 'TargetMode') ?? '');
      if (map.has(id)) add('REL_DUP_ID', x, `Id ${id} twice`);
      const part = external ? null : resolveTarget(ownerPart, target);
      if (part !== null && !byLower.has(part.toLowerCase())) add('REL_DANGLING', x, `${id} -> ${target} (no such part)`);
      map.set(id, { type: (attr(m[0], 'Type') ?? '').replace(/^.*\//, ''), target, part, external });
    }
    rels.set(ownerPart.toLowerCase(), map);
  }

  // ---- workbook
  const rootRels = rels.get('');
  const wbRel = [...(rootRels?.values() ?? [])].find((r) => r.type === 'officeDocument');
  if (!wbRel || !wbRel.part) { add('NO_WORKBOOK', '_rels/.rels', 'no officeDocument relationship'); return out; }
  const wbPath = byLower.get(wbRel.part.toLowerCase()) ?? wbRel.part;
  const wb = await text(wbPath);
  if (wb === null) return out;
  const wbTop = checkOrder(wbPath, wb, 'workbook');
  if (!wbTop || wbTop.root !== 'workbook') add('ROOT', wbPath, `root element ${wbTop?.root}`);
  const wbRels = rels.get(wbPath.toLowerCase()) ?? new Map();
  const sheets: Array<{ name: string; id: string; rid: string; part: string | null; state: string }> = [];
  const sheetsEl = /<(?:[\w.-]+:)?sheets\b[\s\S]*?<\/(?:[\w.-]+:)?sheets>/.exec(wb)?.[0] ?? '';
  for (const m of sheetsEl.matchAll(/<(?:[\w.-]+:)?sheet\b[^>]*>/g)) {
    const rid = attr(m[0], 'id') ?? '';
    sheets.push({ name: unesc(attr(m[0], 'name') ?? ''), id: attr(m[0], 'sheetId') ?? '', rid, part: wbRels.get(rid)?.part ?? null, state: attr(m[0], 'state') ?? 'visible' });
  }
  if (!sheets.length) add('NO_SHEETS', wbPath, 'no sheets');
  const seenName = new Set<string>();
  const seenId = new Set<string>();
  for (const s of sheets) {
    const k = s.name.toLowerCase();
    if (seenName.has(k)) add('SHEET_NAME_DUP', wbPath, s.name);
    seenName.add(k);
    if (seenId.has(s.id)) add('SHEET_ID_DUP', wbPath, s.id);
    seenId.add(s.id);
    if (!s.name || s.name.length > 31 || /[[\]:*?/\\]/.test(s.name) || /^'|'$/.test(s.name)) add('SHEET_NAME_BAD', wbPath, s.name);
    if (!s.part) add('SHEET_REL', wbPath, `${s.name}: relationship ${s.rid} not found`);
  }
  if (sheets.length && sheets.every((s) => s.state !== 'visible')) add('ALL_HIDDEN', wbPath, 'every sheet is hidden');
  const active = Number(attr(/<(?:[\w.-]+:)?workbookView\b[^>]*>/.exec(wb)?.[0] ?? '', 'activeTab') ?? 0);
  if (sheets.length && active >= sheets.length) add('ACTIVE_TAB', wbPath, `activeTab ${active} of ${sheets.length}`);
  else if (sheets.length && sheets[active]?.state !== 'visible') add('ACTIVE_TAB_HIDDEN', wbPath, `activeTab ${active} is a hidden sheet`);
  const dnSeen = new Set<string>();
  for (const m of wb.matchAll(/<(?:[\w.-]+:)?definedName\b([^>]*?)(?:\/>|>([\s\S]*?)<\/(?:[\w.-]+:)?definedName>)/g)) {
    const key = `${(attr(m[0], 'name') ?? '').toLowerCase()}|${attr(m[0], 'localSheetId') ?? ''}`;
    if (dnSeen.has(key)) add('NAME_DUP', wbPath, key);
    dnSeen.add(key);
    const ls = attr(m[0], 'localSheetId');
    if (ls !== undefined && Number(ls) >= sheets.length) add('NAME_SHEET', wbPath, `${key}: localSheetId out of range`);
    if (!(m[2] ?? '').trim() && !/\/>$/.test(m[0])) add('NAME_EMPTY', wbPath, key);
  }

  // ---- styles
  const stylesRel = [...wbRels.values()].find((r) => r.type === 'styles');
  const stylesXml = stylesRel?.part ? await text(stylesRel.part) : null;
  let xfCount = 0;
  let dxfCount = 0;
  if (stylesXml !== null) {
    const sp = stylesRel!.part!;
    checkOrder(sp, stylesXml, 'styleSheet');
    const block = (tag: string) => new RegExp(`<${NS}${tag}\\b[^>]*?(?:/>|>([\\s\\S]*?)</${NS}${tag}>)`).exec(stylesXml)?.[1] ?? '';
    const count = (body: string, tag: string) => (body.match(new RegExp(`<${NS}${tag}(?=[\\s/>])`, 'g')) ?? []).length;
    const fonts = count(block('fonts'), 'font');
    const fills = count(block('fills'), 'fill');
    const borders = count(block('borders'), 'border');
    const cellStyleXfs = count(block('cellStyleXfs'), 'xf');
    const xfBody = block('cellXfs');
    xfCount = count(xfBody, 'xf');
    dxfCount = count(block('dxfs'), 'dxf');
    const fmtIds = new Set<number>();
    for (const m of block('numFmts').matchAll(/<(?:[\w.-]+:)?numFmt\b[^>]*>/g)) {
      fmtIds.add(Number(attr(m[0], 'numFmtId')));
      if (!(attr(m[0], 'formatCode') ?? '').length) add('NUMFMT_EMPTY', sp, `numFmtId ${attr(m[0], 'numFmtId')} has an empty formatCode`);
    }
    if (xfCount === 0) add('NO_CELLXFS', sp, 'cellXfs is empty');
    let k = 0;
    for (const m of xfBody.matchAll(/<(?:[\w.-]+:)?xf\b[^>]*?(?:\/>|>[\s\S]*?<\/(?:[\w.-]+:)?xf>)/g)) {
      const t = /^<[^>]*>/.exec(m[0])![0];
      const f = Number(attr(t, 'fontId') ?? 0), fl = Number(attr(t, 'fillId') ?? 0), b = Number(attr(t, 'borderId') ?? 0), nf = Number(attr(t, 'numFmtId') ?? 0), xi = Number(attr(t, 'xfId') ?? 0);
      if (f >= fonts) add('XF_FONT', sp, `cellXfs[${k}] fontId ${f} of ${fonts}`);
      if (fl >= fills) add('XF_FILL', sp, `cellXfs[${k}] fillId ${fl} of ${fills}`);
      if (b >= borders) add('XF_BORDER', sp, `cellXfs[${k}] borderId ${b} of ${borders}`);
      if (nf >= 164 && !fmtIds.has(nf)) add('XF_NUMFMT', sp, `cellXfs[${k}] numFmtId ${nf} is not defined`);
      if (cellStyleXfs && xi >= cellStyleXfs) add('XF_XFID', sp, `cellXfs[${k}] xfId ${xi} of ${cellStyleXfs}`);
      checkOrder(sp, m[0], 'xf', `cellXfs[${k}]`);
      k++;
    }
    for (const m of block('borders').matchAll(/<(?:[\w.-]+:)?border\b[^>]*?>([\s\S]*?)<\/(?:[\w.-]+:)?border>/g)) checkOrder(sp, m[0], 'border', 'border');
    const cs = block('cellStyles');
    for (const m of cs.matchAll(/<(?:[\w.-]+:)?cellStyle\b[^>]*>/g)) if (cellStyleXfs && Number(attr(m[0], 'xfId') ?? 0) >= cellStyleXfs) add('CELLSTYLE_XFID', sp, m[0]);
  }

  // ---- shared strings
  const sstRel = [...wbRels.values()].find((r) => r.type === 'sharedStrings');
  const sstXml = sstRel?.part ? await text(sstRel.part) : null;
  const sstCount = sstXml === null ? 0 : (sstXml.match(new RegExp(`<${NS}si(?=[\\s/>])`, 'g')) ?? []).length;
  /** shared string text by index (lazy) */
  let sstTexts: string[] | null = null;
  const sstText = (i: number): string => {
    if (sstTexts === null) {
      sstTexts = [];
      for (const m of (sstXml ?? '').matchAll(new RegExp(`<${NS}si\\b[^>]*?(?:/>|>([\\s\\S]*?)</${NS}si>)`, 'g'))) {
        const body = (m[1] ?? '').replace(new RegExp(`<${NS}rPh\\b[\\s\\S]*?</${NS}rPh>`, 'g'), '');
        sstTexts.push(unesc([...body.matchAll(new RegExp(`<${NS}t\\b[^>]*?(?:/>|>([\\s\\S]*?)</${NS}t>)`, 'g'))].map((t) => t[1] ?? '').join('')));
      }
    }
    return sstTexts[i] ?? '';
  };

  // ---- tables (workbook wide: unique names / ids)
  const tableNames = new Map<string, string>();
  const tableIds = new Map<string, string>();
  const tableBoxes: Array<{ sheet: string; part: string; box: NonNullable<ReturnType<typeof parseRef>>; name: string }> = [];

  // ---- sheets
  for (const sh of sheets) {
    if (!sh.part) continue;
    const sp = byLower.get(sh.part.toLowerCase());
    if (!sp) continue;
    const xml = (await text(sp))!;
    const top = checkOrder(sp, xml, 'worksheet');
    if (!top || (top.root !== 'worksheet' && top.root !== 'chartsheet' && top.root !== 'dialogsheet')) { add('ROOT', sp, `root element ${top?.root}`); continue; }
    if (top.root !== 'worksheet') continue;
    const sRels = rels.get(sp.toLowerCase()) ?? new Map();
    const need = (el: string, type: string) => {
      for (const c of top.children.filter((x) => x.name === el)) {
        const rid = attr(c.open, 'id');
        if (rid === undefined) { if (el !== 'tablePart') add('SHEET_REL', sp, `<${el}> without r:id`); continue; }
        const r = sRels.get(rid);
        if (!r) add('SHEET_REL', sp, `<${el} r:id="${rid}"> has no relationship`);
        else if (r.type !== type) add('SHEET_REL_TYPE', sp, `<${el} r:id="${rid}"> points at a ${r.type}`);
      }
    };
    need('drawing', 'drawing');
    need('legacyDrawing', 'vmlDrawing');
    const tp = top.children.find((c) => c.name === 'tableParts');
    if (tp) for (const m of xml.slice(tp.start, tp.end).matchAll(/<(?:[\w.-]+:)?tablePart\b[^>]*>/g)) {
      const rid = attr(m[0], 'id') ?? '';
      const r = sRels.get(rid);
      if (!r) add('SHEET_REL', sp, `<tablePart r:id="${rid}"> has no relationship`);
      else if (r.type !== 'table') add('SHEET_REL_TYPE', sp, `<tablePart r:id="${rid}"> points at a ${r.type}`);
    }
    for (const m of xml.matchAll(/<(?:[\w.-]+:)?hyperlink\b[^>]*>/g)) {
      const rid = attr(m[0], 'id');
      if (rid !== undefined && !sRels.has(rid)) add('SHEET_REL', sp, `<hyperlink r:id="${rid}"> has no relationship`);
      const rf = attr(m[0], 'ref');
      if (rf === undefined || !parseRef(rf)) add('HYPERLINK_REF', sp, `ref ${rf}`);
    }
    // cols: ascending, not overlapping, inside 1..16384
    const colsEl = top.children.find((c) => c.name === 'cols');
    if (colsEl) {
      let prev = 0;
      for (const m of xml.slice(colsEl.start, colsEl.end).matchAll(/<(?:[\w.-]+:)?col\b[^>]*>/g)) {
        const mn = Number(attr(m[0], 'min')), mx = Number(attr(m[0], 'max'));
        if (!(mn >= 1 && mx >= mn && mx <= 16384)) add('COL_RANGE', sp, m[0]);
        else if (mn <= prev) add('COL_ORDER', sp, `<col min=${mn} max=${mx}> after a <col> ending at ${prev}`);
        prev = Math.max(prev, mx);
      }
    }
    // merges
    const mergeBoxes: Array<NonNullable<ReturnType<typeof parseRef>>> = [];
    const me = top.children.find((c) => c.name === 'mergeCells');
    if (me) {
      for (const m of xml.slice(me.start, me.end).matchAll(/<(?:[\w.-]+:)?mergeCell\b[^>]*>/g)) {
        const b = parseRef(attr(m[0], 'ref') ?? '');
        if (!b) { add('MERGE_REF', sp, m[0]); continue; }
        if (b.c1 === b.c2 && b.r1 === b.r2) add('MERGE_SINGLE', sp, m[0]);
        if (b.c2 > 16384 || b.r2 > 1048576) add('MERGE_BOUNDS', sp, m[0]);
        for (const o of mergeBoxes) if (overlap(o, b)) add('MERGE_OVERLAP', sp, `${attr(m[0], 'ref')} overlaps another merge`);
        mergeBoxes.push(b);
      }
      if (!mergeBoxes.length) add('MERGE_EMPTY', sp, '<mergeCells> without a mergeCell');
    }
    // conditional formats: dxf ids inside the table
    for (const m of xml.matchAll(/<(?:[\w.-]+:)?cfRule\b[^>]*>/g)) {
      const d = attr(m[0], 'dxfId');
      if (d !== undefined && Number(d) >= dxfCount) add('CF_DXF', sp, `dxfId ${d} of ${dxfCount}`);
    }
    // rows / cells
    const sd = top.children.find((c) => c.name === 'sheetData');
    const cellText = new Map<string, string>(); // header text lookup for tables: "r,c" -> text
    if (sd) {
      const body = xml.slice(sd.start, sd.end);
      let prevRow = 0;
      const shared = new Map<string, { master: boolean; children: number; ref: string | null; at: string }>();
      for (const rm of body.matchAll(new RegExp(`<${NS}row\\b([^>]*?)(?:/>|>([\\s\\S]*?)</${NS}row>)`, 'g'))) {
        // a writer may leave the row number out (hc-11: Excel counts them)
        const r = attr(rm[0], 'r') === undefined ? prevRow + 1 : Number(attr(rm[0], 'r'));
        if (!(r >= 1 && r <= 1048576)) { add('ROW_R', sp, `row r=${attr(rm[0], 'r')}`); continue; }
        if (r <= prevRow) add('ROW_ORDER', sp, `row ${r} after row ${prevRow}`);
        prevRow = Math.max(prevRow, r);
        let prevCol = 0;
        let implicit = 0;
        for (const cm of (rm[2] ?? '').matchAll(new RegExp(`<${NS}c\\b([^>]*?)(?:/>|>([\\s\\S]*?)</${NS}c>)`, 'g'))) {
          const ref = attr(cm[0], 'r');
          let c: number;
          if (ref !== undefined) {
            const p = /^([A-Z]{1,3})(\d+)$/.exec(ref);
            if (!p) { add('CELL_REF', sp, `r="${ref}"`); continue; }
            c = colNum(p[1]);
            if (Number(p[2]) !== r) add('CELL_ROW', sp, `cell ${ref} inside row ${r}`);
            implicit = c;
          } else c = ++implicit;
          if (c > 16384) add('CELL_COL', sp, `${ref}`);
          if (c <= prevCol) add('CELL_ORDER', sp, `cell ${ref ?? c} after column ${prevCol} in row ${r}`);
          prevCol = Math.max(prevCol, c);
          const open = /^<[^>]*>/.exec(cm[0])![0];
          const s = attr(open, 's');
          if (s !== undefined && !(Number(s) >= 0 && Number(s) < Math.max(xfCount, 1))) add('CELL_STYLE', sp, `${ref} s=${s} of ${xfCount} xfs`);
          const t = attr(open, 't');
          const inner = cm[2] ?? '';
          const v = new RegExp(`<${NS}v>([\\s\\S]*?)</${NS}v>`).exec(inner)?.[1];
          if (t === 's') {
            if (v === undefined || !/^\d+$/.test(v.trim()) || Number(v) >= sstCount) add('CELL_SST', sp, `${ref} t="s" v=${v} of ${sstCount} strings`);
            else cellText.set(`${r},${c}`, sstText(Number(v)));
          } else if (t === 'inlineStr') {
            const m = new RegExp(`<${NS}is>([\\s\\S]*?)</${NS}is>`).exec(inner);
            if (!m) { /* pandas writes an empty inline string for NaN: Excel reads it as an empty cell */ }
            else cellText.set(`${r},${c}`, unesc([...m[1].matchAll(new RegExp(`<${NS}t\\b[^>]*?(?:/>|>([\\s\\S]*?)</${NS}t>)`, 'g'))].map((x) => x[1] ?? '').join('')));
          } else if (t === 'str') { if (v !== undefined) cellText.set(`${r},${c}`, unesc(v)); }
          else if (t === 'b') { if (v !== undefined && !/^\s*(0|1|true|false)\s*$/i.test(v)) add('CELL_BOOL', sp, `${ref} v=${v}`); }
          else if (t === 'e') { if (v === undefined || !/^#(NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|GETTING_DATA|SPILL!|CALC!|FIELD!|BLOCKED!|CONNECT!|BUSY!|UNKNOWN!)$/.test(unesc(v).trim())) add('CELL_ERR', sp, `${ref} v=${v}`); }
          else if (t === undefined || t === 'n') {
            if (v !== undefined && v.trim() !== '') { // <v></v> (XlsxWriter array formulas without a result) is read as no value
              if (!/^\s*[-+]?(\d+\.?\d*|\.\d+)([eE][-+]?\d+)?\s*$/.test(v)) add('CELL_NUM', sp, `${ref} v="${v.slice(0, 20)}"`);
              else cellText.set(`${r},${c}`, v.trim());
            }
          } else if (t === 'd') { /* ISO date, fine */ }
          else add('CELL_TYPE', sp, `${ref} t="${t}"`);
          for (const fm of inner.matchAll(new RegExp(`<${NS}f\\b([^>]*?)(?:/>|>([\\s\\S]*?)</${NS}f>)`, 'g'))) {
            const ft = attr(fm[0], 't');
            const body2 = (fm[2] ?? '').trim();
            if (ft === 'shared') {
              const si = attr(fm[0], 'si') ?? '';
              const e = shared.get(si) ?? { master: false, children: 0, ref: null, at: '' };
              if (body2) { if (e.master) add('SHARED_TWO_MASTERS', sp, `si ${si}`); e.master = true; e.ref = attr(fm[0], 'ref') ?? null; e.at = ref ?? ''; } else e.children++;
              shared.set(si, e);
            } else if (!body2 && ft !== 'dataTable') add('F_EMPTY', sp, `${ref} has an empty <f>`);
            if (body2.length > 8192) add('F_LONG', sp, `${ref} formula of ${body2.length} characters`);
            if (ft === 'array') {
              const ar = parseRef(attr(fm[0], 'ref') ?? '');
              if (!ar || ar.c2 > 16384 || ar.r2 > 1048576) add('ARRAY_REF', sp, `${ref} ref=${attr(fm[0], 'ref')}`);
            }
          }
        }
      }
      for (const [si, e] of shared) {
        if (!e.master && e.children) add('SHARED_NO_MASTER', sp, `si ${si}: ${e.children} cells without a master formula`);
        else if (e.master && e.ref) {
          const b = parseRef(e.ref);
          const at = parseRef(e.at);
          if (!b || (at && !(at.c1 >= b.c1 && at.c1 <= b.c2 && at.r1 >= b.r1 && at.r1 <= b.r2))) add('SHARED_REF', sp, `si ${si} master ${e.at} outside ref ${e.ref}`);
        }
      }
    }
    // tables of this sheet
    for (const [rid, r] of sRels) {
      if (r.type !== 'table' || !r.part) continue;
      const tp2 = byLower.get(r.part.toLowerCase());
      const txml = tp2 ? await text(tp2) : null;
      if (!tp2 || txml === null) continue;
      const tt = checkOrder(tp2, txml, 'table');
      if (!tt || tt.root !== 'table') { add('ROOT', tp2, `root element ${tt?.root}`); continue; }
      const nm = unesc(attr(tt.rootOpen, 'name') ?? '');
      const dn = unesc(attr(tt.rootOpen, 'displayName') ?? '');
      const id = attr(tt.rootOpen, 'id') ?? '';
      if (tableNames.has(dn.toLowerCase())) add('TABLE_NAME_DUP', tp2, `${dn} also in ${tableNames.get(dn.toLowerCase())}`);
      tableNames.set(dn.toLowerCase(), tp2);
      if (tableIds.has(id)) add('TABLE_ID_DUP', tp2, `id ${id} also in ${tableIds.get(id)}`);
      tableIds.set(id, tp2);
      if (!nm || !dn) add('TABLE_NAME_EMPTY', tp2, 'name / displayName');
      const box = parseRef(attr(tt.rootOpen, 'ref') ?? '');
      if (!box) { add('TABLE_REF', tp2, `ref ${attr(tt.rootOpen, 'ref')}`); continue; }
      tableBoxes.push({ sheet: sh.name, part: tp2, box, name: dn });
      const cols = [...txml.matchAll(/<(?:[\w.-]+:)?tableColumn\b[^>]*>/g)].map((m) => ({ id: attr(m[0], 'id') ?? '', name: unesc(attr(m[0], 'name') ?? '') }));
      const tc = tt.children.find((c) => c.name === 'tableColumns');
      const declared = Number(attr(tc?.open ?? '', 'count') ?? cols.length);
      if (declared !== cols.length) add('TABLE_COLS_COUNT', tp2, `tableColumns count=${declared} but ${cols.length} columns`);
      if (cols.length !== box.c2 - box.c1 + 1) add('TABLE_COLS_WIDTH', tp2, `${cols.length} tableColumns for ref ${attr(tt.rootOpen, 'ref')}`);
      const seenC = new Set<string>();
      const seenI = new Set<string>();
      for (const c of cols) {
        const k = c.name.toLowerCase();
        if (seenC.has(k)) add('TABLE_COL_DUP', tp2, `column name ${c.name} twice`);
        seenC.add(k);
        if (seenI.has(c.id)) add('TABLE_COL_ID_DUP', tp2, `column id ${c.id} twice`);
        seenI.add(c.id);
      }
      const headerRows = Number(attr(tt.rootOpen, 'headerRowCount') ?? 1);
      if (headerRows > 0) {
        cols.forEach((c, i) => {
          const t = cellText.get(`${box.r1},${box.c1 + i}`) ?? '';
          // Excel compares the header cell with the column name after turning line breaks into _x000a_-free text; plain equality is the rule
          if (t !== c.name && t.replace(/\r?\n/g, '') !== c.name.replace(/_x000[aA]_|\r?\n/g, '')) add('TABLE_HEADER', tp2, `column ${i + 1}: header cell "${t.slice(0, 30)}" vs name "${c.name.slice(0, 30)}" (row ${box.r1})`);
        });
      }
      for (const mb of mergeBoxes) if (overlap(mb, box)) add('TABLE_MERGE', tp2, `a merged area lies inside the table ${attr(tt.rootOpen, 'ref')}`);
      const af = /<(?:[\w.-]+:)?autoFilter\b[^>]*>/.exec(txml)?.[0];
      if (af) {
        const ab = parseRef(attr(af, 'ref') ?? '');
        const totals = Number(attr(tt.rootOpen, 'totalsRowCount') ?? 0);
        if (!ab) add('TABLE_FILTER_REF', tp2, `autoFilter ref ${attr(af, 'ref')}`);
        else if (ab.c1 !== box.c1 || ab.c2 !== box.c2 || ab.r1 !== box.r1 || ab.r2 !== box.r2 - totals) add('TABLE_FILTER_RANGE', tp2, `autoFilter ${attr(af, 'ref')} vs table ${attr(tt.rootOpen, 'ref')} (totals rows ${totals})`);
      }
      void rid;
    }
  }
  // tables must not overlap each other
  for (let i = 0; i < tableBoxes.length; i++) for (let j = i + 1; j < tableBoxes.length; j++) {
    if (tableBoxes[i].sheet === tableBoxes[j].sheet && overlap(tableBoxes[i].box, tableBoxes[j].box)) add('TABLE_OVERLAP', tableBoxes[i].part, `${tableBoxes[i].name} overlaps ${tableBoxes[j].name}`);
  }
  return out;
}

/** problems of `after` whose code does not occur in `before` (the input's own damage is not the patcher's) */
export function newProblems(before: LintProblem[], after: LintProblem[]): LintProblem[] {
  const had = new Set(before.map((p) => p.code));
  return after.filter((p) => !had.has(p.code));
}

export const formatProblems = (ps: LintProblem[], max = 6) => ps.slice(0, max).map((p) => `${p.code} ${p.part}: ${p.detail}`).join(' | ') + (ps.length > max ? ` (+${ps.length - max} more)` : '');
