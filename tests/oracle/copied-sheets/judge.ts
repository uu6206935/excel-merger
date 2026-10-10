/**
 * The judge of the copied-sheets oracle (task V9): pure functions over the Excel dumps (excel.ts / excel-book.ps1) and the reader's view of the
 * files. Nothing here starts Excel or reads the disk, so it is unit-tested with hand-made dumps (judge.test.ts: a planted fault must be found).
 *
 * What is judged, per scenario and per output (pre = the premerged package, js = replayLog, xl = the app's Excel engine):
 *   open-*             the package opens without repair
 *   names-*            the sheet names, in order, are EXACTLY what planSheetNames gives: merged sheet, the base's other sheets (when kept), then
 *                      every sheet of every participating file in card order (the expectation is computed here from the files, not taken from
 *                      the copier's layout)
 *   copy-visibility-*  visible / hidden / very hidden of a copy = that of its source sheet
 *   copy-value-*       every cell of every copy has the value Excel computes for the source sheet (after CalculateFull on both; 1904 <-> 1900 dates
 *                      shifted; a source with external links is judged leniently: copy-value-lenient is INFO)
 *   copy-formula-*     the Excel formula text of every cell equals the source's modulo the sheet prefixes (the references are renamed)
 *   copy-new-error-*   #REF! / #NAME? in a copy that the source sheet does not have
 *   copy-format-*      what Excel renders (number format, font, fill, alignment, borders, merge of 60 sampled cells), column widths, hidden rows /
 *                      columns, tab colour, conditional-format count, validation cells, hyperlinks, sheet-scoped names
 *   copy-font-name-*   (INFO) the font NAME of a cell whose font is a theme font: the target theme names another font for the script
 *   copy-object-*      charts / shapes / tables / notes / pivot tables / chart sheets (counts, names, series values, pivot values)
 *   copy-reader-*      the project's reader (tests/xlsx/addsheets-helpers compareSheets) sees a different cell / style / column / row / merge /
 *                      rule / hyperlink than in the source (pre and js outputs: the package patcher's own writing)
 *   merged-new-error   #REF! / #NAME? in the merged sheet of the premerged package beyond what the base and the imported sources have
 */
import type { SourceBook } from '../../../src/core/types';
import { planSheetNames } from '../../../src/core/xlsx/sheetnames';
import { compareSheets } from '../../xlsx/addsheets-helpers';
import { sameValue, type Finding, type ScenarioRec } from '../consistency/check';
import { norm } from '../consistency/excel';
import type { BookDump, BookSheet, Deep } from './excel';

export type Which = 'pre' | 'js' | 'xl';

export interface SourceRec {
  id: string;
  /** the name the app saw */
  fileName: string;
  /** the file on disk (what Excel opens) */
  path: string;
  mode: 'join' | 'copy';
  /** xlsx / xlsm: everything is compared; xls: values only (the app reads values, not formulas / styles); csv: through the reader only (Excel parses a csv by its own rules) */
  fidelity: 'full' | 'values' | 'reader';
  /** the sheets of the file as the reader lists them */
  sheets: string[];
  date1904: boolean;
  externalLinks: boolean;
}

export interface CopyInfo {
  sourceId: string;
  file: string;
  original: string;
  name: string;
  suffixed: boolean;
  complete: boolean;
}

export interface CopyRec extends ScenarioRec {
  combo: string;
  group: string;
  drop: boolean;
  ext: 'xlsx' | 'xlsm';
  baseFile: string;
  baseFileName: string;
  sources: SourceRec[];
  /** what the copier says it did (PremergeLayout.addedSheets) */
  copies: CopyInfo[];
  /** messages of the 'sheetCopyLimits' items: what the copier says it left out */
  limits: string[];
}

export interface Expected {
  /** every sheet of the output, in order */
  names: string[];
  copies: Array<{ src: SourceRec; original: string; name: string; type: number; visible: number }>;
  /** problems of the expectation itself (a source that did not open ...) */
  problems: string[];
}

const COPYABLE = new Set([-4167, -4109]);

/** the expected sheet list of the output, from the files alone (planSheetNames is the contract) */
export function expectedSheets(rec: CopyRec, srcDumps: ReadonlyMap<string, BookDump | null>, baseDump: BookDump | null): Expected {
  const problems: string[] = [];
  const keptBase: string[] = [];
  if (!rec.drop) {
    if (!baseDump || baseDump.open === 'failed') problems.push('the base file did not open in Excel');
    else for (const s of baseDump.sheets) if (COPYABLE.has(s.type) && s.name !== rec.mergedSheet) keptBase.push(s.name);
  }
  const requests: Array<{ file: string; sheets: string[] }> = [];
  const meta: Array<Array<{ type: number; visible: number }>> = [];
  for (const s of rec.sources) {
    const d = srcDumps.get(s.id);
    let names: string[];
    let info: Array<{ type: number; visible: number }>;
    if (s.fidelity === 'full') {
      if (!d || d.open === 'failed') {
        problems.push(`source ${s.fileName} did not open in Excel`);
        names = s.sheets;
        info = names.map(() => ({ type: -4167, visible: -1 }));
      } else {
        const list = d.sheets.filter((x) => COPYABLE.has(x.type));
        names = list.map((x) => x.name);
        info = list.map((x) => ({ type: x.type, visible: x.visible }));
      }
    } else {
      names = s.sheets;
      const list = d?.sheets ?? [];
      info = names.map((n) => {
        const x = list.find((y) => y.name === n);
        return { type: x?.type ?? -4167, visible: x?.visible ?? -1 };
      });
    }
    requests.push({ file: s.fileName, sheets: names });
    meta.push(info);
  }
  const planned = planSheetNames([rec.mergedSheet, ...keptBase], requests);
  const copies: Expected['copies'] = [];
  let k = 0;
  rec.sources.forEach((src, i) => {
    requests[i].sheets.forEach((original, j) => {
      const p = planned[k++];
      copies.push({ src, original, name: p.name, type: meta[i][j].type, visible: meta[i][j].visible });
    });
  });
  return { names: [rec.mergedSheet, ...keptBase, ...copies.map((c) => c.name)], copies, problems };
}

// ---------------------------------------------------------------------------------------------------------------------
// helpers
// ---------------------------------------------------------------------------------------------------------------------

const PREFIX = /('([^']|'')*'|[^\s'!=+\-*/^&<>(),:;{}[\]]+)!/g;
/** a formula text modulo sheet prefixes (the references of a copy are renamed) and "_2" table names */
export function stripPrefix(f: string): string {
  return f.replace(PREFIX, '').replace(/_\d+(?=\[)/g, '').replace(/\s+/g, '');
}

const isFatal = (v: unknown): boolean => v === '#REF!' || v === '#NAME?';
const VOLATILE = /\b(TODAY|NOW|RAND|RANDBETWEEN|RANDARRAY|OFFSET|INDIRECT|CELL|INFO)\s*\(/i;
const at = <T,>(rows: T[][] | undefined, r: number, c: number): T | undefined => rows?.[r - 1]?.[c - 1];

function isDateFormat(nf: string | undefined): boolean {
  if (!nf) return false;
  const t = nf.replace(/"[^"]*"|\[[^\]]*\]|\\./g, '');
  return /[ymdhs]/i.test(t) && !/^General$/i.test(t);
}

class Limit {
  constructor(private readonly max = 6) {}
  n = 0;
  examples: string[] = [];
  add(s: string): void {
    this.n++;
    if (this.examples.length < this.max) this.examples.push(s);
  }
  text(): string {
    return `${this.n} difference(s): ${this.examples.join(' ; ')}`;
  }
}

export interface PairOptions {
  /** the source's date system differs from the output's: dates move by 1462 days */
  dateShift: number;
  /** the source is judged leniently (external links: Excel recomputes what the copy holds as a stored value) */
  lenient: boolean;
  fidelity: SourceRec['fidelity'];
  /** the copier says this sheet was copied completely (formulas / styles all kept) */
  complete: boolean;
  /** what the copier says it left out of this sheet (lower-cased messages naming it) */
  limits: string;
  isDateAt: (r: number, c: number) => boolean;
  /** (compareDeep) the cell holds nothing in either sheet: only what can be seen on an empty cell (fill, borders, merge) is compared */
  blankAt?: (r: number, c: number) => boolean;
  /** tables of the copy that were renamed (<name>_2 ...) because the name was taken: copy name -> source name (workbook-wide: a formula may read a table of another sheet) */
  tableRenames?: ReadonlyMap<string, string>;
}

interface Issue {
  cls: string;
  detail: string;
}

/** compares one copied sheet with its source sheet, as Excel dumped both */
export function comparePair(src: BookSheet, out: BookSheet, o: PairOptions): Issue[] {
  const issues: Issue[] = [];
  const add = (cls: string, detail: string) => issues.push({ cls, detail });
  // a table of a later copy of the same file is named <name>_2 (3 ...) and the structured references of the copy follow: compared modulo that
  const renames: Array<[string, string]> = [...(o.tableRenames ?? [])];
  (src.deep?.tables ?? []).forEach((t, i) => {
    const n = out.deep?.tables?.[i]?.n;
    if (n && n !== t.n && n.replace(/_\d+$/, '') === t.n) renames.push([n, t.n]);
  });
  const esc = (s: string) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const tn = (f: string): string => renames.reduce((x, [from, to]) => x.replace(new RegExp(`(?<![\\p{L}\\p{N}_.])${esc(from)}(?![\\p{L}\\p{N}_])`, 'gu'), to), f);
  if (src.type !== out.type) {
    add('copy-object', `sheet type ${src.type} -> ${out.type}`);
    return issues;
  }
  if (src.visible !== out.visible) add('copy-visibility', `${src.visible} (source) -> ${out.visible}`);
  // ---- values / formulas / errors
  const rows = Math.max(src.rows, out.rows);
  const cols = Math.max(src.cols, out.cols);
  const vals = new Limit();
  const fmls = new Limit();
  const f2v = new Limit();
  const errs = new Limit();
  const col = (c: number): string => {
    let s = '';
    for (let n = c; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
    return s;
  };
  let newErr = 0;
  for (let r = 1; r <= rows; r++) {
    for (let c = 1; c <= cols; c++) {
      const a = norm(at(src.v, r, c));
      const b = norm(at(out.v, r, c));
      const fa = at(src.f, r, c) ?? '';
      const fb = at(out.f, r, c) ?? '';
      if (isFatal(b) && !isFatal(a)) {
        newErr++;
        errs.add(`${col(c)}${r}: ${JSON.stringify(a)} [${fa}] -> ${b} [${fb}]`);
      }
      if (VOLATILE.test(fa)) continue;
      let same = sameValue(a, b);
      if (!same && o.dateShift && typeof a === 'number' && typeof b === 'number' && o.isDateAt(r, c)) same = sameValue(a + o.dateShift, b);
      if (!same) vals.add(`${col(c)}${r}: ${JSON.stringify(a)} [${fa}] -> ${JSON.stringify(b)} [${fb}]`);
      if (o.fidelity === 'full') {
        if (fa && !fb) f2v.add(`${col(c)}${r}: ${fa} -> value ${JSON.stringify(b)}`);
        else if ((fa || fb) && stripPrefix(fa) !== stripPrefix(tn(fb))) fmls.add(`${col(c)}${r}: ${fa} -> ${fb}`);
      }
    }
  }
  if (newErr) add('copy-new-error', `${newErr} #REF!/#NAME? cell(s) the source does not have: ${errs.examples.join(' ; ')}`);
  if (vals.n) add(o.lenient ? 'copy-value-lenient' : 'copy-value', vals.text());
  if (fmls.n) add(o.complete ? 'copy-formula' : 'copy-formula-lenient', fmls.text());
  if (f2v.n) add(o.complete && !o.lenient ? 'copy-formula-value' : 'copy-formula-lenient', f2v.text());
  if (o.fidelity !== 'full') return issues;
  if (src.deepError || out.deepError) add('copy-format', `deep dump failed: ${src.deepError ?? out.deepError}`);
  // a cell with nothing in it takes the DEFAULT style of its workbook (vertical alignment, font ...): that differs between two workbooks and shows nothing
  const blank = (s: BookSheet, r: number, c: number): boolean => {
    const v = norm(at(s.v, r, c));
    return (v === null || v === '') && !(at(s.f, r, c) ?? '');
  };
  if (src.deep && out.deep) issues.push(...compareDeep(src.deep, out.deep, { ...o, blankAt: (r, c) => blank(src, r, c) && blank(out, r, c) }));
  else if (src.deep || out.deep) add('copy-format', `deep dump missing for the ${src.deep ? 'copy' : 'source'}`);
  return issues;
}

function eqSet(a: number[] | undefined, b: number[] | undefined): boolean {
  return JSON.stringify(a ?? []) === JSON.stringify(b ?? []);
}

export function compareDeep(a: Deep, b: Deep, o: PairOptions): Issue[] {
  const issues: Issue[] = [];
  const add = (cls: string, detail: string) => issues.push({ cls, detail });
  if (a.tab !== b.tab) add('copy-format', `tab colour ${a.tab} -> ${b.tab}`);
  // column widths: points, what the column HAS (the copy keeps the pixel width; 1.6 pt = 2 px)
  if (a.widths && b.widths) {
    const bad: string[] = [];
    for (let i = 0; i < Math.min(a.widths.length, b.widths.length); i++) if (Math.abs(a.widths[i] - b.widths[i]) > 1.6) bad.push(`col ${i + 1}: ${a.widths[i]}pt -> ${b.widths[i]}pt`);
    if (bad.length) add('copy-format', `column width: ${bad.length} column(s): ${bad.slice(0, 5).join(' ; ')}`);
  }
  if (!eqSet(a.hiddenCols, b.hiddenCols)) add('copy-format', `hidden columns ${JSON.stringify(a.hiddenCols)} -> ${JSON.stringify(b.hiddenCols)}`);
  if (!eqSet(a.hiddenRows, b.hiddenRows)) add('copy-format', `hidden rows ${JSON.stringify(a.hiddenRows)} -> ${JSON.stringify(b.hiddenRows)}`);
  // sampled cells
  if (a.look && b.look) {
    const fmt = new Limit();
    const fonts = new Limit();
    const valign = new Limit();
    const key = (l: { r: number; c: number }) => `${l.r},${l.c}`;
    const bm = new Map(b.look.map((l) => [key(l), l]));
    for (const x of a.look) {
      const y = bm.get(key(x));
      if (!y) continue;
      const diff: string[] = [];
      const empty = o.blankAt?.(x.r, x.c) ?? false;
      for (const k of ['nf', 'fs', 'fb', 'fi', 'fc', 'ic', 'ip', 'ha', 'va', 'wr', 'mg', 'bb', 'bl', 'bt', 'br'] as const) {
        if (empty && !/^(ic|ip|mg|bb|bl|bt|br)$/.test(k)) continue;
        const p = x[k];
        const q = y[k];
        // the vertical alignment of a cell that says nothing about it is the one of the workbook's Normal style: a Japanese Excel gives a file without
        // styles "center", a file with its own Normal style "bottom" — it shows only in a row taller than its text
        if (k === 'va') {
          if (p !== q) valign.add(`R${x.r}C${x.c}: ${p}->${q}`);
          continue;
        }
        if (k === 'fc' || k === 'ic') {
          // colours: the same RGB (a colour that is "automatic" reads as a fixed number in both)
          if (p !== q) diff.push(`${k} ${p}->${q}`);
        } else if (p !== q) diff.push(`${k} ${JSON.stringify(p)}->${JSON.stringify(q)}`);
      }
      if (x.fn !== y.fn && !empty) fonts.add(`R${x.r}C${x.c}: ${x.fn} -> ${y.fn}`);
      if (diff.length) fmt.add(`R${x.r}C${x.c}: ${diff.join(', ')}`);
    }
    if (fmt.n) add('copy-format', `cell format: ${fmt.text()}`);
    if (fonts.n) add('copy-font-name', `font name: ${fonts.text()}`);
    if (valign.n) add('copy-valign', `default vertical alignment: ${valign.text()}`);
  }
  if ((a.cf ?? -1) !== (b.cf ?? -1)) add('copy-format', `conditional formats ${a.cf} -> ${b.cf}`);
  if ((a.dv ?? '') !== (b.dv ?? '')) add('copy-format', `validation cells ${a.dv} -> ${b.dv}`);
  if ((a.links ?? 0) !== (b.links ?? 0)) add('copy-format', `hyperlinks ${a.links} -> ${b.links}`);
  // ---- objects (a kind that the copier says it left out of this sheet is reported, not failed)
  const left = (re: RegExp): boolean => re.test(o.limits);
  const obj = (what: string, re: RegExp, x: string, y: string) => {
    if (x === y) return;
    add(left(re) ? 'copy-object-leftout' : 'copy-object', `${what}: ${x} -> ${y}`);
  };
  const sh = (l: Deep['shapes']) => (l ?? []).map((s) => `${s.t}:${s.n}:${s.x}`).sort().join('|');
  obj(`shapes (${a.shapes?.length ?? 0} -> ${b.shapes?.length ?? 0})`, /図形|画像|ピクチャ|テキストボックス|SmartArt|インク|グループ|描画|ドロー|drawing/i, sh(a.shapes), sh(b.shapes));
  obj('other shapes (notes / form controls / slicers)', /スライサー|タイムライン|コントロール|ActiveX|OLE|コメント|メモ/i, String(a.shapesOther ?? 0), String(b.shapesOther ?? 0));
  const chart = (l: Deep['charts']) => (l ?? []).map((c) => `${c.type}[${c.series.map((s) => `${s.v}/${s.x}`).join(',')}]`).join('|');
  obj(`charts (${a.charts?.length ?? 0} -> ${b.charts?.length ?? 0})`, /グラフ|チャート/i, chart(a.charts), chart(b.charts));
  const tb = (l: Deep['tables']) => (l ?? []).map((t) => `${t.n.replace(/_\d+$/, '')}@${t.a}[${t.cols}]${t.totals}`).join('|');
  obj(`tables (${a.tables?.length ?? 0} -> ${b.tables?.length ?? 0})`, /テーブル/i, tb(a.tables), tb(b.tables));
  const nt = (l: Deep['notes']) => (l ?? []).map((n) => `${n.a}:${n.t}`).join('|');
  obj(`notes (${a.notes?.length ?? 0} -> ${b.notes?.length ?? 0})`, /コメント|メモ|スレッド/i, nt(a.notes), nt(b.notes));
  const pv = (l: Deep['pivots']) => (l ?? []).map((p) => `${p.n.replace(/_\d+$/, '')}@${p.a}<${stripPrefix(p.src).replace(/_\d+$/, '')}>${p.v}`).sort().join('|');
  obj(`pivot tables (${a.pivots?.length ?? 0} -> ${b.pivots?.length ?? 0})`, /ピボット/i, pv(a.pivots), pv(b.pivots));
  if (a.chartSheet || b.chartSheet) obj('chart sheet', /グラフ|チャート/i, JSON.stringify(a.chartSheet ?? null), JSON.stringify(b.chartSheet ?? null));
  return issues;
}

// ---------------------------------------------------------------------------------------------------------------------
// the judge entry points
// ---------------------------------------------------------------------------------------------------------------------

export interface Judged {
  findings: Finding[];
  /** counters for the report */
  stats: { sheetsCompared: number; cellsCompared: number; formatCells: number };
}

/** the sheet list of one output */
export function judgeNames(rec: CopyRec, which: Which, out: BookDump, exp: Expected): Finding[] {
  const found: Finding[] = [];
  const add = (cls: string, detail: string) => found.push({ scenario: rec.id, cls: `${cls}-${which}`, detail });
  for (const p of exp.problems) add('expectation', p);
  if (out.open !== 'normal') {
    add('open', `${out.open}: ${out.log}`);
    return found;
  }
  const got = out.sheets.map((s) => s.name);
  if (JSON.stringify(got) !== JSON.stringify(exp.names)) add('names', `expected ${JSON.stringify(exp.names)} but the workbook has ${JSON.stringify(got)}`);
  return found;
}

export function judgeCopies(rec: CopyRec, which: Which, out: BookDump, exp: Expected, srcDumps: ReadonlyMap<string, BookDump | null>, srcBooks: ReadonlyMap<string, SourceBook>): Judged {
  const findings: Finding[] = [];
  const stats = { sheetsCompared: 0, cellsCompared: 0, formatCells: 0 };
  if (out.open !== 'normal') return { findings, stats };
  // tables renamed in the copies (<name>_2 ...): workbook-wide, a formula can read the table of another sheet
  const tableRenamesOf = (srcDump: BookDump): Map<string, string> => {
    const srcNames = new Set(srcDump.sheets.flatMap((x) => (x.deep?.tables ?? []).map((t) => t.n)));
    const m = new Map<string, string>();
    for (const sh of out.sheets) for (const t of sh.deep?.tables ?? []) if (/_\d+$/.test(t.n) && !srcNames.has(t.n) && srcNames.has(t.n.replace(/_\d+$/, ''))) m.set(t.n, t.n.replace(/_\d+$/, ''));
    return m;
  };
  const limitsOf = (c: CopyRec['copies'][number] | { original: string; name: string }) =>
    rec.limits.filter((m) => m.includes(`「${c.original}」`) || m.includes(`「${c.name}」`)).join(' ').toLowerCase();
  for (const c of exp.copies) {
    const outSheet = out.sheets.find((s) => s.name === c.name);
    const srcDump = srcDumps.get(c.src.id);
    const srcSheet = srcDump?.sheets.find((s) => s.name === c.original);
    if (!outSheet) continue; // reported by judgeNames
    // a csv is parsed by Excel with its own rules (leading zeros, dates): only the reader's view of it is the truth (judgeCopiesByReader)
    if (c.src.fidelity === 'reader') continue;
    if (!srcSheet || !srcDump) {
      findings.push({ scenario: rec.id, cls: `copy-source-missing-${which}`, detail: `${c.src.fileName} / ${c.original}: the source sheet is not in Excel's view of the source file` });
      continue;
    }
    const sb = srcBooks.get(c.src.id)?.sheets.find((s) => s.name === c.original);
    const book = srcBooks.get(c.src.id);
    const dateShift = srcDump.date1904 === out.date1904 ? 0 : srcDump.date1904 ? 1462 : -1462;
    const issues = comparePair(srcSheet, outSheet, {
      dateShift,
      lenient: c.src.externalLinks,
      fidelity: c.src.fidelity,
      complete: rec.copies.find((x) => x.name === c.name)?.complete ?? true,
      limits: limitsOf(c),
      isDateAt: (r, col) => isDateFormat(book?.styles[sb?.rows[r - 1]?.[col - 1]?.s ?? 0]?.numFmt),
      tableRenames: tableRenamesOf(srcDump),
    });
    stats.sheetsCompared++;
    stats.cellsCompared += Math.max(srcSheet.rows, outSheet.rows) * Math.max(srcSheet.cols, outSheet.cols);
    stats.formatCells += srcSheet.deep?.look?.length ?? 0;
    const byCls = new Map<string, string[]>();
    for (const i of issues) byCls.set(i.cls, [...(byCls.get(i.cls) ?? []), i.detail]);
    for (const [cls, details] of byCls) findings.push({ scenario: rec.id, cls: `${cls}-${which}`, detail: `${c.src.fileName} / ${c.original} -> ${c.name}: ${details.join(' | ')}`.slice(0, 700) });
  }
  return { findings, stats };
}

interface ParsedName {
  scope: string | null;
  name: string;
  refers: string;
}
const SCOPED = /^(?:'((?:[^']|'')+)'|([^'!]+))!(.+)$/;
/** the defined names of a dump: workbook level (scope null) and sheet level (Excel prints those as "Sheet!Name"); Excel's own hidden _xlfn.* / _xlnm.* names left out */
export function parseNames(d: BookDump | null): ParsedName[] {
  const out: ParsedName[] = [];
  for (const n of d?.names ?? []) {
    if (/^_xl(fn|pm)\./i.test(n.name) || /_FilterDatabase$/i.test(n.name)) continue;
    const m = SCOPED.exec(n.name);
    out.push(m ? { scope: (m[1] ?? m[2]).replace(/''/g, "'"), name: m[3], refers: n.refers } : { scope: null, name: n.name, refers: n.refers });
  }
  return out;
}

/**
 * Defined names of the copies: every sheet-scoped name of a source sheet is there again on its copy (same text modulo the renamed sheets), and a
 * name the copy has beyond that is a workbook-level name of the SOURCE that the sheet reads (copied as a sheet-scoped name of the copy). No name of a
 * source becomes workbook-level in the output (it would collide with the base's names and with the other files').
 */
export function judgeCopyNames(rec: CopyRec, which: Which, out: BookDump, exp: Expected, srcDumps: ReadonlyMap<string, BookDump | null>, baseDump: BookDump | null): Finding[] {
  const found: Finding[] = [];
  if (out.open !== 'normal') return found;
  const add = (cls: string, detail: string) => found.push({ scenario: rec.id, cls: `${cls}-${which}`, detail: detail.slice(0, 600) });
  const mine = parseNames(out);
  for (const c of exp.copies) {
    const sd = srcDumps.get(c.src.id);
    if (!sd || c.src.fidelity !== 'full') continue;
    const src = parseNames(sd);
    const sScoped = src.filter((n) => n.scope === c.original);
    const sGlobal = src.filter((n) => n.scope === null);
    const oScoped = mine.filter((n) => n.scope === c.name);
    const lost = sScoped.filter((n) => !oScoped.some((m) => m.name === n.name && stripPrefix(m.refers) === stripPrefix(n.refers)));
    if (lost.length) add('copy-names', `${c.src.fileName} / ${c.original} -> ${c.name}: sheet-scoped name(s) not carried over: ${lost.map((n) => `${n.name} ${n.refers}`).join(' ; ')}`);
    const extra = oScoped.filter((m) => !sScoped.some((n) => n.name === m.name) && !sGlobal.some((n) => n.name === m.name && stripPrefix(n.refers) === stripPrefix(m.refers)));
    if (extra.length) add('copy-names', `${c.src.fileName} / ${c.original} -> ${c.name}: name(s) the source does not have: ${extra.map((n) => `${n.name} ${n.refers}`).join(' ; ')}`);
  }
  const allowed = new Set(parseNames(baseDump).filter((n) => n.scope === null).map((n) => n.name.toLowerCase()));
  const leaked = mine.filter((n) => n.scope === null && !allowed.has(n.name.toLowerCase()));
  if (leaked.length) add('copy-names-leak', `workbook-level name(s) that the base does not have: ${leaked.map((n) => `${n.name} ${n.refers}`).join(' ; ')}`);
  return found;
}

/** the project's reader: copy vs source, cell by cell, with resolved styles (the package patcher's own writing: pre and js outputs) */
export function judgeCopiesByReader(rec: CopyRec, which: Which, out: SourceBook, exp: Expected, srcBooks: ReadonlyMap<string, SourceBook>): Finding[] {
  const found: Finding[] = [];
  for (const src of rec.sources) {
    const book = srcBooks.get(src.id);
    if (!book) continue;
    const mine = exp.copies.filter((c) => c.src.id === src.id);
    const rename = new Map(mine.map((c) => [c.original.toLowerCase(), c.name]));
    for (const c of mine) {
      const info = rec.copies.find((x) => x.name === c.name);
      if (!out.sheets.some((s) => s.name === c.name)) continue; // judgeNames
      if (!book.sheets.some((s) => s.name === c.original)) continue; // a chart sheet: the reader has no grid for it
      const shift = book.date1904 === out.date1904 ? 0 : book.date1904 ? 1462 : -1462;
      const base = { rename, styles: true, allowValueFormulas: info ? !info.complete : true };
      let diffs = compareSheets(book, c.original, out, c.name, { ...base, dateShift: 0 });
      // between the date systems only the cells that SHOW a date move by 1462 days: a cell is a real difference when it differs with and without the shift
      if (shift && diffs.length) {
        const shifted = compareSheets(book, c.original, out, c.name, { ...base, dateShift: shift });
        diffs = diffs.filter((d0) => shifted.some((d) => d.where === d0.where));
      }
      // a table of a later copy of the same file is named <name>_2: the structured references follow
      const noSuffix = (f: string) => f.replace(/_\d+(?=\[)/g, '');
      diffs = diffs.filter((d) => {
        const m = /^formula ([\s\S]*) != ([\s\S]*)$/.exec(d.detail);
        return !(m && noSuffix(m[1]) === noSuffix(m[2]));
      });
      if (diffs.length) found.push({ scenario: rec.id, cls: `copy-reader-${which}`, detail: `${src.fileName} / ${c.original} -> ${c.name}: ${diffs.slice(0, 4).map((d) => `${d.where}: ${d.detail}`).join(' ; ')}${diffs.length > 4 ? ` (+${diffs.length - 4})` : ''}`.slice(0, 700) });
    }
  }
  return found;
}

/** the merged sheet of the premerged package must not hold #REF! / #NAME? beyond what the base and the imported sources have */
export function judgeMergedErrors(rec: CopyRec, pre: BookDump, baseDump: BookDump | null, srcDumps: ReadonlyMap<string, BookDump | null>): Finding[] {
  if (pre.open !== 'normal') return [];
  const count = (s: BookSheet | undefined): number => {
    let n = 0;
    for (const row of s?.v ?? []) for (const v of row) if (isFatal(norm(v))) n++;
    return n;
  };
  const m = pre.sheets.find((s) => s.name === rec.mergedSheet);
  if (!m) return [];
  const have = count(m);
  const allow = count(baseDump?.sheets.find((s) => s.name === rec.mergedSheet)) + rec.sources.filter((s) => s.mode === 'join').reduce((n, s) => n + (srcDumps.get(s.id)?.sheets.reduce((k, x) => k + count(x), 0) ?? 0), 0);
  if (have > allow) return [{ scenario: rec.id, cls: 'merged-new-error', detail: `the merged sheet holds ${have} #REF!/#NAME? cell(s), the base + the joined sources only ${allow}` }];
  return [];
}

/**
 * Classes that are REPORTED (RESULTS.md counts them) but are not failures:
 *   copy-font-name      the font NAME of a theme-font cell differs (the target theme names another font for the script; documented gap of A1)
 *   copy-value-lenient  a source with external links: Excel recomputes the link, the copy holds the stored result (documented)
 *   copy-formula-lenient formulas the copier turned into values on purpose (3-D, external, structured references of a table the copier did not copy; warned)
 *   copy-object-leftout an object kind that the 'sheetCopyLimits' warning of the copier names for this sheet
 *   skipped / runner-fallback / circular / convert-blank / cf-quirk / ... : the classes of the consistency oracle
 */
export const INFO_CLASSES = /^(copy-font-name|copy-valign|copy-value-lenient|copy-formula-lenient|copy-object-leftout|skipped|runner-fallback|convert-blank|cf-quirk|circular|ref-error-announced|pivot-value-changed|pivot-engine-package-(cols|widths)|info-)/;
