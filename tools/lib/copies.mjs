// The sheets copied from the other files (plan A, "他のファイルのシートも含める") in a verdict — used by tools/check-output.mjs --with-copies
// (owner: T). Pure: works on inspect.ps1 dumps (tools/lib/inspect.mjs) of the BASE file, the OUTPUT and every SOURCE file.
//
// What a correct output looks like:
//   * the merged sheet first, then the base's other sheets (when the user kept them), then the sheets of every other file in card order, each
//     named "<file name>_<sheet name>" (tools/lib/sheetnames.mjs = planSheetNames: 31 characters, "(1)" for a name that is taken);
//   * every copy equals its source sheet: values (after recalculation), formulas (modulo the renamed sheet references), merges, hidden state,
//     hidden columns, conditional formats, validations, tables, charts, pivot tables, pictures / shapes, notes, hyperlinks, freeze panes, filter.
// Failures: copy-missing / copy-unexpected / copy-order (the names), copy-differs (a copy is not its source), copy-new-error (#REF! / #NAME? the
// source does not have). Info: copied-sheets (the list), copy-leftout (what the source sheet has that the tool knows is not copied: slicers, form
// controls ... see README.md), copy-date-system (a 1904 <-> 1900 shift of dates).
import { arr, compareInspections, sheetByName } from './compare-core.mjs';
import { planSheetNames } from './sheetnames.mjs';

const PREFIX = /('([^']|'')*'|[^\s'!=+\-*/^&<>(),:;{}[\]]+)!/g;
/** a formula text modulo sheet prefixes (the references of a copy are renamed) and "_2" table names */
export function stripPrefix(f) {
  return String(f ?? '').replace(PREFIX, '').replace(/_\d+(?=\[)/g, '').replace(/\s+/g, '');
}

const DATE_SHIFT = 1462;

/**
 * How far a source can be compared with Excel's reading of it: the app reads `.xls` for VALUES only (no formulas, formats, rules), and a text file (`.csv`) by its own
 * rules (text stays text, a leading zero stays) while Excel parses it as numbers / dates — so for those only the names, the visibility and the values of an .xls are compared.
 */
export function sourceKind(fileName) {
  if (/\.(csv|tsv|txt)$/i.test(String(fileName))) return 'text';
  if (/\.xls$/i.test(String(fileName))) return 'values';
  return 'full';
}

/**
 * The sheet name the app gives a text file's only sheet: the file name without its extension, forbidden characters turned into "_", at most 31 characters (the reader's
 * own rule). Excel names that sheet after the file it opened, which can be another name (a temporary copy of the file), so the app's rule is used.
 */
export function textSheetName(fileName) {
  const f = String(fileName).replace(/^.*[\\/]/, '');
  const i = f.lastIndexOf('.');
  let s = (i > 0 ? f.slice(0, i) : f).replace(/[\\/?*[\]:]/g, '_').replace(/^'+|'+$/g, '').trim();
  if (!s) s = 'Sheet';
  return s.length > 31 ? s.slice(0, 31) : s;
}

/** the sheet names of a SOURCE: the dump's, except a text file's only sheet, which has the app's name (see textSheetName) */
function sourceSheetNames(s) {
  const own = sheetNamesOf(s.dump);
  return sourceKind(s.name) === 'text' && own.length === 1 ? [textSheetName(s.name)] : own;
}

/** the names of the sheets of an inspect dump, in workbook order (worksheets by index, chart sheets after them: the dump does not say where they are) */
export function sheetNamesOf(dump) {
  return [...arr(dump?.sheets).slice().sort((a, b) => (a.index ?? 0) - (b.index ?? 0)).map((s) => s.name), ...arr(dump?.chartSheets)];
}

/**
 * What the copies of `sources` must be called in an output whose first sheet is `merged` and which keeps `keptBase` of the base's sheets.
 * @param {{name: string, dump: any}[]} sources  every other file in card order (name = the file name as the app saw it)
 * @returns {{file: string, original: string, name: string, suffixed: boolean, dump: any}[]}
 */
export function expectedCopies(merged, keptBase, sources) {
  const planned = planSheetNames([merged, ...keptBase], sources.map((s) => ({ file: s.name, sheets: sourceSheetNames(s) })));
  let k = 0;
  const out = [];
  for (const s of sources) for (const original of sourceSheetNames(s)) out.push({ ...planned[k++], original, dump: s.dump, strict: s.strict === true });
  return out;
}

/** the sheet `s` under its new name, with every text that names a sheet (formulas, rules, series, pivot sources) modulo the prefixes and "_2" table names */
function renamed(s, name) {
  const sp = (x) => (x == null ? x : stripPrefix(x));
  return {
    ...s,
    name,
    values: { ...(s.values ?? {}) },
    // (a blank cell has the DEFAULT style of its workbook — font size, vertical alignment —, which differs between two workbooks and shows nothing: only cells with something in them are compared)
    probes: Object.fromEntries(Object.entries(s.probes ?? {}).filter(([addr]) => addr in (s.values ?? {}) || addr in (s.formulas ?? {}) || addr in (s.errors ?? {}))),
    formulas: Object.fromEntries(Object.entries(s.formulas ?? {}).map(([k, f]) => [k, stripPrefix(f)])),
    cf: arr(s.cf).map((c) => ({ ...c, formula1: sp(c.formula1), formula2: sp(c.formula2) })),
    dv: arr(s.dv).map((d) => ({ ...d, formula1: sp(d.formula1), formula2: sp(d.formula2) })),
    tables: arr(s.tables).map((t) => ({ ...t, name: String(t.name).replace(/_\d+$/, '') })),
    charts: arr(s.charts).map((c) => ({ ...c, series: arr(c.series).map((x) => ({ ...x, formula: sp(x.formula) })) })),
    // (a pivot table name is unique in a workbook: Excel renames the later one "..._2")
    pivots: arr(s.pivots).map((p) => ({ ...p, name: String(p.name).replace(/_\d+$/, ''), source: sp(p.source) })),
    // what the copier knows it cannot copy is judged apart: shapes / notes (counts), hyperlinks
    shapes: 0,
    comments: 0,
    threadedComments: 0,
  };
}

// (usedRange: Google Sheets writes 1000 empty <row/> — the copy keeps the 33 rows that hold something; the used range follows)
const GROUPS_SKIPPED = new Set(['open', 'sheets', 'names', 'sheet', 'cols', 'usedRange']);

/**
 * @param {{base: any, out: any, merged: string, sources: {name: string, dump: any, strict?: boolean}[], outStrict?: boolean, add: Function, maxExamples?: number}} ctx
 */
export function checkCopies(ctx) {
  const { base, out, merged, sources, add } = ctx;
  const outNames = sheetNamesOf(out);
  const baseNames = sheetNamesOf(base);
  const keptBase = baseNames.filter((n) => n !== merged && outNames.includes(n));
  const usable = sources.filter((s) => s.dump && s.dump.sheets);
  for (const s of sources) if (!s.dump || !s.dump.sheets) add('warn', 'copy-source', `元ファイル「${s.name}」を Excel で読めなかったため、そのシートのコピーは確認できません（${s.dump?.error ?? s.dump?.openMode ?? '?'}）`);
  const exp = expectedCopies(merged, keptBase, usable);
  const expNames = exp.map((e) => e.name);
  const known = new Set([merged, ...keptBase, ...expNames]);
  // ---- names
  const missing = expNames.filter((n) => !outNames.includes(n));
  const unexpected = outNames.filter((n) => !known.has(n) && !baseNames.includes(n));
  if (missing.length) add('fail', 'copy-missing', `他のファイルのシートが出力にありません: ${missing.join(', ')}`);
  if (unexpected.length) add('fail', 'copy-unexpected', `命名規則（ファイル名_シート名、重なったら (1) (2)…）にないシートが出力にあります: ${unexpected.join(', ')}（期待: ${expNames.join(', ')}）`);
  const present = outNames.filter((n) => expNames.includes(n));
  const wanted = expNames.filter((n) => outNames.includes(n));
  const order = sheetNamesOf(out).filter((n) => n === merged || keptBase.includes(n) || expNames.includes(n));
  const wantOrder = [merged, ...keptBase, ...expNames].filter((n) => outNames.includes(n));
  // (chart sheets are only known by name, not by position: the order is judged on the worksheets)
  const ws = (names, d) => names.filter((n) => !arr(d.chartSheets).includes(n));
  if (JSON.stringify(ws(present, out)) !== JSON.stringify(ws(wanted, out)) || JSON.stringify(ws(order, out)) !== JSON.stringify(ws(wantOrder, out)))
    add('fail', 'copy-order', `シートの並びが違います（結合したシート、基準ファイルのほかのシート、他のファイルのシート（カードの順）の順）: ${order.join(' | ')}（期待: ${wantOrder.join(' | ')}）`);
  if (!missing.length && !unexpected.length) add('pass', 'copy-names', `他のファイルのシート ${expNames.length} 枚が、決まりの名前で並んでいます`);
  if (expNames.length)
    add('info', 'copied-sheets', `他のファイルのシート ${expNames.length} 枚: ${exp.map((e) => `${e.name}${e.suffixed ? '（重なったため番号つき）' : ''}`).join(', ')}`);
  // ---- every copy equals its source
  const problems = [];
  const leftOut = [];
  let compared = 0;
  const newErrors = [];
  const textFiles = new Set();
  for (const e of exp) {
    const so = sheetByName(out, e.name);
    const kind = sourceKind(e.file);
    // (a text file's only sheet is the first one of its dump, whatever Excel called it)
    const ss = kind === 'text' ? arr(e.dump.sheets)[0] : sheetByName(e.dump, e.original);
    if (!so || !ss) continue; // names: reported above (chart sheets have no cell dump)
    if (kind === 'text') {
      textFiles.add(e.file);
      if (ss.visible !== so.visible) problems.push(`[${e.name}] 表示状態 ${ss.visible} → ${so.visible}`);
      compared++;
      continue;
    }
    compared++;
    const shiftBy = !!out.date1904 !== !!e.dump.date1904 ? (e.dump.date1904 ? DATE_SHIFT : -DATE_SHIFT) : 0;
    const a = renamed(ss, e.name);
    const b = renamed(so, e.name);
    // dates between the two date systems move by 1462 days (a number that differs by exactly that is the same date)
    if (shiftBy) {
      for (const [k, v] of Object.entries(a.values)) if (typeof v === 'number' && typeof b.values[k] === 'number' && Math.abs(b.values[k] - (v + shiftBy)) < 1e-9) a.values[k] = b.values[k];
    }
    // a formula the copier turned into its stored value (3-D / external / structured reference it could not keep) differs only in the text:
    // the value must still be the source's
    const rows = compareInspections({ openMode: 'normal', sheets: [a], names: [] }, { openMode: 'normal', sheets: [b], names: [] }, { pivotLayout: false });
    // charts / pictures / shapes do not go between a Strict and a normal workbook (the app says so in its warnings: a known limit, a warning here too)
    const strictGap = e.strict !== (ctx.outStrict === true);
    for (const r of rows) {
      if (r.ok || r.filtered || GROUPS_SKIPPED.has(r.group)) continue;
      if (strictGap && r.group === 'charts') {
        leftOut.push(`${e.name}: グラフ ${arr(ss.charts).length} → ${arr(so.charts).length}（Strict 形式と通常の形式の間では写せません。アプリの警告に出ます）`);
        continue;
      }
      if (kind === 'values' && r.group !== 'values') continue; // (the app reads an .xls for its values only)
      problems.push(`[${e.name}] ${r.group}: ${String(r.detail).slice(0, 220)}`);
    }
    // pictures / shapes / notes: the copier does not copy slicers, timelines, form controls / ActiveX / OLE, SmartArt, ink (README) and turns threaded
    // comments into notes — a different count is a warning, not a failure; hyperlinks must be the same
    if (kind === 'values') {
      const eb0 = Object.keys(so.errors ?? {}).filter((k) => /^#(REF!|NAME\?)/.test(String(so.errors[k])) && !(ss.errors ?? {})[k]);
      if (eb0.length) newErrors.push(`${e.name}: ${eb0.slice(0, 4).map((k) => `${k} ${so.errors[k]}`).join(', ')}`);
      if (ss.visible !== so.visible) problems.push(`[${e.name}] 表示状態 ${ss.visible} → ${so.visible}`);
      continue;
    }
    for (const k of ['shapes', 'comments']) if ((ss[k] ?? 0) !== (so[k] ?? 0)) leftOut.push(`${e.name}: ${k === 'shapes' ? '図形・画像' : 'メモ'} ${ss[k] ?? 0} → ${so[k] ?? 0}`);
    if ((ss.hyperlinks ?? 0) !== (so.hyperlinks ?? 0)) problems.push(`[${e.name}] ハイパーリンク ${ss.hyperlinks ?? 0} → ${so.hyperlinks ?? 0}`);
    // visibility, hidden columns (widths follow the default font of each workbook: only hidden / not hidden is compared)
    if (ss.visible !== so.visible) problems.push(`[${e.name}] 表示状態 ${ss.visible} → ${so.visible}`);
    const hid = (s) => arr(s.cols).filter((c) => c.hidden).map((c) => c.letter).join(',');
    if (hid(ss) !== hid(so)) problems.push(`[${e.name}] 非表示の列 ${hid(ss) || 'なし'} → ${hid(so) || 'なし'}`);
    const eb = Object.keys(so.errors ?? {}).filter((k) => /^#(REF!|NAME\?)/.test(String(so.errors[k])) && !(ss.errors ?? {})[k]);
    if (eb.length) newErrors.push(`${e.name}: ${eb.slice(0, 4).map((k) => `${k} ${so.errors[k]}`).join(', ')}`);
  }
  if (newErrors.length) add('fail', 'copy-new-error', `コピーしたシートに、元のシートにはない #REF! / #NAME? があります`, newErrors);
  if (problems.length) add('fail', 'copy-differs', `元のシートと違うコピーがあります（${problems.length} 件）`, problems);
  else if (compared) add('pass', 'copy-equal', `コピーしたシート ${compared} 枚は、元のシートと同じ内容です（値・数式・結合・条件付き書式・入力規則・表・グラフ・ピボット・メモ・リンク）`);
  if (leftOut.length) add('warn', 'copy-leftout', `元のシートにある図形・メモ・コメントのうち、コピーされていないものがあります（スライサーなど、README の「コピーされないもの」）`, leftOut);
  if (textFiles.size) add('info', 'copy-text-source', `${[...textFiles].join('、')}: CSV などの文字ファイルは、Excel で開くと文字の扱い（先頭の 0・日付）が本ツールと違うため、中身の比較はしていません（シート名と表示状態だけ確認）`);
  if (exp.some((e) => !!out.date1904 !== !!e.dump.date1904) && compared) add('info', 'copy-date-system', '元のファイルと日付のしくみ（1900 / 1904）が違うため、日付は 1462 日ずらして保存されています（見た目の日付は同じ）');
  return { compared, expected: expNames };
}
