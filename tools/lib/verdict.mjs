// Verdict for an exported workbook vs the base input it was made from (owner: T). Pure: works on two
// inspect.ps1 dumps (tools/lib/inspect.mjs). Used by tools/check-output.mjs and the E2E specs.
//
// FAIL  = the output is broken for the user: needs repair / new #REF! or #NAME? / lost sheets, names, charts,
//         pivots, tables, CF / formulas now pointing to a different column (by header text).
// WARN  = suspicious but possibly intended (values changed, multi-column ranges now cover other headers …).
// INFO  = facts (formulas converted to values, rows appended …).
import { checkCopies } from './copies.mjs';
import {
  arr,
  cellAddr,
  extractRefs,
  headerIndex,
  headerRowGuess,
  headerText,
  parseCell,
  parseRanges,
  sheetByName,
  totalErrors,
  MAX_ROW,
} from './compare-core.mjs';

const FATAL_ERRORS = ['#REF!', '#NAME?'];

/** names used inside formulas (defined names of the book) that point at single columns -> for header mapping */
function nameTargets(dump) {
  const map = new Map();
  for (const n of arr(dump.names)) {
    const refs = extractRefs(String(n.refersTo ?? ''), null).filter((r) => r.sheet);
    if (refs.length === 1) map.set(n.name, refs[0]);
  }
  return map;
}

function refHeaders(dump, ref, headerRows) {
  const s = sheetByName(dump, ref.sheet);
  const hr = headerRows.get(ref.sheet);
  if (!s || !hr) return null;
  const out = [];
  const c2 = Math.min(ref.box.c2, s.maxCol ?? ref.box.c2, ref.box.c1 + 60);
  for (let c = ref.box.c1; c <= c2; c++) out.push(headerText(s, hr, c) ?? `(${c}列目:空)`);
  return out;
}

/** describe the columns a formula references: [{sheet, text, headers, single}] */
function formulaColumns(dump, formula, sheet, headerRows, names) {
  const res = [];
  for (const ref of extractRefs(formula, sheet)) {
    if (!headerRows.has(ref.sheet)) continue;
    // header cells themselves / rows above the header are not "column" references
    const hr = headerRows.get(ref.sheet);
    if (ref.box.r2 < hr) continue;
    const headers = refHeaders(dump, ref, headerRows);
    if (!headers) continue;
    res.push({ sheet: ref.sheet, text: ref.text, headers, single: ref.box.c1 === ref.box.c2 });
  }
  for (const [name, ref] of names) {
    const re = new RegExp(`(^|[^\\p{L}\\p{N}_.])${name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}(?![\\p{L}\\p{N}_.(!])`, 'u');
    if (!re.test(formula)) continue;
    if (!headerRows.has(ref.sheet)) continue;
    const headers = refHeaders(dump, ref, headerRows);
    if (headers) res.push({ sheet: ref.sheet, text: name, headers, single: ref.box.c1 === ref.box.c2, viaName: true });
  }
  return res;
}

function sig(cols) {
  return cols.map((c) => `${c.sheet}!${c.headers.join('+')}`).join(' , ');
}

/**
 * @param {any} base   inspect dump of the base input workbook
 * @param {any} out    inspect dump of the exported workbook
 * @param {{sheet?: string, headerRow?: number, maxExamples?: number}} [opts]
 *        sheet = the merged sheet (default: auto = the base sheet whose columns changed the most)
 * @returns {{ok: boolean, verdict: 'PASS'|'FAIL', items: Array<{level: 'fail'|'warn'|'info'|'pass', code: string, message: string, examples?: string[]}>, stats: any}}
 */
export function checkOutput(base, out, opts = {}) {
  const items = [];
  const maxEx = opts.maxExamples ?? 8;
  const add = (level, code, message, examples) => items.push({ level, code, message, ...(examples?.length ? { examples: examples.slice(0, maxEx) } : {}) });
  const stats = {};

  // 0. an Excel that answered while busy must never decide the verdict: a dump that still disagrees with its own package
  // after the fresh-instance retries of inspectFiles (tools/lib/dump-check.mjs) is reported as such instead of as a file problem
  for (const [label, d] of [['基準ファイル', base], ['出力ファイル', out]]) {
    if (d?.suspicious?.length) {
      add('fail', 'inspect-mismatch', `${label}: Excel から読み取った内容がファイルの中身と食い違っています（${d.attempts ?? 1} 回試行）。ファイルの問題ではなく検査の不具合かもしれません`, d.suspicious);
    }
  }
  stats.inspection = { base: base?.attempts ?? null, out: out?.attempts ?? null };

  // 1. opens without repair
  if (!out || out.openMode === 'failed' || !out.sheets) {
    add('fail', 'open-failed', `出力ファイルを Excel で開けません (${out?.openError ?? out?.error ?? 'unknown'})`);
    return finish();
  }
  if (out.openMode !== 'normal') {
    add('fail', 'repair', `Excel で開くと修復が必要です (${out.openMode})`, arr(out.repairLog));
  } else add('pass', 'open', 'Excel で修復なしに開けます');
  if (!base || !base.sheets) {
    add('warn', 'no-base', '比較用の基準ファイルを読めませんでした');
    return finish();
  }

  // 2. sheets
  const baseSheets = arr(base.sheets).map((s) => s.name);
  const outSheets = arr(out.sheets).map((s) => s.name);
  const lost = baseSheets.filter((n) => !outSheets.includes(n));
  const lostSet = new Set(lost.map((n) => n.toLowerCase()));
  const singleSheet = opts.singleSheet === true;
  const withCopies = opts.withCopies === true;
  if (singleSheet) {
    // 出力は結合したシートだけ（他のファイルのシートを含める場合は、それも）: 基準ファイルのほかのシートがなくなるのは仕様
    if (lost.length) add('info', 'sheets-dropped', `基準ファイルのほかのシート ${lost.length} 枚は出力に含めない設定のため、ありません: ${lost.join(', ')}`);
    const keptOthers = outSheets.filter((n) => baseSheets.includes(n));
    const extra = outSheets.filter((n) => !baseSheets.includes(n));
    if (keptOthers.length > 1) add('fail', 'not-single-sheet', `出力は結合したシートだけのはずですが、基準ファイルのシートが ${keptOthers.length} 枚残っています: ${keptOthers.join(', ')}`);
    else if (!withCopies && extra.length) add('fail', 'not-single-sheet', `出力は結合したシートだけのはずですが、ほかのシートがあります: ${extra.join(', ')}（他のファイルのシートを含める出力なら --with-copies を付けてください）`);
    else if (!keptOthers.length) add('fail', 'sheets-lost', '結合したシート（基準ファイルのシート）が出力にありません');
    else if (!lost.length || keptOthers.length === 1) add('pass', 'sheets', withCopies ? '基準ファイルのシートは結合したシートだけです（他のファイルのシートは別に確認します）' : '出力は結合したシート 1 枚だけです');
  } else if (lost.length) {
    add('fail', 'sheets-lost', `シートが消えています: ${lost.join(', ')}${outSheets.length === 1 ? '（出力が 1 シートだけなら、結合したシートだけの出力です。--single-sheet を付けて確認してください）' : ''}`);
  } else add('pass', 'sheets', `シート ${baseSheets.length} 枚はすべて残っています`);
  const common = baseSheets.filter((n) => outSheets.includes(n));
  // the sheets copied from the other files are judged apart (against their sources, copies.mjs): the checks below look at the base file's own sheets
  const baseNameSet = new Set(baseSheets);
  const outView = withCopies ? { ...out, sheets: arr(out.sheets).filter((s) => baseNameSet.has(s.name)) } : out;
  if (JSON.stringify(common) !== JSON.stringify(outSheets.filter((n) => baseSheets.includes(n)))) add('warn', 'sheet-order', 'シートの順番が変わっています');

  // merged sheet + header rows (from the BASE: header rows do not move)
  const headerRows = new Map();
  for (const s of arr(base.sheets)) {
    const hr = headerRowGuess(s);
    if (hr) headerRows.set(s.name, hr);
  }
  let merged = opts.sheet ?? null;
  if (!merged) {
    let best = -1;
    for (const n of common) {
      const bs = sheetByName(base, n);
      const os = sheetByName(out, n);
      const d = Math.abs((os.maxCol ?? 0) - (bs.maxCol ?? 0)) * 1000 + Math.abs((os.maxRow ?? 0) - (bs.maxRow ?? 0));
      if (d > best) {
        best = d;
        merged = n;
      }
    }
  }
  if (merged && opts.headerRow) headerRows.set(merged, opts.headerRow);
  stats.mergedSheet = merged;
  stats.headerRow = merged ? headerRows.get(merged) ?? null : null;
  // data block of the merged sheet (header row + contiguous non-empty rows) in the base and in the output: rows
  // appended by join 'full' sit between the base's last data row and the rows below the table (footer)
  const blockB = merged ? dataBlock(sheetByName(base, merged), headerRows.get(merged)) : null;
  const blockO = merged ? dataBlock(sheetByName(out, merged), headerRows.get(merged)) : null;
  const rowShift = blockB && blockO ? blockO.last - blockB.last : 0;
  if (blockB) stats.dataRows = { base: blockB, out: blockO };

  // 3. #REF! in formula text / names / rules / charts
  const refText = (d) => {
    const cells = [];
    for (const s of arr(d.sheets)) for (const [a, f] of Object.entries(s.formulas ?? {})) if (String(f).includes('#REF!')) cells.push(`${s.name}!${a}: ${f}`);
    for (const n of arr(d.names)) if (String(n.refersTo ?? '').includes('#REF!')) cells.push(`名前 ${n.name}: ${n.refersTo}`);
    for (const s of arr(d.sheets)) {
      for (const c of arr(s.cf)) if (`${c.formula1 ?? ''}${c.formula2 ?? ''}`.includes('#REF!')) cells.push(`${s.name} 条件付き書式 ${c.appliesTo}: ${c.formula1}`);
      for (const v of arr(s.dv)) if (`${v.formula1 ?? ''}${v.formula2 ?? ''}`.includes('#REF!')) cells.push(`${s.name} 入力規則 ${v.range}: ${v.formula1}`);
      for (const ch of arr(s.charts)) for (const se of arr(ch.series)) if (String(se.formula ?? '').includes('#REF!')) cells.push(`${s.name} グラフ ${ch.name}: ${se.formula}`);
    }
    return cells;
  };
  const refBase = refText(base);
  const refOut = refText(outView);
  stats.refText = { base: refBase.length, out: refOut.length };
  if (refOut.length > refBase.length) add('fail', 'new-ref-text', `#REF! を含む数式・名前・書式が ${refOut.length - refBase.length} 件増えています`, refOut.filter((x) => !refBase.includes(x)));
  else add('pass', 'ref-text', '新しい #REF! 参照はありません');

  // 4. error values
  const eb = totalErrors(base);
  const eo = totalErrors(outView);
  stats.errors = { base: eb, out: eo };
  const errorCells = (d, kind) => {
    const r = [];
    for (const s of arr(d.sheets)) for (const [a, e] of Object.entries(s.errors ?? {})) if (e === kind) r.push(`${s.name}!${a}${s.formulas?.[a] ? ' ' + s.formulas[a] : ''}`);
    return r;
  };
  let anyErr = false;
  for (const kind of new Set([...Object.keys(eb), ...Object.keys(eo)])) {
    const d = (eo[kind] ?? 0) - (eb[kind] ?? 0);
    if (d <= 0) continue;
    anyErr = true;
    add(FATAL_ERRORS.includes(kind) ? 'fail' : 'warn', `new-error-${kind}`, `エラー値 ${kind} が ${d} 件増えています`, errorCells(outView, kind));
  }
  if (!anyErr) add('pass', 'errors', '新しいエラー値 (#REF! / #NAME? など) はありません');

  // 5. defined names
  const nkey = (n) => `${n.scope ?? ''}!${n.name}`;
  const outNames = new Map(arr(out.names).map((n) => [nkey(n), n]));
  const lostNames = arr(base.names).filter((n) => !/^_xl/i.test(n.name) && !outNames.has(nkey(n)));
  // a name that belongs to / points at a sheet that is not in the output (single-sheet output) goes with the sheet
  const readsRemoved = (text, scope = null) => {
    if (!singleSheet || !lostSet.size) return false;
    if (scope && lostSet.has(String(scope).toLowerCase())) return true;
    return extractRefs(String(text ?? ''), null).some((r) => r.qualified && lostSet.has(String(r.sheet).toLowerCase()));
  };
  // a defined name used inside a rule formula (=大学一覧) that points at a removed sheet
  const removedNames = new Set(arr(base.names).filter((n) => readsRemoved(n.refersTo, n.scope)).map((n) => String(n.name).toLowerCase()));
  const readsRemovedName = (text) => {
    if (!singleSheet || !removedNames.size) return false;
    for (const nm of removedNames) {
      const re = new RegExp(`(^|[^\\p{L}\\p{N}_.])${nm.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}(?![\\p{L}\\p{N}_.(!])`, 'iu');
      if (re.test(String(text ?? ''))) return true;
    }
    return false;
  };
  const droppedNames = lostNames.filter((n) => readsRemoved(n.refersTo, n.scope));
  const reallyLost = lostNames.filter((n) => !droppedNames.includes(n));
  if (droppedNames.length) add('info', 'names-dropped', `出力に含めないシートの名前の定義 ${droppedNames.length} 件は、シートと一緒になくなりました: ${droppedNames.map((n) => n.name).join(', ')}`);
  if (reallyLost.length) add('fail', 'names-lost', `名前の定義が消えています: ${reallyLost.map((n) => n.name).join(', ')}`);

  // 6. formulas still point to the same columns (by header text)
  const namesBase = nameTargets(base);
  const namesOut = nameTargets(out);
  const moved = [];
  const rangeChanged = [];
  const converted = [];
  const vanished = [];
  let checked = 0;
  for (const n of common) {
    const bs = sheetByName(base, n);
    const os = sheetByName(out, n);
    for (const [addr, fb] of Object.entries(bs.formulas ?? {})) {
      const colsB = formulaColumns(base, fb, n, headerRows, namesBase);
      if (!colsB.length) continue;
      // where is this cell in the output?
      let target = addr;
      if (n === merged) {
        target = locateMergedCell(base, out, n, addr, headerRows.get(n), blockB, rowShift);
        if (!target) continue;
      }
      const fo = os.formulas?.[target];
      if (fo == null) {
        const v = os.values?.[target] ?? os.errors?.[target];
        if (v != null) converted.push(`${n}!${addr}: ${fb} -> 値 ${JSON.stringify(v)}`);
        else vanished.push(`${n}!${addr}: ${fb}`);
        continue;
      }
      checked++;
      const colsO = formulaColumns(out, fo, n, headerRows, namesOut);
      const singleB = colsB.filter((c) => c.single);
      const singleO = colsO.filter((c) => c.single);
      if (sig(singleB) !== sig(singleO)) moved.push(`${n}!${addr}${target !== addr ? `→${target}` : ''}: ${fb} → ${fo}  [${sig(singleB)}] → [${sig(singleO)}]`);
      else {
        const multiB = colsB.filter((c) => !c.single);
        const multiO = colsO.filter((c) => !c.single);
        if (sig(multiB) !== sig(multiO)) rangeChanged.push(`${n}!${addr}: ${fb} → ${fo}  [${sig(multiB)}] → [${sig(multiO)}]`);
      }
    }
  }
  stats.formulasChecked = checked;
  if (moved.length) add('fail', 'formula-columns', `参照先の列（見出し）が変わった数式が ${moved.length} 件あります`, moved);
  else add('pass', 'formula-columns', `数式 ${checked} 件はすべて同じ見出しの列を参照しています`);
  if (rangeChanged.length) add('warn', 'range-columns', `複数列の範囲に含まれる列が変わった数式が ${rangeChanged.length} 件あります（列の移動によっては正常）`, rangeChanged);
  if (converted.length) add('info', 'formulas-to-values', `値に変換された数式が ${converted.length} 件あります`, converted);
  if (vanished.length) add('fail', 'formulas-lost', `消えた数式が ${vanished.length} 件あります`, vanished);

  // 6b. rows appended at the end of the table (join 'full'): do the bounded ranges over the table grow with it?
  if (merged && blockB && blockO && rowShift > 0) {
    const cov = appendedCoverage(base, out, merged, common, headerRows, blockB, blockO, rowShift);
    stats.appended = { rows: rowShift, firstRow: blockB.last + 1, lastRow: blockO.last, ranges: cov.total, extended: cov.extended.length, notExtended: cov.excluded.length };
    add('info', 'appended-rows', `[${merged}] 表の末尾に ${rowShift} 行追加されています（${blockB.last + 1}〜${blockO.last} 行目）`);
    if (cov.total) {
      add(
        'info',
        'appended-ranges',
        `表の範囲（${blockB.first}〜${blockB.last} 行目）を指す数式・名前・ピボット・書式 ${cov.total} 件のうち ${cov.extended.length} 件は追加した行まで広がっています`,
        cov.extended,
      );
    }
    if (cov.excluded.length)
      add('warn', 'appended-ranges-excluded', `${cov.excluded.length} 件の範囲は追加した行を含みません（集計・書式に入りません）`, cov.excluded);
  }

  // headers of the merged sheet that are gone from the output (columns the user deleted): a pivot source / CF / DV
  // that loses exactly those columns follows the deletion like Excel does — reported as info, not as a failure
  const goneHeaders = new Set();
  if (merged && headerRows.get(merged)) {
    const hb = headerIndex(sheetByName(base, merged), headerRows.get(merged));
    const ho = headerIndex(sheetByName(out, merged), headerRows.get(merged));
    for (const h of hb.keys()) if (!ho.has(h)) goneHeaders.add(h);
  }
  const allGone = (hs) => hs.length > 0 && hs.every((h) => goneHeaders.has(h));

  // 7. objects: charts / pivots / tables
  for (const n of common) {
    const bs = sheetByName(base, n);
    const os = sheetByName(out, n);
    const bc = arr(bs.charts);
    const oc = arr(os.charts);
    // (single-sheet output: a chart that reads a sheet that is not in the output cannot stay)
    const droppedCharts = bc.filter((ch) => arr(ch.series).some((se) => readsRemoved(se.formula))).length;
    if (droppedCharts) add('info', 'charts-dropped', `[${n}] 出力に含めないシートを参照するグラフ ${droppedCharts} 個は、シートと一緒になくなりました`);
    if (oc.length < bc.length - droppedCharts) add('fail', 'charts-lost', `[${n}] グラフが ${bc.length} → ${oc.length} 個に減っています`);
    for (const ch of bc) {
      if (arr(ch.series).some((se) => readsRemoved(se.formula))) continue;
      const o = oc.find((x) => x.name === ch.name) ?? oc[bc.indexOf(ch)];
      if (!o) continue;
      const sb = arr(ch.series).map((s) => sig(formulaColumns(base, String(s.formula ?? ''), n, headerRows, namesBase)));
      const so = arr(o.series).map((s) => sig(formulaColumns(out, String(s.formula ?? ''), n, headerRows, namesOut)));
      if (JSON.stringify(sb) !== JSON.stringify(so)) add('fail', 'chart-series', `[${n}] グラフ「${ch.name}」の系列の参照先が変わっています`, [`${JSON.stringify(sb)} → ${JSON.stringify(so)}`]);
    }
    const bp = arr(bs.pivots);
    const op = arr(os.pivots);
    if (op.length < bp.length) add('fail', 'pivots-lost', `[${n}] ピボットテーブルが ${bp.length} → ${op.length} 個に減っています`);
    for (const p of bp) {
      const o = op.find((x) => x.name === p.name);
      if (!o) continue;
      const hb = pivotHeaders(base, p.source, headerRows);
      const ho = pivotHeaders(out, o.source, headerRows);
      if (hb && ho) {
        const missing = hb.filter((h) => !ho.includes(h));
        if (missing.length && allGone(missing))
          add('info', 'pivot-source-deleted', `[${n}] ピボット「${p.name}」の元データから、削除した列が外れています: ${missing.join(', ')}`, [`${p.source} → ${o.source}`]);
        else if (missing.length) add('fail', 'pivot-source', `[${n}] ピボット「${p.name}」の元データから列が外れています: ${missing.join(', ')}`, [`${p.source} → ${o.source}`]);
      }
    }
    const bt = arr(bs.tables);
    for (const t of bt) if (!arr(os.tables).some((x) => x.name === t.name)) add('fail', 'table-lost', `[${n}] テーブル「${t.name}」が消えています`);
  }

  // 8. merged sheet features: CF / DV / freeze / filter / hidden columns (compared by header text)
  if (merged) {
    const bs = sheetByName(base, merged);
    const os = sheetByName(out, merged);
    const hr = headerRows.get(merged);
    if (bs && os && hr) {
      // A rule (CF / DV) is described by its kind, the header texts of the columns it covers and the columns its
      // formula refers to (a rule is "kept" when it still covers every column that had the original header).
      const hdrList = (s, ranges) => {
        const hs = new Set();
        for (const b of parseRanges(ranges)) for (let c = b.c1; c <= Math.min(b.c2, b.c1 + 60); c++) hs.add(headerText(s, hr, c) ?? `(${c})`);
        return [...hs].sort();
      };
      const cfRule = (dump, s, c, names) => ({
        kind: String(c.type),
        hs: hdrList(s, c.appliesTo),
        rest: sig(formulaColumns(dump, String(c.formula1 ?? ''), merged, headerRows, names)),
        range: String(c.appliesTo ?? ''),
        raw: c,
      });
      const dvRule = (dump, s, d, names) => ({
        kind: String(d.type),
        hs: hdrList(s, d.range),
        rest: `${sig(formulaColumns(dump, String(d.formula1 ?? ''), merged, headerRows, names))}|${/!/.test(String(d.formula1 ?? '')) ? '' : d.formula1 ?? ''}`,
        range: String(d.range ?? ''),
        raw: d,
      });
      // a rule that only sat on deleted columns is gone with them (Excel deletes it too)
      const ruleGone = (hs) => allGone(hs);
      const showHeader = (h) => (/^\(\d+\)$/.test(h) ? '見出しのない列' : h);
      /**
       * How does output rule `o` relate to base rule `b`? null = not the same rule / an original column lost it.
       * 'same' = the same columns; 'grown' = also covers other columns (Excel extends a rule over a column inserted
       * next to a column carrying it: verified behaviour of Excel and of both engines); 'shrunk' = only columns the
       * user deleted dropped out.
       */
      const relation = (b, o) => {
        if (b.kind !== o.kind || b.rest !== o.rest) return null;
        const need = b.hs.filter((h) => !goneHeaders.has(h));
        if (!need.length || !need.every((h) => o.hs.includes(h))) return null;
        const extras = o.hs.filter((h) => !b.hs.includes(h));
        if (extras.length) return { how: 'grown', extras };
        return { how: need.length < b.hs.length ? 'shrunk' : 'same', extras };
      };
      /** classify every base rule; returns the lost ones and the grown / shrunk ones (with the output range) */
      const classify = (baseRules, outRules) => {
        const lost = [];
        const grown = [];
        let shrunk = 0;
        let gone = 0;
        let kept = 0;
        for (const b of baseRules) {
          const rels = outRules.map((o) => ({ o, r: relation(b, o) })).filter((x) => x.r);
          const best = rels.find((x) => x.r.how === 'same') ?? rels.find((x) => x.r.how === 'shrunk') ?? rels[0];
          if (!best) {
            if (ruleGone(b.hs)) gone++;
            else lost.push(b);
            continue;
          }
          kept++;
          if (best.r.how === 'shrunk') shrunk++;
          if (best.r.how === 'grown') grown.push(`${b.range} → ${best.o.range}（広がった先: ${best.r.extras.map(showHeader).join(', ')}）`);
        }
        return { lost, grown, shrunk, gone, kept };
      };
      const cfBase = arr(bs.cf).map((c) => cfRule(base, bs, c, namesBase));
      const cfC = classify(cfBase, arr(os.cf).map((c) => cfRule(out, os, c, namesOut)));
      if (cfC.gone) add('info', 'cf-deleted', `[${merged}] 削除した列だけに付いていた条件付き書式 ${cfC.gone} 件は列と一緒になくなりました`);
      // (single-sheet output: a rule whose formula reads a sheet that is not in the output is dropped)
      const cfDropped = cfC.lost.filter((c) => readsRemoved(c.raw.formula1) || readsRemoved(c.raw.formula2) || readsRemovedName(c.raw.formula1));
      cfC.lost = cfC.lost.filter((c) => !cfDropped.includes(c));
      if (cfDropped.length) add('info', 'cf-dropped', `[${merged}] 出力に含めないシートを参照する条件付き書式 ${cfDropped.length} 件は、はずしました`, cfDropped.map((c) => `${c.range} ${c.raw.formula1 ?? ''}`));
      if (cfC.lost.length) add('fail', 'cf', `[${merged}] 条件付き書式 ${cfC.lost.length} 件が元の列に付いていません`, cfC.lost.map((c) => `${c.range} ${c.raw.formula1 ?? ''}`));
      else if (cfBase.length) add('pass', 'cf', `[${merged}] 条件付き書式 ${cfC.kept} 件は同じ見出しの列に付いています`);
      if (cfC.grown.length)
        add('info', 'cf-grown', `[${merged}] 条件付き書式 ${cfC.grown.length} 件は、隣に挿入した（移動してきた）列まで範囲が広がっています（Excelの列の挿入と同じ動作）`, cfC.grown);
      if (cfC.shrunk) add('info', 'cf-shrunk', `[${merged}] 条件付き書式 ${cfC.shrunk} 件は、削除した列の分だけ範囲が狭くなっています`);
      const dvBase = arr(bs.dv).map((d) => dvRule(base, bs, d, namesBase));
      const dvC = classify(dvBase, arr(os.dv).map((d) => dvRule(out, os, d, namesOut)));
      if (dvC.gone) add('info', 'dv-deleted', `[${merged}] 削除した列だけに付いていた入力規則 ${dvC.gone} 件は列と一緒になくなりました`);
      // (single-sheet output: a list that reads a sheet that is not in the output becomes a literal list when it fits, else the validation is dropped)
      const dvDropped = dvC.lost.filter((d) => readsRemoved(d.raw.formula1) || readsRemoved(d.raw.formula2) || readsRemovedName(d.raw.formula1));
      dvC.lost = dvC.lost.filter((d) => !dvDropped.includes(d));
      if (dvDropped.length) add('info', 'dv-literalized', `[${merged}] 出力に含めないシートを参照する入力規則 ${dvDropped.length} 件は、リストをそのまま書き込んだか、はずしました`, dvDropped.map((d) => `${d.range} ${d.raw.formula1 ?? ''}`));
      if (dvC.lost.length) add('fail', 'dv', `[${merged}] 入力規則 ${dvC.lost.length} 件が元の列に付いていません`, dvC.lost.map((d) => `${d.range} ${d.raw.formula1 ?? ''}`));
      if (dvC.grown.length)
        add('info', 'dv-grown', `[${merged}] 入力規則 ${dvC.grown.length} 件は、隣に挿入した（移動してきた）列まで範囲が広がっています（Excelの列の挿入と同じ動作）`, dvC.grown);
      if (dvC.shrunk) add('info', 'dv-shrunk', `[${merged}] 入力規則 ${dvC.shrunk} 件は、削除した列の分だけ範囲が狭くなっています`);
      stats.ruleGrowth = { cf: cfC.grown.length, dv: dvC.grown.length };
      if (bs.freeze?.frozen && !os.freeze?.frozen) add('warn', 'freeze', `[${merged}] ウィンドウ枠の固定が外れています`);
      if (bs.freeze?.frozen && os.freeze?.frozen && bs.freeze.splitRow !== os.freeze.splitRow) add('warn', 'freeze', `[${merged}] 固定している行数が変わっています (${bs.freeze.splitRow} → ${os.freeze.splitRow})`);
      if (bs.autoFilter && !os.autoFilter) add('warn', 'autofilter', `[${merged}] オートフィルターが外れています`);
      const hiddenHeaders = (s) => arr(s.hiddenColumns).map((l) => headerText(s, hr, parseCell(`${l}1`).col)).filter(Boolean).sort();
      const hb = hiddenHeaders(bs);
      const ho = hiddenHeaders(os);
      const unhidden = hb.filter((h) => !ho.includes(h));
      if (unhidden.length) add('warn', 'hidden-columns', `[${merged}] 非表示だった列が表示されています: ${unhidden.join(', ')}`);
      stats.rows = { base: bs.maxRow, out: os.maxRow };
      stats.cols = { base: bs.maxCol, out: os.maxCol };
      const hdrB = [...headerIndex(bs, hr).keys()];
      const hdrO = [...headerIndex(os, hr).keys()];
      stats.addedHeaders = hdrO.filter((h) => !hdrB.includes(h));
      stats.removedHeaders = hdrB.filter((h) => !hdrO.includes(h));
      if (stats.addedHeaders.length) add('info', 'added-columns', `追加された列: ${stats.addedHeaders.join(', ')}`);
      if (stats.removedHeaders.length) add('info', 'removed-columns', `なくなった列: ${stats.removedHeaders.join(', ')}`);
    }
  }

  // 9. values of formulas on the other sheets (informational: appended rows legitimately change counts)
  const changed = [];
  for (const n of common) {
    if (n === merged) continue;
    const bs = sheetByName(base, n);
    const os = sheetByName(out, n);
    for (const a of Object.keys(bs.formulas ?? {})) {
      const vb = bs.values?.[a] ?? bs.errors?.[a];
      const vo = os.values?.[a] ?? os.errors?.[a];
      const same = typeof vb === 'number' && typeof vo === 'number' ? Math.abs(vb - vo) <= 1e-9 * Math.max(1, Math.abs(vb)) : vb === vo;
      if (!same) changed.push(`${n}!${a}: ${JSON.stringify(vb)} → ${JSON.stringify(vo)}`);
    }
  }
  if (changed.length) add('warn', 'values-changed', `他のシートの数式の結果が ${changed.length} 件変わっています（行を追加した場合などは正常）`, changed);

  // 10. single-sheet output: a formula that reads a sheet which is not in the output would be #REF! — such formulas must have become values
  if (singleSheet) {
    const have = new Set(outSheets.map((n) => n.toLowerCase()));
    const bad = [];
    for (const s of arr(out.sheets)) {
      for (const [a, f] of Object.entries(s.formulas ?? {})) {
        if (extractRefs(String(f), s.name).some((r) => r.qualified && !have.has(String(r.sheet).toLowerCase()))) bad.push(`${s.name}!${a}: ${f}`);
      }
    }
    if (bad.length) add('fail', 'dangling-sheet-ref', `出力にないシートを参照する数式が ${bad.length} 件あります（値になっているはずです）`, bad);
    else add('pass', 'no-dangling-ref', '出力にないシートを参照する数式はありません（参照していた数式は値になっています）');
  }

  // 11. the sheets copied from the other files (plan A)
  if (withCopies) {
    const { compared } = checkCopies({ base, out, merged: merged ?? outSheets[0], sources: arr(opts.sources), outStrict: opts.outStrict === true, add });
    stats.copies = { compared };
  }

  return finish();

  function finish() {
    const ok = !items.some((i) => i.level === 'fail');
    return { ok, verdict: ok ? 'PASS' : 'FAIL', items, stats };
  }
}

/** header texts covered by a pivot source ('名簿!R3C1:R63C12' or '名簿!$A$3:$L$63') */
function pivotHeaders(dump, source, headerRows) {
  if (!source) return null;
  const s = String(source);
  const bang = s.lastIndexOf('!');
  if (bang < 0) return null;
  const sheet = s.slice(0, bang).replace(/^'|'$/g, '').replace(/''/g, "'");
  const addr = s.slice(bang + 1);
  let box = null;
  const rc = /^R(\d+)C(\d+)(?::R(\d+)C(\d+))?$/i.exec(addr);
  if (rc) box = { r1: +rc[1], c1: +rc[2], r2: +(rc[3] ?? rc[1]), c2: +(rc[4] ?? rc[2]) };
  else box = parseRanges(addr)[0] ?? null;
  if (!box) return null;
  const sh = sheetByName(dump, sheet);
  if (!sh) return null;
  const hr = box.r1 <= MAX_ROW ? box.r1 : headerRows.get(sheet);
  const hs = [];
  for (let c = box.c1; c <= box.c2; c++) hs.push(headerText(sh, hr, c) ?? `(${c})`);
  return hs;
}

/**
 * The data block below the header row: rows headerRow+1 .. last where every row has some value / formula.
 * Returns {first, last} or null when the row under the header is empty.
 */
export function dataBlock(sheet, headerRow) {
  if (!sheet || !headerRow) return null;
  const rows = new Set();
  for (const m of [sheet.values, sheet.formulas, sheet.errors]) for (const a of Object.keys(m ?? {})) rows.add(Number(/\d+$/.exec(a)?.[0]));
  let last = headerRow;
  while (rows.has(last + 1)) last++;
  return last > headerRow ? { first: headerRow + 1, last } : null;
}

/**
 * For appended rows: which references to the data block (rows first..last of the base, bounded — whole columns
 * always include new rows) grew over the appended rows in the output. Formulas (all sheets), names, pivot sources,
 * CF / DV areas and the autofilter of the merged sheet.
 */
function appendedCoverage(base, out, merged, common, headerRows, blockB, blockO, rowShift) {
  const extended = [];
  const excluded = [];
  const covers = (box) => box.r1 <= blockB.first && box.r2 === blockB.last;
  const grew = (box) => box.r2 >= blockO.last;
  const mergedRefs = (f, sheet) => extractRefs(String(f ?? ''), sheet).filter((r) => r.sheet === merged);
  const pairUp = (label, fb, fo, sheet) => {
    const rb = mergedRefs(fb, sheet);
    if (!rb.some((r) => covers(r.box))) return;
    const ro = fo == null ? null : mergedRefs(fo, sheet);
    rb.forEach((r, i) => {
      if (!covers(r.box)) return;
      const o = ro && ro.length === rb.length ? ro[i] : null;
      const line = `${label}: ${r.text} → ${o ? o.text : fo == null ? '(なし)' : '?'}`;
      (o && grew(o.box) ? extended : excluded).push(line);
    });
  };
  for (const n of common) {
    const bs = sheetByName(base, n);
    const os = sheetByName(out, n);
    for (const [addr, fb] of Object.entries(bs.formulas ?? {})) {
      const target = n === merged ? locateMergedCell(base, out, n, addr, headerRows.get(n), blockB, rowShift) : addr;
      if (!target) continue;
      const fo = os.formulas?.[target];
      if (fo == null) continue; // converted to a value / deleted: reported elsewhere
      pairUp(`${n}!${target}`, fb, fo, n);
    }
  }
  const outNames = new Map(arr(out.names).map((x) => [`${x.scope ?? ''}!${x.name}`, x]));
  for (const nb of arr(base.names)) {
    if (/^_xl/i.test(nb.name)) continue;
    const no = outNames.get(`${nb.scope ?? ''}!${nb.name}`);
    if (no) pairUp(`名前 ${nb.name}`, nb.refersTo, no.refersTo, null);
  }
  for (const n of common) {
    const bp = arr(sheetByName(base, n).pivots);
    const op = arr(sheetByName(out, n).pivots);
    for (const p of bp) {
      const b = pivotBox(p.source);
      if (!b || b.sheet !== merged || !covers(b.box)) continue;
      const o = op.find((x) => x.name === p.name);
      const ob = o ? pivotBox(o.source) : null;
      (ob && grew(ob.box) ? extended : excluded).push(`ピボット「${p.name}」の元データ: ${p.source} → ${o?.source ?? '(なし)'}`);
    }
  }
  const bs = sheetByName(base, merged);
  const os = sheetByName(out, merged);
  const hr = headerRows.get(merged);
  const areaRules = (s, kind) => (kind === 'cf' ? arr(s.cf).map((c) => ({ type: c.type, ranges: c.appliesTo })) : arr(s.dv).map((d) => ({ type: d.type, ranges: d.range })));
  for (const kind of ['cf', 'dv']) {
    const label = kind === 'cf' ? '条件付き書式' : '入力規則';
    for (const rule of areaRules(bs, kind)) {
      for (const box of parseRanges(rule.ranges)) {
        if (!covers(box)) continue;
        const h = headerText(bs, hr, box.c1);
        const hit = areaRules(os, kind).some((o) => o.type === rule.type && parseRanges(o.ranges).some((ob) => grew(ob) && ob.r1 <= blockO.first && headerText(os, hr, ob.c1) === h));
        (hit ? extended : excluded).push(`${label}（${h ?? '?'} 列）: ${rule.ranges}`);
      }
    }
  }
  if (bs.autoFilter) {
    const b = parseRanges(bs.autoFilter)[0];
    const o = os.autoFilter ? parseRanges(os.autoFilter)[0] : null;
    if (b && covers(b)) (o && grew(o) ? extended : excluded).push(`オートフィルター: ${bs.autoFilter} → ${os.autoFilter ?? '(なし)'}`);
  }
  return { total: extended.length + excluded.length, extended, excluded };
}

/** {sheet, box} of a pivot source ('名簿!R3C1:R63C12' or "'名簿'!$A$3:$L$63") */
function pivotBox(source) {
  if (!source) return null;
  const s = String(source);
  const bang = s.lastIndexOf('!');
  if (bang < 0) return null;
  const sheet = s.slice(0, bang).replace(/^'|'$/g, '').replace(/''/g, "'");
  const addr = s.slice(bang + 1);
  const rc = /^R(\d+)C(\d+)(?::R(\d+)C(\d+))?$/i.exec(addr);
  const box = rc ? { r1: +rc[1], c1: +rc[2], r2: +(rc[3] ?? rc[1]), c2: +(rc[4] ?? rc[2]) } : parseRanges(addr)[0] ?? null;
  return box ? { sheet, box } : null;
}

/**
 * Position of a merged-sheet cell of the base in the output: the column is found by its header text (unique),
 * the row is the same, or shifted by the number of appended rows for cells below the base data block.
 */
function locateMergedCell(base, out, sheet, addr, headerRow, blockB = null, rowShift = 0) {
  const p = parseCell(addr);
  if (!p || !headerRow) return null;
  const bs = sheetByName(base, sheet);
  const os = sheetByName(out, sheet);
  const h = headerText(bs, headerRow, p.col);
  let col = p.col;
  if (h) {
    const cols = headerIndex(os, headerRow).get(h) ?? [];
    const bcols = headerIndex(bs, headerRow).get(h) ?? [];
    if (cols.length !== 1 || bcols.length !== 1) return null;
    col = cols[0];
  } else return null;
  if (p.row <= headerRow) return cellAddr(p.row, col);
  if (blockB) return cellAddr(p.row > blockB.last ? p.row + rowShift : p.row, col);
  const shift = (os.maxRow ?? 0) - (bs.maxRow ?? 0);
  const a0 = cellAddr(p.row, col);
  if (shift === 0 || os.formulas?.[a0] != null) return a0;
  return cellAddr(p.row + shift, col);
}

/** multi-line human readable verdict */
export function formatVerdict(v, title = '') {
  const mark = { fail: 'NG ', warn: '注意', info: '情報', pass: 'OK ' };
  const lines = [];
  lines.push(`${title ? title + ': ' : ''}${v.verdict === 'PASS' ? '合格 (PASS)' : '不合格 (FAIL)'}`);
  for (const it of v.items) {
    lines.push(`  [${mark[it.level]}] ${it.message}`);
    for (const e of it.examples ?? []) lines.push(`         - ${e}`);
  }
  return lines.join('\n');
}
