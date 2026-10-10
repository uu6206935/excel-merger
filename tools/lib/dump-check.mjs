// Cross-check of an inspect.ps1 dump against the PACKAGE it came from (owner: T). Pure JS, no Excel.
//
// Why: the verdict of check-output trusts what our hidden Excel reports. Once in a while that report was wrong although the file
// was fine (a defined name missing, every formula "pointing at other headers", no chart series / CF / DV) and the same file passed
// in the next batch: Excel answered while it was still busy (loading, calculating, another instance competing) and a swallowed COM
// error turned into "nothing there". The package itself never changes, so what Excel says can be checked against what the file
// contains: defined names, formula cells and the columns their references point to, CF / DV / chart / pivot / table presence.
// `inspectFiles` (tools/lib/inspect.mjs) runs this for every .xlsx / .xlsm and re-inspects in a FRESH Excel instance when it finds
// a disagreement; a disagreement that survives the retries is reported by the verdict as `inspect-mismatch`.
import fs from 'node:fs';
import { extractRefs, arr, sheetByName, parseCell, parseRange } from './compare-core.mjs';
import { loadPackageModel } from './package-diff.mjs';

/** packages bigger than this are not cross-checked (parsing them costs more than an Excel retry) */
export const MAX_CHECK_BYTES = 40 * 1024 * 1024;

/** text constants kept per sheet for the comparison (the first ones in row order: headers and the top of the data) */
const MAX_TEXTS = 600;
const normText = (t) => String(t).replace(/\r\n?/g, '\n').trim();

/** is the cell `addr` inside the area of a pivot table of the sheet facts `f` */
function inPivot(f, addr) {
  const c = parseCell(addr);
  return !!c && arr(f.pivotAreas).some((b) => c.row >= b.r1 && c.row <= b.r2 && c.col >= b.c1 && c.col <= b.c2);
}

/** sorted reference signature of a formula: which ranges it points to ('名簿!1,4,1,63;集計!2,5,2,5'), `$` ignored */
export function refSignature(formula, sheet) {
  const parts = extractRefs(formula, sheet).map((r) => `${r.sheet ?? ''}!${r.box.c1},${r.box.r1},${r.box.c2},${r.box.r2}`);
  return parts.sort().join(';');
}

/**
 * What the package contains, in the terms of an inspect dump. null when the file is not an OOXML workbook (xls / csv / ods:
 * Excel converts those, there is nothing to compare with) or cannot be parsed.
 * @returns {Promise<null | {names: Array<{name: string, scope: string | null}>, sheets: Array<{name: string, formulas: Map<string, string>, texts: Map<string, string>, cf: number, dv: number, tables: number, charts: number, pivots: number, pivotAreas: Array<{r1: number, c1: number, r2: number, c2: number}>}>}>}
 */
export async function packageFacts(file) {
  if (!/\.(xlsx|xlsm)$/i.test(String(file))) return null;
  try {
    if (fs.statSync(file).size > MAX_CHECK_BYTES) return null;
    const model = await loadPackageModel(file);
    const names = model.names.filter((n) => !/^_xl/i.test(n.name)).map((n) => ({ name: n.name, scope: n.scope ?? null }));
    const sheets = [];
    for (const s of model.sheets) {
      const m = model.sheetModels.get(s.name);
      if (!m) continue;
      const formulas = new Map();
      const texts = new Map();
      for (const [addr, c] of m.cells) {
        // data tables (TABLE()) are reported by Excel in its own way; everything else is an ordinary formula cell
        if (typeof c.f === 'string' && c.f !== '' && c.fType !== 'dataTable') formulas.set(addr, refSignature(c.f, s.name));
        else if (c.f == null && c.type === 's' && typeof c.value === 'string' && c.value.trim() !== '' && texts.size < MAX_TEXTS) texts.set(addr, normText(c.value));
      }
      sheets.push({
        name: s.name,
        formulas,
        texts,
        cf: m.cf.length,
        dv: m.dv.length,
        tables: m.tables.length,
        // a drawing anchor holding a chart reads 'twoCell(..) title=..|barChart:col:clustered[series]' (bar3DChart, pieChart, ...)
        charts: m.drawings.filter((d) => /[A-Za-z0-9]Chart[:[]/.test(d)).length,
        pivots: m.pivots.length,
        // the cells a pivot table draws: Excel re-renders them when it refreshes the table (refresh on open, a source that changed), so their text says nothing
        // about whether the dump is complete
        pivotAreas: m.pivots.map((p) => parseRange(String(p.location ?? ''))).filter(Boolean),
      });
    }
    return { names, sheets };
  } catch {
    return null;
  }
}

/**
 * Reasons why `dump` cannot be what Excel would report for the package described by `facts` (empty = consistent).
 * Only normally opened workbooks are compared: a repaired file legitimately differs from its package.
 * @param {Awaited<ReturnType<typeof packageFacts>>} facts
 * @param {any} dump an inspect result
 * @returns {string[]}
 */
export function dumpIssues(facts, dump) {
  const issues = [];
  if (!facts || !dump || dump.openMode !== 'normal' || !Array.isArray(dump.sheets)) return issues;

  // the COM side reports reads it could not complete (Excel busy): those are holes in the dump, never facts
  if (Number(dump.transientFailures ?? 0) > 0) issues.push(`Excel が処理中で ${dump.transientFailures} 回の読み取りに失敗しました（busy）`);
  if (dump.wrongWorkbook) issues.push(`開いたブックが検査対象のファイルではありません（${dump.wrongWorkbook}）`);
  if (dump.nameCountMismatch) issues.push(`名前の定義の数が読み取り中に変わりました（${dump.nameCountMismatch}）`);

  // defined names
  const have = new Set(arr(dump.names).map((n) => `${n.scope ?? ''}!${n.name}`));
  const lostNames = facts.names.filter((n) => !have.has(`${n.scope ?? ''}!${n.name}`)).map((n) => n.name);
  if (lostNames.length) issues.push(`ファイルにある名前の定義が Excel の読み取りにありません: ${lostNames.slice(0, 5).join(', ')}`);

  for (const f of facts.sheets) {
    const d = sheetByName(dump, f.name);
    if (!d) {
      issues.push(`シート「${f.name}」が Excel の読み取りにありません`);
      continue;
    }
    // formula cells: present, and pointing at the same ranges
    if (!d.capped && f.formulas.size) {
      let missing = 0;
      const different = [];
      for (const [addr, sig] of f.formulas) {
        const ef = d.formulas?.[addr];
        if (ef == null) {
          missing++;
          if (different.length < 3) different.push(`${f.name}!${addr} (Excel: 数式なし)`);
          continue;
        }
        if (refSignature(ef, f.name) !== sig) {
          missing++;
          if (different.length < 3) different.push(`${f.name}!${addr}: ファイル [${sig}] / Excel [${refSignature(ef, f.name)}]`);
        }
      }
      if (missing >= 3 && missing / f.formulas.size > 0.02) {
        issues.push(`「${f.name}」の数式 ${f.formulas.size} 件のうち ${missing} 件が、ファイルの中身と違う参照先・数式なしで読み取られました（例: ${different.join(' ; ')}）`);
      }
    }
    // text constants (the header row above all: every "which header does this column have" judgement reads them)
    if (!d.capped && f.texts.size >= 5) {
      let found = 0;
      let skipped = 0;
      const absent = [];
      for (const [addr, text] of f.texts) {
        if (inPivot(f, addr)) {
          skipped++;
          continue;
        }
        const v = d.values?.[addr];
        if (typeof v === 'string' && normText(v) === text) found++;
        else if (absent.length < 3) absent.push(`${addr}「${text.slice(0, 12)}」`);
      }
      const judged = f.texts.size - skipped;
      if (judged >= 5 && found / judged < 0.9) {
        issues.push(`「${f.name}」の文字 ${judged} 件のうち ${judged - found} 件が Excel の読み取りにありません（例: ${absent.join(' ')}）`);
      }
    }
    const presence = (what, inPackage, inDump) => {
      if (inPackage > 0 && inDump === 0) issues.push(`「${f.name}」: ファイルには${what}が ${inPackage} 件あるのに、Excel の読み取りでは 0 件です`);
    };
    presence('条件付き書式', f.cf, arr(d.cf).length);
    presence('入力規則', f.dv, arr(d.dv).length);
    presence('グラフ', f.charts, arr(d.charts).length);
    presence('ピボットテーブル', f.pivots, arr(d.pivots).length);
    presence('テーブル', f.tables, arr(d.tables).length);
  }
  return issues;
}

/** Two dumps of the same file from separate Excel instances must agree: the differences (empty = identical). */
export function dumpDifferences(a, b) {
  const out = [];
  const strip = (d) => {
    const rest = { ...(d ?? {}) };
    delete rest.attempts;
    delete rest.suspicious;
    return rest;
  };
  const x = strip(a);
  const y = strip(b);
  for (const k of new Set([...Object.keys(x), ...Object.keys(y)])) {
    if (k === 'sheets') continue;
    if (JSON.stringify(x[k]) !== JSON.stringify(y[k])) out.push(`${k}`);
  }
  const names = new Set([...arr(x.sheets).map((s) => s.name), ...arr(y.sheets).map((s) => s.name)]);
  for (const n of names) {
    const sa = sheetByName(x, n);
    const sb = sheetByName(y, n);
    if (!sa || !sb) {
      out.push(`sheet ${n}`);
      continue;
    }
    for (const k of new Set([...Object.keys(sa), ...Object.keys(sb)])) {
      if (JSON.stringify(sa[k]) !== JSON.stringify(sb[k])) out.push(`sheet ${n}: ${k}`);
    }
  }
  return out;
}
