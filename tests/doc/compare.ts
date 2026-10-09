// Shared by invariant.test.ts and fuzz.test.ts: compares the doc with a package read back by readBook.
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { numToCol } from '../../src/core/formula/excelRef';
import type { Cell, EditorSnapshot, FeatureRule, MergedDoc, SourceBook } from '../../src/core/types';
import { styleKey as canonStyle } from '../../src/core/doc/insertStyle';

const here = path.dirname(fileURLToPath(import.meta.url));
export const golden = (f: string) => new Uint8Array(fs.readFileSync(path.join(here, '../fixtures/golden', f)));
export const idAt = (s: EditorSnapshot, i: number) => s.doc.columns[Math.max(0, Math.min(i, s.doc.columns.length - 1))].id;

const styleKey = (book: { styles: SourceBook['styles'] } | MergedDoc, xf: number | undefined) => canonStyle(book.styles[xf ?? 0]);
const valueOf = (c: Cell | undefined) => (c && c.f === undefined && c.v !== null && c.v !== '' ? `${c.t === 'd' ? 'n' : c.t}:${String(c.v)}` : '');

export function compare(doc: MergedDoc, book: SourceBook): string[] {
  const diffs: string[] = [];
  const sheet = book.sheets.find((x) => x.name === doc.sheetName)!;
  const rows = Math.max(doc.rowCount, sheet.maxRow);
  doc.columns.forEach((col, i) => {
    for (let r = 1; r <= rows; r++) {
      const a = col.cells[r - 1];
      const b = sheet.rows[r - 1]?.[i];
      const where = `${numToCol(i + 1)}${r}`;
      if ((a?.f ?? null) !== (b?.f ?? null)) diffs.push(`${where} formula doc ${a?.f} file ${b?.f}`);
      if (valueOf(a) !== valueOf(b)) diffs.push(`${where} value doc ${valueOf(a)} file ${valueOf(b)}`);
      // the cached result of a formula. Neither side recalculates, so the caches must stay equal — except a dependent of a
      // deleted column: the doc shows #REF! (what Excel computes on load) while the patcher keeps the old cache (reported)
      if (a?.f !== undefined && b?.f !== undefined && !(a.t === 'e' && a.v === '#REF!')) {
        const ka = a.t === 'z' ? '' : `${a.t === 'd' ? 'n' : a.t}:${a.v}`;
        const kb = b.t === 'z' ? '' : `${b.t === 'd' ? 'n' : b.t}:${b.v}`;
        if (ka !== kb) diffs.push(`${where} cached value doc ${ka} file ${kb}`);
      }
      if ((a?.arr ?? null) !== (b?.arr ?? null)) diffs.push(`${where} array doc ${a?.arr} file ${b?.arr}`);
      if (!!a?.dyn !== !!b?.dyn) diffs.push(`${where} dynamic array doc ${!!a?.dyn} file ${!!b?.dyn}`);
      const sa = styleKey(doc, a?.s ?? col.props.s);
      const sb = styleKey(book, b?.s ?? sheet.colProps[i]?.s);
      if (sa !== sb) diffs.push(`${where} style doc ${sa} file ${sb}`);
    }
    const wa = col.props.width;
    const wb = sheet.colProps[i]?.width;
    if ((wa === undefined) !== (wb === undefined) || (wa !== undefined && wb !== undefined && Math.abs(wa - wb) > 1e-6)) diffs.push(`${numToCol(i + 1)} width doc ${wa} file ${wb}`);
    // ColProps.customWidth = set by the user WITHOUT bestFit (the reader drops the flag of an auto-fitted column), so a
    // width the patcher auto-widened (bestFit customWidth) must still be auto in the file, and a setColumnWidth one custom
    const ca = !!col.props.customWidth;
    const cb = !!sheet.colProps[i]?.customWidth;
    if (ca !== cb) diffs.push(`${numToCol(i + 1)} customWidth doc ${ca} file ${cb}`);
    // the other column properties: hidden, outline level, column style
    const fp = sheet.colProps[i];
    if (!!col.props.hidden !== !!fp?.hidden) diffs.push(`${numToCol(i + 1)} hidden doc ${!!col.props.hidden} file ${!!fp?.hidden}`);
    if ((col.props.outlineLevel ?? 0) !== (fp?.outlineLevel ?? 0)) diffs.push(`${numToCol(i + 1)} outlineLevel doc ${col.props.outlineLevel ?? 0} file ${fp?.outlineLevel ?? 0}`);
    const cs1 = styleKey(doc, col.props.s);
    const cs2 = styleKey(book, fp?.s);
    if (cs1 !== cs2) diffs.push(`${numToCol(i + 1)} column style doc ${cs1} file ${cs2}`);
  });
  for (let c = doc.columns.length; c < sheet.maxCol; c++) {
    for (const row of sheet.rows) if (row?.[c]?.f !== undefined || valueOf(row?.[c])) diffs.push(`${numToCol(c + 1)}: file has content beyond the doc's columns`);
  }
  // other sheets: exactly the doc's formulas
  const docExt = new Map(doc.externalFormulas.map((e) => [`${e.sheet}!${e.col},${e.row}`, e.f]));
  for (const sh of book.sheets) {
    if (sh.name === doc.sheetName) continue;
    sh.rows.forEach((row, ri) =>
      row?.forEach((cell, ci) => {
        if (cell?.f === undefined) return;
        const k = `${sh.name}!${ci + 1},${ri + 1}`;
        if (docExt.get(k) !== cell.f) diffs.push(`${k} doc ${docExt.get(k)} file ${cell.f}`);
        docExt.delete(k);
      }),
    );
  }
  for (const [k, f] of docExt) diffs.push(`${k} doc ${f} file (no formula)`);
  // names (the filter database follows the autofilter; compared through the autofilter)
  const nm = (n: { name: string; scope: string | null }) => `${n.scope ?? ''}!${n.name}`;
  const fileNames = new Map(book.names.map((n) => [nm(n), n.refersTo]));
  for (const n of doc.names) {
    if (/_FilterDatabase$/i.test(n.name)) continue;
    if (fileNames.get(nm(n)) !== n.refersTo) diffs.push(`name ${nm(n)} doc ${n.refersTo} file ${fileNames.get(nm(n))}`);
  }
  // features
  const f1 = doc.features;
  const f2 = sheet.features;
  const boxes = (l: Array<{ r1: number; c1: number; r2: number; c2: number }>) => l.map((b) => `${b.r1},${b.c1},${b.r2},${b.c2}`).sort().join(' ');
  if (boxes(f1.merges) !== boxes(f2.merges)) diffs.push(`merges doc ${boxes(f1.merges)} file ${boxes(f2.merges)}`);
  if (boxes(f1.arrays) !== boxes(f2.arrays)) diffs.push(`arrays doc ${boxes(f1.arrays)} file ${boxes(f2.arrays)}`);
  if ((f1.autoFilter ?? null) !== (f2.autoFilter ?? null)) diffs.push(`autoFilter doc ${f1.autoFilter} file ${f2.autoFilter}`);
  if ((f1.freeze?.cols ?? 0) !== (f2.freeze?.cols ?? 0)) diffs.push(`freeze doc ${JSON.stringify(f1.freeze)} file ${JSON.stringify(f2.freeze)}`);
  const tables = (l: typeof f1.tables) => JSON.stringify(l.map((t) => [t.name, t.ref, t.columns]));
  if (tables(f1.tables) !== tables(f2.tables)) diffs.push(`tables doc ${tables(f1.tables)} file ${tables(f2.tables)}`);
  const links = (l: typeof f1.hyperlinks) => l.map((h) => `${h.ref}|${h.target ?? ''}|${h.location ?? ''}`).sort().join(' ');
  if (links(f1.hyperlinks) !== links(f2.hyperlinks)) diffs.push(`hyperlinks doc ${links(f1.hyperlinks)} file ${links(f2.hyperlinks)}`);
  // CF / DV rules of the merged sheet and of the other sheets: same kind, storage, type, area, formulas (order free)
  const ruleKey = (r: FeatureRule) =>
    [r.kind, r.ext ? 'x14' : 'classic', r.type ?? '', r.sqref.replace(/\$/g, '').trim().split(/\s+/).sort().join(' '), r.formulas.join('\u0001')].join(' | ');
  const ruleDiff = (where: string, a: FeatureRule[], b: FeatureRule[]) => {
    const ka = a.map(ruleKey);
    const kb = b.map(ruleKey);
    const only = (x: string[], y: string[]) => {
      const left = y.slice();
      return x.filter((k) => {
        const at = left.indexOf(k);
        if (at < 0) return true;
        left.splice(at, 1);
        return false;
      });
    };
    const da = only(ka, kb);
    const db = only(kb, ka);
    if (da.length || db.length) diffs.push(`${where} rules only in doc [${da.join(' ; ')}] only in file [${db.join(' ; ')}]`);
  };
  ruleDiff('merged sheet', f1.rules, f2.rules);
  for (const sh of book.sheets) {
    if (sh.name === doc.sheetName) continue;
    ruleDiff(`sheet ${sh.name}`, doc.externalRules.filter((e) => e.sheet === sh.name).map((e) => e.rule), sh.features.rules);
  }
  // row properties (no op touches them): heights, hidden, outline level, row style
  for (let r = 1; r <= doc.rowCount; r++) {
    const pa = doc.rowProps[r - 1];
    const pb = sheet.rowProps[r - 1];
    if ((pa?.height ?? null) !== (pb?.height ?? null)) diffs.push(`row ${r} height doc ${pa?.height} file ${pb?.height}`);
    if (!!pa?.hidden !== !!pb?.hidden) diffs.push(`row ${r} hidden doc ${!!pa?.hidden} file ${!!pb?.hidden}`);
    if ((pa?.outlineLevel ?? 0) !== (pb?.outlineLevel ?? 0)) diffs.push(`row ${r} outlineLevel doc ${pa?.outlineLevel ?? 0} file ${pb?.outlineLevel ?? 0}`);
    const sa = styleKey(doc, pa?.s);
    const sb = styleKey(book, pb?.s);
    if (sa !== sb) diffs.push(`row ${r} style doc ${sa} file ${sb}`);
  }
  // names the file has but the doc lost
  const docNames = new Set(doc.names.map(nm));
  for (const n of book.names) if (!/_FilterDatabase$/i.test(n.name) && !docNames.has(nm(n))) diffs.push(`name ${nm(n)} in the file (${n.refersTo}) but not in the doc`);
  return diffs;
}

