/**
 * Shared pipeline helpers of the consistency oracle (task V1): the app's own pipeline in Node
 * (readBook -> keys -> joinSources -> buildPremerged -> createEditor), exactly like src/state/controller.ts merge().
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import type { ImportColumnSpec, JoinResult, KeyAnalysis, KeySource, PremergeInput, PremergeResult } from '../../../src/core/contracts';
import { createEditor } from '../../../src/core/doc';
import { analyzeSources, defaultPlan, detectRegion, joinSources, pickDataSheet } from '../../../src/core/keys';
import type { EditorSnapshot, MergedDoc, SourceBook } from '../../../src/core/types';
import { readBook } from '../../../src/core/xlsx';
import { buildPremerged, type FormulaTrace } from '../../../src/core/xlsx/premerge';

export const HERE = path.dirname(fileURLToPath(import.meta.url));
export const REPO = path.resolve(HERE, '../../..');
export const SAMPLES = path.join(REPO, 'samples');
export const FIXTURES = path.join(REPO, 'tests', 'fixtures', 'formulaimport');

export async function loadKeySourceBytes(bytes: Uint8Array, fileName: string, id: string, opts: { sheet?: string; headerRow?: number } = {}): Promise<KeySource> {
  const book = await readBook(bytes, fileName, { id });
  const sheet = opts.sheet ?? pickDataSheet(book);
  const region = opts.headerRow ? detectRegion(book, sheet, { headerRow: opts.headerRow }) : detectRegion(book, sheet);
  return { id, book, region, label: fileName.replace(/\.[^.]+$/, '') };
}

export async function loadKeySourceFile(file: string, id: string, opts: { sheet?: string; headerRow?: number } = {}): Promise<KeySource> {
  return loadKeySourceBytes(new Uint8Array(fs.readFileSync(file)), path.basename(file), id, opts);
}

export interface MergeParams {
  joinType: 'left' | 'full';
  /** PremergeOptions.dropOtherSheets (the app's default is true) */
  dropOtherSheets: boolean;
  /** 'default' = the app's default import list (key columns excluded), 'all' = every column of every source */
  imports: 'default' | 'all';
  /** hand JoinResult.keyPairs to buildPremerged (the lead wires it in the controller) */
  keyPairs: boolean;
  duplicatePolicy?: 'first' | 'last' | 'skip';
}

export const DEFAULT_PARAMS: MergeParams = { joinType: 'left', dropOtherSheets: true, imports: 'default', keyPairs: true };

export interface Merged {
  base: KeySource;
  others: KeySource[];
  analysis: KeyAnalysis;
  join: JoinResult;
  imports: ImportColumnSpec[];
  input: PremergeInput;
  result: PremergeResult;
  /** the premerged package read back (what createEditor is given) */
  book: SourceBook;
  editor: EditorSnapshot;
  trace: FormulaTrace[];
  region: { headerRow: number; firstDataRow: number; lastDataRow: number };
  sourceInfo: MergedDoc['sourceInfo'];
}

const COLORS = ['#2563eb', '#15803d', '#c2410c', '#7c3aed', '#0e7490'];

/** The app's merge of `base` + `others` (every other file with a usable key takes part, like controller.merge). */
export async function mergeAll(base: KeySource, others: KeySource[], p: MergeParams): Promise<Merged> {
  const analysis = analyzeSources(base, others);
  const plan = { ...defaultPlan(analysis), joinType: p.joinType, ...(p.duplicatePolicy ? { duplicatePolicy: p.duplicatePolicy } : {}) };
  const join = joinSources(analysis, plan);
  // a source without any enabled tier takes no part (controller.excludedSourceIds)
  const used = others.filter((o) => (plan.tiers[o.id] ?? []).length > 0);
  let imports: ImportColumnSpec[] = join.defaultImports.filter((d) => used.some((o) => o.id === d.sourceId)).map((d) => (d.header ? { sourceId: d.sourceId, col: d.col, header: d.header } : { sourceId: d.sourceId, col: d.col }));
  if (p.imports === 'all') {
    imports = [];
    for (const o of used) for (let c = o.region.firstCol; c <= o.region.lastCol; c++) imports.push({ sourceId: o.id, col: c });
  }
  const input: PremergeInput = {
    base,
    sources: used,
    rows: join.rows,
    imports,
    keyFill: join.keyFill,
    ...(p.keyPairs && join.keyPairs ? { keyPairs: join.keyPairs } : {}),
    options: { formulaPolicy: 'keep-if-possible', extendTable: true, headerConflict: 'suffix', dropOtherSheets: p.dropOtherSheets },
  };
  const trace: FormulaTrace[] = [];
  const result = await buildPremerged(input, undefined, (t) => trace.push(t));
  const book = await readBook(result.bytes, base.book.fileName, { id: 'merged' });
  const sourceInfo: MergedDoc['sourceInfo'] = {};
  [base, ...used].forEach((k, i) => (sourceInfo[k.id] = { label: k.label, color: COLORS[i % COLORS.length] }));
  let last = base.region.lastDataRow;
  for (const a of result.layout.appended) if (a.row > last) last = a.row;
  const region = { headerRow: base.region.headerRow, firstDataRow: base.region.firstDataRow, lastDataRow: last };
  const editor = createEditor(book, result.layout, sourceInfo, region);
  return { base, others: used, analysis, join, imports, input, result, book, editor, trace, region, sourceInfo };
}

/** Excel's column letters */
export function colLetter(n: number): string {
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/** a short stable string of a value for messages */
export const show = (v: unknown): string => (v === null || v === undefined ? '(empty)' : typeof v === 'string' ? JSON.stringify(v) : typeof v === 'object' ? JSON.stringify(v) : String(v));

/** names of the sheets that hold a PivotTable (their cells are rendered by Excel from the pivot definition, not formulas) */
export async function pivotSheetNames(bytes: Uint8Array): Promise<string[]> {
  const JSZip = (await import('jszip')).default;
  const zip = await JSZip.loadAsync(bytes);
  const text = async (p: string): Promise<string> => (zip.file(p) ? await zip.file(p)!.async('string') : '');
  const wb = await text('xl/workbook.xml');
  const rels = await text('xl/_rels/workbook.xml.rels');
  const out: string[] = [];
  for (const m of wb.matchAll(/<sheet\b[^>]*>/g)) {
    const name = /\bname="([^"]*)"/.exec(m[0])?.[1];
    const rid = /\br:id="([^"]*)"/.exec(m[0])?.[1];
    if (!name || !rid) continue;
    const rel = [...rels.matchAll(/<Relationship\b[^>]*>/g)].map((x) => x[0]).find((x) => x.includes(` Id="${rid}"`));
    const target = rel && /\bTarget="([^"]*)"/.exec(rel)?.[1];
    if (!target) continue;
    const part = target.startsWith('/') ? target.slice(1) : `xl/${target}`;
    const relsPart = part.replace(/([^/]+)$/, '_rels/$1.rels');
    if (/pivotTable/i.test(await text(relsPart))) out.push(name.replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'"));
  }
  return out;
}
