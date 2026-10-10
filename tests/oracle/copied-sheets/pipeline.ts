/**
 * The app's pipeline in Node for the copied-sheets oracle (task V9): readBook -> keys -> joinSources -> buildPremerged (includeSourceSheets)
 * -> readBook -> createEditor, exactly like src/state/controller.ts merge(). A source file can take part in two ways:
 *   - 'join' : an ordinary participant (key analysis, imported columns, appended rows);
 *   - 'copy' : takes part with NO imported column — only its sheets are copied (what a file with a key but all columns unticked does, and the
 *              way to put any foreign file next to the merge without it needing a matching key).
 * The order of `others` is the card order: the copies come in that order.
 */
import fs from 'node:fs';
import path from 'node:path';
import type { ImportColumnSpec, KeySource, PremergeInput } from '../../../src/core/contracts';
import { createEditor } from '../../../src/core/doc';
import { analyzeSources, defaultPlan, joinSources } from '../../../src/core/keys';
import type { MergedDoc, SourceBook } from '../../../src/core/types';
import { readBook } from '../../../src/core/xlsx';
import { buildPremerged, type FormulaTrace } from '../../../src/core/xlsx/premerge';
import { loadKeySourceBytes, type Merged } from '../consistency/lib';

export interface FileSpec {
  /** absolute path of the file on disk */
  path: string;
  /** the name the app sees (the dropped file's name); default: the file's own name */
  fileName?: string;
  /** 'join' (default) or 'copy' (no imported column) */
  mode?: 'join' | 'copy';
}

export interface PipelineParams {
  joinType: 'left' | 'full';
  /** PremergeOptions.dropOtherSheets (the app's default is true) */
  dropOtherSheets: boolean;
  includeSourceSheets: boolean;
  imports: 'default' | 'all';
  duplicatePolicy?: 'first' | 'last' | 'skip';
}

export const DEFAULT_PARAMS: PipelineParams = { joinType: 'left', dropOtherSheets: false, includeSourceSheets: true, imports: 'default' };

const COLORS = ['#2563eb', '#15803d', '#c2410c', '#7c3aed', '#0e7490'];

export const baseNameOf = (p: string): string => path.basename(p);

export async function loadSpec(spec: FileSpec, id: string): Promise<KeySource> {
  return loadKeySourceBytes(new Uint8Array(fs.readFileSync(spec.path)), spec.fileName ?? baseNameOf(spec.path), id);
}

export const outputFormatOf = (book: SourceBook): 'xlsx' | 'xlsm' => (book.format === 'xlsm' || book.features.vba ? 'xlsm' : 'xlsx');

export interface MergedWithCopies extends Merged {
  /** every participant in card order (join + copy) */
  participants: KeySource[];
  /** the 'copy'-only ones */
  copyOnly: string[];
  format: 'xlsx' | 'xlsm';
}

/** The app's merge of `base` + `others` (a 'join' source without any enabled key tier takes no part, like controller.excludedSourceIds). */
export async function mergeWithCopies(base: KeySource, others: Array<{ ks: KeySource; mode: 'join' | 'copy' }>, p: PipelineParams): Promise<MergedWithCopies> {
  const joiners = others.filter((o) => o.mode === 'join').map((o) => o.ks);
  const analysis = analyzeSources(base, joiners);
  const plan = { ...defaultPlan(analysis), joinType: p.joinType, ...(p.duplicatePolicy ? { duplicatePolicy: p.duplicatePolicy } : {}) };
  const join = joinSources(analysis, plan);
  const joined = new Set(joiners.filter((o) => (plan.tiers[o.id] ?? []).length > 0).map((o) => o.id));
  const participants = others.filter((o) => o.mode === 'copy' || joined.has(o.ks.id)).map((o) => o.ks);
  const used = joiners.filter((o) => joined.has(o.id));
  let imports: ImportColumnSpec[] = join.defaultImports.filter((d) => joined.has(d.sourceId)).map((d) => (d.header ? { sourceId: d.sourceId, col: d.col, header: d.header } : { sourceId: d.sourceId, col: d.col }));
  if (p.imports === 'all') {
    imports = [];
    for (const o of used) for (let c = o.region.firstCol; c <= o.region.lastCol; c++) imports.push({ sourceId: o.id, col: c });
  }
  let rows = join.rows;
  if (!joiners.length) {
    // nothing to join: the merged sheet is the base data sheet as it is
    rows = [];
    for (let r = base.region.firstDataRow; r <= base.region.lastDataRow; r++) rows.push({ baseRow: r, matches: {} });
  }
  const input: PremergeInput = {
    base,
    sources: participants,
    rows,
    imports,
    keyFill: join.keyFill,
    ...(join.keyPairs ? { keyPairs: join.keyPairs } : {}),
    options: { formulaPolicy: 'keep-if-possible', extendTable: true, headerConflict: 'suffix', dropOtherSheets: p.dropOtherSheets, includeSourceSheets: p.includeSourceSheets },
  };
  const trace: FormulaTrace[] = [];
  const result = await buildPremerged(input, undefined, (t) => trace.push(t));
  const book = await readBook(result.bytes, base.book.fileName, { id: 'merged' });
  const sourceInfo: MergedDoc['sourceInfo'] = {};
  [base, ...participants].forEach((k, i) => (sourceInfo[k.id] = { label: k.label, color: COLORS[i % COLORS.length] }));
  let last = base.region.lastDataRow;
  for (const a of result.layout.appended) if (a.row > last) last = a.row;
  const region = { headerRow: base.region.headerRow, firstDataRow: base.region.firstDataRow, lastDataRow: last };
  const editor = createEditor(book, result.layout, sourceInfo, region);
  return {
    base, others: participants, analysis, join, imports, input, result, book, editor, trace, region, sourceInfo,
    participants,
    copyOnly: others.filter((o) => o.mode === 'copy').map((o) => o.ks.id),
    format: outputFormatOf(base.book),
  };
}
