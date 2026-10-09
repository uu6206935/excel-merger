/**
 * Test fixtures for the UI state tests: realistic books / analysis / join / editor snapshots from the
 * ?mock=1 services (src/state/mock.ts), plus a tiny store that runs the real reducer.
 */
import type { JoinPlan, JoinResult, KeyAnalysis, KeySource } from '../../src/core/contracts';
import type { EditorSnapshot, SourceBook } from '../../src/core/types';
import { createMockServices } from '../../src/state/mock';
import { initialState, reducer, type Action } from '../../src/state/reducer';
import type { Services } from '../../src/state/services';
import type { AppState } from '../../src/state/types';

export const svc: Services = createMockServices('?mock=1&delay=0');

export async function book(name: string, id: string): Promise<SourceBook> {
  return svc.readBook(new Uint8Array(16), name, { id });
}

export async function sources(): Promise<{ base: KeySource; others: KeySource[]; books: SourceBook[] }> {
  const b1 = await book('学生回答.xlsx', 's1');
  const b2 = await book('職場回答.xlsx', 's2');
  const b3 = await book('人事回答.xlsx', 's3');
  const ks = (b: SourceBook, label: string): KeySource => ({ id: b.id, book: b, region: svc.detectRegion(b, svc.pickDataSheet(b)), label });
  return { base: ks(b1, '学生回答'), others: [ks(b2, '職場回答'), ks(b3, '人事回答')], books: [b1, b2, b3] };
}

export interface Joined {
  base: KeySource;
  others: KeySource[];
  analysis: KeyAnalysis;
  plan: JoinPlan;
  join: JoinResult;
}

export async function joined(): Promise<Joined> {
  const { base, others } = await sources();
  const analysis = svc.analyzeSources(base, others);
  const plan = svc.defaultPlan(analysis);
  const join = svc.joinSources(analysis, plan);
  return { base, others, analysis, plan, join };
}

export async function editor(): Promise<{ j: Joined; snapshot: EditorSnapshot; mergedBook: SourceBook }> {
  const j = await joined();
  const res = await svc.buildPremerged({
    base: j.base,
    sources: j.others,
    rows: j.join.rows,
    imports: j.join.defaultImports,
    keyFill: j.join.keyFill,
    options: { formulaPolicy: 'keep-if-possible', extendTable: true, headerConflict: 'suffix' },
  });
  const mergedBook = await svc.readBook(res.bytes, '学生回答.xlsx', { id: 'merged' });
  const snapshot = svc.createEditor(
    mergedBook,
    res.layout,
    { s1: { label: '学生回答', color: '#2563eb' }, s2: { label: '職場回答', color: '#15803d' }, s3: { label: '人事回答', color: '#c2410c' } },
    { headerRow: j.base.region.headerRow, firstDataRow: j.base.region.firstDataRow, lastDataRow: j.base.region.lastDataRow },
  );
  return { j, snapshot, mergedBook };
}

/** Real reducer + recorded actions. */
export function store(start: AppState = initialState) {
  let state = start;
  const actions: Action[] = [];
  return {
    get state() {
      return state;
    },
    actions,
    dispatch(a: Action) {
      actions.push(a);
      state = reducer(state, a);
    },
    getState: () => state,
  };
}

export function file(name: string, size = 1000, lastModified = 1_700_000_000_000): File {
  return new File([new Uint8Array(size)], name, { lastModified });
}
