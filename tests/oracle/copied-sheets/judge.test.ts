/**
 * The judge of the copied-sheets oracle (judge.ts) on hand-made Excel dumps — no Excel: a planted fault must be found, and an honest copy must not
 * be reported. This is the oracle's own test ("does the oracle see a deliberately planted bug?").
 */
import fs from 'node:fs';
import path from 'node:path';
import { describe, expect, it } from 'vitest';
import type { SourceBook } from '../../../src/core/types';
import { isBadDumpOf, isShortDump, packageLastCells, type BookDump, type BookSheet, type Deep } from './excel';
import { comparePair, expectedSheets, judgeCopies, judgeCopyNames, judgeMergedErrors, judgeNames, parseNames, stripPrefix, type CopyRec, type PairOptions, type SourceRec } from './judge';

const src = (id: string, fileName: string, sheets: string[], extra: Partial<SourceRec> = {}): SourceRec => ({ id, fileName, path: `C:\\${fileName}`, mode: 'join', fidelity: 'full', sheets, date1904: false, externalLinks: false, ...extra });

function sheet(name: string, cells: Record<string, string | number>, o: Partial<BookSheet> = {}): BookSheet {
  const rows = 4;
  const cols = 3;
  const v: BookSheet['v'] = Array.from({ length: rows }, () => Array<null | string | number>(cols).fill(null));
  const f: string[][] = Array.from({ length: rows }, () => Array<string>(cols).fill(''));
  for (const [a, x] of Object.entries(cells)) {
    const m = /^([A-C])(\d)$/.exec(a)!;
    if (typeof x === 'string' && x.startsWith('=')) f[Number(m[2]) - 1][m[1].charCodeAt(0) - 65] = x;
    else v[Number(m[2]) - 1][m[1].charCodeAt(0) - 65] = x;
  }
  return { name, type: -4167, visible: -1, rows, cols, v, f, ...o };
}
const deep = (o: Partial<Deep> = {}): Deep => ({
  tab: -1, widths: [8.4, 8.4, 8.4], hiddenCols: [], hiddenRows: [], cf: 0, dv: '', links: 0, shapes: [], shapesOther: 0, charts: [], tables: [], notes: [], pivots: [],
  look: [{ r: 1, c: 1, nf: 'General', fn: 'Calibri', fs: 11, fb: false, fi: false, fc: 0, ic: 16777215, ip: 1, ha: 1, va: -4107, wr: false, mg: false, bb: -4142, bl: -4142, bt: -4142, br: -4142 }],
  ...o,
});
const book = (sheets: BookSheet[], names: BookDump['names'] = [], o: Partial<BookDump> = {}): BookDump => ({ open: 'normal', log: '', date1904: false, names, sheets, ...o });

const opts: PairOptions = { dateShift: 0, lenient: false, fidelity: 'full', complete: true, limits: '', isDateAt: () => false };
const classes = (l: Array<{ cls: string }>) => [...new Set(l.map((x) => x.cls))].sort();

describe('comparePair: an honest copy is clean, every planted fault is found', () => {
  const s = () => sheet('面接評価', { A1: 'ID', B1: '点', A2: 'x1', B2: 5, C2: '=B2*2', C3: '=SUM(B2:B2)' }, { deep: deep() });
  const honest = () => sheet('02_面接評価_面接評価', { A1: 'ID', B1: '点', A2: 'x1', B2: 5, C2: "='02_面接評価_面接評価'!B2*2", C3: '=SUM(B2:B2)' }, { deep: deep() });

  it('equal sheets (formulas differ only by the renamed sheet prefix)', () => {
    expect(comparePair(s(), honest(), opts)).toEqual([]);
  });

  it('a changed value / formula / visibility / new #REF!', () => {
    const a = honest();
    a.v[1][1] = 6;
    expect(comparePair(s(), a, opts).map((i) => i.cls)).toContain('copy-value');
    const b = honest();
    b.f[1][2] = "='02_面接評価_面接評価'!B2*3";
    expect(comparePair(s(), b, opts).map((i) => i.cls)).toContain('copy-formula');
    const c = honest();
    c.visible = 0;
    expect(comparePair(s(), c, opts).map((i) => i.cls)).toContain('copy-visibility');
    const d = honest();
    d.v[1][1] = { e: -2146826265 } as never;
    expect(comparePair(s(), d, opts).map((i) => i.cls)).toEqual(expect.arrayContaining(['copy-new-error', 'copy-value']));
    // a formula that became a value
    const e = honest();
    e.f[1][2] = '';
    expect(comparePair(s(), e, opts).map((i) => i.cls)).toContain('copy-formula-value');
    expect(comparePair(s(), e, { ...opts, complete: false }).map((i) => i.cls)).toContain('copy-formula-lenient');
  });

  it('dates between the 1904 and 1900 systems; a source with external links is judged leniently; csv / xls: no formulas / formats', () => {
    const a = s();
    const b = honest();
    b.v[1][1] = 5 + 1462;
    expect(comparePair(a, b, opts).map((i) => i.cls)).toContain('copy-value');
    expect(comparePair(a, b, { ...opts, dateShift: 1462, isDateAt: (r, c) => r === 2 && c === 2 })).toEqual([]);
    expect(comparePair(a, b, { ...opts, lenient: true }).map((i) => i.cls)).toEqual(['copy-value-lenient']);
    const c = honest();
    c.f[1][2] = '';
    c.deep!.cf = 3;
    expect(comparePair(a, c, { ...opts, fidelity: 'values' })).toEqual([]);
  });

  it('formats Excel renders: number format, fill, borders, widths, hidden columns / rows, tab colour, rules, links', () => {
    const mk = (f: (d: Deep) => void): BookSheet => {
      const h = honest();
      f(h.deep!);
      return h;
    };
    const cls = (x: BookSheet) => comparePair(s(), x, opts).map((i) => i.cls);
    expect(cls(mk((d) => (d.look![0].nf = 'yyyy/m/d')))).toContain('copy-format');
    expect(cls(mk((d) => (d.look![0].ic = 255)))).toContain('copy-format');
    expect(cls(mk((d) => (d.look![0].bb = 1)))).toContain('copy-format');
    expect(cls(mk((d) => (d.widths![1] = 30)))).toContain('copy-format');
    expect(cls(mk((d) => (d.widths![1] = 9.5)))).toEqual([]); // within 1.6 pt
    expect(cls(mk((d) => (d.hiddenCols = [2])))).toContain('copy-format');
    expect(cls(mk((d) => (d.hiddenRows = [3])))).toContain('copy-format');
    expect(cls(mk((d) => (d.tab = 255)))).toContain('copy-format');
    expect(cls(mk((d) => (d.cf = 2)))).toContain('copy-format');
    expect(cls(mk((d) => (d.dv = 'A2:A4')))).toContain('copy-format');
    expect(cls(mk((d) => (d.links = 1)))).toContain('copy-format');
    // the font NAME of a theme font is only reported (INFO)
    expect(cls(mk((d) => (d.look![0].fn = 'ＭＳ Ｐゴシック')))).toEqual(['copy-font-name']);
  });

  it('objects: charts / shapes / tables / notes / pivot tables / chart sheets; what the copier says it left out is only reported', () => {
    const base = () => sheet('S', { A1: 1 }, { deep: deep({ charts: [{ type: '51', series: [{ v: '1;2', x: 'a;b' }] }], shapes: [{ t: 13, n: 'Picture 1', x: '' }], tables: [{ n: 'T1', a: 'A1:B3', cols: 'a|b', totals: false }], notes: [{ a: 'A1', t: 'memo' }], pivots: [{ n: 'P', a: 'A3:B5', src: 'S!R1C1:R9C3', v: 'x;1' }] }) });
    const copy = (f: (d: Deep) => void): BookSheet => {
      const c = base();
      c.name = 'f_S';
      f(c.deep!);
      return c;
    };
    const cls = (x: BookSheet, limits = '') => comparePair(base(), x, { ...opts, limits }).map((i) => i.cls);
    expect(cls(copy(() => undefined))).toEqual([]);
    expect(cls(copy((d) => (d.charts![0].series[0].v = '1;3')))).toContain('copy-object');
    expect(cls(copy((d) => (d.charts = [])))).toContain('copy-object');
    expect(cls(copy((d) => (d.shapes = [])))).toContain('copy-object');
    expect(cls(copy((d) => (d.tables![0].a = 'A1:B4')))).toContain('copy-object');
    expect(cls(copy((d) => (d.tables![0].n = 'T1_2')))).toEqual([]); // a table of the second copy of a file is named T1_2
    expect(cls(copy((d) => (d.notes = [])))).toContain('copy-object');
    expect(cls(copy((d) => (d.pivots![0].v = 'x;2')))).toContain('copy-object');
    expect(cls(copy((d) => (d.pivots![0].src = "'f_S'!R1C1:R9C3")))).toEqual([]); // the source follows the renamed sheet
    // reported, not failed, when the copier names the kind for this sheet
    expect(cls(copy((d) => (d.charts = [])), 'グラフ')).toEqual(['copy-object-leftout']);
    expect(cls(copy((d) => (d.notes = [])), 'コメントは通常のメモ')).toEqual(['copy-object-leftout']);
    // a chart sheet
    const cs = (v: string) => ({ ...sheet('C', {}, { type: -4109, rows: 0, cols: 0, v: [], f: [], deep: { tab: -1, chartSheet: { type: '51', series: [{ v }] } } }) });
    expect(comparePair(cs('1;2'), { ...cs('1;2'), name: 'f_C' }, opts)).toEqual([]);
    expect(comparePair(cs('1;2'), { ...cs('1;9'), name: 'f_C' }, opts).map((i) => i.cls)).toContain('copy-object');
  });
});

describe('the sheet names: exactly what planSheetNames gives', () => {
  const rec = (drop: boolean, sources: SourceRec[]): CopyRec =>
    ({ id: 't', mergedSheet: '名簿', drop, sources, copies: [], limits: [], pivotSheets: [], files: { pre: '', base: '', js: '', xl: '' } } as unknown as CopyRec);
  const dumps = (m: Record<string, string[]>): Map<string, BookDump> => new Map(Object.entries(m).map(([id, names]) => [id, book(names.map((n) => sheet(n, {})))]));

  it('merged sheet, the base sheets that are kept, then every file in card order; (1) for the later one', () => {
    const r = rec(false, [src('s2', '02_面接評価.xlsx', ['面接評価', '評価テーブル']), src('s3', 'C:\\x\\02_面接評価.xlsx', ['面接評価'])]);
    const e = expectedSheets(r, dumps({ s2: ['面接評価', '評価テーブル'], s3: ['面接評価'] }), book([sheet('名簿', {}), sheet('集計', {}), sheet('マスタ', {})]));
    expect(e.names).toEqual(['名簿', '集計', 'マスタ', '02_面接評価_面接評価', '02_面接評価_評価テーブル', '02_面接評価_面接評価(1)']);
    const dropped = expectedSheets(rec(true, r.sources), dumps({ s2: ['面接評価', '評価テーブル'], s3: ['面接評価'] }), book([sheet('名簿', {}), sheet('集計', {})]));
    expect(dropped.names).toEqual(['名簿', '02_面接評価_面接評価', '02_面接評価_評価テーブル', '02_面接評価_面接評価(1)']);
  });

  it('a name that has to be shortened (31 characters), forbidden characters, a csv / xls file takes its sheets from the reader', () => {
    const r = rec(true, [src('s2', '営業部門の月次実績報告書（2026年度第3四半期確定版）.xlsx', ['面接評価']), src('s3', '売上[確定]:v2.csv', ['売上[確定]:v2'], { fidelity: 'reader' })]);
    const e = expectedSheets(r, dumps({ s2: ['面接評価'], s3: ['売上[確定]:v2'] }), null);
    expect(e.names[1].length).toBeLessThanOrEqual(31);
    expect(e.names[1]).toMatch(/_面接評価$/);
    expect(e.names[2]).toBe('売上_確定__v2_売上_確定__v2');
  });

  it('judgeNames: a wrong order, a missing sheet, a name that was not suffixed, a failed open', () => {
    const r = rec(true, [src('s2', 'a.xlsx', ['X', 'Y'])]);
    const e = expectedSheets(r, dumps({ s2: ['X', 'Y'] }), null);
    expect(e.names).toEqual(['名簿', 'a_X', 'a_Y']);
    const ok = book([sheet('名簿', {}), sheet('a_X', {}), sheet('a_Y', {})]);
    expect(judgeNames(r, 'pre', ok, e)).toEqual([]);
    expect(classes(judgeNames(r, 'js', book([sheet('名簿', {}), sheet('a_Y', {}), sheet('a_X', {})]), e))).toEqual(['names-js']);
    expect(classes(judgeNames(r, 'js', book([sheet('名簿', {}), sheet('a_X', {})]), e))).toEqual(['names-js']);
    expect(classes(judgeNames(r, 'xl', book([], [], { open: 'repair', log: 'removed' }), e))).toEqual(['open-xl']);
    expect(classes(judgeNames(r, 'xl', book([], [], { open: 'failed', log: 'x' }), e))).toEqual(['open-xl']);
  });

  it('chart sheets and hidden sheets are copied: type -4109 counts, a macro sheet (3) does not', () => {
    const r = rec(true, [src('s2', 'a.xlsx', ['X'])]);
    const d = new Map([['s2', book([sheet('X', {}), sheet('G', {}, { type: -4109 }), sheet('M', {}, { type: 3 }), sheet('H', {}, { visible: 2 })])]]);
    expect(expectedSheets(r, d, null).names).toEqual(['名簿', 'a_X', 'a_G', 'a_H']);
  });
});

describe('judgeCopies: per output, per copy', () => {
  const r: CopyRec = { id: 't', mergedSheet: '名簿', drop: true, sources: [src('s2', 'a.xlsx', ['X'])], copies: [{ sourceId: 's2', file: 'a.xlsx', original: 'X', name: 'a_X', suffixed: false, complete: true }], limits: [], pivotSheets: [], files: { pre: '', base: '', js: '', xl: '' } } as unknown as CopyRec;
  const srcDump = book([sheet('X', { A1: 'ID', B2: 5 }, { deep: deep() })]);
  const dumps = new Map([['s2', srcDump]]);
  const books = new Map<string, SourceBook>();
  const exp = expectedSheets(r, dumps, null);

  it('findings carry the engine in the class and name the files; honest outputs give none', () => {
    const good = book([sheet('名簿', {}), sheet('a_X', { A1: 'ID', B2: 5 }, { deep: deep() })]);
    expect(judgeCopies(r, 'js', good, exp, dumps, books).findings).toEqual([]);
    const bad = book([sheet('名簿', {}), sheet('a_X', { A1: 'ID', B2: 6 }, { deep: deep() })]);
    const f = judgeCopies(r, 'xl', bad, exp, dumps, books).findings;
    expect(f.map((x) => x.cls)).toEqual(['copy-value-xl']);
    expect(f[0].detail).toContain('a.xlsx / X -> a_X');
  });

  it('a csv source is only judged through the reader (Excel parses a csv by its own rules)', () => {
    const csv = { ...r, sources: [src('s2', 'a.csv', ['X'], { fidelity: 'reader' })] } as CopyRec;
    const bad = book([sheet('名簿', {}), sheet('a_X', { A1: 'ID', B2: 'x' })]);
    expect(judgeCopies(csv, 'js', bad, expectedSheets(csv, dumps, null), dumps, books).findings).toEqual([]);
  });

  it('defined names: the sheet-scoped names of the source are carried over; a workbook-level name of the source may arrive as a name of the copy; none becomes workbook-level', () => {
    const sd = book([sheet('X', {})], [
      { name: 'X!局所', refers: '=X!$C$3:$C$12', visible: true },
      { name: '基準日', refers: '=Y!$D$1', visible: true },
      { name: '_xlfn.IFERROR', refers: '=#NAME?', visible: false },
      { name: 'X!_FilterDatabase', refers: '=X!$A$1:$B$3', visible: false },
    ]);
    const d = new Map([['s2', sd]]);
    const base = book([sheet('名簿', {})], [{ name: '税率', refers: '=名簿!$A$1', visible: true }]);
    const good = book([sheet('名簿', {}), sheet('a_X', {})], [
      { name: '税率', refers: '=名簿!$A$1', visible: true },
      { name: 'a_X!局所', refers: '=a_X!$C$3:$C$12', visible: true },
      { name: 'a_X!基準日', refers: '=a_Y!$D$1', visible: true },
    ]);
    expect(judgeCopyNames(r, 'pre', good, expectedSheets(r, d, null), d, base)).toEqual([]);
    const lost = book([sheet('名簿', {}), sheet('a_X', {})], [{ name: '税率', refers: '=名簿!$A$1', visible: true }]);
    expect(classes(judgeCopyNames(r, 'js', lost, expectedSheets(r, d, null), d, base))).toEqual(['copy-names-js']);
    const leak = book([sheet('名簿', {}), sheet('a_X', {})], [
      { name: '税率', refers: '=名簿!$A$1', visible: true },
      { name: 'a_X!局所', refers: '=a_X!$C$3:$C$12', visible: true },
      { name: '基準日', refers: '=a_Y!$D$1', visible: true },
    ]);
    expect(classes(judgeCopyNames(r, 'xl', leak, expectedSheets(r, d, null), d, base))).toEqual(['copy-names-leak-xl']);
    expect(parseNames(sd).map((n) => `${n.scope ?? ''}!${n.name}`)).toEqual(['X!局所', '!基準日']);
  });

  it('judgeMergedErrors: #REF! in the merged sheet beyond what the base and the joined sources have', () => {
    const clean = book([sheet('名簿', { A1: 'x' })]);
    expect(judgeMergedErrors(r, clean, clean, dumps)).toEqual([]);
    const withErr = book([sheet('名簿', {}, { v: [[{ e: -2146826265 } as never]], f: [[]] })]);
    expect(classes(judgeMergedErrors(r, withErr, clean, dumps))).toEqual(['merged-new-error']);
  });

  it('stripPrefix', () => {
    expect(stripPrefix("='a_X'!B2+Y!C3+SUM('b c'!A1:A2)")).toBe('=B2+C3+SUM(A1:A2)');
    expect(stripPrefix('=T_2[Col]')).toBe('=T[Col]');
  });
});

describe('a dump that is shorter than its package (Excel answered before it had finished)', () => {
  const file = path.resolve(__dirname, '../../../samples/01_学生名簿.xlsx');
  it.skipIf(!fs.existsSync(file))('is read again; an honest dump is not', async () => {
    const last = await packageLastCells(file);
    expect(last).not.toBeNull();
    const roster = last!.get('名簿')!;
    expect(roster.row).toBe(66); // the summary rows under the list
    expect(roster.col).toBeGreaterThanOrEqual(11);
    const honest = book([sheet('名簿', { A1: 'x' }, { used: [roster.row, roster.col] }), sheet('集計', {}, { used: [60, 20] })]);
    expect(await isShortDump(honest, file)).toBe(false);
    const short = book([sheet('名簿', { A1: 'x' }, { used: [roster.row - 1, roster.col] })]);
    expect(await isShortDump(short, file)).toBe(true);
    expect(await isBadDumpOf(short, file)).toBe(true);
    // a truncated dump (a huge sheet cut on purpose) and a file that is not a package are not judged
    expect(await isShortDump(book([sheet('名簿', {}, { used: [10, 3], trunc: true })]), file)).toBe(false);
    expect(await isShortDump(short, 'C:\\none.csv')).toBe(false);
  });
});
