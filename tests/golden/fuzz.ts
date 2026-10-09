/**
 * Package-level fuzzing of replayLog / buildPremerged (round 4). A mutant is a real package (Excel-authored fixtures of
 * tests/fixtures/golden) with a few random XML / zip level damages. The contract under test, for EVERY mutant:
 *  - the call ends in a PatchError with a Japanese message (an `internal` one = the patcher did not foresee the damage:
 *    counted, the tests pin the count to zero for the seeds they run), or
 *  - in a package in which every part the patcher wrote is well-formed XML, that the reader can read when it could read
 *    the mutant, and that replays once more without an error.
 * Never a thrown TypeError / RangeError, never a package that Excel would have to repair.
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import type { PremergeInput } from '../../src/core/contracts';
import type { CellWrite, LoggedOp } from '../../src/core/types';
import { PatchError } from '../../src/core/xlsx/patch/pkg';
import { xmlProblem } from '../../src/core/xlsx/patch/wellformed';
import { buildPremerged } from '../../src/core/xlsx/premerge';
import { readBook } from '../../src/core/xlsx/read';
import { replayLog } from '../../src/core/xlsx/replay';
import { CASES } from './cases';
import { pack, rng, unpack, type Parts } from './package-tools';
import { insideTableScenario, keyFillScenario, richScenario, scenario } from './premerge-scenario';

const here = path.dirname(fileURLToPath(import.meta.url));
const FIX = path.resolve(here, '../fixtures/golden');
const JP = /[ぁ-んァ-ン一-龥]/;

export type Rng = () => number;

// ---------------------------------------------------------------------------------------------------------------
// mutators
// ---------------------------------------------------------------------------------------------------------------

const JUNK_VALUES = ['', '0', '-1', '1', '2', '16385', '1048577', '99999999999', '1e9', '0.5', 'x', 'A1', 'A1:B2', '1:1', '&', '<', '"', '\u0001', 'TRUE', ' ', '&#0;', '_x0000_', '-', 'XFE1', 'Sheet!A1', '#REF!'];
const JUNK_CHARS = ['&', '<', '>', '"', "'", '\u0000', '\u001f', '&#0;', '&#xD800;', '&foo;', '_x0000_', '\ud800', '￾', '/', '=', ' ', ']]>', '<!--', '<![CDATA[', '<?', 'xmlns:q="u" q:a="1"', '</'];
const JUNK_INTS = ['0', '-1', '1', '2', '3', '16384', '16385', '1048576', '1048577', '99999999999', '4294967296', '0.5', '1e3', '007'];

const pick = <T>(r: Rng, list: readonly T[]): T => list[Math.floor(r() * list.length)];
const between = (r: Rng, a: number, b: number) => a + Math.floor(r() * (b - a + 1));

/** [start, end) of a random element: <tag …/> or <tag …>…</tag> (first closing tag of the same name) */
function randomElement(xml: string, r: Rng): [number, number] | null {
  for (let tries = 0; tries < 20; tries++) {
    const at = xml.indexOf('<', between(r, 0, Math.max(0, xml.length - 1)));
    if (at < 0) continue;
    const m = /^<([A-Za-z_][\w:.-]*)(?=[\s/>])[^>]*?(\/?)>/.exec(xml.slice(at, at + 4000));
    if (!m) continue;
    if (m[2] === '/') return [at, at + m[0].length];
    const close = xml.indexOf(`</${m[1]}>`, at + m[0].length);
    if (close < 0) continue;
    return [at, close + m[1].length + 3];
  }
  return null;
}

type Mutator = (xml: string, r: Rng) => string;

const MUTATORS: Array<[string, Mutator]> = [
  ['truncate', (x, r) => x.slice(0, between(r, 0, x.length))],
  ['deleteSpan', (x, r) => { const i = between(r, 0, x.length); return x.slice(0, i) + x.slice(i + between(r, 1, 60)); }],
  ['insertChar', (x, r) => { const i = between(r, 0, x.length); return x.slice(0, i) + pick(r, JUNK_CHARS) + x.slice(i); }],
  ['junkAttr', (x, r) => {
    const all = [...x.matchAll(/\s[\w:.-]+="([^"]*)"/g)];
    if (!all.length) return x;
    const m = pick(r, all);
    const v = m.index! + m[0].indexOf('"') + 1;
    return x.slice(0, v) + pick(r, JUNK_VALUES) + x.slice(v + m[1].length);
  }],
  ['dropAttr', (x, r) => {
    const all = [...x.matchAll(/\s[\w:.-]+="[^"]*"/g)];
    if (!all.length) return x;
    const m = pick(r, all);
    return x.slice(0, m.index!) + x.slice(m.index! + m[0].length);
  }],
  ['numBump', (x, r) => {
    const all = [...x.matchAll(/(?<![\w.])\d+(?![\w.])/g)];
    if (!all.length) return x;
    const m = pick(r, all);
    return x.slice(0, m.index!) + pick(r, JUNK_INTS) + x.slice(m.index! + m[0].length);
  }],
  ['dropElement', (x, r) => { const e = randomElement(x, r); return e ? x.slice(0, e[0]) + x.slice(e[1]) : x; }],
  ['duplicateElement', (x, r) => { const e = randomElement(x, r); return e ? x.slice(0, e[1]) + x.slice(e[0], e[1]) + x.slice(e[1]) : x; }],
  ['swapElements', (x, r) => {
    const a = randomElement(x, r), b = randomElement(x, r);
    if (!a || !b || a[0] === b[0] || (a[0] < b[1] && b[0] < a[1])) return x;
    const [p, q] = a[0] < b[0] ? [a, b] : [b, a];
    return x.slice(0, p[0]) + x.slice(q[0], q[1]) + x.slice(p[1], q[0]) + x.slice(p[0], p[1]) + x.slice(q[1]);
  }],
  ['renameClose', (x, r) => {
    const all = [...x.matchAll(/<\/([A-Za-z_][\w:.-]*)>/g)];
    if (!all.length) return x;
    const m = pick(r, all);
    return x.slice(0, m.index!) + `</${pick(r, ['row', 'c', 'v', 'f', 'x', 'sheetData', m[1] + 'X'])}>` + x.slice(m.index! + m[0].length);
  }],
  ['emptyText', (x, r) => {
    const all = [...x.matchAll(/>([^<>]+)</g)];
    if (!all.length) return x;
    const m = pick(r, all);
    return x.slice(0, m.index! + 1) + pick(r, ['', ' ', '\n', '&amp;', '\u0001', '0', '-1', '#REF!', "'", '[', '1+', 'SUM(', 'A1:']) + x.slice(m.index! + 1 + m[1].length);
  }],
];

/** whole-package damages: [name, fn] */
type PartMutator = (parts: Parts, r: Rng) => string;

function xmlPartNames(parts: Parts): string[] {
  return [...parts.keys()].filter((n) => /\.(xml|rels)$/i.test(n) && typeof parts.get(n) === 'string');
}

/** a worksheet / important part is chosen more often than an arbitrary one */
function pickPart(parts: Parts, r: Rng): string {
  const all = xmlPartNames(parts);
  const sheets = all.filter((n) => /^xl\/worksheets\/sheet\d+\.xml$/.test(n));
  const key = all.filter((n) => /^xl\/(workbook|styles|sharedStrings)\.xml$|\.rels$|Content_Types|tables\/|drawings\/|charts\/chart\d+\.xml$/.test(n));
  const x = r();
  if (x < 0.5 && sheets.length) return pick(r, sheets);
  if (x < 0.8 && key.length) return pick(r, key);
  return pick(r, all);
}

const PART_MUTATORS: Array<[string, PartMutator]> = [
  ...MUTATORS.map(([name, fn]): [string, PartMutator] => [name, (parts, r) => {
    const p = pickPart(parts, r);
    parts.set(p, fn(parts.get(p) as string, r));
    return `${name}@${p}`;
  }]),
  ['deletePart', (parts, r) => { const p = pick(r, [...parts.keys()]); parts.delete(p); return `deletePart@${p}`; }],
  ['emptyPart', (parts, r) => { const p = pickPart(parts, r); parts.set(p, ''); return `emptyPart@${p}`; }],
  ['bomPart', (parts, r) => { const p = pickPart(parts, r); parts.set(p, '﻿' + parts.get(p)); return `bomPart@${p}`; }],
  ['crlfPart', (parts, r) => { const p = pickPart(parts, r); parts.set(p, (parts.get(p) as string).replace(/></g, '>\r\n<')); return `crlfPart@${p}`; }],
  ['singleQuotes', (parts, r) => { const p = pickPart(parts, r); parts.set(p, (parts.get(p) as string).replace(/(\s[\w:.-]+)="([^"']*)"/g, "$1='$2'")); return `singleQuotes@${p}`; }],
];

/** zip / byte level damage (applied to the packed bytes) */
function byteDamage(bytes: Uint8Array, r: Rng): { bytes: Uint8Array; what: string } {
  const out = new Uint8Array(bytes);
  const kind = between(r, 0, 2);
  if (kind === 0) {
    const n = between(r, 1, 6);
    for (let i = 0; i < n; i++) out[between(r, 0, out.length - 1)] ^= 1 << between(r, 0, 7);
    return { bytes: out, what: `bitflips(${n})` };
  }
  if (kind === 1) {
    const cut = between(r, 0, out.length - 1);
    return { bytes: out.slice(0, cut), what: `cut@${cut}` };
  }
  const at = between(r, 0, out.length - 1);
  for (let i = at; i < Math.min(out.length, at + 16); i++) out[i] = between(r, 0, 255);
  return { bytes: out, what: `garble@${at}` };
}

/** damage a package: 1..3 part mutations (mostly 1), sometimes a byte level damage on top; returns bytes + description */
export async function mutate(bytes: Uint8Array, seed: number): Promise<{ bytes: Uint8Array; what: string }> {
  const r = rng(seed * 7919 + 13);
  const parts = await unpack(bytes);
  const n = r() < 0.7 ? 1 : between(r, 2, 3);
  const what: string[] = [];
  for (let i = 0; i < n; i++) {
    const [, fn] = pick(r, PART_MUTATORS);
    try { what.push(fn(parts, r)); } catch { what.push('(mutator skipped)'); }
  }
  let out = await pack(parts);
  if (r() < 0.05) {
    const d = byteDamage(out, r);
    out = d.bytes;
    what.push(d.what);
  }
  return { bytes: out, what: what.join(' + ') };
}

// ---------------------------------------------------------------------------------------------------------------
// fixtures and oracles
// ---------------------------------------------------------------------------------------------------------------

const read = (f: string) => new Uint8Array(fs.readFileSync(path.join(FIX, f)));

/** fixtures used for replayLog fuzzing: [file, ops of the golden cases that start from it (no refused ones)] */
export function replayTargets(): Array<{ file: string; logs: LoggedOp[][] }> {
  const by = new Map<string, LoggedOp[][]>();
  for (const c of CASES) {
    if (c.expectError || c.golden) continue;
    if (!by.has(c.fixture)) by.set(c.fixture, []);
    by.get(c.fixture)!.push(c.ops);
  }
  return [...by.entries()].map(([file, logs]) => ({ file, logs }));
}

export interface FuzzOutcome {
  seed: number;
  what: string;
  kind: 'ok' | 'refused' | 'internal' | 'violation';
  detail?: string;
}

/** XML parts that differ from the mutant and must therefore be well-formed */
async function checkOutput(input: Uint8Array, out: Uint8Array): Promise<string | null> {
  const before = await unpack(input).catch(() => null);
  const after = await unpack(out);
  for (const [name, data] of after) {
    if (typeof data !== 'string' || !/\.(xml|rels)$/i.test(name)) continue;
    if (before && before.get(name) === data) continue; // not written by the patcher
    const why = xmlProblem(data.replace(/^﻿/, ''));
    if (why) return `${name}: ${why}`;
  }
  // the reader can read what it could read before
  const readable = await readBook(input, 'in.xlsx').then(() => true, () => false);
  if (readable) {
    try { await readBook(out, 'out.xlsx'); } catch (e) { return `readBook(output) failed: ${(e as Error).message}`; }
  }
  // the result is an acceptable input itself
  try { await replayLog(out, []); } catch (e) { return `the output does not replay: ${(e as Error).message}`; }
  return null;
}

function classify(e: unknown): { kind: 'refused' | 'internal' | 'violation'; detail: string } {
  if (!(e instanceof PatchError)) return { kind: 'violation', detail: `thrown ${e instanceof Error ? `${e.name}: ${e.message}` : String(e)} (not a PatchError)` };
  if (!JP.test(e.message)) return { kind: 'violation', detail: `message is not Japanese: ${e.message}` };
  if (e.internal) return { kind: 'internal', detail: e.message };
  return { kind: 'refused', detail: e.message };
}

/** fuzz replayLog: `count` mutants spread over the golden fixtures; seeds start..start+count-1 */
export async function fuzzReplay(start: number, count: number): Promise<FuzzOutcome[]> {
  const targets = replayTargets();
  const inputs = new Map<string, Uint8Array>(targets.map((t) => [t.file, read(t.file)]));
  const res: FuzzOutcome[] = [];
  for (let seed = start; seed < start + count; seed++) {
    const t = targets[seed % targets.length];
    const r = rng(seed * 31 + 5);
    const log = t.logs[Math.floor(r() * t.logs.length)];
    const m = await mutate(inputs.get(t.file)!, seed);
    const what = `${t.file}: ${m.what}`;
    try {
      const out = await replayLog(m.bytes, log);
      const bad = await checkOutput(m.bytes, out);
      res.push(bad ? { seed, what, kind: 'violation', detail: bad } : { seed, what, kind: 'ok' });
    } catch (e) {
      res.push({ seed, what, ...classify(e) });
    }
  }
  return res;
}

/** random op logs on the intact fixtures: ops are mostly in range, some are not, some are refused by Excel rules */
export function randomLog(r: Rng, sheets: string[]): LoggedOp[] {
  const n = between(r, 1, 8);
  const log: LoggedOp[] = [];
  const num = (max: number) => (r() < 0.008 ? pick(r, [0, -1, 1.5, 1048577, 16385, 99999999]) : between(r, 1, max));
  // the places where the fixtures keep their tables / totals rows / spills / merges: half of the positions are chosen there
  const hotRow = () => (r() < 0.5 ? pick(r, [1, 2, 3, 4, 5, 6, 7, 13, 14, 15]) : num(40));
  const hotCol = () => (r() < 0.5 ? pick(r, [1, 2, 3, 4, 5, 7, 9]) : num(24));
  for (let i = 0; i < n; i++) {
    const sheet = r() < 0.01 ? 'nonexistent' : pick(r, sheets);
    const k = r();
    if (k < 0.45) {
      const axis = r() < 0.6 ? 'col' : 'row';
      const max = axis === 'col' ? 24 : 40;
      const t = pick(r, ['insert', 'delete', 'move'] as const);
      const num = (m: number) => (r() < 0.5 ? between(r, 1, Math.min(m, 8)) : between(r, 1, m)); // low positions hit the tables more often
      const count = r() < 0.7 ? 1 : between(r, 1, 4);
      log.push({ kind: 'struct', sheet, op: t === 'move' ? { type: 'move', axis, from: num(max), count, to: num(max + 1) } : { type: t, axis, at: num(max), count } } as LoggedOp);
    } else if (k < 0.8) {
      const writes = Array.from({ length: between(r, 1, 4) }, () => {
        const row = hotRow(), col = hotCol();
        const kind = r();
        let cell: CellWrite['cell'];
        if (kind < 0.08) cell = null;
        else if (kind < 0.35) cell = { t: 's', v: pick(r, ['x', 'a\nb', '00123', '=text', '', ' lead', 'a<b>&c', 'あい']), ...(r() < 0.3 ? { quote: r() < 0.5 } : {}) };
        else if (kind < 0.6) cell = { t: 'n', v: pick(r, [0, 1, -5, 3.14159, 1e15, 1234567]), ...(r() < 0.3 ? { numFmt: pick(r, ['0%', '#,##0', 'yyyy/m/d', '@', 'General', '0.00"kg"']) } : {}) };
        else if (kind < 0.7) cell = { t: 'd', v: pick(r, [46295, 1, 60.5]), numFmt: pick(r, ['yyyy/m/d', 'h:mm', 'yyyy"年"m"月"']) };
        else if (kind < 0.75) cell = { t: 'b', v: r() < 0.5 };
        else cell = { t: 'z', v: null, f: pick(r, ['SUM(A1:A3)', 'A1*2', 'B2+C3', '$A$1+1', 'IF(A1>0,"y","n")', "'名簿'!A1", '_xlfn.SEQUENCE(3)']), ...(r() < 0.15 ? { dyn: true } : {}) };
        return { row, col, cell };
      });
      log.push({ kind: 'cells', sheet, writes, reason: r() < 0.85 ? 'edit' : 'toValues' });
    } else if (k < 0.92) {
      log.push({ kind: 'colWidth', sheet, col: num(24), width: r() < 0.02 ? pick(r, [300, -1]) : pick(r, [0, 5, 8.43, 20, 255]) });
    } else {
      const a = between(r, 1, 6), b = between(r, 1, 6);
      log.push({ kind: 'unmerge', sheet, refs: [pick(r, ['A1:F1', 'A3:C5', 'E7:G7', 'I1:K1', `A${a}:C${b + a}`])], centerAcross: r() < 0.5 });
    }
  }
  return log;
}

/** worksheet names of a package */
export async function sheetNames(bytes: Uint8Array): Promise<string[]> {
  const parts = await unpack(bytes);
  const wb = parts.get('xl/workbook.xml');
  return typeof wb === 'string' ? [...wb.matchAll(/<sheet\b[^>]*\sname="([^"]*)"/g)].map((m) => m[1].replace(/&amp;/g, '&').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&lt;/g, '<').replace(/&gt;/g, '>')) : [];
}

/** fuzz replayLog with random op logs on INTACT fixtures (refusals are fine; internal errors and damaged outputs are not) */
export async function fuzzOps(start: number, count: number): Promise<FuzzOutcome[]> {
  const targets = replayTargets();
  const inputs = new Map<string, { bytes: Uint8Array; sheets: string[] }>();
  for (const t of targets) { const bytes = read(t.file); inputs.set(t.file, { bytes, sheets: await sheetNames(bytes) }); }
  const res: FuzzOutcome[] = [];
  for (let seed = start; seed < start + count; seed++) {
    const t = targets[seed % targets.length];
    const inp = inputs.get(t.file)!;
    const log = randomLog(rng(seed * 17 + 3), inp.sheets);
    const what = `${t.file}: ${JSON.stringify(log).slice(0, 400)}`;
    try {
      const out = await replayLog(inp.bytes, log);
      const bad = await checkOutput(inp.bytes, out);
      res.push(bad ? { seed, what, kind: 'violation', detail: bad } : { seed, what, kind: 'ok' });
    } catch (e) {
      res.push({ seed, what, ...classify(e) });
    }
  }
  return res;
}

/** fuzz buildPremerged: mutate the base and / or the source package of the standard scenarios */
export async function fuzzPremerge(start: number, count: number): Promise<FuzzOutcome[]> {
  const res: FuzzOutcome[] = [];
  for (let seed = start; seed < start + count; seed++) {
    const r = rng(seed * 131 + 9);
    const input: PremergeInput = scenario(r() < 0.5 ? 'plain' : 'table');
    const which = r() < 0.4 ? 'base' : r() < 0.7 ? 'source' : 'both';
    const what: string[] = [];
    if (which !== 'source') {
      const m = await mutate(input.base.book.bytes, seed * 2);
      input.base = { ...input.base, book: { ...input.base.book, bytes: m.bytes } };
      what.push(`base: ${m.what}`);
    }
    if (which !== 'base') {
      const s = input.sources[0];
      const m = await mutate(s.book.bytes, seed * 2 + 1);
      input.sources = [{ ...s, book: { ...s.book, bytes: m.bytes } }];
      what.push(`source: ${m.what}`);
    }
    const desc = what.join(' | ');
    try {
      const out = await buildPremerged(input);
      const bad = await checkOutput(input.base.book.bytes, out.bytes);
      res.push(bad ? { seed, what: desc, kind: 'violation', detail: bad } : { seed, what: desc, kind: 'ok' });
    } catch (e) {
      res.push({ seed, what: desc, ...classify(e) });
    }
  }
  return res;
}

export function summarize(list: FuzzOutcome[]): { ok: number; refused: number; internal: FuzzOutcome[]; violations: FuzzOutcome[] } {
  return {
    ok: list.filter((x) => x.kind === 'ok').length,
    refused: list.filter((x) => x.kind === 'refused').length,
    internal: list.filter((x) => x.kind === 'internal'),
    violations: list.filter((x) => x.kind === 'violation'),
  };
}

// ---------------------------------------------------------------------------------------------------------------
// random premerge plans on intact packages
// ---------------------------------------------------------------------------------------------------------------

/**
 * A premerge scenario of tests/golden/premerge-scenario.ts with a random plan: a random subset / order of the imported
 * columns (key columns included), random matches, 0..3 appended rows, random options and key fills. The packages are intact,
 * so every plan is valid input: the result must be a correct package (Excel as the oracle, tools/make-ops-outputs.ts).
 */
export function randomPlan(seed: number): PremergeInput {
  const r = rng(seed * 101 + 7);
  const kind = pick(r, ['plain', 'table', 'inside', 'rich', 'keyfill'] as const);
  const input = kind === 'inside' ? insideTableScenario() : kind === 'rich' ? richScenario() : kind === 'keyfill' ? keyFillScenario() : scenario(pick(r, ['plain', 'table'] as const));
  const src = input.sources[0];
  const srcCols = kind === 'rich' ? 3 : 9;
  const srcFirst = src.region.firstDataRow, srcLast = src.region.lastDataRow;
  const baseRows = input.rows.filter((x) => x.baseRow != null).map((x) => x.baseRow!);
  const pickRow = () => between(r, srcFirst, srcLast);
  const plan = baseRows.map((baseRow) => ({ baseRow, matches: { [src.id]: r() < 0.75 ? { row: pickRow(), via: 'ID', flags: [] } : null } }));
  const appended = Array.from({ length: between(r, 0, 3) }, () => {
    const row = pickRow();
    return { baseRow: null, appendedFrom: { sourceId: src.id, row }, matches: { [src.id]: { row, via: 'ID', flags: [] } } };
  });
  const cols = Array.from({ length: srcCols }, (_x, i) => i + 1).filter(() => r() < 0.6);
  if (!cols.length) cols.push(between(r, 1, srcCols));
  for (let i = cols.length - 1; i > 0; i--) { const j = Math.floor(r() * (i + 1)); [cols[i], cols[j]] = [cols[j], cols[i]]; } // any order
  const keyFill = appended.length ? { [src.id]: [{ baseCol: 1, sourceCol: 1 }, ...(r() < 0.5 ? [{ baseCol: 2, sourceCol: between(r, 1, srcCols) }] : [])] } : {};
  return {
    ...input,
    rows: [...plan, ...appended] as PremergeInput['rows'],
    imports: cols.map((col) => ({ sourceId: src.id, col })),
    keyFill,
    options: { formulaPolicy: pick(r, ['keep-if-possible', 'values'] as const), extendTable: r() < 0.8, headerConflict: pick(r, ['suffix', 'keep'] as const) },
  };
}

/** fuzz buildPremerged with random plans, then replay a random op log on the result (the app's real flow) */
export async function fuzzPlans(start: number, count: number): Promise<FuzzOutcome[]> {
  const res: FuzzOutcome[] = [];
  for (let seed = start; seed < start + count; seed++) {
    const input = randomPlan(seed);
    const what = `plan: ${input.base.label} imports=[${input.imports.map((i) => i.col).join(',')}] rows=${input.rows.length} options=${JSON.stringify(input.options)}`;
    try {
      const pre = await buildPremerged(input);
      let bad = await checkOutput(input.base.book.bytes, pre.bytes);
      if (!bad) {
        const names = await sheetNames(pre.bytes);
        const log = randomLog(rng(seed * 43 + 1), names);
        try {
          const out = await replayLog(pre.bytes, log);
          bad = await checkOutput(pre.bytes, out);
        } catch (e) {
          const c = classify(e);
          if (c.kind !== 'refused') { res.push({ seed, what: `${what} then ${JSON.stringify(log).slice(0, 300)}`, ...c }); continue; }
        }
      }
      res.push(bad ? { seed, what, kind: 'violation', detail: bad } : { seed, what, kind: 'ok' });
    } catch (e) {
      res.push({ seed, what, ...classify(e) });
    }
  }
  return res;
}
