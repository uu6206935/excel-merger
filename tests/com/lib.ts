/** Helpers for the real-Excel tests (EXCEL_TESTS=1). Every script uses its OWN hidden Excel; only that PID is ever killed. */
import { spawn } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { POWERSHELL, killExcelPid } from '../../server/excel/run-job.mjs';
import { numToCol } from '../../src/core/formula/excelRef';

export const EXCEL_TESTS = process.env.EXCEL_TESTS === '1';
export const PS_DIR = path.join(__dirname, 'ps');

export interface PsRun {
  code: number | null;
  lines: Array<Record<string, any>>;
  stderr: string;
  timedOut: boolean;
}

/** Runs a test .ps1 (ASCII-only files) with a watchdog that kills only the script's own Excel PID. */
export function runPs(script: string, args: string[], timeoutMs = 120_000): Promise<PsRun> {
  return new Promise((resolve, reject) => {
    const child = spawn(POWERSHELL, ['-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', path.join(PS_DIR, script), ...args], {
      windowsHide: true,
      stdio: ['ignore', 'pipe', 'pipe'],
    });
    const lines: Array<Record<string, any>> = [];
    let buf = '';
    let err = '';
    let pid = 0;
    let timedOut = false;
    child.stdout.setEncoding('utf8');
    child.stdout.on('data', (d: string) => {
      buf += d;
      let nl;
      while ((nl = buf.indexOf('\n')) >= 0) {
        const line = buf.slice(0, nl).trim();
        buf = buf.slice(nl + 1);
        if (!line) continue;
        try {
          const o = JSON.parse(line);
          if (o.type === 'excel') pid = o.pid;
          lines.push(o);
        } catch {
          lines.push({ type: 'log', text: line });
        }
      }
    });
    child.stderr.on('data', (d: Buffer) => (err += d.toString('latin1')));
    const timer = setTimeout(async () => {
      timedOut = true;
      if (pid) await killExcelPid(pid);
      child.kill();
    }, timeoutMs);
    child.on('error', reject);
    child.on('close', (code) => {
      clearTimeout(timer);
      resolve({ code, lines, stderr: err, timedOut });
    });
  });
}

/** 'f' formula (Range.Formula) · 'd' dynamic-array formula (Formula2) · 'a' legacy array formula over the range · 's' · 'n' */
export type SpecCell = [string, 'f' | 's' | 'n' | 'd' | 'a', string | number];
export interface FixtureSheet {
  name: string;
  cells: SpecCell[];
  merges?: string[];
  tables?: Array<{ name: string; range: string }>;
  widths?: Array<[number, number]>;
  /** columns to hide (1-based) */
  hidden?: number[];
  /** range that gets an AutoFilter (filter arrows, nothing filtered) */
  autoFilter?: string;
  formats?: Array<[string, string]>;
  align?: Array<[string, 'center' | 'left' | 'right' | 'centerAcross']>;
  cf?: Array<{ range: string; type: 'expression' | 'cellIs' | 'colorScale' | 'dataBar'; formula?: string; operator?: number }>;
  dv?: Array<{ range: string; type: 'list' | 'whole'; formula: string; formula2?: string }>;
}
export interface FixtureSpec {
  sheets: FixtureSheet[];
  names?: Array<{ name: string; refersTo: string }>;
}

export async function makeFixture(spec: FixtureSpec, out: string): Promise<void> {
  const specFile = `${out}.spec.json`;
  fs.writeFileSync(specFile, JSON.stringify(spec), 'utf8');
  // Excel is shared with the user's own work and other programs: a transient COM failure is retried once or twice
  let last = '';
  for (let attempt = 1; attempt <= 3; attempt++) {
    const r = await runPs('make-fixture.ps1', ['-Spec', specFile, '-Out', out]);
    const res = r.lines.find((l) => l.type === 'result' || l.type === 'error');
    if (res && res.ok && fs.existsSync(out)) return;
    last = `${JSON.stringify(r.lines)} ${r.stderr}`;
  }
  throw new Error(`make-fixture failed 3 times: ${last}`);
}

export interface DumpCell {
  r: number;
  c: number;
  hf: boolean;
  /** Value2; errors as 'ERR:#N/A' */
  v: unknown;
  t: string;
  /** Range.Formula (legacy view) */
  f?: string;
  /** Range.Formula2 */
  f2?: string;
  /** NumberFormat */
  nf?: string;
  /** HorizontalAlignment (1 general, -4108 center, 7 center across selection, ...) */
  ha?: number;
  /** CurrentArray address of a legacy array cell */
  arr?: string;
  /** SpillingToRange of a spill anchor */
  spill?: string;
}
export interface DumpSheet {
  name: string;
  cells: DumpCell[];
  widths: number[];
  tables: Array<{ name: string; range: string; columns: string[] }>;
  merges: string[];
}
export interface Dump {
  sheets: DumpSheet[];
  names: Array<{ name: string; refersTo: string }>;
}

export async function dumpBook(file: string): Promise<Dump> {
  const out = `${file}.dump.json`;
  let last = '';
  let ok = false;
  for (let attempt = 1; attempt <= 3 && !ok; attempt++) {
    const r = await runPs('dump.ps1', ['-Path', file, '-Out', out]);
    const res = r.lines.find((l) => l.type === 'result' || l.type === 'error');
    if (res && res.ok) ok = true;
    else last = `${JSON.stringify(r.lines)} ${r.stderr}`;
  }
  if (!ok) throw new Error(`dump failed 3 times: ${last}`);
  const d = JSON.parse(fs.readFileSync(out, 'utf8').replace(/^﻿/, '')) as Dump;
  // ConvertTo-Json turns 1-element lists into scalars in some PS versions: normalise
  d.sheets = ([] as DumpSheet[]).concat((d.sheets as unknown as DumpSheet[]) ?? []);
  for (const s of d.sheets) {
    s.cells = ([] as DumpCell[]).concat((s.cells as unknown as DumpCell[]) ?? []);
    s.tables = ([] as DumpSheet['tables']).concat(s.tables ?? []);
    for (const t of s.tables) t.columns = ([] as string[]).concat(t.columns ?? []);
    s.merges = ([] as string[]).concat(s.merges ?? []);
  }
  d.names = ([] as Dump['names']).concat(d.names ?? []);
  return d;
}

export async function clipboardText(): Promise<{ text: string | null; seq: number }> {
  const r = await runPs('clip-text.ps1', [], 30_000);
  const l = r.lines.find((x) => x.type === 'clip');
  if (!l) throw new Error(`clip-text failed: ${r.stderr}`);
  return { text: l.text ?? null, seq: l.seq };
}

/**
 * Every Excel start (even a hidden automation instance) writes ~80 MB of Office diagnostic logs into
 * %TEMP%\Diagnostics\EXCEL and never removes them (measured: 4 files x 20 MB per start, 40 GB after one day of tests).
 * A full C: drive makes Excel fail in confusing ways, so the real-Excel tests stop early with a clear message.
 */
export function ensureFreeDisk(minBytes = 1.5 * 1024 ** 3): void {
  let free = Number.POSITIVE_INFINITY;
  try {
    const s = fs.statfsSync(os.tmpdir());
    free = Number(s.bavail) * Number(s.bsize);
  } catch {
    return; // cannot tell: go on
  }
  if (free < minBytes) {
    throw new Error(
      `Only ${(free / 1048576).toFixed(0)} MB free on the drive of ${os.tmpdir()}: the real-Excel tests need at least ${(minBytes / 1048576).toFixed(0)} MB. ` +
        `Excel writes ~80 MB of diagnostic logs per start into ${path.join(os.tmpdir(), 'Diagnostics', 'EXCEL')} and never deletes them: delete the old files there (Excel recreates them) and run again.`,
    );
  }
}

export function tempDir(label: string): string {
  ensureFreeDisk();
  // ASCII temp path, like the server's job folders
  return fs.mkdtempSync(path.join(os.tmpdir(), `em-com-${label}-`));
}

export const cellAt = (d: Dump, sheet: string, r: number, c: number) =>
  d.sheets.find((s) => s.name === sheet)?.cells.find((x) => x.r === r && x.c === c);

/** Cells covered by an sqref ("C2:D13 C14:D22" == "C2:D22"); a huge area (whole columns) stays as its text. */
export function coverage(sqref: string): Set<string> {
  const out = new Set<string>();
  const col = (s: string) => [...s.toUpperCase()].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
  for (const tok of sqref.split(/\s+/).filter(Boolean)) {
    const m = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})(?::\$?([A-Za-z]{1,3})\$?(\d{1,7}))?$/.exec(tok);
    if (!m) {
      out.add(tok);
      continue;
    }
    const [c1, r1, c2, r2] = [col(m[1]), Number(m[2]), m[3] ? col(m[3]) : col(m[1]), m[4] ? Number(m[4]) : Number(m[2])];
    if ((Math.abs(r2 - r1) + 1) * (Math.abs(c2 - c1) + 1) > 50_000) {
      out.add(tok);
      continue;
    }
    for (let r = Math.min(r1, r2); r <= Math.max(r1, r2); r++) for (let c = Math.min(c1, c2); c <= Math.max(c1, c2); c++) out.add(`${r},${c}`);
  }
  return out;
}

/**
 * CF / DV rules as { rule signature (kind, type, formulas in file form) -> covered cells }. Excel splits a rule's area into
 * several ranges after cut / insert ("C14:D22 C2:D13") and renumbers / reorders priorities; the same rule over the same
 * cells is the same rule, so rules are compared as (signature, coverage) sets, not as text in file order.
 */
export function ruleCoverage(rules: Array<{ kind: string; ext?: boolean; type?: string; sqref: string; formulas: string[] }>): Map<string, Set<string>> {
  const out = new Map<string, Set<string>>();
  for (const r of rules) {
    const sig = `${r.kind}${r.ext ? '(ext)' : ''} ${r.type ?? ''} ${r.formulas.join(' ; ')}`;
    const set = out.get(sig) ?? new Set<string>();
    for (const c of coverage(r.sqref)) set.add(c);
    out.set(sig, set);
  }
  return out;
}

/** Differences of two rule sets (see ruleCoverage); `what` prefixes every line. */
export function diffRuleSets(what: string, a: Map<string, Set<string>>, b: Map<string, Set<string>>): string[] {
  const out: string[] = [];
  const show = (s: Set<string>) => `${s.size} cells`;
  for (const [sig, ca] of a) {
    const cb = b.get(sig);
    if (!cb) out.push(`${what} rule only in A: ${sig} (${show(ca)})`);
    else {
      const onlyA = [...ca].filter((x) => !cb.has(x));
      const onlyB = [...cb].filter((x) => !ca.has(x));
      if (onlyA.length || onlyB.length) out.push(`${what} rule ${sig}: covers ${onlyA.length} cells only in A (${onlyA.slice(0, 6).join(' ')}), ${onlyB.length} only in B (${onlyB.slice(0, 6).join(' ')})`);
    }
  }
  for (const [sig, cb] of b) if (!a.has(sig)) out.push(`${what} rule only in B: ${sig} (${show(cb)})`);
  return out;
}

/**
 * File-level comparison of two saved workbooks (what the packages contain, read with the project's reader):
 * conditional formats and data validations (kind, type, covered cells, formulas in file form; order and range
 * fragmentation ignored), merged areas, tables (ref + column names), autofilter and the STORED <col width> values of the
 * first `cols` columns (1/256 tolerance).
 */
export async function packageDiff(a: Uint8Array, b: Uint8Array, cols = 20): Promise<string[]> {
  const { readBook } = await import('../../src/core/xlsx/read');
  const [ba, bb] = await Promise.all([readBook(a, 'a.xlsx'), readBook(b, 'b.xlsx')]);
  const out: string[] = [];
  const box = (m: { r1: number; c1: number; r2: number; c2: number }) => `${m.r1},${m.c1}:${m.r2},${m.c2}`;
  for (const sa of ba.sheets) {
    const sb = bb.sheets.find((s) => s.name === sa.name);
    if (!sb) {
      out.push(`sheet ${sa.name} missing in B`);
      continue;
    }
    out.push(...diffRuleSets(sa.name, ruleCoverage(sa.features.rules), ruleCoverage(sb.features.rules)));
    const ma = sa.features.merges.map(box).sort();
    const mb = sb.features.merges.map(box).sort();
    if (JSON.stringify(ma) !== JSON.stringify(mb)) out.push(`${sa.name} merges: ${JSON.stringify(ma)} | ${JSON.stringify(mb)}`);
    const tables = (s: typeof sa) => s.features.tables.map((t) => `${t.name} ${t.ref} ${t.columns.join('|')}`).sort();
    if (JSON.stringify(tables(sa)) !== JSON.stringify(tables(sb))) out.push(`${sa.name} tables: ${JSON.stringify(tables(sa))} | ${JSON.stringify(tables(sb))}`);
    if ((sa.features.autoFilter ?? '') !== (sb.features.autoFilter ?? '')) out.push(`${sa.name} autoFilter: ${sa.features.autoFilter} | ${sb.features.autoFilter}`);
    for (let c = 1; c <= cols; c++) {
      const w1 = sa.colProps[c - 1]?.width ?? sa.defaultColWidth;
      const w2 = sb.colProps[c - 1]?.width ?? sb.defaultColWidth;
      if (w1 !== undefined && w2 !== undefined && Math.abs(w1 - w2) > 1 / 256) out.push(`${sa.name} stored width col ${c}: ${w1} | ${w2}`);
    }
  }
  const names = (bk: typeof ba) => bk.names.map((n) => `${n.scope ?? ''}!${n.name}=${n.refersTo}`).sort();
  const na = names(ba);
  const nb = names(bb);
  for (const n of na) if (!nb.includes(n)) out.push(`name only in A: ${n}`);
  for (const n of nb) if (!na.includes(n)) out.push(`name only in B: ${n}`);
  return out;
}

/** Per-cell comparison of two dumps (what Excel reports for each cell). */
export function diffDumps(a: Dump, b: Dump): string[] {
  const out: string[] = [];
  const keys: Array<keyof DumpCell> = ['hf', 'f', 'f2', 'v', 'nf', 'ha', 'arr', 'spill'];
  for (const sa of a.sheets) {
    const sb = b.sheets.find((s) => s.name === sa.name);
    if (!sb) {
      out.push(`sheet ${sa.name} missing`);
      continue;
    }
    const at = new Map(sb.cells.map((c) => [`${c.r},${c.c}`, c]));
    const seen = new Set<string>();
    for (const ca of sa.cells) {
      const k = `${ca.r},${ca.c}`;
      seen.add(k);
      const cb = at.get(k);
      const addr = `${sa.name}!${numToCol(ca.c)}${ca.r}`;
      if (!cb) {
        out.push(`${addr}: only in A ${JSON.stringify(ca)}`);
        continue;
      }
      for (const key of keys) {
        const x = ca[key];
        const y = cb[key];
        const same = typeof x === 'number' && typeof y === 'number' ? Math.abs(x - y) < 1e-9 : JSON.stringify(x) === JSON.stringify(y);
        if (!same) out.push(`${addr} ${key}: ${JSON.stringify(x)} | ${JSON.stringify(y)}`);
      }
    }
    for (const cb of sb.cells) if (!seen.has(`${cb.r},${cb.c}`)) out.push(`${sa.name}!${numToCol(cb.c)}${cb.r}: only in B ${JSON.stringify(cb)}`);
    if (JSON.stringify(sa.merges) !== JSON.stringify(sb.merges)) out.push(`${sa.name} merges: ${JSON.stringify(sa.merges)} | ${JSON.stringify(sb.merges)}`);
    for (let i = 0; i < 30; i++) if (Math.abs(sa.widths[i] - sb.widths[i]) > 1e-6) out.push(`${sa.name} width ${numToCol(i + 1)}: ${sa.widths[i]} | ${sb.widths[i]}`);
    if (JSON.stringify(sa.tables) !== JSON.stringify(sb.tables)) out.push(`${sa.name} tables: ${JSON.stringify(sa.tables)} | ${JSON.stringify(sb.tables)}`);
  }
  const na = a.names.map((n) => `${n.name}=${n.refersTo}`).sort();
  const nb = b.names.map((n) => `${n.name}=${n.refersTo}`).sort();
  if (JSON.stringify(na) !== JSON.stringify(nb)) out.push(`names: ${JSON.stringify(na)} | ${JSON.stringify(nb)}`);
  return out;
}

// ---------------------------------------------------------------------------------------------------------
// Raw package cells (no Excel): what the FILE says about a cell (xf attributes), for things Excel's object model hides
// (quotePrefix, the exact fill / border records)
// ---------------------------------------------------------------------------------------------------------

export interface RawCell {
  addr: string;
  /** resolved text of a string cell (shared / inline), the raw <v> of any other cell */
  value?: string;
  /** cell type attribute (s, n, b, e, str ...) */
  t?: string;
  formula?: string;
  quotePrefix: boolean;
  numFmtId: number;
  /** format code of a custom format, 'General' for id 0, '@' for 49, 'yyyy/m/d' for 14 (ja-JP built-in), else undefined */
  numFmt?: string;
  horizontal?: string;
  /** e.g. 'solid rgb=FF1F4E78' / 'solid theme=5 tint=0.39997558519241921' / 'none' */
  fill: string;
  /** per side: 'thin rgb=FFC00000' / 'thin auto' / undefined (no line) */
  border: { left?: string; right?: string; top?: string; bottom?: string };
}

/**
 * The lines that are DRAWN, independent of which of the two neighbouring cells stores them: Excel keeps a line shared by two
 * adjacent cells on one of them only (A1.right thin is saved as B1.left thin), the JS patcher leaves the records as they are.
 * Key 'V:row:col' = the vertical line at the left of cell (row, col), 'H:row:col' = the horizontal line at its top.
 */
export function drawnEdges(cells: Map<string, RawCell>): Map<string, string> {
  const out = new Map<string, string>();
  const put = (key: string, style: string | undefined) => {
    if (!style) return;
    const old = out.get(key);
    out.set(key, old && old !== style ? [old, style].sort().join(' | ') : style);
  };
  for (const c of cells.values()) {
    const m = /^([A-Z]+)(\d+)$/.exec(c.addr);
    if (!m) continue;
    const col = [...m[1]].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
    const row = Number(m[2]);
    put(`V:${row}:${col}`, c.border.left);
    put(`V:${row}:${col + 1}`, c.border.right);
    put(`H:${row}:${col}`, c.border.top);
    put(`H:${row + 1}:${col}`, c.border.bottom);
  }
  return out;
}

/**
 * Per-cell differences of two packages' sheets as the user sees them: value, type, quotePrefix, number format, fill, horizontal
 * alignment and the drawn border lines (see drawnEdges). A missing cell equals a default cell.
 */
export function diffRawCells(a: Map<string, RawCell>, b: Map<string, RawCell>, labels: [string, string] = ['A', 'B']): string[] {
  const out: string[] = [];
  const blank = (addr: string): RawCell => ({ addr, quotePrefix: false, numFmtId: 0, numFmt: 'General', fill: 'none', border: {} });
  const plain = (t: string | undefined) => (t ?? '').replace(/_x000D_/g, '').replace(/\r/g, '');
  const norm = (c: RawCell) => ({
    value: c.value === undefined ? undefined : plain(c.value),
    t: c.value === undefined ? undefined : c.t === 'str' || c.t === 'inlineStr' ? 's' : c.t ?? 'n',
    formula: c.formula,
    quotePrefix: c.quotePrefix,
    numFmt: c.numFmt ?? `#${c.numFmtId}`,
    fill: c.fill,
    horizontal: c.horizontal === 'general' ? undefined : c.horizontal,
  });
  const addrs = [...new Set([...a.keys(), ...b.keys()])].sort();
  for (const addr of addrs) {
    const x = norm(a.get(addr) ?? blank(addr));
    const y = norm(b.get(addr) ?? blank(addr));
    for (const k of Object.keys(x) as Array<keyof typeof x>) {
      if (JSON.stringify(x[k]) !== JSON.stringify(y[k])) out.push(`${addr} ${k}: ${labels[0]}=${JSON.stringify(x[k])} ${labels[1]}=${JSON.stringify(y[k])}`);
    }
  }
  const ea = drawnEdges(a);
  const eb = drawnEdges(b);
  for (const k of [...new Set([...ea.keys(), ...eb.keys()])].sort()) {
    if (ea.get(k) !== eb.get(k)) out.push(`line ${k}: ${labels[0]}=${ea.get(k) ?? '-'} ${labels[1]}=${eb.get(k) ?? '-'}`);
  }
  return out;
}

/** Cells of one sheet of a package, straight from the XML. */
export async function rawCells(bytes: Uint8Array, sheetName: string): Promise<Map<string, RawCell>> {
  const JSZip = (await import('jszip')).default;
  const { parseXml } = await import('../../tools/lib/package-diff.mjs');
  type N = NonNullable<ReturnType<typeof parseXml>>;
  const zip = await JSZip.loadAsync(bytes);
  const xml = async (name: string): Promise<N | null> => {
    const f = zip.file(name);
    return f ? parseXml(await f.async('string')) : null;
  };
  const kids = (n: N | null | undefined, local: string): N[] => (n ? n.children.filter((c) => c.local === local) : []);
  const textOf = (n: N): string => n.text + n.children.map(textOf).join('');
  const wb = await xml('xl/workbook.xml');
  const rels = await xml('xl/_rels/workbook.xml.rels');
  const sheetEl = kids(kids(wb, 'sheets')[0], 'sheet').find((x) => x.attrs.name === sheetName);
  if (!sheetEl) throw new Error(`sheet ${sheetName} not found`);
  const rid = Object.entries(sheetEl.attrs).find(([k]) => k === 'r:id' || k.endsWith(':id'))?.[1];
  const target = kids(rels, 'Relationship').find((r) => r.attrs.Id === rid)?.attrs.Target ?? '';
  const part = target.startsWith('/') ? target.slice(1) : `xl/${target}`;
  const sheet = await xml(part);
  const sst = kids(await xml('xl/sharedStrings.xml'), 'si').map(textOf);
  const styles = await xml('xl/styles.xml');
  const customFmts = new Map(kids(kids(styles, 'numFmts')[0], 'numFmt').map((f) => [Number(f.attrs.numFmtId), f.attrs.formatCode]));
  const fills = kids(kids(styles, 'fills')[0], 'fill');
  const borders = kids(kids(styles, 'borders')[0], 'border');
  const xfs = kids(kids(styles, 'cellXfs')[0], 'xf');
  const colorText = (c: N | undefined) => (c ? (c.attrs.rgb ? `rgb=${c.attrs.rgb}` : c.attrs.theme ? `theme=${c.attrs.theme}${c.attrs.tint ? ` tint=${c.attrs.tint}` : ''}` : c.attrs.indexed === '64' || c.attrs.auto ? 'auto' : c.attrs.indexed ? `indexed=${c.attrs.indexed}` : '') : '');
  const fillText = (i: number): string => {
    const pf = kids(fills[i], 'patternFill')[0];
    if (!pf) return 'none';
    const t = pf.attrs.patternType ?? 'none';
    return t === 'none' ? 'none' : `${t} ${colorText(kids(pf, 'fgColor')[0])}`.trim();
  };
  const sideText = (b: N | undefined, side: string): string | undefined => {
    const e = kids(b, side)[0];
    return e && e.attrs.style ? `${e.attrs.style} ${colorText(kids(e, 'color')[0])}`.trim() : undefined;
  };
  const out = new Map<string, RawCell>();
  for (const row of kids(kids(sheet, 'sheetData')[0], 'row')) {
    for (const c of kids(row, 'c')) {
      const xf = xfs[Number(c.attrs.s ?? 0)];
      const id = Number(xf?.attrs.numFmtId ?? 0);
      const t = c.attrs.t;
      const v = kids(c, 'v')[0];
      let value: string | undefined;
      if (t === 's' && v) value = sst[Number(textOf(v))];
      else if (t === 'inlineStr') value = kids(c, 'is').map(textOf).join('');
      else if (v) value = textOf(v);
      const f = kids(c, 'f')[0];
      const b = borders[Number(xf?.attrs.borderId ?? 0)];
      out.set(c.attrs.r, {
        addr: c.attrs.r,
        value,
        t,
        formula: f ? textOf(f) : undefined,
        quotePrefix: xf?.attrs.quotePrefix === '1',
        numFmtId: id,
        numFmt: id === 0 ? 'General' : id === 49 ? '@' : id === 14 ? 'yyyy/m/d' : customFmts.get(id),
        horizontal: kids(xf, 'alignment')[0]?.attrs.horizontal,
        fill: fillText(Number(xf?.attrs.fillId ?? 0)),
        border: { left: sideText(b, 'left'), right: sideText(b, 'right'), top: sideText(b, 'top'), bottom: sideText(b, 'bottom') },
      });
    }
  }
  return out;
}
