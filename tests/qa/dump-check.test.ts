// tools/lib/dump-check.mjs: the cross-check of an Excel dump against its package (the fix for the check-output flake: a good output
// flagged FAIL because Excel answered while busy — no defined names, headers unread, formulas "pointing at other headers").
//   npx vitest run tests/qa/dump-check.test.ts                      pure: synthetic dumps built from the real sample packages
//   QA_EXCEL=1 npx vitest run tests/qa/dump-check.test.ts           + real Excel: consistent dumps, self-check, the retry path
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { afterAll, describe, expect, it } from 'vitest';
import { dumpDifferences, dumpIssues, packageFacts, refSignature } from '../../tools/lib/dump-check.mjs';
import { excelAvailable, inspectFiles } from '../../tools/lib/inspect.mjs';
import type { InspectResult } from '../../tools/lib/inspect.mjs';
import { loadPackageModel } from '../../tools/lib/package-diff.mjs';
import { checkOutput } from '../../tools/lib/verdict.mjs';
import { sheet } from './helpers';

const root = path.resolve(__dirname, '../..');
const base = path.join(root, 'samples', '01_学生名簿.xlsx');
const interview = path.join(root, 'samples', '02_面接評価.xlsx');
const haveSamples = fs.existsSync(base) && fs.existsSync(interview);

/** What Excel would report for the package, built from the package model itself (a stand-in for inspect.ps1). */
async function fakeDump(file: string): Promise<InspectResult> {
  const model = await loadPackageModel(file);
  const sheets = [...model.sheetModels.values()].map((m, i) => {
    const formulas: Record<string, string> = {};
    const values: Record<string, string | number | boolean> = {};
    for (const [addr, c] of m.cells as Map<string, { f?: string; value: unknown; type: string }>) {
      if (typeof c.f === 'string' && c.f) formulas[addr] = `=${c.f}`;
      else if (c.type === 's' || c.type === 'n' || c.type === 'b') values[addr] = c.value as string | number | boolean;
    }
    return {
      name: m.name,
      index: i + 1,
      visible: -1,
      capped: false,
      formulas,
      values,
      cf: (m.cf as string[]).map(() => ({ type: 1, appliesTo: 'A1', operator: null, formula1: null, formula2: null, priority: 1 })),
      dv: (m.dv as string[]).map(() => ({ range: 'A1', type: 3, operator: null, formula1: null, formula2: null, alertStyle: 1 })),
      tables: (m.tables as string[]).map((t) => ({ name: t, range: 'A1', showTotals: false, columns: [] })),
      charts: (m.drawings as string[]).filter((d) => /[A-Za-z0-9]Chart[:[]/.test(d)).map((d) => ({ name: d, type: 51, title: null, topLeft: 'A1', series: [] })),
      pivots: (m.pivots as Array<{ name: string }>).map((p) => ({ name: p.name, range: null, source: null, rowFields: [], columnFields: [], dataFields: [] })),
    };
  });
  const names = model.names
    .filter((n: { name: string }) => !n.name.startsWith('_xl'))
    .map((n: { name: string; scope: string | null; text: string }) => ({ name: n.name, scope: n.scope, refersTo: `=${n.text}`, visible: true }));
  return { tool: 'excel-merger/inspect', version: 1, file: path.basename(file), path: file, openMode: 'normal', names, sheets } as unknown as InspectResult;
}

describe('refSignature', () => {
  it('ignores $, sorts, keeps the sheet of every reference', () => {
    expect(refSignature('=SUM($A$1:B2)+Sheet2!C3', '名簿')).toBe('Sheet2!3,3,3,3;名簿!1,1,2,2');
    expect(refSignature('=SUM(A1:B2)+Sheet2!C3', '名簿')).toBe(refSignature('=SUM($A$1:$B$2)+Sheet2!$C$3', '名簿'));
    expect(refSignature('="A1"&"B2"', '名簿')).toBe('');
    expect(refSignature('=A1', '名簿')).not.toBe(refSignature('=B1', '名簿'));
  });
});

describe.skipIf(!haveSamples)('dumpIssues against the real sample packages', () => {
  it('reads the facts of 01_学生名簿', async () => {
    const f = await packageFacts(base);
    expect(f).not.toBeNull();
    expect(f!.names.map((n) => n.name).sort()).toEqual(['基準日', '志望度列', '大学一覧'].sort());
    const roster = f!.sheets.find((s) => s.name === '名簿')!;
    expect(roster.formulas.size).toBeGreaterThan(100);
    expect(roster.texts.size).toBeGreaterThan(20);
    expect(roster.cf).toBeGreaterThanOrEqual(2);
    expect(roster.dv).toBeGreaterThanOrEqual(2);
    expect(f!.sheets.find((s) => s.name === '集計')!.charts).toBe(1);
    expect(f!.sheets.find((s) => s.name === 'ピボット')!.pivots).toBe(1);
  });

  it('is not applicable to xls / csv / missing / broken files', async () => {
    expect(await packageFacts(path.join(root, 'samples', '04_適性検査.xls'))).toBeNull();
    expect(await packageFacts(path.join(root, 'samples', '03_連絡先.csv'))).toBeNull();
    expect(await packageFacts(path.join(root, 'samples', 'no-such-file.xlsx'))).toBeNull();
    const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'em-dc-'));
    try {
      const bad = path.join(tmp, 'broken.xlsx');
      fs.writeFileSync(bad, Buffer.from('not a zip'));
      expect(await packageFacts(bad)).toBeNull();
    } finally {
      fs.rmSync(tmp, { recursive: true, force: true });
    }
    expect(dumpIssues(null, await fakeDump(base))).toEqual([]);
  });

  it.each([
    ['01_学生名簿.xlsx', base],
    ['02_面接評価.xlsx', interview],
  ])('a dump that matches the package is consistent (%s)', async (_n, file) => {
    const facts = await packageFacts(file);
    expect(dumpIssues(facts, await fakeDump(file))).toEqual([]);
  });

  it('flags the flake signature: no defined names', async () => {
    const d = await fakeDump(base);
    d.names = [];
    expect(dumpIssues(await packageFacts(base), d).join('\n')).toMatch(/名前の定義/);
    d.names = (await fakeDump(base)).names!.filter((n) => n.name !== '大学一覧');
    expect(dumpIssues(await packageFacts(base), d).join('\n')).toMatch(/大学一覧/);
  });

  it('flags the flake signature: cell values (header row) not read', async () => {
    const d = await fakeDump(base);
    for (const s of d.sheets!) s.values = {};
    const msg = dumpIssues(await packageFacts(base), d).join('\n');
    expect(msg).toMatch(/文字 \d+ 件のうち \d+ 件が Excel の読み取りにありません/);
  });

  it('flags formulas that point at other columns / are missing', async () => {
    const facts = await packageFacts(base);
    const d = await fakeDump(base);
    const roster = d.sheets!.find((s) => s.name === '名簿')!;
    // every reference one column to the right: what "formulas pointing at other headers" looked like
    for (const a of Object.keys(roster.formulas)) roster.formulas[a] = roster.formulas[a].replace(/\b([A-K])(\d+)/g, (_m, c: string, r: string) => `${String.fromCharCode(c.charCodeAt(0) + 1)}${r}`);
    expect(dumpIssues(facts, d).join('\n')).toMatch(/名簿」の数式 \d+ 件のうち \d+ 件/);
    const d2 = await fakeDump(base);
    d2.sheets!.find((s) => s.name === '集計')!.formulas = {};
    expect(dumpIssues(facts, d2).join('\n')).toMatch(/集計」の数式/);
  });

  it('flags missing CF / DV / chart / pivot and unknown sheets', async () => {
    const facts = await packageFacts(base);
    const d = await fakeDump(base);
    d.sheets!.find((s) => s.name === '名簿')!.cf = [];
    d.sheets!.find((s) => s.name === '名簿')!.dv = [];
    d.sheets!.find((s) => s.name === '集計')!.charts = [];
    d.sheets!.find((s) => s.name === 'ピボット')!.pivots = [];
    const msg = dumpIssues(facts, d).join('\n');
    for (const w of ['条件付き書式', '入力規則', 'グラフ', 'ピボットテーブル']) expect(msg).toContain(w);
    const d2 = await fakeDump(base);
    d2.sheets = d2.sheets!.filter((s) => s.name !== 'マスタ');
    expect(dumpIssues(facts, d2).join('\n')).toMatch(/シート「マスタ」/);
  });

  it('flags the signs the PowerShell side reports itself (busy reads, wrong workbook, short name table)', async () => {
    const facts = await packageFacts(base);
    for (const [k, v, re] of [
      ['transientFailures', 2, /busy/],
      ['wrongWorkbook', 'C:\\x\\other.xlsx', /ファイルではありません/],
      ['nameCountMismatch', '2/3', /名前の定義の数/],
    ] as const) {
      const d = await fakeDump(base);
      (d as unknown as Record<string, unknown>)[k] = v;
      expect(dumpIssues(facts, d).join('\n')).toMatch(re);
    }
  });

  it('does not judge repaired / capped dumps', async () => {
    const facts = await packageFacts(base);
    const d = await fakeDump(base);
    d.names = [];
    d.openMode = 'repair';
    expect(dumpIssues(facts, d)).toEqual([]);
    const c = await fakeDump(base);
    for (const s of c.sheets!) {
      s.capped = true;
      s.values = {};
      s.formulas = {};
    }
    expect(dumpIssues(facts, c).filter((x) => /数式|文字/.test(x))).toEqual([]);
  });

  it('dumpDifferences lists what differs between two readings', async () => {
    const a = await fakeDump(base);
    const b = await fakeDump(base);
    expect(dumpDifferences(a, b)).toEqual([]);
    b.names = [];
    b.sheets!.find((s) => s.name === '名簿')!.values = {};
    const diff = dumpDifferences(a, b);
    expect(diff).toContain('names');
    expect(diff).toContain('sheet 名簿: values');
    // bookkeeping of the inspection itself is not a difference
    const c = await fakeDump(base);
    (c as unknown as Record<string, unknown>).attempts = 2;
    expect(dumpDifferences(a, c)).toEqual([]);
  });

  it('the verdict reports a dump that still disagrees with its package as inspect-mismatch (not as a broken file)', () => {
    const b = { tool: 'x', openMode: 'normal', sheets: [sheet({ name: 'S', cells: { A1: 'h' } })], names: [] } as unknown as InspectResult;
    const o = { ...b, attempts: 3, suspicious: ['名前の定義が Excel の読み取りにありません: 大学一覧'] } as unknown as InspectResult;
    const v = checkOutput(b, o);
    const item = v.items.find((i: { code: string }) => i.code === 'inspect-mismatch');
    expect(item?.level).toBe('fail');
    expect(item?.message).toMatch(/3 回試行/);
    expect(v.ok).toBe(false);
    expect(checkOutput(b, { ...b, attempts: 1 } as unknown as InspectResult).items.some((i: { code: string }) => i.code === 'inspect-mismatch')).toBe(false);
  });
});

const enabled = process.env.QA_EXCEL === '1' && haveSamples && excelAvailable();

describe.skipIf(!enabled)('dump-check with real Excel', () => {
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'em-dc-xl-'));
  afterAll(() => fs.rmSync(tmp, { recursive: true, force: true }));

  it('real dumps are consistent with their packages on the first attempt, and two instances read the same', async () => {
    const files = [base, interview];
    const a = await inspectFiles(files, { recalc: true });
    a.forEach((d) => {
      expect(d.openMode).toBe('normal');
      expect(d.suspicious, `${d.file}: ${d.suspicious?.join(' / ')}`).toBeUndefined();
      expect(d.attempts).toBe(1);
      expect(d.transientFailures).toBeUndefined();
      expect(d.notReady).toBeUndefined();
    });
    const b = await inspectFiles(files, { recalc: true });
    files.forEach((_f, i) => expect(dumpDifferences(a[i], b[i]), path.basename(files[i])).toEqual([]));
  }, 8 * 60_000);

  it('an Excel that answers like the flake (fault injected once) is inspected again in a fresh instance', async () => {
    const marker = path.join(tmp, 'fault-marker.txt');
    const before = process.env.EM_INSPECT_FAULT_ONCE;
    process.env.EM_INSPECT_FAULT_ONCE = marker;
    try {
      const out = await inspectFiles([base], { recalc: true });
      expect(fs.existsSync(marker), 'the fault was injected into the first attempt').toBe(true);
      expect(out[0].attempts, 'inspected again').toBe(2);
      expect(out[0].suspicious).toBeUndefined();
      expect((out[0].names ?? []).map((n) => n.name)).toContain('大学一覧');
      expect(Object.keys(out[0].sheets!.find((s) => s.name === '名簿')!.values).length).toBeGreaterThan(100);
      // the same fault without retries is what check-output used to swallow: now it is at least reported
      fs.rmSync(marker);
      const single = await inspectFiles([base], { recalc: true, noRetry: true });
      expect(single[0].attempts).toBe(1);
      expect(single[0].suspicious?.join('\n')).toMatch(/名前の定義/);
    } finally {
      if (before === undefined) delete process.env.EM_INSPECT_FAULT_ONCE;
      else process.env.EM_INSPECT_FAULT_ONCE = before;
    }
  }, 8 * 60_000);
});
