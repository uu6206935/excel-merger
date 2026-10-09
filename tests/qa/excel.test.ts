// QA tools against REAL Excel (our own hidden instance): inspect.ps1 dumps, repair detection and the
// check-output verdict on workbooks modified by Excel itself. Opt-in (starts Excel):
//   QA_EXCEL=1 npx vitest run tests/qa/excel.test.ts
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import JSZip from 'jszip';
import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { excelAvailable, inspectFiles } from '../../tools/lib/inspect.mjs';
import type { InspectResult } from '../../tools/lib/inspect.mjs';
import { runPs } from '../../tools/lib/ps.mjs';
import { checkOutput } from '../../tools/lib/verdict.mjs';
import { sheetByName } from '../../tools/lib/compare-core.mjs';

const root = path.resolve(__dirname, '../..');
const base = path.join(root, 'samples', '01_学生名簿.xlsx');
const enabled = process.env.QA_EXCEL === '1' && fs.existsSync(base) && excelAvailable();

describe.skipIf(!enabled)('QA tools with real Excel', () => {
  let tmp = '';
  const dumps = new Map<string, InspectResult>();
  const names = ['good', 'converted', 'inserted', 'ref', 'name', 'wrongcol', 'bad-merge', 'bad-trunc'];

  beforeAll(async () => {
    tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'em-qa-'));
    const res = await runPs(path.join(__dirname, 'fixtures', 'make-variants.ps1'), ['-Base', base, '-OutDir', tmp], { timeoutMs: 5 * 60 * 1000 });
    if (res.code !== 0) throw new Error(`make-variants.ps1 failed: ${JSON.stringify(res.lines.filter((l) => l.type === 'error'))} ${res.stderr}`);
    // corrupt packages: an overlapping merge (Excel repairs it) and a truncated zip (Excel cannot open it)
    const src = fs.readFileSync(path.join(root, 'samples', '05_研修出欠.xlsx'));
    const zip = await JSZip.loadAsync(src);
    const xml = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
    zip.file('xl/worksheets/sheet1.xml', xml.replace(/<mergeCells count="1"><mergeCell ref="A1:G1"\/><\/mergeCells>/, '<mergeCells count="2"><mergeCell ref="A1:G1"/><mergeCell ref="B1:C2"/></mergeCells>'));
    fs.writeFileSync(path.join(tmp, 'bad-merge.xlsx'), await zip.generateAsync({ type: 'nodebuffer' }));
    fs.writeFileSync(path.join(tmp, 'bad-trunc.xlsx'), src.subarray(0, src.length - 200));
    const files = [base, ...names.map((n) => path.join(tmp, `${n}.xlsx`))];
    const out = await inspectFiles(files, { recalc: true, timeoutMs: 8 * 60 * 1000 });
    dumps.set('base', out[0]);
    names.forEach((n, i) => dumps.set(n, out[i + 1]));
  }, 12 * 60 * 1000);

  afterAll(() => {
    if (tmp) fs.rmSync(tmp, { recursive: true, force: true });
  });

  it('dumps the base workbook as Excel sees it', () => {
    const b = dumps.get('base')!;
    expect(b.openMode).toBe('normal');
    const roster = sheetByName(b, '名簿')!;
    expect(roster.values.A4).toBe('000100');
    expect(roster.formulas.J4).toBe('=DATEDIF(I4,基準日,"Y")');
    expect(roster.merges).toEqual(['A1:K1']);
    expect(roster.freeze).toEqual({ frozen: true, splitRow: 3, splitCol: 2 });
    expect(roster.autoFilter).toBe('A3:L63');
    expect(roster.hiddenColumns).toEqual(['L']);
    expect(roster.cf.map((c) => c.formula1)).toEqual(['=4', '=$L4>1']);
    expect(roster.dv.map((d) => `${d.range} ${d.formula1}`)).toEqual(['K4:K63 1,2,3,4,5', 'D4:D63 =マスタ!$A$2:$A$20']);
    expect(roster.probes.A1).toMatch(/\|b1\|fc16777215\|/); // bold white title on a theme fill
    const sum = sheetByName(b, '集計')!;
    expect(sum.charts[0].series[0].formula).toBe('=SERIES(集計!$B$4,集計!$A$5:$A$14,集計!$B$5:$B$14,1)');
    expect(Object.keys(sum.errorCounts)).toEqual([]);
    const piv = sheetByName(b, 'ピボット')!;
    expect(piv.pivots[0]).toMatchObject({ name: '大学別学年ピボット', source: '名簿!R3C1:R63C12', rowFields: ['大学'], columnFields: ['学年'] });
    expect((b.names ?? []).filter((n) => !n.name.startsWith('_')).map((n) => n.name).sort()).toEqual(['基準日', '志望度列', '大学一覧'].sort());
  });

  it('detects repair and unopenable files', () => {
    expect(dumps.get('bad-merge')!.openMode).toBe('repair');
    expect(dumps.get('bad-merge')!.repairLog!.join(' ')).toMatch(/sheet1\.xml/);
    expect(dumps.get('bad-trunc')!.openMode).toBe('failed');
    const v = checkOutput(dumps.get('base'), dumps.get('bad-merge'));
    expect(v.ok).toBe(false);
    expect(v.items.some((i) => i.code === 'repair')).toBe(true);
  });

  it('passes an output that Excel itself produced (columns added and moved)', () => {
    const v = checkOutput(dumps.get('base'), dumps.get('good'));
    expect(v.items.filter((i) => i.level === 'fail')).toEqual([]);
    expect(v.stats.addedHeaders?.sort()).toEqual(['判定', '面接官', '面接平均'].sort());
    expect(v.stats.formulasChecked).toBeGreaterThan(150);
    const c = checkOutput(dumps.get('base'), dumps.get('converted'));
    expect(c.ok).toBe(true);
    expect(c.items.some((i) => i.code === 'formulas-to-values')).toBe(true);
  });

  it('passes CF / DV that Excel itself extended over a column inserted right of 志望度 and reports the growth as info', () => {
    const v = checkOutput(dumps.get('base'), dumps.get('inserted'));
    expect(v.items.filter((i) => i.level === 'fail')).toEqual([]);
    expect(v.ok).toBe(true);
    const codes = v.items.map((i) => i.code);
    expect(codes).toEqual(expect.arrayContaining(['cf', 'cf-grown', 'dv-grown']));
    expect(v.stats.ruleGrowth).toEqual({ cf: 1, dv: 1 });
    const cf = v.items.find((i) => i.code === 'cf-grown')!;
    expect(cf.level).toBe('info');
    expect(cf.examples!.join()).toMatch(/K4:K63 → K4:L63/);
    expect(v.stats.formulasChecked).toBeGreaterThan(150);
  });

  it('fails outputs with #REF!, #NAME? or formulas pointing to another column', () => {
    const ref = checkOutput(dumps.get('base'), dumps.get('ref'));
    expect(ref.ok).toBe(false);
    expect(ref.items.filter((i) => i.level === 'fail').map((i) => i.code)).toEqual(expect.arrayContaining(['new-ref-text', 'new-error-#REF!']));
    // the pivot source and the DV that sat on the deleted column (大学) follow the deletion like Excel does: info, not fail
    expect(ref.items.filter((i) => i.level === 'info').map((i) => i.code)).toEqual(expect.arrayContaining(['pivot-source-deleted', 'dv-deleted']));
    const name = checkOutput(dumps.get('base'), dumps.get('name'));
    expect(name.items.filter((i) => i.level === 'fail').map((i) => i.code)).toEqual(['new-error-#NAME?']);
    const wrong = checkOutput(dumps.get('base'), dumps.get('wrongcol'));
    expect(wrong.items.filter((i) => i.level === 'fail').map((i) => i.code)).toEqual(['formula-columns']);
  });
});
