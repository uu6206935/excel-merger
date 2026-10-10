/**
 * The copied-sheets oracle (oracle.ts) under vitest.
 *  - always: the Node half runs without Excel — the app's pipeline builds the premerged package WITH copies for a handful of combinations (samples in
 *    other card orders, a foreign file, an object fixture, three files with the same name, a rich formula set), a random edit log is made through the doc
 *    API, replayLog exports it, and the doc's own view of the result must equal the package read back; the copies' names must obey the rule and
 *    the layout / report must agree (static findings);
 *  - EXCEL_TESTS=1: the real run (run.ps1): every combination (>= 70), both engines, one hidden Excel per batch of workbooks, every copied sheet
 *    compared with its source. EXCEL_ORACLE_ONLY (a regex of combination names) and EXCEL_ORACLE_LIMIT narrow it (default: a 14-combination
 *    selection over every group); the whole run is run.ps1 without arguments.
 */
import { spawn } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { describe, expect, it } from 'vitest';
import { KNOWN_OPEN } from '../consistency/check';
import { buildCombos } from './combos';
import { HERE } from './excel';
import { INFO_CLASSES } from './judge';
import { refineEngineFindings, runOracle } from './oracle';

const EXCEL = process.env.EXCEL_TESTS === '1';

describe('copied-sheets oracle: the combinations', () => {
  it('there are at least 40 of them, over every group, with unique ids', () => {
    const combos = buildCombos();
    expect(combos.length).toBeGreaterThanOrEqual(40);
    expect(new Set(combos.map((c) => c.id)).size).toBe(combos.length);
    for (const g of ['smp', 'cor', 'ob', 'nam', 'rich']) expect(combos.some((c) => c.group === g), g).toBe(true);
    // every sample as the base, three files with the same name, a name that needs shortening, 1904, xls, csv, xlsm, a hidden sheet
    const ids = combos.map((c) => c.id).join(' ');
    for (const re of [/smp-01-/, /smp-02-/, /smp-03csv/, /smp-04xls/, /smp-05-/, /nam-same-name-x3/, /nam-long-file-name/, /1904/, /macro/, /hidden/]) expect(ids).toMatch(re);
  });
});

describe('copied-sheets oracle: refineEngineFindings (what is only how Excel writes the file)', () => {
  const f = (cls: string, detail: string, scenario = 's') => ({ scenario, cls, detail });
  it('a note line break (CR LF vs LF), the localized name of the Normal style, a default column width written another way, an x14 data validation', () => {
    const got = refineEngineFindings([
      f('engine-package-comments', 'シート: a のみ: B5: 本文\n返信  b のみ: B5: 本文\r\n返信'),
      f('engine-package-comments', 'シート: a のみ: B5: 本文  b のみ: B5: 別の本文'),
      f('engine-package-cols', 'copy: A 列の書式 style: Normal -> 標準 ; B 列の書式 style: Normal -> 標準'),
      f('engine-package-cols', 'copy: A 列の書式 style: Normal -> 標準 ; B 幅 8.43 -> 20'),
      f('engine-package-sheetProps', 'copy: "tab=/baseColWidth=10,defaultColWidth=,zeroHeight=" -> "tab=/baseColWidth=,defaultColWidth=11,zeroHeight="'),
      f('engine-package-sheetProps', 'copy: "tab=/baseColWidth=10,defaultColWidth=,zeroHeight=" -> "tab=/baseColWidth=,defaultColWidth=25,zeroHeight="'),
      f('engine-package-rules', '名簿: 名簿 rule only in A: dv list マスタ!$A$2:$A$9 (30 cells)'),
      f('engine-package-rules', '名簿: 名簿 rule only in B: dv(ext) list マスタ!$A$2:$A$9 (30 cells)'),
      f('engine-package-rules', '名簿: 名簿 rule only in A: dv list 他!$A$2:$A$9 (30 cells)'),
    ]).map((x) => x.cls);
    expect(got).toEqual([
      'info-engine-note-linebreak',
      'engine-package-comments',
      'info-engine-style-name',
      'engine-package-cols',
      'info-engine-default-width',
      'engine-package-sheetProps',
      'info-engine-dv-ext',
      'info-engine-dv-ext',
      'engine-package-rules',
    ]);
  });
});

describe('copied-sheets oracle: Node half (no Excel)', () => {
  it('builds copies for a selection of combinations; names obey the rule; the doc equals replayLog for every random log', async () => {
    const out = fs.mkdtempSync(path.join(os.tmpdir(), 'copied-oracle-'));
    try {
      const only = '^(smp-01-rev-keep|smp-05-base-full-drop|cor-hc-05-mac-excel-2011-1904|ob-all-keep|nam-same-name-x3|nam-long-file-name|nam-forbidden-chars|rich-f07_othersheet-c|cor-ex-macro-as-base)$';
      const r = await runOracle(['--out', out, '--no-excel', '--only', only]);
      expect(r.recs.length).toBeGreaterThanOrEqual(9);
      // (without Excel there are no dumps: that is the only thing the judge may report)
      const failing = r.findings.filter((f) => !INFO_CLASSES.test(f.cls) && !KNOWN_OPEN.test(f.cls) && f.cls !== 'no-dump');
      expect(failing.map((f) => `${f.scenario} ${f.cls}: ${f.detail}`.slice(0, 300))).toEqual([]);
      expect(r.recs.filter((x) => x.docVsReplay.length).map((x) => `${x.id}: ${x.docVsReplay.join(' | ')}`)).toEqual([]);
      for (const rec of r.recs) expect(rec.copies.length, rec.id).toBeGreaterThan(0);
      // the names of the three files with the same name: (1) and (2) go to the later ones
      const same = r.recs.find((x) => x.combo === 'nam-same-name-x3')!;
      expect(same.copies.map((c) => c.name).filter((n) => /_名簿(\(\d+\))?$/.test(n))).toEqual(['01_学生名簿_名簿', '01_学生名簿_名簿(1)', '01_学生名簿_名簿(2)']);
      const long = r.recs.find((x) => x.combo === 'nam-long-file-name')!;
      expect(long.copies.every((c) => c.name.length <= 31)).toBe(true);
      expect(r.recs.some((x) => x.opsKinds.move)).toBe(true);
      expect(r.recs.some((x) => x.opsKinds.edit)).toBe(true);
    } finally {
      fs.rmSync(out, { recursive: true, force: true });
    }
  }, 600_000);
});

describe.skipIf(!EXCEL)('copied-sheets oracle: real Excel (EXCEL_TESTS=1)', () => {
  it('every copied sheet equals its source, both engines agree, edits never change a copy', async () => {
    const out = fs.mkdtempSync(path.join(os.tmpdir(), 'copied-oracle-xl-'));
    try {
      const ps1 = path.join(HERE, 'run.ps1');
      const only = process.env.EXCEL_ORACLE_ONLY ?? '^(smp-01-all-keep|smp-02-base-full-drop|smp-05-base-keep|cor-hc-05-mac-excel-2011-1904|cor-hc-17-as-base|cor-py-04-openpyxl-table|cor-ex-macro-as-base|ob-all-keep|ob-pivot|nam-same-name-x3|nam-long-file-name|rich-f07_othersheet-c|rich-f12_footer-c|cor-ex-1904-as-base)$';
      const args = ['-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', ps1, '-Out', out, '-Only', only];
      if (process.env.EXCEL_ORACLE_LIMIT) args.push('-Limit', process.env.EXCEL_ORACLE_LIMIT);
      // asynchronous: a blocking call would starve vitest's worker heartbeat for the minutes Excel needs
      const code = await new Promise<number | null>((resolve, reject) => {
        const child = spawn('powershell.exe', args, { stdio: ['ignore', 'inherit', 'inherit'] });
        const timer = setTimeout(() => child.kill(), 120 * 60_000);
        child.on('error', reject);
        child.on('close', (c) => {
          clearTimeout(timer);
          resolve(c);
        });
      });
      const findings = JSON.parse(fs.readFileSync(path.join(out, 'findings.json'), 'utf8')) as Array<{ scenario: string; cls: string; detail: string }>;
      const failing = findings.filter((f) => !INFO_CLASSES.test(f.cls) && !KNOWN_OPEN.test(f.cls));
      expect(failing.map((f) => `${f.scenario} ${f.cls}: ${f.detail}`.slice(0, 300))).toEqual([]);
      expect(code).toBe(0);
    } finally {
      fs.rmSync(out, { recursive: true, force: true });
    }
  }, 130 * 60_000);
});
