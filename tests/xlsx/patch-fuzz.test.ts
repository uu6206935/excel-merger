/**
 * Round 4: package-level robustness sweep of replayLog / buildPremerged (the explicit cases are in patch-robust.test.ts).
 *  1. metamorphic: the same package spelled differently (namespace prefix, single quotes, white space between elements,
 *     BOM, UTF-16) replays to the same result as the ordinary spelling, for real Excel fixtures and the golden op logs;
 *  2. buildPremerged on unusual but valid packages (no styles.xml / sharedStrings.xml, inline strings, xlsm, sheet names
 *     with quotes, UTF-16, prefixes, spans, a row of 16,384 styled cells) gives the result of the ordinary package;
 *  3. seeded fuzz (tests/golden/fuzz.ts): damaged packages end as a Japanese PatchError or a correct package. FUZZ_N=2000
 *     (env) runs many more mutants; tests/golden/tools/fuzz-cli.ts prints the seeds that failed.
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import type { PremergeInput } from '../../src/core/contracts';
import { buildPremerged } from '../../src/core/xlsx/premerge';
import { replayLog } from '../../src/core/xlsx/replay';
import { CASES } from '../golden/cases';
import { fuzzOps, fuzzPlans, fuzzPremerge, fuzzReplay, mutate, summarize } from '../golden/fuzz';
import { BASE_SST, SRC_SST, SH1, SH2, inlineStrings, mini, miniInput, noStyles, type Mini } from '../golden/mini-scenario';
import { edit, packageProblems, prefixMain, text, unpack, utf16le, type Parts } from '../golden/package-tools';
import { diffSnap, snapshot } from '../golden/snapshot';

const here = path.dirname(fileURLToPath(import.meta.url));
const FIX = path.resolve(here, '../fixtures/golden');
const fixture = (f: string) => new Uint8Array(fs.readFileSync(path.join(FIX, f)));
const FUZZ_N = Number(process.env.FUZZ_N) || 0; // 0 = the default sizes below

// ---------------------------------------------------------------------------------------------------------------
// 1. the same package, spelled differently
// ---------------------------------------------------------------------------------------------------------------

const xmlNames = (p: Parts) => [...p.keys()].filter((n) => /\.(xml|rels)$/i.test(n) && typeof p.get(n) === 'string');
const mainParts = (p: Parts) => xmlNames(p).filter((n) => /^xl\/(worksheets\/[^/]+|workbook|styles|sharedStrings|tables\/[^/]+)\.xml$/.test(n));

const SPELLINGS: Array<[string, (p: Parts) => void]> = [
  ['namespace prefix x:', (p) => { for (const n of xmlNames(p)) p.set(n, prefixMain(p.get(n) as string, 'x')); }],
  ['single quotes and line breaks between elements', (p) => {
    for (const n of mainParts(p)) p.set(n, (p.get(n) as string).replace(/(\s[\w:.-]+)="([^"']*)"/g, "$1='$2'").replace(/></g, '>\r\n  <'));
  }],
  // .NET writers put a space before the slash of an empty element: <c r="A1" s="1" />, <si />
  ['space before />', (p) => { for (const n of mainParts(p)) p.set(n, (p.get(n) as string).replace(/([\w"'])\/>/g, '$1 />')); }],
  ['BOM', (p) => { for (const n of xmlNames(p)) p.set(n, '﻿' + (p.get(n) as string)); }],
  ['UTF-16LE', (p) => { for (const n of mainParts(p)) p.set(n, utf16le(p.get(n) as string)); }],
];

describe('the same package spelled differently replays to the same result', () => {
  // every third golden case (all of them pass through the same code; the full set takes too long for one run)
  const cases = CASES.filter((c) => !c.expectError && !c.golden).filter((_c, i) => process.env.FUZZ_ALL || i % 3 === 0);
  const plain = new Map<string, Awaited<ReturnType<typeof snapshot>>>();

  for (const [label, respell] of SPELLINGS) {
    it(`${label} (${cases.length} golden cases)`, async () => {
      const bad: string[] = [];
      for (const c of cases) {
        const input = fixture(c.fixture);
        let a = plain.get(c.name);
        if (!a) { a = await snapshot(await replayLog(input, c.ops)); plain.set(c.name, a); }
        const out = await replayLog(await edit(input, respell), c.ops);
        const b = await snapshot(out);
        for (const d of diffSnap(a, b)) bad.push(`${c.name}: ${d.path} plain=${JSON.stringify(d.a)} respelled=${JSON.stringify(d.b)}`);
        for (const p of await packageProblems(out)) bad.push(`${c.name}: ${p}`);
      }
      expect(bad).toEqual([]);
    }, 120000);
  }
});

// ---------------------------------------------------------------------------------------------------------------
// 2. buildPremerged on unusual but valid packages
// ---------------------------------------------------------------------------------------------------------------

/** content of every sheet of a premerged package: values + formulas (and styles when asked) */
async function content(bytes: Uint8Array, styles = false) {
  const snap = await snapshot(bytes);
  return snap.sheets.map((x) => ({ name: x.name, values: x.values, formulas: x.formulas, ...(styles ? { styles: x.styles } : {}) }));
}

describe('buildPremerged on unusual but valid packages', () => {
  it('the reference result: imported columns, the appended row, the extended 合計 range, the Report formula', async () => {
    const m = await mini();
    const r = await buildPremerged(miniInput(m));
    expect(await packageProblems(r.bytes)).toEqual([]);
    const [main, report] = await content(r.bytes);
    expect(main.values).toMatchObject({ C1: 's:点数', D1: 's:備考', E1: 's:計算', C2: 'n:90', D2: 's:a', A5: 's:S9', C5: 'n:50', D5: 's:z', A7: 's:合計' });
    expect(main.formulas).toMatchObject({ E2: 'C2*2', E5: 'C5*2', B7: 'COUNTA(B2:B5)' });
    expect(report.formulas.A1).toMatch(/B7\*2$/);
    expect(r.layout.insertAt).toBe(3);
  });

  const variants: Array<[string, (m: Mini) => Promise<PremergeInput>, boolean]> = [
    ['namespace prefixes in base and source', async (m) => miniInput(m, await edit(m.base, (p) => { for (const k of xmlNames(p)) p.set(k, prefixMain(p.get(k) as string, 'ss')); }), await edit(m.src, (p) => { for (const k of xmlNames(p)) p.set(k, prefixMain(p.get(k) as string, 'x')); })), true],
    ['inline strings, no sharedStrings.xml in base and source', async (m) => miniInput(m, await edit(m.base, (p) => inlineStrings(p, BASE_SST, [SH1])), await edit(m.src, (p) => inlineStrings(p, SRC_SST, [SH1]))), true],
    ['no sharedStrings.xml in the base only (the part is created)', async (m) => miniInput(m, await edit(m.base, (p) => inlineStrings(p, BASE_SST, [SH1])), m.src), true],
    ['no styles.xml in the base (the part is created for the imported styles)', async (m) => miniInput(m, await edit(m.base, (p) => noStyles(p, [SH1, SH2])), m.src), false],
    ['no styles.xml in the source (imported cells have no style)', async (m) => miniInput(m, m.base, await edit(m.src, (p) => noStyles(p, [SH1]))), false],
    ['UTF-16 parts in base and source', async (m) => miniInput(m, await edit(m.base, (p) => { for (const k of [SH1, SH2, 'xl/workbook.xml', 'xl/sharedStrings.xml', 'xl/styles.xml']) p.set(k, utf16le(p.get(k) as string)); }), await edit(m.src, (p) => { for (const k of [SH1, 'xl/workbook.xml', 'xl/sharedStrings.xml', 'xl/styles.xml']) p.set(k, utf16le(p.get(k) as string)); })), true],
    ['spans, no <dimension>, rows and cells in reverse order', async (m) => {
      const respell = (xml: string) => xml.replace(/<dimension[^>]*\/>/, '').replace(/<row r="(\d+)">(.*?)<\/row>/g, (_x, r: string, cells: string) => `<row r="${r}" spans="1:9">${(cells.match(/<c [^>]*?(?:\/>|>.*?<\/c>)/g) ?? [cells]).reverse().join('')}</row>`);
      const rev = (xml: string) => xml.replace(/(<sheetData>)(.*)(<\/sheetData>)/, (_x, a: string, rows: string, c: string) => a + (rows.match(/<row .*?<\/row>/g) ?? [rows]).reverse().join('') + c);
      return miniInput(m, await edit(m.base, (p) => p.set(SH1, rev(respell(text(p, SH1))))), await edit(m.src, (p) => p.set(SH1, rev(respell(text(p, SH1))))));
    }, true],
  ];
  for (const [label, make, styles] of variants) {
    it(label, async () => {
      const m = await mini();
      const plain = await buildPremerged(miniInput(m));
      const r = await buildPremerged(await make(m));
      expect(await packageProblems(r.bytes)).toEqual([]);
      expect(r.layout).toEqual(plain.layout);
      expect(r.report.formulasKept).toBe(plain.report.formulasKept);
      const a = await content(plain.bytes, styles), b = await content(r.bytes, styles);
      expect(b).toEqual(a);
    });
  }

  it('sheet names with a quote, an ampersand and Japanese text (base and source)', async () => {
    const m = await mini("It's 名簿 & 売上", "A&B 'src'");
    const r = await buildPremerged(miniInput(m));
    expect(await packageProblems(r.bytes)).toEqual([]);
    const [main, report] = await content(r.bytes);
    expect(main.name).toBe("It's 名簿 & 売上");
    expect(main.values).toMatchObject({ C2: 'n:90', A5: 's:S9' });
    expect(main.formulas.B7).toBe('COUNTA(B2:B5)');
    // the Report formula quotes the name (doubled apostrophe) and follows the appended row
    expect(report.formulas.A1).toMatch(/^'It''s 名簿 & 売上'!B7\*2$/);
    // an unknown sheet name is a Japanese PatchError, not a crash
    const bad = miniInput(m);
    bad.base = { ...bad.base, region: { ...bad.base.region, sheet: '存在しない' } };
    await expect(buildPremerged(bad)).rejects.toThrow(/シート「存在しない」/);
  });

  it('a macro-enabled base (xlsm): vbaProject.bin and its content type / relationship are carried through', async () => {
    const m = await mini();
    const bin = new Uint8Array(2000).map((_v, i) => (i * 13 + 5) & 255);
    const xlsm = await edit(m.base, (p) => {
      p.set('xl/vbaProject.bin', bin);
      p.set('[Content_Types].xml', text(p, '[Content_Types].xml')
        .replace('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml', 'application/vnd.ms-excel.sheet.macroEnabled.main+xml')
        .replace('<Default Extension="xml"', '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/><Default Extension="xml"'));
      p.set('xl/_rels/workbook.xml.rels', text(p, 'xl/_rels/workbook.xml.rels').replace('</Relationships>', '<Relationship Id="rId99" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/></Relationships>'));
    });
    const r = await buildPremerged(miniInput(m, xlsm));
    const out = await unpack(r.bytes);
    expect(out.get('xl/vbaProject.bin')).toEqual(bin);
    expect(text(out, '[Content_Types].xml')).toContain('macroEnabled.main+xml');
    expect(text(out, 'xl/_rels/workbook.xml.rels')).toContain('vbaProject.bin');
    expect(await packageProblems(r.bytes)).toEqual([]);
    expect((await content(r.bytes))[0].values).toMatchObject({ C2: 'n:90', A5: 's:S9' });
  });

  it('a base row of 16,384 styled cells: the imported columns push the empty styled cells off the grid, nothing else changes', async () => {
    const m = await mini();
    let extra = '';
    for (let c = 3; c <= 16384; c++) {
      let t = '', x = c;
      while (x > 0) { t = String.fromCharCode(65 + ((x - 1) % 26)) + t; x = Math.floor((x - 1) / 26); }
      extra += `<c r="${t}1" s="1"/>`;
    }
    const wide = await edit(m.base, (p) => p.set(SH1, text(p, SH1).replace('</c></row><row r="2">', `</c>${extra}</row><row r="2">`)));
    const r = await buildPremerged(miniInput(m, wide));
    expect(await packageProblems(r.bytes)).toEqual([]);
    const xml = text(await unpack(r.bytes), SH1);
    expect(xml).not.toMatch(/<c r="XF[E-Z]/);
    expect((await content(r.bytes))[0].values).toMatchObject({ C1: 's:点数', C2: 'n:90' });
  }, 30000);

  it('a base cell beyond XFD is refused in Japanese when the imported columns have to move it', async () => {
    const m = await mini();
    const bad = await edit(m.base, (p) => p.set(SH1, text(p, SH1).replace('</row><row r="2">', '<c r="XFE1"><v>1</v></c></row><row r="2">')));
    await expect(buildPremerged(miniInput(m, bad))).rejects.toThrow(/XFD|範囲/);
  });

  it('a source with a cell beyond XFD in a row that is never read still imports (nothing of it is written)', async () => {
    const m = await mini();
    const odd = await edit(m.src, (p) => p.set(SH1, text(p, SH1).replace('<row r="5">', '<row r="9"><c r="ZZZ9"><v>1</v></c></row><row r="5">')));
    const r = await buildPremerged(miniInput(m, m.base, odd));
    expect((await content(r.bytes))[0].values).toMatchObject({ C2: 'n:90', A5: 's:S9' });
    expect(await packageProblems(r.bytes)).toEqual([]);
  });
});

// ---------------------------------------------------------------------------------------------------------------
// 3. seeded fuzz
// ---------------------------------------------------------------------------------------------------------------

describe('seeded package fuzz', () => {
  const report = (list: Awaited<ReturnType<typeof fuzzReplay>>) => {
    const sum = summarize(list);
    // internal = an unforeseen exception that only the safety net caught; violations = a non-PatchError / a damaged output
    expect(sum.violations.map((x) => `seed ${x.seed} ${x.what}: ${x.detail}`)).toEqual([]);
    expect(sum.internal.map((x) => `seed ${x.seed} ${x.what}: ${x.detail}`)).toEqual([]);
    return sum;
  };

  it('replayLog: damaged Excel fixtures end as a Japanese PatchError or a well-formed package', async () => {
    const count = FUZZ_N || 100;
    const sum = report(await fuzzReplay(0, count));
    // the fuzzer really exercises both outcomes
    expect(sum.ok).toBeGreaterThan(count * 0.2);
    expect(sum.refused).toBeGreaterThan(count * 0.05);
  }, 600000);

  it('buildPremerged: damaged base / source packages end as a Japanese PatchError or a well-formed package', async () => {
    const count = FUZZ_N || 80;
    const sum = report(await fuzzPremerge(0, count));
    expect(sum.ok).toBeGreaterThan(count * 0.2);
    expect(sum.refused).toBeGreaterThan(count * 0.05);
  }, 600000);

  it('replayLog: random op logs (in range and out of range) on the intact fixtures', async () => {
    const count = FUZZ_N || 120;
    const sum = report(await fuzzOps(0, count));
    expect(sum.ok).toBeGreaterThan(count * 0.3);
  }, 600000);

  it('buildPremerged: random plans (column subsets / orders, matches, appended rows, options) on intact packages, then random ops on the result', async () => {
    const count = FUZZ_N || 60;
    const sum = report(await fuzzPlans(0, count));
    expect(sum.ok).toBe(count);
  }, 600000);

  it('mutate() is deterministic (a seed reproduces its mutant)', async () => {
    const b = fixture('g-table.xlsx');
    const a1 = await mutate(b, 42), a2 = await mutate(b, 42), a3 = await mutate(b, 43);
    expect(a1.what).toBe(a2.what);
    expect(Buffer.from(a1.bytes).equals(Buffer.from(a2.bytes))).toBe(true);
    expect(a1.what === a3.what && Buffer.from(a1.bytes).equals(Buffer.from(a3.bytes))).toBe(false);
  });
});
