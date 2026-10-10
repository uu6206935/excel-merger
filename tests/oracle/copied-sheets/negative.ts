/**
 * Negative controls of the copied-sheets oracle, in real Excel (task V9): "does the oracle see a deliberately planted fault?".
 * A premerged package of a scenario is damaged in ONE way at a time (the way a faulty copier would): a value, a formula, a sheet name, a sheet's
 * visibility, the conditional formats, a column width, a merged range, a data validation, a defined name, a table, a chart series, a picture, a pivot
 * cache refresh flag, a leaked workbook-level name ... — written as <combo>.neg-<name>.xlsx, opened and recalculated by the same hidden Excel batch as
 * everything else, and judged by the SAME judge (judgeNames / judgeCopies / judgeCopyNames). Every mutation must be found, in the expected class.
 * (judge.test.ts does the same on hand-made dumps; this one proves that the whole chain — dump script, judge, classes — sees them.)
 */
import fs from 'node:fs';
import path from 'node:path';
import JSZip from 'jszip';
import { readBook as readDump, runBooks, type BookDump } from './excel';
import { expectedSheets, judgeCopies, judgeCopyNames, judgeNames, type CopyRec, type Expected } from './judge';
import type { Finding } from '../consistency/check';
import type { SourceBook } from '../../../src/core/types';

export interface Mutation {
  name: string;
  /** the finding classes (prefix match on the class WITHOUT the "-pre") of which at least one must be reported */
  expect: RegExp;
  /** only for the scenarios whose combination matches (the object parts of a package belong to the copies only in some of them) */
  only?: RegExp;
  /** returns the damaged package, or null when this package has nothing to damage that way */
  apply(zip: JSZip, ctx: { sheetPart: (sheetName: string) => string; copy: (n: number) => string; count: number }): Promise<string | null>;
}

const unq = (s: string) => s.replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&');

async function partOf(zip: JSZip, sheetName: string): Promise<string> {
  const wb = await zip.file('xl/workbook.xml')!.async('string');
  const rels = await zip.file('xl/_rels/workbook.xml.rels')!.async('string');
  for (const m of wb.matchAll(/<sheet\b[^>]*>/g)) {
    const name = unq(/\bname="([^"]*)"/.exec(m[0])?.[1] ?? '');
    const rid = /\br:id="([^"]*)"/.exec(m[0])?.[1];
    if (name !== sheetName || !rid) continue;
    const rel = [...rels.matchAll(/<Relationship\b[^>]*>/g)].map((x) => x[0]).find((x) => x.includes(` Id="${rid}"`));
    const target = /\bTarget="([^"]*)"/.exec(rel ?? '')?.[1];
    if (target) return target.startsWith('/') ? target.slice(1) : `xl/${target}`;
  }
  throw new Error(`sheet ${sheetName} not found`);
}

/** the mutations; `copy(n)` = the name of the n-th copied sheet (0-based) of the scenario */
export const MUTATIONS: Mutation[] = [
  {
    name: 'value',
    expect: /^copy-value/,
    async apply(zip, ctx) {
      const part = await partOf(zip, ctx.copy(0));
      const x = await zip.file(part)!.async('string');
      // the first numeric constant of the sheet
      const m = /(<c [^>]*>)<v>(-?\d+(?:\.\d+)?)<\/v>/.exec(x.replace(/<c [^>]*t="(?:s|str|b|e)"[^>]*>[\s\S]*?<\/c>/g, ''));
      if (!m) return null;
      const out = x.replace(`<v>${m[2]}</v>`, `<v>${Number(m[2]) + 7}</v>`);
      if (out === x) return null;
      zip.file(part, out);
      return `a number of ${ctx.copy(0)} changed by 7`;
    },
  },
  {
    name: 'formula',
    expect: /^copy-(formula|value)/,
    async apply(zip, ctx) {
      for (let k = 0; k < Math.min(8, ctx.count); k++) {
        const part = await partOf(zip, ctx.copy(k));
        const x = await zip.file(part)!.async('string');
        const m = /<f>([^<]*[A-Z]+\d+[^<]*)<\/f>/.exec(x);
        if (!m) continue;
        zip.file(part, x.replace(m[0], `<f>${m[1]}+1</f>`));
        return `a formula of ${ctx.copy(k)} got +1`;
      }
      return null;
    },
  },
  {
    name: 'sheet-name',
    expect: /^(names|copy-source-missing)/,
    async apply(zip, ctx) {
      let wb = await zip.file('xl/workbook.xml')!.async('string');
      const from = ctx.copy(1);
      if (!from) return null;
      const esc = from.replace(/&/g, '&amp;');
      if (!wb.includes(`name="${esc}"`)) return null;
      wb = wb.replace(`name="${esc}"`, `name="${esc}X"`);
      zip.file('xl/workbook.xml', wb);
      return `${from} renamed to ${from}X (formulas elsewhere keep pointing at the old name: Excel may report #REF!)`;
    },
  },
  {
    name: 'visibility',
    expect: /^(copy-visibility|names)/,
    async apply(zip, ctx) {
      let wb = await zip.file('xl/workbook.xml')!.async('string');
      const esc = ctx.copy(0).replace(/&/g, '&amp;');
      const m = new RegExp(`<sheet\\b[^>]*name="${esc.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"[^>]*>`).exec(wb);
      if (!m || /state=/.test(m[0])) return null;
      wb = wb.replace(m[0], m[0].replace(/\/>$/, ' state="hidden"/>'));
      zip.file('xl/workbook.xml', wb);
      return `${ctx.copy(0)} hidden`;
    },
  },
  {
    name: 'conditional-formats',
    expect: /^copy-format/,
    async apply(zip, ctx) {
      for (let k = 0; k < Math.min(10, ctx.count); k++) {
        const part = await partOf(zip, ctx.copy(k));
        const x = await zip.file(part)!.async('string');
        if (!/<conditionalFormatting\b/.test(x)) continue;
        zip.file(part, x.replace(/<conditionalFormatting\b[\s\S]*?<\/conditionalFormatting>/g, ''));
        return `the conditional formats of ${ctx.copy(k)} removed`;
      }
      return null;
    },
  },
  {
    name: 'column-width',
    expect: /^copy-format/,
    async apply(zip, ctx) {
      for (let k = 0; k < Math.min(10, ctx.count); k++) {
        const part = await partOf(zip, ctx.copy(k));
        const x = await zip.file(part)!.async('string');
        const m = /<col\b[^>]*\bwidth="(\d+(?:\.\d+)?)"[^>]*>/.exec(x);
        if (!m) continue;
        zip.file(part, x.replace(m[0], m[0].replace(`width="${m[1]}"`, `width="${Number(m[1]) + 12}"`)));
        return `a column of ${ctx.copy(k)} 12 characters wider`;
      }
      return null;
    },
  },
  {
    name: 'data-validation',
    expect: /^copy-format/,
    async apply(zip, ctx) {
      for (let k = 0; k < Math.min(10, ctx.count); k++) {
        const part = await partOf(zip, ctx.copy(k));
        const x = await zip.file(part)!.async('string');
        if (!/<dataValidations\b/.test(x)) continue;
        zip.file(part, x.replace(/<dataValidations\b[\s\S]*?<\/dataValidations>/g, ''));
        return `the data validations of ${ctx.copy(k)} removed`;
      }
      return null;
    },
  },
  {
    name: 'hidden-column',
    expect: /^copy-format/,
    async apply(zip, ctx) {
      for (let k = 0; k < Math.min(10, ctx.count); k++) {
        const part = await partOf(zip, ctx.copy(k));
        const x = await zip.file(part)!.async('string');
        const m = /<col\b(?![^>]*\bhidden=)[^>]*\/>/.exec(x);
        if (!m) continue;
        zip.file(part, x.replace(m[0], m[0].replace(/\/>$/, ' hidden="1"/>')));
        return `a column of ${ctx.copy(k)} hidden`;
      }
      return null;
    },
  },
  {
    name: 'chart-series',
    expect: /^copy-object/,
    only: /^(ob-|cor-py-(18|27))/,
    async apply(zip) {
      for (const n of Object.keys(zip.files).filter((f) => /^xl\/charts\/chart\d*\.xml$/.test(f))) {
        const x = await zip.file(n)!.async('string');
        const m = /<c:val>[\s\S]*?<c:f>([^<]*)<\/c:f>/.exec(x);
        if (!m) continue;
        const ref = m[1].replace(/\$([A-Z]+)\$(\d+):\$([A-Z]+)\$(\d+)/, (_a, c1, r1, c2, r2) => `$${c1}$${Number(r1) + 1}:$${c2}$${r2}`);
        if (ref === m[1]) continue;
        zip.file(n, x.replace(`<c:f>${m[1]}</c:f>`, `<c:f>${ref}</c:f>`));
        return `a series of ${n} reads one row less`;
      }
      return null;
    },
  },
  {
    name: 'picture',
    expect: /^copy-object/,
    only: /^(ob-|cor-py-(18|27))/,
    async apply(zip) {
      for (const n of Object.keys(zip.files).filter((f) => /^xl\/drawings\/drawing\d+\.xml$/.test(f))) {
        const x = await zip.file(n)!.async('string');
        const m = /<(?:xdr:)?pic\b[\s\S]*?<\/(?:xdr:)?pic>/.exec(x);
        if (!m) continue;
        zip.file(n, x.replace(/<(xdr:)?(oneCellAnchor|twoCellAnchor)\b([\s\S]*?)<\/(xdr:)?\2>/, (a) => (a.includes('pic') ? '' : a)));
        return `a picture of ${n} removed`;
      }
      return null;
    },
  },
  {
    name: 'leaked-name',
    expect: /^copy-names/,
    async apply(zip, ctx) {
      let wb = await zip.file('xl/workbook.xml')!.async('string');
      const sheet = ctx.copy(0).replace(/'/g, "''");
      const def = `<definedName name="LeakedByOracle">'${sheet}'!$A$1</definedName>`;
      wb = /<definedNames>/.test(wb) ? wb.replace('<definedNames>', `<definedNames>${def}`) : wb.replace('</sheets>', `</sheets><definedNames>${def}</definedNames>`);
      zip.file('xl/workbook.xml', wb);
      return 'a workbook-level name that no base has';
    },
  },
];

export interface NegativeResult {
  findings: Finding[];
  table: Array<{ combo: string; mutation: string; what: string; found: string[]; ok: boolean }>;
}

/**
 * Damages the premerged package of each scenario in `recs` once per mutation, asks Excel (one batch) and judges. `expectedOf` gives the expectation
 * (sheet names) the judge compares with; the dumps of the sources come from the oracle's own dump folder.
 */
export async function runNegative(
  recs: CopyRec[],
  dirs: { pkgDir: string; dumpDir: string; tempRoot: string },
  srcDumpOf: (rec: CopyRec) => { srcDumps: Map<string, BookDump | null>; base: BookDump | null; srcBooks: Map<string, SourceBook> },
  jobIdOf: (file: string) => string,
  log: (s: string) => void,
): Promise<NegativeResult> {
  const made: Array<{ rec: CopyRec; mutation: Mutation; file: string; what: string; id: string }> = [];
  for (const rec of recs) {
    const bytes = fs.readFileSync(path.join(dirs.pkgDir, rec.files.pre));
    for (const mutation of MUTATIONS) {
      if (mutation.only && !mutation.only.test(rec.combo)) continue;
      const zip = await JSZip.loadAsync(bytes);
      const what = await mutation.apply(zip, { sheetPart: () => '', copy: (n) => rec.copies[n]?.name ?? '', count: rec.copies.length });
      if (what === null) continue;
      const file = path.join(dirs.pkgDir, `${rec.combo}.neg-${mutation.name}.${rec.ext}`);
      fs.writeFileSync(file, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }));
      made.push({ rec, mutation, file, what, id: jobIdOf(file) });
    }
  }
  log(`negative controls: ${made.length} damaged packages`);
  await runBooks(
    made.map((m) => ({ id: m.id, file: m.file, deep: m.rec.copies.map((c) => c.name) })),
    dirs.dumpDir,
    { tempRoot: dirs.tempRoot, batchSize: 24, log },
  );
  const findings: Finding[] = [];
  const table: NegativeResult['table'] = [];
  for (const m of made) {
    const d = readDump(dirs.dumpDir, m.id);
    const { srcDumps, base, srcBooks } = srcDumpOf(m.rec);
    const exp: Expected = expectedSheets(m.rec, srcDumps, base);
    const got: Finding[] = d ? [...judgeNames(m.rec, 'pre', d, exp), ...judgeCopies(m.rec, 'pre', d, exp, srcDumps, srcBooks).findings, ...judgeCopyNames(m.rec, 'pre', d, exp, srcDumps, base)] : [];
    // a damaged package that Excel refuses to open is a find as well (open-pre)
    const classes = [...new Set(got.map((f) => f.cls.replace(/-pre$/, '')))];
    const ok = classes.some((c) => m.mutation.expect.test(c)) || classes.some((c) => /^open/.test(c));
    table.push({ combo: m.rec.combo, mutation: m.mutation.name, what: m.what, found: classes, ok });
    if (!ok) findings.push({ scenario: m.rec.id, cls: 'negative-not-found', detail: `${m.mutation.name}: ${m.what} — the oracle reported ${classes.join(', ') || 'nothing'}` });
  }
  return { findings, table };
}
