/**
 * Dev CLI for the golden harness (run with `npx tsx tests/golden/tools/cli.ts ...`).
 *   replay <in.xlsx> <out.xlsx> <ops.json>     ops.json = LoggedOp[] (or {ops:[...]})
 *   diff <ours.xlsx> <excel.xlsx> [keys]       prints snapshot differences
 *   snap <file.xlsx>                           prints the snapshot JSON
 */
import fs from 'node:fs';
import { replayLog } from '../../../src/core/xlsx/replay';
import { diffSnap, snapshot } from '../snapshot';
import type { LoggedOp } from '../../../src/core/types';

async function main(): Promise<void> {
  const [cmd, ...args] = process.argv.slice(2);
  if (cmd === 'replay') {
    const [inp, out, opsFile] = args;
    const raw = JSON.parse(fs.readFileSync(opsFile, 'utf8').replace(/^﻿/, ''));
    const ops: LoggedOp[] = Array.isArray(raw) ? raw : raw.ops;
    const t0 = Date.now();
    const res = await replayLog(new Uint8Array(fs.readFileSync(inp)), ops);
    fs.writeFileSync(out, res);
    console.log(`wrote ${out} (${res.length} bytes, ${Date.now() - t0} ms)`);
  } else if (cmd === 'diff') {
    const [a, b, keys] = args;
    const sa = await snapshot(new Uint8Array(fs.readFileSync(a)));
    const sb = await snapshot(new Uint8Array(fs.readFileSync(b)));
    const d = diffSnap(sa, sb, keys ? (keys.split(',') as never) : undefined);
    for (const x of d) console.log(`${x.path}\n   ours : ${JSON.stringify(x.a)}\n   excel: ${JSON.stringify(x.b)}`);
    console.log(`${d.length} differences`);
  } else if (cmd === 'cases') {
    const { CASES } = await import('../cases');
    // cases with their own golden file (probes) were produced by their probe script
    fs.writeFileSync(args[0], JSON.stringify(CASES.filter((c) => !c.golden).map((c) => ({ name: c.name, fixture: c.fixture, ops: c.ops })), null, 1));
    console.log(`wrote ${CASES.length} cases to ${args[0]}`);
  } else if (cmd === 'ours') {
    // replay every case into <outDir>/<name>.ours.xlsx (for opening in Excel)
    const { CASES } = await import('../cases');
    const [fixDir, outDir] = args;
    for (const c of CASES) {
      try {
        const res = await replayLog(new Uint8Array(fs.readFileSync(`${fixDir}/${c.fixture}`)), c.ops);
        fs.writeFileSync(`${outDir}/${c.name}.ours.xlsx`, res);
      } catch (e) {
        console.log(`${c.name}: ${(e as Error).message}`);
      }
    }
    console.log(`wrote ${CASES.length} files`);
  } else if (cmd === 'diffall') {
    // replay + diff every case against Excel's golden (fixtures in tests/fixtures/golden)
    const { CASES } = await import('../cases');
    const [only, max] = args;
    for (const c of CASES) {
      if (only && only !== '-' && !c.name.startsWith(only)) continue;
      const golden = `tests/fixtures/golden/${c.golden ?? `out/${c.name}.xlsx`}`;
      let ours: Uint8Array;
      try {
        ours = await replayLog(new Uint8Array(fs.readFileSync(`tests/fixtures/golden/${c.fixture}`)), c.ops);
      } catch (e) {
        console.log(`== ${c.name}: ${c.expectError ? 'refused as expected' : 'ERROR ' + (e as Error).message}`);
        continue;
      }
      if (c.expectError) { console.log(`== ${c.name}: NOT REFUSED`); continue; }
      const d = diffSnap(await snapshot(ours), await snapshot(new Uint8Array(fs.readFileSync(golden))))
        .filter((x) => !Object.keys(c.known ?? {}).some((k) => x.path.startsWith(k)));
      console.log(`== ${c.name}: ${d.length} differences`);
      for (const x of d.slice(0, Number(max ?? 8))) console.log(`  ${x.path}\n     ours : ${JSON.stringify(x.a)}\n     excel: ${JSON.stringify(x.b)}`);
    }
  } else if (cmd === 'premerge') {
    // premerge scenario -> <outDir>/pm-<variant>.xlsx (+ report on stdout)
    const { scenario } = await import('../premerge-scenario');
    const { buildPremerged } = await import('../../../src/core/xlsx/premerge');
    const [outDir] = args;
    for (const v of ['plain', 'table'] as const) {
      const t0 = Date.now();
      const r = await buildPremerged(scenario(v));
      fs.writeFileSync(`${outDir}/pm-${v}.xlsx`, r.bytes);
      console.log(`pm-${v}: ${Date.now() - t0} ms`, JSON.stringify(r.layout), JSON.stringify(r.report, null, 1));
    }
  } else if (cmd === 'excelcheck') {
    // every output the tests produce, for tests/golden/com/opentest.ps1 (open without repair)
    const { CASES } = await import('../cases');
    const { scenario } = await import('../premerge-scenario');
    const { buildPremerged } = await import('../../../src/core/xlsx/premerge');
    const [outDir] = args;
    for (const c of CASES) {
      if (c.expectError) continue;
      fs.writeFileSync(`${outDir}/${c.name}.xlsx`, await replayLog(new Uint8Array(fs.readFileSync(`tests/fixtures/golden/${c.fixture}`)), c.ops));
    }
    // round 3: a block from inside a table landing right after its last column (in-table reorder), then everything of the
    // round in one log (unmerge with the look spread, quote prefixes, columns inserted into a table, a totals label)
    const fix = (f: string) => new Uint8Array(fs.readFileSync(`tests/fixtures/golden/${f}`));
    fs.writeFileSync(`${outDir}/r3-mov-table-end.xlsx`, await replayLog(fix('g-table.xlsx'), [{ kind: 'struct', sheet: '名簿', op: { type: 'move', axis: 'col', from: 2, count: 1, to: 8 } }]));
    fs.writeFileSync(`${outDir}/r3-mov-table-end2.xlsx`, await replayLog(fix('g-table.xlsx'), [{ kind: 'struct', sheet: '名簿', op: { type: 'move', axis: 'col', from: 4, count: 2, to: 8 } }, { kind: 'struct', sheet: '名簿', op: { type: 'move', axis: 'col', from: 1, count: 1, to: 8 } }]));
    fs.writeFileSync(`${outDir}/r3-mixed.xlsx`, await replayLog(fix('r3-misc.xlsx'), [
      { kind: 'unmerge', sheet: 'Band', refs: ['A1:F1', 'A3:C5'], centerAcross: true },
      { kind: 'unmerge', sheet: 'Outline', refs: ['A1:F1', 'A3:C5', 'A7:C7'], centerAcross: false },
      { kind: 'cells', sheet: 'Q', reason: 'edit', writes: [
        { row: 2, col: 1, cell: { t: 's', v: '00123', quote: true } }, { row: 3, col: 2, cell: { t: 's', v: 'x' } }, { row: 1, col: 4, cell: { t: 's', v: '00123', numFmt: '@', quote: true } },
      ] },
      { kind: 'struct', sheet: 'Tb', op: { type: 'insert', axis: 'col', at: 3, count: 2 } },
      { kind: 'struct', sheet: 'Tb', op: { type: 'move', axis: 'col', from: 2, count: 1, to: 7 } },
    ]));
    const variants: Array<[string, 'plain' | 'table', Record<string, unknown>]> = [
      ['pm-plain', 'plain', {}], ['pm-table', 'table', {}], ['pm-plain-values', 'plain', { formulaPolicy: 'values' }],
      ['pm-table-keep', 'table', { headerConflict: 'keep' }], ['pm-plain-noext', 'plain', { extendTable: false }], ['pm-table-noext', 'table', { extendTable: false }],
    ];
    for (const [name, v, o] of variants) {
      const r = await buildPremerged(scenario(v, o));
      fs.writeFileSync(`${outDir}/${name}.xlsx`, r.bytes);
      // the premerged package after typical user edits (moves / insert / delete / width / cell edit)
      const edited = await replayLog(r.bytes, [
        // (without the table extension the imported columns are outside the table: keep the move outside it)
        { kind: 'struct', sheet: '学生', op: { type: 'move', axis: 'col', from: 9, count: 2, to: name.includes('noext') ? 6 : 3 } },
        { kind: 'struct', sheet: '学生', op: { type: 'insert', axis: 'col', at: 6, count: 1 } },
        { kind: 'struct', sheet: '学生', op: { type: 'delete', axis: 'col', at: 12, count: 1 } },
        { kind: 'colWidth', sheet: '学生', col: 2, width: 15 },
        { kind: 'cells', sheet: '学生', reason: 'edit', writes: [{ row: 3, col: 6, cell: { t: 's', v: 'メモ' } }] },
      ]);
      fs.writeFileSync(`${outDir}/${name}-edited.xlsx`, edited);
      // round-2 ops on the premerged package: header rename / typed date with number format / dynamic array / an
      // unmerge that matches nothing (the merged sheet of these scenarios has no merges)
      const edited2 = await replayLog(r.bytes, [
        { kind: 'cells', sheet: '学生', reason: 'edit', writes: [
          { row: 1, col: 6, cell: { t: 's', v: '志望（改）' } },
          { row: 3, col: 13, cell: { t: 'd', v: 46295, numFmt: 'yyyy/m/d' } },
          { row: 12, col: 2, cell: { t: 'z', v: null, f: '_xlfn.SEQUENCE(3)', dyn: true } },
        ] },
        { kind: 'unmerge', sheet: '学生', refs: ['A1:B1'], centerAcross: true },
      ]);
      fs.writeFileSync(`${outDir}/${name}-r2edit.xlsx`, edited2);
    }
    const { insideTableScenario, richScenario, keyFillScenario } = await import('../premerge-scenario');
    fs.writeFileSync(`${outDir}/pm-inside-table.xlsx`, (await buildPremerged(insideTableScenario())).bytes);
    // round 2: rich text imports, combined / split key cells
    fs.writeFileSync(`${outDir}/pm-rich.xlsx`, (await buildPremerged(richScenario())).bytes);
    fs.writeFileSync(`${outDir}/pm-keyfill.xlsx`, (await buildPremerged(keyFillScenario())).bytes);
    console.log('written');
  } else if (cmd === 'snap') {
    console.log(JSON.stringify(await snapshot(new Uint8Array(fs.readFileSync(args[0]))), null, 1));
  } else {
    console.log('usage: replay|diff|snap');
  }
}

main().catch((e) => { console.error(e); process.exit(1); });
