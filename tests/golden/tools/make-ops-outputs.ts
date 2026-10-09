/**
 * Writes the outputs of random op logs (tests/golden/fuzz.ts randomLog) into a folder so that real Excel can be asked
 * whether they open without repair — the structural oracle the XML checks cannot replace (a file can be well-formed and
 * still be "repaired" by Excel, e.g. an array range over a cell that holds something else):
 *   npx tsx tests/golden/tools/make-ops-outputs.ts <folder> [count=120] [firstSeed=0] [plans=0]   (plans: random premerge plans too)
 *   powershell -NoProfile -ExecutionPolicy Bypass -File tests\golden\com\opentest.ps1 -Dir <folder>
 * Inputs: the Excel-authored golden fixtures and the premerged packages of the premerge scenarios (the real app flow).
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { buildPremerged } from '../../../src/core/xlsx/premerge';
import { replayLog } from '../../../src/core/xlsx/replay';
import { randomLog, randomPlan, replayTargets, sheetNames } from '../fuzz';
import { rng } from '../package-tools';
import { insideTableScenario, keyFillScenario, richScenario, scenario } from '../premerge-scenario';

const here = path.dirname(fileURLToPath(import.meta.url));
const FIX = path.resolve(here, '../../fixtures/golden');
const out = path.resolve(process.argv[2] ?? 'ops-out');
const count = Number(process.argv[3] ?? 120);
const first = Number(process.argv[4] ?? 0);
fs.mkdirSync(out, { recursive: true });

const inputs: Array<{ name: string; bytes: Uint8Array }> = [];
for (const t of replayTargets()) inputs.push({ name: t.file.replace(/\.xlsx$/, ''), bytes: new Uint8Array(fs.readFileSync(path.join(FIX, t.file))) });
for (const [name, input] of [['pm-plain', scenario('plain')], ['pm-table', scenario('table')], ['pm-inside', insideTableScenario()], ['pm-rich', richScenario()], ['pm-keyfill', keyFillScenario()]] as const) {
  inputs.push({ name, bytes: (await buildPremerged(input)).bytes });
}
const sheets = new Map<string, string[]>();
for (const i of inputs) sheets.set(i.name, await sheetNames(i.bytes));

let made = 0, refused = 0;
for (let seed = first; made < count && seed < first + count * 6; seed++) {
  const inp = inputs[seed % inputs.length];
  const log = randomLog(rng(seed * 29 + 11), sheets.get(inp.name)!);
  try {
    fs.writeFileSync(path.join(out, `op-${seed}-${inp.name}.xlsx`), await replayLog(inp.bytes, log));
    fs.writeFileSync(path.join(out, `op-${seed}-${inp.name}.json`), JSON.stringify(log));
    made++;
  } catch { refused++; }
}
console.log(`wrote ${made} outputs to ${out} (${refused} op logs were refused)`);

// random premerge plans (fuzz.ts randomPlan): the premerged package itself and the same package after a random op log
const plans = Number(process.argv[5] ?? 0);
let pm = 0, pmRefused = 0;
for (let seed = first; pm < plans && seed < first + plans * 4; seed++) {
  try {
    const pre = await buildPremerged(randomPlan(seed));
    fs.writeFileSync(path.join(out, `pmr-${seed}.xlsx`), pre.bytes);
    const log = randomLog(rng(seed * 43 + 1), await sheetNames(pre.bytes));
    try { fs.writeFileSync(path.join(out, `pmo-${seed}.xlsx`), await replayLog(pre.bytes, log)); } catch { pmRefused++; }
    pm++;
  } catch { pmRefused++; }
}
if (plans) console.log(`wrote ${pm} premerged packages (and their op-log results) (${pmRefused} refused)`);
