/**
 * Run the package fuzzers for many seeds and print what they found:
 *   npx tsx tests/golden/tools/fuzz-cli.ts replay|premerge|ops|plans <startSeed> <count> [--all]
 *   (replay = damaged packages + golden op logs, premerge = damaged base / source, ops = intact packages + random op logs,
 *    plans = intact packages + random premerge plans, then random op logs on the result)
 * Internal errors (an unforeseen exception that the patcher converted to a Japanese PatchError) and contract violations
 * are listed with the seed that reproduces them (mutate(bytes, seed) is deterministic).
 */
import { fuzzOps, fuzzPlans, fuzzPremerge, fuzzReplay, summarize } from '../fuzz';

const [mode = 'replay', start = '0', count = '500', flag] = process.argv.slice(2);
const t0 = Date.now();
const res = await (mode === 'premerge' ? fuzzPremerge(Number(start), Number(count)) : mode === 'ops' ? fuzzOps(Number(start), Number(count)) : mode === 'plans' ? fuzzPlans(Number(start), Number(count)) : fuzzReplay(Number(start), Number(count)));
const s = summarize(res);
console.log(`${mode}: ${res.length} mutants in ${((Date.now() - t0) / 1000).toFixed(1)} s: ok ${s.ok}, refused ${s.refused}, internal ${s.internal.length}, violations ${s.violations.length}`);
for (const x of s.violations) console.log(`VIOLATION seed=${x.seed} ${x.what}\n    ${x.detail}`);
for (const x of s.internal) console.log(`INTERNAL  seed=${x.seed} ${x.what}\n    ${(x.detail ?? '').slice(0, 300)}`);
if (flag === '--all') for (const x of res.filter((y) => y.kind === 'refused')) console.log(`refused   seed=${x.seed} ${x.what}\n    ${(x.detail ?? '').slice(0, 200)}`);
