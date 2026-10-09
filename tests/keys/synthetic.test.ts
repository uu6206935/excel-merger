/**
 * Synthetic "tricky" set from research R6 (tests/fixtures/keys/*.xlsx + truth.json) evaluated through the facade:
 * 0 wrong automatic matches, every missed true pair surfaced as a suggestion or an ambiguous choice.
 */
import fs from 'node:fs';
import path from 'node:path';
import { describe, expect, it } from 'vitest';
import type { KeySource, JoinPlan, JoinResult } from '../../src/core/contracts';
import { analyzeSources, defaultPlan, joinSources } from '../../src/core/keys';
import { buildTable } from '../../src/core/keys/region';
import { analyzeKeys, runJoin } from '../../src/core/keys/api';
import { FIXTURES, bookFromFile, sourceOf } from './helpers';

const truth = JSON.parse(fs.readFileSync(path.join(FIXTURES, 'truth.json'), 'utf8')) as { persons: string[]; rows: Record<string, number[]> };
const FILES: Record<string, string> = { A: 'A_応募者名簿.xlsx', B: 'B_一次面接結果.xlsx', C: 'C_学籍名簿.xlsx', D: 'D_謎の表.xlsx' };
const src: Record<string, KeySource> = {};
for (const [k, f] of Object.entries(FILES)) src[k] = sourceOf(bookFromFile(path.join(FIXTURES, f), k), k, k);
/** Excel row → truth person index */
function personOf(k: string): Map<number, number> {
  const t = buildTable(src[k]);
  const m = new Map<number, number>();
  t.excelRow.forEach((r, i) => m.set(r, truth.rows[k][i]));
  return m;
}

interface Eval {
  truePairs: number;
  correct: number;
  wrong: string[];
  missed: string[];
  missedSurfaced: number;
  wrongSuggestions: string[];
  res: JoinResult;
}

function evaluate(a: string, b: string, patch: Partial<JoinPlan> = {}): Eval {
  const an = analyzeSources(src[a], [src[b]]);
  const plan = { ...defaultPlan(an), ...patch };
  const res = joinSources(an, plan);
  const pa = personOf(a);
  const pb = personOf(b);
  let correct = 0;
  const wrong: string[] = [];
  const matched = new Set<string>();
  for (const row of res.rows) {
    const m = row.matches[b];
    if (row.baseRow === null || !m) continue;
    const x = pa.get(row.baseRow);
    const y = pb.get(m.row);
    if (x !== undefined && x >= 0 && x === y) {
      correct++;
      matched.add(`${row.baseRow}|${m.row}`);
    } else wrong.push(`${a}${row.baseRow} ⟷ ${b}${m.row} via ${m.via}`);
  }
  const truePairs: [number, number][] = [];
  for (const [ra, x] of pa) for (const [rb, y] of pb) if (x >= 0 && x === y) truePairs.push([ra, rb]);
  const rep = res.reports[b];
  const sugg = new Set(rep.suggestions.map((s) => `${s.base.excelRow}|${s.other.excelRow}`));
  const ambig = new Set(rep.ambiguous.flatMap((x) => x.candidates.map((c) => `${x.base.excelRow}|${c.excelRow}`)));
  // a person whose duplicate submission was resolved by the policy counts once
  const matchedPersons = new Set([...matched].map((k) => pa.get(Number(k.split('|')[0]))));
  const missed = truePairs.filter(([x, y]) => !matched.has(`${x}|${y}`) && !matchedPersons.has(pa.get(x)));
  const missedSurfaced = missed.filter(([x, y]) => sugg.has(`${x}|${y}`) || ambig.has(`${x}|${y}`)).length;
  const wrongSuggestions = rep.suggestions.filter((s) => pa.get(s.base.excelRow) !== pb.get(s.other.excelRow)).map((s) => `${s.base.preview.join(',')} ⟷ ${s.other.preview.join(',')} ${s.score}`);
  return { truePairs: new Set(truePairs.map(([x]) => pa.get(x))).size, correct, wrong, missed: missed.map(([x, y]) => `${x}|${y}`), missedSurfaced, wrongSuggestions, res };
}

describe('synthetic truth (R6 set)', () => {
  it('A ⟷ B with 重複=結合しない reproduces R6: 24 correct, 0 wrong, 5 missed all surfaced', () => {
    const e = evaluate('A', 'B', { duplicatePolicy: 'skip' });
    expect(e.wrong).toEqual([]);
    expect(e.correct).toBe(24);
    expect(e.missed.length).toBe(e.missedSurfaced);
    expect(e.wrongSuggestions).toEqual([]);
    const rep = e.res.reports.B;
    // Round 5: メール / 電話番号 are tiers of their own (unique on ≥ 90% of the rows), tried after the composite keys
    expect(rep.plan.map((p) => p.name)).toEqual(['ID', '氏名＋大学', '氏名カナ＋大学', 'メール', '電話番号', '氏名', '氏名カナ']);
    expect(rep.ambiguous.length).toBe(1); // 小林 大輝 submitted twice
  });

  it('A ⟷ B with the default policy (先頭の行) also resolves the duplicate submission', () => {
    const e = evaluate('A', 'B');
    expect(e.wrong).toEqual([]);
    expect(e.correct).toBe(25);
    expect(e.missed.length).toBe(e.missedSurfaced);
    expect(e.wrongSuggestions).toEqual([]);
    const kobayashi = e.res.rows.find((r) => r.matches.B?.flags.includes('duplicateInSource'));
    expect(kobayashi?.baseRow).toBe(16); // A row 16 = 000112 小林 大輝
    // flags: ID一致・氏名不一致 (中村/木村) is a conflict, 異体字 matches are variants, 氏名＋大学 matches are fallbacks
    const byRow = new Map(e.res.rows.map((r) => [r.baseRow, r.matches.B]));
    expect(byRow.get(17)?.flags).toContain('conflict'); // 000113 中村 優子 ⟷ 木村 拓也
    expect(byRow.get(7)?.flags).toContain('variant'); // 000103 髙橋 一郎 ⟷ 高橋一郎
    const viaComposite = e.res.rows.filter((r) => r.matches.B?.via === '氏名＋大学');
    expect(viaComposite.length).toBe(2);
    for (const r of viaComposite) expect(r.matches.B?.flags).toContain('fallback');
    const items = e.res.items.map((i) => i.code);
    expect(items).toEqual(expect.arrayContaining(['unmatched', 'duplicate', 'conflict', 'fallback', 'variant', 'suggestion']));
    const conflict = e.res.items.find((i) => i.code === 'conflict');
    expect(conflict?.severity).toBe('error');
    expect(conflict?.rows).toEqual([17]);
    expect(conflict?.message).toContain('ID一致・氏名不一致');
  });

  it('A ⟷ C: 同姓同名 resolved by 氏名＋大学 (姓＋名 virtual column), 10 / 10', () => {
    const e = evaluate('A', 'C');
    expect(e.wrong).toEqual([]);
    expect(e.correct).toBe(10);
    expect(e.res.reports.C.plan[0].name).toBe('氏名＋大学');
  });

  it('A ⟷ D: header-less ID column typed from values, 20 / 20', () => {
    const e = evaluate('A', 'D');
    expect(e.wrong).toEqual([]);
    expect(e.correct).toBe(20);
    expect(e.res.reports.D.plan.map((p) => p.name)).toEqual(['ID']);
  });

  it('B ⟷ C: different ID systems ignored, 9 / 10 + 斎藤/斉藤 suggested', () => {
    const e = evaluate('B', 'C');
    expect(e.wrong).toEqual([]);
    expect(e.correct).toBe(9);
    expect(e.missed.length).toBe(1);
    expect(e.missedSurfaced).toBe(1);
    expect(e.wrongSuggestions).toEqual([]);
  });

  it('R6 two-table API: flags aggregated by kind (R6 §9.3 table)', () => {
    const an = analyzeKeys(buildTable(src.A), buildTable(src.B));
    expect(an.mode).toBe('oneToOne');
    const r = runJoin(an, { policy: 'skip' });
    expect(r.matched).toBe(24);
    const count = (k: string) => r.flags[k]?.count ?? 0;
    // R6 counted 20 「ID表記の違いを補正」; the 000000-formatted number 115 now also shows as 000115 (vs '115')
    expect(count('ID表記の違いを補正')).toBe(21);
    // R6 skipped the name check on 氏名-tier matches; 異体字 matches there are now reported too
    expect(count('異体字の違い')).toBeGreaterThanOrEqual(4);
    expect(count('似た字の違い')).toBe(1);
    expect(count('ID一致・氏名不一致')).toBe(1);
    expect(count('フリガナ不一致')).toBe(1);
    expect(count('姓名の順序違い')).toBe(1);
    expect(r.flags['ID一致・氏名不一致'].level).toBe('error');
    expect(r.suggestions.every((s) => s.score >= 0.6)).toBe(true);
  });

  it('A ⟷ B ⟷ C ⟷ D at once (per-pair plans are independent)', () => {
    const an = analyzeSources(src.A, [src.B, src.C, src.D]);
    const res = joinSources(an, defaultPlan(an));
    const pa = personOf('A');
    for (const k of ['B', 'C', 'D']) {
      const pk = personOf(k);
      for (const r of res.rows) {
        const m = r.matches[k];
        if (m && r.baseRow !== null) expect(pk.get(m.row)).toBe(pa.get(r.baseRow));
      }
    }
    expect(res.summary.baseRows).toBe(30);
    // everyone but 一ノ瀬 葵 (only in B, spelled 一之瀬, no ID → suggestion) is matched by some file
    expect(res.summary.matchedAny).toBe(29);
    expect(res.summary.matchedAll).toBe(0 + res.rows.filter((r) => r.baseRow !== null && ['B', 'C', 'D'].every((k) => r.matches[k])).length);
    expect(res.summary.matchedAll).toBeGreaterThan(0);
  });
});
