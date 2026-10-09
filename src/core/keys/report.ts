/** Plain-text reporting helpers (debugging / test output only; the UI uses the DTOs). */
import { type ColumnProfile, TYPE_LABEL } from './detect';
import type { Candidates, JoinResult, PairStats, Tier } from './match';
import type { Table } from './table';

const pct = (x: number) => `${(x * 100).toFixed(1)}%`;

export function printProfile(t: Table, ps: ColumnProfile[], all = false): string {
  const lines: string[] = [];
  lines.push(`\n=== ${t.file} / ${t.sheet}  (header row ${t.headerRow}, ${t.rows.length} data rows)`);
  if (t.layoutHint) lines.push(`  ! ${t.layoutHint}`);
  const shown = ps.filter((p) => all || p.type !== 'unknown').sort((a, b) => b.keyScore - a.keyScore || b.typeScore - a.typeScore);
  for (const p of shown) {
    lines.push(
      `  [${p.col}] ${p.header.padEnd(10)} → ${(TYPE_LABEL[p.type] ?? p.type).padEnd(6)} score=${p.typeScore.toFixed(2)}(${p.confidence})` +
        ` keyScore=${p.keyScore.toFixed(2)} fill=${pct(p.fill)} uniq=${pct(p.uniqueness)} dupRows=${p.dupRows} blank=${p.blank}` +
        `${p.entity ? ` entity=${p.entity}` : ''}${p.notes.length ? `  notes: ${p.notes.join(' / ')}` : ''}`,
    );
  }
  const hidden = ps.length - shown.length;
  if (hidden) {
    const names = ps.filter((p) => p.type === 'unknown').map((p) => p.header);
    lines.push(`  (${hidden} columns not key-like: ${names.slice(0, 8).join(', ')}${hidden > 8 ? ' …' : ''})`);
  }
  return lines.join('\n');
}

export function fmtPair(s: PairStats): string {
  return (
    `${s.a.label} ⟷ ${s.b.label}: match(B→A)=${pct(s.matchRate)} (${s.matchedB}/${s.nonBlankB}) coverA=${pct(s.coverageA)} ` +
    `uniqA=${pct(s.uniqA)} uniqB=${pct(s.uniqB)} dupA=${s.dupRowsA} dupB=${s.dupRowsB} blankA=${s.blankA} blankB=${s.blankB} ` +
    `ambig=${s.ambiguousB} card=${s.cardinality} score=${s.score.toFixed(3)} → ${s.verdict}`
  );
}

export function printCandidates(c: Candidates, max = 5): string {
  const l: string[] = [];
  l.push('  -- single-column candidates (same type family):');
  for (const s of c.singles.slice(0, max)) l.push('    ' + fmtPair(s));
  if (!c.singles.length) l.push('    (none)');
  l.push('  -- discovered by value overlap (header-independent):');
  for (const s of c.discovered.slice(0, max)) l.push('    ' + fmtPair(s));
  if (!c.discovered.length) l.push('    (none)');
  l.push('  -- composite candidates:');
  for (const s of c.composites.slice(0, max)) l.push('    ' + fmtPair(s));
  if (!c.composites.length) l.push('    (none)');
  return l.join('\n');
}

export function printJoin(tA: Table, tB: Table, tiers: Tier[], r: JoinResult, show: (t: Table, i: number) => string): string {
  const l: string[] = [];
  l.push(`  -- plan: ${tiers.map((t, i) => `${i + 1}.${t.name}`).join(' → ') || '(no usable key)'}`);
  l.push(
    `  -- result: matched=${r.pairs.length} (${Object.entries(r.perTier).map(([k, v]) => `${k}:${v}`).join(', ')}) ` +
      `unmatchedA=${r.unmatchedA.length} unmatchedB=${r.unmatchedB.length} ambiguous=${r.ambiguous.length} conflicts=${r.conflicts.length}`,
  );
  if (r.dupA.length) l.push(`     duplicate keys in A: ${r.dupA.map(([k, n]) => `${k}×${n}`).join(', ')}`);
  if (r.dupB.length) l.push(`     duplicate keys in B: ${r.dupB.map(([k, n]) => `${k}×${n}`).join(', ')}`);
  for (const p of r.pairs) for (const f of p.flags) l.push(`     [${f.level}] A:${show(tA, p.a)} ⟷ B:${show(tB, p.b)} (${p.tier}) ${f.msg}`);
  for (const a of r.ambiguous) l.push(`     [ambiguous] A:${show(tA, a.a)} → B候補 ${a.candidatesB.map((j) => show(tB, j)).join(' | ')} (${a.tier})`);
  const seen = new Set<string>();
  for (const c of r.conflicts) {
    const k = `${c.a}|${c.b}`;
    if (seen.has(k)) continue;
    seen.add(k);
    l.push(`     [conflict] A:${show(tA, c.a)} ⟷ B:${show(tB, c.b)} ${c.msg}`);
  }
  if (r.unmatchedA.length) l.push(`     unmatched A: ${r.unmatchedA.slice(0, 12).map((i) => show(tA, i)).join(' | ')}${r.unmatchedA.length > 12 ? ' …' : ''}`);
  if (r.unmatchedB.length) l.push(`     unmatched B: ${r.unmatchedB.slice(0, 12).map((i) => show(tB, i)).join(' | ')}${r.unmatchedB.length > 12 ? ' …' : ''}`);
  for (const s of r.suggestions.slice(0, 12)) l.push(`     [suggest ${s.score.toFixed(2)}] A:${show(tA, s.a)} ⟷ B:${show(tB, s.b)}  ${s.reasons.join('・')}`);
  return l.join('\n');
}
