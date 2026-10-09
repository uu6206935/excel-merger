/**
 * Round 5 — randomised messy lists against the ground truth (generator: messy-fuzz-lib.ts): whatever columns two files share
 * and however their values are written (prefixes, zeros, 全角, 様, 異体字, typos, blanks, resubmissions, 同姓同名 …) the default
 * plan never joins two different people SILENTLY: a wrong match may only happen where the result flags the row for checking
 * (2nd-tier key, name / birth-date conflict, duplicate-row rule). What is left unjoined is fine — it must show up as unmatched /
 * suggestion / ambiguous.
 *
 * Unavoidable "wrong" matches are kept out of the generator: a typo never produces another person's ID / name / address, and a
 * pair of namesakes is always in BOTH files or in neither (a namesake that exists in one file only cannot be told apart by anyone).
 */
import { describe, expect, it } from 'vitest';
import { analyzeSources, defaultPlan, joinSources } from '../../src/core/keys';
import { FLAGGED, fuzz, fuzzMulti } from './messy-fuzz-lib';

describe('randomised messy lists', () => {
  const N = 250;
  it(`${N} random pairs: the default plan never joins two different people (any join type / duplicate policy)`, () => {
    let matched = 0;
    let flaggedWrong = 0;
    let wrongTotal = 0;
    const failures: string[] = [];
    for (let seed = 1; seed <= N; seed++) {
      const f = fuzz(seed);
      const an = analyzeSources(f.base, [f.other]);
      const plan = { ...defaultPlan(an), joinType: f.joinType, duplicatePolicy: f.policy };
      const res = joinSources(an, plan);
      const wrong: string[] = [];
      for (const r of res.rows) {
        const m = r.matches.o;
        if (!m || r.baseRow === null) continue;
        matched++;
        const x = f.baseTruth.get(r.baseRow);
        const y = f.otherTruth.get(m.row);
        if (!x || x !== y) {
          // a wrong match that the result FLAGS (2nd-tier key, name conflict, duplicate rule …) is listed for checking;
          // a silent one is the failure
          if (m.flags.some((fl) => FLAGGED.has(fl))) flaggedWrong++;
          else wrong.push(`base ${r.baseRow} (${x}) ⟷ other ${m.row} (${y}) via ${m.via} flags=[${m.flags.join(',')}]`);
        }
      }
      // rows appended by 'full' are people of the other file that joined nothing: none of them may be a base row's partner twice
      if (wrong.length) {
        wrongTotal += wrong.length;
        failures.push(`${f.desc} [${f.joinType}/${f.policy}] plan=${an.pairs.o.candidates.filter((c) => c.selected).map((c) => c.label).join('→')}: ${wrong.join('; ')}`);
      }
    }
    expect(failures).toEqual([]);
    expect(wrongTotal).toBe(0);
    expect(matched).toBeGreaterThan(1000); // the generator produces real joins, not only non-matches
    console.log(`fuzz: ${matched} automatic matches, ${flaggedWrong} of them between different people but FLAGGED for checking`);
  });

  it('the same pair analysed with every normalisation option off never joins two different people either', () => {
    const failures: string[] = [];
    for (let seed = 1; seed <= 60; seed++) {
      const f = fuzz(seed);
      const an = analyzeSources(f.base, [f.other], { itaiji: false, leadingZeros: false, honorifics: false, maidenName: false, universityAbbrev: false });
      const res = joinSources(an, defaultPlan(an));
      for (const r of res.rows) {
        const m = r.matches.o;
        if (!m || r.baseRow === null) continue;
        if (f.baseTruth.get(r.baseRow) !== f.otherTruth.get(m.row) && !m.flags.some((fl) => FLAGGED.has(fl))) failures.push(`${f.desc}: base ${r.baseRow} ⟷ other ${m.row} flags=[${m.flags.join(',')}]`);
      }
    }
    expect(failures).toEqual([]);
  });

  it('150 random triples / quadruples with "full" joins: rows appended from one file are matched by the next one without mixing people up', () => {
    const failures: string[] = [];
    let cross = 0;
    for (let seed = 1; seed <= 150; seed++) {
      const f = fuzzMulti(seed);
      const an = analyzeSources(f.base, f.others);
      const res = joinSources(an, { ...defaultPlan(an), joinType: 'full', duplicatePolicy: f.policy });
      for (const r of res.rows) {
        // who is this merged row? a base row, or a row appended from one of the files
        const person = r.baseRow !== null ? f.baseTruth.get(r.baseRow) : f.otherTruth[r.appendedFrom!.sourceId].get(r.appendedFrom!.row);
        for (const o of f.others) {
          const m = r.matches[o.id];
          if (!m) continue;
          if (r.baseRow === null && r.appendedFrom?.sourceId !== o.id) cross++;
          if (r.baseRow === null && r.appendedFrom?.sourceId === o.id) continue; // the row's own source
          const other = f.otherTruth[o.id].get(m.row);
          if (!person || person !== other) {
            if (!m.flags.some((fl) => FLAGGED.has(fl))) failures.push(`${f.desc}: row ${r.baseRow ?? 'appended'} (${person}) ⟷ ${o.id} row ${m.row} (${other}) via ${m.via} flags=[${m.flags.join(',')}]`);
          }
        }
      }
    }
    expect(failures).toEqual([]);
    expect(cross).toBeGreaterThan(20); // later files did find rows appended by earlier ones
  });
});
