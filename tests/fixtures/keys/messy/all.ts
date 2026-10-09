/** Every messy case of Round 5, in one list (names are unique). */
import type { MessyCase } from './lib';
import { cases as ids } from './make-ids';
import { cases as layout } from './make-layout';
import { cases as names } from './make-names';
import { cases as relations } from './make-relations';

export const allCases: MessyCase[] = [...names, ...ids, ...layout, ...relations];

const seen = new Set<string>();
for (const c of allCases) {
  if (seen.has(c.name)) throw new Error(`duplicate messy case name: ${c.name}`);
  seen.add(c.name);
}
