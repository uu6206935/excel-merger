/**
 * The two engines' outputs as PACKAGES (no recalculation): everything that is not a cell value / formula text (those are compared in
 * Excel, check.ts judgeParity): column properties (width, hidden, outline), names, merges, tables, autofilter, freeze, hyperlinks,
 * conditional formats and data validations (compared by the cells they cover, not by how Excel fragments them: tests/com/lib.ts
 * packageDiff). Known-benign differences are filtered and named:
 *   - cf-quirk        a CF rule with an absolute column reference ($C2) covers other cells in Excel after a move inside its range (documented)
 *   - richData        Excel writes xl/richData parts when it saves a dynamic-array workbook
 *   - calculated columns of tables (Excel infers <calculatedColumnFormula> when it saves; no effect on values)
 */
import { comparePackages } from '../../../tools/lib/package-diff.mjs';
import { packageDiff } from '../../com/lib';
import type { Finding, ScenarioRec } from './check';

export async function packageFindings(rec: ScenarioRec, js: Uint8Array, xl: Uint8Array): Promise<Finding[]> {
  const out: Finding[] = [];
  const isPivot = (sheet: string | undefined) => !!sheet && rec.pivotSheets.includes(sheet);
  const add = (cls: string, sheet: string | undefined, detail: string) => {
    if (isPivot(sheet) && !rec.pivotComparable) return;
    out.push({ scenario: rec.id, cls: `${isPivot(sheet) ? 'pivot-' : ''}${cls}`, detail: `${sheet ?? ''}: ${detail}`.slice(0, 420) });
  };
  const cmp = await comparePackages(js, xl);
  // groups the Excel dumps judge better (values, formulas), the styles (not the topic), and the groups replaced below
  const skip = new Set(['styles', 'formulas', 'values', 'constants', 'cf', 'dv', 'tables']);
  const rows = (cmp.rows as Array<{ group: string; ok: boolean; sheet?: string; detail: string }>).filter((x) => !x.ok && !skip.has(x.group));
  for (const r of rows.slice(0, 8)) {
    if (r.group === 'parts' && /リッチデータ/.test(r.detail)) continue;
    add(`engine-package-${r.group}`, r.sheet, r.detail);
  }
  // CF / DV / tables / merges / autofilter / names / stored widths by coverage
  for (const line of (await packageDiff(js, xl, 40)).slice(0, 8)) {
    const sheet = line.split(/[ :]/)[0];
    // Excel's documented CF quirk (tests/formula/cfProbe.test.ts KNOWN abs-ref-to-moved-column, tests/golden/RESULTS.md CF_INSIDE_MOVE): a rule
    // whose formula holds an ABSOLUTE column reference ($C2) loses the cells left of the moved column when that column moves inside its range.
    // Only the cells a conditional format covers differ, never a value.
    if (/ rule cf /.test(line) && /\$[A-Z]{1,3}\d/.test(line)) {
      add('cf-quirk', sheet, line);
      continue;
    }
    const kind = /rule/.test(line) ? 'rules' : /tables:/.test(line) ? 'tables' : /merges/.test(line) ? 'merges' : /autoFilter/.test(line) ? 'autofilter' : /stored width/.test(line) ? 'widths' : /^name only/.test(line) ? 'names' : 'other';
    add(`engine-package-${kind}`, sheet, line);
  }
  return out;
}
