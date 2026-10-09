/**
 * Join 'full' appends the rows that only exist in the other files below the table. A formula below the table whose range ends exactly at
 * the base's last data row (a total row: =SUM(D4:D33)) must grow over the appended rows ("range-extension rule" of buildPremerged), and
 * every other reference of the base keeps pointing at the same cells (rows below the insertion point move down). Checked on the premerged
 * sheet without Excel: the formula texts of the base book (readBook) against the ones of the premerged book.
 */
import { tokenize } from '../../../src/core/formula/excelRef';
import type { Finding } from './check';
import type { Merged } from './lib';

export function checkAppendedRows(combo: string, m: Merged): Finding[] {
  const out: Finding[] = [];
  const appended = m.result.layout.appended.length;
  if (!appended) return out;
  const baseSheet = m.base.book.sheets.find((s) => s.name === m.base.region.sheet);
  const mergedSheet = m.book.sheets.find((s) => s.name === m.result.layout.sheet);
  if (!baseSheet || !mergedSheet) return out;
  const r = m.base.region;
  const lastBase = r.lastDataRow;
  const lastMerged = m.region.lastDataRow;
  let checked = 0;
  for (let ri = 0; ri < baseSheet.rows.length; ri++) {
    const row = baseSheet.rows[ri];
    if (!row) continue;
    // rows inside the table keep their number, rows below it move down by the number of appended rows
    const target = ri + 1 > lastBase ? ri + appended : ri;
    row.forEach((cell, ci) => {
      if (!cell || cell.f === undefined) return;
      const mc = mergedSheet.rows[target]?.[ci];
      if (!mc || mc.f === undefined) {
        // single-sheet output: a formula that read a removed sheet (or a name pointing at one) is a value now
        if (m.input.options.dropOtherSheets) return;
        out.push({ scenario: combo, cls: 'append-range', detail: `${baseSheet.name}!R${ri + 1}C${ci + 1} (=${cell.f}) has no formula at R${target + 1} after the rows were appended` });
        return;
      }
      const tb = tokenize(cell.f).filter((t) => t.type === 'ref' && t.area && !t.prefix);
      const tm = tokenize(mc.f).filter((t) => t.type === 'ref' && t.area && !t.prefix);
      if (tb.length !== tm.length) {
        out.push({ scenario: combo, cls: 'append-range', detail: `R${ri + 1}C${ci + 1}: ${cell.f} -> ${mc.f} (different number of references)` });
        return;
      }
      tb.forEach((t, i) => {
        const a = t.area!;
        const b = tm[i].area!;
        if (a.kind !== 'area' && a.kind !== 'cell') return;
        checked++;
        // a range that ends at the old last data row and starts at / above the first data row grows over the new rows
        const inRange = a.kind === 'area' && Math.min(a.r1, a.r2) <= r.firstDataRow && Math.max(a.r1, a.r2) === lastBase;
        // a running total in the last data row (=SUM($D$2:D31) in row 31) ends at ITS OWN row: "up to here", not "the whole table"
        const running = inRange && !a.ar2 && ri + 1 === lastBase;
        const grows = inRange && !running;
        const wantR2 = grows ? lastMerged : a.r2 > lastBase ? a.r2 + appended : a.r2;
        const wantR1 = a.r1 > lastBase ? a.r1 + appended : a.r1;
        if (b.r2 !== wantR2 || b.r1 !== wantR1) out.push({ scenario: combo, cls: running ? 'append-range-running' : 'append-range', detail: `R${ri + 1}C${ci + 1}: ${cell.f} -> ${mc.f}: expected rows ${wantR1}..${wantR2} for the reference ${t.text}, got ${b.r1}..${b.r2}` });
      });
    });
  }
  void checked;
  return out;
}
