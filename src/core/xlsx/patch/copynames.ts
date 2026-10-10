/**
 * Defined names of a source workbook for the sheets copied from it (addsheets.ts).
 *
 * A copied sheet gets its names as SHEET-scoped names, so nothing can collide with the names of the base workbook or of another copied
 * file: the sheet-scoped names of the source sheet itself (print area / titles, _FilterDatabase, the user's own), and every
 * workbook-scoped name that one of its formulas, rules or links (or another name it needs) uses. The text of a name has its sheet
 * references re-pointed at the renamed copies. A name that cannot be copied (macro / function names, references into another workbook,
 * tables or sheets that are not copied ...) is dropped; formulas that use it keep their cached values (FormulaCopier fails them with
 * 'name').
 */
import { escAttr, esc, findAll, findFirst, getAttr, unesc } from './xml';
import { FormulaCopier, namePattern, type CopyFailReason, type NameView } from './copyformula';
import type { Engine } from './engine';

export interface SrcName {
  key: string;
  name: string;
  lower: string;
  /** index of the source sheet the name belongs to, null = workbook scope */
  local: number | null;
  hidden: boolean;
  comment?: string;
  text: string;
  macro: boolean;
}

export class SourceNames {
  readonly all: SrcName[] = [];
  readonly global = new Map<string, SrcName>();
  readonly scoped = new Map<number, Map<string, SrcName>>();
  /** key -> why the name cannot be copied */
  readonly bad = new Map<string, CopyFailReason | 'macro'>();
  private readonly results = new Map<string, { text: string; used: string[] }>();
  private readonly views = new Map<number | null, NameView>();

  constructor(eng: Engine, readonly sheetLower: (index: number) => string | null) {
    const el = findFirst(eng.wbXml, 'definedNames');
    if (!el || el.selfClosing) return;
    for (const d of findAll(el.inner, 'definedName')) {
      const name = getAttr(d.open, 'name');
      if (!name || d.selfClosing) continue;
      const ls = getAttr(d.open, 'localSheetId');
      const local = ls === undefined || ls === '' ? null : Number(ls);
      if (local !== null && !Number.isInteger(local)) continue;
      const lower = name.toLowerCase();
      const key = local === null ? `g:${lower}` : `s${local}:${lower}`;
      const text = unesc(d.inner).trim();
      if (!text) continue;
      const n: SrcName = {
        key, name, lower, local, hidden: /\shidden="(1|true)"/.test(d.open), comment: getAttr(d.open, 'comment'), text,
        macro: /\s(function|vbProcedure|xlm)="(1|true)"/.test(d.open),
      };
      if (local === null) { if (!this.global.has(lower)) this.global.set(lower, n); else continue; }
      else {
        let m = this.scoped.get(local);
        if (!m) this.scoped.set(local, (m = new Map()));
        if (m.has(lower)) continue;
        m.set(lower, n);
      }
      this.all.push(n);
    }
  }

  /** the names a formula on source sheet `index` (null: a workbook-level name) sees */
  view(index: number | null): NameView {
    let v = this.views.get(index);
    if (v) return v;
    const sc = index === null ? undefined : this.scoped.get(index);
    const names = [...(sc?.keys() ?? []), ...this.global.keys()];
    v = {
      has: (l) => !!sc?.has(l) || this.global.has(l),
      bad: (l) => (sc?.has(l) ? this.bad.has(`s${index}:${l}`) : this.bad.has(`g:${l}`)),
      pattern: namePattern(names),
    };
    this.views.set(index, v);
    return v;
  }

  /** decide which names can be copied (the text of every name is rewritten once; names that use a bad name are bad themselves) */
  analyse(copier: FormulaCopier): void {
    for (const n of this.all) if (n.macro) this.bad.set(n.key, 'macro');
    for (let round = 0; round < 12; round++) {
      let changed = false;
      for (const n of this.all) {
        if (this.bad.has(n.key)) continue;
        const r = copier.rewrite(n.text, n.local === null ? null : this.sheetLower(n.local), this.view(n.local));
        if (!r.ok) { this.bad.set(n.key, r.reason); changed = true; continue; }
        this.results.set(n.key, { text: r.text, used: r.names });
      }
      if (!changed) break;
    }
  }

  /**
   * The <definedName> elements for the copy of source sheet `index` (its new position `newIndex`): its own sheet-scoped names and the
   * workbook names in `used` (and the names those need). `dropped`: names of the sheet that could not be copied.
   */
  emit(index: number, newIndex: number, used: Iterable<string>): { xml: string[]; dropped: string[] } {
    const xml: string[] = [];
    const dropped: string[] = [];
    const done = new Set<string>();
    const write = (n: SrcName) => {
      const res = this.results.get(n.key);
      if (!res || done.has(n.lower)) return;
      done.add(n.lower);
      xml.push(`<definedName name="${escAttr(n.name)}" localSheetId="${newIndex}"${n.hidden ? ' hidden="1"' : ''}${n.comment ? ` comment="${escAttr(n.comment)}"` : ''}>${esc(res.text)}</definedName>`);
    };
    const sc = this.scoped.get(index);
    for (const n of sc?.values() ?? []) {
      if (this.bad.has(n.key)) { if (!/^_xlnm\._FilterDatabase$/i.test(n.name)) dropped.push(n.name.replace(/^_xlnm\./, '')); continue; }
      write(n);
    }
    const queue = [...used];
    while (queue.length) {
      const l = queue.pop()!;
      if (sc?.has(l)) continue; // the sheet's own name hides the workbook name (already written, or dropped)
      const g = this.global.get(l);
      if (!g || done.has(g.lower)) continue;
      if (this.bad.has(g.key)) continue;
      write(g);
      queue.push(...(this.results.get(g.key)?.used ?? []));
    }
    return { xml, dropped };
  }
}
