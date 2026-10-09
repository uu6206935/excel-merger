/**
 * Conditional formatting / data validation on the sheet an op is performed on (owner: patcher P).
 * Units: one <conditionalFormatting> block (its cfRules share the sqref), one <dataValidation>, one
 * <x14:conditionalFormatting>, one <x14:dataValidation>. Formulas ("slots": <formula>, <formula1/2>, <xm:f>,
 * cfvo val) are anchored at the top-left of the sqref bounding box (xlsx convention). excelRef.adjustAnchoredRule
 * does all the Excel math (fragments, format inheritance on insert, the move "leak", per-cell splitting); this file
 * runs it per cfRule (the leak depends on the rule type), regroups rules that end up with the same sqref, adds the
 * premerge row extension for validations, and re-renders the XML.
 */
import { adjustAnchoredRule, type AnchoredRule, type Area, type StructOp } from '../../formula/excelRef';
import { parseSqref, renderSqref } from './formula';
import { esc, findAll, getAttr, setAttr, unesc } from './xml';

export type UnitKind = 'cf' | 'dv' | 'x14cf' | 'x14dv';

export interface GrowRule {
  /** Excel "format same as left/above": CF areas ending right before an insert grow, moved blocks pick up rules */
  leak: boolean;
  /** premerge row append: DV areas spanning [<= firstDataRow, lastDataRow] grow over the appended rows */
  ext?: { firstDataRow: number; lastDataRow: number };
}

/** Formula slots of a unit / rule (document order). */
export function getSlots(xml: string, kind: UnitKind): string[] {
  const out: string[] = [];
  if (kind === 'cf') {
    for (const m of xml.matchAll(/<formula>([\s\S]*?)<\/formula>|<cfvo\b[^>]*>/g)) {
      if (m[1] !== undefined) out.push(unesc(m[1]));
      else { const v = getAttr(m[0], 'val'); if (v !== undefined) out.push(v); }
    }
  } else if (kind === 'dv') {
    for (const m of xml.matchAll(/<(formula1|formula2)>([\s\S]*?)<\/\1>/g)) out.push(unesc(m[2]));
  } else {
    for (const m of xml.matchAll(/<xm:f>([\s\S]*?)<\/xm:f>|<x14:cfvo\b[^>]*>/g)) {
      if (m[1] !== undefined) out.push(unesc(m[1]));
      else { const v = getAttr(m[0], 'val'); if (v !== undefined) out.push(v); }
    }
  }
  return out;
}

export function setSlots(xml: string, kind: UnitKind, vals: string[]): string {
  let i = 0;
  if (kind === 'cf') {
    return xml.replace(/<formula>([\s\S]*?)<\/formula>|<cfvo\b[^>]*>/g, (m0, f) => {
      if (f !== undefined) return `<formula>${esc(vals[i++])}</formula>`;
      if (getAttr(m0, 'val') === undefined) return m0;
      return setAttr(m0, 'val', vals[i++]);
    });
  }
  if (kind === 'dv') return xml.replace(/<(formula1|formula2)>([\s\S]*?)<\/\1>/g, (_m, tag) => `<${tag}>${esc(vals[i++])}</${tag}>`);
  return xml.replace(/<xm:f>([\s\S]*?)<\/xm:f>|<x14:cfvo\b[^>]*>/g, (m0, f) => {
    if (f !== undefined) return `<xm:f>${esc(vals[i++])}</xm:f>`;
    if (getAttr(m0, 'val') === undefined) return m0;
    return setAttr(m0, 'val', vals[i++]);
  });
}

export function getSqref(xml: string, kind: UnitKind): string {
  if (kind === 'cf' || kind === 'dv') return getAttr(/^<[^>]*>/.exec(xml)?.[0] ?? '', 'sqref') ?? '';
  const m = /<xm:sqref>([\s\S]*?)<\/xm:sqref>/.exec(xml);
  return m ? unesc(m[1]) : '';
}

export function setSqref(xml: string, kind: UnitKind, sqref: string): string {
  if (kind === 'cf' || kind === 'dv') return xml.replace(/^<[^>]*>/, (open) => setAttr(open, 'sqref', sqref));
  return xml.replace(/<xm:sqref>[\s\S]*?<\/xm:sqref>/, `<xm:sqref>${esc(sqref)}</xm:sqref>`);
}

const sameAreas = (a: Area[], b: Area[]) => a.length === b.length && a.every((x, i) => x.c1 === b[i].c1 && x.c2 === b[i].c2 && x.r1 === b[i].r1 && x.r2 === b[i].r2);

/** premerge: validation areas spanning [<= first, last] grow over rows inserted right after `last`. */
function extGrow(rules: AnchoredRule[], op: StructOp, g: GrowRule): AnchoredRule[] {
  const e = g.ext;
  if (!e || op.type !== 'insert' || op.axis !== 'row' || op.at !== e.lastDataRow + 1) return rules;
  return rules.map((r) => ({
    formulas: r.formulas,
    sqref: r.sqref.map((a) => {
      const lo = Math.min(a.r1, a.r2), hi = Math.max(a.r1, a.r2);
      if (a.kind !== 'area' && a.kind !== 'cell') return a;
      if (!(lo <= e.firstDataRow && hi === e.lastDataRow)) return a;
      return { ...a, kind: 'area' as const, r1: lo, r2: hi + op.count };
    }),
  }));
}

export interface UnitResult {
  /** new XML for this unit ('' = removed) */
  xml: string;
  /** number of copies produced (> 1 = split) */
  copies: number;
  changed: boolean;
}

/**
 * The move "leak" (spread) depends on the rule type: Excel does NOT extend colorScale, dataBar, iconSet, top10,
 * uniqueValues, duplicateValues, aboveAverage onto a moved block that lands right after / inside their area; cellIs,
 * expression, containsText, containsBlanks (and validations) spread (tests/golden/com/probe-cf-leak.ps1). Every type
 * grows over an inserted column. excelRef.adjustAnchoredRule takes the xlsx <cfRule type> / <dataValidation type>
 * as `type` and applies that table (NO_SPREAD_CF_TYPES).
 */

/** One CF block (classic or x14): every rule adjusted on its own, rules regrouped by resulting sqref. */
function adjustCfBlock(xml: string, x14: boolean, op: StructOp, sheet: string, g: GrowRule): UnitResult {
  const kind: UnitKind = x14 ? 'x14cf' : 'cf';
  const tag = x14 ? 'x14:conditionalFormatting' : 'conditionalFormatting';
  const ruleTag = x14 ? 'x14:cfRule' : 'cfRule';
  const open = /^<[^>]*>/.exec(xml)![0];
  if (open.endsWith('/>')) return { xml, copies: 1, changed: false };
  const inner = xml.slice(open.length, xml.length - `</${tag}>`.length);
  const rules = findAll(inner, ruleTag);
  const areas = parseSqref(getSqref(xml, kind));
  if (!areas.length || !rules.length) return { xml, copies: 1, changed: false };
  // what is left besides the rules (x14: <xm:sqref>; classic: an extLst)
  let extras = inner;
  for (let i = rules.length - 1; i >= 0; i--) extras = extras.slice(0, rules[i].start) + extras.slice(rules[i].end);
  if (x14) extras = extras.replace(/<xm:sqref>[\s\S]*?<\/xm:sqref>/, '');
  const groups = new Map<string, { areas: Area[]; rules: string[] }>();
  let changed = false;
  // Excel keeps ONE range object per <conditionalFormatting> block: every rule of the block is fragmented at the same
  // places, and on a move the cells that joined the block by format inheritance leave it again when ANY rule of the
  // block refers to the moved cells / the destination -> each rule sees the formulas of the others (AnchoredRuleOptions.siblings)
  const allSlots = rules.map((r) => getSlots(r.text, kind));
  for (const [ri, r] of rules.entries()) {
    const slots = allSlots[ri];
    const siblings = allSlots.filter((_s, i) => i !== ri);
    const res = adjustAnchoredRule({ sqref: areas, formulas: slots }, op, sheet, { kind: 'cf', type: getAttr(r.open, 'type'), inheritFormat: g.leak, siblings });
    if (res.length !== 1 || !sameAreas(res[0].sqref, areas) || res[0].formulas.some((f, i) => f !== slots[i])) changed = true;
    res.forEach((ar, k) => {
      if (!ar.sqref.length) return;
      const sq = renderSqref(ar.sqref);
      let rx = setSlots(r.text, kind, ar.formulas);
      if (k > 0) rx = rx.replace(/<extLst>[\s\S]*?<\/extLst>/g, '').replace(/\sid="\{[^"]*\}"/, ''); // copies must not reuse x14 links
      const grp = groups.get(sq) ?? { areas: ar.sqref, rules: [] };
      grp.rules.push(rx);
      groups.set(sq, grp);
    });
  }
  if (!changed) return { xml, copies: 1, changed: false };
  if (!groups.size) return { xml: '', copies: 0, changed: true };
  let first = true;
  const parts: string[] = [];
  for (const [sq, grp] of groups) {
    const ex = first ? extras : x14 ? '' : extras.replace(/<extLst>[\s\S]*?<\/extLst>/g, '');
    parts.push(x14
      ? `${open}${grp.rules.join('')}${ex}<xm:sqref>${esc(sq)}</xm:sqref></${tag}>`
      : `${setAttr(open, 'sqref', sq)}${grp.rules.join('')}${ex}</${tag}>`);
    first = false;
  }
  const copies = [...groups.values()].reduce((n, gr) => n + gr.rules.length, 0);
  return { xml: parts.join(''), copies: copies > rules.length ? copies : 1, changed: true };
}

/** One data validation (classic or x14). */
function adjustDv(xml: string, x14: boolean, op: StructOp, sheet: string, g: GrowRule): UnitResult {
  const kind: UnitKind = x14 ? 'x14dv' : 'dv';
  const areas = parseSqref(getSqref(xml, kind));
  if (!areas.length) return { xml, copies: 1, changed: false };
  const slots = getSlots(xml, kind);
  // Excel's inserted / moved cells inherit validation exactly like CF membership (goldens g-ins-D, g-ins-row14,
  // t-mov-L-H), so validations use the same format-inheritance rules as CF
  let res = adjustAnchoredRule({ sqref: areas, formulas: slots }, op, sheet, { kind: 'dv', type: getAttr(/^<[^>]*>/.exec(xml)?.[0] ?? '', 'type'), inheritFormat: g.leak });
  res = extGrow(res, op, g).filter((r) => r.sqref.length);
  if (!res.length) return { xml: '', copies: 0, changed: true };
  if (res.length === 1 && sameAreas(res[0].sqref, areas) && res[0].formulas.every((f, i) => f === slots[i])) return { xml, copies: 1, changed: false };
  const parts = res.map((r, i) => {
    let x = setSqref(xml, kind, renderSqref(r.sqref));
    x = setSlots(x, kind, r.formulas);
    if (i > 0) x = x.replace(/\sxr:uid="[^"]*"/g, '');
    return x;
  });
  return { xml: parts.join(''), copies: res.length, changed: true };
}

/** Apply one op to one unit. */
export function adjustUnit(xml: string, kind: UnitKind, op: StructOp, sheet: string, g: GrowRule): UnitResult {
  if (kind === 'cf' || kind === 'x14cf') return adjustCfBlock(xml, kind === 'x14cf', op, sheet, g);
  return adjustDv(xml, kind === 'x14dv', op, sheet, g);
}

/**
 * Apply an op to every CF / DV unit in a sheet tail. Returns the new tail and whether CF rules were duplicated
 * (then priorities must be renumbered).
 */
export function adjustRulesInTail(tail: string, op: StructOp, sheet: string, g: GrowRule): { tail: string; split: boolean } {
  let split = false;
  tail = replaceAll(tail, 'conditionalFormatting', (text) => {
    const r = adjustUnit(text, 'cf', op, sheet, g);
    if (r.copies > 1) split = true;
    return r.xml;
  });
  tail = replaceInContainer(tail, 'dataValidations', 'dataValidation', (text) => adjustUnit(text, 'dv', op, sheet, g).xml);
  tail = replaceAll(tail, 'x14:conditionalFormatting', (text) => {
    const r = adjustUnit(text, 'x14cf', op, sheet, g);
    if (r.copies > 1) split = true;
    return r.xml;
  });
  tail = replaceInContainer(tail, 'x14:dataValidations', 'x14:dataValidation', (text) => adjustUnit(text, 'x14dv', op, sheet, g).xml);
  tail = tail.replace(/<x14:conditionalFormattings>\s*<\/x14:conditionalFormattings>/g, '');
  tail = removeEmptyExt(tail);
  return { tail, split };
}

function replaceAll(xml: string, tag: string, fn: (text: string) => string): string {
  const list = findAll(xml, tag);
  if (!list.length) return xml;
  let out = '', pos = 0;
  for (const m of list) { out += xml.slice(pos, m.start) + fn(m.text); pos = m.end; }
  return out + xml.slice(pos);
}

/** Rewrite children of a counted container (<dataValidations count>), fixing the count / removing it when empty. */
function replaceInContainer(xml: string, container: string, child: string, fn: (text: string) => string): string {
  return replaceAll(xml, container, (ctext) => {
    const open = /^<[^>]*>/.exec(ctext)![0];
    if (open.endsWith('/>')) return ctext;
    const inner = ctext.slice(open.length, ctext.length - (`</${container}>`).length);
    const newInner = replaceAll(inner, child, fn);
    if (newInner === inner) return ctext;
    const n = findAll(newInner, child).length;
    if (!n) return '';
    return setAttr(open, 'count', String(n)) + newInner + `</${container}>`;
  });
}

/** Drop <ext> elements (and an empty <extLst>) that became empty after rule removal. */
export function removeEmptyExt(xml: string): string {
  let x = xml.replace(/<ext\b[^>]*>\s*<\/ext>/g, '');
  x = x.replace(/<extLst>\s*<\/extLst>/g, '');
  return x;
}

/** Renumber cfRule priorities 1..n keeping the relative order (classic + x14). */
export function renumberPriorities(tail: string): string {
  const occ: Array<{ p: number; i: number }> = [];
  let i = 0;
  const re = /<(?:x14:)?cfRule\b[^>]*>/g;
  for (const m of tail.matchAll(re)) {
    const p = getAttr(m[0], 'priority');
    if (p !== undefined) occ.push({ p: Number(p), i });
    i++;
  }
  const order = occ.slice().sort((a, b) => a.p - b.p || a.i - b.i);
  const newP = new Map<number, number>();
  order.forEach((o, k) => newP.set(o.i, k + 1));
  let j = 0;
  return tail.replace(re, (m0) => {
    const idx = j++;
    const np = newP.get(idx);
    return np !== undefined ? setAttr(m0, 'priority', String(np)) : m0;
  });
}

/** Adjust only the formulas of every unit (rules hosted on a sheet whose cells do not move). */
export function mapRuleFormulas(tail: string, fn: (f: string) => string): string {
  const map = (kind: UnitKind) => (t: string) => {
    const slots = getSlots(t, kind);
    const next = slots.map(fn);
    return next.every((f, i) => f === slots[i]) ? t : setSlots(t, kind, next);
  };
  let x = replaceAll(tail, 'conditionalFormatting', map('cf'));
  x = replaceAll(x, 'dataValidation', map('dv'));
  x = replaceAll(x, 'x14:conditionalFormatting', map('x14cf'));
  x = replaceAll(x, 'x14:dataValidation', map('x14dv'));
  return x;
}
