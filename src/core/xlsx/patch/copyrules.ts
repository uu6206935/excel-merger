/**
 * Conditional formats, data validations and sparklines of a sheet that is copied into another workbook (addsheets.ts): classic elements
 * and their x14 extension twins. What changes: dxf ids (the dxf records are imported into the target styles.xml), formulas (sheet
 * references follow the renamed copies, defined names are collected), colours of the source theme (converted when the themes differ),
 * the GUIDs that link a classic rule to its x14 extension (fresh ones), `xr:uid` attributes (dropped). A rule whose formula cannot be
 * copied (external workbook, 3D reference ...) is dropped and counted.
 */
import { getSlots, setSlots } from './rules';
import type { Styles } from './styles';
import { convertColorTags } from './styles';
import { esc, findAll, getAttr, replaceElems, setAttr, unesc } from './xml';

export interface RuleCounts {
  cfDropped: number;
  dvDropped: number;
  sparklinesDropped: number;
  dvDateShifted: number;
}

export interface RuleEnv {
  /** the formula of a slot of a rule, re-pointed; null = it cannot be copied (the rule is dropped) */
  formula(f: string): string | null;
  /** the target styles.xml / the source styles.xml and the colour conversion between them */
  base: Styles;
  src: Styles;
  themeRgb: string[] | null;
  indexedRgb: string[] | null;
  /** days to add to a date constant (1904 <-> 1900) */
  dateShift: number;
  /** dxf id of the source -> dxf id in the target (null: the dxf does not exist) */
  dxf(id: number): number | null;
  /** a fresh GUID for a source GUID (stable for the same input) */
  guid(old: string): string;
  counts: RuleCounts;
}

const colors = (x: string, env: RuleEnv) => convertColorTags(x, env.themeRgb, env.indexedRgb);

/** rewrite every formula slot of a rule / validation; null when one of them cannot be copied */
function mapSlots(xml: string, kind: Parameters<typeof getSlots>[1], env: RuleEnv): string | null {
  const slots = getSlots(xml, kind);
  if (!slots.length) return xml;
  const next: string[] = [];
  let changed = false;
  for (const s of slots) {
    const r = env.formula(s);
    if (r === null) return null;
    if (r !== s) changed = true;
    next.push(r);
  }
  return changed ? setSlots(xml, kind, next) : xml;
}

const X14_ID = /<x14:id>([^<]*)<\/x14:id>/;

/**
 * One classic <conditionalFormatting> block. Returns the new block ('' = nothing left). `dropped`: the x14 GUIDs of the rules that were
 * dropped (their x14 twins must go too), `ids`: old GUID -> new GUID of the kept rules.
 */
export function copyCfBlock(block: string, env: RuleEnv, dropped: Set<string>): string {
  const open = /^<conditionalFormatting\b[^>]*>/.exec(block)?.[0];
  if (!open || open.endsWith('/>')) return '';
  const inner = block.slice(open.length, block.length - '</conditionalFormatting>'.length);
  const rules = findAll(inner, 'cfRule');
  let kept = 0;
  const body = replaceElems(inner, rules, (r) => {
    let x = mapSlots(r.text, 'cf', env);
    if (x === null) {
      env.counts.cfDropped++;
      const id = X14_ID.exec(r.text)?.[1];
      if (id) dropped.add(id);
      return '';
    }
    x = x.replace(/^<cfRule\b[^>]*>/, (o) => {
      const d = getAttr(o, 'dxfId');
      if (d === undefined) return o;
      const nd = env.dxf(Number(d));
      return setAttr(o, 'dxfId', nd === null ? null : String(nd));
    });
    x = colors(x, env);
    x = x.replace(X14_ID, (_m, id: string) => `<x14:id>${env.guid(id)}</x14:id>`);
    kept++;
    return x;
  });
  if (!kept) return '';
  // a pivot-table rule has no pivot table to belong to in the copy
  return setAttr(open, 'pivot', null) + body + '</conditionalFormatting>';
}

/** one <x14:conditionalFormatting> block of the extension list */
export function copyX14CfBlock(block: string, env: RuleEnv, dropped: Set<string>): string {
  const open = /^<x14:conditionalFormatting\b[^>]*>/.exec(block)?.[0];
  if (!open || open.endsWith('/>')) return '';
  const inner = block.slice(open.length, block.length - '</x14:conditionalFormatting>'.length);
  const rules = findAll(inner, 'x14:cfRule');
  let kept = 0;
  const body = replaceElems(inner, rules, (r) => {
    const id = getAttr(r.open, 'id');
    if (id && dropped.has(id)) return '';
    let x = mapSlots(r.text, 'x14cf', env);
    if (x === null) { env.counts.cfDropped++; return ''; }
    x = x.replace(/<x14:dxf\b[\s\S]*?<\/x14:dxf>/g, (d) => env.base.convertDxfXml(d, env.themeRgb, env.indexedRgb));
    x = colors(x, env);
    if (id) x = x.replace(/^<x14:cfRule\b[^>]*>/, (o) => setAttr(o, 'id', env.guid(id)));
    kept++;
    return x;
  });
  return kept ? open + body + '</x14:conditionalFormatting>' : '';
}

const DATE_LITERAL = /^\s*\d+(?:\.\d+)?\s*$/;

function validation(x: string, tag: 'dataValidation' | 'x14:dataValidation', env: RuleEnv): string {
  const x14 = tag === 'x14:dataValidation';
  const mapped = mapSlots(x, x14 ? 'x14dv' : 'dv', env);
  if (mapped === null) { env.counts.dvDropped++; return ''; }
  let y = mapped.replace(/\sxr:uid="[^"]*"/g, '');
  // date validations with constant dates ("2020/1/1" = 43831): the serial depends on the date system of the workbook
  if (env.dateShift !== 0 && /^<[^>]*\stype="date"/.test(y)) {
    const slots = getSlots(y, x14 ? 'x14dv' : 'dv');
    if (slots.length && slots.every((s) => DATE_LITERAL.test(s))) {
      y = setSlots(y, x14 ? 'x14dv' : 'dv', slots.map((s) => String(Number(s) + env.dateShift)));
      env.counts.dvDateShifted++;
    }
  }
  return y;
}

/** the whole <dataValidations> container (classic) or <x14:dataValidations> (extension); '' = nothing left */
export function copyDvContainer(container: string, env: RuleEnv, x14: boolean): string {
  const ctag = x14 ? 'x14:dataValidations' : 'dataValidations';
  const vtag = x14 ? 'x14:dataValidation' : 'dataValidation';
  const open = new RegExp(`^<${ctag}\\b[^>]*>`).exec(container)?.[0];
  if (!open || open.endsWith('/>')) return '';
  const inner = container.slice(open.length, container.length - `</${ctag}>`.length);
  const items = findAll(inner, vtag);
  let kept = 0;
  const body = replaceElems(inner, items, (it) => {
    const y = validation(it.text, vtag, env);
    if (y) kept++;
    return y;
  });
  return kept ? setAttr(open, 'count', String(kept)) + body + `</${ctag}>` : '';
}

/** <x14:sparklineGroups>: references re-pointed, colours converted; a sparkline that cannot be copied is dropped with its group when it was the last one */
export function copySparklineGroups(container: string, env: RuleEnv): string {
  const open = /^<x14:sparklineGroups\b[^>]*>/.exec(container)?.[0];
  if (!open || open.endsWith('/>')) return '';
  const inner = container.slice(open.length, container.length - '</x14:sparklineGroups>'.length);
  let kept = 0;
  const body = replaceElems(inner, findAll(inner, 'x14:sparklineGroup'), (g) => {
    let n = 0;
    let x = g.text.replace(/<x14:sparkline>([\s\S]*?)<\/x14:sparkline>/g, (m0, sp: string) => {
      const f = /<xm:f>([\s\S]*?)<\/xm:f>/.exec(sp);
      if (!f) { env.counts.sparklinesDropped++; return ''; }
      const r = env.formula(unesc(f[1]));
      if (r === null) { env.counts.sparklinesDropped++; return ''; }
      n++;
      return r === unesc(f[1]) ? m0 : m0.replace(f[0], `<xm:f>${esc(r)}</xm:f>`);
    });
    if (!n) return '';
    x = x.replace(/\sxr2?:uid="[^"]*"/g, '');
    kept++;
    return colors(x, env);
  });
  return kept ? open + body + '</x14:sparklineGroups>' : '';
}
