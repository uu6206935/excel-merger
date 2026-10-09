// Pure helpers over inspect.ps1 dumps (owner: T): reference extraction, header lookup, summaries and a generic
// feature-by-feature diff of two dumps. No I/O — unit-tested in tests/qa/compare-core.test.ts.

export const MAX_ROW = 1048576;
export const MAX_COL = 16384;

export function colToNum(letters) {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

export function numToCol(n) {
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/** 'B12' -> {row: 12, col: 2} (no sheet, no $) */
export function parseCell(addr) {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(addr);
  if (!m) return null;
  return { row: Number(m[2]), col: colToNum(m[1]) };
}

export function cellAddr(row, col) {
  return numToCol(col) + row;
}

/** 'A1:C5' / 'C:C' / '3:3' / 'B2' -> box; null when not an A1 range */
export function parseRange(text) {
  const t = text.replace(/\$/g, '');
  let m = /^([A-Za-z]{1,3})(\d+)(?::([A-Za-z]{1,3})(\d+))?$/.exec(t);
  if (m) {
    const r1 = Number(m[2]);
    const c1 = colToNum(m[1]);
    const r2 = m[4] ? Number(m[4]) : r1;
    const c2 = m[3] ? colToNum(m[3]) : c1;
    return { r1: Math.min(r1, r2), c1: Math.min(c1, c2), r2: Math.max(r1, r2), c2: Math.max(c1, c2) };
  }
  m = /^([A-Za-z]{1,3}):([A-Za-z]{1,3})$/.exec(t);
  if (m) {
    const c1 = colToNum(m[1]);
    const c2 = colToNum(m[2]);
    return { r1: 1, c1: Math.min(c1, c2), r2: MAX_ROW, c2: Math.max(c1, c2) };
  }
  m = /^(\d+):(\d+)$/.exec(t);
  if (m) {
    const r1 = Number(m[1]);
    const r2 = Number(m[2]);
    return { r1: Math.min(r1, r2), c1: 1, r2: Math.max(r1, r2), c2: MAX_COL };
  }
  return null;
}

/** 'A1:B2 D4' (space separated, as Excel's Address for multi-area ranges uses ',') -> boxes */
export function parseRanges(text) {
  if (!text) return [];
  return text
    .split(/[ ,]+/)
    .map((p) => p.trim())
    .filter(Boolean)
    .map((p) => parseRange(p.includes('!') ? p.slice(p.lastIndexOf('!') + 1) : p))
    .filter(Boolean);
}

const IDENT_CH = /[\p{L}\p{N}_.$\\?]/u;
// sheet prefix (quoted or unquoted) + A1 reference (cell, range, whole columns, whole rows)
const REF_RE =
  /(?:('(?:[^']|'')+'|[\p{L}_\\][\p{L}\p{N}_.]*)!)?(\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?\d+:\$?\d+)(?![\p{L}\p{N}_(!.])/uy;

function unquoteSheet(s) {
  if (s.startsWith("'") && s.endsWith("'")) return s.slice(1, -1).replace(/''/g, "'");
  return s;
}

/**
 * Walks the A1 references of a formula outside string literals / structured references. `visit(m)` gets the
 * REF_RE match (m[1] = sheet prefix without '!', m[2] = the A1 part) and may return replacement text for it.
 * Returns the formula with the replacements applied.
 */
function scanRefs(f, visit) {
  let out = '';
  let last = 0;
  let i = 0;
  let bracket = 0;
  while (i < f.length) {
    const ch = f[i];
    if (ch === '"') {
      // string literal
      i++;
      while (i < f.length) {
        if (f[i] === '"') {
          if (f[i + 1] === '"') i += 2;
          else break;
        } else i++;
      }
      i++;
      continue;
    }
    if (ch === '[') {
      bracket++;
      i++;
      continue;
    }
    if (ch === ']') {
      bracket = Math.max(0, bracket - 1);
      i++;
      continue;
    }
    if (bracket > 0) {
      i++;
      continue;
    }
    const prev = i > 0 ? f[i - 1] : '';
    // inside an identifier, right after a sheet prefix we could not parse (3-D / external refs), or the second
    // sheet of a 3-D reference (Sheet1:Sheet3!A1)
    if (prev && (IDENT_CH.test(prev) || prev === "'" || prev === ']' || prev === '!' || (prev === ':' && /[\p{L}\p{N}$]/u.test(f[i - 2] ?? '')))) {
      i++;
      continue;
    }
    if (ch === "'" || /[\p{L}_$\\0-9]/u.test(ch)) {
      REF_RE.lastIndex = i;
      const m = REF_RE.exec(f);
      if (m) {
        const end = REF_RE.lastIndex;
        const rep = visit(m);
        if (rep != null && rep !== m[0]) {
          out += f.slice(last, i) + rep;
          last = end;
        }
        i = end;
        continue;
      }
      // skip the rest of this identifier / quoted sheet name
      if (ch === "'") {
        i++;
        while (i < f.length && !(f[i] === "'" && f[i + 1] !== "'")) i += f[i] === "'" ? 2 : 1;
        i++;
      } else {
        i++;
        while (i < f.length && IDENT_CH.test(f[i])) i++;
      }
      continue;
    }
    i++;
  }
  return out + f.slice(last);
}

/**
 * A1 references in a formula (UI or file form, with or without '='), outside string literals.
 * Returns [{sheet, box: {r1,c1,r2,c2}, text, qualified}] where sheet = explicit sheet or `currentSheet`.
 * Structured references, external books and 3-D references (Sheet1:Sheet3!A1) are ignored.
 */
export function extractRefs(formula, currentSheet = null) {
  const out = [];
  if (typeof formula !== 'string') return out;
  scanRefs(formula, (m) => {
    const box = parseRange(m[2]);
    // a 3-D reference 'Sheet1:Sheet3!A1' shows up as a ref 'Sheet1' + ':'... -> skip anything followed by '!'
    if (box && box.c2 <= MAX_COL && box.r2 <= MAX_ROW && box.r1 >= 1 && box.c1 >= 1) {
      const qualified = m[1] != null;
      out.push({ sheet: qualified ? unquoteSheet(m[1]) : currentSheet, box, text: m[0], qualified });
    }
    return null;
  });
  return out;
}

function shiftPart(abs, n, delta, max) {
  if (abs || !delta) return { ok: true, n };
  const v = n + delta;
  return { ok: v >= 1 && v <= max, n: v };
}

/**
 * Shifts the RELATIVE parts of every A1 reference by (dr, dc) — how Excel derives the formula of a shared-formula
 * child cell from its master (<f t="shared" si="0"/>). Out-of-sheet results become #REF!.
 */
export function shiftFormula(formula, dr, dc) {
  if (typeof formula !== 'string' || (!dr && !dc)) return formula;
  return scanRefs(formula, (m) => {
    const prefix = m[1] != null ? m[0].slice(0, m[0].length - m[2].length) : '';
    const t = m[2];
    let bad = false;
    const cell = (s) =>
      s.replace(/^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$/, (_all, ca, cl, ra, rw) => {
        const c = shiftPart(!!ca, colToNum(cl), dc, MAX_COL);
        const r = shiftPart(!!ra, Number(rw), dr, MAX_ROW);
        if (!c.ok || !r.ok) bad = true;
        return `${ca}${numToCol(c.n)}${ra}${r.n}`;
      });
    let res;
    let mm;
    if ((mm = /^(\$?[A-Za-z]{1,3}\$?\d+)(?::(\$?[A-Za-z]{1,3}\$?\d+))?$/.exec(t))) {
      res = cell(mm[1]) + (mm[2] ? ':' + cell(mm[2]) : '');
    } else if ((mm = /^(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})$/.exec(t))) {
      const a = shiftPart(!!mm[1], colToNum(mm[2]), dc, MAX_COL);
      const b = shiftPart(!!mm[3], colToNum(mm[4]), dc, MAX_COL);
      if (!a.ok || !b.ok) bad = true;
      res = `${mm[1]}${numToCol(a.n)}:${mm[3]}${numToCol(b.n)}`;
    } else if ((mm = /^(\$?)(\d+):(\$?)(\d+)$/.exec(t))) {
      const a = shiftPart(!!mm[1], Number(mm[2]), dr, MAX_ROW);
      const b = shiftPart(!!mm[3], Number(mm[4]), dr, MAX_ROW);
      if (!a.ok || !b.ok) bad = true;
      res = `${mm[1]}${a.n}:${mm[3]}${b.n}`;
    } else return null;
    return prefix + (bad ? '#REF!' : res);
  });
}

/** functions whose result differs between two recalculations (time, random, environment) */
export const VOLATILE_RE = /(^|[^\p{L}\p{N}_.])(?:_xlfn\.)?(NOW|TODAY|RAND|RANDBETWEEN|RANDARRAY|CELL|INFO)\s*\(/iu;

export function sheetByName(dump, name) {
  return (dump?.sheets ?? []).find((s) => s.name === name) ?? null;
}

/** header row guess of a sheet: the row (among the first 40) with the most text cells; ties -> first */
export function headerRowGuess(sheet) {
  const counts = sheet?.rowTextCounts ?? [];
  let best = -1;
  let bestN = 0;
  counts.forEach((n, i) => {
    if (n > bestN) {
      bestN = n;
      best = i;
    }
  });
  if (best < 0) return null;
  return (sheet.firstRow ?? 1) + best;
}

export function cellValue(sheet, addr) {
  if (!sheet) return undefined;
  if (sheet.errors && addr in sheet.errors) return sheet.errors[addr];
  return sheet.values?.[addr];
}

export function headerText(sheet, headerRow, col) {
  if (!sheet || !headerRow) return null;
  const v = cellValue(sheet, cellAddr(headerRow, col));
  return v == null ? null : String(v).trim();
}

/** map of header text -> [columns] on `headerRow` */
export function headerIndex(sheet, headerRow) {
  const map = new Map();
  if (!sheet || !headerRow) return map;
  const maxCol = sheet.maxCol ?? 0;
  for (let c = 1; c <= maxCol; c++) {
    const h = headerText(sheet, headerRow, c);
    if (!h) continue;
    if (!map.has(h)) map.set(h, []);
    map.get(h).push(c);
  }
  return map;
}

export function arr(x) {
  return x == null ? [] : Array.isArray(x) ? x : [x];
}

export function totalErrors(dump) {
  const t = {};
  for (const s of arr(dump?.sheets)) for (const [k, n] of Object.entries(s.errorCounts ?? {})) t[k] = (t[k] ?? 0) + n;
  return t;
}

/** one-line-per-sheet text summary of a dump */
export function summarize(dump) {
  const lines = [];
  lines.push(`${dump.file}: open=${dump.openMode}${dump.openError && dump.openMode !== 'normal' ? ` (${dump.openError})` : ''}${dump.error ? ` ERROR ${dump.error}` : ''}`);
  for (const r of arr(dump.repairLog)) lines.push(`  repair: ${r}`);
  if (dump.names) lines.push(`  names: ${arr(dump.names).map((n) => `${n.scope ? n.scope + '!' : ''}${n.name}${n.refersTo}`).join('  ')}`);
  for (const s of arr(dump.sheets)) {
    const errs = Object.entries(s.errorCounts ?? {}).map(([k, n]) => `${k}x${n}`).join(' ');
    const feats = [
      s.tables?.length ? `tables=${s.tables.map((t) => `${t.name}(${t.range})`).join(',')}` : '',
      s.charts?.length ? `charts=${s.charts.length}` : '',
      s.pivots?.length ? `pivots=${s.pivots.map((p) => `${p.name}<-${p.source}`).join(',')}` : '',
      s.cf?.length ? `cf=${s.cf.length}` : '',
      s.dv?.length ? `dv=${s.dv.length}` : '',
      s.merges?.length ? `merges=${s.merges.join(',')}` : '',
      s.freeze?.frozen ? `freeze=${s.freeze.splitRow}r/${s.freeze.splitCol}c` : '',
      s.autoFilter ? `filter=${s.autoFilter}` : '',
      s.hiddenColumns?.length ? `hidden=${s.hiddenColumns.join(',')}` : '',
    ]
      .filter(Boolean)
      .join(' ');
    lines.push(`  [${s.name}] ${s.usedRange} header~${headerRowGuess(s) ?? '-'} formulas=${s.formulaCount}${s.refTextCount ? ` #REF!-text=${s.refTextCount}` : ''}${errs ? ` errors: ${errs}` : ''} ${feats}`);
  }
  return lines.join('\n');
}

// ---------------------------------------------------------------------------------------------------------------
// Generic diff of two dumps (e.g. JS-engine output vs Excel-engine output of the same op log)
// ---------------------------------------------------------------------------------------------------------------

const eq = (a, b) => JSON.stringify(a ?? null) === JSON.stringify(b ?? null);

/** skip(k) -> true: a known-benign difference that is counted in `filtered` instead of `bad` */
function diffMap(a = {}, b = {}, limit = 8, skip = null) {
  const keys = new Set([...Object.keys(a ?? {}), ...Object.keys(b ?? {})]);
  let same = 0;
  const bad = [];
  const filtered = [];
  for (const k of keys) {
    const x = a?.[k];
    const y = b?.[k];
    const same1 = typeof x === 'number' && typeof y === 'number' ? Math.abs(x - y) <= 1e-9 * Math.max(1, Math.abs(x)) : x === y;
    if (same1) same++;
    else if (skip?.(k)) filtered.push(`${k}: ${JSON.stringify(x)} -> ${JSON.stringify(y)}`);
    else bad.push(`${k}: ${JSON.stringify(x)} -> ${JSON.stringify(y)}`);
  }
  return { same, total: keys.size, bad: bad.slice(0, limit), badCount: bad.length, filtered };
}

/** cells whose value is not reproducible: volatile formulas and the cells they spill into */
function volatileCells(s) {
  const out = new Set();
  for (const [addr, f] of Object.entries(s?.formulas ?? {})) if (VOLATILE_RE.test(String(f))) out.add(addr);
  for (const sp of arr(s?.spills)) {
    const [master, range] = String(sp).split('=>');
    if (!out.has(master)) continue;
    const box = parseRange(range ?? '');
    if (!box || (box.r2 - box.r1 + 1) * (box.c2 - box.c1 + 1) > 100000) continue;
    for (let r = box.r1; r <= box.r2; r++) for (let c = box.c1; c <= box.c2; c++) out.add(cellAddr(r, c));
  }
  return out;
}

/**
 * Feature-by-feature comparison of two inspect dumps (what Excel shows). Returns rows {group, sheet, ok, detail};
 * `detail` lists the differences. Rows with `filtered: true` are known-benign differences that were NOT counted
 * (values of volatile functions NOW/TODAY/RAND/RANDBETWEEN/RANDARRAY/CELL/INFO, pivot-sheet layout).
 * Options: {ignoreValues?: boolean, pivotLayout?: boolean}
 */
export function compareInspections(a, b, opts = {}) {
  const rows = [];
  const push = (group, sheet, ok, detail = '', filtered = false) => rows.push(filtered ? { group, sheet, ok, detail, filtered } : { group, sheet, ok, detail });
  push('open', '', a.openMode === b.openMode, `${a.openMode} -> ${b.openMode}`);
  if (!a.sheets || !b.sheets) return rows;
  const an = arr(a.sheets).map((s) => `${s.name}:${s.visible}`);
  const bn = arr(b.sheets).map((s) => `${s.name}:${s.visible}`);
  push('sheets', '', eq(an, bn), eq(an, bn) ? '' : `${an.join(',')} -> ${bn.join(',')}`);
  const names = (d) => arr(d.names).filter((n) => !/^_xl/i.test(n.name)).map((n) => `${n.scope ?? ''}!${n.name}=${n.refersTo}`).sort();
  const na = names(a);
  const nb = names(b);
  push('names', '', eq(na, nb), eq(na, nb) ? '' : `missing: ${na.filter((x) => !nb.includes(x)).join(' | ')}  extra: ${nb.filter((x) => !na.includes(x)).join(' | ')}`);
  for (const sa of arr(a.sheets)) {
    const sb = sheetByName(b, sa.name);
    if (!sb) {
      push('sheet', sa.name, false, 'missing');
      continue;
    }
    const fd = diffMap(sa.formulas, sb.formulas);
    push('formulas', sa.name, fd.badCount === 0, `${fd.same}/${fd.total}${fd.badCount ? ' ; ' + fd.bad.join(' ; ') : ''}`);
    if (!opts.ignoreValues) {
      const vol = new Set([...volatileCells(sa), ...volatileCells(sb)]);
      const vd = diffMap({ ...sa.values, ...sa.errors }, { ...sb.values, ...sb.errors }, 8, (k) => vol.has(k));
      push('values', sa.name, vd.badCount === 0, `${vd.same}/${vd.total}${vd.badCount ? ' ; ' + vd.bad.join(' ; ') : ''}`);
      if (vd.filtered.length)
        push('values', sa.name, true, `揮発性関数 (NOW/TODAY/RAND…) の値 ${vd.filtered.length} セル: ${vd.filtered.slice(0, 4).join(' ; ')}`, true);
    }
    const simple = {
      usedRange: (s) => s.usedRange,
      arrays: (s) => s.arrays,
      spills: (s) => s.spills,
      merges: (s) => s.merges,
      cf: (s) => arr(s.cf).map((c) => `${c.type}|${c.appliesTo}|${c.operator ?? ''}|${c.formula1 ?? ''}|${c.formula2 ?? ''}`),
      dv: (s) => arr(s.dv).map((d) => `${d.range}|${d.type}|${d.operator ?? ''}|${d.formula1 ?? ''}|${d.formula2 ?? ''}`),
      tables: (s) => arr(s.tables).map((t) => `${t.name}|${t.range}|${t.showTotals}|${arr(t.columns).join(',')}`),
      charts: (s) => arr(s.charts).map((c) => `${c.name}|${c.type}|${c.title ?? ''}|${arr(c.series).map((x) => x.formula).join(';')}`),
      pivots: (s) => arr(s.pivots).map((p) => `${p.name}|${p.range}|${p.source}|${arr(p.rowFields).join(',')}|${arr(p.columnFields).join(',')}|${arr(p.dataFields).join(',')}`),
      cols: (s) => arr(s.cols).map((c) => `${c.letter}=${c.width}${c.hidden ? 'h' : ''}`),
      styles: (s) => s.probes ?? null,
      freeze: (s) => s.freeze,
      autoFilter: (s) => s.autoFilter,
      objects: (s) => ({ shapes: s.shapes, comments: s.comments, threaded: s.threadedComments, links: s.hyperlinks }),
    };
    // Excel re-lays-out a pivot table when it refreshes it (autofit widths, used range, formats); a package patcher
    // cannot reproduce that, so the layout of pivot sheets is only compared with {pivotLayout: true}
    const pivotSheet = arr(sa.pivots).length > 0 || arr(sb.pivots).length > 0;
    const layoutKeys = new Set(['cols', 'styles', 'usedRange']);
    for (const [k, get] of Object.entries(simple)) {
      const x = get(sa);
      const y = get(sb);
      if (pivotSheet && !opts.pivotLayout && layoutKeys.has(k) && !eq(x, y)) {
        push(k, sa.name, true, `ピボットの再レイアウト (pivot layout): ${JSON.stringify(x)} -> ${JSON.stringify(y)}`.slice(0, 400), true);
        continue;
      }
      push(k, sa.name, eq(x, y), eq(x, y) ? '' : `${JSON.stringify(x)} -> ${JSON.stringify(y)}`);
    }
  }
  for (const sb of arr(b.sheets)) if (!sheetByName(a, sb.name)) push('sheet', sb.name, false, 'extra sheet');
  return rows;
}

/**
 * Markdown table of the rows (only the differences with onlyDiffs), followed by the list of filtered
 * (known-benign) differences unless showFiltered is false.
 */
export function formatComparison(rows, { onlyDiffs = false, showFiltered = true } = {}) {
  const shown = rows.filter((r) => !r.filtered && !(onlyDiffs && r.ok));
  // no table at all when there is nothing to show (a report with only differences stays short)
  const lines = shown.length ? ['| group | sheet | result | detail |', '|---|---|---|---|'] : [];
  for (const r of shown) {
    lines.push(`| ${r.group} | ${r.sheet} | ${r.ok ? 'OK' : 'DIFF'} | ${String(r.detail).replace(/\|/g, '\\|').replace(/\r?\n/g, ' ').slice(0, 600)} |`);
  }
  const filtered = rows.filter((r) => r.filtered);
  if (showFiltered && filtered.length) {
    lines.push('', 'Filtered (known-benign, not counted):');
    for (const r of filtered) lines.push(`- [${r.group}${r.sheet ? ' / ' + r.sheet : ''}] ${String(r.detail).replace(/\r?\n/g, ' ').slice(0, 400)}`);
  }
  return lines.join('\n').replace(/^\n+/, '');
}

/** 'formulas 2, styles 1' — the groups with real differences and how many rows differ in each */
export function diffSummary(rows) {
  const counts = new Map();
  for (const r of rows) if (!r.ok && !r.filtered) counts.set(r.group, (counts.get(r.group) ?? 0) + 1);
  return [...counts].map(([g, n]) => `${g} ${n}`).join(', ');
}
