/**
 * Writes representative outputs of the robustness sweep into a folder, so that real Excel can be asked whether they open
 * without repair (tests/golden/com/run-ops.ps1 -OpenDir <folder>, or com/opentest.ps1 -Dir <folder>):
 *   npx tsx tests/golden/tools/make-fuzz-outputs.ts <folder>
 *  - rs-*   : golden fixtures (Excel-authored) spelled differently (prefix / single quotes + line breaks / UTF-16 / BOM), then
 *             replayed with their golden op logs;
 *  - pm-*   : buildPremerged on the small hand-written scenario (tests/golden/mini-scenario.ts) in unusual but valid forms;
 *  - h-*    : replayLog on hostile-but-valid hand-written books (8,000-character formula, a full row of 16,384 styled cells,
 *             awkward sheet names, no styles.xml / sharedStrings.xml, inline strings, rich strings, spans ...);
 *  - fz-*   : accepted mutants of the seeded fuzz that only change the spelling (BOM, line breaks, quotes).
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import type { LoggedOp, StructOp } from '../../../src/core/types';
import { buildPremerged } from '../../../src/core/xlsx/premerge';
import { replayLog } from '../../../src/core/xlsx/replay';
import { buildXlsx } from '../build';
import { CASES } from '../cases';
import { mutate, randomLog, replayTargets, sheetNames } from '../fuzz';
import { rng } from '../package-tools';
import { SH1, SH2, SRC_SST, BASE_SST, inlineStrings, mini, miniInput, noStyles } from '../mini-scenario';
import { edit, prefixMain, text, utf16le, type Parts } from '../package-tools';

const here = path.dirname(fileURLToPath(import.meta.url));
const FIX = path.resolve(here, '../../fixtures/golden');
const out = path.resolve(process.argv[2] ?? 'fuzz-out');
fs.mkdirSync(out, { recursive: true });

const save = (name: string, bytes: Uint8Array) => { fs.writeFileSync(path.join(out, name + '.xlsx'), bytes); console.log('wrote', name); };
const fixture = (f: string) => new Uint8Array(fs.readFileSync(path.join(FIX, f)));
const xmlNames = (p: Parts) => [...p.keys()].filter((n) => /\.(xml|rels)$/i.test(n) && typeof p.get(n) === 'string');
const mainParts = (p: Parts) => xmlNames(p).filter((n) => /^xl\/(worksheets\/[^/]+|workbook|styles|sharedStrings|tables\/[^/]+)\.xml$/.test(n));

const struct = (sheet: string, op: StructOp): LoggedOp => ({ kind: 'struct', sheet, op });
const insC = (sheet: string, at: number, count = 1) => struct(sheet, { type: 'insert', axis: 'col', at, count });
const delC = (sheet: string, at: number, count = 1) => struct(sheet, { type: 'delete', axis: 'col', at, count });
const movC = (sheet: string, from: number, count: number, to: number) => struct(sheet, { type: 'move', axis: 'col', from, count, to });
const cells = (sheet: string, writes: Array<{ row: number; col: number; cell: { t: 's' | 'n' | 'd' | 'z'; v: string | number | null; f?: string; numFmt?: string } | null }>): LoggedOp => ({ kind: 'cells', sheet, writes, reason: 'edit' });

// ---- rs-*: respelled Excel fixtures -----------------------------------------------------------------------------
const SPELLINGS: Array<[string, (p: Parts) => void]> = [
  ['prefix', (p) => { for (const n of xmlNames(p)) p.set(n, prefixMain(p.get(n) as string, 'x')); }],
  ['quotes', (p) => { for (const n of mainParts(p)) p.set(n, (p.get(n) as string).replace(/(\s[\w:.-]+)="([^"']*)"/g, "$1='$2'").replace(/></g, '>\r\n  <')); }],
  ['utf16', (p) => { for (const n of mainParts(p)) p.set(n, utf16le(p.get(n) as string)); }],
  ['bom', (p) => { for (const n of xmlNames(p)) p.set(n, '﻿' + (p.get(n) as string)); }],
];
for (const cname of ['t-seq', 'g-mov-D-B', 'h2-totals-ops', 'u-band-center', 'q-wrap', 'r-dyn']) {
  const c = CASES.find((x) => x.name === cname);
  if (!c) continue;
  for (const [label, respell] of SPELLINGS) {
    save(`rs-${cname}-${label}`, await replayLog(await edit(fixture(c.fixture), respell), c.ops));
  }
}

// ---- pm-*: buildPremerged -----------------------------------------------------------------------------------------
{
  const m = await mini();
  const prem = async (name: string, base = m.base, src = m.src) => save(`pm-${name}`, (await buildPremerged(miniInput(m, base, src))).bytes);
  await prem('plain');
  await prem('prefixed', await edit(m.base, (p) => { for (const k of xmlNames(p)) p.set(k, prefixMain(p.get(k) as string, 'ss')); }), await edit(m.src, (p) => { for (const k of xmlNames(p)) p.set(k, prefixMain(p.get(k) as string, 'x')); }));
  await prem('inline-both', await edit(m.base, (p) => inlineStrings(p, BASE_SST, [SH1])), await edit(m.src, (p) => inlineStrings(p, SRC_SST, [SH1])));
  await prem('base-no-sst', await edit(m.base, (p) => inlineStrings(p, BASE_SST, [SH1])));
  await prem('base-no-styles', await edit(m.base, (p) => noStyles(p, [SH1, SH2])));
  await prem('src-no-styles', m.base, await edit(m.src, (p) => noStyles(p, [SH1])));
  await prem('utf16', await edit(m.base, (p) => { for (const k of [SH1, SH2, 'xl/workbook.xml', 'xl/sharedStrings.xml', 'xl/styles.xml']) p.set(k, utf16le(p.get(k) as string)); }),
    await edit(m.src, (p) => { for (const k of [SH1, 'xl/workbook.xml', 'xl/sharedStrings.xml', 'xl/styles.xml']) p.set(k, utf16le(p.get(k) as string)); }));
  const q = await mini("It's 名簿 & 売上", "A&B 'src'");
  save('pm-quoted-names', (await buildPremerged(miniInput(q))).bytes);
  let extra = '';
  for (let c = 3; c <= 16384; c++) { let t = '', x = c; while (x > 0) { t = String.fromCharCode(65 + ((x - 1) % 26)) + t; x = Math.floor((x - 1) / 26); } extra += `<c r="${t}1" s="1"/>`; }
  await prem('wide-row', await edit(m.base, (p) => p.set(SH1, text(p, SH1).replace('</c></row><row r="2">', `</c>${extra}</row><row r="2">`))));
}

// ---- h-*: replayLog on hostile-but-valid books --------------------------------------------------------------------
const DATA = '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" t="s"><v>2</v></c></row>'
  + '<row r="2"><c r="A2"><v>1</v></c><c r="B2"><v>2</v></c><c r="C2"><f>A2+B2</f><v>3</v></c></row>'
  + '<row r="3"><c r="A3"><v>4</v></c><c r="B3"><v>5</v></c><c r="C3"><f>A3+B3</f><v>9</v></c></row>';
const base = () => buildXlsx([{ name: 'S', data: DATA }], { sst: ['a', 'b', 'c'] });
{
  // ~5,500 characters: 1,000 references plus 10 text constants of 250 characters. Excel's limits are 8,192 characters AND
  // 16,384 bytes of compiled formula (a reference costs about 8 bytes, a text constant 2 bytes per character), so a chain of
  // 2,600 references is not a formula Excel itself can hold ("repair"): the compiled size of this one is about 13,000 bytes.
  const f = Array.from({ length: 1000 }, (_x, i) => `B${(i % 3) + 1}`).join('+') + Array.from({ length: 10 }, (_x, i) => `&"${String.fromCharCode(97 + i).repeat(250)}"`).join('');
  save('h-long-formula', await replayLog(await buildXlsx([{ name: 'S', data: `<row r="1"><c r="A1"><f>${f}</f><v>0</v></c></row><row r="2"><c r="B2"><v>1</v></c></row>` }]), [insC('S', 2), struct('S', { type: 'insert', axis: 'row', at: 1, count: 1 })]));
  let row = '<row r="1">';
  for (let c = 1; c <= 16384; c++) { let t = '', x = c; while (x > 0) { t = String.fromCharCode(65 + ((x - 1) % 26)) + t; x = Math.floor((x - 1) / 26); } row += `<c r="${t}1" s="${c % 3 === 0 ? 1 : 0}"/>`; }
  save('h-wide-row', await replayLog(await buildXlsx([{ name: 'S', data: row + '</row><row r="2"><c r="A2"><v>7</v></c></row>', head: '<cols><col min="1" max="16384" width="9" style="2" customWidth="1"/></cols>' }], { sst: ['a'] }), [insC('S', 1, 2), delC('S', 5), movC('S', 10, 3, 1)]));
  const NAMES = ["It's 売上", 'A1', '2024年度', 'a b', 'AT&T 売上', 'R1C1'];
  const esc = (s: string) => s.replace(/&/g, '&amp;');
  const q = (s: string) => `'${s.replace(/'/g, "''")}'`;
  const sheets = NAMES.map((n, i) => ({ name: esc(n), data: `<row r="1"><c r="A1"><v>${i + 1}</v></c><c r="B1"><v>10</v></c></row><row r="2"><c r="A2"><f>A1*2</f><v>0</v></c></row>` }));
  sheets.push({ name: 'Main', data: NAMES.map((n, i) => `<row r="${i + 1}"><c r="A${i + 1}"><f>${esc(q(n))}!B1+SUM(${esc(q(n))}!A1:B2)</f><v>0</v></c></row>`).join('') });
  const names = NAMES.map((n, i) => `<definedName name="Nm${i}">${esc(q(n))}!$B$1</definedName>`).join('');
  save('h-awkward-names', await replayLog(await buildXlsx(sheets, { names }), NAMES.map((n) => insC(n, 1))));
  const nost = await edit(await buildXlsx([{ name: 'S', data: DATA.replace(/ s="\d+"/g, '') }], { sst: ['a', 'b', 'c'] }), (p) => noStyles(p, [SH1]));
  save('h-no-styles', await replayLog(nost, [cells('S', [{ row: 5, col: 1, cell: { t: 'd', v: 46295, numFmt: 'yyyy/m/d' } }, { row: 5, col: 2, cell: { t: 's', v: 'a\nb' } }]), movC('S', 3, 1, 1)]));
  const nosst = await edit(await buildXlsx([{ name: 'S', data: '<row r="1"><c r="A1"><v>1</v></c><c r="B1" t="inlineStr"><is><t>i</t></is></c></row>' }]), (p) => {
    p.delete('xl/sharedStrings.xml');
    p.set('xl/_rels/workbook.xml.rels', text(p, 'xl/_rels/workbook.xml.rels').replace(/<Relationship [^>]*\/sharedStrings"[^>]*\/>/, ''));
    p.set('[Content_Types].xml', text(p, '[Content_Types].xml').replace(/<Override PartName="\/xl\/sharedStrings.xml"[^>]*\/>/, ''));
  });
  save('h-no-sst', await replayLog(nosst, [cells('S', [{ row: 3, col: 1, cell: { t: 's', v: '文字' } }]), insC('S', 1)]));
  const inline = '<row r="1"><c r="A1" t="inlineStr"><is><t>ID</t></is></c><c r="B1" t="inlineStr"><is><r><rPr><b/></rPr><t>太</t></r><r><t xml:space="preserve"> 普通</t></r></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>a&amp;b</t></is></c><c r="B2"><v>2</v></c></row>';
  save('h-inline', await replayLog(await buildXlsx([{ name: 'S', data: inline }]), [movC('S', 2, 1, 1), cells('S', [{ row: 3, col: 1, cell: { t: 's', v: '新' } }])]));
  const items = ['<si><t>plain</t></si>', '<si><r><rPr><b/><sz val="11"/></rPr><t>Bold</t></r><r><t xml:space="preserve"> normal</t></r></si>', '<si><t>山田</t><rPh sb="0" eb="2"><t>ヤマダ</t></rPh><phoneticPr fontId="1"/></si>', '<si><t xml:space="preserve">  lead</t></si>', '<si><t/></si>'];
  const rich = await edit(await buildXlsx([{ name: 'S', data: '<row r="1"><c r="A1" t="s"><v>1</v></c><c r="B1" t="s"><v>2</v></c><c r="C1" t="s"><v>3</v></c></row>' }], { sst: ['x'] }), (p) => p.set('xl/sharedStrings.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="4" uniqueCount="${items.length}">${items.join('')}</sst>`));
  save('h-rich-sst', await replayLog(rich, [cells('S', [{ row: 2, col: 1, cell: { t: 's', v: 'plain' } }, { row: 2, col: 2, cell: { t: 's', v: '新' } }, { row: 2, col: 3, cell: { t: 's', v: 'a<b>&c' } }]), movC('S', 3, 1, 1)]));
  const spans = await buildXlsx([{ name: 'S', data: '<row r="2" spans="3:3" ht="30" customHeight="1"><c r="C2"><v>1</v></c></row><row r="1" spans="1:2"><c r="B1" t="s"><v>1</v></c><c r="A1" t="s"><v>0</v></c><c r="D1"><v>9</v></c></row>' }], { sst: ['a', 'b'] });
  save('h-spans', await replayLog(await edit(spans, (p) => p.set(SH1, text(p, SH1).replace(/<dimension[^>]*\/>/, ''))), [insC('S', 2)]));
  save('h-plain', await replayLog(await base(), [movC('S', 3, 1, 1), insC('S', 2), cells('S', [{ row: 4, col: 1, cell: { t: 's', v: 'x' } }, { row: 4, col: 2, cell: { t: 'd', v: 46295, numFmt: 'yyyy/m/d' } }])]));
}

// ---- fz-*: accepted mutants that only change the spelling -----------------------------------------------------------
{
  const targets: Array<[string, string]> = [['g-table.xlsx', 'g-ins-D'], ['r2-table.xlsx', 'h2-totals-ops'], ['torture.xlsx', 't-seq'], ['r3-misc.xlsx', 'u-band']];
  let made = 0;
  for (let seed = 0; seed < 400 && made < 6; seed++) {
    const [file, cname] = targets[seed % targets.length];
    const c = CASES.find((x) => x.name === cname);
    const ops = c ? c.ops : [];
    const m = await mutate(fixture(file), seed);
    if (!/^(bomPart|crlfPart|singleQuotes)@/.test(m.what) || m.what.includes(' + ')) continue;
    try {
      save(`fz-${seed}-${m.what.split('@')[0]}`, await replayLog(m.bytes, ops));
      made++;
    } catch { /* refused: fine */ }
  }
}

// ---- fo-*: random op logs (fuzzOps) on the intact fixtures: the first accepted ones per fixture ------------------------
{
  const targets = replayTargets();
  const made = new Map<string, number>();
  for (let seed = 0; seed < 600; seed++) {
    const t = targets[seed % targets.length];
    if ((made.get(t.file) ?? 0) >= 2) continue;
    const bytes = fixture(t.file);
    const log = randomLog(rng(seed * 17 + 3), await sheetNames(bytes));
    if (log.length < 3) continue;
    try {
      save(`fo-${seed}-${t.file.replace(/\.xlsx$/, '')}`, await replayLog(bytes, log));
      made.set(t.file, (made.get(t.file) ?? 0) + 1);
    } catch { /* refused: fine */ }
  }
}
