// Producer-shaped rewrites of a package (round 5, module H1): the same workbook spelled the way other writers spell it. Every mutation must
// be semantically neutral; tests/xlsx/foreign-mutations.test.ts runs the pipeline on the mutated and on the original package with one op log
// and demands the same output. (tests/xlsx/patch-fuzz.test.ts has the generic spellings: x: prefix, single quotes, BOM, UTF-16.)
import JSZip from 'jszip';

export interface Mutation {
  name: string;
  /** what writer does this */
  like: string;
  /** rewrite one XML part (path, text) -> text; return the text itself to leave it alone */
  part?: (path: string, xml: string, all: Map<string, string>) => string;
  /** zip-level options (store, data descriptors, folder entries, entry order) */
  zip?: { store?: boolean; streamFiles?: boolean; reverse?: boolean; folders?: boolean; comment?: string; upper?: boolean };
}

const isSheet = (p: string) => /^xl\/worksheets\/sheet\d+\.xml$/i.test(p);
const BOOL_ATTRS = 'hidden|customWidth|customHeight|bestFit|collapsed|customFormat|showGridLines|tabSelected|showRowColHeaders|showZeros|defaultGridColor';

/** `<c ...>` / `<row ...>` start tags of a sheet, rewritten by `fn` (attribute string without the tag name, selfClosing flag) */
function mapTags(xml: string, tag: 'c' | 'row', fn: (attrs: string, selfClosing: boolean) => string): string {
  return xml.replace(new RegExp(`<${tag}(\\s[^>]*?)?(/?)>`, 'g'), (_m, attrs: string | undefined, sc: string) => `<${tag}${fn(attrs ?? '', sc === '/')}${sc}>`);
}

export const MUTATIONS: Mutation[] = [
  {
    name: 'boolean words',
    like: 'LibreOffice / Apache POI: hidden="true" customWidth="false" instead of "1" / "0"',
    part: (p, x) => (isSheet(p) ? x.replace(new RegExp(`(\\s(?:${BOOL_ATTRS})=)"1"`, 'g'), '$1"true"').replace(new RegExp(`(\\s(?:${BOOL_ATTRS})=)"0"`, 'g'), '$1"false"') : x),
  },
  {
    name: 'explicit t="n"',
    like: 'LibreOffice: every numeric cell says t="n"',
    part: (p, x) => (isSheet(p) ? mapTags(x, 'c', (a) => (/\st=/.test(a) ? a : a + ' t="n"')) : x),
  },
  {
    name: 'inline strings instead of shared strings',
    like: 'XlsxWriter constant_memory, Apache POI SXSSF, ExcelJS streaming, SheetJS',
    part: (p, x, all) => {
      if (!isSheet(p)) return x;
      const sst = all.get('xl/sharedStrings.xml');
      if (!sst) return x;
      const items = [...sst.matchAll(/<si\b[^>]*?(?:\/>|>([\s\S]*?)<\/si>)/g)].map((m) => m[1] ?? '');
      return x.replace(/<c(\s[^>]*?)\st="s"([^>]*)>\s*<v>(\d+)<\/v>\s*<\/c>/g, (m0, a: string, b: string, i: string) => (items[Number(i)] === undefined ? m0 : `<c${a}${b} t="inlineStr"><is>${items[Number(i)]}</is></c>`));
    },
  },
  {
    name: 'expanded empty elements, whitespace and comments between elements',
    like: 'pretty printers, DOM serializers (.NET XmlDocument, Python minidom)',
    part: (p, x) => {
      if (!isSheet(p) && p !== 'xl/workbook.xml' && p !== 'xl/styles.xml') return x;
      return x
        .replace(/<(c|row|col|sheet|xf|font|fill|border|numFmt)(\s[^>]*?)?\/>/g, (_m, t: string, a: string | undefined) => `<${t}${a ?? ''}></${t}>`)
        .replace(/(<\/[\w:.-]+>|\/>)(?=<)/g, '$1\r\n    ')
        .replace(/(<\/row>)/g, '$1<!-- row -->');
    },
  },
  {
    name: 'attribute order reversed',
    like: 'writers that emit r / s / t in another order (Open XML SDK: r, s, t; Google: s, t, r)',
    part: (p, x) => {
      if (!isSheet(p)) return x;
      const rev = (a: string) => [...a.matchAll(/\s[^\s=]+\s*=\s*"[^"]*"/g)].map((m) => m[0]).reverse().join('');
      return mapTags(mapTags(x, 'c', rev), 'row', rev);
    },
  },
  {
    name: 'numeric character references for every non-ASCII character',
    like: 'Python ElementTree ascii output, PHP htmlspecialchars pipelines, old Java serializers',
    part: (p, x) => (/\.xml$/i.test(p) && /sharedStrings|worksheets\/sheet|workbook\.xml/.test(p) ? x.replace(/>([^<]+)</g, (_m, t: string) => `>${[...t].map((ch) => (ch.codePointAt(0)! > 0x7f ? `&#x${ch.codePointAt(0)!.toString(16)};` : ch)).join('')}<`) : x),
  },
  {
    name: 'lower-case cell references',
    like: 'hand-rolled writers: r="a1"',
    part: (p, x) => (isSheet(p) ? mapTags(x, 'c', (a) => a.replace(/(\sr=")([A-Z]+)(\d+")/, (_m, a1: string, col: string, rest: string) => a1 + col.toLowerCase() + rest)) : x),
  },
  {
    name: 'no dimension, no spans, a wrong dimension',
    like: 'Google Sheets, POI SXSSF (dimension A1), streaming writers',
    part: (p, x) => (isSheet(p) ? x.replace(/<dimension\b[^>]*\/>/, '<dimension ref="A1"/>').replace(/\sspans="[^"]*"/g, '') : x),
  },
  {
    name: 'integers written with a decimal point',
    like: 'Google Sheets / Apache POI: <v>85.0</v>, width="18.0"',
    part: (p, x) => (isSheet(p) ? x.replace(/<c(\s(?![^>]*\st=")[^>]*)><v>(\d+)<\/v><\/c>/g, (m0, a: string, n: string) => (n.length < 9 ? `<c${a}><v>${n}.0</v></c>` : m0)).replace(/(\swidth=")(\d+)"/g, '$1$2.0"') : x),
  },
  {
    name: 'stored zip, data descriptors, folder entries, [Content_Types].xml last, zip comment',
    like: '.NET System.IO.Packaging, Java ZipOutputStream, hand zipped packages',
    zip: { store: true, streamFiles: true, reverse: true, folders: true, comment: 'made by another tool' },
  },
  {
    name: 'every sheet name quoted in formulas',
    like: 'XlsxWriter / openpyxl / ExcelJS: \'名簿\'!$E$2 even when no quotes are needed',
    part: (p, x, all) => {
      if (!/^xl\/(worksheets\/sheet\d+|workbook)\.xml$/i.test(p)) return x;
      const names = [...(all.get('xl/workbook.xml') ?? '').matchAll(/<sheet\b[^>]*?\sname="([^"]*)"/g)].map((m) => m[1]).filter((n) => /^[^\s'!&<>"]+$/.test(n));
      let out = x;
      for (const n of names) {
        const re = new RegExp(`(<f(?:\\s[^>]*)?>|<formula>|<formula1>|<formula2>|<definedName\\b[^>]*>|[,(+\\-*/=&:\\s])${n.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}!`, 'g');
        out = out.replace(re, (_m, pre: string) => `${pre}'${n}'!`);
      }
      return out;
    },
  },
  {
    name: 'references to the own sheet written with the sheet name',
    like: 'writers that qualify every reference (SheetJS, Apache POI formula evaluators)',
    part: (p, x, all) => {
      const m = /^xl\/worksheets\/sheet(\d+)\.xml$/i.exec(p);
      if (!m) return x;
      const sheets = [...(all.get('xl/workbook.xml') ?? '').matchAll(/<sheet\b[^>]*?\sname="([^"]*)"/g)].map((s) => s[1]);
      const name = sheets[Number(m[1]) - 1];
      if (!name || !/^[^\s'!&<>"]+$/.test(name)) return x;
      // only plain arithmetic / SUM style formulas (no strings, no other sheets, no names): every cell reference gets the sheet name
      return x.replace(/<f>([A-Z0-9$:+\-*/(),.\s]+)<\/f>/g, (m0, f: string) => (/!/.test(f) ? m0 : `<f>${f.replace(/(?<![A-Z$0-9.])(\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)(?![A-Z0-9(])/g, `'${name}'!$1`)}</f>`));
    },
  },
  {
    name: 'absolute relationship targets',
    like: 'Google Sheets, some PHP / Java writers: Target="/xl/worksheets/sheet1.xml"',
    part: (p, x) => {
      if (!/\.rels$/i.test(p) || p === '_rels/.rels') return x;
      const base = p.replace(/\/_rels\/[^/]*$/, '');
      return x.replace(/(Target=")(?!\/)([^"]*)(")(?![^>]*TargetMode="External")/g, (m0, a: string, t: string, c: string) => {
        if (/^[a-z]+:/i.test(t) || t.startsWith('#')) return m0;
        const parts = base.split('/');
        for (const seg of t.split('/')) { if (seg === '..') parts.pop(); else if (seg !== '.') parts.push(seg); }
        return `${a}/${parts.join('/')}${c}`;
      });
    },
  },
  {
    name: 'upper-case part names in the zip',
    like: 'Windows tools that upper-case file names (OPC part names are case-insensitive)',
    zip: { upper: true },
  },
];

/** The package with the mutation applied (text parts are rewritten, everything else is carried over). */
export async function mutate(bytes: Uint8Array, m: Mutation): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(bytes);
  const names = Object.keys(zip.files).filter((n) => !zip.files[n].dir);
  const text = new Map<string, string>();
  for (const n of names) if (/\.(xml|rels)$/i.test(n)) text.set(n, await zip.file(n)!.async('string'));
  const out = new JSZip();
  const order = m.zip?.reverse ? names.slice().reverse() : names;
  if (m.zip?.folders) for (const d of new Set(order.map((n) => n.replace(/[^/]*$/, '')).filter(Boolean))) out.folder(d.replace(/\/$/, ''));
  for (const n of order) {
    const t = text.get(n);
    const target = m.zip?.upper && n !== '[Content_Types].xml' ? n.toUpperCase() : n;
    if (t === undefined) out.file(target, await zip.file(n)!.async('uint8array'));
    else out.file(target, m.part ? m.part(n, t, text) : t);
  }
  return out.generateAsync({ type: 'uint8array', compression: m.zip?.store ? 'STORE' : 'DEFLATE', streamFiles: !!m.zip?.streamFiles, comment: m.zip?.comment });
}
