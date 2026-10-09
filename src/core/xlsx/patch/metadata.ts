/**
 * Dynamic-array cell metadata (owner: patcher P).
 *
 * A spilling formula is stored as <c cm="N"><f t="array" ref="A1">…</f></c> where N (1-based) points at a
 * <cellMetadata> block of xl/metadata.xml whose record is the XLDAPR future-metadata entry
 * {fDynamic=1, fCollapsed=0}. The XML below is exactly what Excel 365 writes (inspected in a workbook saved by Excel
 * with spilled formulas: storage-test.xlsx of research R2, and torture.xlsx). An existing metadata part is merged:
 * other metadata types (rich values XLRICHVALUE, MDX ...) and their indexes are kept, the XLDAPR type / future
 * record / cell block are added only when missing.
 */
import { findAll, findFirst, getAttr, setAttr } from './xml';

export const METADATA_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata';
export const METADATA_CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml';
const XDA_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray';
const DA_EXT_URI = '{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}';
const XLDAPR_TYPE = '<metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/>';

export const NEW_METADATA_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
  `<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xda="${XDA_NS}">` +
  `<metadataTypes count="1">${XLDAPR_TYPE}</metadataTypes>` +
  `<futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="${DA_EXT_URI}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata>` +
  '<cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>';

/** Element order inside <metadata> (CT_Metadata). */
const ORDER = ['metadataTypes', 'metadataStrings', 'mdxMetadata', 'futureMetadata', 'cellMetadata', 'valueMetadata', 'extLst'];

/** Insert `el` before the first existing element that must follow `after` in CT_Metadata order. */
function insertInOrder(xml: string, name: string, el: string): string {
  const idx = ORDER.indexOf(name);
  for (let k = idx + 1; k < ORDER.length; k++) {
    const m = new RegExp(`<${ORDER[k]}\\b`).exec(xml);
    if (m) return xml.slice(0, m.index) + el + xml.slice(m.index);
  }
  return xml.replace(/<\/metadata>\s*$/, el + '</metadata>');
}

/** the xda namespace prefix declared on <metadata> (declared as "xda" when missing) */
function ensureXdaPrefix(xml: string): { xml: string; prefix: string } {
  const open = /<metadata\b[^>]*>/.exec(xml)![0];
  for (const m of open.matchAll(/\sxmlns:([\w.-]+)="([^"]*)"/g)) if (m[2] === XDA_NS) return { xml, prefix: m[1] };
  let p = 'xda';
  for (let n = 2; new RegExp(`\\sxmlns:${p}=`).test(open); n++) p = 'xda' + n;
  return { xml: xml.replace(open, setAttr(open, `xmlns:${p}`, XDA_NS)), prefix: p };
}

/**
 * Make sure `xml` (a metadata part, or null when the package has none) has the dynamic-array cell metadata.
 * Returns the new XML and the 1-based cellMetadata index to use as cm="…".
 */
export function ensureDynamicArrayMetadata(xml: string | null): { xml: string; cm: number } {
  if (!xml || !/<metadata\b/.test(xml)) return { xml: NEW_METADATA_XML, cm: 1 };
  let x = xml;
  // 1) metadata type (1-based index = t of the <rc> record)
  let types = findFirst(x, 'metadataTypes');
  let typeIdx = -1;
  if (types && !types.selfClosing) {
    findAll(types.inner, 'metadataType').forEach((t, i) => { if (typeIdx < 0 && getAttr(t.open, 'name') === 'XLDAPR') typeIdx = i + 1; });
  }
  if (typeIdx < 0) {
    if (types && !types.selfClosing) {
      const n = findAll(types.inner, 'metadataType').length + 1;
      x = x.slice(0, types.start) + setAttr(types.open, 'count', String(n)) + types.inner + XLDAPR_TYPE + '</metadataTypes>' + x.slice(types.end);
      typeIdx = n;
    } else {
      const el = `<metadataTypes count="1">${XLDAPR_TYPE}</metadataTypes>`;
      x = types ? x.slice(0, types.start) + el + x.slice(types.end) : insertInOrder(x, 'metadataTypes', el);
      // existing <rc t> records refer to types by index: a missing <metadataTypes> means there were none
      typeIdx = 1;
    }
    types = findFirst(x, 'metadataTypes');
  }
  // 2) future metadata record {fDynamic=1, fCollapsed=0} (0-based index = v of the <rc> record)
  const pre = ensureXdaPrefix(x);
  x = pre.xml;
  const bkXml = `<bk><extLst><ext uri="${DA_EXT_URI}"><${pre.prefix}:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk>`;
  let futIdx = -1;
  const futs = findAll(x, 'futureMetadata');
  const fut = futs.find((f) => getAttr(f.open, 'name') === 'XLDAPR');
  if (fut && !fut.selfClosing) {
    findAll(fut.inner, 'bk').forEach((b, i) => {
      if (futIdx >= 0) return;
      const p = /<(?:[\w.-]+:)?dynamicArrayProperties\b[^>]*>/.exec(b.inner);
      if (p && /\sfDynamic="(1|true)"/.test(p[0]) && !/\sfCollapsed="(1|true)"/.test(p[0])) futIdx = i;
    });
    if (futIdx < 0) {
      const n = findAll(fut.inner, 'bk').length;
      x = x.slice(0, fut.start) + setAttr(fut.open, 'count', String(n + 1)) + fut.inner + bkXml + '</futureMetadata>' + x.slice(fut.end);
      futIdx = n;
    }
  } else {
    const el = `<futureMetadata name="XLDAPR" count="1">${bkXml}</futureMetadata>`;
    if (fut) x = x.slice(0, fut.start) + el + x.slice(fut.end);
    else if (futs.length) { const last = futs[futs.length - 1]; x = x.slice(0, last.end) + el + x.slice(last.end); }
    else x = insertInOrder(x, 'futureMetadata', el);
    futIdx = 0;
  }
  // 3) cell metadata block with exactly that record
  const rc = `<rc t="${typeIdx}" v="${futIdx}"/>`;
  const cm = findFirst(x, 'cellMetadata');
  if (cm && !cm.selfClosing) {
    const bks = findAll(cm.inner, 'bk');
    for (let i = 0; i < bks.length; i++) {
      const rcs = [...bks[i].inner.matchAll(/<rc\b[^>]*>/g)];
      if (rcs.length === 1 && getAttr(rcs[0][0], 't') === String(typeIdx) && getAttr(rcs[0][0], 'v') === String(futIdx)) return { xml: x, cm: i + 1 };
    }
    x = x.slice(0, cm.start) + setAttr(cm.open, 'count', String(bks.length + 1)) + cm.inner + `<bk>${rc}</bk>` + '</cellMetadata>' + x.slice(cm.end);
    return { xml: x, cm: bks.length + 1 };
  }
  const el = `<cellMetadata count="1"><bk>${rc}</bk></cellMetadata>`;
  x = cm ? x.slice(0, cm.start) + el + x.slice(cm.end) : insertInOrder(x, 'cellMetadata', el);
  return { xml: x, cm: 1 };
}
