/**
 * Comments (notes) of a copied sheet (task A2, see addsheets-objects.ts).
 *
 * A sheet's notes are two parts: commentsN.xml (authors, the text of every comment by cell) and vmlDrawingN.vml (one shape per note: its
 * box, size and visibility). Both are copied; the VML is cleaned of everything that is not a note (form controls, ActiveX, OLE objects
 * live in the same file but are not copied) and numbered anew (<o:idmap data> and the shape ids are workbook-unique in files written by
 * Excel). The authors stay as they are.
 *
 * Threaded comments (the newer kind: xl/threadedComments + xl/persons) are not copied as such: Excel stores a plain note next to each of
 * them as a fallback ("[Threaded comment] Your version of Excel ...") and a reader that does not know the newer kind shows that. The copy
 * writes a clean note instead: the comment text of the thread (the replies below it), the name of the person who started it as the author.
 */
import { esc, findAll, findFirst, getAttr, nsPrefix, setAttr, unesc } from './xml';
import type { SheetObjectsContext } from './copytypes';
import { isStrict } from './bookinfo';
import { CT, REL_MS, freshPart, partFlavour, relIs, readSrcRels, relType, type ObjState, type OutRel, type Stage } from './addsheets-objects-core';

export interface CommentCopy {
  rels: OutRel[];
  legacyDrawingId: string | null;
  notes: number;
  threaded: number;
  /** shapes of the VML that are not notes (form controls ...) and were left out */
  otherShapes: number;
}

interface Thread {
  text: string;
  replies: Array<{ text: string; person: string }>;
  person: string;
}

async function readThreads(ctx: SheetObjectsContext): Promise<Map<string, Thread>> {
  const out = new Map<string, Thread>();
  const eng = ctx.source.eng;
  const rel = ctx.source.rels.find((r) => !r.external && r.type === REL_MS.threadedComment) ?? ctx.source.rels.find((r) => !r.external && /threadedComment$/i.test(r.type));
  if (!rel || !rel.path) return out;
  const x = await eng.pkg.read(rel.path);
  if (!x) return out;
  // persons: workbook relationship
  const persons = new Map<string, string>();
  const wbRel = (await readSrcRels(eng, eng.wbPath)).find((r) => !r.external && /\/person$/i.test(r.type));
  const px = wbRel?.path ? await eng.pkg.read(wbRel.path) : null;
  for (const m of (px ?? '').matchAll(/<person\b[^>]*>/g)) {
    const id = getAttr(m[0], 'id');
    if (id) persons.set(id.toLowerCase(), getAttr(m[0], 'displayName') ?? '');
  }
  const items = [...x.matchAll(/<threadedComment\b([^>]*?)(?:\/>|>([\s\S]*?)<\/threadedComment>)/g)].map((m) => ({
    id: (getAttr(m[1], 'id') ?? '').toLowerCase(),
    parent: (getAttr(m[1], 'parentId') ?? '').toLowerCase(),
    person: persons.get((getAttr(m[1], 'personId') ?? '').toLowerCase()) ?? '',
    text: unesc(/<text\b[^>]*>([\s\S]*?)<\/text>/.exec(m[2] ?? '')?.[1] ?? ''),
  }));
  for (const t of items) if (!t.parent) out.set(t.id, { text: t.text, person: t.person, replies: [] });
  for (const t of items) if (t.parent) out.get(t.parent)?.replies.push({ text: t.text, person: t.person });
  return out;
}

/** the notes of the sheet: comments part + VML, threaded comments turned into notes */
export async function copyComments(ctx: SheetObjectsContext, st: ObjState, stage: Stage): Promise<CommentCopy | null> {
  const eng = ctx.source.eng;
  const rels = ctx.source.rels;
  const cRel = rels.find((r) => !r.external && r.path && relIs(r, 'comments'));
  const legacy = /<legacyDrawing\b[^>]*\sr:id="([^"]*)"/.exec(ctx.source.xml)?.[1];
  const vRel = rels.find((r) => r.id === legacy && !r.external && r.path) ?? rels.find((r) => !r.external && r.path && relIs(r, 'vmlDrawing'));
  if (!cRel || !vRel) return null;
  const cxRaw = await eng.pkg.read(cRel.path);
  const vml = await eng.pkg.read(vRel.path);
  if (!cxRaw || !vml) return null;
  const srcStrict = isStrict(eng.wbXml);
  // Strict into Strict: the notes are copied as they are (comments part with <commentPr>, shape ids, idmap): Excel's Strict reader ties every
  // comment to its VML shape through shapeId and repairs the file (it removes the comments) when the numbers are not the ones it wrote — measured
  // with tests/golden/com/check-addsheets-objects.ps1 on variants of one workbook. Everywhere else the comment carries shapeId 0 (a transitional
  // package never ties them) and the box of the note is the VML shape's.
  const verbatim = srcStrict && st.strict;
  const cx0 = verbatim ? cxRaw : partFlavour(cxRaw, srcStrict, st.strict).replace(/<mc:AlternateContent\b[^>]*>(?:(?!<\/mc:AlternateContent>)[\s\S])*?<commentPr\b[\s\S]*?<\/mc:AlternateContent>/g, '');
  const cx = verbatim || /<xdr:/.test(cx0) ? cx0 : cx0.replace(/\sxmlns:xdr="[^"]*"/, '');

  // ---- the VML: notes only ----
  const vp = nsPrefix(vml, 'urn:schemas-microsoft-com:vml', 'v:');
  const xp = nsPrefix(vml, 'urn:schemas-microsoft-com:office:excel', 'x:');
  const op = nsPrefix(vml, 'urn:schemas-microsoft-com:office:office', 'o:');
  const shapeRe = new RegExp(`<${vp}shape\\b[^>]*?(?:/>|>[\\s\\S]*?</${vp}shape>)`, 'g');
  const noteRe = new RegExp(`<${xp}ClientData\\b[^>]*\\sObjectType="Note"`);
  let otherShapes = 0;
  const shapes: string[] = [];
  for (const m of vml.matchAll(shapeRe)) {
    if (noteRe.test(m[0])) shapes.push(m[0]);
    else otherShapes++;
  }
  if (!shapes.length) return null;
  // shapetypes that a note shape uses
  const types = new Set(shapes.map((s) => /\stype="#([^"]*)"/.exec(s)?.[1]).filter((x): x is string => !!x));
  const shapetypes = [...vml.matchAll(new RegExp(`<${vp}shapetype\\b[^>]*?(?:/>|>[\\s\\S]*?</${vp}shapetype>)`, 'g'))].map((m) => m[0]).filter((s) => types.has(/\sid="([^"]*)"/.exec(s)?.[1] ?? ''));
  if (!verbatim) st.vmlMap += 1;
  const map = st.vmlMap;
  let k = 0;
  const newShapes = verbatim ? shapes : shapes.map((s) => {
    const id = `_x0000_s${map * 1024 + 1 + k++}`;
    return s.replace(/(\s(?:[\w.-]+:)?(?:id|spid)=")_x0000_s\d+(")/g, `$1${id}$2`);
  });
  const xmlOpen = /<xml\b[^>]*>/.exec(vml)?.[0] ?? '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">';
  const layout = verbatim
    ? new RegExp(`<${op}shapelayout\\b[^>]*?(?:/>|>[\\s\\S]*?</${op}shapelayout>)`).exec(vml)?.[0] ?? ''
    : `<${op}shapelayout ${vp}ext="edit"><${op}idmap ${vp}ext="edit" data="${map}"/></${op}shapelayout>`;
  const newVml = `${xmlOpen}\r\n ${layout}${shapetypes.join('')}${newShapes.join('')}</xml>`;

  // ---- the comments ----
  const threads = await readThreads(ctx);
  const authors = (findFirst(cx, 'authors')?.inner.match(/<author\b[^>]*?(?:\/>|>[\s\S]*?<\/author>)/g) ?? []).map((a) => unesc(a.replace(/^<author\b[^>]*>|<\/author>$/g, '')));
  const list = findFirst(cx, 'commentList');
  const comments = list ? findAll(list.inner, 'comment') : [];
  const outAuthors: string[] = [];
  const authorIndex = (name: string): number => {
    let i = outAuthors.indexOf(name);
    if (i < 0) { outAuthors.push(name); i = outAuthors.length - 1; }
    return i;
  };
  let threaded = 0;
  const outComments = comments.map((c) => {
    const ai = Number(getAttr(c.open, 'authorId') ?? 0);
    const author = authors[ai] ?? '';
    const th = /^tc=\{?([^}]*)\}?$/.exec(author);
    let open = c.open.replace(/\sxr\d*:uid="[^"]*"/g, '');
    let body = c.selfClosing ? '' : c.inner;
    let name = author;
    const thread = th ? threads.get(`{${th[1]}}`.toLowerCase()) ?? threads.get(th[1].toLowerCase()) : undefined;
    if (thread) {
      threaded++;
      name = thread.person || author;
      const lines = [thread.text, ...thread.replies.map((r) => `返信（${r.person || '不明'}）: ${r.text}`)];
      body = `<text><t xml:space="preserve">${esc(lines.join('\n'))}</t></text>`;
    }
    open = setAttr(open.replace(/\sauthorId="[^"]*"/, ` authorId="${authorIndex(name)}"`), 'shapeId', verbatim ? (getAttr(c.open, 'shapeId') ?? '0') : '0');
    return c.selfClosing ? open : `${open}${body}</comment>`;
  });
  if (!outComments.length) return null;
  const root = /<comments\b[^>]*>/.exec(cx)?.[0] ?? '<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">';
  const newCx = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n${root.replace(/\sxr\d*:uid="[^"]*"/g, '')}<authors>${outAuthors.map((a) => `<author>${esc(a)}</author>`).join('')}</authors><commentList>${outComments.join('')}</commentList></comments>`;

  const cPath = freshPart(st, 'xl/', 'comments', '.xml');
  const vPath = freshPart(st, 'xl/drawings/', 'vmlDrawing', '.vml');
  stage.text(cPath, newCx, CT.comments);
  stage.text(vPath, newVml);
  stage.defaultType('vml', CT.vml);
  const legacyId = ctx.newRelId();
  const out: OutRel[] = [
    { id: ctx.newRelId(), type: relType(st, 'comments'), target: cPath },
    { id: legacyId, type: relType(st, 'vmlDrawing'), target: vPath },
  ];
  return { rels: out, legacyDrawingId: legacyId, notes: outComments.length, threaded, otherShapes };
}

