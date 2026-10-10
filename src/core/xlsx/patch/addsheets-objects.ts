/**
 * The object phase of the sheet copier (task A2, round 9 plan A): everything a source sheet owns besides its cells — drawings (charts,
 * pictures, shapes, text boxes, groups, connectors), notes, tables, pivot tables, the sheet background — and chart sheets.
 *
 * addsheets.ts calls two functions of this file: planSheetObjects (before the formulas of a source are copied: which tables are copied under
 * which names, so that structured references can stay formulas) and copySheetObjects (after the core copy of every sheet: the new sheet XML
 * and relationships are in `ctx`). The work is split by kind:
 *   addsheets-objects-core.ts      state of the target workbook, relationships, unique part names, content types, staging (an object that
 *                                  fails half way leaves nothing behind)
 *   addsheets-objects-drawing.ts   the drawing part, anchor by anchor
 *   addsheets-objects-chart.ts     charts (formulas follow the copies, literal data for what cannot follow, theme colours / fonts)
 *   addsheets-objects-table.ts     tables, custom table styles
 *   addsheets-objects-pivot.ts     pivot tables and their caches
 *   addsheets-objects-comments.ts  notes (comments + VML), threaded comments turned into notes
 *
 * What is NOT copied (named, with the sheet and a count, in the sheet's one 'sheetCopyLimits' warning): slicers and timelines, SmartArt,
 * ink, form controls / ActiveX / OLE objects, header / footer pictures (left to addsheets.ts: they stay in ctx.present), pivot tables whose
 * data is not a range of a copied sheet (they stay as cell values), XML-mapped and query tables (they stay as cells). Nothing is ever left
 * half done: no dangling r:id, no orphan part, every part has its content type.
 */
import { isStrict } from './bookinfo';
import type { Engine } from './engine';
import { PatchError } from './pkg';
import { findAll, findFirst, getAttr } from './xml';
import type { ObjectKind, SheetCopy, SheetObjectsContext } from './copytypes';
import { copyComments } from './addsheets-objects-comments';
import { addInfo, collectTakenNames, getState, readSrcRels, relIs, relType, takeInfo, Stage, type ObjState, type OutRel } from './addsheets-objects-core';
import { copyDrawing, splitChildren } from './addsheets-objects-drawing';
import { copyPivots } from './addsheets-objects-pivot';
import { copyTables, planTables, type TablePlan } from './addsheets-objects-table';

export { collectTakenNames, planTables, takeInfo as takeObjectInfo };
export type { TablePlan };

/** an unexpected defect while copying one object: turned into a warning for the user, but loud in the tests */
const FATAL = typeof process !== 'undefined' && !!process.env && process.env.VITEST !== undefined;

async function guarded<T>(ctx: SheetObjectsContext, st: ObjState, what: string, fn: (stage: Stage) => Promise<T>): Promise<T | null> {
  const stage = new Stage(st);
  try {
    const r = await fn(stage);
    await stage.commit();
    return r;
  } catch (e) {
    stage.discard();
    if (FATAL && !(e instanceof PatchError)) throw e; // a defect of the copier is loud in the tests; a damaged part of the source is a note
    const why = e instanceof PatchError ? `（${e.message}）` : '';
    ctx.note(`${what}をコピーできませんでした${why}`);
    return null;
  }
}

const pushRels = (ctx: SheetObjectsContext, rels: readonly OutRel[]): void => {
  for (const r of rels) ctx.rels.push({ id: r.id, type: r.type, target: r.target, external: !!r.external });
};

/**
 * Kind `kind` of the sheet is copied except for `label` (what was left out): the caller reports `label` instead of the whole kind. Without a
 * label the kind is copied completely.
 */
function leave(ctx: SheetObjectsContext, kind: ObjectKind, label: string | null, count = 1): void {
  if (label === null) { ctx.handled.add(kind); return; }
  const i = ctx.present.findIndex((o) => o.kind === kind);
  if (i >= 0) ctx.present[i] = { kind, label, count };
  else ctx.present.push({ kind, label, count });
}

/** the source and the target are written in different flavours of OOXML (Strict / transitional) */
const mixedFlavours = (ctx: SheetObjectsContext, st: ObjState): boolean => isStrict(ctx.source.eng.wbXml) !== st.strict;

const isChartSheet = (xml: string): boolean => /^\s*(?:<\?xml[^>]*\?>\s*)?(?:<!--[\s\S]*?-->\s*)*<chartsheet\b/.test(xml);

/**
 * Copies what the source sheet owns besides its cells into the new sheet (ctx.xml / ctx.rels) and the target package. Kinds that were
 * copied (completely, or with a note for the part that was left out) are marked in ctx.handled; the rest stays reported by the caller.
 */
export async function copySheetObjects(ctx: SheetObjectsContext): Promise<void> {
  const st = await getState(ctx.target.eng);
  if (isChartSheet(ctx.source.xml)) return copyChartSheetObjects(ctx, st);
  const src = ctx.source.xml;
  const srcRels = ctx.source.rels;
  // What the sheet owns is read from its XML and relationships, not from ctx.present (which only lists what the core copy saw).
  const has = (tag: string): boolean => new RegExp(`<${tag}(?=[\\s/>])`).test(src);
  const present = new Set(ctx.present.map((o) => o.kind));

  // ---- pivot tables first: pivot charts follow their copies ----
  let pivotNames = new Map<string, string>();
  if (srcRels.some((r) => !r.external && relIs(r, 'pivotTable'))) {
    const r = await guarded(ctx, st, 'ピボットテーブル', async (stage) => copyPivots(ctx, st, stage));
    if (r) {
      pivotNames = r.names;
      pushRels(ctx, r.rels);
      leave(ctx, 'pivotTables', r.values ? `ピボットテーブル ${r.values} 個（元のデータを参照できないため、セルの値だけを残しました）` : null, r.values);
    }
  }

  // ---- tables ----
  if (has('tableParts')) {
    const tp = findFirst(src, 'tableParts');
    const rids = tp ? findAll(tp.inner, 'tablePart').map((t) => getAttr(t.open, 'r:id') ?? '').filter(Boolean) : [];
    const r = await guarded(ctx, st, 'テーブル', async (stage) => copyTables(ctx, st, stage, rids));
    if (r) {
      if (r.parts.length) {
        pushRels(ctx, r.parts.map((p) => p.rel));
        ctx.insertElement('tableParts', `<tableParts count="${r.parts.length}">${r.parts.map((p) => `<tablePart r:id="${p.rel.id}"/>`).join('')}</tableParts>`);
      }
      for (const n of r.renamed) {
        addInfo(ctx.target.eng, { severity: 'info', code: 'tableRenamed', sourceId: ctx.source.ks.id, message: `シート「${ctx.target.name}」のテーブル「${n.from}」は、名前が重なったため「${n.to}」にしました（数式の参照も付け替えました）。` });
      }
      leave(ctx, 'tables', r.skipped ? `テーブル ${r.skipped} 個（外部データ接続・XML 対応のため、テーブルの機能は外し、セルの値と書式だけを残しました）` : null, r.skipped);
    }
  }

  // ---- drawing ----
  // (a Strict OOXML package and a transitional one differ in the values of DrawingML too — percentages, angles — so charts, pictures and shapes are
  // copied between packages of the same kind only)
  if (has('drawing') && mixedFlavours(ctx, st)) {
    leave(ctx, 'drawing', 'グラフ・画像・図形（Strict 形式のファイルと通常の形式のファイルの間では写せません）', 1);
  } else if (has('drawing')) {
    const dEl = /<drawing\b[^>]*\sr:id="([^"]*)"/.exec(src);
    const rel = srcRels.find((r) => r.id === dEl?.[1] && !r.external && r.path);
    if (rel) {
      const d = await guarded(ctx, st, '図形・画像・グラフ', async (stage) => copyDrawing(ctx, st, stage, rel.path, { pivotNames }));
      if (d) {
        if (d.path) {
          const id = ctx.newRelId();
          ctx.rels.push({ id, type: relType(st, 'drawing'), target: d.path, external: false });
          ctx.insertElement('drawing', `<drawing r:id="${id}"/>`);
        }
        const lost: string[] = [];
        if (d.chartsLost) lost.push(`グラフ ${d.chartsLost} 個`);
        if (d.slicers) lost.push(`スライサー・タイムライン ${d.slicers} 個`);
        if (d.diagrams) lost.push(`SmartArt ${d.diagrams} 個`);
        if (d.other) lost.push(`その他の図 ${d.other} 個`);
        if (d.literal) ctx.note(`グラフの参照先をコピーできなかった系列・見出し ${d.literal} 個は、元の値のまま残しました（セルの変更には追従しません）`);
        leave(ctx, 'drawing', lost.length ? lost.join('、') : null, d.chartsLost + d.slicers + d.diagrams + d.other);
        if (d.slicers) ctx.handled.add('slicers'); // named with the drawing
      }
    }
  }

  // ---- notes ----
  if (srcRels.some((r) => !r.external && (relIs(r, 'comments') || relIs(r, 'vmlDrawing')))) {
    const c = await guarded(ctx, st, 'コメント', async (stage) => copyComments(ctx, st, stage));
    if (c) {
      pushRels(ctx, c.rels);
      if (c.legacyDrawingId) ctx.insertElement('legacyDrawing', `<legacyDrawing r:id="${c.legacyDrawingId}"/>`);
      if (c.threaded) ctx.note(`スレッドコメント ${c.threaded} 件は、通常のコメント（メモ）にしました`);
      ctx.handled.add('comments');
      if (!c.otherShapes) ctx.handled.add('legacyDrawing');
      // the form controls of the VML and the <controls> of the sheet are the same objects: named once (as form controls)
      else if (present.has('controls') || has('controls')) ctx.handled.add('controls');
    }
  }

  // Strict OOXML wraps <legacyDrawing> (the notes) in an <mc:AlternateContent>: the core copy takes it for a control
  if (!has('controls') && /<mc:AlternateContent[^>]*>(?:(?!<\/mc:AlternateContent>)[\s\S])*?<legacyDrawing\b/.test(src)) ctx.handled.add('controls');

  // ---- the sheet background ----
  if (has('picture')) {
    const pic = /<picture\b[^>]*\sr:id="([^"]*)"/.exec(src);
    const rel = srcRels.find((r) => r.id === pic?.[1] && !r.external && r.path && relIs(r, 'image'));
    if (rel) {
      const id = await guarded(ctx, st, '背景画像', async (stage) => {
        const ext = (/\.([^./]+)$/.exec(rel.path)?.[1] ?? 'png').toLowerCase();
        const path = await stage.copyBinary(ctx.source.eng, rel.path, 'xl/media/', 'image', `.${ext}`, true);
        if (!path) return null;
        const nid = ctx.newRelId();
        ctx.rels.push({ id: nid, type: relType(st, 'image'), target: path, external: false });
        return nid;
      });
      if (id) {
        ctx.insertElement('picture', `<picture r:id="${id}"/>`);
        ctx.handled.add('background');
      }
    }
  }
}

// ---------------------------------------------------------------------------------------------------------------------
// chart sheets
// ---------------------------------------------------------------------------------------------------------------------

/** the worksheets part of a chart sheet is copied as a chart sheet: sheet properties / views / print setup + its drawing */
export function prepareChartSheet(xml: string): SheetCopy | { problem: string } {
  const open = /<chartsheet\b[^>]*>/.exec(xml);
  const close = /<\/chartsheet>\s*$/.exec(xml);
  if (!open || !close) return { problem: 'グラフシートの内容を読み取れませんでした' };
  const prolog = xml.slice(0, open.index);
  const kept: string[] = [];
  const left: SheetCopy['left'] = [];
  for (const c of splitChildren(xml.slice(open.index + open[0].length, close.index))) {
    switch (c.tag) {
      case 'sheetPr': kept.push(c.text.replace(/^<sheetPr\b[^>]*>/, (o) => o.replace(/\scodeName="[^"]*"/, ''))); break;
      case 'sheetViews': kept.push(c.text.replace(/<sheetView\b[^>]*>/g, (o) => o.replace(/\stabSelected="[^"]*"/, ''))); break;
      case 'sheetProtection': case 'pageMargins': case 'headerFooter': kept.push(c.text); break;
      case 'pageSetup': kept.push(c.text.replace(/\sr:id="[^"]*"/, '')); break;
      case 'drawing': left.push({ kind: 'drawing', label: 'グラフ', count: 1 }); break;
      case 'legacyDrawingHF': case 'drawingHF': left.push({ kind: 'headerFooterImages', label: 'ヘッダー・フッターの画像', count: 1 }); break;
      case 'customSheetViews': left.push({ kind: 'customViews', label: 'ユーザー設定のビュー', count: 1 }); break;
      default: break;
    }
  }
  const root = open[0].replace(/\sxr\d*:uid="[^"]*"/g, '');
  return {
    xml: `${prolog}${root}${kept.join('')}</chartsheet>`,
    rels: [], usedNames: new Set(), failures: [], noCache: 0, left, vmCells: 0, xfOverflow: 0,
    dropped: { cf: 0, dv: 0, sparklines: 0, hyperlinks: 0, other: [] }, notes: [], formulaCount: 0,
  };
}

async function copyChartSheetObjects(ctx: SheetObjectsContext, st: ObjState): Promise<void> {
  if (mixedFlavours(ctx, st)) return; // no drawing: the caller makes an empty sheet of it and says so
  const dEl = /<drawing\b[^>]*\sr:id="([^"]*)"/.exec(ctx.source.xml);
  const rel = ctx.source.rels.find((r) => r.id === dEl?.[1] && !r.external && r.path);
  if (!rel) return;
  const d = await guarded(ctx, st, 'グラフ', async (stage) => copyDrawing(ctx, st, stage, rel.path, { pivotNames: new Map() }));
  if (!d) return;
  if (d.path) {
    const id = ctx.newRelId();
    ctx.rels.push({ id, type: relType(st, 'drawing'), target: d.path, external: false });
    const frag = `<drawing r:id="${id}"/>`;
    const at = /<legacyDrawing\b|<legacyDrawingHF\b|<picture\b|<webPublishItems\b|<extLst\b|<\/chartsheet>/.exec(ctx.xml);
    ctx.xml = at ? ctx.xml.slice(0, at.index) + frag + ctx.xml.slice(at.index) : ctx.xml;
  }
  if (d.chartsLost) ctx.note(`グラフ ${d.chartsLost} 個はコピーできませんでした`);
  if (d.literal) ctx.note(`グラフの参照先をコピーできなかった系列・見出し ${d.literal} 個は、元の値のまま残しました（セルの変更には追従しません）`);
  ctx.handled.add('drawing');
}

/** the relationship ids (workbook.xml.rels) of the chart sheets of a workbook: they take part in the naming like worksheets */
export async function chartSheetRids(eng: Engine): Promise<Set<string>> {
  const out = new Set<string>();
  for (const r of await readSrcRels(eng, eng.wbPath)) if (!r.external && r.path && relIs(r, 'chartsheet')) out.add(r.id);
  return out;
}
