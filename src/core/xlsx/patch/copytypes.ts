/**
 * Types shared by the sheet copier (addsheets.ts, copysheet.ts, copynames.ts). See addsheets.ts for the overview.
 */
import type { KeySource } from '../../contracts';
import type { Sst } from './sst';
import type { Styles } from './styles';
import type { Engine, SheetInfo } from './engine';
import type { Rel } from './pkg';
import type { CopyFailReason, FormulaCopier, NameView, Rewrite } from './copyformula';

/** a relationship of a sheet that is created in the output package (hyperlink targets, and whatever copySheetObjects adds) */
export interface NewRel {
  id: string;
  /** full relationship type URL */
  type: string;
  /** a part path inside the package ("xl/drawings/drawing3.xml"), or the external target text when `external` */
  target: string;
  external: boolean;
}

/**
 * Kinds of things a worksheet can own besides its cells. The core copy leaves all of them out and tells copySheetObjects which ones the
 * source sheet had; whatever copySheetObjects does not mark as handled is reported in the 'sheetCopyLimits' warning of the sheet.
 */
export type ObjectKind =
  | 'drawing' // <drawing>: charts, pictures, shapes, text boxes, SmartArt, slicers placed on the sheet
  | 'legacyDrawing' // <legacyDrawing>: notes (comments) and form controls (VML)
  | 'comments' // the comments part / threaded comments
  | 'tables' // <tableParts>
  | 'pivotTables' // pivot tables (relationships of the sheet)
  | 'background' // <picture>
  | 'oleObjects' // <oleObjects>
  | 'controls' // <controls> (ActiveX)
  | 'slicers' // slicers / timelines (extension list)
  | 'headerFooterImages' // <legacyDrawingHF>
  | 'customViews' // <customSheetViews>
  | 'customProperties' // <customProperties>
  | 'other'; // extension elements the copier does not know

export interface ObjectItem {
  kind: ObjectKind;
  /** plain Japanese for the warning ("グラフ 2 個") */
  label: string;
  count: number;
}

/** the user-visible state of a sheet tab */
export type SheetState3 = 'visible' | 'hidden' | 'veryHidden';

/** one sheet of a source workbook with the name it gets in the output */
export interface PlannedSheet {
  info: SheetInfo;
  state: SheetState3;
  /** the planned name; null = the sheet is not copied (chart sheet, macro sheet, dialog sheet) */
  name: string | null;
  original: string;
  suffixed: boolean;
}

export interface SourcePlan {
  ks: KeySource;
  eng: Engine;
  sheets: PlannedSheet[];
}

export interface CopyPlan {
  sources: SourcePlan[];
}

/** everything about the SOURCE package that converting its cells needs */
export interface CopySource {
  ks: KeySource;
  eng: Engine;
  styles: Styles;
  sst: Sst;
  /** source theme colours when they differ from the target's */
  themeRgb: string[] | null;
  /** the source's indexed palette when it differs from the target's */
  indexedRgb: string[] | null;
  /** days to add to date serials (1904 <-> 1900), 0 = same date system */
  dateShift: number;
  /** the source could only be read as values (xls / csv / ods) */
  valuesOnly: boolean;
  /** source xf -> target xf (shared with premerge's imported columns when the source is merged too) */
  xfCache: Map<number, number>;
  /** max digit width (pixels) of the source's default font (column widths are stored in characters of it), undefined when the font is not known */
  mdw: number | undefined;
  /** lower-cased source sheet name -> name in the output */
  rename: Map<string, string>;
}

/** the TARGET package (the premerged base) the copies are written into */
export interface CopyTarget {
  eng: Engine;
  styles: Styles;
  sst: Sst;
  /** Excel's limit of cell formats (cellXfs); a format that would exceed it becomes the default format */
  xfMax: number;
  /** the target package is written in the Strict OOXML namespaces */
  strict: boolean;
  /** max digit width (pixels) of the target's default font, undefined when the font is not known */
  mdw: number | undefined;
}

/** how the formulas of a sheet were treated */
export interface FormulaFailures {
  reason: CopyFailReason;
  count: number;
  examples: string[];
}

/** what the core copy of ONE sheet produced */
export interface SheetCopy {
  xml: string;
  rels: NewRel[];
  /** lower-cased workbook-level defined names the sheet's formulas / rules use */
  usedNames: Set<string>;
  failures: FormulaFailures[];
  /** formulas turned into values that had no stored result (the cell is empty) */
  noCache: number;
  /** left out: see ObjectItem */
  left: ObjectItem[];
  /** cells whose special data type (vm) became an ordinary value */
  vmCells: number;
  /** cell formats that did not fit into the 64,000 formats of a workbook (shown with the default format) */
  xfOverflow: number;
  dropped: { cf: number; dv: number; sparklines: number; hyperlinks: number; other: string[] };
  /** extra sentences for the warning (copySheetObjects can add some) */
  notes: string[];
  /** the sheet had hidden / visible state as in the source */
  formulaCount: number;
}

/** what copySheetObjects (patch/addsheets-objects.ts, task A2) is given for ONE copied sheet, after the core copy */
export interface SheetObjectsContext {
  source: {
    ks: KeySource;
    eng: Engine;
    info: SheetInfo;
    /** the source sheet part (as the package patcher reads it: attributes in double quotes, default namespace) */
    xml: string;
    /** every relationship of the source sheet (internal ones with their resolved part path) */
    rels: Array<Rel & { path: string }>;
    styles: Styles;
    sst: Sst;
  };
  target: {
    eng: Engine;
    /** the name of the copy */
    name: string;
    /** position in <sheets> = localSheetId */
    index: number;
    /** part path of the new sheet ("xl/worksheets/sheet5.xml") and its relationships part */
    path: string;
    relsPath: string;
    styles: Styles;
    sst: Sst;
  };
  /** the new sheet XML (a complete worksheet part without any of the object elements); assign a new text to change it */
  xml: string;
  /** the relationships of the new sheet; objects append theirs (use addRel for ids) */
  rels: NewRel[];
  /** a fresh relationship id of the new sheet */
  newRelId(): string;
  /** insert `<tag …/>` (a drawing, legacyDrawing, tableParts … element) where the worksheet schema wants it; declares xmlns:r when needed */
  insertElement(tag: string, fragment: string): void;
  /** the objects the source sheet has (kind / count), what the core copy left out */
  present: ObjectItem[];
  /** mark kinds as copied completely: they are not reported as left out */
  handled: Set<ObjectKind>;
  /** an additional sentence for the sheet's 'sheetCopyLimits' warning (partial copies, things skipped inside a kind) */
  note(text: string): void;
  /** sheet name maps */
  names: {
    /** lower-cased source sheet name -> name of the copy */
    rename: ReadonlyMap<string, string>;
    /** the formula copier of the source (rename sheet references; structured references need `tables`) */
    copier: FormulaCopier;
    rewrite(formula: string): Rewrite;
    nameView: NameView;
  };
  maps: {
    xf(sourceXf: number): number;
    dxf(sourceDxf: number): number | null;
    sst(sourceIndex: number): number;
    themeRgb: string[] | null;
    indexedRgb: string[] | null;
  };
}
