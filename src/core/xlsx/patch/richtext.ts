/**
 * Rich text of a shared string (<si> body) moved from one package to another (premerge imported cells, the sheet copier).
 */
import { convertColors, type Styles } from './styles';
import { getAttr, setAttr } from './xml';

/** what convertRich needs to know about the SOURCE package */
export interface RichSource {
  /** the source's styles.xml (phonetic properties refer to a font of it) */
  styles: Styles;
  /** source theme colours when they differ from the base theme (theme refs are then converted to RGB) */
  themeRgb: string[] | null;
  /** the source's legacy indexed palette when it differs from the base's */
  indexedRgb: string[] | null;
}

/**
 * A source <si> body in the base package: run fonts are inline (<rPr>), theme colours / theme font schemes are
 * converted when the themes differ (like imported cell styles); <phoneticPr fontId> refers to a font of the source
 * styles.xml and is re-pointed at the same font imported into the base (dropped when it does not exist).
 */
export function convertRich(xml: string, c: RichSource, base: Styles): string {
  let x = xml;
  if (c.indexedRgb && !c.themeRgb) x = x.replace(/<rPr>([\s\S]*?)<\/rPr>/g, (_m, inner: string) => `<rPr>${convertColors(inner, null, c.indexedRgb)}</rPr>`);
  if (c.themeRgb) {
    const rgb = c.themeRgb;
    x = x.replace(/<rPr>([\s\S]*?)<\/rPr>/g, (_m, inner: string) => `<rPr>${convertColors(inner, rgb, c.indexedRgb)}</rPr>`);
  }
  x = x.replace(/<phoneticPr\b[^>]*\/>/g, (el) => {
    const id = getAttr(el, 'fontId');
    if (id === undefined) return el;
    const nid = base.importFont(c.styles, Number(id), c.themeRgb, c.indexedRgb);
    return nid === null ? '' : setAttr(el, 'fontId', String(nid));
  });
  return x;
}
