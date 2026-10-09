/**
 * Column autofit (double-click on a resize handle): measures the displayed texts with a canvas, like Excel's
 * "列の幅の自動調整" (all rows, the header row included). DOM-only helper.
 */

let ctx: CanvasRenderingContext2D | null | undefined;
let lastFont = '';

function context(): CanvasRenderingContext2D | null {
  if (ctx === undefined) {
    try {
      ctx = document.createElement('canvas').getContext('2d');
    } catch {
      ctx = null;
    }
  }
  return ctx;
}

/** Width of `text` in px for a CSS font shorthand; a rough estimate when no canvas is available. */
export function measureText(text: string, font: string): number {
  if (!text) return 0;
  const c = context();
  if (!c) {
    let w = 0;
    for (const ch of text) w += ch.charCodeAt(0) > 0xff ? 15 : 8;
    return w;
  }
  if (font !== lastFont) {
    c.font = font;
    lastFont = font;
  }
  return c.measureText(text).width;
}

export interface FitText {
  text: string;
  font: string;
}

/** Max number of rows measured; beyond that an even sample is taken (Excel measures all, but 1M rows would freeze the UI). */
const MAX_MEASURED = 30000;

/**
 * Autofit width in px: widest line of all texts + cell padding, clamped to [min, max].
 * `texts(i)` returns the displayed text of row index i (or null for empty cells).
 */
export function autoFitWidth(
  count: number,
  texts: (i: number) => FitText | null,
  header: { text: string; font: string; extra: number },
  limits: { min: number; max: number; padding: number },
): number {
  let w = header.text ? measureText(header.text, header.font) + header.extra : 0;
  const step = count > MAX_MEASURED ? count / MAX_MEASURED : 1;
  for (let k = 0; k < Math.min(count, MAX_MEASURED); k++) {
    const i = step === 1 ? k : Math.floor(k * step);
    const t = texts(i);
    if (!t || !t.text) continue;
    if (t.text.includes('\n')) {
      for (const line of t.text.split('\n')) w = Math.max(w, measureText(line, t.font) + limits.padding);
    } else {
      w = Math.max(w, measureText(t.text, t.font) + limits.padding);
    }
  }
  return Math.max(limits.min, Math.min(limits.max, Math.ceil(w)));
}
