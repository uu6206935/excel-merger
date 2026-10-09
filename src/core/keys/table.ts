/**
 * The plain Table the key detector works on + header-row detection and the 横持ち (wide layout) hint.
 * (R6 prototype's table.ts without the file reader: the SourceBook adapter lives in region.ts.)
 */
import { type Cell, cellText, nfkc, isIdShape, isKanjiNameShape } from './normalize';
import { headerNorm, bestHeaderType, isNonKeyHeader } from './detect';

export interface Table {
  /** file label (for messages) */
  file: string;
  sheet: string;
  /** 1-based Excel row of the header row */
  headerRow: number;
  headers: string[];
  /** 1-based Excel column of each table column */
  cols: number[];
  /** matchable data rows (blank and summary rows excluded); each array has headers.length cells */
  rows: Cell[][];
  /** text Excel displays for data cell [row][col] (number format applied: 000123) — computed lazily */
  display: (row: number, col: number) => string;
  /** Excel number format of each column (mode over data cells) */
  numFmt: string[];
  /** column number format is a date/time format */
  dateFmt: boolean[];
  /** 1-based Excel row number of each data row (for messages: "12行目") */
  excelRow: number[];
  layoutHint?: string;
}

export const MAX_HEADER_SCAN = 30;

const ANSWER_CELL_RE = /^\d{1,2}:\d{2}|^[○×△✕✓◯-]$|^(ng|ok|終日ng)$/i;
/** kanji column titles that look like 2–4 character names (受検日 / 非言語 / 合計 / 総合判定 / 出席回数) */
const HEADER_WORD_RE = /(日|日時|時間|期間|語|計|数|率|点|額|金|費|型|判定|区分|番号|結果|評価|状況|有無|備考|欄|順|順位|種別|分類|形式|方法|内容|理由|予定|場所|会場|時刻|回)$/u;

function isEmptyValue(v: Cell): boolean {
  return v === null || v === undefined || (typeof v === 'string' && v.trim() === '');
}

export interface HeaderRowScore {
  /** index into the scanned grid */
  r: number;
  /** cells that match a key synonym (score >= 0.7) or a known non-key header word */
  hits: number;
  /** density-style score (pass 2) */
  text: number;
  nonEmpty: number;
}

/** Scores every candidate header row of a grid (rows of cell values). */
export function scoreHeaderRows(grid: Cell[][]): HeaderRowScore[] {
  const limit = Math.min(MAX_HEADER_SCAN, grid.length);
  const scored: HeaderRowScore[] = [];
  for (let r = 0; r < limit; r++) {
    const cells = (grid[r] ?? []).filter((v) => !isEmptyValue(v));
    if (cells.length < 2) continue;
    let hasDataBelow = false;
    for (let k = r + 1; k < Math.min(grid.length, r + 6) && !hasDataBelow; k++) hasDataBelow = (grid[k] ?? []).some((v) => !isEmptyValue(v));
    if (!hasDataBelow) continue;
    let hits = 0;
    let text = 0;
    for (const v of cells) {
      if (typeof v !== 'string') {
        text -= 1;
        continue;
      }
      const t = nfkc(v);
      const hn = headerNorm(t);
      if (bestHeaderType(hn).score >= 0.7 || isNonKeyHeader(hn)) hits++;
      if (t.length > 40) text += 0.5; // long form-question headers
      else if (isIdShape(t) && t.length >= 5) text -= 0.5; // TEST00001 is data
      else if (ANSWER_CELL_RE.test(t)) text -= 0.5; // answer cells
      else text += 1;
    }
    scored.push({ r, hits, text, nonEmpty: cells.length });
  }
  return scored;
}

/**
 * Header row = the row (within the first 30) with the most dictionary hits (ties → earliest row);
 * when no row has any hit, the densest short-text row. Returns an index into grid, or -1.
 */
export function detectHeaderRow(grid: Cell[][]): number {
  const scored = scoreHeaderRows(grid);
  let best = -1;
  let bestHits = 0;
  for (const s of scored) {
    if (s.hits > bestHits) {
      bestHits = s.hits;
      best = s.r;
    }
  }
  if (bestHits >= 1) return best;
  let bestText = -Infinity;
  for (const s of scored) {
    if (s.text > bestText) {
      bestText = s.text;
      best = s.r;
    }
  }
  return bestText > 0 ? best : scored[0]?.r ?? -1;
}

/**
 * "Wide" (横持ち) layout: persons are laid out across COLUMNS (the header row, or the row just above it, is full of
 * IDs / names). Row-wise key detection is meaningless there; the UI says so.
 * @param h index of the header row in grid
 */
export function wideLayoutHint(grid: Cell[][], h: number, nCols: number, rowOffset = 0): string | undefined {
  for (const r of [h - 1, h]) {
    if (r < 0 || !grid[r]) continue;
    // a row with two or more key headers (学籍番号 / 氏名カナ …) is an ordinary header row: its other kanji words
    // (受検日 / 言語 / 合計 / 総合判定) are column titles, not people — only ID-shaped cells count there
    let keyHeaders = 0;
    for (const v of grid[r]) if (typeof v === 'string' && bestHeaderType(headerNorm(v)).score >= 0.7) keyHeaders++;
    const cells = grid[r].filter((v) => {
      if (typeof v !== 'string') return false;
      const hn = headerNorm(v);
      if (bestHeaderType(hn).score >= 0.5 || isNonKeyHeader(hn)) return false; // 氏名/大学名 are titles, not people
      if (isIdShape(v) && nfkc(v).length >= 5) return true;
      return keyHeaders < 2 && isKanjiNameShape(v) && !HEADER_WORD_RE.test(nfkc(v).trim());
    });
    if (cells.length >= 5 && cells.length >= nCols * 0.4) {
      return `${r + 1 + rowOffset}行目に ID/氏名 が横方向に${cells.length}個並んでいます（横持ちレイアウト）。行方向の結合キーは検出できません。`;
    }
  }
  return undefined;
}

/** Header text of a cell ('' when empty). */
export function headerText(v: Cell): string {
  return isEmptyValue(v) ? '' : nfkc(cellText(v)).replace(/\s+/g, ' ');
}
