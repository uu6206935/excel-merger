/**
 * Round 5 helpers (owner: T): hostile inputs, settled-state waits, batch verification, reading exports without Excel.
 * Built on e2e/helpers.ts (the data-testid contract of docs/ARCHITECTURE.md §5).
 */
import fs from 'node:fs';
import path from 'node:path';
import { expect, test, type Page } from '@playwright/test';
import * as XLSX from 'xlsx';
import { excelAvailable, inspectFiles } from '../tools/lib/inspect.mjs';
import { checkOutput as checkOutputDumps } from '../tools/lib/verdict.mjs';
import {
  assertValid,
  compareOutputs,
  noteInspectionRetries,
  requireTestIds,
  SAMPLE_FILES,
  SAMPLES_DIR,
  type CheckOptions,
  type CheckOutputResult,
} from './helpers';

/**
 * Waits until the app has settled: no busy overlay, no file card still loading, no 「結合キーを探しています」 spinner.
 * (State, not time: every spec that changes files / keys / the base waits with this or with waitMergeReady.)
 */
export async function waitIdle(page: Page, timeout = 90_000): Promise<void> {
  await expect(page.getByTestId('busy-overlay')).toHaveCount(0, { timeout });
  await expect(page.locator('.card-loading')).toHaveCount(0, { timeout });
  await expect(page.getByText('結合キーを探しています')).toHaveCount(0, { timeout });
}

/** A file given to the hidden file input without a file on disk. */
export interface Payload {
  name: string;
  mimeType: string;
  buffer: Buffer;
}
export const payload = (name: string, buffer: Buffer): Payload => ({ name, mimeType: 'application/octet-stream', buffer });

/** `setInputFiles` with in-memory files or paths; with `expectCards`, waits for that many file cards and for them to finish loading. */
export async function dropFiles(page: Page, files: Array<string | Payload>, expectCards?: number): Promise<void> {
  await requireTestIds(page, ['file-input'], 'ファイル読み込み');
  // Playwright does not take paths and buffers in one call: a mixed drop (a real folder with one broken file) reads the paths
  const mixed = files.some((f) => typeof f !== 'string') && files.some((f) => typeof f === 'string');
  const list = mixed ? files.map((f) => (typeof f === 'string' ? payload(path.basename(f), fs.readFileSync(f)) : f)) : files;
  await page.getByTestId('file-input').setInputFiles(list as string[] | Payload[]);
  if (expectCards != null) {
    await expect(page.locator('[data-testid^="file-card-"]')).toHaveCount(expectCards);
    await expect(page.locator('.card-loading')).toHaveCount(0, { timeout: 60_000 });
  }
}

/** The first `frac` of a real workbook: a zip that stops in the middle (an interrupted download / copy). */
export function truncatedXlsx(name = SAMPLE_FILES[0], frac = 0.6): Buffer {
  const good = fs.readFileSync(path.join(SAMPLES_DIR, name));
  return good.subarray(0, Math.floor(good.length * frac));
}

/**
 * A structurally valid OLE2 (CFB v3) container with the two streams Excel writes for 「パスワードを使用して暗号化」
 * (EncryptionInfo, EncryptedPackage) — what the app sees when a password-protected .xlsx is dropped. The streams hold no real
 * ciphertext (nothing here is decrypted); both are 4096 bytes so that no mini stream is needed.
 */
export function passwordProtectedXlsx(): Buffer {
  const SECTOR = 512;
  const ENDOFCHAIN = 0xfffffffe;
  const FREESECT = 0xffffffff;
  const STREAM = 4096; // 8 sectors each
  const buf = Buffer.alloc(SECTOR * (1 + 1 + 1 + 16), 0); // header, FAT, directory, two streams
  // header
  Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]).copy(buf, 0);
  buf.writeUInt16LE(0x003e, 24); // minor version
  buf.writeUInt16LE(3, 26); // major version 3 (512-byte sectors)
  buf.writeUInt16LE(0xfffe, 28); // little endian
  buf.writeUInt16LE(9, 30); // sector shift
  buf.writeUInt16LE(6, 32); // mini sector shift
  buf.writeUInt32LE(1, 44); // FAT sectors
  buf.writeUInt32LE(1, 48); // first directory sector
  buf.writeUInt32LE(4096, 56); // mini stream cutoff
  buf.writeUInt32LE(ENDOFCHAIN, 60); // first mini FAT sector
  buf.writeUInt32LE(ENDOFCHAIN, 68); // first DIFAT sector
  buf.writeUInt32LE(0, 76); // DIFAT[0] = sector 0 (the FAT)
  for (let i = 1; i < 109; i++) buf.writeUInt32LE(FREESECT, 76 + i * 4);
  // FAT (sector 0, at byte 512): itself, the directory (1), EncryptionInfo 2..9, EncryptedPackage 10..17
  const fat = SECTOR;
  for (let i = 0; i < 128; i++) buf.writeUInt32LE(FREESECT, fat + i * 4);
  buf.writeUInt32LE(0xfffffffd, fat); // FATSECT
  buf.writeUInt32LE(ENDOFCHAIN, fat + 4);
  for (let s = 2; s <= 9; s++) buf.writeUInt32LE(s === 9 ? ENDOFCHAIN : s + 1, fat + s * 4);
  for (let s = 10; s <= 17; s++) buf.writeUInt32LE(s === 17 ? ENDOFCHAIN : s + 1, fat + s * 4);
  // directory (sector 1, at byte 1024): Root Entry, EncryptionInfo, EncryptedPackage
  const dir = SECTOR * 2;
  const entry = (i: number, name: string, type: number, left: number, right: number, child: number, start: number, size: number) => {
    const o = dir + i * 128;
    Buffer.from(name, 'utf16le').copy(buf, o);
    buf.writeUInt16LE(name ? (name.length + 1) * 2 : 0, o + 64);
    buf.writeUInt8(type, o + 66);
    buf.writeUInt8(1, o + 67); // black
    buf.writeUInt32LE(left, o + 68);
    buf.writeUInt32LE(right, o + 72);
    buf.writeUInt32LE(child, o + 76);
    buf.writeUInt32LE(start, o + 116);
    buf.writeUInt32LE(size, o + 120);
  };
  entry(0, 'Root Entry', 5, FREESECT, FREESECT, 1, ENDOFCHAIN, 0);
  entry(1, 'EncryptionInfo', 2, FREESECT, 2, FREESECT, 2, STREAM);
  entry(2, 'EncryptedPackage', 2, FREESECT, FREESECT, FREESECT, 10, STREAM);
  entry(3, '', 0, FREESECT, FREESECT, FREESECT, 0, 0);
  // EncryptionInfo stream (sector 2, at byte 1536): version 4.4 + flags
  buf.writeUInt16LE(4, SECTOR * 3);
  buf.writeUInt16LE(4, SECTOR * 3 + 2);
  buf.writeUInt32LE(0x40, SECTOR * 3 + 4);
  return buf;
}

/** One output to verify against its own base. */
export interface VerifyCase {
  label: string;
  base: string;
  file: string;
  check?: CheckOptions;
}

/**
 * check-output verdicts of several outputs against their own bases with ONE Excel instance (each base is inspected once).
 * Every verdict is asserted with assertValid. Returns [] without Excel (or with QA_NO_EXCEL=1).
 */
export async function verifyCases(cases: VerifyCase[]): Promise<CheckOutputResult[]> {
  if (!excelAvailable()) {
    test.info().annotations.push({ type: 'check-output', description: 'Excel がないため check-output は実行していません' });
    return [];
  }
  const bases = [...new Set(cases.map((c) => c.base))];
  const files = [...bases, ...cases.map((c) => c.file)];
  const dumps = await inspectFiles(files, { recalc: true });
  noteInspectionRetries(dumps, files);
  dumps.forEach((d, i) => {
    if (d.error && d.openMode !== 'failed') throw new Error(`${path.basename(files[i])} を調べられませんでした: ${d.error}`);
  });
  return cases.map((c, i) => {
    const b = dumps[bases.indexOf(c.base)];
    const d = dumps[bases.length + i];
    if (b.openMode === 'failed') throw new Error(`基準ファイルを Excel で開けませんでした: ${c.base}\n${b.error ?? ''}`);
    const verdict = checkOutputDumps(b, d, { sheet: c.check?.sheet, headerRow: c.check?.headerRow });
    const v: CheckOutputResult = { ...verdict, openMode: d.openMode } as CheckOutputResult;
    assertValid(v, c.label);
    return v;
  });
}

export interface SheetGrid {
  sheetNames: string[];
  sheet: string;
  /** every row of the sheet as displayed text ('' for empty cells), row 1 first */
  rows: string[][];
}

const toRows = (ws: XLSX.WorkSheet): string[][] =>
  (XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1, raw: false, defval: '', blankrows: true }) as unknown[][]).map((r) => r.map((c) => String(c ?? '')));

/** Reads an exported .xlsx without Excel (SheetJS): the displayed text of every cell of `sheet` (default: the first sheet). */
export function readSheet(file: string, sheet?: string): SheetGrid {
  const wb = XLSX.read(fs.readFileSync(file), { type: 'buffer' });
  const name = sheet ?? wb.SheetNames[0];
  const ws = wb.Sheets[name];
  if (!ws) throw new Error(`シート「${name}」がありません: ${wb.SheetNames.join(', ')}`);
  return { sheetNames: wb.SheetNames, sheet: name, rows: toRows(ws) };
}

/** Reads a sample source file (xlsx / xls / csv in Shift_JIS) the same way: rows of displayed text (row 1 first). */
export function readSample(name: string, sheet?: string): string[][] {
  const buf = fs.readFileSync(path.join(SAMPLES_DIR, name));
  // SheetJS (ESM build) has no code page tables: the Shift_JIS csv is decoded here and read as text
  const wb = /\.csv$/i.test(name) ? XLSX.read(new TextDecoder('shift_jis').decode(buf), { type: 'string', raw: true }) : XLSX.read(buf, { type: 'buffer' });
  return toRows(wb.Sheets[sheet ?? wb.SheetNames[0]]);
}

/** package-only compare (no Excel needed): identical when the tools know no real difference and no formula differs. */
export function samePackage(a: string, b: string): { identical: boolean; text: string } {
  const c = compareOutputs(a, b, ['--package-only']);
  return { identical: c.identical && c.formulaDiffs === 0, text: c.text };
}

/** Switches the fake Excel runner of the fallback server (playwright.config.ts → e2e/fixtures/fake-excel-runner.mjs). */
export function setFakeExcel(mode: { mode: 'fail' | 'crash' | 'slow' | 'noout'; code?: string; message?: string; ms?: number }): void {
  const file = process.env.E2E_FAKE_CONTROL;
  if (!file) throw new Error('E2E_FAKE_CONTROL が設定されていません（playwright.config.ts が設定します）');
  fs.mkdirSync(path.dirname(file), { recursive: true });
  fs.writeFileSync(file, JSON.stringify(mode), 'utf8');
}

// ---------------------------------------------------------------------------------------------------------------
// alignment of the merged rows against samples/truth.json
// ---------------------------------------------------------------------------------------------------------------

interface SourceSpec {
  /** header row of the data table in the file */
  headerRow: number;
  /** constant text columns that identify a row's content (formula columns have no cached value in the built-in engine's output) */
  probes: string[];
  /** file row -> roster row (samples/01_学生名簿.xlsx) of the same person, null = not in the roster */
  rows: Array<{ row: number; baseRow: number | null }>;
}

/**
 * Two files that can only be matched by the NAME (04 has kana only; 02 / 03 / 05 share no other column with each other) cannot tell
 * the two 佐藤 翔太 (roster rows 12 / 41) apart: neither answer is wrong for those two rows, so they are not judged.
 */
const blindToSameName = (a: string, b: string): boolean => {
  const nameOnly = new Set(['02_面接評価.xlsx', '03_連絡先.csv', '05_研修出欠.xlsx']);
  return a.startsWith('04_') || b.startsWith('04_') || (nameOnly.has(a) && nameOnly.has(b));
};

const TRUTH = JSON.parse(fs.readFileSync(path.join(SAMPLES_DIR, 'truth.json'), 'utf8')) as {
  base: { students: Array<{ row: number }>; sameNamePair: number[] };
  interview: { rows: SourceSpec['rows'] };
  contacts: { rows: SourceSpec['rows'] };
  aptitude: { rows: SourceSpec['rows'] };
  attendance: { rows: SourceSpec['rows'] };
};

export const SOURCE: Record<string, SourceSpec> = {
  '01_学生名簿.xlsx': { headerRow: 3, probes: ['フリガナ'], rows: TRUTH.base.students.map((s) => ({ row: s.row, baseRow: s.row })) },
  '02_面接評価.xlsx': { headerRow: 1, probes: ['面接官', 'コメント'], rows: TRUTH.interview.rows },
  '03_連絡先.csv': { headerRow: 1, probes: ['メール'], rows: TRUTH.contacts.rows },
  '04_適性検査.xls': { headerRow: 1, probes: ['性格タイプ'], rows: TRUTH.aptitude.rows },
  '05_研修出欠.xlsx': { headerRow: 2, probes: ['第1回(9/10)', '第2回(9/17)', '第3回(9/24)', '備考'], rows: TRUTH.attendance.rows },
};

const stemOf = (f: string) => f.replace(/\.[^.]+$/, '');
const squash = (s: string) => s.replace(/\s+/g, '');

/**
 * Checks every row of an exported merge against the ground truth: the row of the base file that belongs to roster student S must
 * show, in the columns imported from file F, what F holds for S (one of F's rows when F has S twice), and nothing when F has no S.
 * `others` = the joined files (not the excluded ones). Rows of the base that are not in the roster are not judged (truth.json
 * cannot say whether two files' "extra" people are the same person). Returns the problems and how many cells were compared.
 */
export function alignmentProblems(out: SheetGrid, baseFile: string, others: string[]): { problems: string[]; checked: number } {
  const problems: string[] = [];
  let checked = 0;
  const b = SOURCE[baseFile];
  const seen = new Set<number>();
  const baseTwice = new Set<number>();
  for (const r of b.rows) {
    if (r.baseRow == null) continue;
    if (seen.has(r.baseRow)) baseTwice.add(r.baseRow);
    seen.add(r.baseRow);
  }
  const titles = out.rows[b.headerRow - 1] ?? [];
  const colOf = (file: string, header: string): number => {
    const suffixed = `${header}（${stemOf(file)}）`;
    let i = file === baseFile ? -1 : titles.indexOf(suffixed);
    if (i < 0) i = titles.indexOf(header);
    return i;
  };
  for (const file of others) {
    const spec = SOURCE[file];
    const src = readSample(file);
    const srcHeader = src[spec.headerRow - 1];
    const srcCols = spec.probes.map((p) => srcHeader.indexOf(p));
    if (srcCols.some((c) => c < 0)) throw new Error(`${file}: 見出し ${spec.probes.join(' / ')} がありません: ${srcHeader.join(' | ')}`);
    const outCols = spec.probes.map((p) => colOf(file, p));
    if (outCols.some((c) => c < 0)) {
      problems.push(`${file}: 出力の見出しに ${spec.probes.join(' / ')} がありません: ${titles.join(' | ')}`);
      continue;
    }
    // roster row -> what this file holds for that person
    const held = new Map<number, string[]>();
    for (const r of spec.rows) {
      if (r.baseRow == null) continue;
      const v = srcCols.map((c) => squash(src[r.row - 1]?.[c] ?? '')).join('|');
      held.set(r.baseRow, [...(held.get(r.baseRow) ?? []), v]);
    }
    for (const r of b.rows) {
      if (r.baseRow == null) continue;
      if (blindToSameName(baseFile, file) && TRUTH.base.sameNamePair.includes(r.baseRow)) continue;
      const cells = outCols.map((c) => squash(out.rows[r.row - 1]?.[c] ?? '')).join('|');
      const want = held.get(r.baseRow) ?? [];
      checked++;
      // a person who is in the base TWICE (02's resubmissions) is joined by ID, but never by a name alone (keys/multi.ts: several
      // base rows sharing a non-ID key are "never decided by that tier"): blank is an acceptable answer there, a wrong value is not
      if (want.length && cells.replace(/\|/g, '') === '' && baseTwice.has(r.baseRow)) continue;
      if (want.length ? !want.includes(cells) : cells.replace(/\|/g, '') !== '') {
        problems.push(`${baseFile} の ${r.row} 行目（名簿 ${r.baseRow} 行目）: ${file} の列が「${cells}」、正しくは ${want.length ? want.map((w) => `「${w}」`).join(' か ') : '空'}`);
      }
    }
  }
  return { problems, checked };
}
