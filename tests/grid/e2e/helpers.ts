/** Shared helpers for the grid browser tests (demo page: src/grid-demo.tsx). */
import { expect, type Page } from '@playwright/test';
import { fileURLToPath } from 'node:url';
import fs from 'node:fs';

export const SHOTS = fileURLToPath(new URL('../shots/', import.meta.url));
fs.mkdirSync(SHOTS, { recursive: true });

export interface Call {
  name: string;
  args: unknown[];
}

/** Open the demo and collect page errors / React warnings. */
export async function openDemo(page: Page, query = ''): Promise<string[]> {
  const errors: string[] = [];
  page.on('pageerror', (e) => errors.push(String(e)));
  page.on('console', (m) => {
    if (m.type() === 'error' || m.type() === 'warning') errors.push(`${m.type()}: ${m.text()}`);
  });
  await page.goto('/grid-demo.html' + query);
  await page.waitForSelector('[data-testid="grid-cell-2-0"]');
  await frames(page, 2);
  return errors;
}

export async function frames(page: Page, n = 2): Promise<void> {
  await page.evaluate(async (k) => {
    for (let i = 0; i < k; i++) await new Promise((r) => requestAnimationFrame(() => r(null)));
  }, n);
}

export const titles = (page: Page) => page.evaluate(() => window.__gridDemo!.titles());
export const order = (page: Page) => page.evaluate(() => window.__gridDemo!.order());
export const idOf = async (page: Page, title: string) => {
  const id = await page.evaluate((t) => window.__gridDemo!.idOf(t), title);
  if (!id) throw new Error('no column ' + title);
  return id;
};
export const indexOfTitle = async (page: Page, title: string) => {
  const i = (await titles(page)).indexOf(title);
  if (i < 0) throw new Error('no column ' + title);
  return i;
};
export const widthOf = (page: Page, id: string) => page.evaluate((c) => window.__gridDemo!.width(c), id);
export const history = (page: Page) => page.evaluate(() => window.__gridDemo!.history());
export const cellValue = (page: Page, row: number, id: string) => page.evaluate(({ r, c }) => window.__gridDemo!.cellText(r, c), { r: row, c: id });

export async function calls(page: Page, name?: string): Promise<Call[]> {
  const all = await page.evaluate(() => JSON.parse(JSON.stringify(window.__gridDemo!.calls)) as Call[]);
  return name ? all.filter((c) => c.name === name) : all;
}
export async function clearCalls(page: Page): Promise<void> {
  await page.evaluate(() => {
    window.__gridDemo!.calls.length = 0;
  });
}

/** header cell locator by exact title */
export const header = (page: Page, title: string) =>
  page.locator('.g-header .g-hcell', { has: page.locator('.g-htitle', { hasText: new RegExp(`^${title.replace(/[()（）]/g, '\\$&')}$`) }) });

/** scroll the column into view through the grid handle and return its header box */
export async function reveal(page: Page, title: string) {
  const id = await idOf(page, title);
  await page.evaluate((c) => window.__gridDemo!.grid()!.scrollToColumn(c), id);
  await frames(page, 3);
  const box = await header(page, title).boundingBox();
  if (!box) throw new Error('header not rendered: ' + title);
  return box;
}

/** scroll a data row into view (no selection) */
export async function revealRow(page: Page, row: number) {
  await page.evaluate((r) => window.__gridDemo!.grid()!.scrollToRow(r), row);
  await frames(page, 3);
}

/** the cell locator of an Excel row and a column title (current position) */
export async function cellOf(page: Page, row: number, title: string) {
  await reveal(page, title);
  await revealRow(page, row);
  const i = await indexOfTitle(page, title);
  return page.getByTestId(`grid-cell-${row}-${i}`);
}

export const selectedTitles = (page: Page) => page.$$eval('.g-hcell.is-sel .g-htitle', (els) => els.map((e) => e.textContent));
export const address = (page: Page) => page.getByTestId('formula-bar-address').textContent();
export const fbValue = (page: Page) => page.getByTestId('formula-bar-input').inputValue();

export async function expectNoErrors(errors: string[]) {
  expect(errors.filter((e) => !/Download the React DevTools/.test(e))).toEqual([]);
}
