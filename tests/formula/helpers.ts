// Shared helpers for the formula-engine tests (fixture loading).
import fs from 'node:fs';
import path from 'node:path';
import zlib from 'node:zlib';
import { fileURLToPath } from 'node:url';

const here = path.dirname(fileURLToPath(import.meta.url));
export const FIXTURES = path.resolve(here, '../fixtures/formula');

/** Read a JSON fixture from tests/fixtures/formula (BOM tolerant; *.json.gz is gunzipped). */
export function readFixture<T = unknown>(name: string): T {
  const p = path.join(FIXTURES, name);
  let buf = fs.readFileSync(p);
  if (name.endsWith('.gz')) buf = zlib.gunzipSync(buf);
  return JSON.parse(buf.toString('utf8').replace(/^﻿/, '')) as T;
}

/** Collects failures so that one `expect(fails).toEqual([])` shows every mismatch at once. */
export class Tally {
  pass = 0;
  fails: string[] = [];
  ok(cond: boolean, msg: string | (() => string)): void {
    if (cond) this.pass++;
    else this.fails.push(typeof msg === 'function' ? msg() : msg);
  }
}
