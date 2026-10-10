/**
 * Damaged object parts (task A2): every part a drawing / chart / picture / note / table / pivot table owns is removed, truncated or emptied in the
 * Excel-authored fixtures of round 9 (and their relationships files removed), then the file is the other file of a merge. A damaged part of
 * an object is that object's problem, never the merge's: the output is a package that passes the linter (no dangling relationship, no orphan
 * part, every XML part well-formed) with the objects that could be copied — or, where the core copy cannot read the file at all, a
 * refusal with a Japanese message (as without objects). An error of the copier itself (TypeError ...) fails the test: the object phase
 * rethrows those under vitest (addsheets-objects.ts `guarded`) instead of turning them into a warning.
 */
import fs from 'node:fs';
import path from 'node:path';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { REPO, premergeCopies, type FileIn } from './addsheets-helpers';
import { lintObjects, lintObjectsText, objectsOf } from './addsheets-objects-helpers';

const FIX = ['ob-charts', 'ob-shapes', 'ob-comments', 'ob-table', 'ob-pivot', 'ob-slicer-control', 'ob-strict'];
const OBJECT_PART = /^xl\/(drawings|charts|tables|pivotTables|pivotCache|comments|threadedComments|persons|media|embeddings|slicers|slicerCaches|ctrlProps|activeX)\b|^xl\/comments\d*\.xml$/i;
const base = (): FileIn => ({ bytes: new Uint8Array(fs.readFileSync(path.join(REPO, 'samples/02_面接評価.xlsx'))), fileName: '02_面接評価.xlsx' });

async function mutate(bytes: Uint8Array, part: string, how: 'removed' | 'truncated' | 'emptied' | 'rels'): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(bytes);
  if (how === 'removed') zip.remove(part);
  else if (how === 'rels') {
    const i = part.lastIndexOf('/');
    zip.remove(`${part.slice(0, i + 1)}_rels/${part.slice(i + 1)}.rels`);
  } else {
    const text = await zip.file(part)!.async('string');
    zip.file(part, how === 'emptied' ? '' : text.slice(0, Math.floor(text.length / 2)));
  }
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

describe('a damaged part of an object never breaks the merge', () => {
  for (const name of FIX) {
    it(name, async () => {
      const bytes = new Uint8Array(fs.readFileSync(path.join(REPO, 'tests/fixtures/addsheets-objects', `${name}.xlsx`)));
      const parts = Object.keys((await JSZip.loadAsync(bytes)).files).filter((p) => !p.endsWith('/') && !/\.rels$/.test(p) && OBJECT_PART.test(p));
      expect(parts.length).toBeGreaterThanOrEqual(3);
      const refused: string[] = [];
      const problems: string[] = [];
      let runs = 0;
      for (const part of parts) {
        const textual = /\.(xml|vml)$/i.test(part);
        const hows: Array<'removed' | 'truncated' | 'emptied' | 'rels'> = textual ? ['removed', 'truncated', 'emptied', 'rels'] : ['removed'];
        for (const how of hows) {
          runs++;
          const damaged = await mutate(bytes, part, how);
          let res;
          try {
            res = await premergeCopies(base(), [{ bytes: damaged, fileName: `${name}.xlsx` }], { dropOtherSheets: true });
          } catch (e) {
            // refused: with a Japanese message (the reader / the core copy found the file unusable), nothing else
            const msg = (e as Error).message;
            if (!/[぀-ヿ一-鿿]/.test(msg)) problems.push(`${part} ${how}: ${(e as Error).name}: ${msg.slice(0, 120)}`);
            else refused.push(`${part} ${how}`);
            continue;
          }
          const lint = (await lintObjects(res.bytes)).filter((l) => !/^APP_/.test(l.code));
          if (lint.length) problems.push(`${part} ${how}: ${lintObjectsText(lint).replace(/\n/g, ' | ').slice(0, 200)}`);
          await objectsOf(res.bytes); // readable
        }
      }
      expect(problems).toEqual([]);
      // most damage is survivable: the file is refused only when the reader needs the part (tables / comments are read for the preview)
      expect(refused.length).toBeLessThan(runs * 0.5);
    }, 120_000);
  }
});
