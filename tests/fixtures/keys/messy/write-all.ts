/**
 * Regenerates the committed xlsx files of the messy fixtures:  npx tsx tests/fixtures/keys/messy/write-all.ts
 * Output: ./out/<case name>/<file name>.xlsx (+ ./out/manifest.json with the ground truth of every file).
 * Deterministic: the same sources always produce the same bytes (tests/keys/messy.test.ts checks that the committed
 * files are in sync with these scripts).
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { allCases } from './all';
import { toXlsxBytes } from './lib';

const here = path.dirname(fileURLToPath(import.meta.url));
const out = path.join(here, 'out');

function main(): void {
  fs.rmSync(out, { recursive: true, force: true });
  const manifest: Record<string, unknown> = {};
  let files = 0;
  let bytes = 0;
  for (const c of allCases) {
    const dir = path.join(out, c.name);
    fs.mkdirSync(dir, { recursive: true });
    for (const f of c.files) {
      const data = toXlsxBytes(f);
      fs.writeFileSync(path.join(dir, f.fileName), data);
      files++;
      bytes += data.length;
    }
    manifest[c.name] = {
      note: c.note,
      plan: c.plan ?? null,
      expect: c.expect,
      files: c.files.map((f) => ({ id: f.id, fileName: f.fileName, dataSheet: f.dataSheet ?? f.sheets[0].name, truth: f.truth })),
    };
  }
  fs.writeFileSync(path.join(out, 'manifest.json'), JSON.stringify(manifest, null, 1));
  console.log(`${allCases.length} cases, ${files} files, ${(bytes / 1024).toFixed(0)} KB`);
}

main();
