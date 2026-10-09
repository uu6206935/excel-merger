// Dev helper: print the border sides of probe cells (see tests/golden/com/probe-insert-format.ps1).
import fs from 'node:fs';
import { snapshot } from '../snapshot';

function sides(sig: string | undefined): string {
  if (!sig) return '-';
  const b = /border=(<border>.*?<\/border>)/.exec(sig)?.[1] ?? '';
  const out: string[] = [];
  for (const s of ['left', 'right', 'top', 'bottom']) {
    const m = new RegExp(`<${s} style="([A-Za-z]+)"`).exec(b);
    if (m) out.push(s[0].toUpperCase() + (m[1] === 'thin' ? '' : ':' + m[1]));
  }
  const fill = /patternType="solid"/.test(sig) ? ' fill' : '';
  return (out.join('') || 'none') + fill;
}

async function main() {
  const [before, after] = process.argv.slice(2);
  const a = await snapshot(new Uint8Array(fs.readFileSync(before)));
  const b = await snapshot(new Uint8Array(fs.readFileSync(after)));
  const ca = a.sheets.find((s) => s.name === 'Cols')!, cb = b.sheets.find((s) => s.name === 'Cols')!;
  console.log('COLS: #  left(B)  right(C)  -> new C');
  for (let i = 0; i < 16; i++) {
    const r = 2 + 2 * i;
    console.log(`${i + 1}\t${sides(ca.styles['B' + r])}\t${sides(ca.styles['C' + r])}\t-> ${sides(cb.styles['C' + r])}\t(D=${sides(cb.styles['D' + r])})`);
  }
  const ra = a.sheets.find((s) => s.name === 'Rows')!, rb = b.sheets.find((s) => s.name === 'Rows')!;
  console.log('ROWS: #  up(2)  down(3)  -> new 3   rows:', JSON.stringify(rb.rows));
  const col = (n: number) => String.fromCharCode(64 + n);
  for (let i = 0; i < 16; i++) {
    const c = col(2 + 2 * i);
    console.log(`${i + 1}\t${sides(ra.styles[c + 2])}\t${sides(ra.styles[c + 3])}\t-> ${sides(rb.styles[c + 3])}\t(4=${sides(rb.styles[c + 4])})`);
  }
}
main();
