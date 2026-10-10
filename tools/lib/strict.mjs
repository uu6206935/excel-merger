// Is a workbook file in the "Strict Open XML" format (owner: T)? The app cannot copy charts / pictures / shapes between a Strict and a normal workbook (the copied
// sheet is still complete otherwise, and the app says so in its warnings), so check-output --with-copies needs to know which of the files are Strict.
import fs from 'node:fs';
import JSZip from 'jszip';

/** true when `file` is a .xlsx / .xlsm whose workbook part uses the Strict namespace (http://purl.oclc.org/ooxml/...) */
export async function isStrictOoxml(file) {
  try {
    if (!/\.(xlsx|xlsm|xltx|xltm)$/i.test(String(file))) return false;
    const zip = await JSZip.loadAsync(fs.readFileSync(file));
    const part = zip.file('xl/workbook.xml');
    if (!part) return false;
    return /purl\.oclc\.org\/ooxml/.test((await part.async('string')).slice(0, 4000));
  } catch {
    return false;
  }
}
