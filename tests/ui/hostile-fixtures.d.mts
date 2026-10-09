export function nameOf(i: number): string;
export function idOf(i: number): string;
/** Writes the hostile / messy input files (see hostile-fixtures.mjs); returns name -> absolute path. */
export function makeHostileFixtures(dir: string, opts?: { big?: boolean; huge?: boolean }): Promise<Record<string, string>>;
