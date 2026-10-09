// Types for tools/lib/ps.mjs (owner: T).

export const POWERSHELL: string;
export function ensureBom(files: string[]): void;
export function runPs(
  script: string,
  args: string[],
  opts?: { timeoutMs?: number; onLine?: (obj: Record<string, unknown> & { type?: string }, raw: string) => void; cwd?: string },
): Promise<{ code: number; lines: Array<Record<string, unknown> & { type?: string }>; stderr: string; timedOut: boolean; excelPids: number[] }>;
