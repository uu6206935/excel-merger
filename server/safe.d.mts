// Types for safe.mjs (used by the TypeScript tests).
export declare const OUTPUT_EXTS: Set<string>;
export declare function safeDecode(s: string): string;
export declare function isReservedDeviceName(stem: string): boolean;
export declare function cleanOutputName(raw: unknown): { name: string; error?: undefined } | { error: string; name?: undefined };
export declare function isInside(dir: string, file: string): boolean;
export declare function textCheckOpenPath(raw: unknown): { status: number; error: string } | null;
export interface OpenCommand {
  cmd: string;
  args: string[];
  env: Record<string, string>;
  verbatim: boolean;
}
export declare function buildOpenCommand(target: string, select: boolean, sysRoot?: string): OpenCommand;
export declare function validateRunnerOps(ops: unknown[], options?: { allowDebug?: boolean }): { ops: Array<Record<string, unknown>>; error?: undefined } | { error: string; ops?: undefined };
