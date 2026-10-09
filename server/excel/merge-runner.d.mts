// Types for merge-runner.mjs (used by the TypeScript tests).
export declare const PLACEHOLDER: string;
export declare function mergeRunner(srcText: string, csText: string): string;
export declare function buildRunnerFile(dir: string, outFile?: string): { file: string; bytes: number };
export declare function runnerIsStale(dir: string): boolean;
