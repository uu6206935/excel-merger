// Types for run-job.mjs (used by the TypeScript tests).
export declare const POWERSHELL: string;

export interface ExcelDetection {
  installed: boolean;
  version?: string;
  reason?: string;
  progId?: string | null;
  exePath?: string | null;
  exeExists?: boolean;
}
export declare function detectExcel(): Promise<ExcelDetection>;
export declare function decode(buf: Uint8Array | null | undefined): string;
export declare function isExcelPid(pid: number): Promise<boolean>;
export declare function pidAlive(pid: number): boolean;
export declare function processStartMs(pid: number): Promise<number | null>;
/** Kills the PID only when it is an EXCEL.EXE (and, when `startedMs` is given, one started at that time). */
export declare function killExcelPid(pid: number, startedMs?: number): Promise<boolean>;
/** PowerShell single-quoted string literal (also doubles the typographic quotes U+2018 / U+2019 / U+201A / U+201B). */
export declare function psQuote(s: string): string;

export interface RunnerArg {
  name: string;
  value?: string;
  switch?: boolean;
}
export declare function runnerCommand(runner: string, runnerArgs: RunnerArg[], invoke?: string): { cmd: string; args: string[] };
/** process.env + overrides; on Windows an override replaces every case variant of the name (null removes it). */
export declare function childEnv(extra?: Record<string, string | null | undefined>, base?: NodeJS.ProcessEnv | Record<string, string | undefined>): Record<string, string>;

/** Every JSON line of the runner ({type:'start'|'excel'|'progress'|'warn'|'result'|'error', ...}). */
export type RunnerEvent = { type: string; [key: string]: unknown };

/** What is known once powershell.exe and our Excel are gone (the `closed` promise of an earlyResult run). */
export interface RunnerClosed {
  exitCode: number | null;
  excelKilled: boolean;
  stderr?: string;
  error?: string;
  wallMs: number;
}

export interface RunnerResult {
  type: 'result' | 'error';
  ok: boolean;
  code?: string;
  message?: string;
  output?: string;
  warnings?: Array<{ code: string; message: string }>;
  timings?: Array<{ i: number; op: string; ms: number }>;
  refErrors?: { cells: number; names: number; samples: string[] };
  clipboardTouched?: boolean;
  /** null when the result was delivered early (earlyResult) and powershell.exe has not ended yet */
  exitCode: number | null;
  excelPid: number;
  /** process start time (Unix ms) of our Excel as reported by the runner (used to avoid killing a recycled PID) */
  excelStartedMs?: number;
  excelKilled: boolean;
  started: boolean;
  cancelled: boolean;
  stderr?: string;
  wallMs: number;
  /** earlyResult runs only */
  earlyResult?: boolean;
  closed?: Promise<RunnerClosed>;
  reap?: () => Promise<void>;
  [key: string]: unknown;
}

export interface RunnerOptions {
  runner: string;
  jobFile?: string;
  probe?: boolean;
  timeoutMs?: number;
  idleMs?: number;
  onEvent?: (ev: RunnerEvent) => void;
  signal?: AbortSignal;
  invoke?: 'auto' | 'scriptblock' | 'file' | string;
  env?: Record<string, string>;
  /** resolve as soon as the final result line arrived; the process shutdown continues in the background (see run-job.mjs) */
  earlyResult?: boolean;
  /** earlyResult: ms the process may keep running after the final line before our Excel PID is killed (default 6000) */
  exitGraceMs?: number;
  extraArgs?: RunnerArg[];
}
export declare function runRunner(options: RunnerOptions): Promise<RunnerResult>;
