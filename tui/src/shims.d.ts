declare const process: {
  cwd(): string;
  argv: string[];
  env: Record<string, string | undefined>;
  exit(code?: number): never;
  stdin?: {
    isTTY?: boolean;
    setRawMode?(mode: boolean): void;
  };
  stdout?: {
    isTTY?: boolean;
    columns?: number;
    rows?: number;
    write?(text: string): boolean;
  };
};

declare module "node:path" {
  export function resolve(...parts: string[]): string;
  export function dirname(path: string): string;
  export function basename(path: string): string;
  export function join(...parts: string[]): string;
}

declare module "node:fs" {
  export function existsSync(path: string): boolean;
  export function mkdtempSync(prefix: string): string;
  export function renameSync(oldPath: string, newPath: string): void;
  export function rmSync(path: string, options?: { recursive?: boolean; force?: boolean }): void;
  export function unlinkSync(path: string): void;
  export function writeFileSync(path: string, data: string, encoding?: string): void;
}

declare module "node:os" { export function tmpdir(): string; }
declare module "node:test" { const test: (name: string, fn: () => void | Promise<void>) => void; export default test; }
declare module "node:assert/strict" {
  const assert: {
    equal(actual: unknown, expected: unknown, message?: string): void;
    deepEqual(actual: unknown, expected: unknown, message?: string): void;
    ok(value: unknown, message?: string): void;
    throws(fn: () => unknown, expected?: RegExp | ((error: unknown) => boolean)): void;
  };
  export default assert;
}

declare module "node:crypto" {
  export function createHash(algorithm: string): { update(data: string): any; digest(encoding: string): string };
}

declare module "node:sqlite" {
  export class StatementSync {
    run(...params: unknown[]): { changes: number; lastInsertRowid: bigint };
    get(...params: unknown[]): any;
    all(...params: unknown[]): any[];
  }

  export class DatabaseSync {
    constructor(path: string, options?: { readOnly?: boolean });
    prepare(sql: string): StatementSync;
    exec(sql: string): void;
    close(): void;
  }
}
