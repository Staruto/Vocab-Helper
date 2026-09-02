declare const process: {
  cwd(): string;
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
}

declare module "node:sqlite" {
  export class StatementSync {
    run(...params: unknown[]): { changes: number; lastInsertRowid: bigint };
    get(...params: unknown[]): any;
    all(...params: unknown[]): any[];
  }

  export class DatabaseSync {
    constructor(path: string);
    prepare(sql: string): StatementSync;
    exec(sql: string): void;
    close(): void;
  }
}
