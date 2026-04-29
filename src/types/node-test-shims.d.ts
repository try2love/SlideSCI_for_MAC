declare module "node:child_process" {
  export const execFile: any;
}

declare module "node:fs" {
  export const existsSync: any;
}

declare module "node:fs/promises" {
  export const mkdtemp: any;
  export const readFile: any;
  export const rm: any;
  export const writeFile: any;
}

declare module "node:os" {
  export const tmpdir: any;
}

declare module "node:path" {
  export const join: any;
  export const resolve: any;
}

declare module "node:util" {
  export const promisify: any;
}

declare const process: any;
