import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, resolve } from 'node:path';

const initializedWasm = new WeakMap<object, unknown>();

/** Load and synchronously initialize a wasm-pack `--target web` artifact in
 *  Node. The `--target web` build expects to be `fetch`ed from a URL; in Node
 *  we sidestep that path by reading the .wasm bytes off disk and feeding them
 *  into the generated `initSync` helper. */
export function loadWasmModule<T>(jsModule: T & { initSync: (init: { module: WebAssembly.Module }) => unknown }, wasmPath: string): T {
  const bytes = readFileSync(wasmPath);
  const module = new WebAssembly.Module(bytes);
  initializedWasm.set(jsModule as object, jsModule.initSync({ module }));
  return jsModule;
}

/** Diagnostic linear-memory size for fresh-process benchmarks. */
export function wasmMemoryPages(jsModule: object): number | undefined {
  const initialized = initializedWasm.get(jsModule) as { memory?: WebAssembly.Memory } | undefined;
  return initialized?.memory ? initialized.memory.buffer.byteLength / (64 * 1024) : undefined;
}

/** Resolve a path relative to a workspace-package source file. Used by the
 *  per-format entry points to locate the `.wasm` artifact emitted by
 *  `wasm-pack build --out-dir ../src/wasm`. */
export function resolveWasm(metaUrl: string, relPath: string): string {
  const here = dirname(fileURLToPath(metaUrl));
  return resolve(here, relPath);
}
