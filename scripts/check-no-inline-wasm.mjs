#!/usr/bin/env node
// PD13 — guard against a WASM parser being re-inlined into a JS bundle as a
// base64 data-URL (which once bloated the package ~6 MB; fixed by shipping the
// .wasm as a real asset — commits b004e8c / 3ac69ca). This is a *preventive*
// gate: the current wasm-bindgen toolchain emits a real `new URL(...wasm)`
// reference in both release and dev profiles, so nothing base64-inlines today.
// The gate fails a build if that ever regresses.
//
// Usage: node scripts/check-no-inline-wasm.mjs [dist-dir ...]
//   Defaults to ./dist. Scans every .js/.mjs/.cjs file for:
//     1. the base64 encoding of the WASM magic bytes (\0asm) — "AGFzbQ" — which
//        is the unmistakable head of a base64-inlined module, and
//     2. any wasm data-URL literal ("data:application/wasm;base64,").
//   Also flags a JS chunk larger than MAX_JS_BYTES that also contains a long
//   base64 run, as a size backstop for encodings that dodge the markers above.
//
// Real .wasm asset files are expected and ignored — only JS text is scanned.

import { readdirSync, readFileSync, statSync, existsSync } from 'node:fs';
import { join, extname } from 'node:path';

const WASM_MAGIC_B64 = 'AGFzbQ'; // base64("\0asm")
const WASM_DATA_URL = 'data:application/wasm;base64,';
// A parser .wasm is ~0.6–0.8 MB; base64-inlined it would be ~1 MB of text in a
// single JS chunk. Real JS chunks in this repo top out well under that (the
// mathjax bundle ~3 MB is the sole large JS asset and contains no base64 wasm).
const MAX_JS_BYTES = 1_500_000;
const LONG_B64_RUN = /[A-Za-z0-9+/]{50000,}={0,2}/; // a >~37 KB contiguous base64 blob

const JS_EXT = new Set(['.js', '.mjs', '.cjs']);

/** @param {string} dir @param {string[]} out */
function collectJs(dir, out) {
  for (const entry of readdirSync(dir, { withFileTypes: true })) {
    const p = join(dir, entry.name);
    if (entry.isDirectory()) collectJs(p, out);
    else if (JS_EXT.has(extname(entry.name))) out.push(p);
  }
}

const dirs = process.argv.slice(2);
if (dirs.length === 0) dirs.push('dist');

/** @type {string[]} */
const problems = [];
let scanned = 0;

for (const dir of dirs) {
  if (!existsSync(dir)) {
    console.error(`check-no-inline-wasm: directory not found: ${dir}`);
    process.exit(2);
  }
  /** @type {string[]} */
  const files = [];
  collectJs(dir, files);
  for (const file of files) {
    scanned++;
    const text = readFileSync(file, 'utf8');
    const size = statSync(file).size;
    if (text.includes(WASM_DATA_URL)) {
      problems.push(`${file}: contains a "${WASM_DATA_URL}" data-URL`);
    } else if (text.includes(WASM_MAGIC_B64)) {
      problems.push(`${file}: contains base64 WASM magic "${WASM_MAGIC_B64}" (inlined module?)`);
    } else if (size > MAX_JS_BYTES && LONG_B64_RUN.test(text)) {
      problems.push(
        `${file}: ${(size / 1e6).toFixed(1)} MB JS with a long base64 run — possible inlined binary`,
      );
    }
  }
}

if (problems.length > 0) {
  console.error('check-no-inline-wasm: FAILED — base64-inlined WASM detected:');
  for (const p of problems) console.error(`  - ${p}`);
  console.error(
    '\nThe .wasm parsers must ship as real asset files referenced by URL, not ' +
      'base64 data-URLs. See scripts/check-no-inline-wasm.mjs and commits ' +
      'b004e8c / 3ac69ca.',
  );
  process.exit(1);
}

console.log(`check-no-inline-wasm: OK — ${scanned} JS file(s) scanned, no inlined WASM.`);
