#!/usr/bin/env node
import { spawnSync } from 'node:child_process';
import { mkdtempSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));

/** Run the iterator in a fresh Vitest/Node process and return its metrics. */
export function benchmarkXlsxRowIterator(inputPath, options = {}) {
  const vitest = resolve(here, '../../../node_modules/.bin/vitest');
  const childTest = resolve(here, '../src/xlsx-row-iterator-bench-child.test.ts');
  const metricsDirectory = mkdtempSync(resolve(tmpdir(), 'ooxml-xlsx-bench-'));
  const metricsPath = resolve(metricsDirectory, 'metrics.json');
  try {
    const child = spawnSync(vitest, ['run', childTest], {
      cwd: resolve(here, '../../..'),
      encoding: 'utf8',
      env: {
        ...process.env,
        OOXML_XLSX_BENCH_INPUT: resolve(inputPath),
        OOXML_XLSX_BENCH_OUTPUT: metricsPath,
        ...(options.maxArchiveEntryBytes === undefined
          ? {}
          : { OOXML_XLSX_BENCH_MAX_ENTRY: String(options.maxArchiveEntryBytes) }),
      },
    });
    if (child.status !== 0) {
      throw new Error(`benchmark child failed (${child.status})\n${child.stdout}\n${child.stderr}`);
    }
    return JSON.parse(readFileSync(metricsPath, 'utf8'));
  } finally {
    rmSync(metricsDirectory, { recursive: true, force: true });
  }
}

if (import.meta.url === pathToFileURL(process.argv[1] ?? '').href) {
  if (!process.argv[2]) throw new Error('usage: bench-xlsx-row-iterator.mjs INPUT.xlsx [MAX_ARCHIVE_ENTRY_BYTES]');
  const maxArchiveEntryBytes = process.argv[3] === undefined ? undefined : Number(process.argv[3]);
  process.stdout.write(`${JSON.stringify(benchmarkXlsxRowIterator(process.argv[2], { maxArchiveEntryBytes }))}\n`);
}
