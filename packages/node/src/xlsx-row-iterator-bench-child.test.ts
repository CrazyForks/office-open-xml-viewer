import { readFile, writeFile } from 'node:fs/promises';
import { performance } from 'node:perf_hooks';
import { expect, test } from 'vitest';
// @ts-ignore — wasm-pack generated JavaScript is local build output.
import * as xlsxWasm from '../../xlsx/src/wasm/xlsx_parser.js';
import { iterateXlsxWorksheetRows } from './xlsx.ts';
import { wasmMemoryPages } from './wasm-loader.ts';

test.skipIf(!process.env.OOXML_XLSX_BENCH_INPUT)('records bounded XLSX iterator metrics', async () => {
  const bytes = await readFile(process.env.OOXML_XLSX_BENCH_INPUT as string);
  const started = performance.now();
  let rows = 0;
  let cells = 0;
  let pulls = 0;
  let wireBytes = 0;
  const maxArchiveEntryBytes = process.env.OOXML_XLSX_BENCH_MAX_ENTRY;
  for await (const chunk of iterateXlsxWorksheetRows(bytes, 0, {
    ...(maxArchiveEntryBytes === undefined
      ? {}
      : { resourceLimits: { maxArchiveEntryBytes: Number(maxArchiveEntryBytes), maxTotalInflatedBytes: null } }),
  })) {
    pulls++;
    wireBytes += chunk.wireBytes;
    if (chunk.kind === 'rows') {
      rows += chunk.rows.length;
      for (const row of chunk.rows) cells += row.cells.length;
    }
  }
  const metrics = {
    elapsedMs: performance.now() - started,
    rows,
    cells,
    pulls,
    wireBytes,
    wasmPages: wasmMemoryPages(xlsxWasm),
    maxRssKiB: process.resourceUsage().maxRSS,
  };
  expect(metrics.wasmPages).toBeGreaterThan(0);
  await writeFile(process.env.OOXML_XLSX_BENCH_OUTPUT as string, JSON.stringify(metrics));
});
