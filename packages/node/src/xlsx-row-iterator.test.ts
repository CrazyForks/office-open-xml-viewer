import { mkdtemp, readFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { spawnSync } from 'node:child_process';
import { afterAll, beforeAll, describe, expect, it, vi } from 'vitest';
import type { Row, Worksheet } from '@silurus/ooxml-xlsx';
import { OoxmlResourceLimitError } from '@silurus/ooxml-core';
// @ts-ignore — wasm-pack generated JavaScript is local build output.
import * as xlsxWasm from '../../xlsx/src/wasm/xlsx_parser.js';
import { generateSyntheticXlsx } from '../scripts/generate-synthetic-xlsx.mjs';
import {
  iterateXlsxWorksheetRows,
  parseXlsx,
  parseSheet,
} from './xlsx.ts';

let directory = '';
let fixture = '';
let bytes: Buffer;
let worksheetBytes = 0;
let corruptBytes: Buffer;

beforeAll(async () => {
  directory = await mkdtemp(join(tmpdir(), 'ooxml-xlsx-row-'));
  fixture = join(directory, 'parity.xlsx');
  const generated = await generateSyntheticXlsx(fixture, { rows: 2049, columns: 32 });
  worksheetBytes = generated.worksheetBytes;
  bytes = await readFile(fixture);
  const corrupt = join(directory, 'malformed-tail.xlsx');
  await generateSyntheticXlsx(corrupt, {
    rows: 513,
    columns: 8,
    malformedWorksheetTail: true,
  });
  corruptBytes = await readFile(corrupt);
});

afterAll(async () => {
  if (directory) await rm(directory, { recursive: true, force: true });
});

describe('Node bounded XLSX worksheet row iterator', () => {
  it('matches the synchronous compatibility parse across multiple pulls', async () => {
    const expected = parseSheet(bytes, 0, 'Synthetic');
    const rows: Row[] = [];
    let terminal: Worksheet | undefined;
    let pulls = 0;
    for await (const chunk of iterateXlsxWorksheetRows(bytes, 0)) {
      pulls++;
      if (chunk.kind === 'rows') rows.push(...chunk.rows);
      else terminal = chunk.worksheet;
    }

    expect(pulls).toBeGreaterThan(2);
    for (const chunk of await collectChunks(bytes)) {
      if (chunk.kind === 'rows') expect(chunk.rows.length).toBeLessThanOrEqual(128);
      expect(chunk.wireBytes).toBeLessThanOrEqual(64 * 1024 * 1024);
    }
    expect(rows).toEqual(expected.rows);
    expect(terminal).toBeDefined();
    expect(terminal?.rows).toEqual([]);
    expect({ ...terminal, rows }).toEqual(expected);
  });

  it('honors an exact per-entry boundary and returns the canonical typed error one byte below it', async () => {
    let rowCount = 0;
    for await (const chunk of iterateXlsxWorksheetRows(bytes, 0, {
      resourceLimits: { maxArchiveEntryBytes: worksheetBytes, maxTotalInflatedBytes: null },
    })) {
      if (chunk.kind === 'rows') rowCount += chunk.rows.length;
    }
    expect(rowCount).toBe(2049);

    let error: unknown;
    try {
      for await (const _chunk of iterateXlsxWorksheetRows(bytes, 0, {
        resourceLimits: { maxArchiveEntryBytes: worksheetBytes - 1, maxTotalInflatedBytes: null },
      })) { /* drain */ }
    } catch (caught) {
      error = caught;
    }
    expect(error).toMatchObject({ name: 'OoxmlResourceLimitError', code: 'ooxml-resource-limit' });
    expect(String(error)).toContain('xl/worksheets/sheet1.xml');
  });

  it('cancels cleanly when a consumer returns after its first batch', async () => {
    const prototype = (xlsxWasm as unknown as {
      XlsxArchive: { prototype: { cancel_sheet_cursor(): void } };
    }).XlsxArchive.prototype;
    const cancel = vi.spyOn(prototype, 'cancel_sheet_cursor');
    const iterator = iterateXlsxWorksheetRows(bytes, 0);
    const first = await iterator.next();
    expect(first.value?.kind).toBe('rows');
    await expect(iterator.return()).resolves.toEqual({ done: true, value: undefined });
    expect(cancel).toHaveBeenCalledOnce();
    cancel.mockRestore();

    // A fresh archive/session can immediately drain after early-return cleanup.
    let rows = 0;
    for await (const chunk of iterateXlsxWorksheetRows(bytes, 0)) {
      if (chunk.kind === 'rows') rows += chunk.rows.length;
    }
    expect(rows).toBe(2049);
  });

  it('acknowledges the terminal only after the consumer advances past it', async () => {
    const prototype = (xlsxWasm as unknown as {
      XlsxArchive: { prototype: { acknowledge_sheet_cursor_terminal(): void } };
    }).XlsxArchive.prototype;
    const acknowledge = vi.spyOn(prototype, 'acknowledge_sheet_cursor_terminal');
    const iterator = iterateXlsxWorksheetRows(bytes, 0);
    try {
      for (;;) {
        const result = await iterator.next();
        expect(result.done).toBe(false);
        if (result.value?.kind !== 'finished') continue;
        expect(result.value.worksheet.rows).toEqual([]);
        expect(acknowledge).not.toHaveBeenCalled();
        await expect(iterator.next()).resolves.toEqual({ done: true, value: undefined });
        expect(acknowledge).toHaveBeenCalledOnce();
        break;
      }
    } finally {
      await iterator.return();
      acknowledge.mockRestore();
    }
  });

  it('frees the native archive exactly once after normal completion and an ordinary error', async () => {
    const free = vi.spyOn(archivePrototype(), 'free');
    try {
      for await (const _chunk of iterateXlsxWorksheetRows(bytes, 0)) { /* drain */ }
      expect(free).toHaveBeenCalledTimes(1);
      await expect(iterateXlsxWorksheetRows(bytes, 99).next()).rejects.toThrow(/out of range/);
      expect(free).toHaveBeenCalledTimes(2);
    } finally {
      free.mockRestore();
    }
  });

  it('cancels and frees when the consumer body throws', async () => {
    const cancel = vi.spyOn(archivePrototype(), 'cancel_sheet_cursor');
    const free = vi.spyOn(archivePrototype(), 'free');
    const consumerError = new Error('consumer failed');
    try {
      await expect(async () => {
        for await (const _chunk of iterateXlsxWorksheetRows(bytes, 0)) throw consumerError;
      }).rejects.toBe(consumerError);
      expect(cancel).toHaveBeenCalledOnce();
      expect(free).toHaveBeenCalledOnce();
    } finally {
      cancel.mockRestore();
      free.mockRestore();
    }
  });

  it('cancels without ACK when the consumer returns after receiving the terminal', async () => {
    const acknowledge = vi.spyOn(archivePrototype(), 'acknowledge_sheet_cursor_terminal');
    const cancel = vi.spyOn(archivePrototype(), 'cancel_sheet_cursor');
    const free = vi.spyOn(archivePrototype(), 'free');
    const iterator = iterateXlsxWorksheetRows(bytes, 0);
    try {
      for (;;) {
        const result = await iterator.next();
        if (result.value?.kind !== 'finished') continue;
        await iterator.return();
        break;
      }
      expect(acknowledge).not.toHaveBeenCalled();
      expect(cancel).toHaveBeenCalledOnce();
      expect(free).toHaveBeenCalledOnce();
    } finally {
      acknowledge.mockRestore();
      cancel.mockRestore();
      free.mockRestore();
    }
  });

  it('honors an already-aborted signal and an abort after a yielded batch', async () => {
    const cancel = vi.spyOn(archivePrototype(), 'cancel_sheet_cursor');
    const free = vi.spyOn(archivePrototype(), 'free');
    try {
      const pre = new AbortController();
      pre.abort();
      await expect(iterateXlsxWorksheetRows(bytes, 0, { signal: pre.signal }).next())
        .rejects.toMatchObject({ name: 'AbortError' });
      expect(cancel).toHaveBeenCalledTimes(1);
      expect(free).toHaveBeenCalledTimes(1);

      const post = new AbortController();
      const iterator = iterateXlsxWorksheetRows(bytes, 0, { signal: post.signal });
      expect((await iterator.next()).value?.kind).toBe('rows');
      post.abort();
      await expect(iterator.next()).rejects.toMatchObject({ name: 'AbortError' });
      expect(cancel).toHaveBeenCalledTimes(2);
      expect(free).toHaveBeenCalledTimes(2);
    } finally {
      cancel.mockRestore();
      free.mockRestore();
    }
  });

  it('cancels and frees after malformed JSON or a terminal-marker mismatch', async () => {
    const prototype = archivePrototype();
    const cancel = vi.spyOn(prototype, 'cancel_sheet_cursor');
    const free = vi.spyOn(prototype, 'free').mockImplementation(() => {
      throw new Error('free failed');
    });
    try {
      const malformed = vi.spyOn(prototype, 'pull_sheet_cursor')
        .mockImplementationOnce(() => new TextEncoder().encode('{'));
      await expect(iterateXlsxWorksheetRows(bytes, 0).next()).rejects.toBeInstanceOf(SyntaxError);
      malformed.mockRestore();

      const marker = vi.spyOn(prototype, 'sheet_cursor_pull_finished')
        .mockImplementationOnce(() => true);
      await expect(iterateXlsxWorksheetRows(bytes, 0).next())
        .rejects.toThrow(/terminal marker mismatch/);
      marker.mockRestore();

      expect(cancel).toHaveBeenCalledTimes(2);
      expect(free).toHaveBeenCalledTimes(2);
    } finally {
      vi.restoreAllMocks();
    }
  });

  it('does not mask a resource error when lifecycle cleanup also fails', async () => {
    const prototype = archivePrototype();
    const resourceError = new OoxmlResourceLimitError('synthetic row limit', {
      stage: 'parsing',
      violation: {
        format: 'xlsx',
        operation: 'parse_sheet',
        resource: 'worksheet-row',
        metric: 'projected-bytes',
        part: 'xl/worksheets/sheet1.xml',
        limit: 1,
        observed: 2,
        configurable: false,
        usage: {
          archiveEntryCount: 7,
          declaredInflatedBytes: worksheetBytes,
          distinctInflatedBytes: 1,
          operationInflatedBytes: 1,
        },
      },
    });
    const pull = vi.spyOn(prototype, 'pull_sheet_cursor').mockImplementationOnce(() => {
      throw resourceError;
    });
    const cancel = vi.spyOn(prototype, 'cancel_sheet_cursor').mockImplementation(() => {
      throw new Error('cleanup failed');
    });
    const free = vi.spyOn(prototype, 'free').mockImplementation(() => {
      throw new Error('free failed');
    });
    try {
      await expect(iterateXlsxWorksheetRows(bytes, 0).next()).rejects.toMatchObject({
        name: 'OoxmlResourceLimitError',
        code: 'ooxml-resource-limit',
      });
      expect(cancel).toHaveBeenCalled();
      expect(free).toHaveBeenCalledOnce();
    } finally {
      pull.mockRestore();
      cancel.mockRestore();
      free.mockRestore();
    }
  });

  it('matches compatibility placeholder semantics for a malformed worksheet tail', async () => {
    const expected = parseSheet(corruptBytes, 0, 'Synthetic');
    const provisionalRows: Row[] = [];
    let terminal: Worksheet | undefined;
    for await (const chunk of iterateXlsxWorksheetRows(corruptBytes, 0)) {
      if (chunk.kind === 'rows') provisionalRows.push(...chunk.rows);
      else terminal = chunk.worksheet;
    }
    expect(provisionalRows.length).toBeGreaterThan(0);
    expect(terminal?.parseError).toBeTruthy();
    expect(terminal?.rows).toEqual([]);
    // The terminal placeholder invalidates the already-yielded provisional rows.
    provisionalRows.length = 0;
    expect(provisionalRows).toEqual([]);
    expect(terminal).toEqual(expected);
  });

  it('never routes a full iterator drain through the materializing parse_sheet export', async () => {
    const parseSheetExport = vi.spyOn(xlsxWasm, 'parse_sheet');
    try {
      for await (const _chunk of iterateXlsxWorksheetRows(bytes, 0)) { /* drain */ }
      expect(parseSheetExport).not.toHaveBeenCalled();
    } finally {
      parseSheetExport.mockRestore();
    }
  });

  it('drains more than 250,000 cells without adding a Node-global cell cap', async () => {
    const output = join(directory, 'over-retained-model-cap.xlsx');
    await generateSyntheticXlsx(output, { rows: 7_813, columns: 32 });
    const large = await readFile(output);
    let rows = 0;
    let cells = 0;
    for await (const chunk of iterateXlsxWorksheetRows(large, 0)) {
      if (chunk.kind !== 'rows') continue;
      rows += chunk.rows.length;
      for (const row of chunk.rows) cells += row.cells.length;
    }
    expect(rows).toBe(7_813);
    expect(cells).toBe(250_016);
  }, 30_000);

  it('keeps the existing synchronous workbook API materialized and unchanged', () => {
    expect(parseXlsx(bytes).workbook.sheets.map((sheet) => sheet.name)).toEqual(['Synthetic']);
  });
});

type ArchivePrototype = {
  free(): void;
  acknowledge_sheet_cursor_terminal(): void;
  cancel_sheet_cursor(): void;
  pull_sheet_cursor(rowCredit: number): Uint8Array;
  sheet_cursor_pull_finished(): boolean;
};

function archivePrototype(): ArchivePrototype {
  return (xlsxWasm as unknown as { XlsxArchive: { prototype: ArchivePrototype } })
    .XlsxArchive.prototype;
}

async function collectChunks(input: Buffer) {
  const chunks = [];
  for await (const chunk of iterateXlsxWorksheetRows(input, 0)) chunks.push(chunk);
  return chunks;
}

describe.runIf(process.env.RUN_XLSX_ISSUE_1102 === '1')('Issue #1102 synthetic boundary and memory benchmark', () => {
  it('rejects an over-default inflated worksheet and records fresh-process metrics', async () => {
    const large = join(directory, 'over-default.xlsx');
    const target = 128 * 1024 * 1024 + 1;
    const generated = await generateSyntheticXlsx(large, {
      rows: 2049,
      columns: 32,
      targetWorksheetBytes: target,
    });
    expect(generated.worksheetBytes).toBe(target);
    const largeBytes = await readFile(large);
    await expect(async () => {
      for await (const _chunk of iterateXlsxWorksheetRows(largeBytes, 0)) { /* drain */ }
    }).rejects.toMatchObject({ code: 'ooxml-resource-limit' });

    const child = spawnSync(process.execPath, [
      join(import.meta.dirname, '../scripts/bench-xlsx-row-iterator.mjs'),
      large,
      String(target),
    ], { encoding: 'utf8', env: { ...process.env, NODE_NO_WARNINGS: '1' } });
    expect(child.stderr).toBe('');
    expect(child.status).toBe(0);
    const metrics = JSON.parse(child.stdout) as Record<string, number>;
    expect(metrics).toMatchObject({ rows: 2049, cells: 2049 * 32 });
    expect(metrics.pulls).toBeGreaterThan(2);
    expect(metrics.wireBytes).toBeGreaterThan(0);
    expect(metrics.wasmPages).toBeGreaterThan(0);
    expect(metrics.maxRssKiB).toBeGreaterThan(0);
    expect(metrics.elapsedMs).toBeGreaterThan(0);

    const maxRssMiB = process.env.XLSX_BENCH_MAX_RSS_MIB;
    if (maxRssMiB) expect(metrics.maxRssKiB).toBeLessThanOrEqual(Number(maxRssMiB) * 1024);
  }, 120_000);
});
