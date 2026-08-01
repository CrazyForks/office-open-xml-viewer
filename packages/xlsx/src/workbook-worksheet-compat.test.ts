import { describe, expect, it, vi } from 'vitest';
import { XlsxWorkbook } from './workbook.js';
import type { ParsedWorkbook, Worksheet, WorkerRequest } from './types.js';
import type { RenderWorkerRequest } from './worker-protocol.js';

const WORKSHEET: Worksheet = {
  name: 'Sheet1',
  rows: [{
    index: 1,
    height: null,
    cells: [{ col: 1, row: 1, value: { type: 'shared', si: 0 } }],
  }],
  colWidths: {},
  rowHeights: {},
  defaultColWidth: 8.43,
  defaultRowHeight: 15,
  mergeCells: [],
  freezeRows: 0,
  freezeCols: 0,
  conditionalFormats: [],
  images: [],
  charts: [],
};

const PARSED_WORKBOOK = {
  workbook: { sheets: [{ name: 'Sheet1' }] },
  sharedStrings: [{ text: 'resolved' }],
  styles: {},
} as ParsedWorkbook;

interface WorkbookProbe {
  getWorksheet(index: number): Promise<Worksheet>;
}

function makeWorkbook(
  mode: 'main' | 'worker',
  respond: (
    request: WorkerRequest | RenderWorkerRequest,
  ) => Promise<Record<string, unknown>>,
) {
  const request = vi.fn(
    (build: (id: number) => WorkerRequest | RenderWorkerRequest) => respond(build(7)),
  );
  const instance = Object.create(XlsxWorkbook.prototype) as Record<string, unknown>;
  instance._mode = mode;
  instance.rawData = new ArrayBuffer(1);
  instance.parsedWorkbook = structuredClone(PARSED_WORKBOOK);
  instance.sheetCache = new Map();
  instance.sheetLoads = new Map();
  instance.bridge = { request };
  return { workbook: instance as unknown as WorkbookProbe, request };
}

describe('XlsxWorkbook.getWorksheet compatibility materializer', () => {
  it('decodes main-mode bytes once, resolves shared strings, and preserves cache identity', async () => {
    const { workbook, request } = makeWorkbook('main', async (message) => {
      expect(message).toMatchObject({
        type: 'parseSheet',
        id: 7,
        sheetIndex: 0,
        sheetName: 'Sheet1',
      });
      expect(message).not.toHaveProperty('data');
      return {
        type: 'parsedSheet',
        id: 7,
        worksheetJson: new TextEncoder().encode(JSON.stringify(WORKSHEET)).buffer,
      };
    });

    const first = await workbook.getWorksheet(0);
    expect(first.rows[0].cells[0].value).toEqual({ type: 'text', text: 'resolved' });
    first.rows.push({ index: 2, height: null, cells: [] });
    const second = await workbook.getWorksheet(0);

    expect(second).toBe(first);
    expect(second.rows).toHaveLength(2);
    expect(request).toHaveBeenCalledTimes(1);
  });

  it('deduplicates concurrent worker-mode materialization into one request and object', async () => {
    let resolveResponse: ((value: Record<string, unknown>) => void) | undefined;
    const response = new Promise<Record<string, unknown>>((resolve) => {
      resolveResponse = resolve;
    });
    const { workbook, request } = makeWorkbook('worker', async () => response);

    const first = workbook.getWorksheet(0);
    const second = workbook.getWorksheet(0);
    expect(request).toHaveBeenCalledTimes(1);
    resolveResponse?.({
      type: 'parsedSheet',
      id: 7,
      worksheet: structuredClone(WORKSHEET),
    });

    const [left, right] = await Promise.all([first, second]);
    expect(left).toBe(right);
    expect(left.rows[0].cells[0].value).toEqual({ type: 'text', text: 'resolved' });
  });

  it('rejects an out-of-range sheet before opening a worker operation', async () => {
    const { workbook, request } = makeWorkbook('main', async () => {
      throw new Error('must not be called');
    });

    await expect(workbook.getWorksheet(1)).rejects.toThrow('Sheet index 1 out of range');
    expect(request).not.toHaveBeenCalled();
  });

  it('does not retain a rejected in-flight materialization', async () => {
    let attempt = 0;
    const { workbook, request } = makeWorkbook('main', async () => {
      attempt += 1;
      if (attempt === 1) throw new Error('injected sheet failure');
      return {
        type: 'parsedSheet',
        id: 7,
        worksheetJson: new TextEncoder().encode(JSON.stringify(WORKSHEET)).buffer,
      };
    });

    await expect(workbook.getWorksheet(0)).rejects.toThrow('injected sheet failure');
    await expect(workbook.getWorksheet(0)).resolves.toMatchObject({ name: 'Sheet1' });
    expect(request).toHaveBeenCalledTimes(2);
  });
});
