import { readFile } from 'node:fs/promises';
import { Canvas, loadImage } from 'skia-canvas';
import { beforeAll, describe, expect, it, vi } from 'vitest';
// @ts-ignore — wasm-pack generated JavaScript is local build output.
import * as docxWasm from '../../docx/src/wasm/docx_parser.js';
import { createLayoutServices } from '../../docx/src/layout-runtime.ts';
import { layoutDocument } from '../../docx/src/document-layout.ts';
import {
  openDocxDocument,
  parseDocx,
} from './docx.ts';
import type { NodeCanvasFactory } from './render.ts';

const factory: NodeCanvasFactory = {
  createCanvas: (width, height) =>
    new Canvas(width, height) as unknown as ReturnType<NodeCanvasFactory['createCanvas']>,
  loadImage: (async (buffer: ArrayBuffer | Uint8Array | Buffer) =>
    loadImage(Buffer.from(buffer as Uint8Array))) as unknown as NodeCanvasFactory['loadImage'],
};

let bytes: Buffer;

beforeAll(async () => {
  bytes = await readFile(new URL('../../docx/public/demo/sample-1.docx', import.meta.url));
});

describe('Node bounded DOCX document session', () => {
  it('matches compatibility pagination and renders one caller-owned canvas at a time', async () => {
    const expected = parseDocx(bytes);
    const measure = factory.createCanvas(1, 1).getContext('2d');
    const expectedPages = layoutDocument(
      expected,
      createLayoutServices(expected, { measureContext: measure }),
      { currentDateMs: 0 },
    ).pages.length;

    const session = await openDocxDocument(bytes, { factory, currentDate: 0 });
    expect(session.pageCount).toBe(expectedPages);
    expect(session.resourceUsage?.operationInflatedBytes).toBeGreaterThan(0);

    let count = 0;
    for await (const page of session) {
      expect(page.pageIndex).toBe(count);
      expect(page.widthPt).toBeGreaterThan(0);
      expect(page.heightPt).toBeGreaterThan(0);
      expect(page.canvas.width).toBeGreaterThan(0);
      expect(page.canvas.height).toBeGreaterThan(0);
      count += 1;
    }
    expect(count).toBe(expectedPages);
  });

  it('does not route the bounded session through the materializing parse export', async () => {
    const parse = vi.spyOn(docxWasm, 'parse_docx');
    try {
      const session = await openDocxDocument(bytes, { factory, currentDate: 0 });
      await session.close();
      expect(parse).not.toHaveBeenCalled();
    } finally {
      parse.mockRestore();
    }
  });

  it('frees exactly once after completion, early return, and explicit close', async () => {
    const free = vi.spyOn(archivePrototype(), 'free');
    try {
      const completed = await openDocxDocument(bytes, { factory, currentDate: 0 });
      for await (const _page of completed.pages({ dpr: 1 })) { /* drain */ }
      await completed.close();
      expect(free).toHaveBeenCalledTimes(1);

      const early = await openDocxDocument(bytes, { factory, currentDate: 0 });
      for await (const _page of early.pages({ dpr: 1 })) break;
      await early.close();
      expect(free).toHaveBeenCalledTimes(2);

      const unopened = await openDocxDocument(bytes, { factory, currentDate: 0 });
      await unopened.close();
      await unopened.close();
      expect(free).toHaveBeenCalledTimes(3);
    } finally {
      free.mockRestore();
    }
  });

  it('reports metrics for the owned session when it closes', async () => {
    const onResourceMetrics = vi.fn();
    const session = await openDocxDocument(bytes, {
      factory,
      currentDate: 0,
      onResourceMetrics,
    });
    expect(onResourceMetrics).not.toHaveBeenCalled();
    await session.close();
    expect(onResourceMetrics).toHaveBeenCalledOnce();
    expect(onResourceMetrics).toHaveBeenCalledWith(expect.objectContaining({
      format: 'docx',
      scope: 'session',
      status: 'ok',
    }));
  });

  it('normalizes package limits and aborts between page pulls', async () => {
    await expect(openDocxDocument(bytes, {
      factory,
      resourceLimits: { maxArchiveEntryBytes: 1, maxTotalInflatedBytes: null },
    })).rejects.toMatchObject({
      name: 'OoxmlResourceLimitError',
      code: 'ooxml-resource-limit',
    });

    const before = new AbortController();
    before.abort();
    await expect(openDocxDocument(bytes, { factory, signal: before.signal }))
      .rejects.toMatchObject({ name: 'AbortError' });

    const after = new AbortController();
    const session = await openDocxDocument(bytes, {
      factory,
      currentDate: 0,
      signal: after.signal,
    });
    const iterator = session.pages({ dpr: 1 });
    expect((await iterator.next()).value?.pageIndex).toBe(0);
    after.abort();
    await expect(iterator.next()).rejects.toMatchObject({ name: 'AbortError' });
  });
});

type ArchivePrototype = {
  free(): void;
};

function archivePrototype(): ArchivePrototype {
  return (docxWasm as unknown as { DocxArchive: { prototype: ArchivePrototype } })
    .DocxArchive.prototype;
}
