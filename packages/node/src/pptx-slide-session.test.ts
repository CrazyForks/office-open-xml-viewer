import { readFile } from 'node:fs/promises';
import { beforeAll, describe, expect, it, vi } from 'vitest';
// @ts-ignore — wasm-pack generated JavaScript is local build output.
import * as pptxWasm from '../../pptx/src/wasm/pptx_parser.js';
import {
  openPptxPresentation,
  parsePptx,
} from './pptx.ts';

let bytes: Buffer;

beforeAll(async () => {
  bytes = await readFile(new URL('../../pptx/public/demo/sample-1.pptx', import.meta.url));
});

describe('Node bounded PPTX presentation session', () => {
  it('matches the materializing compatibility model one canonical slide at a time', async () => {
    const expected = parsePptx(bytes);
    const session = await openPptxPresentation(bytes);
    expect(session.slideCount).toBe(expected.slides.length);
    expect(session.slideWidth).toBe(expected.slideWidth);
    expect(session.slideHeight).toBe(expected.slideHeight);

    let index = 0;
    for await (const slide of session) {
      expect(slide).toEqual(expected.slides[index]);
      expect(session.resourceUsage?.operationInflatedBytes).toBeGreaterThan(0);
      index += 1;
    }
    expect(index).toBe(expected.slides.length);
  });

  it('does not route a bounded drain through the materializing parse export', async () => {
    const parse = vi.spyOn(pptxWasm, 'parse_pptx');
    try {
      const session = await openPptxPresentation(bytes);
      for await (const _slide of session) { /* drain */ }
      expect(parse).not.toHaveBeenCalled();
    } finally {
      parse.mockRestore();
    }
  });

  it('closes and frees exactly once after completion, early return, and explicit close', async () => {
    const free = vi.spyOn(archivePrototype(), 'free');
    try {
      const completed = await openPptxPresentation(bytes);
      for await (const _slide of completed) { /* drain */ }
      await completed.close();
      expect(free).toHaveBeenCalledTimes(1);

      const early = await openPptxPresentation(bytes);
      for await (const _slide of early) break;
      await early.close();
      expect(free).toHaveBeenCalledTimes(2);

      const unopened = await openPptxPresentation(bytes);
      await unopened.close();
      await unopened.close();
      expect(free).toHaveBeenCalledTimes(3);
    } finally {
      free.mockRestore();
    }
  });

  it('is one-pass and rejects iteration after ownership has closed', async () => {
    const session = await openPptxPresentation(bytes);
    const iterator = session.slides();
    expect((await iterator.next()).value?.index).toBe(0);
    await iterator.return();
    await expect(session.slides().next()).rejects.toThrow(/closed|one-pass/);
  });

  it('normalizes package limits and aborts deterministically', async () => {
    await expect(openPptxPresentation(bytes, {
      resourceLimits: { maxArchiveEntryBytes: 1, maxTotalInflatedBytes: null },
    })).rejects.toMatchObject({
      name: 'OoxmlResourceLimitError',
      code: 'ooxml-resource-limit',
    });

    const before = new AbortController();
    before.abort();
    await expect(openPptxPresentation(bytes, { signal: before.signal }))
      .rejects.toMatchObject({ name: 'AbortError' });

    const after = new AbortController();
    const session = await openPptxPresentation(bytes, { signal: after.signal });
    const iterator = session.slides();
    expect((await iterator.next()).value?.index).toBe(0);
    after.abort();
    await expect(iterator.next()).rejects.toMatchObject({ name: 'AbortError' });
  });
});

type ArchivePrototype = {
  free(): void;
};

function archivePrototype(): ArchivePrototype {
  return (pptxWasm as unknown as { PptxArchive: { prototype: ArchivePrototype } })
    .PptxArchive.prototype;
}
