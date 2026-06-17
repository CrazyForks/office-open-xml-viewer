import { describe, it, expect, vi } from 'vitest';
import { XlsxWorkbook } from './workbook';
import type { WorkerRequest } from './types';

/**
 * `XlsxWorkbook.getImage(path, mime)` routes through the persistent worker via
 * the `extractImage` message (xlsx uses the `type` discriminant, like docx),
 * wraps the returned bytes in a Blob of the requested MIME, and serves repeat
 * calls from its per-instance cache so the worker is hit at most once per path.
 * Mirrors docx's `document.image.test.ts` / pptx's `presentation.image.test.ts`.
 *
 * The constructor opens a real Worker, so we build the instance off-prototype
 * and inject a fake `bridge` whose `request` resolves an `imageExtracted`
 * response. This isolates the cache + Blob-wrapping contract from the worker.
 */
/** The subset of XlsxWorkbook this test exercises. Kept separate from the class
 *  type so the off-prototype build doesn't intersect its private fields (which
 *  would collapse to `never` under `tsc`). */
interface GetImageProbe {
  getImage(imagePath: string, mimeType: string): Promise<Blob>;
}

describe('XlsxWorkbook.getImage', () => {
  function makeWorkbook(requestImpl: (req: WorkerRequest) => unknown) {
    const request = vi.fn((build: (id: number) => WorkerRequest) =>
      Promise.resolve(requestImpl(build(1))),
    );
    // Build off the real prototype (so the real getImage runs) but inject only
    // the private collaborators it touches. Cast through unknown to avoid
    // intersecting the class's private members.
    const instance = Object.create(XlsxWorkbook.prototype) as Record<string, unknown>;
    instance.bridge = { request };
    instance.imageBlobCache = new Map<string, Promise<Blob>>();
    const wb = instance as unknown as GetImageProbe;
    return { wb, request };
  }

  const bytesFor = (s: string) => new TextEncoder().encode(s).buffer;

  it('wraps extracted bytes in a Blob of the requested MIME', async () => {
    const payload = bytesFor('PNGDATA');
    const { wb, request } = makeWorkbook((req) => {
      expect(req.type).toBe('extractImage');
      expect((req as Extract<WorkerRequest, { type: 'extractImage' }>).path).toBe(
        'xl/media/image1.png',
      );
      return { type: 'imageExtracted', id: 1, bytes: payload };
    });

    const blob = await wb.getImage('xl/media/image1.png', 'image/png');
    expect(blob).toBeInstanceOf(Blob);
    expect(blob.type).toBe('image/png');
    expect(new Uint8Array(await blob.arrayBuffer())).toEqual(
      new Uint8Array(bytesFor('PNGDATA')),
    );
    expect(request).toHaveBeenCalledTimes(1);
  });

  it('serves a second call for the same path from cache (one worker request)', async () => {
    const { wb, request } = makeWorkbook(() => ({
      type: 'imageExtracted',
      id: 1,
      bytes: bytesFor('X'),
    }));

    const a = await wb.getImage('xl/media/image1.png', 'image/png');
    const b = await wb.getImage('xl/media/image1.png', 'image/png');
    // Same cached promise → identical Blob, single underlying request.
    expect(a).toBe(b);
    expect(request).toHaveBeenCalledTimes(1);
  });
});
