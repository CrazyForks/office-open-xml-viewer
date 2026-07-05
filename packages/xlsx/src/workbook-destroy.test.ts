import { describe, it, expect, vi, afterEach } from 'vitest';
import { WorkerBridge, preloadGoogleFonts, type WorkerLike, type FontPreloadEntry } from '@silurus/ooxml-core';
import { XlsxWorkbook } from './workbook.js';

/**
 * `XlsxWorkbook.destroy()` tears the parser worker down via
 * `WorkerBridge.terminate()`. That must reject any request still in flight so a
 * `load()` / image extraction awaiting the worker cannot hang after the
 * workbook is disposed. Pinned with a real {@link WorkerBridge} over an
 * in-memory worker (the constructor opens a real Worker, so we build
 * off-prototype and inject the collaborators destroy() touches — the pattern
 * from `workbook.image.test.ts`).
 */

class SilentWorker implements WorkerLike {
  postMessage(): void {}
  addEventListener(): void {}
  removeEventListener(): void {}
  terminated = false;
  terminate(): void {
    this.terminated = true;
  }
}

interface DestroyProbe {
  destroy(): void;
}

// ── Fake FontFaceSet so destroy()'s Google-Fonts release is observable ───────
const G = globalThis as Record<string, unknown>;
const ORIG_FONTS = { document: G.document, self: G.self, fetch: G.fetch, FontFace: G.FontFace };
afterEach(() => {
  G.document = ORIG_FONTS.document;
  G.self = ORIG_FONTS.self;
  G.fetch = ORIG_FONTS.fetch;
  G.FontFace = ORIG_FONTS.FontFace;
});

const CSS = `@font-face { font-family: 'Carlito'; font-style: normal; font-weight: 400; src: url(https://fonts.gstatic.com/s/carlito/y.woff2) format('woff2'); }`;
interface FakeFace { family: string }
function installFontFaceSet(): { added: FakeFace[] } {
  const added: FakeFace[] = [];
  class FakeFontFace {
    constructor(public family: string, public source: string, public descriptors?: object) {}
    load(): Promise<FakeFontFace> { return Promise.resolve(this); }
  }
  const set = {
    add: (f: FakeFace) => { added.push(f); },
    delete: (f: FakeFace) => { const i = added.indexOf(f); if (i >= 0) added.splice(i, 1); return i >= 0; },
    [Symbol.iterator]() { return added[Symbol.iterator](); },
    ready: Promise.resolve(),
  };
  G.FontFace = FakeFontFace;
  G.document = { fonts: set };
  G.fetch = async () => ({ ok: true, text: async () => CSS });
  delete G.self;
  return { added };
}
const MAP: Record<string, FontPreloadEntry> = {
  calibri: { url: 'https://fonts.googleapis.com/css2?family=Carlito', loadFamily: 'Carlito' },
};

describe('XlsxWorkbook.destroy() — rejects in-flight worker requests', () => {
  function makeWorkbook() {
    const worker = new SilentWorker();
    const bridge = new WorkerBridge<{ id?: number }>(worker, {
      correlate: (r) => r.id,
    });
    const instance = Object.create(XlsxWorkbook.prototype) as Record<string, unknown>;
    instance.bridge = bridge;
    // Fields destroy() clears after terminate(); undefined would throw.
    instance.sheetCache = new Map();
    instance.imageCache = new Map();
    instance.imageBlobCache = new Map();
    instance.googleFontFaces = [];
    instance._fetchImage = () => Promise.resolve(new Blob());
    return { wb: instance as unknown as DestroyProbe, bridge, worker };
  }

  it('rejects a pending request when destroy() terminates the worker', async () => {
    const { wb, bridge, worker } = makeWorkbook();
    const inFlight = bridge.request((id) => ({ id }));
    wb.destroy();
    expect(worker.terminated).toBe(true);
    await expect(inFlight).rejects.toThrow(/terminated/i);
  });

  it('is safe to call destroy() twice', () => {
    const { wb } = makeWorkbook();
    wb.destroy();
    expect(() => wb.destroy()).not.toThrow();
  });

  // Wiring guard: destroy() must actually release the Google-Fonts substitutes
  // the workbook preloaded. The other tests set `googleFontFaces = []`, so they
  // never exercise the unload branch — a dropped call would go unnoticed.
  it('destroy() releases the workbook’s Google fonts from the FontFaceSet', async () => {
    const { added } = installFontFaceSet();
    const held = await preloadGoogleFonts(['Calibri'], MAP);
    expect(added).toHaveLength(1);

    const { wb } = makeWorkbook();
    (wb as unknown as { googleFontFaces: FontFace[] }).googleFontFaces = held;
    wb.destroy();

    expect(added).toHaveLength(0);
    expect((wb as unknown as { googleFontFaces: FontFace[] }).googleFontFaces).toHaveLength(0);
  });
});

/**
 * `destroy()` must close every cached `ImageBitmap` (GPU-backed) before
 * dropping `imageCache`, not just `.clear()` it — a bare `.clear()` drops the
 * last reference without releasing the GPU backing, leaking it until GC (which
 * is not guaranteed to run promptly for GPU-backed objects). See
 * `closeAndClearImageCache` in render-orchestrator.ts.
 */
describe('XlsxWorkbook.destroy() — closes cached ImageBitmaps (GPU-leak guard)', () => {
  function makeWorkbookWithImageCache(imageCache: Map<string, CanvasImageSource | null>) {
    const worker = new SilentWorker();
    const bridge = new WorkerBridge<{ id?: number }>(worker, {
      correlate: (r) => r.id,
    });
    const instance = Object.create(XlsxWorkbook.prototype) as Record<string, unknown>;
    instance.bridge = bridge;
    instance.sheetCache = new Map();
    instance.imageCache = imageCache;
    instance.imageBlobCache = new Map();
    instance.googleFontFaces = [];
    instance._fetchImage = () => Promise.resolve(new Blob());
    return instance as unknown as DestroyProbe;
  }

  it('calls .close() on each cached ImageBitmap and empties the cache', () => {
    const close1 = vi.fn();
    const close2 = vi.fn();
    const bmp1 = { close: close1 } as unknown as ImageBitmap;
    const bmp2 = { close: close2 } as unknown as ImageBitmap;
    const imageCache = new Map<string, CanvasImageSource | null>([
      ['xl/media/image1.png', bmp1],
      ['xl/media/image2.png', bmp2],
    ]);
    const wb = makeWorkbookWithImageCache(imageCache);

    wb.destroy();

    expect(close1).toHaveBeenCalledTimes(1);
    expect(close2).toHaveBeenCalledTimes(1);
    expect(imageCache.size).toBe(0);
  });

  it('skips a cached null (unsupported metafile) without throwing', () => {
    const imageCache = new Map<string, CanvasImageSource | null>([
      ['xl/media/diagram.emf', null],
    ]);
    const wb = makeWorkbookWithImageCache(imageCache);
    expect(() => wb.destroy()).not.toThrow();
    expect(imageCache.size).toBe(0);
  });

  it('is safe to destroy() twice — the second call does not re-close an already-closed bitmap', () => {
    const close = vi.fn();
    const bmp = { close } as unknown as ImageBitmap;
    const imageCache = new Map<string, CanvasImageSource | null>([['xl/media/image1.png', bmp]]);
    const wb = makeWorkbookWithImageCache(imageCache);

    wb.destroy();
    expect(() => wb.destroy()).not.toThrow();
    // The map is empty after the first destroy(), so the second pass has
    // nothing to iterate — close() is called exactly once total.
    expect(close).toHaveBeenCalledTimes(1);
  });
});
