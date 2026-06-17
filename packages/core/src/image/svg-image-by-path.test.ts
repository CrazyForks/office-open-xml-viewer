import { describe, it, expect, afterEach, vi } from 'vitest';
import { getCachedSvgImageByPath } from './svg-image-by-path';

describe('getCachedSvgImageByPath', () => {
  afterEach(() => vi.unstubAllGlobals());
  it('fetches bytes, makes an object URL, loads an <img>, dedupes by path', async () => {
    let created = 0;
    vi.stubGlobal('URL', { createObjectURL: () => { created++; return `blob:${created}`; },
                           revokeObjectURL: () => {} });
    class FakeImg { onload: (() => void) | null = null; onerror: (() => void) | null = null;
      set src(_v: string) { queueMicrotask(() => this.onload && this.onload()); } }
    vi.stubGlobal('Image', FakeImg);
    const fetchImage = vi.fn(async () => new Blob(['<svg/>'], { type: 'image/svg+xml' }));
    const a = await getCachedSvgImageByPath('word/media/i.svg', fetchImage);
    const b = await getCachedSvgImageByPath('word/media/i.svg', fetchImage);
    expect(a).toBe(b);
    expect(fetchImage).toHaveBeenCalledTimes(1);
    expect(created).toBe(1);
  });
});
