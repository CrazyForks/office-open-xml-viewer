import { describe, expect, it, vi } from 'vitest';
import { BoundedRawPartCache } from './bounded-raw-part-cache.js';

function blob(bytes: number, type = ''): Blob {
  return new Blob([new Uint8Array(bytes)], { type });
}

function deferred<T>() {
  let resolve!: (value: T) => void;
  let reject!: (error: unknown) => void;
  const promise = new Promise<T>((resolvePromise, rejectPromise) => {
    resolve = resolvePromise;
    reject = rejectPromise;
  });
  return { promise, resolve, reject };
}

describe('BoundedRawPartCache', () => {
  it('deduplicates one raw part across media MIME views', async () => {
    const pending = deferred<Blob>();
    const load = vi.fn(() => pending.promise);
    const cache = new BoundedRawPartCache({ maxEntries: 2, maxBytes: 10 });

    const image = cache.get('ppt/media/shared.bin', 'image/png', load);
    const media = cache.get('ppt/media/shared.bin', 'video/mp4', load);
    await vi.waitFor(() => expect(load).toHaveBeenCalledTimes(1));
    pending.resolve(blob(4, 'application/octet-stream'));

    await expect(image).resolves.toMatchObject({ size: 4, type: 'image/png' });
    await expect(media).resolves.toMatchObject({ size: 4, type: 'video/mp4' });
    expect(cache.usage).toEqual({ entries: 1, bytes: 4, pending: 0 });
  });

  it('preserves Blob identity for repeat calls using the retained MIME', async () => {
    const cache = new BoundedRawPartCache({ maxEntries: 1, maxBytes: 10 });
    const load = vi.fn(() => blob(2, 'image/png'));
    const first = await cache.get('ppt/media/image.png', 'image/png', load);
    const second = await cache.get('ppt/media/image.png', 'image/png', load);
    expect(second).toBe(first);
    expect(load).toHaveBeenCalledTimes(1);
  });

  it('evicts resolved parts until both count and byte ceilings fit', async () => {
    const cache = new BoundedRawPartCache({ maxEntries: 2, maxBytes: 5 });
    const loads = new Map<string, ReturnType<typeof vi.fn>>();
    const get = (path: string, bytes: number) => {
      const load = vi.fn(() => blob(bytes));
      loads.set(path, load);
      return cache.get(path, '', load);
    };

    await get('a', 2);
    await get('b', 3);
    await get('c', 4);
    expect(cache.usage).toEqual({ entries: 1, bytes: 4, pending: 0 });
    await get('a', 2);
    expect(loads.get('a')).toHaveBeenCalledTimes(1);
    expect(cache.usage.bytes).toBeLessThanOrEqual(5);
  });

  it('returns an oversized part without retaining it', async () => {
    const cache = new BoundedRawPartCache({ maxEntries: 1, maxBytes: 3 });
    const load = vi.fn(() => blob(4));

    await expect(cache.get('large', '', load)).resolves.toMatchObject({ size: 4 });
    await expect(cache.get('large', '', load)).resolves.toMatchObject({ size: 4 });
    expect(load).toHaveBeenCalledTimes(2);
    expect(cache.usage).toEqual({ entries: 0, bytes: 0, pending: 0 });
  });

  it('removes rejected loads and never admits a completion cleared in flight', async () => {
    const cache = new BoundedRawPartCache({ maxEntries: 1, maxBytes: 10 });
    const failed = vi.fn()
      .mockRejectedValueOnce(new Error('read failed'))
      .mockResolvedValueOnce(blob(2));
    await expect(cache.get('retry', '', failed)).rejects.toThrow('read failed');
    await expect(cache.get('retry', '', failed)).resolves.toMatchObject({ size: 2 });

    const stale = deferred<Blob>();
    const staleResult = cache.get('stale', '', () => stale.promise);
    await vi.waitFor(() => expect(cache.usage.pending).toBe(1));
    cache.clear();
    stale.resolve(blob(3));
    await expect(staleResult).resolves.toMatchObject({ size: 3 });
    expect(cache.usage).toEqual({ entries: 0, bytes: 0, pending: 0 });
  });
});
