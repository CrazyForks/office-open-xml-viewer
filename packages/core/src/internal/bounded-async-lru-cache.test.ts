import { describe, expect, it, vi } from 'vitest';
import { BoundedAsyncLruCache } from './bounded-async-lru-cache.js';

const deferred = <T>() => {
  let resolve!: (value: T) => void;
  let reject!: (error: unknown) => void;
  const promise = new Promise<T>((res, rej) => {
    resolve = res;
    reject = rej;
  });
  return { promise, resolve, reject };
};

describe('BoundedAsyncLruCache', () => {
  it('evicts the oldest resolved entries until both count and weight fit', async () => {
    const removed: string[] = [];
    const cache = new BoundedAsyncLruCache<string, string>({
      maxEntries: 3,
      maxWeight: 6,
      measure: (value) => value.length,
      onRemove: (_value, key, reason) => removed.push(`${key}:${reason}`),
    });

    await cache.getOrLoad('a', () => 'aa');
    await cache.getOrLoad('b', () => 'bbb');
    await cache.getOrLoad('c', () => 'c');
    await cache.getOrLoad('d', () => 'dddd');

    expect(cache.usage).toEqual({ entries: 2, weight: 5, pending: 0 });
    expect(cache.has('a')).toBe(false);
    expect(cache.has('b')).toBe(false);
    expect(cache.get('c')).toBe('c');
    expect(cache.get('d')).toBe('dddd');
    expect(removed).toEqual(['a:evicted', 'b:evicted']);
  });

  it('refreshes retained entries on access before count eviction', async () => {
    const cache = new BoundedAsyncLruCache<string, number>({
      maxEntries: 2,
      maxWeight: 10,
      measure: () => 1,
    });
    await cache.getOrLoad('a', () => 1);
    await cache.getOrLoad('b', () => 2);
    expect(cache.get('a')).toBe(1);
    await cache.getOrLoad('c', () => 3);
    expect(cache.has('a')).toBe(true);
    expect(cache.has('b')).toBe(false);
    expect(cache.has('c')).toBe(true);
  });

  it('deduplicates concurrent loaders and removes rejected loads for retry', async () => {
    const first = deferred<number>();
    const loader = vi.fn(() => first.promise);
    const cache = new BoundedAsyncLruCache<string, number>({
      maxEntries: 2,
      maxWeight: 10,
      measure: () => 1,
    });

    const left = cache.getOrLoad('same', loader);
    const right = cache.getOrLoad('same', loader);
    expect(left).toBe(right);
    expect(cache.usage.pending).toBe(1);
    first.reject(new Error('load failed'));
    await expect(left).rejects.toThrow('load failed');
    await expect(right).rejects.toThrow('load failed');
    expect(cache.usage).toEqual({ entries: 0, weight: 0, pending: 0 });

    await expect(cache.getOrLoad('same', () => 7)).resolves.toBe(7);
    expect(loader).toHaveBeenCalledTimes(1);
    expect(cache.get('same')).toBe(7);
  });

  it('returns an oversized value without retaining or removing it', async () => {
    const onRemove = vi.fn();
    const cache = new BoundedAsyncLruCache<string, string>({
      maxEntries: 2,
      maxWeight: 3,
      measure: (value) => value.length,
      onRemove,
    });

    await expect(cache.getOrLoad('large', () => 'large')).resolves.toBe('large');
    expect(cache.usage).toEqual({ entries: 0, weight: 0, pending: 0 });
    expect(onRemove).not.toHaveBeenCalled();
  });

  it('does not re-admit stale resolutions after delete or clear', async () => {
    const deleted = deferred<string>();
    const cleared = deferred<string>();
    const onRemove = vi.fn();
    const cache = new BoundedAsyncLruCache<string, string>({
      maxEntries: 2,
      maxWeight: 10,
      measure: (value) => value.length,
      onRemove,
    });

    const deletedResult = cache.getOrLoad('deleted', () => deleted.promise);
    expect(cache.delete('deleted')).toBe(true);
    const clearedResult = cache.getOrLoad('cleared', () => cleared.promise);
    cache.clear();
    expect(cache.usage).toEqual({ entries: 0, weight: 0, pending: 0 });

    deleted.resolve('old-delete');
    cleared.resolve('old-clear');
    await expect(deletedResult).resolves.toBe('old-delete');
    await expect(clearedResult).resolves.toBe('old-clear');
    expect(cache.usage).toEqual({ entries: 0, weight: 0, pending: 0 });
    expect(onRemove).not.toHaveBeenCalled();

    await expect(cache.getOrLoad('deleted', () => 'new')).resolves.toBe('new');
    expect(cache.get('deleted')).toBe('new');
  });

  it('keeps a replacement same-key load when an invalidated load resolves first', async () => {
    const oldLoad = deferred<string>();
    const replacement = deferred<string>();
    const cache = new BoundedAsyncLruCache<string, string>({
      maxEntries: 2,
      maxWeight: 20,
      measure: (value) => value.length,
    });

    const oldResult = cache.getOrLoad('same', () => oldLoad.promise);
    expect(cache.delete('same')).toBe(true);
    const replacementResult = cache.getOrLoad('same', () => replacement.promise);
    expect(cache.usage.pending).toBe(1);

    oldLoad.resolve('stale');
    await expect(oldResult).resolves.toBe('stale');
    expect(cache.usage).toEqual({ entries: 0, weight: 0, pending: 1 });

    replacement.resolve('replacement');
    await expect(replacementResult).resolves.toBe('replacement');
    expect(cache.get('same')).toBe('replacement');
    expect(cache.usage).toEqual({ entries: 1, weight: 11, pending: 0 });
  });

  it('does not let a cleared same-key load overwrite a replacement that resolves first', async () => {
    const oldLoad = deferred<string>();
    const replacement = deferred<string>();
    const cache = new BoundedAsyncLruCache<string, string>({
      maxEntries: 2,
      maxWeight: 20,
      measure: (value) => value.length,
    });

    const oldResult = cache.getOrLoad('same', () => oldLoad.promise);
    cache.clear();
    const replacementResult = cache.getOrLoad('same', () => replacement.promise);

    replacement.resolve('replacement');
    await expect(replacementResult).resolves.toBe('replacement');
    oldLoad.resolve('stale');
    await expect(oldResult).resolves.toBe('stale');

    expect(cache.get('same')).toBe('replacement');
    expect(cache.usage).toEqual({ entries: 1, weight: 11, pending: 0 });
  });

  it('retains and deduplicates an undefined value', async () => {
    const loader = vi.fn(() => undefined);
    const cache = new BoundedAsyncLruCache<string, undefined>({
      maxEntries: 1,
      maxWeight: 1,
      measure: () => 0,
    });

    await expect(cache.getOrLoad('undefined', loader)).resolves.toBeUndefined();
    expect(cache.has('undefined')).toBe(true);
    expect(cache.get('undefined')).toBeUndefined();
    await expect(cache.getOrLoad('undefined', loader)).resolves.toBeUndefined();
    expect(loader).toHaveBeenCalledTimes(1);
    expect(cache.usage).toEqual({ entries: 1, weight: 0, pending: 0 });
  });

  it('removes a synchronously thrown load so the key can retry', async () => {
    const cache = new BoundedAsyncLruCache<string, number>({
      maxEntries: 1,
      maxWeight: 1,
      measure: () => 1,
    });

    await expect(cache.getOrLoad('a', () => {
      throw new Error('synchronous failure');
    })).rejects.toThrow('synchronous failure');
    expect(cache.usage).toEqual({ entries: 0, weight: 0, pending: 0 });
    await expect(cache.getOrLoad('a', () => 2)).resolves.toBe(2);
    expect(cache.get('a')).toBe(2);
  });

  it('applies the entry-count bound to zero-weight values', async () => {
    const removed: string[] = [];
    const cache = new BoundedAsyncLruCache<string, string>({
      maxEntries: 2,
      maxWeight: 1,
      measure: () => 0,
      onRemove: (_value, key, reason) => removed.push(`${key}:${reason}`),
    });

    await cache.getOrLoad('a', () => 'A');
    await cache.getOrLoad('b', () => 'B');
    await cache.getOrLoad('c', () => 'C');

    expect(cache.has('a')).toBe(false);
    expect(cache.has('b')).toBe(true);
    expect(cache.has('c')).toBe(true);
    expect(cache.usage).toEqual({ entries: 2, weight: 0, pending: 0 });
    expect(removed).toEqual(['a:evicted']);
  });

  it('removes a load after measurement failure so the key can retry', async () => {
    const measure = vi.fn((value: number) => {
      if (value === 1) throw new Error('cannot measure');
      return value;
    });
    const cache = new BoundedAsyncLruCache<string, number>({
      maxEntries: 2,
      maxWeight: 10,
      measure,
    });

    await expect(cache.getOrLoad('a', () => 1)).rejects.toThrow('cannot measure');
    expect(cache.usage).toEqual({ entries: 0, weight: 0, pending: 0 });
    await expect(cache.getOrLoad('a', () => 2)).resolves.toBe(2);
    expect(cache.get('a')).toBe(2);
  });

  it('does not re-admit when measurement re-enters and clears the cache', async () => {
    let cache!: BoundedAsyncLruCache<string, string>;
    cache = new BoundedAsyncLruCache({
      maxEntries: 2,
      maxWeight: 10,
      measure: (value) => {
        cache.clear();
        return value.length;
      },
    });

    await expect(cache.getOrLoad('a', () => 'value')).resolves.toBe('value');
    expect(cache.usage).toEqual({ entries: 0, weight: 0, pending: 0 });
  });

  it('removes each retained value once and isolates callback exceptions', async () => {
    const removals: string[] = [];
    const cache = new BoundedAsyncLruCache<string, string>({
      maxEntries: 2,
      maxWeight: 2,
      measure: () => 1,
      onRemove: (_value, key, reason) => {
        removals.push(`${key}:${reason}`);
        if (key === 'a') throw new Error('cleanup failed');
      },
    });
    await cache.getOrLoad('a', () => 'A');
    await cache.getOrLoad('b', () => 'B');
    await cache.getOrLoad('c', () => 'C');

    expect(cache.delete('b')).toBe(true);
    expect(cache.delete('b')).toBe(false);
    await cache.getOrLoad('d', () => 'D');
    cache.clear();
    cache.clear();
    expect(removals).toEqual(['a:evicted', 'b:deleted', 'c:cleared', 'd:cleared']);
    expect(cache.usage).toEqual({ entries: 0, weight: 0, pending: 0 });
  });

  it('supports clear and delete re-entry during eviction with exactly-once removal', async () => {
    const removals: string[] = [];
    const deleteResults: boolean[] = [];
    let cache!: BoundedAsyncLruCache<string, string>;
    cache = new BoundedAsyncLruCache({
      maxEntries: 2,
      maxWeight: 2,
      measure: () => 1,
      onRemove: (_value, key, reason) => {
        removals.push(`${key}:${reason}`);
        if (key === 'a') {
          deleteResults.push(cache.delete('b'));
          cache.clear();
        }
      },
    });

    await cache.getOrLoad('a', () => 'A');
    await cache.getOrLoad('b', () => 'B');
    await cache.getOrLoad('c', () => 'C');

    expect(removals).toEqual(['a:evicted', 'b:deleted', 'c:cleared']);
    expect(deleteResults).toEqual([true]);
    expect(cache.usage).toEqual({ entries: 0, weight: 0, pending: 0 });
  });

  it('accepts and accounts for the maximum safe integer boundary exactly', async () => {
    const max = Number.MAX_SAFE_INTEGER;
    const cache = new BoundedAsyncLruCache<string, number>({
      maxEntries: max,
      maxWeight: max,
      measure: (value) => value,
    });

    await cache.getOrLoad('full', () => max);
    await cache.getOrLoad('zero', () => 0);
    expect(cache.usage).toEqual({ entries: 2, weight: max, pending: 0 });

    await cache.getOrLoad('one', () => 1);
    expect(cache.has('full')).toBe(false);
    expect(cache.has('zero')).toBe(true);
    expect(cache.has('one')).toBe(true);
    expect(cache.usage).toEqual({ entries: 2, weight: 1, pending: 0 });
  });

  it('validates limits and measured weights as safe integers', async () => {
    for (const maxEntries of [0, -1, 1.5, Number.MAX_SAFE_INTEGER + 1]) {
      expect(() => new BoundedAsyncLruCache({
        maxEntries,
        maxWeight: 1,
        measure: () => 0,
      })).toThrow(/maxEntries.*positive safe integer/);
    }
    for (const maxWeight of [0, -1, 1.5, Number.MAX_SAFE_INTEGER + 1]) {
      expect(() => new BoundedAsyncLruCache({
        maxEntries: 1,
        maxWeight,
        measure: () => 0,
      })).toThrow(/maxWeight.*positive safe integer/);
    }
    for (const measured of [-1, 1.5, Number.MAX_SAFE_INTEGER + 1]) {
      const cache = new BoundedAsyncLruCache<string, string>({
        maxEntries: 1,
        maxWeight: 1,
        measure: () => measured,
      });
      await expect(cache.getOrLoad('a', () => 'a')).rejects.toThrow(
        /weight.*non-negative safe integer/,
      );
      expect(cache.usage).toEqual({ entries: 0, weight: 0, pending: 0 });
    }
  });
});
