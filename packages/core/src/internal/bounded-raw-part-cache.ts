import { BoundedAsyncLruCache } from './bounded-async-lru-cache.js';

export interface BoundedRawPartCacheOptions {
  /** Maximum resolved package parts retained by the library. */
  readonly maxEntries: number;
  /** Maximum sum of retained raw Blob bytes. */
  readonly maxBytes: number;
}

export interface BoundedRawPartCacheUsage {
  readonly entries: number;
  readonly bytes: number;
  readonly pending: number;
}

/**
 * Count- and byte-bounded ownership for lazily extracted raw package parts.
 *
 * Image/media callers share one retained byte owner for the same OPC part.
 * The first loader's MIME is retained, so repeat calls through the same API
 * preserve Blob identity; a different MIME receives a cheap `slice()` view.
 * Pending loads are deduplicated,
 * rejected loads are retryable, oversized values are returned but not retained,
 * and `clear()` prevents stale completions from entering a new generation.
 */
export class BoundedRawPartCache {
  readonly #cache: BoundedAsyncLruCache<string, Blob>;

  constructor(options: BoundedRawPartCacheOptions) {
    this.#cache = new BoundedAsyncLruCache({
      maxEntries: options.maxEntries,
      maxWeight: options.maxBytes,
      measure: (blob) => blob.size,
    });
  }

  get usage(): BoundedRawPartCacheUsage {
    const usage = this.#cache.usage;
    return { entries: usage.entries, bytes: usage.weight, pending: usage.pending };
  }

  async get(
    partPath: string,
    mimeType: string,
    loadRaw: () => Blob | PromiseLike<Blob>,
  ): Promise<Blob> {
    if (typeof partPath !== 'string' || partPath.length === 0) {
      throw new TypeError('raw package part path must be a non-empty string');
    }
    if (typeof mimeType !== 'string') {
      throw new TypeError('raw package part MIME type must be a string');
    }
    const raw = await this.#cache.getOrLoad(partPath, async () => {
      const loaded = await loadRaw();
      if (!(loaded instanceof Blob)) {
        throw new TypeError('raw package part loader must return a Blob');
      }
      return loaded;
    });
    return mimeType === '' || raw.type === mimeType
      ? raw
      : raw.slice(0, raw.size, mimeType);
  }

  clear(): void {
    this.#cache.clear();
  }
}
