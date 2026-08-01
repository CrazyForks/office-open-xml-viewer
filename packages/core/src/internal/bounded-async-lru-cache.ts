export type CacheRemovalReason = 'evicted' | 'deleted' | 'cleared';

export interface BoundedAsyncLruCacheOptions<K, V> {
  maxEntries: number;
  maxWeight: number;
  /** Returns the deterministic retained weight of one resolved value. */
  measure: (value: V) => number;
  /**
   * Observes values after the cache releases them. Exceptions are isolated from
   * cache state and ignored. The callback must not invalidate references held
   * by callers; it describes cache ownership, not caller lifetime.
   */
  onRemove?: (value: V, key: K, reason: CacheRemovalReason) => void;
}

export interface BoundedAsyncLruCacheUsage {
  entries: number;
  weight: number;
  pending: number;
}

interface RetainedEntry<V> {
  value: V;
  weight: number;
}

interface PendingEntry<V> {
  token: object;
  promise: Promise<V>;
}

function assertPositiveSafeInteger(value: number, name: string): void {
  if (!Number.isSafeInteger(value) || value <= 0) {
    throw new TypeError(`${name} must be a positive safe integer`);
  }
}

function assertWeight(value: number): void {
  if (!Number.isSafeInteger(value) || value < 0) {
    throw new TypeError('cache entry weight must be a non-negative safe integer');
  }
}

/**
 * A count- and weight-bounded LRU for values produced asynchronously.
 *
 * Pending loads are not retained entries and therefore carry no estimated
 * weight. Callers that need to bound producer work must do so at the producer
 * boundary. This cache admits a value only after the loader resolves and its
 * exact retained weight is available.
 */
export class BoundedAsyncLruCache<K, V> {
  readonly #maxEntries: number;
  readonly #maxWeight: number;
  readonly #measure: (value: V) => number;
  readonly #onRemove?: (value: V, key: K, reason: CacheRemovalReason) => void;

  readonly #entries = new Map<K, RetainedEntry<V>>();
  readonly #pending = new Map<K, PendingEntry<V>>();
  #weight = 0;

  constructor(options: BoundedAsyncLruCacheOptions<K, V>) {
    assertPositiveSafeInteger(options.maxEntries, 'maxEntries');
    assertPositiveSafeInteger(options.maxWeight, 'maxWeight');
    this.#maxEntries = options.maxEntries;
    this.#maxWeight = options.maxWeight;
    this.#measure = options.measure;
    this.#onRemove = options.onRemove;
  }

  get usage(): BoundedAsyncLruCacheUsage {
    return {
      entries: this.#entries.size,
      weight: this.#weight,
      pending: this.#pending.size,
    };
  }

  has(key: K): boolean {
    return this.#entries.has(key);
  }

  /** Returns a retained value and refreshes its LRU position. */
  get(key: K): V | undefined {
    const entry = this.#entries.get(key);
    if (entry === undefined) return undefined;
    this.#entries.delete(key);
    this.#entries.set(key, entry);
    return entry.value;
  }

  /** Returns the existing value/load, or starts exactly one load for this key. */
  getOrLoad(key: K, loader: () => V | PromiseLike<V>): Promise<V> {
    const retained = this.get(key);
    if (retained !== undefined || this.#entries.has(key)) {
      return Promise.resolve(retained as V);
    }

    const existing = this.#pending.get(key);
    if (existing !== undefined) return existing.promise;

    const token = {};
    const promise = Promise.resolve()
      .then(loader)
      .then(
        (value) => this.#admitResolved(key, token, value),
        (error: unknown) => {
          this.#forgetPending(key, token);
          throw error;
        },
      );
    this.#pending.set(key, { token, promise });
    return promise;
  }

  /** Invalidates both a retained value and any in-flight load for this key. */
  delete(key: K): boolean {
    const invalidatedPending = this.#pending.delete(key);
    const entry = this.#entries.get(key);
    if (entry === undefined) return invalidatedPending;

    this.#entries.delete(key);
    this.#weight -= entry.weight;
    this.#notifyRemoval(entry.value, key, 'deleted');
    return true;
  }

  /** Invalidates every retained value and in-flight load. */
  clear(): void {
    this.#pending.clear();
    const removed = [...this.#entries];
    this.#entries.clear();
    this.#weight = 0;
    for (const [key, entry] of removed) {
      this.#notifyRemoval(entry.value, key, 'cleared');
    }
  }

  #forgetPending(key: K, token: object): void {
    if (this.#pending.get(key)?.token === token) this.#pending.delete(key);
  }

  #admitResolved(key: K, token: object, value: V): V {
    if (this.#pending.get(key)?.token !== token) return value;
    let weight: number;
    try {
      weight = this.#measure(value);
      assertWeight(weight);
    } catch (error) {
      this.#forgetPending(key, token);
      throw error;
    }
    // `measure` is caller code and may re-enter the cache.
    if (this.#pending.get(key)?.token !== token) return value;
    this.#pending.delete(key);
    if (weight > this.#maxWeight) return value;

    const evicted: Array<[K, RetainedEntry<V>]> = [];
    while (
      this.#entries.size >= this.#maxEntries ||
      weight > this.#maxWeight - this.#weight
    ) {
      const oldest = this.#entries.entries().next().value as
        | [K, RetainedEntry<V>]
        | undefined;
      if (oldest === undefined) break;
      const [oldestKey, oldestEntry] = oldest;
      this.#entries.delete(oldestKey);
      this.#weight -= oldestEntry.weight;
      evicted.push([oldestKey, oldestEntry]);
    }

    this.#entries.set(key, { value, weight });
    this.#weight += weight;
    // Notify only after the cache reaches a valid bounded state. If a callback
    // re-enters (for example, by clearing the cache), it observes and mutates
    // the completed admission rather than accidentally resurrecting it.
    for (const [evictedKey, entry] of evicted) {
      this.#notifyRemoval(entry.value, evictedKey, 'evicted');
    }
    return value;
  }

  #notifyRemoval(value: V, key: K, reason: CacheRemovalReason): void {
    try {
      this.#onRemove?.(value, key, reason);
    } catch {
      // Cleanup/observation hooks cannot roll back a completed cache mutation.
    }
  }
}
