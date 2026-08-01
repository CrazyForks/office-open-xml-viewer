import type { DeepReadonly } from './types.js';

function assertPlainData(
  value: unknown,
  path: string,
  visiting = new WeakSet<object>(),
  completed = new WeakSet<object>(),
): void {
  if (
    value === null
    || value === undefined
    || typeof value === 'string'
    || typeof value === 'boolean'
  ) return;
  if (typeof value === 'number') {
    if (!Number.isFinite(value)) throw new TypeError(`${path} must contain finite numbers`);
    return;
  }
  if (typeof value !== 'object') {
    throw new TypeError(`${path} must be structured-clone-safe plain data`);
  }
  if (visiting.has(value)) {
    throw new TypeError(`${path} must be structured-clone-safe plain data`);
  }
  if (completed.has(value)) return;
  const prototype = Object.getPrototypeOf(value);
  if (!Array.isArray(value) && prototype !== Object.prototype && prototype !== null) {
    throw new TypeError(`${path} must be structured-clone-safe plain data`);
  }
  if (Object.getOwnPropertySymbols(value).length !== 0) {
    throw new TypeError(`${path} must contain only enumerable string data properties`);
  }
  visiting.add(value);
  try {
    for (const key of Object.getOwnPropertyNames(value)) {
      if (Array.isArray(value) && key === 'length') continue;
      const descriptor = Object.getOwnPropertyDescriptor(value, key);
      if (!descriptor || !descriptor.enumerable || !('value' in descriptor)) {
        throw new TypeError(`${path}.${key} must be an enumerable data property`);
      }
      assertPlainData(descriptor.value, `${path}.${key}`, visiting, completed);
    }
  } finally {
    visiting.delete(value);
  }
  completed.add(value);
}

export function deepFreezePlainData<T>(
  value: T,
  seen = new WeakSet<object>(),
): DeepReadonly<T> {
  if (value === null || typeof value !== 'object' || seen.has(value)) {
    return value as DeepReadonly<T>;
  }
  seen.add(value);
  for (const child of Object.values(value)) deepFreezePlainData(child, seen);
  return Object.freeze(value) as DeepReadonly<T>;
}

export function snapshotPlainData<T>(value: T, label: string): DeepReadonly<T> {
  assertPlainData(value, label);
  try {
    return deepFreezePlainData(structuredClone(value));
  } catch {
    throw new TypeError(`${label} must be structured-clone-safe plain data`);
  }
}

/** Validate and recursively seal builder-owned plain data in place. Unlike
 * snapshotPlainData this has no second structured-clone peak; callers must own
 * the supplied graph and must not expose it for later mutation. */
export function sealPlainData<T>(value: T, label: string): DeepReadonly<T> {
  assertPlainData(value, label);
  return deepFreezePlainData(value);
}
