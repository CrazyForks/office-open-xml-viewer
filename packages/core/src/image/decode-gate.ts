import { MAX_CONCURRENT_IMAGE_DECODES } from './pixel-budget.js';

interface DecodeGate {
  active: number;
  readonly waiters: Array<() => void>;
}

const gates = new WeakMap<object, DecodeGate>();

/** Run all decoded-image work under one per-document concurrency gate. */
export async function withDecodedImageSlot<T>(
  owner: object,
  operation: () => Promise<T>,
): Promise<T> {
  let gate = gates.get(owner);
  if (!gate) {
    gate = { active: 0, waiters: [] };
    gates.set(owner, gate);
  }
  if (gate.active >= MAX_CONCURRENT_IMAGE_DECODES || gate.waiters.length > 0) {
    await new Promise<void>((resolve) => gate!.waiters.push(resolve));
  } else {
    gate.active++;
  }
  try {
    // Let callers publish their cache entry before a queued producer checks
    // whether eviction or teardown superseded it.
    await Promise.resolve();
    return await operation();
  } finally {
    const next = gate.waiters.shift();
    if (next) {
      // Transfer this exact slot to the oldest waiter. Keeping `active`
      // unchanged prevents a new request from stealing the transiently free
      // slot before the resumed waiter reaches its next microtask.
      next();
    } else {
      gate.active--;
      if (gate.active === 0) gates.delete(owner);
    }
  }
}
