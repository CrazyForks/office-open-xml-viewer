import type { WorkerLike } from './bridge.js';

/**
 * Release a worker owned by a factory whose load did not return an engine.
 *
 * Once the factory rejects, ownership cannot reach the caller. Prefer the
 * partially constructed engine's full disposer because it also releases font,
 * bitmap, object-URL, and bridge state. If construction did not finish, or that
 * disposer itself fails, terminate the worker directly. Cleanup is deliberately
 * best-effort and never replaces the original load error.
 */
export function disposeRejectedLoad(worker: WorkerLike, dispose?: () => void): void {
  if (dispose) {
    try {
      dispose();
      return;
    } catch {
      // Fall through to the worker-level ownership boundary.
    }
  }
  try {
    worker.terminate();
  } catch {
    // The factory must rethrow the original load failure, not a cleanup failure.
  }
}
