import { describe, expect, it, vi } from 'vitest';
import type { WorkerLike } from './bridge.js';
import { disposeRejectedLoad } from './rejected-load.js';

function worker(terminate = vi.fn()): WorkerLike {
  return {
    postMessage() {},
    addEventListener() {},
    removeEventListener() {},
    terminate,
  };
}

describe('disposeRejectedLoad', () => {
  it('prefers the engine disposer when construction completed', () => {
    const terminate = vi.fn();
    const dispose = vi.fn();

    disposeRejectedLoad(worker(terminate), dispose);

    expect(dispose).toHaveBeenCalledOnce();
    expect(terminate).not.toHaveBeenCalled();
  });

  it('terminates directly when no engine exists', () => {
    const terminate = vi.fn();

    disposeRejectedLoad(worker(terminate));

    expect(terminate).toHaveBeenCalledOnce();
  });

  it('falls back to termination when the engine disposer throws', () => {
    const terminate = vi.fn();

    expect(() =>
      disposeRejectedLoad(worker(terminate), () => {
        throw new Error('dispose failed');
      }),
    ).not.toThrow();
    expect(terminate).toHaveBeenCalledOnce();
  });

  it('does not throw when direct termination fails', () => {
    expect(() =>
      disposeRejectedLoad(
        worker(
          vi.fn(() => {
            throw new Error('terminate failed');
          }),
        ),
      ),
    ).not.toThrow();
  });
});
