import { describe, expect, it, vi } from 'vitest';
import { withDecodedImageSlot } from './decode-gate';

describe('withDecodedImageSlot', () => {
  it('hands a released slot to its waiter before admitting a new request', async () => {
    const owner = {};
    let active = 0;
    let maximum = 0;
    const releases = new Map<string, () => void>();
    const run = (name: string) => withDecodedImageSlot(owner, async () => {
      active++;
      maximum = Math.max(maximum, active);
      await new Promise<void>((resolve) => releases.set(name, resolve));
      active--;
    });

    const a = run('a');
    const b = run('b');
    await vi.waitFor(() => expect(active).toBe(2));
    const c = run('c');

    // A's continuation releases C's gate waiter. D is deliberately queued
    // immediately behind A's continuation and before C resumes, reproducing
    // the handoff race that an increment-after-await semaphore permits.
    releases.get('a')?.();
    let d!: Promise<void>;
    queueMicrotask(() => { d = run('d'); });
    await vi.waitFor(() => expect(releases.has('c')).toBe(true));
    expect(maximum).toBe(2);

    releases.get('b')?.();
    await vi.waitFor(() => expect(releases.has('d')).toBe(true));
    releases.get('c')?.();
    releases.get('d')?.();
    await Promise.all([a, b, c, d]);
    expect(maximum).toBe(2);
  });
});
