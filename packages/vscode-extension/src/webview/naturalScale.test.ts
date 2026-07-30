import { describe, expect, it, vi } from 'vitest';
import { loadAtNaturalScale } from './naturalScale';

describe('loadAtNaturalScale', () => {
  it('latches absolute 100% before loading instead of fitting to the view width', async () => {
    const calls: string[] = [];
    const buffer = new ArrayBuffer(4);
    const viewer = {
      setScale: vi.fn((scale: number) => {
        calls.push(`scale:${scale}`);
      }),
      load: vi.fn(async (source: string | ArrayBuffer) => {
        expect(source).toBe(buffer);
        calls.push('load');
      }),
    };

    await loadAtNaturalScale(viewer, buffer);

    expect(viewer.setScale).toHaveBeenCalledWith(1);
    expect(calls).toEqual(['scale:1', 'load']);
  });
});
