import { describe, expect, it, vi } from 'vitest';
import { loadAtContainerFit, loadAtNaturalScale } from './naturalScale';

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

  it('leaves the scale unlatched when loading a presentation to fit its container', async () => {
    const buffer = new ArrayBuffer(4);
    const viewer = {
      setScale: vi.fn(),
      load: vi.fn(async (source: string | ArrayBuffer) => {
        expect(source).toBe(buffer);
      }),
    };

    await loadAtContainerFit(viewer, buffer);

    expect(viewer.setScale).not.toHaveBeenCalled();
    expect(viewer.load).toHaveBeenCalledOnce();
  });
});
