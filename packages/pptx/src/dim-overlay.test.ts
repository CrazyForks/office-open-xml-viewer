import { describe, it, expect } from 'vitest';
import { applyDimOverlay } from './renderer.js';
import type { DimOptions } from './types.js';

/** Recording mock 2D context: captures the ordered calls applyDimOverlay makes. */
function mockCtx() {
  const calls: Array<[string, ...unknown[]]> = [];
  const ctx = {
    save: () => calls.push(['save']),
    restore: () => calls.push(['restore']),
    fillRect: (x: number, y: number, w: number, h: number) => calls.push(['fillRect', x, y, w, h]),
    set globalAlpha(v: number) { calls.push(['globalAlpha', v]); },
    set fillStyle(v: string) { calls.push(['fillStyle', v]); },
  } as unknown as CanvasRenderingContext2D;
  return { ctx, calls };
}

describe('applyDimOverlay', () => {
  it('fills the whole canvas with the dim color at the given opacity, save/restore-wrapped', () => {
    const { ctx, calls } = mockCtx();
    const dim: DimOptions = { color: '#ffffff', opacity: 0.6 };
    applyDimOverlay(ctx, dim, 960, 540);
    expect(calls).toEqual([
      ['save'],
      ['globalAlpha', 0.6],
      ['fillStyle', '#ffffff'],
      ['fillRect', 0, 0, 960, 540],
      ['restore'],
    ]);
  });
});
