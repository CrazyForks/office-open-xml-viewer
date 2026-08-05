import { afterEach, describe, expect, it, vi } from 'vitest';
import { CanvasViewerErrorRouter, StaticCanvasRenderDispatcher } from './canvas-viewer-mechanics.js';

afterEach(() => vi.restoreAllMocks());

describe('StaticCanvasRenderDispatcher', () => {
  it('closes a stale worker bitmap instead of committing it', () => {
    const transferFromImageBitmap = vi.fn();
    const canvas = {
      width: 0,
      height: 0,
      style: {},
      getContext: vi.fn(() => ({ transferFromImageBitmap })),
    } as unknown as HTMLCanvasElement;
    const dispatcher = new StaticCanvasRenderDispatcher(canvas, true);
    const stale = dispatcher.begin();
    dispatcher.begin();
    const bitmap = { width: 10, height: 20, close: vi.fn() } as unknown as ImageBitmap;

    expect(dispatcher.commitBitmap(stale, bitmap)).toBe(false);
    expect(bitmap.close).toHaveBeenCalledOnce();
    expect(transferFromImageBitmap).not.toHaveBeenCalled();
  });

  it('closes the worker bitmap when transfer fails', () => {
    const failure = new Error('context lost');
    const canvas = {
      width: 0,
      height: 0,
      style: {},
      getContext: vi.fn(() => ({
        transferFromImageBitmap: () => { throw failure; },
      })),
    } as unknown as HTMLCanvasElement;
    const dispatcher = new StaticCanvasRenderDispatcher(canvas, true);
    const bitmap = { width: 10, height: 20, close: vi.fn() } as unknown as ImageBitmap;

    expect(() => dispatcher.commitBitmap(dispatcher.begin(), bitmap)).toThrow(failure);
    expect(bitmap.close).toHaveBeenCalledOnce();
  });
});

describe('CanvasViewerErrorRouter', () => {
  it('normalizes failures and becomes silent after close', () => {
    const onError = vi.fn();
    const router = new CanvasViewerErrorRouter('TestViewer', onError);
    router.report('failure');
    router.close();
    router.report(new Error('late'));
    expect(onError).toHaveBeenCalledOnce();
    expect(onError.mock.calls[0][0]).toEqual(new Error('failure'));
  });
});
