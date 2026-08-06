import { afterEach, describe, expect, it, vi } from 'vitest';
import { CanvasViewerErrorRouter, StaticCanvasRenderDispatcher } from './canvas-viewer-mechanics.js';

afterEach(() => vi.restoreAllMocks());

describe('StaticCanvasRenderDispatcher', () => {
  it('acquires the bitmap context once and avoids redundant backing-store resets', () => {
    const bitmapContext = { transferFromImageBitmap: vi.fn() };
    let width = 10;
    let height = 20;
    const setWidth = vi.fn((value: number) => { width = value; });
    const setHeight = vi.fn((value: number) => { height = value; });
    const canvas = {
      get width() { return width; },
      set width(value: number) { setWidth(value); },
      get height() { return height; },
      set height(value: number) { setHeight(value); },
      style: {},
      getContext: vi.fn(() => bitmapContext),
    } as unknown as HTMLCanvasElement;
    const dispatcher = new StaticCanvasRenderDispatcher(canvas, true);
    const bitmap = { width: 10, height: 20, close: vi.fn() } as unknown as ImageBitmap;

    expect(dispatcher.commitBitmap(dispatcher.begin(), bitmap)).toBe(true);
    expect(canvas.getContext).toHaveBeenCalledOnce();
    expect(setWidth).not.toHaveBeenCalled();
    expect(setHeight).not.toHaveBeenCalled();
  });

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
