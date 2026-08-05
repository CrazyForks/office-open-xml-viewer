import { describe, expect, it, vi } from 'vitest';
import type { XlsxWorkbook } from '../workbook.js';
import {
  SelectionController,
  SheetAcquisition,
  SheetRenderDispatcher,
  ViewportState,
} from './sheet-viewer-runtime.js';

function workbook() {
  const destroy = vi.fn();
  return { value: { destroy } as unknown as XlsxWorkbook, destroy };
}

describe('XLSX viewer composition roles', () => {
  it('SheetAcquisition installs only the latest generation and closes every loser', async () => {
    const acquisition = new SheetAcquisition();
    const first = workbook();
    const second = workbook();
    let resolveFirst: ((value: XlsxWorkbook) => void) | undefined;
    const firstLoad = new Promise<XlsxWorkbook>((resolve) => { resolveFirst = resolve; });

    const stale = acquisition.replace(() => firstLoad);
    await acquisition.replace(() => Promise.resolve(second.value));
    resolveFirst?.(first.value);

    await expect(stale).resolves.toBeNull();
    expect(acquisition.current).toBe(second.value);
    expect(first.destroy).toHaveBeenCalledOnce();
    acquisition.destroy();
    expect(second.destroy).toHaveBeenCalledOnce();
  });

  it('ViewportState clamps logical offsets without a native scroll element', () => {
    const viewport = new ViewportState(1);
    viewport.setExtent(1000, 800);
    viewport.setViewportSize(300, 200);
    viewport.setOffset(900, 700);
    expect({ x: viewport.x, y: viewport.y }).toEqual({ x: 700, y: 600 });
    viewport.setScale(1.5);
    expect(viewport.scale).toBe(1.5);
  });

  it('SelectionController owns independent immutable cell coordinates', () => {
    const selection = new SelectionController();
    const cell = { row: 2, col: 3 };
    selection.select(cell);
    cell.row = 9;
    expect(selection.anchor).toEqual({ row: 2, col: 3 });
    selection.reset();
    expect(selection.active).toBeNull();
  });

  it('SelectionController owns range extension and renderer header projection', () => {
    const selection = new SelectionController();
    selection.select({ row: 2, col: 1 }, 'rows');
    selection.extend({ row: 5, col: 1 });
    expect(selection.snapshot()).toEqual({
      anchor: { row: 2, col: 1 },
      active: { row: 5, col: 1 },
      mode: 'rows',
    });
    expect(selection.headerHighlight()).toEqual({
      selectedRowRange: { start: 2, end: 5, strong: true },
      selectedColRange: { start: 1, end: Number.MAX_SAFE_INTEGER, strong: false },
    });
  });

  it('SheetRenderDispatcher invalidates in-flight generations on destroy', () => {
    const dispatcher = new SheetRenderDispatcher();
    const generation = dispatcher.begin();
    expect(dispatcher.isCurrent(generation)).toBe(true);
    dispatcher.destroy();
    expect(dispatcher.isCurrent(generation)).toBe(false);
  });

  it('SheetRenderDispatcher owns stale ImageBitmap disposal', () => {
    const close = vi.fn();
    const dispatcher = new SheetRenderDispatcher();
    const stale = dispatcher.begin();
    dispatcher.begin();
    expect(dispatcher.commitBitmap(
      stale,
      { close, width: 1, height: 1 } as unknown as ImageBitmap,
      1,
      1,
    )).toBe(false);
    expect(close).toHaveBeenCalledOnce();
  });

  it('SheetRenderDispatcher closes an owned bitmap when commit fails', () => {
    const failure = new Error('context lost');
    const transfer = vi.fn(() => { throw failure; });
    const canvas = {
      width: 0,
      height: 0,
      style: { width: '', height: '' },
      getContext: vi.fn(() => ({ transferFromImageBitmap: transfer })),
    } as unknown as HTMLCanvasElement;
    const close = vi.fn();
    const bitmap = { close, width: 2, height: 3 } as unknown as ImageBitmap;
    const dispatcher = new SheetRenderDispatcher(canvas, true);
    const generation = dispatcher.begin();

    expect(() => dispatcher.commitBitmap(generation, bitmap, 20, 30)).toThrow(failure);
    expect(close).toHaveBeenCalledOnce();
  });
});
