import { afterEach, describe, expect, it, vi } from 'vitest';
import { XlsxSheetViewer } from './viewer.js';
import { XlsxWorkbook } from './workbook.js';
import type { OoxmlResourceMetrics } from '@silurus/ooxml-core';
import { installDom, makeContainer, makeEl, type FakeEl } from './viewer-destroy-test-dom.js';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

function descendants(root: FakeEl): FakeEl[] {
  return root.children.flatMap((child) => [child, ...descendants(child)]);
}

describe('XlsxSheetViewer canvas mount', () => {
  it('uses the caller canvas without constructing workbook footer chrome', () => {
    installDom();
    const parent = makeContainer();
    const canvas = makeEl('canvas');
    canvas.clientWidth = 640;
    canvas.clientHeight = 360;
    parent.appendChild(canvas);

    const viewer = new XlsxSheetViewer(canvas as unknown as HTMLCanvasElement);

    expect(viewer.canvasElement).toBe(canvas);
    expect(descendants(parent).filter((element) => element.tag === 'canvas')).toContain(canvas);
    expect(descendants(parent).filter((element) => element.tag === 'button')).toHaveLength(0);
    const viewportInput = descendants(parent).find(
      (element) => element.getAttribute('data-xlsx-viewport-input') === 'sheet',
    );
    expect(viewportInput?.style.overflow).toBe('clip');
    expect(viewportInput?.children).toHaveLength(0);
    expect(descendants(parent).some((element) => element.style.overflow === 'auto')).toBe(false);

    viewer.destroy();
  });

  it('restores the exact caller-owned canvas position, style, and bitmap dimensions', () => {
    installDom();
    const parent = makeContainer();
    const before = makeEl('span');
    const canvas = makeEl('canvas');
    const after = makeEl('span');
    canvas.setAttribute('style', 'width:320px;height:180px;border:1px solid red');
    canvas.style.cssText = 'width:320px;height:180px;border:1px solid red';
    canvas.width = 960;
    canvas.height = 540;
    parent.appendChild(before);
    parent.appendChild(canvas);
    parent.appendChild(after);

    const viewer = new XlsxSheetViewer(canvas as unknown as HTMLCanvasElement);
    expect(parent.children).toEqual([before, parent.children[1], after]);
    expect(parent.children[1]).not.toBe(canvas);

    viewer.destroy();
    viewer.destroy();

    expect(parent.children).toEqual([before, canvas, after]);
    expect(canvas.getAttribute('style')).toBe('width:320px;height:180px;border:1px solid red');
    expect(canvas.width).toBe(960);
    expect(canvas.height).toBe(540);
  });

  it('retains immutable query snapshots and closes every mutation after destroy', async () => {
    installDom();
    const canvas = makeEl('canvas');
    const viewer = new XlsxSheetViewer(canvas as unknown as HTMLCanvasElement, {
      cellScale: 1.25,
      hiddenSheetMode: 'dim',
    });
    viewer.destroy();

    expect(viewer.canvasElement).toBe(canvas);
    expect(viewer.sheetIndex).toBe(0);
    expect(viewer.sheetCount).toBe(0);
    expect(viewer.sheetNames).toEqual([]);
    expect(viewer.getViewportOffset()).toEqual({ x: 0, y: 0 });
    expect(viewer.selection).toBeNull();
    expect(viewer.getScale()).toBe(1.25);
    expect(viewer.hiddenSheetMode).toBe('dim');
    expect(viewer.visibleSheetCount).toBe(0);
    expect(viewer.getCellAt(0, 0)).toBeNull();

    const closed = 'XlsxSheetViewer is destroyed';
    await expect(viewer.load(new ArrayBuffer(0))).rejects.toThrow(closed);
    await expect(viewer.goToSheet(0)).rejects.toThrow(closed);
    await expect(viewer.nextSheet()).rejects.toThrow(closed);
    await expect(viewer.prevSheet()).rejects.toThrow(closed);
    await expect(viewer.setViewportOffset({ x: 0, y: 0 })).rejects.toThrow(closed);
    await expect(viewer.scrollToCell('A1')).rejects.toThrow(closed);
    await expect(viewer.relayout()).rejects.toThrow(closed);
    await expect(viewer.setHiddenSheetMode('show')).rejects.toThrow(closed);
    await expect(viewer.findText('x')).rejects.toThrow(closed);
    await expect(viewer.findNext()).rejects.toThrow(closed);
    await expect(viewer.findPrev()).rejects.toThrow(closed);
    await expect(viewer.getResourceMetrics()).rejects.toThrow(closed);
    expect(() => viewer.setScale(1)).toThrow(closed);
    expect(() => viewer.zoomIn()).toThrow(closed);
    expect(() => viewer.zoomOut()).toThrow(closed);
    expect(() => viewer.fitWidth()).toThrow(closed);
    expect(() => viewer.fitPage()).toThrow(closed);
    expect(() => viewer.select('A1')).toThrow(closed);
    expect(() => viewer.setSelectionColor('#000')).toThrow(closed);
    expect(() => viewer.clearFind()).toThrow(closed);
  });

  it('retains the successful load metrics snapshot after destroy without an explicit metrics query', async () => {
    installDom();
    const metrics: OoxmlResourceMetrics = {
      schemaVersion: 1,
      scope: 'load',
      format: 'xlsx',
      mode: 'main',
      status: 'ok',
      sourceBytes: 12,
      elapsedMs: 3,
      policy: {
        maxArchiveEntryBytes: null,
        maxTotalInflatedBytes: null,
        maxArchiveEntries: null,
      },
      checkpoints: [],
    };
    const workbook = {
      sheetNames: ['Sheet1'],
      tabColors: {} as Record<number, string>,
      destroy: vi.fn(),
      getResourceMetrics: vi.fn().mockResolvedValue(metrics),
    } as unknown as XlsxWorkbook;
    vi.spyOn(XlsxWorkbook, 'load').mockImplementation(async (_source, options) => {
      options?.onResourceMetrics?.(metrics);
      return workbook;
    });

    const canvas = makeEl('canvas');
    const viewer = new XlsxSheetViewer(canvas as unknown as HTMLCanvasElement);
    const engine = (viewer as unknown as { engine: { showSheet(index: number): Promise<void> } }).engine;
    vi.spyOn(engine, 'showSheet').mockResolvedValue(undefined);

    await viewer.load(new ArrayBuffer(0));
    viewer.destroy();

    await expect(viewer.getResourceMetrics()).resolves.toEqual(metrics);
    expect(workbook.getResourceMetrics).not.toHaveBeenCalled();
  });

  it('retains terminal error metrics after a rejected load and destroy', async () => {
    installDom();
    const metrics: OoxmlResourceMetrics = {
      schemaVersion: 1,
      scope: 'load',
      format: 'xlsx',
      mode: 'main',
      status: 'error',
      sourceBytes: 12,
      elapsedMs: 3,
      policy: {
        maxArchiveEntryBytes: null,
        maxTotalInflatedBytes: null,
        maxArchiveEntries: null,
      },
      checkpoints: [],
      error: { code: 'ooxml-resource-limit', stage: 'package-open' },
    };
    const failure = new Error('load failed');
    vi.spyOn(XlsxWorkbook, 'load').mockImplementation(async (_source, options) => {
      options?.onResourceMetrics?.(metrics);
      throw failure;
    });

    const viewer = new XlsxSheetViewer(makeEl('canvas') as unknown as HTMLCanvasElement);
    await expect(viewer.load(new ArrayBuffer(0))).rejects.toBe(failure);
    viewer.destroy();

    await expect(viewer.getResourceMetrics()).resolves.toEqual(metrics);
  });
});
