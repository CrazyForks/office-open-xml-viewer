import { afterEach, describe, expect, it, vi } from 'vitest';
import { XlsxSheetViewer } from './viewer.js';
import { XlsxWorkbook } from './workbook.js';
import type { OoxmlResourceMetrics } from '@silurus/ooxml-core';
import type { Worksheet } from './types.js';
import {
  installDom,
  makeContainer,
  makeDocument,
  makeEl,
  type FakeEl,
} from './viewer-destroy-test-dom.js';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

function descendants(root: FakeEl): FakeEl[] {
  return root.children.flatMap((child) => [child, ...descendants(child)]);
}

function worksheet(name: string): Worksheet {
  return {
    name,
    rows: [],
    colWidths: {},
    rowHeights: {},
    defaultColWidth: 64,
    defaultRowHeight: 20,
    mergeCells: [],
    freezeRows: 0,
    freezeCols: 0,
    conditionalFormats: [],
    charts: [],
    images: [],
    shapeGroups: [],
  } as unknown as Worksheet;
}

describe('XlsxSheetViewer canvas mount', () => {
  it('creates chrome, styles, and document listeners in the canvas owner window', () => {
    const openerDocument = installDom();
    const popupDocument = makeDocument(2);
    const parent = makeContainer(800, 600, popupDocument);
    const canvas = makeEl('canvas', popupDocument);
    parent.appendChild(canvas);

    const viewer = new XlsxSheetViewer(canvas as unknown as HTMLCanvasElement);
    const mounted = descendants(parent);

    expect(mounted.every((element) => element.ownerDocument === popupDocument)).toBe(true);
    expect(popupDocument.head.querySelector('style[data-xlsx-viewer-styles]')).not.toBeNull();
    expect(openerDocument.head.querySelector('style[data-xlsx-viewer-styles]')).toBeNull();
    expect(popupDocument.listenerCount('keydown')).toBe(1);
    expect(openerDocument.listenerCount('keydown')).toBe(0);

    const popupWrite = vi.fn(() => Promise.resolve());
    const openerWrite = vi.fn(() => Promise.resolve());
    Object.assign(popupDocument.defaultView, { navigator: { clipboard: { writeText: popupWrite } } });
    Object.assign(openerDocument.defaultView, { navigator: { clipboard: { writeText: openerWrite } } });
    const engine = (viewer as unknown as { engine: {
      currentWorksheet: Worksheet;
      selectionController: { select(cell: { row: number; col: number }): void };
      copySelection(): void;
    } }).engine;
    engine.currentWorksheet = {
      ...worksheet('Popup'),
      rows: [{
        index: 1,
        cells: [{ row: 1, col: 1, value: { type: 'text', text: 'popup' } }],
      }],
    } as unknown as Worksheet;
    engine.selectionController.select({ row: 1, col: 1 });
    engine.copySelection();
    expect(popupWrite).toHaveBeenCalledWith('popup');
    expect(openerWrite).not.toHaveBeenCalled();

    viewer.destroy();
    expect(popupDocument.listenerCount('keydown')).toBe(0);
  });

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

  it('borrows one loaded workbook, opens the requested sheet, and leaves workbook cleanup to the caller', async () => {
    installDom();
    const destroy = vi.fn();
    const getWorksheet = vi.fn((index: number) => Promise.resolve(worksheet(`Sheet${index + 1}`)));
    const workbook = {
      mode: 'main',
      sheetCount: 2,
      sheetNames: ['Sheet1', 'Sheet2'],
      tabColors: {} as Record<number, string>,
      getWorksheet,
      isHidden: () => false,
      destroy,
    } as unknown as XlsxWorkbook;
    const parent = makeContainer();
    const canvas = makeEl('canvas');
    parent.appendChild(canvas);

    const viewer = new XlsxSheetViewer(canvas as unknown as HTMLCanvasElement, { workbook });
    await viewer.goToSheet(1);

    expect(viewer.sheetIndex).toBe(1);
    expect(getWorksheet).toHaveBeenCalledWith(1);
    expect(getWorksheet).not.toHaveBeenCalledWith(0);
    await expect(viewer.load(new ArrayBuffer(0))).rejects.toThrow(
      'XlsxSheetViewer.load() is unsupported when an engine is injected',
    );

    viewer.destroy();
    expect(destroy).not.toHaveBeenCalled();
    expect(parent.children).toEqual([canvas]);
  });

  it('validates an injected mode conflict before mounting the caller canvas', () => {
    installDom();
    const parent = makeContainer();
    const canvas = makeEl('canvas');
    parent.appendChild(canvas);
    const workbook = {
      mode: 'worker',
      sheetCount: 1,
      sheetNames: ['Sheet1'],
      destroy: vi.fn(),
    } as unknown as XlsxWorkbook;

    expect(() => new XlsxSheetViewer(canvas as unknown as HTMLCanvasElement, {
      workbook,
      mode: 'main',
    })).toThrow("opts.mode='main' conflicts with the injected engine's mode='worker'");
    expect(parent.children).toEqual([canvas]);
  });

  it('keeps mutable view state independent when two viewers borrow the same cached sheet', async () => {
    installDom();
    const source = worksheet('Shared');
    const workbook = {
      mode: 'main',
      sheetCount: 1,
      sheetNames: ['Shared'],
      tabColors: {} as Record<number, string>,
      getWorksheet: vi.fn().mockResolvedValue(source),
      isHidden: () => false,
      destroy: vi.fn(),
    } as unknown as XlsxWorkbook;
    const first = new XlsxSheetViewer(
      makeEl('canvas') as unknown as HTMLCanvasElement,
      { workbook },
    );
    const second = new XlsxSheetViewer(
      makeEl('canvas') as unknown as HTMLCanvasElement,
      { workbook },
    );
    await Promise.all([first.goToSheet(0), second.goToSheet(0)]);
    const firstWorksheet = (first as unknown as {
      engine: { currentWorksheet: Worksheet };
    }).engine.currentWorksheet;
    const secondWorksheet = (second as unknown as {
      engine: { currentWorksheet: Worksheet };
    }).engine.currentWorksheet;

    firstWorksheet.rowHeights[1] = 40;
    firstWorksheet.colWidths[1] = 16;

    expect(firstWorksheet).not.toBe(secondWorksheet);
    expect(secondWorksheet.rowHeights[1]).toBeUndefined();
    expect(secondWorksheet.colWidths[1]).toBeUndefined();
    expect(source.rowHeights[1]).toBeUndefined();
    expect(source.colWidths[1]).toBeUndefined();

    first.destroy();
    second.destroy();
  });
});
