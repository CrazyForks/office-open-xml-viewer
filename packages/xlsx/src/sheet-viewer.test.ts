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
  it('selects an A1 range as geometry without inventing ActiveCell direction', () => {
    installDom();
    const parent = makeContainer();
    const canvas = makeEl('canvas');
    parent.appendChild(canvas);
    const viewer = new XlsxSheetViewer(canvas as unknown as HTMLCanvasElement);

    viewer.setSelection('D5:B2');

    expect(viewer.selectionState).toEqual({
      areas: [{ kind: 'cells', top: 2, left: 2, bottom: 5, right: 4 }],
      activeAreaIndex: 0,
      activeCell: { row: 2, col: 2 },
      extensionAnchor: { row: 2, col: 2 },
    });
    viewer.destroy();
  });

  it('keeps bounded cell rectangles distinct from row and column selections', () => {
    installDom();
    const parent = makeContainer();
    const canvas = makeEl('canvas');
    parent.appendChild(canvas);
    const viewer = new XlsxSheetViewer(canvas as unknown as HTMLCanvasElement);

    viewer.setSelection('A2:XFD4');
    expect(viewer.selectionState?.areas[0]).toEqual({
      kind: 'cells', top: 2, left: 1, bottom: 4, right: 16_384,
    });
    viewer.setSelection('2:4');
    expect(viewer.selectionState?.areas[0]).toEqual({ kind: 'rows', firstRow: 2, lastRow: 4 });
    viewer.setSelection('B:D');
    expect(viewer.selectionState?.areas[0]).toEqual({
      kind: 'columns', firstColumn: 2, lastColumn: 4,
    });

    viewer.destroy();
  });

  it('represents ActiveCell, extension anchor, and multiple areas independently', () => {
    installDom();
    const parent = makeContainer();
    const canvas = makeEl('canvas');
    parent.appendChild(canvas);
    const viewer = new XlsxSheetViewer(canvas as unknown as HTMLCanvasElement);

    viewer.setSelection({
      areas: [
        { kind: 'cells', top: 1, left: 1, bottom: 4, right: 4 },
        { kind: 'cells', top: 8, left: 2, bottom: 9, right: 3 },
      ],
      activeAreaIndex: 0,
      activeCell: { row: 2, col: 2 },
      extensionAnchor: { row: 4, col: 4 },
    });

    expect(viewer.selectionState?.activeCell).toEqual({ row: 2, col: 2 });
    expect(viewer.selectionState?.extensionAnchor).toEqual({ row: 4, col: 4 });
    expect(viewer.selectionState?.areas).toHaveLength(2);
    viewer.destroy();
  });

  it('emits canonical state only for semantic changes and validates invariants', () => {
    installDom();
    const parent = makeContainer();
    const canvas = makeEl('canvas');
    parent.appendChild(canvas);
    const onSelectionStateChange = vi.fn();
    const viewer = new XlsxSheetViewer(canvas as unknown as HTMLCanvasElement, {
      onSelectionStateChange,
    });

    viewer.setSelection('B2:D5');
    viewer.setSelection('D5:B2');
    expect(onSelectionStateChange).toHaveBeenCalledOnce();
    expect(() => viewer.setSelection({
      areas: [{ kind: 'cells', top: 1, left: 1, bottom: 2, right: 2 }],
      activeAreaIndex: 0,
      activeCell: { row: 3, col: 1 },
      extensionAnchor: { row: 1, col: 1 },
    })).toThrow(/activeCell/);
    viewer.destroy();
  });

  it('serializes reentrant selection callbacks in semantic order', () => {
    installDom();
    const parent = makeContainer();
    const canvas = makeEl('canvas');
    parent.appendChild(canvas);
    const activeCells: Array<{ row: number; col: number }> = [];
    let viewer: XlsxSheetViewer;
    viewer = new XlsxSheetViewer(canvas as unknown as HTMLCanvasElement, {
      onSelectionStateChange(selection) {
        if (!selection) return;
        activeCells.push(selection.activeCell);
        if (selection.activeCell.row === 2) viewer.setSelection('C3');
      },
    });

    viewer.setSelection('B2');

    expect(activeCells).toEqual([{ row: 2, col: 2 }, { row: 3, col: 3 }]);
    expect(viewer.selectionState?.activeCell).toEqual({ row: 3, col: 3 });
    viewer.destroy();
  });

  it('projects a logical selection into frozen-pane fragments without fake pane edges', () => {
    installDom();
    const parent = makeContainer();
    const canvas = makeEl('canvas');
    parent.appendChild(canvas);
    const viewer = new XlsxSheetViewer(canvas as unknown as HTMLCanvasElement);
    const engine = (viewer as unknown as { engine: {
      currentWorksheet: Worksheet;
      canvasArea: FakeEl;
      overlayHost: { selection: FakeEl };
    } }).engine;
    engine.currentWorksheet = { ...worksheet('Frozen'), freezeRows: 1, freezeCols: 1 };
    engine.canvasArea.clientWidth = 640;
    engine.canvasArea.clientHeight = 360;

    viewer.setSelection('A1:C3');

    const fragments = descendants(engine.overlayHost.selection).filter(
      (element) => element.getAttribute('data-xlsx-selection-fragment') === 'cells',
    );
    expect(fragments).toHaveLength(4);
    expect(fragments[0].style['border-right']).toBe('none');
    expect(fragments[0].style['border-bottom']).toBe('none');
    expect(fragments.every((fragment) => fragment.style['border-right'] === 'none')).toBe(true);
    expect(fragments.some((fragment) => fragment.style['border-bottom'].includes('solid'))).toBe(true);
    viewer.destroy();
  });

  it('bounds overlay work and geometry when legal freeze counts cover the sheet', () => {
    installDom();
    const parent = makeContainer();
    const canvas = makeEl('canvas');
    parent.appendChild(canvas);
    const viewer = new XlsxSheetViewer(canvas as unknown as HTMLCanvasElement);
    const engine = (viewer as unknown as { engine: {
      currentWorksheet: Worksheet;
      canvasArea: FakeEl;
      overlayHost: { selection: FakeEl };
    } }).engine;
    engine.currentWorksheet = {
      ...worksheet('Fully frozen'),
      freezeRows: 1_048_576,
      freezeCols: 16_384,
      rightToLeft: true,
    };
    engine.canvasArea.clientWidth = 640;
    engine.canvasArea.clientHeight = 360;

    viewer.setSelection('A1:XFD1048576');

    const fragments = descendants(engine.overlayHost.selection).filter(
      (element) => element.getAttribute('data-xlsx-selection-fragment') === 'cells',
    );
    expect(fragments).toHaveLength(1);
    expect(Number.parseFloat(fragments[0].style.width)).toBeLessThanOrEqual(640);
    expect(Number.parseFloat(fragments[0].style.height)).toBeLessThanOrEqual(360);
    viewer.destroy();
  });

  it('does not materialize a programmatic clipboard range above the cell limit', async () => {
    const document = installDom();
    const parent = makeContainer();
    const canvas = makeEl('canvas');
    parent.appendChild(canvas);
    const viewer = new XlsxSheetViewer(canvas as unknown as HTMLCanvasElement);
    const writeText = vi.fn(() => Promise.resolve());
    Object.assign(document.defaultView, { navigator: { clipboard: { writeText } } });
    const engine = (viewer as unknown as {
      engine: {
        currentWorksheet: Worksheet;
        copySelection(): Promise<unknown>;
      };
    }).engine;
    engine.currentWorksheet = {
      ...worksheet('Sparse'),
      rows: [
        { index: 2, cells: [{ row: 2, col: 2, value: { type: 'text', text: 'first' } }] },
        {
          index: 1_048_575,
          cells: [{ row: 1_048_575, col: 16_383, value: { type: 'text', text: 'last' } }],
        },
      ],
    } as unknown as Worksheet;

    viewer.select('B2:XFC1048575');
    const result = await engine.copySelection();

    expect(writeText).not.toHaveBeenCalled();
    expect(result).toEqual({ status: 'too-large', limit: 'cells' });
    viewer.destroy();
  });

  it('applies the clipboard cell limit to pointer-created selections too', async () => {
    const document = installDom();
    const parent = makeContainer();
    const canvas = makeEl('canvas');
    parent.appendChild(canvas);
    const viewer = new XlsxSheetViewer(canvas as unknown as HTMLCanvasElement);
    const writeText = vi.fn(() => Promise.resolve());
    Object.assign(document.defaultView, { navigator: { clipboard: { writeText } } });
    const engine = (viewer as unknown as {
      engine: {
        currentWorksheet: Worksheet;
        selectionController: {
          select(cell: { row: number; col: number }): void;
          extend(cell: { row: number; col: number }): void;
        };
        copySelection(): Promise<unknown>;
      };
    }).engine;
    engine.currentWorksheet = worksheet('Large selection');
    engine.selectionController.select({ row: 1, col: 1 });
    engine.selectionController.extend({ row: 501, col: 500 });

    const result = await engine.copySelection();

    expect(writeText).not.toHaveBeenCalled();
    expect(result).toEqual({ status: 'too-large', limit: 'cells' });
    viewer.destroy();
  });

  it('reports unsupported multi-area copy without flattening selection semantics', async () => {
    installDom();
    const parent = makeContainer();
    const canvas = makeEl('canvas');
    parent.appendChild(canvas);
    const viewer = new XlsxSheetViewer(canvas as unknown as HTMLCanvasElement);
    const engine = (viewer as unknown as { engine: { currentWorksheet: Worksheet } }).engine;
    engine.currentWorksheet = worksheet('Areas');
    viewer.setSelection({
      areas: [
        { kind: 'cells', top: 1, left: 1, bottom: 1, right: 1 },
        { kind: 'cells', top: 3, left: 3, bottom: 3, right: 3 },
      ],
      activeAreaIndex: 0,
      activeCell: { row: 1, col: 1 },
      extensionAnchor: { row: 1, col: 1 },
    });

    await expect(viewer.copySelection()).resolves.toEqual({ status: 'unsupported-multiple-areas' });
    viewer.destroy();
  });

  it('extracts bounded, serializable selection context without workbook mutation APIs', () => {
    installDom();
    const parent = makeContainer();
    const canvas = makeEl('canvas');
    parent.appendChild(canvas);
    const viewer = new XlsxSheetViewer(canvas as unknown as HTMLCanvasElement);
    const engine = (viewer as unknown as { engine: { currentWorksheet: Worksheet } }).engine;
    engine.currentWorksheet = {
      ...worksheet('Analysis'),
      rows: [
        { index: 1, cells: [
          { row: 1, col: 1, value: { type: 'text', text: 'Revenue' } },
          { row: 1, col: 2, value: { type: 'number', number: 42 }, formula: 'SUM(B2:B4)' },
        ] },
        { index: 2, cells: [
          { row: 2, col: 1, value: { type: 'bool', bool: true } },
        ] },
      ],
    } as unknown as Worksheet;
    viewer.setSelection('A1:B2');

    const context = viewer.getSelectionContext({ maxCells: 2 });

    expect(context).toMatchObject({
      format: 'xlsx',
      sheetName: 'Analysis',
      coordinateCountUpperBound: 4,
      truncated: true,
      maxCells: 2,
    });
    expect(context?.cells).toEqual([
      { address: { row: 1, col: 1 }, displayText: '', valueType: 'text', value: 'Revenue' },
      {
        address: { row: 1, col: 2 }, displayText: '', valueType: 'number', value: 42,
        formula: 'SUM(B2:B4)',
      },
    ]);
    viewer.destroy();
  });

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

  it('uses the caller canvas with native scrollbars and without workbook footer chrome', () => {
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
    expect(viewportInput?.style.overflow).toBe('auto');
    expect(viewportInput?.children).toHaveLength(1);
    expect(descendants(parent).some((element) => element.style.overflow === 'auto')).toBe(true);

    viewer.destroy();
  });

  it('can hide native sheet scrollbars without adding workbook footer chrome', () => {
    installDom();
    const parent = makeContainer();
    const canvas = makeEl('canvas');
    parent.appendChild(canvas);

    const viewer = new XlsxSheetViewer(canvas as unknown as HTMLCanvasElement, {
      showScrollbars: false,
    });

    const viewportInput = descendants(parent).find(
      (element) => element.getAttribute('data-xlsx-viewport-input') === 'sheet',
    );
    expect(viewportInput?.style.overflow).toBe('clip');
    expect(viewportInput?.children).toHaveLength(0);
    expect(descendants(parent).filter((element) => element.tag === 'button')).toHaveLength(0);

    viewer.destroy();
  });

  it('continues drag selection beyond the visible viewport while auto-scrolling', () => {
    const doc = installDom();
    const parent = makeContainer();
    const canvas = makeEl('canvas');
    parent.appendChild(canvas);
    const viewer = new XlsxSheetViewer(canvas as unknown as HTMLCanvasElement);
    const engine = (viewer as unknown as { engine: {
      currentWorksheet: Worksheet;
      canvasArea: FakeEl;
      scrollHost: FakeEl;
      viewport: {
        setExtent(width: number, height: number): void;
        setViewportSize(width: number, height: number): void;
      };
      getCellAt(clientX: number, clientY: number): { row: number; col: number } | null;
      scheduleRender(): void;
    } }).engine;
    engine.currentWorksheet = {
      ...worksheet('Long sheet'),
      defaultColWidth: 8.43,
    };
    engine.canvasArea.clientWidth = 800;
    engine.canvasArea.clientHeight = 600;
    // Simulate classic scrollbars that reserve a 20px gutter inside the
    // canvasArea (unlike overlay scrollbars on macOS).
    engine.scrollHost.clientWidth = 780;
    engine.scrollHost.clientHeight = 580;
    engine.scrollHost.scrollWidth = 5_000;
    engine.scrollHost.scrollHeight = 5_000;
    engine.viewport.setExtent(5_000, 5_000);
    engine.viewport.setViewportSize(800, 600);
    engine.scheduleRender = () => undefined;

    const frames: FrameRequestCallback[] = [];
    Object.assign(doc.defaultView, {
      requestAnimationFrame: (callback: FrameRequestCallback) => {
        frames.push(callback);
        return frames.length;
      },
      cancelAnimationFrame: () => undefined,
    });
    const pointer = (overrides: Record<string, unknown>) => ({
      button: 0,
      pointerId: 1,
      pointerType: 'mouse',
      clientX: 100,
      clientY: 100,
      shiftKey: false,
      preventDefault: () => undefined,
      ...overrides,
    });

    engine.scrollHost.dispatch('pointerdown', pointer({
      pointerId: 2,
      pointerType: 'touch',
      clientX: 400,
      clientY: 300,
    }));
    engine.scrollHost.dispatch('pointerdown', pointer({}));
    const primarySelection = viewer.selection;
    engine.scrollHost.dispatch('pointerup', pointer({
      pointerId: 2,
      pointerType: 'touch',
      clientX: 400,
      clientY: 300,
    }));
    expect(viewer.selection).toEqual(primarySelection);

    engine.scrollHost.dispatch('pointerdown', pointer({
      pointerId: 2,
      pointerType: 'touch',
      clientX: 400,
      clientY: 300,
    }));
    engine.scrollHost.dispatch('pointerup', pointer({
      pointerId: 2,
      pointerType: 'touch',
      clientX: 400,
      clientY: 300,
    }));
    engine.scrollHost.dispatch('pointermove', pointer({
      pointerId: 2,
      clientX: 1_600,
      clientY: 1_200,
    }));
    engine.scrollHost.dispatch('pointerup', pointer({ pointerId: 2 }));
    engine.scrollHost.dispatch('pointercancel', pointer({ pointerId: 2 }));
    expect(frames).toHaveLength(0);
    expect(viewer.selection).toEqual(primarySelection);

    engine.scrollHost.dispatch('pointermove', pointer({ clientX: 1_600, clientY: 1_200 }));
    expect(viewer.selection?.active).toEqual(engine.getCellAt(779, 579));
    const visibleEdgeSelection = viewer.selection;
    engine.scrollHost.dispatch('pointerup', pointer({ pointerId: 2 }));
    engine.scrollHost.dispatch('pointercancel', pointer({ pointerId: 2 }));
    for (let frame = 1; frame <= 20; frame += 1) {
      const callback = frames.shift();
      expect(callback).toBeDefined();
      callback?.(frame * 16);
    }

    expect(viewer.getViewportOffset().x).toBeGreaterThan(0);
    expect(viewer.getViewportOffset().y).toBeGreaterThan(0);
    expect(viewer.selection?.active.row).toBeGreaterThan(visibleEdgeSelection?.active.row ?? 0);
    expect(viewer.selection?.active.col).toBeGreaterThan(visibleEdgeSelection?.active.col ?? 0);

    engine.scrollHost.dispatch('pointerup', pointer({ clientX: 1_600, clientY: 1_200 }));
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

    const viewer = XlsxSheetViewer.fromWorkbook(
      canvas as unknown as HTMLCanvasElement,
      workbook,
    );
    await viewer.goToSheet(1);

    expect(viewer.sheetIndex).toBe(1);
    expect(getWorksheet).toHaveBeenCalledWith(1);
    expect(getWorksheet).not.toHaveBeenCalledWith(0);
    await expect((viewer as XlsxSheetViewer).load(new ArrayBuffer(0))).rejects.toThrow(
      'XlsxSheetViewer.load() is unsupported on a Viewer created by fromWorkbook()',
    );

    viewer.destroy();
    expect(destroy).not.toHaveBeenCalled();
    expect(parent.children).toEqual([canvas]);
  });

  it('validates a borrowed mode conflict before mounting the caller canvas', () => {
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

    expect(() => XlsxSheetViewer.fromWorkbook(
      canvas as unknown as HTMLCanvasElement,
      workbook,
      { mode: 'main' } as never,
    )).toThrow("opts.mode='main' conflicts with the borrowed engine's mode='worker'");
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
    const first = XlsxSheetViewer.fromWorkbook(
      makeEl('canvas') as unknown as HTMLCanvasElement,
      workbook,
    );
    const second = XlsxSheetViewer.fromWorkbook(
      makeEl('canvas') as unknown as HTMLCanvasElement,
      workbook,
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
