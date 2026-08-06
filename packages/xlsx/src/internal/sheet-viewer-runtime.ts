import type { XlsxWorkbook } from '../workbook.js';
import { StaticCanvasRenderDispatcher } from '@silurus/ooxml-core/internal/canvas-viewer-mechanics';

/** Generation-safe workbook ownership for one viewer instance. */
export class SheetAcquisition {
  private generation = 0;
  private currentWorkbook: XlsxWorkbook | null = null;
  private closed = false;

  get current(): XlsxWorkbook | null {
    return this.currentWorkbook;
  }

  async replace(load: () => Promise<XlsxWorkbook>): Promise<XlsxWorkbook | null> {
    this.assertOpen();
    const generation = ++this.generation;
    let candidate: XlsxWorkbook;
    try {
      candidate = await load();
    } catch (error) {
      if (this.closed) throw this.closedError();
      if (generation !== this.generation) return null;
      throw error;
    }
    if (this.closed) {
      candidate.destroy();
      throw this.closedError();
    }
    if (generation !== this.generation) {
      candidate.destroy();
      return null;
    }
    this.install(candidate);
    return candidate;
  }

  /** Commit an already acquired workbook, closing the previously owned one. */
  install(candidate: XlsxWorkbook): void {
    this.assertOpen();
    const previous = this.currentWorkbook;
    this.currentWorkbook = candidate;
    previous?.destroy();
  }

  destroy(): void {
    if (this.closed) return;
    this.closed = true;
    this.generation++;
    this.currentWorkbook?.destroy();
    this.currentWorkbook = null;
  }

  private assertOpen(): void {
    if (this.closed) throw this.closedError();
  }

  private closedError(): Error {
    return new Error('SheetAcquisition is closed');
  }
}

/** Logical viewport independent of native browser scrolling. */
export class ViewportState {
  private offsetX = 0;
  private offsetY = 0;
  private contentWidth = 0;
  private contentHeight = 0;
  private viewportWidth = 0;
  private viewportHeight = 0;

  constructor(private scaleValue: number) {}

  get x(): number { return this.offsetX; }
  get y(): number { return this.offsetY; }
  get scale(): number { return this.scaleValue; }
  get maxX(): number { return Math.max(0, this.contentWidth - this.viewportWidth); }
  get maxY(): number { return Math.max(0, this.contentHeight - this.viewportHeight); }

  setScale(scale: number): void {
    this.scaleValue = scale;
  }

  setExtent(contentWidth: number, contentHeight: number): void {
    this.contentWidth = Math.max(0, contentWidth);
    this.contentHeight = Math.max(0, contentHeight);
    this.setOffset(this.offsetX, this.offsetY);
  }

  /** Expand to include an extent reported by a native scroll host without
   * shrinking the format-computed worksheet extent. Browser layout and test
   * doubles may publish scrollWidth/scrollHeight after the worksheet geometry
   * has already been installed. */
  ensureExtent(contentWidth: number, contentHeight: number): void {
    this.contentWidth = Math.max(this.contentWidth, Math.max(0, contentWidth));
    this.contentHeight = Math.max(this.contentHeight, Math.max(0, contentHeight));
  }

  setViewportSize(width: number, height: number): void {
    this.viewportWidth = Math.max(0, width);
    this.viewportHeight = Math.max(0, height);
    this.setOffset(this.offsetX, this.offsetY);
  }

  setOffset(x: number, y: number): void {
    this.offsetX = Math.min(this.maxX, Math.max(0, x));
    this.offsetY = Math.min(this.maxY, Math.max(0, y));
  }

  /** Mirror offsets already clamped by a native browser scroll container. This
   * intentionally does not re-clamp against logical extents: the DOM is the
   * authority for the composite viewer and may publish its geometry later. */
  adoptNativeOffset(x: number, y: number): void {
    this.offsetX = Math.max(0, x);
    this.offsetY = Math.max(0, y);
  }

  reset(): void {
    this.offsetX = 0;
    this.offsetY = 0;
  }
}

/** XLSX render scheduling around core-owned static bitmap lifecycle mechanics. */
export class SheetRenderDispatcher {
  private animationFrame: number | null = null;
  private readonly staticDispatcher: StaticCanvasRenderDispatcher | null;
  private generation = 0;
  private destroyed = false;

  constructor(
    canvas?: HTMLCanvasElement,
    workerBitmapMode = false,
  ) {
    this.staticDispatcher = canvas
      ? new StaticCanvasRenderDispatcher(canvas, workerBitmapMode)
      : null;
  }

  begin(): number {
    if (this.staticDispatcher) return this.staticDispatcher.begin();
    return ++this.generation;
  }

  isCurrent(generation: number): boolean {
    if (this.staticDispatcher) return this.staticDispatcher.isCurrent(generation);
    return !this.destroyed && generation === this.generation;
  }

  /** Delegate stale disposal and atomic bitmap replacement to the core owner. */
  commitBitmap(
    generation: number,
    bitmap: ImageBitmap,
    cssWidth: number,
    cssHeight: number,
  ): boolean {
    if (!this.isCurrent(generation)) {
      bitmap.close();
      return false;
    }
    if (!this.staticDispatcher) {
      bitmap.close();
      throw new Error('SheetRenderDispatcher is not configured for worker bitmap rendering');
    }
    return this.staticDispatcher.commitBitmap(generation, bitmap, {
      cssWidth,
      cssHeight,
    });
  }

  schedule(render: () => void): void {
    if (this.animationFrame !== null) return;
    if (typeof requestAnimationFrame !== 'function') {
      render();
      return;
    }
    this.animationFrame = requestAnimationFrame(() => {
      this.animationFrame = null;
      render();
    });
  }

  destroy(): void {
    if (this.destroyed) return;
    this.destroyed = true;
    if (this.animationFrame !== null && typeof cancelAnimationFrame === 'function') {
      cancelAnimationFrame(this.animationFrame);
      this.animationFrame = null;
    }
    this.staticDispatcher?.destroy();
    this.generation++;
  }
}

export interface SheetCellAddress {
  readonly row: number;
  readonly col: number;
}

export type SheetSelectionMode = 'cells' | 'rows' | 'cols' | 'all';

/** Selection state and immutable snapshots for one sheet viewer instance. */
export class SelectionController {
  private anchorCell: SheetCellAddress | null = null;
  private activeCell: SheetCellAddress | null = null;
  private selectionMode: SheetSelectionMode = 'cells';
  private dragActive = false;

  get anchor(): SheetCellAddress | null {
    return this.anchorCell ? { ...this.anchorCell } : null;
  }

  get active(): SheetCellAddress | null {
    return this.activeCell ? { ...this.activeCell } : null;
  }

  get mode(): SheetSelectionMode { return this.selectionMode; }
  get dragging(): boolean { return this.dragActive; }

  setAnchor(cell: SheetCellAddress | null): void {
    this.anchorCell = cell ? { ...cell } : null;
  }

  setActive(cell: SheetCellAddress | null): void {
    this.activeCell = cell ? { ...cell } : null;
  }

  setMode(mode: SheetSelectionMode): void { this.selectionMode = mode; }
  setDragging(dragging: boolean): void { this.dragActive = dragging; }

  reset(): void {
    this.anchorCell = null;
    this.activeCell = null;
    this.selectionMode = 'cells';
    this.dragActive = false;
  }

  select(cell: SheetCellAddress, mode: SheetSelectionMode = 'cells'): void {
    this.anchorCell = { ...cell };
    this.activeCell = { ...cell };
    this.selectionMode = mode;
  }

  extend(cell: SheetCellAddress): void {
    if (!this.anchorCell) this.anchorCell = { ...cell };
    this.activeCell = { ...cell };
  }

  snapshot(): Readonly<{
    anchor: SheetCellAddress;
    active: SheetCellAddress;
    mode: SheetSelectionMode;
  }> | null {
    if (!this.anchorCell || !this.activeCell) return null;
    return {
      anchor: { ...this.anchorCell },
      active: { ...this.activeCell },
      mode: this.selectionMode,
    };
  }

  headerHighlight(): {
    selectedRowRange: { start: number; end: number; strong: boolean } | null;
    selectedColRange: { start: number; end: number; strong: boolean } | null;
  } {
    if (!this.anchorCell || !this.activeCell) {
      return { selectedRowRange: null, selectedColRange: null };
    }
    const r1 = Math.min(this.anchorCell.row, this.activeCell.row);
    const r2 = Math.max(this.anchorCell.row, this.activeCell.row);
    const c1 = Math.min(this.anchorCell.col, this.activeCell.col);
    const c2 = Math.max(this.anchorCell.col, this.activeCell.col);
    const all = Number.MAX_SAFE_INTEGER;
    switch (this.selectionMode) {
      case 'cells':
        return {
          selectedRowRange: { start: r1, end: r2, strong: false },
          selectedColRange: { start: c1, end: c2, strong: false },
        };
      case 'rows':
        return {
          selectedRowRange: { start: r1, end: r2, strong: true },
          selectedColRange: { start: 1, end: all, strong: false },
        };
      case 'cols':
        return {
          selectedRowRange: { start: 1, end: all, strong: false },
          selectedColRange: { start: c1, end: c2, strong: true },
        };
      case 'all':
        return {
          selectedRowRange: { start: 1, end: all, strong: true },
          selectedColRange: { start: 1, end: all, strong: true },
        };
    }
  }
}
