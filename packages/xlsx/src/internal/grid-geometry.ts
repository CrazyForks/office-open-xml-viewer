import type { ViewportRange, Worksheet } from '../types.js';
import { colWidthToPx, rowHeightToPx } from '../renderer.js';

export const MAX_WORKSHEET_ROW = 1_048_576;
export const MAX_WORKSHEET_COL = 16_384;

/** Sparse cumulative axis geometry. Offsets are logical, unscaled pixels. */
export class GridAxisGeometry {
  private readonly indices: number[];
  private readonly cumulativeDelta: number[];
  private readonly customPx: number[];

  constructor(
    customs: Record<number, number>,
    private readonly defaultPx: number,
    toPx: (raw: number) => number,
    private readonly maxIndex: number,
  ) {
    this.indices = Object.keys(customs)
      .map(Number)
      .filter((value) => value >= 1 && value <= maxIndex)
      .sort((a, b) => a - b);
    this.cumulativeDelta = new Array(this.indices.length);
    this.customPx = new Array(this.indices.length);
    let accumulated = 0;
    for (let index = 0; index < this.indices.length; index++) {
      const px = toPx(customs[this.indices[index]]);
      this.customPx[index] = px;
      accumulated += px - defaultPx;
      this.cumulativeDelta[index] = accumulated;
    }
  }

  private deltaBefore(index: number): number {
    let low = 0;
    let high = this.indices.length;
    while (low < high) {
      const middle = (low + high) >> 1;
      if (this.indices[middle] < index) low = middle + 1;
      else high = middle;
    }
    return low === 0 ? 0 : this.cumulativeDelta[low - 1];
  }

  offsetOf(index: number): number {
    return (index - 1) * this.defaultPx + this.deltaBefore(index);
  }

  indexAt(offset: number): { index: number; partial: number } {
    if (offset <= 0) return { index: 1, partial: 0 };
    let low = 1;
    let high = this.maxIndex;
    while (low < high) {
      const middle = (low + high + 1) >> 1;
      if (this.offsetOf(middle) <= offset) low = middle;
      else high = middle - 1;
    }
    return { index: low, partial: offset - this.offsetOf(low) };
  }

  scrollableIndexAt(content: number, firstScrollable: number): number | null {
    const absoluteOffset = content + this.offsetOf(firstScrollable);
    if (absoluteOffset >= this.offsetOf(this.maxIndex) + this.sizeOf(this.maxIndex)) {
      return null;
    }
    return this.indexAt(absoluteOffset).index;
  }

  sizeOf(index: number): number {
    return this.offsetOf(index + 1) - this.offsetOf(index);
  }

  /** Same sparse axis after Excel-style per-band scale rounding. */
  scaled(scale: number): GridAxisGeometry {
    const customs: Record<number, number> = {};
    for (let index = 0; index < this.indices.length; index++) {
      customs[this.indices[index]] = Math.round(this.customPx[index] * scale);
    }
    return new GridAxisGeometry(
      customs,
      Math.round(this.defaultPx * scale),
      (value) => value,
      this.maxIndex,
    );
  }

  /** Number of bands required to cover `distance` from the start of `index`. */
  countToCover(index: number, distance: number): number {
    if (index > this.maxIndex || distance <= 0) return 0;
    const target = this.offsetOf(index) + distance;
    const end = this.offsetOf(this.maxIndex) + this.sizeOf(this.maxIndex);
    if (target >= end) return this.maxIndex - index + 1;
    const located = this.indexAt(target);
    return located.index - index + (located.partial > 0 ? 1 : 0);
  }
}

export interface GridCellRectOptions {
  readonly scale: number;
  readonly scrollX: number;
  readonly scrollY: number;
  readonly headerWidth: number;
  readonly headerHeight: number;
}

export interface GridVisibleRangeOptions extends GridCellRectOptions {
  readonly width: number;
  readonly height: number;
  readonly buffer?: number;
}

export interface GridVisibleGeometry {
  readonly range: ViewportRange;
  readonly offsetX: number;
  readonly offsetY: number;
  readonly frozenWidth: number;
  readonly frozenHeight: number;
}

export interface GridScrollToCellOptions {
  readonly scale: number;
  readonly viewportWidth: number;
  readonly viewportHeight: number;
  readonly currentX: number;
  readonly currentY: number;
  readonly headerWidth: number;
  readonly headerHeight: number;
  readonly align: 'nearest' | 'start' | 'center' | 'end';
}

/** Pure worksheet geometry shared by both XLSX viewer facades. */
export class GridGeometry {
  private static readonly cache = new WeakMap<Worksheet, GridGeometry>();

  static forWorksheet(worksheet: Worksheet, mdw: number): GridGeometry {
    const cached = this.cache.get(worksheet);
    if (cached) return cached;
    const geometry = new GridGeometry(worksheet, mdw);
    this.cache.set(worksheet, geometry);
    return geometry;
  }

  static invalidate(worksheet: Worksheet): void {
    this.cache.delete(worksheet);
  }

  readonly col: GridAxisGeometry;
  readonly row: GridAxisGeometry;
  private readonly freezeRows: number;
  private readonly freezeCols: number;
  private scaledCache: Readonly<{
    scale: number;
    row: GridAxisGeometry;
    col: GridAxisGeometry;
  }> | null = null;

  private constructor(worksheet: Worksheet, mdw: number) {
    this.freezeRows = Math.min(MAX_WORKSHEET_ROW, Math.max(0, worksheet.freezeRows ?? 0));
    this.freezeCols = Math.min(MAX_WORKSHEET_COL, Math.max(0, worksheet.freezeCols ?? 0));
    this.col = new GridAxisGeometry(
      worksheet.colWidths,
      colWidthToPx(worksheet.defaultColWidth, mdw),
      (raw) => colWidthToPx(raw, mdw),
      MAX_WORKSHEET_COL,
    );
    this.row = new GridAxisGeometry(
      worksheet.rowHeights,
      rowHeightToPx(worksheet.defaultRowHeight),
      rowHeightToPx,
      MAX_WORKSHEET_ROW,
    );
  }

  logicalFrozenExtent(): { width: number; height: number } {
    return {
      width: this.col.offsetOf(this.freezeCols + 1),
      height: this.row.offsetOf(this.freezeRows + 1),
    };
  }

  roundedFrozenExtent(scale: number): { width: number; height: number } {
    const axes = this.scaledAxes(scale);
    return {
      width: axes.col.offsetOf(this.freezeCols + 1),
      height: axes.row.offsetOf(this.freezeRows + 1),
    };
  }

  logicalContentExtent(
    maxRow: number,
    maxCol: number,
    headerWidth: number,
    headerHeight: number,
  ): { width: number; height: number } {
    return {
      width: headerWidth + this.col.offsetOf(Math.min(MAX_WORKSHEET_COL, maxCol) + 1),
      height: headerHeight + this.row.offsetOf(Math.min(MAX_WORKSHEET_ROW, maxRow) + 1),
    };
  }

  roundedContentExtent(
    maxRow: number,
    maxCol: number,
    scale: number,
    headerWidth: number,
    headerHeight: number,
  ): { width: number; height: number } {
    const axes = this.scaledAxes(scale);
    return {
      width: Math.round(headerWidth * scale)
        + axes.col.offsetOf(Math.min(MAX_WORKSHEET_COL, maxCol) + 1),
      height: Math.round(headerHeight * scale)
        + axes.row.offsetOf(Math.min(MAX_WORKSHEET_ROW, maxRow) + 1),
    };
  }

  cellAt(
    innerX: number,
    innerY: number,
    viewport: { readonly scrollX: number; readonly scrollY: number },
  ): { row: number; col: number } | null {
    if (innerX < 0 || innerY < 0) return null;
    const row = this.rowAt(innerY, viewport.scrollY);
    if (row === null) return null;
    const col = this.colAt(innerX, viewport.scrollX);
    return col === null ? null : { row, col };
  }

  rowAt(innerY: number, scrollY: number): number | null {
    if (innerY < 0) return null;
    const frozenHeight = this.row.offsetOf(this.freezeRows + 1);
    return innerY < frozenHeight
      ? this.indexWithinFrozen(this.row, innerY, this.freezeRows)
      : this.row.scrollableIndexAt(innerY - frozenHeight + scrollY, this.freezeRows + 1);
  }

  colAt(innerX: number, scrollX: number): number | null {
    if (innerX < 0) return null;
    const frozenWidth = this.col.offsetOf(this.freezeCols + 1);
    return innerX < frozenWidth
      ? this.indexWithinFrozen(this.col, innerX, this.freezeCols)
      : this.col.scrollableIndexAt(innerX - frozenWidth + scrollX, this.freezeCols + 1);
  }

  cellRect(
    row: number,
    col: number,
    options: GridCellRectOptions,
  ): { x: number; y: number; w: number; h: number } | null {
    if (row < 1 || row > MAX_WORKSHEET_ROW || col < 1 || col > MAX_WORKSHEET_COL) return null;
    const axes = this.scaledAxes(options.scale);
    const headerX = Math.round(options.headerWidth * options.scale);
    const headerY = Math.round(options.headerHeight * options.scale);
    const frozen = this.roundedFrozenExtent(options.scale);

    const x = col <= this.freezeCols
      ? headerX + axes.col.offsetOf(col)
      : this.scrollableCellPosition(
          this.col,
          axes.col,
          col,
          this.freezeCols,
          options.scrollX,
          options.scale,
          headerX + frozen.width,
        );
    const y = row <= this.freezeRows
      ? headerY + axes.row.offsetOf(row)
      : this.scrollableCellPosition(
          this.row,
          axes.row,
          row,
          this.freezeRows,
          options.scrollY,
          options.scale,
          headerY + frozen.height,
        );
    return { x, y, w: axes.col.sizeOf(col), h: axes.row.sizeOf(row) };
  }

  visibleRange(options: GridVisibleRangeOptions): GridVisibleGeometry {
    const frozen = this.logicalFrozenExtent();
    const colStart = this.col.indexAt(options.scrollX + this.col.offsetOf(this.freezeCols + 1));
    const rowStart = this.row.indexAt(options.scrollY + this.row.offsetOf(this.freezeRows + 1));
    const cellWidth = options.width / options.scale - options.headerWidth - frozen.width;
    const cellHeight = options.height / options.scale - options.headerHeight - frozen.height;
    const buffer = options.buffer ?? 0;
    return {
      range: {
        row: rowStart.index,
        col: colStart.index,
        rows: this.row.countToCover(rowStart.index, cellHeight + rowStart.partial * 2) + buffer,
        cols: this.col.countToCover(colStart.index, cellWidth + colStart.partial * 2) + buffer,
      },
      offsetX: colStart.partial,
      offsetY: rowStart.partial,
      frozenWidth: frozen.width,
      frozenHeight: frozen.height,
    };
  }

  scrollOffsetForCell(
    row: number,
    col: number,
    options: GridScrollToCellOptions,
  ): { x: number; y: number } {
    const scaledFrozen = this.roundedFrozenExtent(options.scale);
    const viewTop = Math.round(options.headerHeight * options.scale) + scaledFrozen.height;
    const viewLeft = Math.round(options.headerWidth * options.scale) + scaledFrozen.width;
    let y = options.currentY;
    if (row > this.freezeRows && row <= MAX_WORKSHEET_ROW) {
      const cellStart = (this.row.offsetOf(row) - this.row.offsetOf(this.freezeRows + 1))
        * options.scale;
      const cellSize = this.row.sizeOf(row) * options.scale;
      y = this.alignedOffset(
        cellStart,
        cellSize,
        options.currentY,
        viewTop,
        options.viewportHeight,
        options.align,
      );
    }
    let x = options.currentX;
    if (col > this.freezeCols && col <= MAX_WORKSHEET_COL) {
      const cellStart = (this.col.offsetOf(col) - this.col.offsetOf(this.freezeCols + 1))
        * options.scale;
      const cellSize = this.col.sizeOf(col) * options.scale;
      x = this.alignedOffset(
        cellStart,
        cellSize,
        options.currentX,
        viewLeft,
        options.viewportWidth,
        options.align,
      );
    }
    return { x: Math.max(0, x), y: Math.max(0, y) };
  }

  private scaledAxes(scale: number): { row: GridAxisGeometry; col: GridAxisGeometry } {
    if (this.scaledCache?.scale === scale) return this.scaledCache;
    const scaled = { scale, row: this.row.scaled(scale), col: this.col.scaled(scale) };
    this.scaledCache = scaled;
    return scaled;
  }

  private indexWithinFrozen(
    axis: GridAxisGeometry,
    offset: number,
    frozenCount: number,
  ): number | null {
    if (frozenCount === 0) return null;
    const index = axis.indexAt(offset).index;
    return index <= frozenCount ? index : null;
  }

  private scrollableCellPosition(
    logicalAxis: GridAxisGeometry,
    scaledAxis: GridAxisGeometry,
    index: number,
    frozenCount: number,
    scroll: number,
    scale: number,
    scrollAreaStart: number,
  ): number {
    const start = logicalAxis.indexAt(scroll + logicalAxis.offsetOf(frozenCount + 1));
    return scrollAreaStart - start.partial * scale
      + scaledAxis.offsetOf(index) - scaledAxis.offsetOf(start.index);
  }

  private alignedOffset(
    cellStart: number,
    cellSize: number,
    current: number,
    viewportStart: number,
    viewportEnd: number,
    align: GridScrollToCellOptions['align'],
  ): number {
    const cellOnScreen = viewportStart + cellStart - current;
    if (align === 'start') return cellStart;
    if (align === 'center') return cellStart - (viewportEnd - viewportStart - cellSize) / 2;
    if (align === 'end') return cellStart - (viewportEnd - viewportStart - cellSize);
    if (cellOnScreen < viewportStart) return cellStart;
    if (cellOnScreen + cellSize > viewportEnd) {
      return cellStart - (viewportEnd - viewportStart - cellSize);
    }
    return current;
  }
}
