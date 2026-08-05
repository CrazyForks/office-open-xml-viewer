import type { Worksheet } from '../types.js';
import { colWidthToPx, rowHeightToPx } from '../renderer.js';

/** Sparse cumulative axis geometry. Offsets are logical, unscaled pixels. */
export class GridAxisGeometry {
  private readonly indices: number[];
  private readonly cumulativeDelta: number[];

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
    let accumulated = 0;
    for (let index = 0; index < this.indices.length; index++) {
      accumulated += toPx(customs[this.indices[index]]) - defaultPx;
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

  private sizeOf(index: number): number {
    return this.offsetOf(index + 1) - this.offsetOf(index);
  }
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

  private constructor(worksheet: Worksheet, mdw: number) {
    this.col = new GridAxisGeometry(
      worksheet.colWidths,
      colWidthToPx(worksheet.defaultColWidth, mdw),
      (raw) => colWidthToPx(raw, mdw),
      16_384,
    );
    this.row = new GridAxisGeometry(
      worksheet.rowHeights,
      rowHeightToPx(worksheet.defaultRowHeight),
      rowHeightToPx,
      1_048_576,
    );
  }
}
