import { describe, expect, it } from 'vitest';
import type { Worksheet } from '../types.js';
import { colWidthToPx, getMdwForWorksheet, rowHeightToPx } from '../renderer.js';
import { GridGeometry } from './grid-geometry.js';

function worksheet(): Worksheet {
  return {
    name: 'Geometry',
    rows: [],
    colWidths: { 1: 12, 2: 20, 4: 30, 16_384: 5 },
    rowHeights: { 1: 25, 2: 40, 5: 60, 1_048_576: 12 },
    defaultColWidth: 8.43,
    defaultRowHeight: 15,
    mergeCells: [],
    freezeRows: 2,
    freezeCols: 2,
    conditionalFormats: [],
    images: [],
    charts: [],
  } as Worksheet;
}

describe('GridGeometry', () => {
  it('computes far-cell rectangles from cumulative axes while preserving per-cell scale rounding', () => {
    const ws = worksheet();
    const geometry = GridGeometry.forWorksheet(ws, getMdwForWorksheet(ws));
    const scale = 1.25;
    const scrollX = 231.5;
    const scrollY = 417.25;
    const rect = geometry.cellRect(1_048_576, 16_384, {
      scale,
      scrollX,
      scrollY,
      headerWidth: 48,
      headerHeight: 24,
    });

    const sp = (value: number) => Math.round(value * scale);
    const defaultCol = sp(colWidthToPx(ws.defaultColWidth, getMdwForWorksheet(ws)));
    const defaultRow = sp(rowHeightToPx(ws.defaultRowHeight));
    const scaledCol = (index: number) => sp(colWidthToPx(ws.colWidths[index] ?? ws.defaultColWidth, getMdwForWorksheet(ws)));
    const scaledRow = (index: number) => sp(rowHeightToPx(ws.rowHeights[index] ?? ws.defaultRowHeight));
    const sparseColDelta = [1, 2, 4]
      .reduce((sum, index) => sum + scaledCol(index) - defaultCol, 0);
    const sparseRowDelta = [1, 2, 5]
      .reduce((sum, index) => sum + scaledRow(index) - defaultRow, 0);
    const frozenW = scaledCol(1) + scaledCol(2);
    const frozenH = scaledRow(1) + scaledRow(2);
    const logicalColStart = geometry.col.indexAt(scrollX + geometry.col.offsetOf(3));
    const logicalRowStart = geometry.row.indexAt(scrollY + geometry.row.offsetOf(3));
    const expectedX = sp(48) + frozenW - logicalColStart.partial * scale
      + (16_384 - 1) * defaultCol + sparseColDelta
      - ((logicalColStart.index - 1) * defaultCol
        + [1, 2, 4].filter((index) => index < logicalColStart.index)
          .reduce((sum, index) => sum + scaledCol(index) - defaultCol, 0));
    const expectedY = sp(24) + frozenH - logicalRowStart.partial * scale
      + (1_048_576 - 1) * defaultRow + sparseRowDelta
      - ((logicalRowStart.index - 1) * defaultRow
        + [1, 2, 5].filter((index) => index < logicalRowStart.index)
          .reduce((sum, index) => sum + scaledRow(index) - defaultRow, 0));

    expect(rect).toEqual({
      x: expectedX,
      y: expectedY,
      w: scaledCol(16_384),
      h: scaledRow(1_048_576),
    });
  });

  it('rejects coordinates outside the worksheet limits', () => {
    const geometry = GridGeometry.forWorksheet(worksheet(), 7);
    const options = {
      scale: 1,
      scrollX: 0,
      scrollY: 0,
      headerWidth: 48,
      headerHeight: 24,
    };
    expect(geometry.cellRect(0, 1, options)).toBeNull();
    expect(geometry.cellRect(1_048_577, 1, options)).toBeNull();
    expect(geometry.cellRect(1, 0, options)).toBeNull();
    expect(geometry.cellRect(1, 16_385, options)).toBeNull();
  });

  it('derives hit tests and visible ranges from the same frozen-pane axes', () => {
    const ws = worksheet();
    const geometry = GridGeometry.forWorksheet(ws, getMdwForWorksheet(ws));
    const frozen = geometry.logicalFrozenExtent();
    const scrollX = 51;
    const scrollY = 37;

    expect(geometry.cellAt(frozen.width + 1, frozen.height + 1, { scrollX, scrollY })).toEqual({
      row: geometry.row.scrollableIndexAt(1 + scrollY, 3),
      col: geometry.col.scrollableIndexAt(1 + scrollX, 3),
    });
    const visible = geometry.visibleRange({
      width: 640,
      height: 480,
      scale: 1,
      scrollX,
      scrollY,
      headerWidth: 48,
      headerHeight: 24,
      buffer: 2,
    });
    expect(visible.range.row).toBeGreaterThanOrEqual(3);
    expect(visible.range.col).toBeGreaterThanOrEqual(3);
    expect(visible.range.rows).toBeGreaterThan(2);
    expect(visible.range.cols).toBeGreaterThan(2);
  });

  it('matches the former visible-band walk across scales and partial scroll offsets', () => {
    const ws = worksheet();
    const geometry = GridGeometry.forWorksheet(ws, getMdwForWorksheet(ws));
    const oldCount = (
      axis: typeof geometry.row,
      start: number,
      partial: number,
      available: number,
      max: number,
    ) => {
      let accumulated = -partial;
      let index = start;
      let count = 0;
      while (accumulated < available + partial && index <= max) {
        accumulated += axis.sizeOf(index);
        count++;
        index++;
      }
      return count + 2;
    };

    for (const scale of [0.75, 1, 1.25]) {
      for (const [scrollX, scrollY] of [[0, 0], [51, 37], [1234.5, 987.25]]) {
        const width = 641;
        const height = 479;
        const visible = geometry.visibleRange({
          width,
          height,
          scale,
          scrollX,
          scrollY,
          headerWidth: 48,
          headerHeight: 24,
          buffer: 2,
        });
        const frozen = geometry.logicalFrozenExtent();
        const availableWidth = width / scale - 48 - frozen.width;
        const availableHeight = height / scale - 24 - frozen.height;
        expect(visible.range.cols).toBe(oldCount(
          geometry.col,
          visible.range.col,
          visible.offsetX,
          availableWidth,
          16_384,
        ));
        expect(visible.range.rows).toBe(oldCount(
          geometry.row,
          visible.range.row,
          visible.offsetY,
          availableHeight,
          1_048_576,
        ));
      }
    }
  });

  it('does not inspect every preceding band when locating a far cell', () => {
    let rowReads = 0;
    const ws = worksheet();
    ws.rowHeights = new Proxy(ws.rowHeights, {
      get(target, key, receiver) {
        if (typeof key === 'string' && /^\d+$/.test(key)) rowReads++;
        return Reflect.get(target, key, receiver);
      },
    });
    const geometry = GridGeometry.forWorksheet(ws, getMdwForWorksheet(ws));

    expect(geometry.cellRect(1_048_576, 16_384, {
      scale: 1.25,
      scrollX: 0,
      scrollY: 0,
      headerWidth: 48,
      headerHeight: 24,
    })).not.toBeNull();
    expect(rowReads).toBeLessThan(100);
  });
});
