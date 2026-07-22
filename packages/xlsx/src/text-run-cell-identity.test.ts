import { describe, expect, it } from 'vitest';
import { renderViewport } from './renderer.js';
import type { Cell, Styles, Worksheet, XlsxTextRunInfo } from './types.js';

const STYLES: Styles = {
  fonts: [{
    bold: false,
    italic: false,
    underline: false,
    strike: false,
    size: 11,
    color: null,
    name: null,
  }],
  fills: [],
  borders: [],
  cellXfs: [{ fontId: 0, fillId: 0, borderId: 0, numFmtId: 0 } as Styles['cellXfs'][number]],
  numFmts: [],
  dxfs: [],
};

function cell(col: number, text: string): Cell {
  return {
    col,
    row: 1,
    value: { type: 'text', text },
    styleIndex: 0,
  } as Cell;
}

function sheet(): Worksheet {
  return {
    name: 'Quarter 1',
    rows: [{ index: 1, height: null, cells: [cell(1, 'alpha'), cell(2, 'beta')] }],
    colWidths: {},
    rowHeights: {},
    defaultColWidth: 8.43,
    defaultRowHeight: 15,
    mergeCells: [],
    freezeRows: 0,
    freezeCols: 0,
    conditionalFormats: [],
    images: [],
    charts: [],
    defaultFontFamily: 'Calibri',
    defaultFontSize: 11,
  } as Worksheet;
}

function recordingCtx(width = 400, height = 200): CanvasRenderingContext2D {
  let font = '11px sans-serif';
  const ctx: Record<string, unknown> = {
    canvas: { width, height },
    get font() { return font; },
    set font(value: string) { font = value; },
    fillStyle: '#000',
    strokeStyle: '#000',
    lineWidth: 1,
    textBaseline: 'alphabetic',
    textAlign: 'left',
    letterSpacing: '0px',
    direction: 'ltr',
    globalAlpha: 1,
    measureText: (text: string) => ({ width: [...text].length * 8 }),
    fillText: () => {},
    strokeText: () => {},
    fillRect: () => {},
    strokeRect: () => {},
    clearRect: () => {},
    beginPath: () => {},
    closePath: () => {},
    moveTo: () => {},
    lineTo: () => {},
    rect: () => {},
    arc: () => {},
    fill: () => {},
    stroke: () => {},
    clip: () => {},
    save: () => {},
    restore: () => {},
    translate: () => {},
    rotate: () => {},
    scale: () => {},
    setLineDash: () => {},
    setTransform: () => {},
    createLinearGradient: () => ({ addColorStop: () => {} }),
  };
  return ctx as unknown as CanvasRenderingContext2D;
}

describe('renderViewport text-run cell identity', () => {
  it('emits the worksheet name and A1 cell reference used by Office CLI', () => {
    const runs: XlsxTextRunInfo[] = [];
    renderViewport(recordingCtx(), sheet(), STYLES, { row: 1, col: 1, rows: 1, cols: 2 }, {
      onTextRun: (run) => runs.push(run),
    });

    expect(runs.map(({ text, sheetName, cellRef, row, col }) => (
      { text, sheetName, cellRef, row, col }
    ))).toEqual([
      { text: 'alpha', sheetName: 'Quarter 1', cellRef: 'A1', row: 1, col: 1 },
      { text: 'beta', sheetName: 'Quarter 1', cellRef: 'B1', row: 1, col: 2 },
    ]);
  });
});
