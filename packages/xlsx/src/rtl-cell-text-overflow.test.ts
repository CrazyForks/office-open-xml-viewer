import { describe, expect, it } from 'vitest';
import { colWidthToPx, getMdwForWorksheet, renderViewport } from './renderer.js';
import type { Styles, Worksheet } from './types.js';

const TEXT = 'ورقة ثانية بالعربية';

const STYLES: Styles = {
  fonts: [{
    bold: false,
    italic: false,
    underline: false,
    strike: false,
    size: 14,
    color: null,
    name: 'Amiri',
  }],
  fills: [],
  borders: [],
  cellXfs: [{
    fontId: 0,
    fillId: 0,
    borderId: 0,
    numFmtId: 0,
    alignH: 'right',
    readingOrder: 2,
  } as Styles['cellXfs'][number]],
  numFmts: [],
  dxfs: [],
};

function worksheet(): Worksheet {
  return {
    name: 'ورقة٢',
    rows: [{
      index: 1,
      height: null,
      cells: [{
        row: 1,
        col: 1,
        styleIndex: 0,
        value: { type: 'text', text: TEXT },
      }],
    }],
    colWidths: { 1: 8.43, 2: 8.43 },
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
    rightToLeft: true,
  } as Worksheet;
}

function recordingContext(): {
  ctx: CanvasRenderingContext2D;
  textClips: Array<{ text: string; x: number; width: number }>;
} {
  let lastRect = { x: 0, width: 0 };
  const textClips: Array<{ text: string; x: number; width: number }> = [];
  const state: Record<string, unknown> = {
    canvas: { width: 240, height: 100 },
    fillStyle: '#000000',
    strokeStyle: '#000000',
    lineWidth: 1,
    font: '14px Amiri',
    textAlign: 'left',
    textBaseline: 'alphabetic',
    direction: 'ltr',
    globalAlpha: 1,
    letterSpacing: '0px',
    measureText: (text: string) => ({ width: [...text].length * 10 }),
    rect: (x: number, _y: number, width: number) => { lastRect = { x, width }; },
    fillText: (text: string) => { textClips.push({ text, ...lastRect }); },
    createLinearGradient: () => ({ addColorStop: () => {} }),
  };
  const noOp = () => {};
  const ctx = new Proxy(state, {
    get(target, property) {
      return property in target ? target[property as string] : noOp;
    },
    set(target, property, value) {
      target[property as string] = value;
      return true;
    },
  });
  return { ctx: ctx as unknown as CanvasRenderingContext2D, textClips };
}

describe('RTL cell text overflow', () => {
  it('extends right-aligned A1 text physically left into the empty B1 cell', () => {
    const recording = recordingContext();
    const sheet = worksheet();

    renderViewport(
      recording.ctx,
      sheet,
      STYLES,
      { row: 1, col: 1, rows: 1, cols: 2 },
    );

    const textDraw = recording.textClips.find((call) => call.text === TEXT);
    const cellWidth = colWidthToPx(sheet.colWidths[1], getMdwForWorksheet(sheet));
    expect(textDraw).toBeDefined();
    expect(textDraw?.width).toBe(cellWidth * 2);
  });
});
