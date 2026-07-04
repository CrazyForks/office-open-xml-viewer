import { describe, it, expect } from 'vitest';
import { renderDocumentToCanvas } from './renderer.js';
import { MAX_CANVAS_DIMENSION, MAX_CANVAS_AREA } from '@silurus/ooxml-core';
import type { BodyElement, DocParagraph, DocxDocumentModel, SectionProps } from './types';

// RB5 — the docx renderer must clamp its backing store to browser canvas limits.
// A pathological `opts.width` (or page size × dpr) would otherwise trip the
// silent browser clamp and render blank. Here we drive a mock canvas that records
// the assigned backing-store size and the `ctx.scale()` factor, and assert the
// backing store stays within the caps while the CSS box keeps the requested size
// and the scale is reduced to compensate (aspect preserved).

interface Recording {
  scaleX: number | null;
  scaleY: number | null;
}

function makeRecordingCanvas(): { canvas: HTMLCanvasElement; rec: Recording } {
  let font = '10px serif';
  const rec: Recording = { scaleX: null, scaleY: null };
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const ctx = {
    get font() {
      return font;
    },
    set font(v: string) {
      font = v;
    },
    letterSpacing: '0px',
    measureText: (s: string) => {
      const p = px();
      return {
        width: [...s].length * p * 0.5,
        fontBoundingBoxAscent: p * 0.8,
        fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8,
        actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() {},
    restore() {},
    closePath() {},
    beginPath() {},
    moveTo() {},
    lineTo() {},
    arc() {},
    fill() {},
    stroke() {},
    fillRect() {},
    strokeRect() {},
    clip() {},
    rect() {},
    // Only the FIRST scale() is the dpr transform we care about; later flip
    // scales are local and wrapped in save/restore, but the render sets the dpr
    // scale exactly once up front, so record only the first.
    scale(x: number, y: number) {
      if (rec.scaleX === null) {
        rec.scaleX = x;
        rec.scaleY = y;
      }
    },
    translate() {},
    rotate() {},
    setLineDash() {},
    clearRect() {},
    quadraticCurveTo() {},
    bezierCurveTo() {},
    createLinearGradient() {
      return { addColorStop() {} };
    },
    drawImage() {},
    fillText() {},
    strokeText() {},
    fillStyle: '#000',
    strokeStyle: '#000',
    lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign,
    direction: 'ltr' as CanvasDirection,
    globalAlpha: 1,
    lineCap: 'butt' as CanvasLineCap,
    lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = {
    width: 0,
    height: 0,
    style: {} as Record<string, string>,
    getContext: () => ctx,
  };
  return { canvas: canvas as unknown as HTMLCanvasElement, rec };
}

function para(text = 'x'): DocParagraph {
  return {
    type: 'paragraph',
    alignment: 'left',
    indentLeft: 0,
    indentRight: 0,
    indentFirst: 0,
    spaceBefore: 0,
    spaceAfter: 0,
    lineSpacing: null,
    numbering: null,
    tabStops: [],
    runs: [
      {
        type: 'text',
        text,
        bold: false,
        italic: false,
        underline: false,
        strikethrough: false,
        fontSize: 12,
        color: null,
        fontFamily: 'Times New Roman',
        fontFamilyEastAsia: '',
        isLink: false,
        background: null,
        vertAlign: null,
        hyperlink: null,
      } as DocParagraph['runs'][number],
    ],
    defaultFontSize: 12,
    defaultFontFamily: 'Times New Roman',
    widowControl: false,
  } as unknown as DocParagraph;
}

function doc(pageWidth: number, pageHeight: number): DocxDocumentModel {
  const section = {
    pageWidth,
    pageHeight,
    marginTop: 5,
    marginRight: 5,
    marginBottom: 5,
    marginLeft: 5,
    headerDistance: 4,
    footerDistance: 4,
    titlePage: false,
    evenAndOddHeaders: false,
  } as SectionProps;
  return {
    section,
    body: [para() as unknown as BodyElement],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { 'Times New Roman': 'roman' },
  } as unknown as DocxDocumentModel;
}

describe('docx renderer clamps the backing store to browser limits (RB5)', () => {
  it('a normal page renders at the requested dpr (no clamp)', async () => {
    const { canvas, rec } = makeRecordingCanvas();
    // width 800, page 800×600 → 1× scale; dpr 2 → 1600×1200 backing, in budget.
    await renderDocumentToCanvas(doc(800, 600), canvas, 0, { dpr: 2, width: 800 });
    expect(canvas.width).toBe(1600);
    expect(canvas.height).toBe(1200);
    expect(rec.scaleX).toBe(2);
    expect(rec.scaleY).toBe(2);
  });

  it('a pathologically wide request is clamped, aspect preserved, scale reduced', async () => {
    const { canvas, rec } = makeRecordingCanvas();
    // Request a 50000-px-wide render of a 50000×500 page: 25 MP at dpr 1, both
    // over the area cap and (width) over the axis cap.
    await renderDocumentToCanvas(doc(50000, 500), canvas, 0, { dpr: 1, width: 50000 });
    // Backing store within BOTH caps.
    expect(canvas.width).toBeLessThanOrEqual(MAX_CANVAS_DIMENSION);
    expect(canvas.height).toBeLessThanOrEqual(MAX_CANVAS_DIMENSION);
    expect(canvas.width * canvas.height).toBeLessThanOrEqual(MAX_CANVAS_AREA);
    // The dpr transform was reduced below 1 to fit (folded clamp factor).
    expect(rec.scaleX).toBeLessThan(1);
    expect(rec.scaleX).toBeGreaterThan(0);
    // Uniform scale on both axes → aspect preserved.
    expect(rec.scaleX).toBeCloseTo(rec.scaleY as number, 10);
  });

  it('a huge dpr on a moderate page is clamped by the area cap', async () => {
    const { canvas, rec } = makeRecordingCanvas();
    // 4000×3000 page at dpr 4 → 16000×12000 = 192 MP, far over the area cap.
    await renderDocumentToCanvas(doc(4000, 3000), canvas, 0, { dpr: 4, width: 4000 });
    expect(canvas.width * canvas.height).toBeLessThanOrEqual(MAX_CANVAS_AREA);
    // Effective scale < requested dpr (4).
    expect(rec.scaleX).toBeLessThan(4);
    expect(rec.scaleX).toBeGreaterThan(0);
  });
});
